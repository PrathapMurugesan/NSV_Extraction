"""
pipeline.py - the steps a deliverable script calls, in order:

    files = find_workbooks(run, "d2", WORKBOOKS)            1. which expected workbooks are in the folder
    extract_workbook(run, files, "comparison", WORKBOOKS)   2. per workbook: sheets -> read -> check -> save
    extract_csv_folder(run, "d3", GROUPS)                   (D3: csv files instead of workbooks)
    extract_everything(run, "d2")                           (settings.EXTRACT_ALL: no config at all)

What happens to each sheet of a workbook (extract_workbook):
    a. find it by name (exact or fuzzy)                     missing          -> logged, manifest
    b. read it as a table (header found, merges handled)    unreadable       -> logged, manifest
    c. header check against schemas/...csv  (the GATE)      headers missing  -> REJECTED, not saved
    d. formulas -> formulas.json, legends -> data_dictionary.json
    e. profile -> profile.json
    f. reports: empty sheet, defect columns, lanes (never stop anything)
    g. save the table as parquet + csv (all text, exactly as delivered), named d2_<workbook>_<sheet>

Sheet options in WORKBOOKS (deliverables/dN.py):
    sheet_names   known names of the sheet (required)
    kind          "table" (default) or "cells" (report-style sheet saved cell by cell as JSON)
    skip          True = known sheet, deliberately not extracted
    header_row, header_rows, info_rows, skip_rows, columns    -> see nsv/excel.py
    cells         also save this column range cell by cell, e.g. "I:M" (side tables next to the main table)
"""
from pathlib import PurePosixPath

import settings
from nsv.checks import check_headers, defect_columns_report, empty_sheet_report, load_schema
from nsv.excel import ExcelWorkbook
from nsv.headers import final_names, match_headers
from nsv.formulas import describe_formulas
from nsv.log import get_logger
from nsv.matching import classify_path, identify_names, match_score, normalize_name, safe_name
from nsv.profiler import profile_table

logger = get_logger("nsv.pipeline")

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


# =============================================================================
# 1. Which expected workbooks are in the folder?
# =============================================================================
def find_workbooks(run, folder, workbooks):
    """
    Returns {workbook key: file path} for the workbooks that were found,
    or None when one is missing and settings.STOP_IF_WORKBOOK_MISSING is on.
    Only files directly in the folder are considered (not sub-folders); Excel lock files (~$) are ignored.

    1. by file name (exact / normalized / fuzzy, see matching.py)
    2. if not found by name and the workbook config has "file_keywords" + "identify_by_sheets":
       a file whose name contains every keyword AND that has every one of those sheets
       (e.g. D2: the name changes per road and month, but always contains 'comparison')
    """
    paths = {PurePosixPath(p).name: p for p in _excel_files(folder)}
    logger.info(f"{len(paths)} Excel file(s) in '{settings.INPUT.describe()}/{folder}': {list(paths)}")

    # ---- 1. by file name ----
    expected = {key: cfg["file_names"] for key, cfg in workbooks.items()}
    found, missing, unexpected = identify_names(list(paths), expected, settings.FUZZY_THRESHOLD, "workbook",
                                                quiet=True)
    how = {key: "file name" for key in found}

    # ---- 2. by keywords in the name + the sheets inside ----
    for key in list(missing):
        cfg = workbooks[key]
        if not (cfg.get("file_keywords") and cfg.get("identify_by_sheets")):
            continue
        name = _find_by_content(key, cfg, [n for n in unexpected], paths)
        if name:
            found[key] = name
            how[key] = f"name contains {cfg['file_keywords']} + has sheets {cfg['identify_by_sheets']}"
            missing.remove(key)
            unexpected.remove(name)

    # ---- log what is still missing / unexpected ----
    for key in missing:
        logger.error(f"[{key}] no workbook found (expected like '{workbooks[key]['file_names'][0]}'"
                     + (f", or a name containing {workbooks[key]['file_keywords']} with sheets "
                        f"{workbooks[key]['identify_by_sheets']})" if workbooks[key].get("file_keywords") else ")"))
    for name in unexpected:
        logger.warning(f"unexpected workbook, ignored: '{name}'")

    for key in workbooks:
        run.manifest["workbooks"][key] = {"file": found.get(key), "status": "found" if key in found else "missing",
                                          "identified_by": how.get(key), "sheets": {}}
    run.manifest["unexpected_files"] = unexpected

    if missing and settings.STOP_IF_WORKBOOK_MISSING:
        logger.error(f"[{run.survey.label}] {run.deliverable} NOT extracted: missing workbook(s) {missing} "
                     f"(settings.STOP_IF_WORKBOOK_MISSING = True)")
        return None
    return {key: paths[name] for key, name in found.items()}


def _find_by_content(key, cfg, names, paths):
    """The one file whose name contains all keywords and that has all identify_by_sheets; None otherwise."""
    keywords = [normalize_name(k, is_file=False) for k in cfg["file_keywords"]]
    candidates = [n for n in names if all(k in normalize_name(n) for k in keywords)]
    accepted = []
    for name in candidates:
        try:
            wb = ExcelWorkbook(name, settings.INPUT.read_bytes(paths[name]))
            sheets = wb.sheet_names
            wb.close()
        except Exception as e:
            logger.warning(f"[{key}] '{name}' has {cfg['file_keywords']} in its name but cannot be opened: {e}")
            continue
        lacking = [s for s in cfg["identify_by_sheets"]
                   if not any(match_score(x, [s], is_file=False) >= settings.FUZZY_THRESHOLD for x in sheets)]
        if lacking:
            logger.warning(f"[{key}] '{name}' has {cfg['file_keywords']} in its name but not the sheets {lacking}")
        else:
            accepted.append(name)

    if len(accepted) == 1:
        logger.warning(f"[{key}] workbook '{accepted[0]}' recognised by its name containing "
                       f"{cfg['file_keywords']} and its sheets {cfg['identify_by_sheets']} (file name not in config)")
        return accepted[0]
    if len(accepted) > 1:
        logger.error(f"[{key}] ambiguous - several files have {cfg['file_keywords']} in their name and the "
                     f"expected sheets: {accepted}")
    return None


def _excel_files(folder):
    return [p for p in settings.INPUT.list_files(folder)
            if PurePosixPath(p).suffix.lower() in EXCEL_EXTENSIONS
            and not PurePosixPath(p).name.startswith("~$")
            and PurePosixPath(p).parent.as_posix() == folder.strip("/")]


# =============================================================================
# 2. One workbook: every configured sheet
# =============================================================================
def extract_workbook(run, files, wb_key, workbooks, check_nulls=False):
    """
    files:       {workbook key: file path}   (from find_workbooks)
    wb_key:      which workbook of WORKBOOKS to extract
    check_nulls: also run the defect-column / lane report (D2, D4)
    """
    config = workbooks[wb_key]
    entry = run.manifest["workbooks"].setdefault(wb_key, {"file": None, "status": "missing", "sheets": {}})
    if wb_key not in files:
        logger.error(f"[{wb_key}] workbook not found - nothing extracted")
        return

    # ---- open the workbook (from local disk or S3, via settings.INPUT) ----
    path = files[wb_key]
    logger.info(f"---------- [{wb_key}] {path} ----------")
    data = settings.INPUT.read_bytes(path)
    entry.update(_file_facts(path, data))
    try:
        wb = ExcelWorkbook(PurePosixPath(path).name, data)
    except Exception as e:
        logger.error(f"[{wb_key}] cannot open workbook: {e}")
        entry["status"] = "read_error"
        entry["error"] = str(e)
        return
    entry["sheet_names_in_file"] = wb.sheet_names

    # ---- a. find the configured sheets by name ----
    sheet_cfgs = config["sheets"]
    expected = {key: cfg["sheet_names"] for key, cfg in sheet_cfgs.items()}
    found, missing, unexpected = identify_names(wb.sheet_names, expected, settings.FUZZY_THRESHOLD,
                                                label=f"sheet in '{wb.file_name}'", is_file=False)
    entry["unexpected_sheets"] = unexpected
    for key in missing:
        entry["sheets"][key] = {"status": "missing", "expected_names": sheet_cfgs[key]["sheet_names"]}

    # ---- b-g. every sheet that was found ----
    meta = {"schema": {}, "profile": {}, "formulas": {}, "data_dictionary": {}}
    for sheet_key, sheet_name in found.items():
        options = sheet_cfgs[sheet_key]
        label = f"{wb_key}/{sheet_key}"
        result = {"sheet_name": sheet_name}
        entry["sheets"][sheet_key] = result

        if options.get("skip"):
            logger.info(f"[{label}] skipped (configured skip: True)")
            result["status"] = "skipped_by_config"
            continue
        try:
            if options.get("kind") == "cells":
                _save_cells(run, wb, wb_key, sheet_key, sheet_name, None, result)
                result["status"] = "saved_cells"
            else:
                _extract_sheet(run, wb, wb_key, sheet_key, sheet_name, options, check_nulls, result, meta)
        except Exception as e:
            logger.error(f"[{label}] could not read sheet '{sheet_name}': {e}")
            result.update(status="read_error", error=str(e))

    # ---- workbook-level JSONs ----
    for name, content in meta.items():
        if content:
            entry.setdefault("meta_files", []).append(run.save_json(content, wb_key, name))
    entry["status"] = "extracted"
    wb.close()


def _extract_sheet(run, wb, wb_key, sheet_key, sheet_name, options, check_nulls, result, meta):
    label = f"{wb_key}/{sheet_key}"

    # ---- b. read as a table ----
    table = wb.read_table(sheet_name, options)
    df = table.to_frame()
    result.update(header_excel_rows=table.header_excel_rows, rows=len(df), columns=len(table.columns),
                  columns_read=options.get("columns", "all"))
    logger.info(f"[{label}] read '{sheet_name}': header rows {table.header_excel_rows}, "
                f"{len(df)} data rows, {len(table.columns)} columns")

    # ---- c. GATE: headers - match to the schema (exact / normalized / unit / alias), then final names ----
    schema = None if settings.EXTRACT_ALL else load_schema(
        settings.SCHEMAS_DIR / run.deliverable / wb_key / f"{sheet_key}.csv")
    expected, aliases, info = schema if schema else ([], {}, {})
    matched = match_headers(table.names, expected, aliases, info,
                            letters=[c["letter"] for c in table.columns],
                            values=[[row[i] for row in table.rows] for i in range(len(table.columns))])
    if settings.EXTRACT_ALL:
        result["header_check"] = "off (EXTRACT_ALL)"
    elif schema is None:
        logger.warning(f"[{label}] no schema file - headers not checked "
                       f"(create it with: python tools/make_schema.py {run.deliverable})")
        result["header_check"] = "no schema file"
    else:
        result["header_check"] = check_headers(matched, label, gate=settings.CHECK_HEADERS)

    names = final_names(matched["columns"])            # schema name without unit, e.g. 'Chainage from'
    meta["schema"][sheet_key] = {
        "sheet_name": sheet_name,
        "status": "rejected" if result["header_check"] not in ("off (EXTRACT_ALL)", "no schema file")
                  and not result["header_check"]["ok"] else "accepted",
        "missing_headers": matched["missing"],
        "columns": [{"column": n, "excel_column": col["letter"], **c}
                    for n, col, c in zip(names, table.columns, matched["columns"])],
    }
    if meta["schema"][sheet_key]["status"] == "rejected":
        result["status"] = "rejected_headers"
        return
    table.rename(names)
    df = table.to_frame()

    # ---- d. formulas + legends ----
    formulas = wb.read_formulas(sheet_name, table)
    if formulas:
        meta["formulas"][sheet_key] = {"sheet_name": sheet_name, "columns": describe_formulas(formulas, table)}
        logger.info(f"[{label}] {len(formulas)} calculated column(s) documented in formulas.json")
    if table.dictionary:
        meta["data_dictionary"][sheet_key] = {"sheet_name": sheet_name, "columns": table.dictionary}

    # ---- e. profile ----
    profile = profile_table(df, table.columns, formulas.keys(), table.header_excel_rows)
    meta["profile"][sheet_key] = {"sheet_name": sheet_name, **profile}
    result["formula_only_rows"] = profile["formula_only_rows"]
    if profile["formula_only_rows"]:
        logger.warning(f"[{label}] {profile['formula_only_rows']} row(s) have values ONLY in formula columns "
                       f"(input columns empty - e.g. unfilled template rows) - saved as-is, see profile.json")

    # ---- f. reports (never stop anything) ----
    result["empty_check"] = empty_sheet_report(df, label)
    if check_nulls and options.get("check_nulls", True) and len(df):
        result["defect_check"] = defect_columns_report(df, settings.DEFECT_KEYWORDS, settings.NULL_LIMIT,
                                                       settings.NULL_IGNORE, label)

    # ---- g. save ----
    result["outputs"] = run.save_table(df, wb_key, sheet_key)
    if options.get("cells"):
        _save_cells(run, wb, wb_key, sheet_key, sheet_name, options["cells"], result)
    result["status"] = "saved"
    logger.info(f"[{label}] saved {len(df)} rows -> {run.file_name(wb_key, sheet_key)}.parquet + .csv")


def _save_cells(run, wb, wb_key, sheet_key, sheet_name, columns, result):
    """Report-style sheet (or a side area of a sheet): every non-empty cell with its address."""
    cells = wb.read_cells(sheet_name, columns)
    content = {"sheet_name": sheet_name, "columns": columns or "all",
               "merged_cells": wb.merged_refs(sheet_name), "cells": cells}
    area = [columns.replace(":", "")] if columns else []                 # e.g. 'IM' for columns I:M
    path = run.save_json(content, wb_key, sheet_key, *area, "cells")
    result.setdefault("outputs", []).append(path)
    logger.info(f"[{wb_key}/{sheet_key}] {len(cells)} cells{' in ' + columns if columns else ''} saved -> {path}")


def _file_facts(path, data):
    import hashlib
    return {"file": PurePosixPath(path).name, "path": path, "size_bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest()}


# =============================================================================
# 3. D3: every csv under a folder, any depth
# =============================================================================
def extract_csv_folder(run, folder, groups):
    """
    Each csv is saved as d3_<group>_<file>.parquet + .csv. The group comes from its folder / file name
    (groups = {"ca": ["as per ca", "ca"], ...}); no match -> group 'ungrouped'.
    """
    import io
    import pandas as pd

    files = [p for p in settings.INPUT.list_files(folder)
             if p.lower().endswith(".csv") and not PurePosixPath(p).name.startswith("~$")]
    logger.info(f"{len(files)} csv file(s) under '{settings.INPUT.describe()}/{folder}'")
    profiles, schemas, used = {}, {}, set()

    for path in files:
        relative = PurePosixPath(path).relative_to(folder).as_posix()
        group = classify_path(relative, groups) or "ungrouped"
        data = settings.INPUT.read_bytes(path)
        entry = {"group": group, **_file_facts(path, data)}
        run.manifest["csv_files"][relative] = entry

        # ---- read as text, exactly as in the file ----
        try:
            try:
                df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding="latin-1")
        except Exception as e:
            logger.error(f"[{relative}] could not read: {e}")
            entry.update(status="read_error", error=str(e))
            continue
        df = df.astype(object).where(df != "", None)
        matched = match_headers(list(df.columns), [])                  # no schema for csv: just unit-free names
        names = final_names(matched["columns"])
        schemas[relative] = {"columns": [{"column": n, **c} for n, c in zip(names, matched["columns"])]}
        columns = [{"name": n, "vendor_header": c["vendor_header"]} for n, c in zip(names, matched["columns"])]
        df.columns = names
        df.insert(0, "source_row", range(2, len(df) + 2))             # line 1 is the header

        # ---- report + profile + save ----
        entry["empty_check"] = empty_sheet_report(df, relative)
        profiles[relative] = profile_table(df, columns, header_excel_rows=[1])
        name = safe_name(PurePosixPath(relative).stem)
        while (group, name) in used:                                   # same file name in two folders
            name += "_2"
        used.add((group, name))
        entry["outputs"] = run.save_table(df, group, name)
        entry.update(status="saved", rows=len(df), columns=len(df.columns) - 1)
        logger.info(f"[{relative}] -> {group}: saved {len(df)} rows -> {run.file_name(group, name)}.parquet + .csv")

    if profiles:
        run.save_json(schemas, "schema")
        run.save_json(profiles, "profile")


# =============================================================================
# 4. Survey image folders (D4: crack_images_DHMS/L1/*.jpg, pothole images ...) - optional
# =============================================================================
IMAGE_COLUMNS_FIRST = ["image_folder", "lane", "file_name", "image_path", "status", "issues", "fields_found",
                       "chainage", "latitude", "longitude", "location_source", "location_check_m",
                       "chainage_start", "chainage_end", "nsv_cracks_chainage",
                       "ocr_chainage", "ocr_latitude", "ocr_longitude"]


def _image_location_tables(files, location):
    """location = {"defect_report": wb key, "sheet": name, "nsv_cracks": wb key} -> (lanes, offsets)."""
    from nsv.images import read_lane_locations, read_offsets
    lanes, offsets = {}, {}
    if not location:
        return lanes, offsets
    files = files or {}
    try:
        if files.get(location.get("defect_report")):
            lanes = read_lane_locations(settings.INPUT.read_bytes(files[location["defect_report"]]),
                                        location.get("sheet", "Data Sheet"))
            logger.info(f"[images] defect report: lat/long for lane(s) {sorted(lanes)}")
        else:
            logger.warning("[images] no defect report - images located by EXIF / OCR only")
        if files.get(location.get("nsv_cracks")):
            offsets = read_offsets(settings.INPUT.read_bytes(files[location["nsv_cracks"]]))
            logger.info(f"[images] NSV Cracks chainage offset per survey run: {offsets}")
    except Exception as e:
        logger.error(f"[images] could not read the location workbooks ({e}) - images located by EXIF / OCR only")
    return lanes, offsets


def extract_image_folders(run, folder, image_folders, files=None, location=None):
    """
    image_folders: {group: {"keywords": ["crack"], "required": False}}
    A sub-folder (directly under the deliverable folder) whose name contains all keywords holds that
    group's images, at any depth below it (lane folders L1..R4). One table per group:
    d4_images_<group>.parquet + .csv, one row per image (see nsv/images.py).
    A group with no folder: 'missing' - highlighted in the summary; a problem only if required.
    files + location (optional): workbooks used to give every image its chainage and lat/long
    (see images.locate); without them the OCR values stay as read.
    """
    import pandas as pd
    from nsv.images import lane_from_path, locate, ocr_available, read_images

    base = folder.strip("/")
    by_folder = {}                                         # first-level sub-folder -> image paths
    for path in settings.INPUT.list_files(folder):
        rel = PurePosixPath(path).relative_to(base)
        if len(rel.parts) >= 2 and rel.suffix.lower() in settings.IMAGE_EXTENSIONS:
            by_folder.setdefault(rel.parts[0], []).append(path)
    manifest = run.manifest.setdefault("image_folders", {})

    ocr = settings.OCR_IMAGES and ocr_available(settings.TESSERACT_CMD)
    if settings.OCR_IMAGES and not ocr and by_folder:
        logger.error("OCR is not available (tesseract not found) - images are listed, but the text on them "
                     "is not read. Install tesseract and set settings.TESSERACT_CMD.")
    cfg = {"ocr": ocr, "tesseract_cmd": settings.TESSERACT_CMD, "regions": settings.IMAGE_TEXT_REGIONS,
           "fields": settings.IMAGE_FIELDS}
    lanes, offsets = _image_location_tables(files, location) if by_folder else ({}, {})
    if offsets:
        run.manifest["image_chainage_offsets"] = offsets

    used = set()
    for group, gcfg in image_folders.items():
        keywords = [normalize_name(k, is_file=False) for k in gcfg["keywords"]]
        folders = sorted(f for f in by_folder
                         if f not in used and all(k in normalize_name(f, is_file=False) for k in keywords))
        entry = manifest[group] = {"required": gcfg.get("required", False), "folders": folders}

        # ---- folder not delivered ----
        if not folders:
            entry["status"] = "missing"
            say = logger.error if entry["required"] else logger.warning
            say(f"[{group}] no image folder with {gcfg['keywords']} in its name"
                f"{' (REQUIRED)' if entry['required'] else ' (optional)'}")
            continue
        used.update(folders)

        # ---- read every image (parallel, in batches) ----
        paths = sorted(p for f in folders for p in by_folder[f])
        logger.info(f"[{group}] {len(paths)} image(s) in {folders} - "
                    f"{'reading text with OCR' if ocr else 'OCR off'} on {settings.IMAGE_WORKERS} core(s)")

        def make_job(path):
            rel = PurePosixPath(path).relative_to(base)
            return (settings.INPUT.full_path(path), settings.INPUT.read_bytes(path),
                    lane_from_path(rel.parts[1:-1]), rel.parts[0], cfg)

        def progress(done, total):
            logger.info(f"[{group}] {done}/{total} images read")

        rows = read_images(paths, make_job, settings.IMAGE_WORKERS, progress=progress)
        if location:
            rows = [locate(r, lanes, offsets, settings.IMAGE_LOCATION_AGREE_M) for r in rows]
        df = pd.DataFrame(rows)
        first = [c for c in IMAGE_COLUMNS_FIRST if c in df.columns]
        std = [c for c in settings.IMAGE_FIELDS if c in df.columns and c not in first]
        rest = sorted(c for c in df.columns if c not in first and c not in std)
        df = df[first + std + rest]

        counts = df["status"].value_counts().to_dict()
        with_issues = int((df["issues"].fillna("") != "").sum())
        sources = df["location_source"].value_counts().to_dict() if "location_source" in df else {}
        entry.update(status="saved", images=len(df), image_status=counts, images_with_issues=with_issues,
                     location_source=sources, ocr=ocr, outputs=run.save_table(df, "images", group))
        logger.info(f"[{group}] saved {len(df)} image rows -> {run.file_name('images', group)}.parquet + .csv "
                    f"(status: {counts}, {with_issues} with issues, location: {sources})")

    # ---- image folders that belong to no group ----
    other = sorted(f for f in by_folder if f not in used)
    if other:
        logger.warning(f"image folder(s) not in the config, ignored: {other}")
        run.manifest["unexpected_image_folders"] = other


# =============================================================================
# 5. settings.EXTRACT_ALL: every workbook, every sheet, header found automatically, no header check
# =============================================================================
def extract_everything(run, folder):
    logger.warning(f"{run.deliverable}: EXTRACT_ALL is on - config ignored, every workbook and sheet is read")
    for path in _excel_files(folder):
        key = safe_name(PurePosixPath(path).stem)
        wb = ExcelWorkbook(PurePosixPath(path).name, settings.INPUT.read_bytes(path))
        sheets = {safe_name(s): {"sheet_names": [s]} for s in wb.sheet_names}
        wb.close()
        workbooks = {key: {"file_names": [PurePosixPath(path).name], "sheets": sheets}}
        run.manifest["workbooks"][key] = {"file": PurePosixPath(path).name, "status": "found", "sheets": {}}
        extract_workbook(run, {key: path}, key, workbooks)
