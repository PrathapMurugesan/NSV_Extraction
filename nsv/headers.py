"""
headers.py - match a sheet's headers to its schema, and decide the column names to save.

Matching, strictest first. An expected header takes the first tier that finds it; exact matches are
assigned before any looser tier, so a vendor column with the exact name always wins.

    tier          example (vendor -> schema)                          accepted?
    exact         'Chainage from (km)'  -> 'Chainage from (km)'         yes
    normalized    'Scheme Name'         -> 'SchemeName'                 yes  (case, spaces, _ - . ignored)
    unit          'Chainage from (m)'   -> 'Chainage from (km)'         yes  (same name, unit of the same kind)
                  'Chainage from'       -> 'Chainage from (km)'         yes  (no unit in the vendor header)
    alias         'NH No.'              -> 'NH Number'                  yes  (listed in the schema csv, 'aliases')
    position      column CC 'Raveling Lane Area (sq. m) | L4'          yes, WARNING 'expected X, got Y'
                    -> 'Raveling Lane Area (%) @ 10 mtr | L4'  (schema column CC)
                  = the header differs, but the column is in the SAME place as in the template and holds
                    the SAME kind of values (number / text / date). Only columns no other tier claimed.
                    Different kind of values -> not a match (the sheet is rejected).
    fuzzy         'Chainage Frm (km)'   -> 'Chainage from (km)'         NO - only a hint in the log
Fuzzy is never accepted because of near-twins: 'Chainage from' / 'Chainage To', 'Severity' /
'Pothole Severity', L1 / R1 ... look alike but are different columns.

Saved column names: the schema header WITHOUT its unit ('Chainage from (km)' -> 'Chainage from').
Columns not in the schema keep the vendor header, also without its unit. The units (in the file and in
the schema) are recorded in schema.json. If removing units would give two columns the same name
(e.g. 'Cracking Lane Area (%)' and 'Cracking Lane Area (sq. m.)'), those columns keep their unit.
"""
import re
from difflib import get_close_matches

from nsv.excel import clean_text

# -----------------------------------------------------------------------------
# Units that may be split off a header. Only brackets whose WHOLE content is one of these
# (optionally starting with 'in ') are units; '(CA)', '(from)', '[dd-mm-yyyy]' stay in the name.
#   written form (lower case) -> (standard unit, kind of unit)
# -----------------------------------------------------------------------------
UNITS = {
    "km": ("km", "length"), "kms": ("km", "length"), "kilometer": ("km", "length"), "kilometre": ("km", "length"),
    "m": ("m", "length"), "mtr": ("m", "length"), "mtrs": ("m", "length"), "meter": ("m", "length"),
    "meters": ("m", "length"), "metre": ("m", "length"), "metres": ("m", "length"),
    "mm": ("mm", "length"), "cm": ("cm", "length"),
    "mm/km": ("mm/km", "roughness"), "m/km": ("m/km", "roughness"),
    "%": ("%", "percent"), "% area": ("%", "percent"), "percent": ("%", "percent"),
    "sq m": ("sq m", "area"), "sq. m.": ("sq m", "area"), "sq.m": ("sq m", "area"), "sq. m": ("sq m", "area"),
    "sqm": ("sq m", "area"), "m2": ("sq m", "area"),
}
BRACKETS = re.compile(r"\s*[\(\[]\s*([^\)\]]*?)\s*[\)\]]")


def split_unit(header):
    """'Chainage from (km)' -> ('Chainage from', 'km', 'length');  'NH Number' -> ('NH Number', None, None)"""
    text = clean_text(header)
    for m in reversed(list(BRACKETS.finditer(text))):                    # the last unit in the header counts
        inner = re.sub(r"^in\s+", "", m.group(1).lower()).strip()
        if inner in UNITS:
            unit, kind = UNITS[inner]
            return clean_text(text[:m.start()] + " " + text[m.end():]), unit, kind
    return text, None, None


def value_type(values):
    """
    The kind of values a column holds: 'number', 'text', 'date', or 'empty'.
    Numbers / dates stored as text count as number / date. The most common kind wins.
    """
    from nsv.profiler import DATE_TEXT, NUMBER_TEXT, _kind
    counts = {}
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()) or (isinstance(v, float) and v != v):
            continue
        kind = _kind(v)
        if kind == "text":
            kind = "number" if NUMBER_TEXT.match(v) else "date" if DATE_TEXT.match(v) else "text"
        if kind == "bool":
            kind = "text"
        counts[kind] = counts.get(kind, 0) + 1
    return max(counts, key=counts.get) if counts else "empty"


def _norm(text):
    """Comparison key: lower case, letters/digits/% only ('Scheme Name' == 'SchemeName' == 'SCHEME_NAME').
    '%' is kept: 'Extent %' must never equal 'Extent'."""
    return re.sub(r"[^a-z0-9%]", "", str(text).lower())


# -----------------------------------------------------------------------------
# Matching
# -----------------------------------------------------------------------------
def match_headers(found, expected, aliases=None, info=None, letters=None, values=None):
    """
    found:    vendor headers as read (multi-row headers already combined)
    expected: schema headers
    aliases:  {schema header: [other accepted names]}
    info:     {schema header: {"excel_column", "value_type"}}   (schema csv)  -> for the position tier
    letters:  Excel column of every found header; values: the values of every found column
    Returns {"columns": [...one entry per found header...], "missing": [...], "hints": {missing: closest found}}
    """
    aliases, info = aliases or {}, info or {}
    letters = letters or [None] * len(found)
    values = values or [[] for _ in found]
    f_info = [dict(zip(("base", "unit", "kind"), split_unit(h)), clean=clean_text(h)) for h in found]
    match = {}                                          # found index -> (expected header, tier)

    def candidates(name):
        e_base, e_unit, e_kind = split_unit(name)
        e_clean = clean_text(name)
        return {
            "exact": lambda f: f["clean"] == e_clean,
            "normalized": lambda f: _norm(f["clean"]) == _norm(e_clean),
            "unit": lambda f: (_norm(f["base"]) == _norm(e_base)
                               and (f["unit"] is None or e_unit is None or f["kind"] == e_kind)),
        }

    tiers = ["exact", "normalized", "unit"]
    for tier in tiers + ["alias"]:                      # tier by tier over ALL expected headers
        for e in expected:
            if e in {m[0] for m in match.values()}:
                continue
            names = aliases.get(e, []) if tier == "alias" else [e]
            for name in names:
                checks = candidates(name)
                for sub in (tiers if tier == "alias" else [tier]):
                    hits = [i for i, f in enumerate(f_info) if i not in match and checks[sub](f)]
                    if len(hits) == 1:
                        match[hits[0]] = (e, tier)
                        break
                if e in {m[0] for m in match.values()}:
                    break

    # ---- position tier: same Excel column as in the template + same kind of values ----
    position_types = {}
    by_letter = {letters[i]: i for i in range(len(found)) if letters[i]}
    for e in expected:
        if e in {m[0] for m in match.values()}:
            continue
        letter = info.get(e, {}).get("excel_column")
        i = by_letter.get(letter) if letter else None
        if i is None or i in match:
            continue
        expected_type = info[e].get("value_type") or "empty"
        found_type = value_type(values[i])
        if "empty" in (expected_type, found_type) or expected_type == found_type:
            match[i] = (e, "position")
            position_types[i] = (found_type, expected_type)

    matched = {m[0] for m in match.values()}
    missing = [e for e in expected if e not in matched]
    unmatched = [found[i] for i in range(len(found)) if i not in match]
    lower = {clean_text(h).lower(): h for h in unmatched}
    hints = {}
    for e in missing:
        close = get_close_matches(clean_text(e).lower(), list(lower), n=1, cutoff=0.6)
        if close:
            hints[e] = lower[close[0]]

    columns = []
    for i, header in enumerate(found):
        e, tier = match.get(i, (None, "not in schema"))
        e_unit = split_unit(e)[1] if e else None
        f_unit = f_info[i]["unit"]
        types = position_types.get(i)
        columns.append({
            "vendor_header": header,
            "schema_header": e,
            "match": tier,
            "excel_column": letters[i],
            "values_in_file": types[0] if types else None,          # only for position matches
            "values_in_schema": types[1] if types else None,
            "unit_in_file": f_unit,
            "unit_in_schema": e_unit,
            "unit_differs": bool(e and e_unit != f_unit),
        })
    return {"columns": columns, "missing": missing, "hints": hints}


# -----------------------------------------------------------------------------
# Final names
# -----------------------------------------------------------------------------
def final_names(columns):
    """
    Name to save for every column: schema header (or vendor header) without unit.
    If two columns would end up with the same name, those keep their full header (with unit);
    any remaining duplicates get _2, _3.
    """
    full = [c["schema_header"] or clean_text(c["vendor_header"]) for c in columns]
    short = [split_unit(h)[0] for h in full]
    counts = {}
    for name in short:
        counts[name] = counts.get(name, 0) + 1
    names, seen = [], {}
    for s, f in zip(short, full):
        name = s if counts[s] == 1 else f              # collision -> keep the unit in the name
        seen[name] = seen.get(name, 0) + 1
        names.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    return names
