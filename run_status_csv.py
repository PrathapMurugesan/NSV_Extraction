"""
run_status_csv.py - run run_download.py for every UPC in the download status CSV that has
New Structure = 1, up to 3 UPCs at the same time.

The CSV has the columns  Zone, Phase, Cycle, UPC code ..., Status, New Structure, ...
("Phase  3", "Cycle 1", "N/02007/11001/HR" are fine - only the number / code is used).

    python run_status_csv.py --csv "Downloading_status(UPC Download Status).csv" --root "$SRC"
    python run_status_csv.py --csv s3://bucket/path/status.csv --root "$SRC"      CSV on S3
    python run_status_csv.py --csv status.csv --root "$SRC" --parallel 2         2 at a time
    python run_status_csv.py --csv status.csv --root "$SRC" --list               only show the UPCs
    python run_status_csv.py --csv status.csv --root "$SRC" d1 d2                any run_download option

One log per UPC:   logs/status_<timestamp>/<zone>_<phase>_<cycle>_<UPC>.log
Summary:           logs/status_<timestamp>/summary.csv   (UPC, exit code, log file)
"""
import argparse
import csv
import io
import re
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent


def read_csv_text(path):
    if path.startswith("s3://"):
        import boto3
        bucket, _, key = path[len("s3://"):].partition("/")
        data = boto3.client("s3").get_object(Bucket=bucket, Key=key)["Body"].read()
    else:
        data = Path(path).read_bytes()
    return data.decode("utf-8-sig", errors="replace")


def surveys_from_csv(path):
    """[(zone, phase, cycle, upc)] of the rows with New Structure = 1, duplicates removed."""
    reader = csv.DictReader(io.StringIO(read_csv_text(path)))
    columns = {c.strip().lower(): c for c in reader.fieldnames or []}
    upc_col = next((c for k, c in columns.items() if k.startswith("upc")), None)
    new_col = columns.get("new structure")
    missing = [n for n, c in (("Zone", columns.get("zone")), ("Phase", columns.get("phase")),
                              ("Cycle", columns.get("cycle")), ("UPC ...", upc_col),
                              ("New Structure", new_col)) if c is None]
    if missing:
        sys.exit(f"column(s) {missing} not found in {path} - columns are {reader.fieldnames}")
    out = []
    for row in reader:
        if (row.get(new_col) or "").strip() not in ("1", "1.0"):
            continue
        zone = (row[columns["zone"]] or "").strip().upper()
        phase = re.findall(r"\d+", row[columns["phase"]] or "")
        cycle = re.findall(r"\d+", row[columns["cycle"]] or "")
        upc = (row[upc_col] or "").strip()
        if not (zone and phase and cycle and upc):
            print(f"skipped (incomplete row): {row}")
            continue
        survey = (zone, phase[0], cycle[0], upc)
        if survey not in out:
            out.append(survey)
    return out


def run_one(survey, root, extra, log_dir):
    zone, phase, cycle, upc = survey
    log_path = log_dir / f"{zone}_{phase}_{cycle}_{upc.replace('/', '_')}.log"
    command = [sys.executable, str(PROJECT_ROOT / "run_download.py"), "--root", root,
               "--upc", upc, "--zone", zone, "--phase", phase, "--cycle", cycle, *extra]
    with open(log_path, "w", encoding="utf-8") as log:
        code = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT, cwd=PROJECT_ROOT).returncode
    return code, log_path


def main(args):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--csv", required=True, help="download status CSV, local path or s3://bucket/key")
    parser.add_argument("--root", required=True, help="download root, local folder or s3://bucket/prefix/")
    parser.add_argument("--parallel", type=int, default=3, help="UPCs at the same time (default 3)")
    parser.add_argument("--list", action="store_true", help="only show the UPCs, run nothing")
    o, extra = parser.parse_known_args(args)

    surveys = surveys_from_csv(o.csv)
    print(f"{len(surveys)} UPC(s) with New Structure = 1 in {o.csv}")
    for zone, phase, cycle, upc in surveys:
        print(f"  zone {zone}  phase {phase}  cycle {cycle}  {upc}")
    if o.list or not surveys:
        return 0

    log_dir = PROJECT_ROOT / "logs" / f"status_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    log_dir.mkdir(parents=True, exist_ok=True)
    results = []
    with ThreadPoolExecutor(max_workers=max(1, o.parallel)) as pool:
        jobs = {pool.submit(run_one, s, o.root, extra, log_dir): s for s in surveys}
        for number, job in enumerate(as_completed(jobs), start=1):
            zone, phase, cycle, upc = jobs[job]
            try:
                code, log_path = job.result()
            except Exception as e:
                code, log_path = f"crashed: {e}", ""
            print(f"[{number}/{len(surveys)}] {zone}/{phase}/{cycle} {upc:20} exit code {code}  {log_path}")
            results.append([zone, phase, cycle, upc, code, log_path])

    summary = log_dir / "summary.csv"
    with open(summary, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["zone", "phase", "cycle", "upc", "exit_code", "log"])
        w.writerows(sorted(results, key=lambda r: r[:4]))
    failed = [r for r in results if r[4] != 0]
    print(f"done: {len(results) - len(failed)} ok (exit 0), {len(failed)} with problems - summary: {summary}")
    return 0 if not failed else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
