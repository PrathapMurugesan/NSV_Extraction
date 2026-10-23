"""
settings.py - the ONE file you edit to change how the pipeline behaves.

What each deliverable contains (workbooks, sheets, header layout) lives in deliverables/d1..d4.py.
Expected column headers live in schemas/<deliverable>/<workbook>/<sheet>.csv.
"""
from pathlib import Path

from nsv.storage import LocalStorage, S3Storage  # noqa: F401  (S3Storage is used when you switch to S3)

PROJECT_ROOT = Path(__file__).resolve().parent


# =============================================================================
# 1. WHERE FILES COME FROM AND WHERE RESULTS GO
#    Raw data and bronze output live under the same root today (local folder "data").
#    To use S3, comment the Local lines and uncomment the S3 lines.
# =============================================================================
INPUT = LocalStorage(PROJECT_ROOT / "data")
# INPUT = S3Storage(bucket="your-nhai-bucket", prefix="", region_name="ap-south-1")

OUTPUT = LocalStorage(PROJECT_ROOT / "data")
# OUTPUT = S3Storage(bucket="your-nhai-bucket", prefix="", region_name="ap-south-1")

SCHEMAS_DIR = PROJECT_ROOT / "schemas"          # expected headers (part of the code, always local)


# =============================================================================
# 2. FOLDER LAYOUT
#    {zone} {phase} {cycle} {upc} {deliverable} {run_id} are filled in per survey and run.
#    {upc} is the UPC as used in folder names: '/' replaced by '-'  (N/08004/03001/HR -> N-08004-03001-HR)
#
#    raw:     nsv/zone=A/phase=1/cycle=1/N-08004-03001-HR/d2/<files as downloaded>
#    bronze:  bronze/nsv/zone=A/phase=1/cycle=1/N-08004-03001-HR/d2/run=20261020_101500/<extracted files>
#
#    If your real folders look different (e.g. "upc=N-08..." or "Deliverable 2"), change only these lines.
# =============================================================================
RAW_LAYOUT = "nsv/zone={zone}/phase={phase}/cycle={cycle}/{upc}/{deliverable}"
BRONZE_LAYOUT = "bronze/nsv/zone={zone}/phase={phase}/cycle={cycle}/{upc}/{deliverable}"
RUN_FOLDER = "run={run_id}"                      # one folder per extraction run, inside the bronze folder
BATCH_FOLDER = "bronze/nsv/_batches"             # summary of every loop over several surveys

ZONES = ["A", "B", "C", "D", "E"]                # allowed values - other folders are reported and ignored
PHASES = ["1", "2", "3"]
CYCLES = ["1", "2"]
DELIVERABLES = ["d1", "d2", "d3", "d4"]          # deliverable folder names inside a UPC folder


# =============================================================================
# 3. SWITCHES
# =============================================================================
CHECK_HEADERS = True             # True = a sheet whose headers differ from its schema CSV is rejected (not saved)
STOP_IF_WORKBOOK_MISSING = True  # True = if an expected workbook is not found, the deliverable is not extracted
EXTRACT_ALL = False              # True = ignore the deliverable config: read every workbook and every sheet,
                                 #        header found automatically, no header check (for a new vendor layout)


# =============================================================================
# 4. MATCHING FILE / SHEET NAMES
# =============================================================================
FUZZY_THRESHOLD = 0.85           # 0-1: how similar a file or sheet name must be to an expected name
HEADER_SCAN_ROWS = 20            # automatic header detection looks at this many top rows


# =============================================================================
# 5. DEFECT-COLUMN NULL REPORT (D2, D4) - REPORT ONLY, a sheet is never skipped
#    A defect column = its name contains one of DEFECT_KEYWORDS.
#    It is reported when completely empty, or more than NULL_LIMIT % empty.
# =============================================================================
NULL_LIMIT = 25
DEFECT_KEYWORDS = [
    "roughness", "rut", "crack", "ravelling", "raveling", "distress", "depth", "area", "severity",
    "chainage", "pothole", "defect", "lane", "carriageway", "image", "latitude", "longitude",
    "pci", "condition", "bi ", "iri",
]
NULL_IGNORE = ["remark", "structure", "limitation"]   # naturally sparse text / contract limits: not reported


# =============================================================================
# 6. PROFILE JSON
# =============================================================================
PROFILE_EXAMPLES = 3             # example values stored per column
PROFILE_MAX_DISTINCT = 1000      # distinct values are counted up to this number


# =============================================================================
# 8. SURVEY IMAGES (e.g. d4/crack_images_DHMS/L1/*.jpg)
#    Text printed on the images is read with OCR (tesseract). Tesseract is a separate program:
#      Windows: install from https://github.com/UB-Mannheim/tesseract/wiki and set TESSERACT_CMD below
#      Linux:   apt install tesseract-ocr
#    Without it, images are still listed (file name, path, lane, EXIF) and the report says OCR is missing.
# =============================================================================
import os as _os

OCR_IMAGES = True
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"             # e.g. r"C:\Program Files\Tesseract-OCR\tesseract.exe"  (None = found on PATH)
IMAGE_WORKERS = _os.cpu_count() or 1        # CPU cores used in parallel (~0.6 s per image per core)
IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")

# Where the text is, as a fraction of the image height: (name, from, to). Add e.g. ("bottom", 0.9, 1.0)
# for a vendor with a bottom banner. Text drawn on top of the pavement picture is not reliable to read.
IMAGE_TEXT_REGIONS = [("top banner", 0.0, 0.17)]

# Image location: chainage from the file name, lat/long from EXIF GPS / Defect Report / OCR (see nsv/images.py).
# OCR lat/long further than this from the chosen location -> flagged in 'issues' (correct reads agree to < 1 m).
IMAGE_LOCATION_AGREE_M = 10

# Standard name <- labels as printed by vendors (case / spaces / symbols ignored) + check.
# Labels not listed here are still extracted, as columns 'label_<their name>'.
IMAGE_FIELDS = {
    "survey_id":    {"labels": ["Survey ID", "SurveyID", "Survey Id"]},
    "description":  {"labels": ["Description", "Section", "Road"]},
    "direction":    {"labels": ["Direction"]},
    "date":         {"labels": ["Date", "Survey Date", "Captured On"], "type": "date"},
    "survey_start": {"labels": ["Survey Start", "Start"], "type": "number"},
    "frame_no":     {"labels": ["Frame No", "Frame", "Frame Number"], "type": "number"},
    "latitude":     {"labels": ["Latitude", "Lat", "GPS Lat"], "type": "number", "range": (6, 38)},
    "longitude":    {"labels": ["Longitude", "Long", "Lon", "GPS Long"], "type": "number", "range": (68, 98)},
    "altitude":     {"labels": ["Altitude", "Alt"], "type": "number"},
    "distance":     {"labels": ["Distance"], "type": "number"},
    "chainage":     {"labels": ["Chainage KM", "Chainage", "Chainage (m)", "Ch"], "type": "number"},
    "speed":        {"labels": ["Speed Km/hr", "Speed", "Speed (km/h)"], "type": "number", "range": (0, 150)},
}
