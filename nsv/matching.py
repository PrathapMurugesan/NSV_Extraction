"""
matching.py - find the expected workbooks / sheets by name, even when vendors rename them slightly.

How a name is compared:
  1. exactly the same text as a known name                       -> best possible match
  2. same after normalizing: case, spaces, _ - . (1) v1.0, month-year ('Feb-26') ignored -> score 1.0
  3. one name contains the other: 'Jun-25 Comparison as per CA Kaithal..' holds 'Comparison as per CA',
     'Defect Report.xlsx' is inside 'Defect Report_DHMS'  -> score CONTAINS_SCORE
     (several sheets like that: the one with the same name wins, see identify_names)
  4. otherwise a similarity score 0-1; at least FUZZY_THRESHOLD counts as a match
Each found name goes to the expected key it resembles MOST, so 'RIS_...' is never
claimed by 'lrms' just because the two names look 90 % alike.
"""
import re
from difflib import SequenceMatcher
from pathlib import PurePosixPath

from nsv.log import get_logger

logger = get_logger("nsv.matching")

CONTAINS_SCORE = 0.95       # a name containing the whole expected name (or the other way round)
CONTAINS_MIN_LENGTH = 7     # ... only for names of at least 7 letters/digits ('crripci'), so 'report' is not enough
# month + year that vendors put around names: 'Feb-26 Summary', 'Sep-26 Comparison ...', 'Report July 2026'
MONTH_YEAR = re.compile(r"(?<![a-z])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s_\-'.]*\d{2,4}(?!\d)")


def normalize_name(name, is_file=True):
    """'LRMS_DataTemplate v1.0 (1).xlsx' -> 'lrmsdatatemplate', 'Feb-26 Summary' -> 'summary'
    (sheet names: is_file=False)"""
    name = (PurePosixPath(name).stem if is_file else str(name)).lower()
    name = re.sub(r"\(\d+\)", "", name)                  # Windows copies: "(1)"
    name = re.sub(r"[ _\-]*v\d+(\.\d+)*", "", name)      # versions: v1.0, _v2, -v1.2.3
    stripped = MONTH_YEAR.sub("", name)
    name = stripped if re.search(r"[a-z]", stripped) else name   # a name that is only a month stays as it is
    return re.sub(r"[^a-z0-9]", "", name)                # spaces, _, -, dots, brackets


def safe_name(name):
    """'Comparison as per CA (1)' -> 'comparison_as_per_ca_1'  (for output file names)"""
    return re.sub(r"[^a-z0-9]+", "_", str(name).lower()).strip("_") or "unnamed"


def match_score(name, aliases, is_file=True):
    """Best similarity between a name and any known alias."""
    if name in aliases:
        return 2.0                                       # identical text beats look-alikes ('Cycle-1' / 'Cycle- 1')
    norm = normalize_name(name, is_file)
    return max(_similarity(norm, normalize_name(a, is_file), is_file) for a in aliases)


def _similarity(a, b, is_file):
    ratio = SequenceMatcher(None, a, b).ratio()
    short, long_ = sorted((a, b), key=len)
    if len(short) >= CONTAINS_MIN_LENGTH and short in long_:
        return max(ratio, CONTAINS_SCORE)               # vendors add road / '_DHMS' / '@100m' around the name
    return ratio


def identify_names(found, expected, threshold, label="workbook", is_file=True, quiet=False):
    """
    Map every expected key to exactly one of the found names.

    found:    names that exist (file names, or the sheet names of one workbook)
    expected: {key: [known names]}
    quiet:    True = don't log missing / unexpected (the caller logs them, e.g. after a fallback)
    Returns:  ({key: found_name}, [missing keys], [unexpected names])
    """
    # ---- step 1: every found name goes to the key it resembles most (if good enough) ----
    candidates = {key: [] for key in expected}
    unexpected = []
    for name in found:
        scores = {key: match_score(name, aliases, is_file) for key, aliases in expected.items()}
        best = max(scores, key=scores.get) if scores else None
        if best is not None and scores[best] >= threshold:
            candidates[best].append((scores[best], name))
        else:
            unexpected.append(name)

    # ---- step 2: every key needs exactly one match ----
    identified, missing = {}, []
    for key, matches in candidates.items():
        if not matches:
            missing.append(key)
            if not quiet:
                logger.error(f"[{key}] no {label} found (expected like '{expected[key][0]}')")
        elif len(matches) > 1:
            matches.sort(reverse=True)
            if matches[0][0] >= 1.0 and matches[0][0] > matches[1][0]:
                # one has the same name ('Comparison as per CA'), the others only look alike
                # ('Comparison As per CA @100m') -> the same name wins, the look-alikes are unexpected
                identified[key] = matches[0][1]
                unexpected += [n for _, n in matches[1:]]
                logger.info(f"[{key}] {label} '{matches[0][1]}' (same name - look-alikes "
                            f"{[n for _, n in matches[1:]]} left as unexpected)")
            else:
                missing.append(key)
                unexpected += [n for _, n in matches]
                logger.warning(f"[{key}] several {label}s match: {[n for _, n in matches]} - left to the caller")
        else:
            score, name = matches[0]
            identified[key] = name
            how = ("exact match" if score > 1.0 else "same name ignoring case/spaces/version" if score == 1.0
                   else f"fuzzy match {score:.2f}")
            logger.info(f"[{key}] {label} '{name}' ({how})")

    if not quiet:
        for name in unexpected:
            logger.warning(f"{label} not in the config: '{name}'")
    return identified, missing, unexpected


def classify_path(relative_path, groups):
    """
    D3: which group does a csv belong to, judging by its folder + file name?
    groups: {"ca": ["as per ca", "ca"], "irc": ["irc"]} - whole words only ('ca' does not match 'Location').
    Returns the group name, or None.
    """
    text = " " + re.sub(r"[^a-z0-9]+", " ", str(relative_path).lower()) + " "
    for group, phrases in groups.items():
        if any(f" {re.sub(r'[^a-z0-9]+', ' ', p.lower()).strip()} " in text for p in phrases):
            return group
    return None
