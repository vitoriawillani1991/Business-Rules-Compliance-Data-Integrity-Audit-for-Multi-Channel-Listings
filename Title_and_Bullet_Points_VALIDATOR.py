"""
PURPOSE
-------
Validate that already-built Titles and Bullet Points follow Business Rules templates AND that
the data inserted in place of placeholders matches expected NetSuite (Saved Search) values.

WHAT THIS SCRIPT DOES
---------------------
For each SKU row in the Titles/Bullets input:
1) Select the correct template row from the Business Rules workbook (sheet "Templates")
   using Titles column "SKU Type" == Rules column "SKU Type Updated".
2) For each text field (Amazon Title, bullets, eBay/Walmart/Webstore/Google fields):
   A) RULE (STRUCTURE) VALIDATION
      - Convert the template into a tolerant regex:
        * Literal text matched flexibly (case-insensitive, whitespace tolerant, punctuation tolerant).
        * Placeholders [LIKE THIS] captured as groups.
        * OUTER DIAMETER MM has a specialized capture to reduce regex drift.
      - Match built text against this regex to validate structure.

   B) PLACEHOLDER (DATA) VALIDATION
      - If structure matches: validate captured placeholder values normally.
      - If structure FAILS: run a FALLBACK placeholder validation that does NOT depend on captures:
        * For each placeholder present in the template, try to confirm the expected value exists
          somewhere in the built text (case-insensitive, whitespace-tolerant).
        * If the placeholder is allowed to remain literal when the Saved Search value is missing
          (Boolean placeholders, MIRROR COLOR, MIRROR FEATURES), we accept literal placeholder presence.

   Placeholder values are validated against allowed candidates from:
      - Saved Search (Inventory Item / Assembly/Bill of Materials)
      - Kit members via Parts_in_Package + Saved Search (Kit/Package) (any-match wins)
      - Position list for [Position], [Left/Right], [Driver/Passenger], and [Position for <Category>]
   Numeric validations are tolerant (units, rounding, formatting differences).
   Spec strings are normalized.

SPECIAL PLACEHOLDER BEHAVIOR
----------------------------
Ignored placeholders (completely ignored in validation):
  [2PIECE DESIGN M1], [2PIECE DESIGN M2], [2PIECE DESIGN M3], [2PIECE DESIGN],
  [SKU M1], [SKU M2], [SKU M3], [YMM], [MMY], [YMM], [MMY]

Boolean placeholders (ONLY those listed on Rules sheet "Boolean"):
  - If Saved Search value exists and is recognized as boolean (Yes/No/True/False/1/0):
      True  -> expected text must match Boolean.Yes
      False -> expected text must match Boolean.No (may be blank, meaning placeholder deleted)
  - If Saved Search value is missing/blank/NaN OR not boolean:
      It is OK for the built text to keep the literal placeholder (e.g. "[WITH BALL JOINT]").

[MIRROR COLOR]
  - If Saved Search is Black/Chrome/Gray/Satin -> expected text is that value.
  - If Saved Search is "Paint to Match" -> expected text is empty (placeholder deleted).
  - If missing -> OK for built text to keep literal placeholder.

[MIRROR FEATURES]
  - "Extendable, Heated" or "Heated" -> expected "Heated"
  - "Extendable" -> expected "Non-Heated"
  - If missing -> OK for built text to keep literal placeholder.

[Left/Right] and [Driver/Passenger]
  - Derived from Position list (or inferred from kit members if needed):
      * If any position is generic (Front/Rear without side) -> "Left or Right" / "Driver or Passenger"
      * Left only -> "Left" / "Driver"
      * Right only -> "Right" / "Passenger"
      * Left and Right -> "Left or Right" / "Driver or Passenger"

[Position] acceptance enhancement:
  - If expected positions include both Front and Rear, accept:
      * "Front and Rear"
      * "Front or Rear"
  - If ANY validated field template for a SKU uses side-carrying placeholders
    ([Left/Right] or [Driver/Passenger]), then [Position] is expected without side (Front/Rear)
    consistently across all fields for that SKU.

Dual-measure placeholders (may appear in titles as:  xx.xx" (yyy.y mm))
  These placeholders do NOT always appear in that format, but when a data mismatch is detected,
  the script tries a dual-measure interpretation before concluding it's a real data error.
  Placeholders included:
    - CALIPER INLET PORT SIZE
    - CALIPER PISTON SIZE
    - CATALYTIC CONVERTER OUTLET OUTSIDE DIAMETER
    - DRIVE SHAFT LENGTH INCHES
    - OVERALL COMPRESSED LENGTH MM
    - RADIATOR CORE HEIGHT
    - RADIATOR CORE WIDTH
    - OUTER DIAMETER MM

OUTPUT
------
An XLSX report with ONE ROW PER (SKU, FIELD) that has:
- SKU, Type, Rule Key, Field
- rule_status: OK / ERROR / NO_RULE_MATCH
- placeholder_status: OK / ERROR / SKIPPED
- placeholder_check_mode: CAPTURE / FALLBACK / SKIPPED
- rule_error: details if rule_status is ERROR (structure/template mismatch)
- placeholder_error: details if placeholder_status is ERROR (missing/mismatched placeholders)

NOTE:
- Even when rule_status is ERROR, placeholder validation still runs (FALLBACK mode).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pandas as pd

# =========================
# FILE PATHS
# =========================

BUSINESS_RULES_XLSX_PATH = "Master_File-New_Templates_202602.xlsx"
TITLES_BUILDS_FILE_PATH = "INPUT_FILE.xlsx"
SAVED_SEARCH_CSV_PATH = "Placeholders_for_Title_Validation_Script.csv"
PARTS_IN_PACKAGE_CSV_PATH = "Parts_in_Package.csv"
POSITION_LIST_CSV_PATH = "Position_complete_list.csv"
OUTPUT_REPORT_XLSX_PATH = "OUPUT_FILE.xlsx"

# =========================
# SETTINGS
# =========================

RULES_TEMPLATES_SHEET_NAME = "Templates"
RULES_BOOLEAN_SHEET_NAME = "Boolean"

COL_TITLES_SKU = "Item Name/Number"
COL_TITLES_TYPE = "Type"
COL_TITLES_RULE_TYPE = "SKU Type"

COL_TITLES_CATEGORY = None
COL_RULES_CATEGORY = None

TEXT_FIELDS_TO_VALIDATE = [
    "Amazon Title",
    "Amazon Bullet Point 1",
    "Amazon Bullet Point 2",
    "Amazon Bullet Point 3",
    "Amazon Bullet Point 4",
    "Amazon Bullet Point 5",
    "eBay Title",
    "eBay Subtitle",
    "eBay Description",
    "Walmart Title",
    "Webstore Title",
    "Google Title",
]

COL_RULES_TYPE = "SKU Type Updated"
REQUIRED_RULE_TEMPLATE_COLUMNS = TEXT_FIELDS_TO_VALIDATE.copy()

PLACEHOLDER_PATTERN = re.compile(r"\[[^\[\]]+?\]")

IGNORED_PLACEHOLDERS: Set[str] = {
    "[2PIECE DESIGN M1]",
    "[2PIECE DESIGN M2]",
    "[2PIECE DESIGN M3]",
    "[2PIECE DESIGN]",
    "[SKU M1]",
    "[SKU M2]",
    "[SKU M3]",
    "[YMM]",
    "[MMY]",
}

POSITION_SUFFIX_TOKENS = {"F", "R", "FL", "FR", "RL", "RR", "L", "R"}

PLACEHOLDER_TO_SAVEDSEARCH_COLUMN_OVERRIDES: Dict[str, str] = {
    "CALIPER PISTON SIZE": "Caliper Piston Size (IN)",
    "WHEEL STUD QUANTITY": "Wheel Stud Quantity",
    "WHEEL STUDS": "Wheel Studs",
    "BRAKE CALIPER PISTON COUNT": "Brake Caliper Piston Count",
}

ABS_TOLERANCE_DEFAULT = 0.08
ABS_TOLERANCE_MM = 0.12
ABS_TOLERANCE_IN = 0.06

COL_PIP_PACKAGE_NAME_CANDIDATES = ["Package Name", "Package", "SKU", "Name"]
COL_PIP_MEMBERS = "Members"
COL_PIP_WMS_CATEGORY = "WMS Category"

PH_MIRROR_COLOR = "[MIRROR COLOR]"
PH_MIRROR_FEATURES = "[MIRROR FEATURES]"
PH_LEFT_RIGHT = "[Left/Right]"
PH_DRIVER_PASSENGER = "[Driver/Passenger]"

SIDE_CARRYING_PLACEHOLDERS: Set[str] = {PH_LEFT_RIGHT, PH_DRIVER_PASSENGER}

DUAL_MEASURE_PLACEHOLDERS: Set[str] = {
    "OUTER DIAMETER MM",
    "CALIPER INLET PORT SIZE",
    "CALIPER PISTON SIZE",
    "CATALYTIC CONVERTER OUTLET OUTSIDE DIAMETER",
    "DRIVE SHAFT LENGTH INCHES",
    "OVERALL COMPRESSED LENGTH MM",
    "RADIATOR CORE HEIGHT",
    "RADIATOR CORE WIDTH",
}

# =========================
# DATA MODELS
# =========================

@dataclass(frozen=True)
class PlaceholderCapture:
    placeholder: str
    observed_text: str


@dataclass
class FieldValidationResult:
    sku: str
    field_name: str
    built_text: str
    template_text: str
    structure_match: bool
    missing_placeholders: List[str]
    mismatched_placeholders: List[str]
    ignored_placeholders_seen: List[str]


# =========================
# NORMALIZATION HELPERS
# =========================

def norm_space_case(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def normalize_sku_key(s: object) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    x = str(s)

    x = x.replace("\ufeff", "")
    x = x.replace("\u200b", "")
    x = x.replace("\u200c", "")
    x = x.replace("\u200d", "")
    x = x.replace("\xa0", " ")
    x = x.replace("\t", " ")
    x = x.replace("\r", " ").replace("\n", " ")

    x = re.sub(r"\s+", " ", x).strip()

    if x.startswith("'") and len(x) > 1:
        x = x[1:].strip()

    if re.fullmatch(r"[A-Za-z0-9\-]+\.0", x):
        x = x[:-2]

    return x.upper()


def norm_placeholder_token(ph: str) -> str:
    inner = ph.strip()[1:-1] if ph.startswith("[") and ph.endswith("]") else ph.strip()
    inner = re.sub(r"\s+", " ", inner.strip())
    return inner.upper()


def strip_positional_suffix(placeholder_inner_upper: str) -> str:
    parts = placeholder_inner_upper.split()
    if len(parts) >= 2 and parts[-1] in POSITION_SUFFIX_TOKENS:
        return " ".join(parts[:-1])
    return placeholder_inner_upper


def canonical_spec_string(s: str) -> str:
    x = str(s).strip().lower()
    x = x.replace("×", "x")
    x = re.sub(r"\s+", "", x)
    return x


def parse_numeric_units(s: str) -> Dict[str, List[float]]:
    text = str(s).lower()
    results: Dict[str, List[float]] = {"in": [], "mm": []}

    for m in re.finditer(r"(-?\d+(?:\.\d+)?)\s*(?:\"|in\.?\b|inch(?:es)?\b)", text):
        results["in"].append(float(m.group(1)))

    for m in re.finditer(r"(-?\d+(?:\.\d+)?)\s*mm\b", text):
        results["mm"].append(float(m.group(1)))

    if not results["in"] and not results["mm"]:
        bare = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?", text)]
        results["bare"] = bare

    return results


def floats_match(a: float, b: float, tol: float) -> bool:
    return abs(a - b) <= tol


def numeric_match_flexible(observed: str, expected_candidates: Sequence[str]) -> bool:
    obs = parse_numeric_units(observed)

    for cand in expected_candidates:
        exp = parse_numeric_units(cand)

        for unit in ("in", "mm"):
            if obs.get(unit) and exp.get(unit):
                tol = ABS_TOLERANCE_IN if unit == "in" else ABS_TOLERANCE_MM
                for ov in obs[unit]:
                    for ev in exp[unit]:
                        if floats_match(ov, ev, tol):
                            return True

        obs_bare = obs.get("bare", [])
        exp_bare = exp.get("bare", [])

        if obs_bare:
            for ov in obs_bare:
                for ev in exp_bare:
                    if floats_match(ov, ev, ABS_TOLERANCE_DEFAULT):
                        return True
                for unit in ("in", "mm"):
                    for ev in exp.get(unit, []):
                        tol = ABS_TOLERANCE_IN if unit == "in" else ABS_TOLERANCE_MM
                        if floats_match(ov, ev, tol):
                            return True

        if exp_bare:
            for ev in exp_bare:
                for unit in ("in", "mm"):
                    for ov in obs.get(unit, []):
                        tol = ABS_TOLERANCE_IN if unit == "in" else ABS_TOLERANCE_MM
                        if floats_match(ov, ev, tol):
                            return True

    return False


def is_likely_spec_string(s: str) -> bool:
    t = str(s).lower()
    return ("/" in t and "x" in t) or ("×" in t)


def normalize_bool_value(val: object) -> Optional[bool]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None

    s = str(val).strip().lower()
    if s == "" or s == "nan":
        return None

    if s in {"yes", "y", "true", "t", "1"}:
        return True
    if s in {"no", "n", "false", "f", "0"}:
        return False

    return None


def placeholder_literal_equals_observed(placeholder: str, observed: str) -> bool:
    return norm_space_case(placeholder) == norm_space_case(observed)


# =========================
# DUAL MEASURE (in + mm) HELPERS
# =========================

RX_SAVED_DUAL = re.compile(r"(?i)\s*(-?\d+(?:\.\d+)?)\s*in\.?\s*/\s*(-?\d+(?:\.\d+)?)\s*mm\s*")
RX_TITLE_DUAL = re.compile(r'(?i)\s*(-?\d+(?:\.\d+)?)\s*"\s*\(\s*(-?\d+(?:\.\d+)?)\s*mm\s*\)?\s*')


def parse_saved_dual_measure(s: str) -> Optional[Tuple[float, float]]:
    m = RX_SAVED_DUAL.search(str(s))
    if not m:
        return None
    return float(m.group(1)), float(m.group(2))


def parse_title_dual_measure(s: str) -> Optional[Tuple[float, float]]:
    m = RX_TITLE_DUAL.search(str(s))
    if not m:
        return None
    return float(m.group(1)), float(m.group(2))


def dual_measure_match(observed: str, expected_candidates: Sequence[str]) -> bool:
    obs_pair = parse_title_dual_measure(observed)
    if not obs_pair:
        return False

    obs_in, obs_mm = obs_pair

    for cand in expected_candidates:
        exp_pair = parse_saved_dual_measure(cand)
        if not exp_pair:
            continue
        exp_in, exp_mm = exp_pair

        if floats_match(obs_in, exp_in, ABS_TOLERANCE_IN) and floats_match(obs_mm, exp_mm, ABS_TOLERANCE_MM):
            return True

    return False


# =========================
# LOADING HELPERS
# =========================

def read_table(path: str | Path) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")

    suf = p.suffix.lower()

    if suf == ".csv":
        return pd.read_csv(p, low_memory=False)

    if suf in (".xlsx", ".xls"):
        try:
            return pd.read_excel(p, engine="openpyxl")
        except Exception:
            return pd.read_csv(p, low_memory=False)

    try:
        return pd.read_excel(p, engine="openpyxl")
    except Exception:
        return pd.read_csv(p, low_memory=False)


def read_rules_templates(xlsx_path: str | Path) -> pd.DataFrame:
    p = Path(xlsx_path)
    if not p.exists():
        raise FileNotFoundError(f"Rules file not found: {p}")
    return pd.read_excel(p, engine="openpyxl", sheet_name=RULES_TEMPLATES_SHEET_NAME)


def read_rules_boolean_sheet(xlsx_path: str | Path) -> pd.DataFrame:
    p = Path(xlsx_path)
    if not p.exists():
        raise FileNotFoundError(f"Rules file not found: {p}")
    return pd.read_excel(p, engine="openpyxl", sheet_name=RULES_BOOLEAN_SHEET_NAME)


def build_boolean_placeholder_map(boolean_df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    col_map = {norm_space_case(c): c for c in boolean_df.columns}

    for r in ["placeholder", "yes", "no"]:
        if r not in col_map:
            raise ValueError(
                "Boolean sheet must contain columns: Placeholder, Yes, No. "
                f"Missing: '{r}'. Found: {list(boolean_df.columns)}"
            )

    c_placeholder = col_map["placeholder"]
    c_yes = col_map["yes"]
    c_no = col_map["no"]

    out: Dict[str, Dict[str, str]] = {}
    for _, r in boolean_df.iterrows():
        ph = r.get(c_placeholder, "")
        if pd.isna(ph):
            continue
        ph_str = str(ph).strip()
        if not ph_str or ph_str.lower() == "nan":
            continue

        ph_inner = norm_placeholder_token(ph_str)

        yes_text = "" if pd.isna(r.get(c_yes)) else str(r.get(c_yes)).strip()
        no_text = "" if pd.isna(r.get(c_no)) else str(r.get(c_no)).strip()

        out[ph_inner] = {"yes": yes_text, "no": no_text}

    return out


# =========================
# RULE SELECTION + TEMPLATE MATCHING
# =========================

def choose_rule_row(rules_df: pd.DataFrame, row: pd.Series) -> Optional[pd.Series]:
    if COL_TITLES_RULE_TYPE not in row.index:
        return None

    rule_key = str(row[COL_TITLES_RULE_TYPE]).strip()
    if not rule_key or rule_key.lower() == "nan":
        return None

    mask = rules_df[COL_RULES_TYPE].astype(str).str.strip().str.lower().eq(rule_key.lower())

    if (
        COL_TITLES_CATEGORY
        and COL_RULES_CATEGORY
        and COL_TITLES_CATEGORY in row.index
        and COL_RULES_CATEGORY in rules_df.columns
    ):
        cat_val = str(row[COL_TITLES_CATEGORY]).strip()
        mask = mask & rules_df[COL_RULES_CATEGORY].astype(str).str.strip().str.lower().eq(cat_val.lower())

    matches = rules_df[mask]
    if matches.empty:
        return None

    return matches.iloc[0]


def _normalize_unicode_punct(s: str) -> str:
    if s is None:
        return ""
    x = str(s)
    x = x.replace("“", '"').replace("”", '"').replace("„", '"')
    x = x.replace("’", "'").replace("‘", "'")
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    x = x.replace("\u00A0", " ")
    return x


def _literal_to_flexible_regex(literal: str) -> str:
    lit = _normalize_unicode_punct(literal)
    out = []
    for ch in lit:
        if ch.isalnum():
            out.append(re.escape(ch))
        elif ch.isspace():
            out.append(r"\s+")
        else:
            out.append(r"\W*")
    return "".join(out)


def _capture_regex_for_placeholder(ph_inner_upper: str) -> str:
    base = strip_positional_suffix(ph_inner_upper)

    if base in DUAL_MEASURE_PLACEHOLDERS:
        dual = r'\s*-?\d+(?:\.\d+)?\s*"\s*\(\s*-?\d+(?:\.\d+)?\s*mm\s*\)?\s*'
        return rf"({dual}|.*?)"

    return r"(.*?)"


def build_regex_from_template(template: str) -> Tuple[re.Pattern, List[str]]:
    template = _normalize_unicode_punct(template or "")
    placeholders = PLACEHOLDER_PATTERN.findall(template)

    parts: List[str] = []
    last = 0

    for ph in placeholders:
        start = template.find(ph, last)
        literal = template[last:start]
        parts.append(_literal_to_flexible_regex(literal))

        ph_inner_upper = norm_placeholder_token(ph)
        parts.append(_capture_regex_for_placeholder(ph_inner_upper))

        last = start + len(ph)

    parts.append(_literal_to_flexible_regex(template[last:]))

    pattern_str = r"^\s*" + "".join(parts) + r"\s*$"
    return re.compile(pattern_str, flags=re.IGNORECASE | re.DOTALL), placeholders


def extract_captures(template: str, built: str) -> Tuple[bool, List[PlaceholderCapture], List[str]]:
    rx, placeholders = build_regex_from_template(template)
    m = rx.match(built or "")
    if not m:
        return False, [], placeholders

    captures: List[PlaceholderCapture] = []
    for ph, g in zip(placeholders, m.groups()):
        captures.append(PlaceholderCapture(placeholder=ph, observed_text=str(g).strip()))
    return True, captures, placeholders


# =========================
# LOOKUPS: SAVED SEARCH / KITS / POSITION LIST
# =========================

def build_saved_search_index(df: pd.DataFrame) -> Tuple[Dict[str, pd.Series], Dict[str, str]]:
    sku_col = None
    for c in ["Name", "SKU", "Item", "itemid"]:
        if c in df.columns:
            sku_col = c
            break
    if sku_col is None:
        raise ValueError("Saved search file must contain a SKU column (e.g., 'Name' or 'SKU').")

    idx: Dict[str, pd.Series] = {}
    for _, r in df.iterrows():
        sku_key = normalize_sku_key(r.get(sku_col))
        if sku_key:
            idx[sku_key] = r

    norm_map: Dict[str, str] = {}
    for c in df.columns:
        norm_map[norm_space_case(c)] = c

    return idx, norm_map


def placeholder_to_savedsearch_column(placeholder_inner_upper: str, savedsearch_norm_map: Dict[str, str]) -> Optional[str]:
    base_upper = strip_positional_suffix(placeholder_inner_upper)

    if base_upper in PLACEHOLDER_TO_SAVEDSEARCH_COLUMN_OVERRIDES:
        return PLACEHOLDER_TO_SAVEDSEARCH_COLUMN_OVERRIDES[base_upper]

    candidates_norm = {
        norm_space_case(base_upper),
        norm_space_case(base_upper.title()),
        norm_space_case(base_upper.replace("  ", " ")),
    }
    for cn in candidates_norm:
        if cn in savedsearch_norm_map:
            return savedsearch_norm_map[cn]

    return None


def parse_members_list(members_str: str) -> List[str]:
    if not members_str or str(members_str).lower() == "nan":
        return []
    out: List[str] = []
    for token in str(members_str).split(","):
        token = token.strip()
        if not token:
            continue
        token = re.sub(r"\(\d+\)\s*$", "", token).strip()
        if token:
            out.append(token)
    return out


def parse_categories_list(categories_str: str) -> List[str]:
    if not categories_str or str(categories_str).lower() == "nan":
        return []
    return [c.strip() for c in str(categories_str).split(",") if c.strip()]


def build_parts_in_package_index(df: pd.DataFrame) -> Dict[str, pd.Series]:
    key_col = None
    for c in COL_PIP_PACKAGE_NAME_CANDIDATES:
        if c in df.columns:
            key_col = c
            break
    if key_col is None:
        raise ValueError("Parts_in_Package file must contain a Package Name column (or similar).")

    idx: Dict[str, pd.Series] = {}
    for _, r in df.iterrows():
        k = normalize_sku_key(r.get(key_col))
        if k:
            idx[k] = r
    return idx


def build_position_index(df: pd.DataFrame) -> Dict[str, List[Dict[str, str]]]:
    required = {"SKU", "Position"}
    if not required.issubset(set(df.columns)):
        raise ValueError("Position_List must contain columns: SKU, Position")

    idx: Dict[str, List[Dict[str, str]]] = {}
    for _, r in df.iterrows():
        sku_key = normalize_sku_key(r.get("SKU"))
        if not sku_key:
            continue

        rec = {"Position": str(r.get("Position", "")).strip()}
        idx.setdefault(sku_key, []).append(rec)

    return idx


def split_positions(position_cell: str) -> List[str]:
    if not position_cell or str(position_cell).lower() == "nan":
        return []
    return [p.strip() for p in str(position_cell).split(",") if p.strip()]


def get_saved_row(saved_idx: Dict[str, pd.Series], sku: str) -> Optional[pd.Series]:
    sku_key = normalize_sku_key(sku)
    if not sku_key:
        return None
    return saved_idx.get(sku_key)


def get_parts_row(parts_idx: Dict[str, pd.Series], kit_sku: str) -> Optional[pd.Series]:
    sku_key = normalize_sku_key(kit_sku)
    if not sku_key:
        return None
    return parts_idx.get(sku_key)


# =========================
# EXPECTED VALUE COLLECTION
# =========================

def get_expected_candidates_inventory_or_assembly(
    sku: str,
    placeholder_inner_upper: str,
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
) -> Tuple[List[str], Optional[str]]:
    row = get_saved_row(saved_idx, sku)
    if row is None:
        return [], None

    col = placeholder_to_savedsearch_column(placeholder_inner_upper, saved_norm_map)
    if not col or col not in row.index:
        return [], col

    val = row[col]
    if pd.isna(val) or str(val).strip() == "":
        return [], col

    return [str(val).strip()], col


def get_expected_candidates_kit_any_member(
    kit_sku: str,
    placeholder_inner_upper: str,
    parts_idx: Dict[str, pd.Series],
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
) -> Tuple[List[str], Optional[str]]:
    pkg_row = get_parts_row(parts_idx, kit_sku)
    if pkg_row is None:
        return [], None

    members = parse_members_list(str(pkg_row.get(COL_PIP_MEMBERS, "")))
    if not members:
        return [], None

    col = placeholder_to_savedsearch_column(placeholder_inner_upper, saved_norm_map)
    if not col:
        return [], None

    candidates: List[str] = []
    for msku in members:
        mrow = get_saved_row(saved_idx, msku)
        if mrow is None or col not in mrow.index:
            continue
        v = mrow[col]
        if pd.isna(v):
            continue
        vs = str(v).strip()
        if vs:
            candidates.append(vs)

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in candidates:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x)

    return uniq, col


# =========================
# POSITION / SIDE EXPECTATIONS
# =========================

def normalize_position_label(p: str) -> str:
    s = norm_space_case(p)
    s = s.replace("&", "and")
    s = re.sub(r"\s+", " ", s).strip()

    mapping = {
        "front": "Front",
        "rear": "Rear",
        "front and rear": "Front and Rear",
        "front/rear": "Front and Rear",
        "front or rear": "Front or Rear",
    }
    return mapping.get(s, p.strip())


def strip_side_from_position_label(p: str) -> str:
    s = re.sub(r"\s+", " ", str(p).strip())
    s_low = s.lower()
    s_low = re.sub(r"\s+(left|right)\s*$", "", s_low).strip()
    return normalize_position_label(s_low)


def _collect_positions_for_sku(sku: str, pos_idx: Dict[str, List[Dict[str, str]]]) -> List[str]:
    sku_key = normalize_sku_key(sku)
    recs = pos_idx.get(sku_key)
    if not recs:
        return []
    out: List[str] = []
    for rec in recs:
        out.extend(split_positions(rec.get("Position", "")))
    return [x.strip() for x in out if str(x).strip()]


def _collect_positions_for_kit_members(kit_sku: str, parts_idx: Dict[str, pd.Series], pos_idx: Dict[str, List[Dict[str, str]]]) -> List[str]:
    pkg_row = get_parts_row(parts_idx, kit_sku)
    if pkg_row is None:
        return []
    members = parse_members_list(str(pkg_row.get(COL_PIP_MEMBERS, "")))
    if not members:
        return []
    out: List[str] = []
    for msku in members:
        out.extend(_collect_positions_for_sku(msku, pos_idx))
    return out


def combine_positions_to_title_candidates(positions: List[str]) -> List[str]:
    normed = {norm_space_case(normalize_position_label(p)) for p in positions if str(p).strip()}
    if not normed:
        return []

    has_front = "front" in normed
    has_rear = "rear" in normed

    if has_front and has_rear:
        return ["Front and Rear", "Front or Rear"]

    out: List[str] = []
    seen: Set[str] = set()
    for p in positions:
        pp = normalize_position_label(p)
        k = norm_space_case(pp)
        if k and k not in seen:
            seen.add(k)
            out.append(pp)
    return out


def expected_positions_for_sku(
    sku: str,
    sku_type: str,
    pos_idx: Dict[str, List[Dict[str, str]]],
    parts_idx: Dict[str, pd.Series],
    strip_sides: bool,
) -> List[str]:
    raw = _collect_positions_for_sku(sku, pos_idx)
    if not raw and sku_type.strip() == "Kit/Package":
        raw = _collect_positions_for_kit_members(sku, parts_idx, pos_idx)

    if not raw:
        return []

    if strip_sides:
        base = [strip_side_from_position_label(p) for p in raw]
    else:
        base = [normalize_position_label(p) for p in raw]

    candidates = combine_positions_to_title_candidates(base)

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in candidates:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x.strip())
    return uniq


def expected_positions_for_kit_member_category(
    kit_sku: str,
    category_in_placeholder: str,
    parts_idx: Dict[str, pd.Series],
    pos_idx: Dict[str, List[Dict[str, str]]],
    strip_sides: bool,
) -> List[str]:
    pkg_row = get_parts_row(parts_idx, kit_sku)
    if pkg_row is None:
        return []

    members = parse_members_list(str(pkg_row.get(COL_PIP_MEMBERS, "")))
    cats = parse_categories_list(str(pkg_row.get(COL_PIP_WMS_CATEGORY, "")))

    if not members or not cats or len(members) != len(cats):
        return []

    target = norm_space_case(category_in_placeholder)

    raw: List[str] = []
    for msku, mcat in zip(members, cats):
        if norm_space_case(mcat) != target:
            continue
        raw.extend(_collect_positions_for_sku(msku, pos_idx))

    if not raw:
        return []

    if strip_sides:
        base = [strip_side_from_position_label(p) for p in raw]
    else:
        base = [normalize_position_label(p) for p in raw]

    candidates = combine_positions_to_title_candidates(base)

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in candidates:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x.strip())
    return uniq


def _position_side_token(p: str) -> Optional[str]:
    s = norm_space_case(p)
    if "left" in s:
        return "left"
    if "right" in s:
        return "right"
    return None


def expected_left_right_from_positions(raw_positions: List[str]) -> List[str]:
    if not raw_positions:
        return []

    sides: Set[str] = set()
    has_generic = False

    for p in raw_positions:
        ps = norm_space_case(p)
        side = _position_side_token(ps)
        if side:
            sides.add(side)
            continue
        if "front" in ps or "rear" in ps:
            has_generic = True

    if has_generic:
        return ["Left or Right"]
    if sides == {"left"}:
        return ["Left"]
    if sides == {"right"}:
        return ["Right"]
    if "left" in sides and "right" in sides:
        return ["Left or Right"]

    return ["Left or Right"]


def expected_driver_passenger_from_positions(raw_positions: List[str]) -> List[str]:
    lr = expected_left_right_from_positions(raw_positions)
    if not lr:
        return []
    mapping = {"left": "Driver", "right": "Passenger", "left or right": "Driver or Passenger"}
    out: List[str] = []
    for x in lr:
        out.append(mapping.get(norm_space_case(x), "Driver or Passenger"))
    return out


# =========================
# VALIDATION HELPERS
# =========================

def is_position_placeholder(placeholder: str) -> bool:
    return placeholder.strip() == "[Position]"


def is_left_right_placeholder(placeholder: str) -> bool:
    return placeholder.strip() == PH_LEFT_RIGHT


def is_driver_passenger_placeholder(placeholder: str) -> bool:
    return placeholder.strip() == PH_DRIVER_PASSENGER


def is_position_for_category_placeholder(placeholder: str) -> bool:
    p = placeholder.strip()
    return p.lower().startswith("[position for ") and p.endswith("]")


def extract_category_from_position_for(placeholder: str) -> str:
    inner = placeholder.strip()[1:-1]
    m = re.match(r"(?i)position\s+for\s+(.+)$", inner.strip())
    return m.group(1).strip() if m else ""


def validate_observed_against_candidates(observed: str, expected_candidates: List[str]) -> bool:
    obs = "" if observed is None else str(observed).strip()

    if any(str(c).strip() == "" for c in expected_candidates):
        if obs == "":
            return True

    expected_non_empty = [str(c).strip() for c in expected_candidates if str(c).strip() != ""]
    if not expected_non_empty:
        return obs == ""

    obs_has_digit = bool(re.search(r"\d", obs))

    if is_likely_spec_string(obs) or any(is_likely_spec_string(c) for c in expected_non_empty):
        obs_c = canonical_spec_string(obs)
        return any(obs_c == canonical_spec_string(c) for c in expected_non_empty)

    if obs_has_digit or any(re.search(r"\d", str(c)) for c in expected_non_empty):
        if numeric_match_flexible(obs, expected_non_empty):
            return True

    o = norm_space_case(obs)
    return any(o == norm_space_case(c) for c in expected_non_empty)


def validate_observed_with_dual_fallback(
    ph_inner_upper: str,
    observed: str,
    expected_candidates: List[str],
) -> bool:
    if validate_observed_against_candidates(observed, expected_candidates):
        return True

    base_upper = strip_positional_suffix(ph_inner_upper)
    if base_upper in DUAL_MEASURE_PLACEHOLDERS:
        if dual_measure_match(observed, expected_candidates):
            return True

    return False


def _normalize_for_contains_search(s: str) -> str:
    x = _normalize_unicode_punct(s or "")
    x = re.sub(r"\s+", " ", x).strip().lower()
    return x


def _contains_fuzzy(haystack: str, needle: str) -> bool:
    n = _normalize_for_contains_search(needle)
    if n == "":
        return True
    h = _normalize_for_contains_search(haystack)
    return n in h


# =========================
# EXPECTED TEXTS FOR SPECIALS
# =========================

def expected_texts_for_boolean_placeholder(
    sku: str,
    sku_type: str,
    placeholder_inner_upper: str,
    boolean_map: Dict[str, Dict[str, str]],
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
) -> Tuple[List[str], Optional[str], bool]:
    yes_no = boolean_map.get(placeholder_inner_upper)
    if not yes_no:
        return [], None, False

    if sku_type.strip() in ("Inventory Item", "Assembly/Bill of Materials"):
        raw_vals, mapped_col = get_expected_candidates_inventory_or_assembly(
            sku=sku,
            placeholder_inner_upper=placeholder_inner_upper,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    elif sku_type.strip() == "Kit/Package":
        raw_vals, mapped_col = get_expected_candidates_kit_any_member(
            kit_sku=sku,
            placeholder_inner_upper=placeholder_inner_upper,
            parts_idx=parts_idx,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    else:
        raw_vals, mapped_col = [], None

    if not raw_vals:
        return [], mapped_col, True

    flags: Set[bool] = set()
    for v in raw_vals:
        b = normalize_bool_value(v)
        if b is None:
            continue
        flags.add(b)

    if not flags:
        return [], mapped_col, True

    expected: List[str] = []
    if True in flags:
        expected.append(yes_no.get("yes", "").strip())
    if False in flags:
        expected.append(yes_no.get("no", "").strip())

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in expected:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x)

    return uniq, mapped_col, False


def expected_texts_for_mirror_color(
    sku: str,
    sku_type: str,
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
) -> Tuple[List[str], Optional[str], bool]:
    ph_inner = "MIRROR COLOR"
    if sku_type.strip() in ("Inventory Item", "Assembly/Bill of Materials"):
        raw_vals, mapped_col = get_expected_candidates_inventory_or_assembly(
            sku=sku,
            placeholder_inner_upper=ph_inner,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    elif sku_type.strip() == "Kit/Package":
        raw_vals, mapped_col = get_expected_candidates_kit_any_member(
            kit_sku=sku,
            placeholder_inner_upper=ph_inner,
            parts_idx=parts_idx,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    else:
        raw_vals, mapped_col = [], None

    if not raw_vals:
        return [], mapped_col, True

    expected: List[str] = []
    for v in raw_vals:
        sv = str(v).strip()
        if not sv:
            continue
        if norm_space_case(sv) == "paint to match":
            expected.append("")
        else:
            expected.append(sv)

    if not expected:
        return [], mapped_col, True

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in expected:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x)

    return uniq, mapped_col, False


def expected_texts_for_mirror_features(
    sku: str,
    sku_type: str,
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
) -> Tuple[List[str], Optional[str], bool]:
    ph_inner = "MIRROR FEATURES"
    if sku_type.strip() in ("Inventory Item", "Assembly/Bill of Materials"):
        raw_vals, mapped_col = get_expected_candidates_inventory_or_assembly(
            sku=sku,
            placeholder_inner_upper=ph_inner,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    elif sku_type.strip() == "Kit/Package":
        raw_vals, mapped_col = get_expected_candidates_kit_any_member(
            kit_sku=sku,
            placeholder_inner_upper=ph_inner,
            parts_idx=parts_idx,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    else:
        raw_vals, mapped_col = [], None

    if not raw_vals:
        return [], mapped_col, True

    expected: List[str] = []
    for v in raw_vals:
        sv = str(v).strip()
        if not sv:
            continue
        s = norm_space_case(sv)
        if s in {"extendable, heated", "heated"}:
            expected.append("Heated")
        elif s == "extendable":
            expected.append("Non-Heated")
        else:
            expected.append(sv)

    if not expected:
        return [], mapped_col, True

    seen: Set[str] = set()
    uniq: List[str] = []
    for x in expected:
        nx = norm_space_case(x)
        if nx not in seen:
            seen.add(nx)
            uniq.append(x)

    return uniq, mapped_col, False


# =========================
# TEMPLATE ANALYSIS HELPERS
# =========================

def placeholders_in_text(template_text: object) -> List[str]:
    templ = "" if template_text is None or (isinstance(template_text, float) and pd.isna(template_text)) else str(template_text)
    return [p.strip() for p in PLACEHOLDER_PATTERN.findall(templ)]


def has_any_side_carrying_placeholder(rule_row: pd.Series, fields_validated: List[str]) -> bool:
    for field in fields_validated:
        templ_val = rule_row.get(field, "")
        phs = set(placeholders_in_text(templ_val))
        if phs.intersection(SIDE_CARRYING_PLACEHOLDERS):
            return True
    return False


def consecutive_placeholder_runs(placeholders: List[str], template_text: str) -> List[List[int]]:
    """
    Returns index runs (lists of indices into 'placeholders') where placeholders are consecutive
    in the template with only whitespace between them.
    """
    templ = _normalize_unicode_punct(template_text or "")
    raw_phs = PLACEHOLDER_PATTERN.findall(templ)
    if not raw_phs:
        return []

    runs: List[List[int]] = []
    last = 0
    indices: List[Tuple[int, int]] = []

    for ph in raw_phs:
        start = templ.find(ph, last)
        end = start + len(ph)
        indices.append((start, end))
        last = end

    i = 0
    while i < len(indices) - 1:
        j = i
        run = [i]
        while j < len(indices) - 1:
            end_cur = indices[j][1]
            start_next = indices[j + 1][0]
            between = templ[end_cur:start_next]
            if between.strip() == "":
                run.append(j + 1)
                j += 1
            else:
                break
        if len(run) >= 2:
            runs.append(run)
        i = j + 1

    return runs


def tokenize_for_split(s: str) -> List[str]:
    return [t for t in re.split(r"\s+", str(s).strip()) if t]


def join_tokens(tokens: List[str]) -> str:
    return " ".join(tokens).strip()


def resolve_consecutive_run_by_expected(
    run_indices: List[int],
    captures: List[PlaceholderCapture],
    expected_map: Dict[int, List[str]],
    placeholder_inner_map: Dict[int, str],
) -> bool:
    """
    Repartition text across a consecutive placeholder run to improve validation outcomes.

    This is a best-effort optimizer:
      - It evaluates possible token splits and chooses the one that maximizes the number of
        placeholders that validate successfully.
      - It will apply an improved split even if not all placeholders can be satisfied
        (e.g., when one placeholder is inherently inconsistent with the built text).

    Supports runs of length 2 or 3.
    Returns True if it updated captures with a better split.
    """
    if len(run_indices) < 2 or len(run_indices) > 3:
        return False

    texts = [captures[i].observed_text for i in run_indices]
    combined_tokens = tokenize_for_split(" ".join([t for t in texts if t]))
    if not combined_tokens:
        return False

    def score_split(parts: List[str]) -> Tuple[int, int, int]:
        """
        Returns a tuple score used for comparison (higher is better):
          1) number of placeholders that validate (primary)
          2) total matched length among placeholders that validate (tie-breaker)
          3) negative total length among placeholders that fail (tie-breaker)
        """
        ok_count = 0
        ok_len_sum = 0
        fail_len_sum = 0

        for idx_in_run, obs in enumerate(parts):
            cap_idx = run_indices[idx_in_run]
            expected = expected_map.get(cap_idx, [])
            ph_inner = placeholder_inner_map.get(cap_idx, "")

            if not expected:
                # Unconstrained (missing data or intentionally skipped). Do not penalize or reward.
                continue

            if validate_observed_with_dual_fallback(ph_inner, obs, expected):
                ok_count += 1
                ok_len_sum += len(obs)
            else:
                fail_len_sum += len(obs)

        return ok_count, ok_len_sum, -fail_len_sum

    def current_parts() -> List[str]:
        return [captures[i].observed_text for i in run_indices]

    best_parts = current_parts()
    best_score = score_split(best_parts)

    if len(run_indices) == 2:
        for k in range(0, len(combined_tokens) + 1):
            a = join_tokens(combined_tokens[:k])
            b = join_tokens(combined_tokens[k:])
            parts = [a, b]
            sc = score_split(parts)
            if sc > best_score:
                best_score = sc
                best_parts = parts

    else:  # len == 3
        for i in range(0, len(combined_tokens) + 1):
            for j in range(i, len(combined_tokens) + 1):
                a = join_tokens(combined_tokens[:i])
                b = join_tokens(combined_tokens[i:j])
                c = join_tokens(combined_tokens[j:])
                parts = [a, b, c]
                sc = score_split(parts)
                if sc > best_score:
                    best_score = sc
                    best_parts = parts

    if best_parts == current_parts():
        return False

    for idx_in_run, new_text in enumerate(best_parts):
        cap_idx = run_indices[idx_in_run]
        captures[cap_idx] = PlaceholderCapture(captures[cap_idx].placeholder, new_text)

    return True


# =========================
# EXPECTED CANDIDATES FOR PLACEHOLDERS
# =========================

def _expected_candidates_for_placeholder(
    sku: str,
    sku_type: str,
    ph: str,
    strip_sides_for_position: bool,
    boolean_map: Dict[str, Dict[str, str]],
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
    pos_idx: Dict[str, List[Dict[str, str]]],
) -> Tuple[List[str], str, bool]:
    ph_stripped = ph.strip()

    if ph_stripped in IGNORED_PLACEHOLDERS:
        return [], "IGNORED", True

    if is_left_right_placeholder(ph_stripped):
        raw = _collect_positions_for_sku(sku, pos_idx)
        if not raw and sku_type.strip() == "Kit/Package":
            raw = _collect_positions_for_kit_members(sku, parts_idx, pos_idx)
        return expected_left_right_from_positions(raw), "Position_List(LeftRight)", False

    if is_driver_passenger_placeholder(ph_stripped):
        raw = _collect_positions_for_sku(sku, pos_idx)
        if not raw and sku_type.strip() == "Kit/Package":
            raw = _collect_positions_for_kit_members(sku, parts_idx, pos_idx)
        return expected_driver_passenger_from_positions(raw), "Position_List(DriverPassenger)", False

    if is_position_placeholder(ph_stripped):
        allowed = expected_positions_for_sku(
            sku=sku,
            sku_type=sku_type,
            pos_idx=pos_idx,
            parts_idx=parts_idx,
            strip_sides=strip_sides_for_position,
        )
        return allowed, "Position_List", False

    if is_position_for_category_placeholder(ph_stripped):
        if sku_type.strip() != "Kit/Package":
            return [], "Position_List(category)", False
        category = extract_category_from_position_for(ph_stripped)
        allowed = expected_positions_for_kit_member_category(
            kit_sku=sku,
            category_in_placeholder=category,
            parts_idx=parts_idx,
            pos_idx=pos_idx,
            strip_sides=strip_sides_for_position,
        )
        return allowed, "Position_List(category)", False

    if ph_stripped.upper() == PH_MIRROR_COLOR:
        allowed, col, allow_lit = expected_texts_for_mirror_color(
            sku=sku, sku_type=sku_type, saved_idx=saved_idx, saved_norm_map=saved_norm_map, parts_idx=parts_idx
        )
        return allowed, f"SavedSearch:{col}", allow_lit

    if ph_stripped.upper() == PH_MIRROR_FEATURES:
        allowed, col, allow_lit = expected_texts_for_mirror_features(
            sku=sku, sku_type=sku_type, saved_idx=saved_idx, saved_norm_map=saved_norm_map, parts_idx=parts_idx
        )
        return allowed, f"SavedSearch:{col}", allow_lit

    ph_inner_upper = norm_placeholder_token(ph_stripped)

    if ph_inner_upper in boolean_map:
        allowed, col, allow_lit = expected_texts_for_boolean_placeholder(
            sku=sku,
            sku_type=sku_type,
            placeholder_inner_upper=ph_inner_upper,
            boolean_map=boolean_map,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
            parts_idx=parts_idx,
        )
        return allowed, f"Boolean:{col}", allow_lit

    if sku_type.strip() in ("Inventory Item", "Assembly/Bill of Materials"):
        allowed, col = get_expected_candidates_inventory_or_assembly(
            sku=sku, placeholder_inner_upper=ph_inner_upper, saved_idx=saved_idx, saved_norm_map=saved_norm_map
        )
    elif sku_type.strip() == "Kit/Package":
        allowed, col = get_expected_candidates_kit_any_member(
            kit_sku=sku,
            placeholder_inner_upper=ph_inner_upper,
            parts_idx=parts_idx,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
        )
    else:
        allowed, col = [], None

    return allowed, f"SavedSearch:{col}", False


# =========================
# FALLBACK PLACEHOLDER VALIDATION
# =========================

def fallback_validate_placeholders(
    sku: str,
    sku_type: str,
    built_text: str,
    template_text: str,
    boolean_map: Dict[str, Dict[str, str]],
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
    pos_idx: Dict[str, List[Dict[str, str]]],
    strip_sides_for_position: bool,
) -> Tuple[List[str], List[str]]:
    templ = "" if template_text is None else str(template_text)
    placeholders = PLACEHOLDER_PATTERN.findall(templ)
    built = "" if built_text is None else str(built_text)

    missing: List[str] = []
    mismatched: List[str] = []

    for ph in placeholders:
        phs = ph.strip()
        if phs in IGNORED_PLACEHOLDERS:
            continue

        expected, source, allow_literal_if_missing = _expected_candidates_for_placeholder(
            sku=sku,
            sku_type=sku_type,
            ph=phs,
            strip_sides_for_position=strip_sides_for_position,
            boolean_map=boolean_map,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
            parts_idx=parts_idx,
            pos_idx=pos_idx,
        )

        ph_inner_upper = norm_placeholder_token(phs)

        if not expected:
            if allow_literal_if_missing and phs in built:
                continue
            missing.append(phs)
            continue

        non_empty_expected = [e for e in expected if str(e).strip() != ""]
        if not non_empty_expected:
            continue

        base_upper = strip_positional_suffix(ph_inner_upper)
        if base_upper in DUAL_MEASURE_PLACEHOLDERS:
            title_duals = [m.group(0) for m in RX_TITLE_DUAL.finditer(built)]
            if title_duals:
                ok_any = any(dual_measure_match(obs_fragment, expected) for obs_fragment in title_duals)
                if ok_any:
                    continue
                mismatched.append(
                    f"{phs} | fallback_dual_measure_no_match | found={title_duals} | allowed={expected} | source='{source}'"
                )
                continue

        ok = any(_contains_fuzzy(built, cand) for cand in non_empty_expected)
        if not ok:
            mismatched.append(f"{phs} | fallback_contains_fail | allowed={expected} | source='{source}'")

    return sorted(set(missing)), mismatched


# =========================
# CAPTURE-BASED VALIDATION
# =========================

def validate_field_capture_mode(
    sku: str,
    sku_type: str,
    field_name: str,
    built_text: object,
    template_text: object,
    saved_idx: Dict[str, pd.Series],
    saved_norm_map: Dict[str, str],
    parts_idx: Dict[str, pd.Series],
    pos_idx: Dict[str, List[Dict[str, str]]],
    boolean_map: Dict[str, Dict[str, str]],
    strip_sides_for_position: bool,
) -> FieldValidationResult:
    built = "" if pd.isna(built_text) else str(built_text)
    templ = "" if pd.isna(template_text) else str(template_text)

    structure_ok, captures, placeholders_in_template = extract_captures(templ, built)

    missing: List[str] = []
    mismatched: List[str] = []
    ignored_seen: List[str] = []

    if not structure_ok:
        for ph in placeholders_in_template:
            if ph in IGNORED_PLACEHOLDERS:
                ignored_seen.append(ph)
        return FieldValidationResult(
            sku=sku,
            field_name=field_name,
            built_text=built,
            template_text=templ,
            structure_match=False,
            missing_placeholders=[],
            mismatched_placeholders=[],
            ignored_placeholders_seen=sorted(set(ignored_seen)),
        )

    exp_by_index: Dict[int, List[str]] = {}
    inner_by_index: Dict[int, str] = {}

    for idx, cap in enumerate(captures):
        ph = cap.placeholder.strip()
        if ph in IGNORED_PLACEHOLDERS:
            continue

        expected, _, allow_literal_if_missing = _expected_candidates_for_placeholder(
            sku=sku,
            sku_type=sku_type,
            ph=ph,
            strip_sides_for_position=strip_sides_for_position,
            boolean_map=boolean_map,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
            parts_idx=parts_idx,
            pos_idx=pos_idx,
        )
        exp_by_index[idx] = expected
        inner_by_index[idx] = norm_placeholder_token(ph)

        if not expected and allow_literal_if_missing and placeholder_literal_equals_observed(ph, cap.observed_text):
            exp_by_index[idx] = []

    runs = consecutive_placeholder_runs(placeholders_in_template, templ)
    for run in runs:
        any_failed = False
        for ridx in run:
            expected = exp_by_index.get(ridx, [])
            if not expected:
                any_failed = True
                break
            obs = captures[ridx].observed_text
            if not validate_observed_with_dual_fallback(inner_by_index.get(ridx, ""), obs, expected):
                any_failed = True
                break

        if any_failed:
            resolve_consecutive_run_by_expected(run, captures, exp_by_index, inner_by_index)

    for idx, cap in enumerate(captures):
        ph = cap.placeholder.strip()
        observed = "" if cap.observed_text is None else cap.observed_text.strip()

        if ph in IGNORED_PLACEHOLDERS:
            ignored_seen.append(ph)
            continue

        expected, source, allow_literal_if_missing = _expected_candidates_for_placeholder(
            sku=sku,
            sku_type=sku_type,
            ph=ph,
            strip_sides_for_position=strip_sides_for_position,
            boolean_map=boolean_map,
            saved_idx=saved_idx,
            saved_norm_map=saved_norm_map,
            parts_idx=parts_idx,
            pos_idx=pos_idx,
        )

        if not expected:
            if allow_literal_if_missing and placeholder_literal_equals_observed(ph, observed):
                continue
            missing.append(ph)
            continue

        ph_inner_upper = norm_placeholder_token(ph)

        if not validate_observed_with_dual_fallback(ph_inner_upper, observed, expected):
            mismatched.append(f"{ph} | observed='{observed}' | allowed={expected} | source='{source}'")

    return FieldValidationResult(
        sku=sku,
        field_name=field_name,
        built_text=built,
        template_text=templ,
        structure_match=True,
        missing_placeholders=sorted(set(missing)),
        mismatched_placeholders=mismatched,
        ignored_placeholders_seen=sorted(set(ignored_seen)),
    )


# =========================
# MAIN + OUTPUT
# =========================

def main() -> None:
    rules_df = read_rules_templates(BUSINESS_RULES_XLSX_PATH)
    boolean_df = read_rules_boolean_sheet(BUSINESS_RULES_XLSX_PATH)

    titles_df = read_table(TITLES_BUILDS_FILE_PATH)
    saved_df = read_table(SAVED_SEARCH_CSV_PATH)
    parts_df = read_table(PARTS_IN_PACKAGE_CSV_PATH)
    pos_df = read_table(POSITION_LIST_CSV_PATH)

    boolean_map = build_boolean_placeholder_map(boolean_df)

    saved_idx, saved_norm_map = build_saved_search_index(saved_df)
    parts_idx = build_parts_in_package_index(parts_df)
    pos_idx = build_position_index(pos_df)

    for col in [COL_TITLES_SKU, COL_TITLES_TYPE, COL_TITLES_RULE_TYPE]:
        if col not in titles_df.columns:
            raise ValueError(f"Titles/Bullets file is missing required column: '{col}'")

    if COL_RULES_TYPE not in rules_df.columns:
        raise ValueError(f"Business Rules Templates sheet is missing required column: '{COL_RULES_TYPE}'")

    if not any(c in rules_df.columns for c in REQUIRED_RULE_TEMPLATE_COLUMNS):
        raise ValueError(
            "Business Rules Templates sheet does not contain any of the expected template columns. "
            f"Expected one of: {REQUIRED_RULE_TEMPLATE_COLUMNS}"
        )

    existing_text_fields = [c for c in TEXT_FIELDS_TO_VALIDATE if c in titles_df.columns]
    if not existing_text_fields:
        raise ValueError(
            "None of the configured TEXT_FIELDS_TO_VALIDATE exist in the Titles/Bullets file. "
            f"Configured: {TEXT_FIELDS_TO_VALIDATE}. File columns: {list(titles_df.columns)}"
        )

    fields_validated = [c for c in existing_text_fields if c in rules_df.columns]
    if not fields_validated:
        raise ValueError(
            "No validation fields exist in BOTH Titles/Bullets file and Business Rules Templates sheet. "
            f"Titles fields found: {existing_text_fields}. Rules columns: {list(rules_df.columns)}"
        )

    print(f"SKUs loaded from Titles/Bullets file: {len(titles_df)}")
    print(f"Saved search rows indexed by SKU: {len(saved_idx)}")
    print(f"Parts_in_Package rows indexed by kit SKU: {len(parts_idx)}")
    print(f"Position_List SKUs indexed: {len(pos_idx)}")
    print(f"Boolean placeholders loaded: {len(boolean_map)}")
    print(f"Fields validated (present in titles + rules): {len(fields_validated)}")

    output_rows: List[Dict[str, object]] = []

    count_no_rule = 0
    count_rule_error = 0
    count_placeholder_error = 0
    count_ok = 0
    count_fallback_checked = 0
    count_capture_checked = 0

    for _, row in titles_df.iterrows():
        sku = str(row[COL_TITLES_SKU]).strip()
        sku_type = str(row[COL_TITLES_TYPE]).strip()
        rule_key = str(row[COL_TITLES_RULE_TYPE]).strip()

        if not sku or sku.lower() == "nan":
            continue

        rule_row = choose_rule_row(rules_df, row)
        if rule_row is None:
            count_no_rule += 1
            output_rows.append(
                {
                    "SKU": sku,
                    "Type": sku_type,
                    "Rule Key": rule_key,
                    "Field": "",
                    "rule_status": "NO_RULE_MATCH",
                    "placeholder_status": "SKIPPED",
                    "placeholder_check_mode": "SKIPPED",
                    "rule_error": "No matching rule row found in Templates sheet",
                    "placeholder_error": "",
                }
            )
            continue

        sku_side_handled = has_any_side_carrying_placeholder(rule_row, fields_validated)

        for field in fields_validated:
            built_val = row.get(field, "")
            templ_val = rule_row.get(field, "")

            field_placeholders = set(placeholders_in_text(templ_val))
            field_side_handled = bool(field_placeholders.intersection(SIDE_CARRYING_PLACEHOLDERS))
            strip_sides_for_position = sku_side_handled or field_side_handled

            fr = validate_field_capture_mode(
                sku=sku,
                sku_type=sku_type,
                field_name=field,
                built_text=built_val,
                template_text=templ_val,
                saved_idx=saved_idx,
                saved_norm_map=saved_norm_map,
                parts_idx=parts_idx,
                pos_idx=pos_idx,
                boolean_map=boolean_map,
                strip_sides_for_position=strip_sides_for_position,
            )

            if not fr.structure_match:
                rule_status = "ERROR"
                rule_error = "STRUCTURE_FAIL (built text does not match template structure)"
                count_rule_error += 1

                missing_fb, mismatched_fb = fallback_validate_placeholders(
                    sku=sku,
                    sku_type=sku_type,
                    built_text="" if pd.isna(built_val) else str(built_val),
                    template_text="" if pd.isna(templ_val) else str(templ_val),
                    boolean_map=boolean_map,
                    saved_idx=saved_idx,
                    saved_norm_map=saved_norm_map,
                    parts_idx=parts_idx,
                    pos_idx=pos_idx,
                    strip_sides_for_position=strip_sides_for_position,
                )
                count_fallback_checked += 1
                placeholder_check_mode = "FALLBACK"

                if missing_fb or mismatched_fb:
                    placeholder_status = "ERROR"
                    details: List[str] = []
                    if missing_fb:
                        details.append(f"missing={missing_fb}")
                    if mismatched_fb:
                        details.append(f"mismatched={mismatched_fb}")
                    placeholder_error = " | ".join(details)
                    count_placeholder_error += 1
                else:
                    placeholder_status = "OK"
                    placeholder_error = ""
                    count_ok += 1
            else:
                rule_status = "OK"
                rule_error = ""
                count_capture_checked += 1
                placeholder_check_mode = "CAPTURE"

                if fr.missing_placeholders or fr.mismatched_placeholders:
                    placeholder_status = "ERROR"
                    details = []
                    if fr.missing_placeholders:
                        details.append(f"missing={fr.missing_placeholders}")
                    if fr.mismatched_placeholders:
                        details.append(f"mismatched={fr.mismatched_placeholders}")
                    placeholder_error = " | ".join(details)
                    count_placeholder_error += 1
                else:
                    placeholder_status = "OK"
                    placeholder_error = ""
                    count_ok += 1

            if rule_status == "OK" and placeholder_status == "OK":
                continue

            output_rows.append(
                {
                    "SKU": sku,
                    "Type": sku_type,
                    "Rule Key": rule_key,
                    "Field": field,
                    "rule_status": rule_status,
                    "placeholder_status": placeholder_status,
                    "placeholder_check_mode": placeholder_check_mode,
                    "rule_error": rule_error,
                    "placeholder_error": placeholder_error,
                }
            )

    out_df = pd.DataFrame(output_rows)
    Path(OUTPUT_REPORT_XLSX_PATH).parent.mkdir(parents=True, exist_ok=True)
    out_df.to_excel(OUTPUT_REPORT_XLSX_PATH, index=False, engine="openpyxl")

    print(f"Rows written (errors only): {len(out_df)}")
    print(f"Rows with NO_RULE_MATCH: {count_no_rule}")
    print(f"Fields with RULE errors (STRUCTURE_FAIL): {count_rule_error}")
    print(f"Fields with PLACEHOLDER errors: {count_placeholder_error}")
    print(f"Fields checked via CAPTURE: {count_capture_checked}")
    print(f"Fields checked via FALLBACK: {count_fallback_checked}")
    print(f"Fields fully OK (not written): {count_ok}")
    print(f"Report XLSX written: {OUTPUT_REPORT_XLSX_PATH}")


if __name__ == "__main__":
    main()