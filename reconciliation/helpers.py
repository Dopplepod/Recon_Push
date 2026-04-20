import re
from typing import Optional, Dict
import pandas as pd
import xml.etree.ElementTree as ET

from .config import LEVEL2_LABELS


def normalize_code(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text.strip()


def normalize_header(value) -> str:
    text = normalize_code(value)
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def standardize_headers(df: pd.DataFrame, aliases: Dict[str, str]) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        norm = normalize_header(col)
        if norm in aliases:
            rename_map[col] = aliases[norm]
    return df.rename(columns=rename_map)


def detect_header_row(excel_file, sheet_name: str, required_aliases: Dict[str, str], scan_rows: int = 10) -> int:
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=scan_rows)
    required_targets = set(required_aliases.values())
    best_idx = 0
    best_score = -1
    for idx in range(len(preview)):
        row_values = [normalize_header(v) for v in preview.iloc[idx].tolist()]
        matched = {required_aliases[v] for v in row_values if v in required_aliases}
        score = len(matched.intersection(required_targets))
        if score > best_score:
            best_idx = idx
            best_score = score
        if score == len(required_targets):
            return idx
    return best_idx


def extract_local_coa_code(value) -> str:
    text = normalize_code(value)
    if not text:
        return ""
    match = re.match(r"^\s*([^\-]+?)\s*\-", text)
    if match:
        return normalize_code(match.group(1))
    token = text.split()[0]
    return normalize_code(token)


def extract_local_coa_desc(value) -> str:
    text = normalize_code(value)
    if not text:
        return ""
    parts = re.split(r"\s+-\s+", text, maxsplit=1)
    return parts[1].strip() if len(parts) == 2 else ""


def first_three_digits(value) -> str:
    text = normalize_code(value)
    digits = re.sub(r"\D", "", text)
    return digits[:3] if len(digits) >= 3 else ""


def extract_os_level2_code(line_item: str) -> str:
    text = normalize_code(line_item)
    match = re.match(r"^(\d{7})", text)
    return match.group(1) if match else ""


def canonical_line_item(line_item: str) -> str:
    code = extract_os_level2_code(line_item)
    if code:
        bucket = first_three_digits(code)
        return LEVEL2_LABELS.get(bucket, line_item.strip())
    return normalize_code(line_item)


def pick_sheet_name(excel_file, preferred: Optional[str] = None) -> str:
    if preferred and preferred in excel_file.sheet_names:
        return preferred
    return excel_file.sheet_names[0]


def line_item_from_bucket(bucket: str) -> str:
    return LEVEL2_LABELS.get(bucket, f"{bucket}0000 - Unassigned")


def parse_hierarchy_level2_map(xml_path) -> Dict[str, str]:
    """Build a fallback first-3-digits -> full level2 label mapping from hierarchy.xml."""
    result: Dict[str, str] = {}
    try:
        root = ET.parse(xml_path).getroot()
    except Exception:
        return dict(LEVEL2_LABELS)

    for elem in root.iter():
        code = normalize_code(elem.attrib.get("code"))
        name = normalize_code(elem.attrib.get("name"))
        if not code or not name:
            continue
        bucket = first_three_digits(code)
        if len(code) == 7 and bucket and bucket not in result:
            result[bucket] = f"{code} - {name}"
    merged = dict(result)
    merged.update(LEVEL2_LABELS)
    return merged
