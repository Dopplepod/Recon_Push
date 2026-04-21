
from __future__ import annotations

from pathlib import Path
import pandas as pd

from .config import REFERENCE_FILES
from .helpers import normalize_code, first_three_digits, line_item_from_bucket, parse_hierarchy_level2_map


class MappingRepository:
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self._hierarchy_map = parse_hierarchy_level2_map(self.base_dir / REFERENCE_FILES["hierarchy"])

    def load_bfc_to_os(self) -> pd.DataFrame:
        path = self.base_dir / REFERENCE_FILES["bfc_to_os"]
        df = pd.read_excel(path)
        df = df.rename(columns={df.columns[0]: "sap_mapping", df.columns[1]: "os_level2"})
        df = df.dropna(subset=["sap_mapping"]).copy()
        df["sap_mapping_raw"] = df["sap_mapping"].map(normalize_code)
        df["os_level2"] = df["os_level2"].fillna("").astype(str).str.strip()
        # support either 410 / 501 style or full 4100000 - Revenue style
        df["bucket"] = df["os_level2"].map(first_three_digits)
        df.loc[df["os_level2"].eq(""), "os_level2"] = df.loc[df["os_level2"].eq(""), "bucket"].map(line_item_from_bucket)
        short_mask = df["os_level2"].map(lambda x: x.isdigit() and len(x) <= 3 if isinstance(x, str) else False)
        df.loc[short_mask, "os_level2"] = df.loc[short_mask, "os_level2"].map(line_item_from_bucket)
        # use hierarchy/config fallback for anything still blank
        df.loc[df["os_level2"].eq(""), "os_level2"] = df.loc[df["os_level2"].eq(""), "bucket"].map(
            lambda x: self._hierarchy_map.get(x, line_item_from_bucket(x))
        )
        return df[["sap_mapping_raw", "bucket", "os_level2"]].drop_duplicates()

    def hierarchy_map(self):
        return dict(self._hierarchy_map)
