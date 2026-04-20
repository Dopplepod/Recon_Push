from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List
import pandas as pd

from .config import SUMMARY_ORDER, OPERATING_EXPENSE_BUCKETS
from .helpers import (
    normalize_code,
    extract_local_coa_code,
    extract_local_coa_desc,
    first_three_digits,
    pick_sheet_name,
    detect_header_row,
    standardize_headers,
    line_item_from_bucket,
    canonical_line_item,
)
from .mappings import MappingRepository


class ReconError(Exception):
    pass


@dataclass
class ReconResult:
    payload: Dict


class ReconciliationService:
    def __init__(self, base_dir: Path):
        self.base_dir = base_dir
        self.repo = MappingRepository(base_dir)

    def _load_sap_raw(self, file_obj) -> pd.DataFrame:
        xl = pd.ExcelFile(file_obj)
        sheet = pick_sheet_name(xl, "TB BFC")
        aliases = {
            "gl code": "GL Code",
            "sap mapping": "SAP Mapping",
            "amount": "Amount",
            "gl name": "GL Name",
            "sap description": "SAP Description",
        }
        header_row = detect_header_row(xl, sheet, aliases)
        df = pd.read_excel(xl, sheet_name=sheet, header=header_row)
        df = standardize_headers(df, aliases)
        required = ["GL Code", "SAP Mapping", "Amount"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ReconError(f"SAP raw file is missing required columns: {', '.join(missing)}")
        work = pd.DataFrame({
            "gl_code": df["GL Code"].map(normalize_code),
            "gl_name": df["GL Name"].fillna("").astype(str).str.strip() if "GL Name" in df.columns else "",
            "sap_mapping_raw": df["SAP Mapping"].map(normalize_code),
            "amount": pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0),
        })
        work = work[(work["gl_code"] != "") & (work["sap_mapping_raw"] != "")]
        return work

    def _detect_os_amount_col(self, df: pd.DataFrame):
        if "Amount" in df.columns and pd.to_numeric(df["Amount"], errors="coerce").notna().any():
            return "Amount", []
        warnings = ["OS raw file has no usable Amount column; structure-only mode loaded with zero amounts."]
        return None, warnings

    def _load_os_raw(self, file_obj):
        xl = pd.ExcelFile(file_obj)
        sheet = pick_sheet_name(xl, "OS TB")
        df = pd.read_excel(xl, sheet_name=sheet)
        if "Local COA" not in df.columns:
            raise ReconError("OS raw file is missing required column: Local COA")
        amount_col, warnings = self._detect_os_amount_col(df)
        amt = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0) if amount_col else pd.Series([0.0] * len(df))
        work = pd.DataFrame({
            "local_coa_raw": df["Local COA"].fillna(""),
            "local_coa_code": df["Local COA"].map(extract_local_coa_code),
            "local_coa_desc": df["Local COA"].map(extract_local_coa_desc),
            "sap_code": df["SAP COA"].map(normalize_code) if "SAP COA" in df.columns else "",
            "os_coa": df["OS COA"].map(normalize_code) if "OS COA" in df.columns else "",
            "amount": amt,
        })
        work = work[work["local_coa_code"] != ""].copy()
        return work, warnings

    @staticmethod
    def _is_pl_bucket(value: str) -> bool:
        value = str(value or '').strip()
        digits = ''.join(ch for ch in value if ch.isdigit())
        if len(digits) < 3:
            return False
        try:
            return int(digits[:3]) >= 410
        except ValueError:
            return False

    @staticmethod
    def _join_unique_text(series: pd.Series) -> str:
        vals: List[str] = []
        for v in series.fillna("").astype(str):
            v = v.strip()
            if v and v not in vals:
                vals.append(v)
        return " | ".join(vals)

    @staticmethod
    def _safe_number(value) -> float:
        try:
            return float(value)
        except Exception:
            return 0.0

    def _summary_from_buckets(self, bucket_totals: Dict[str, float]) -> Dict[str, float]:
        revenue = bucket_totals.get("410", 0.0)
        opex = sum(bucket_totals.get(code, 0.0) for code in OPERATING_EXPENSE_BUCKETS)
        ebitda = revenue + opex
        da = bucket_totals.get("504", 0.0)
        ebit = ebitda + da
        finance_income = bucket_totals.get("602", 0.0)
        finance_expense = bucket_totals.get("603", 0.0)
        net_finance = finance_income + finance_expense
        share_results = bucket_totals.get("801", 0.0)
        non_operating = bucket_totals.get("601", 0.0)
        exceptional = bucket_totals.get("699", 0.0)
        pbt = ebit + net_finance + share_results + non_operating + exceptional
        income_tax = bucket_totals.get("607", 0.0)
        discontinued = bucket_totals.get("861", 0.0)
        pat = pbt + income_tax + discontinued
        minority = 0.0
        att_own = pat + minority
        return {
            "Revenue": revenue,
            "Operating Expense (Ex-D And A)": opex,
            "EBITDA": ebitda,
            "EBIT": ebit,
            "6020000 - Finance Income": finance_income,
            "6030000 - Finance Expense": finance_expense,
            "Net Finance (Income / Expense)": net_finance,
            "8010000 - Share Of Results Of AJV": share_results,
            "6010000 - Non Operating Gain Loss": non_operating,
            "6990000 - Exceptional Items": exceptional,
            "Profit Or Loss Before Tax": pbt,
            "6070000 - Income Tax Expense": income_tax,
            "8610000 - Profit Or Loss From Discontinued Operation (Net Of Tax)": discontinued,
            "Net Profit Or Loss (PAT)": pat,
            "PL_MI - Minority Interest": minority,
            "Profit Or Loss Attributable To Owners Of The Company": att_own,
        }

    def reconcile(self, sap_file, os_file, entity: str = "") -> ReconResult:
        warnings: List[str] = []
        sap_raw = self._load_sap_raw(sap_file)
        os_raw, os_warnings = self._load_os_raw(os_file)
        warnings.extend(os_warnings)

        bfc_map = self.repo.load_bfc_to_os()
        hierarchy_map = self.repo.hierarchy_map()

        sap = sap_raw.merge(
            bfc_map[["sap_mapping_raw", "os_level2", "bucket"]].drop_duplicates(),
            on="sap_mapping_raw",
            how="left",
        )
        sap["bucket"] = sap["bucket"].fillna(sap["os_level2"].map(first_three_digits))
        sap["line_item"] = sap["bucket"].map(lambda x: hierarchy_map.get(x, line_item_from_bucket(x))).map(canonical_line_item)
        sap["detail_id"] = sap["line_item"].fillna("") + "_" + sap["gl_code"]

        os_raw["bucket"] = os_raw["os_coa"].map(first_three_digits)
        os_raw["line_item"] = os_raw["bucket"].map(lambda x: hierarchy_map.get(x, line_item_from_bucket(x))).map(canonical_line_item)
        os_raw["detail_id"] = os_raw["line_item"].fillna("") + "_" + os_raw["local_coa_code"]

        sap_unmapped = sap[sap["line_item"].fillna("").eq("")].copy()
        os_unmapped = os_raw[os_raw["line_item"].fillna("").eq("")].copy()

        sap_non_pl = sap[sap["line_item"].fillna("").ne("") & ~sap["bucket"].map(self._is_pl_bucket)].copy()
        os_non_pl = os_raw[os_raw["line_item"].fillna("").ne("") & ~os_raw["bucket"].map(self._is_pl_bucket)].copy()

        sap_visible = sap[sap["line_item"].fillna("").ne("") & sap["bucket"].map(self._is_pl_bucket)].copy()
        os_visible = os_raw[os_raw["line_item"].fillna("").ne("") & os_raw["bucket"].map(self._is_pl_bucket)].copy()

        sap_detail = (
            sap_visible.groupby(["bucket", "line_item", "detail_id", "gl_code"], as_index=False)
            .agg(
                sap_amount=("amount", "sum"),
                description=("gl_name", self._join_unique_text),
                route=("sap_mapping_raw", self._join_unique_text),
            )
        )
        sap_detail["display_code"] = sap_detail["gl_code"]

        os_detail = (
            os_visible.groupby(["bucket", "line_item", "detail_id", "local_coa_code"], as_index=False)
            .agg(
                os_amount=("amount", "sum"),
                description=("local_coa_desc", self._join_unique_text),
                raw_os_coa=("os_coa", self._join_unique_text),
                sap_code=("sap_code", self._join_unique_text),
            )
        )
        os_detail["display_code"] = os_detail["local_coa_code"]

        detail_compare = sap_detail.merge(
            os_detail[["bucket", "line_item", "detail_id", "os_amount", "display_code"]],
            on=["bucket", "line_item", "detail_id"],
            how="outer",
            suffixes=("_sap", "_os"),
        )
        detail_compare["sap_amount"] = detail_compare["sap_amount"].fillna(0.0)
        detail_compare["os_amount"] = detail_compare["os_amount"].fillna(0.0)
        detail_compare["difference"] = detail_compare["sap_amount"] - detail_compare["os_amount"]
        detail_compare["display_code"] = detail_compare["display_code_sap"].fillna(detail_compare["display_code_os"]).fillna("")

        bucket_compare = (
            detail_compare.groupby(["bucket", "line_item"], as_index=False)
            .agg(sap_amount=("sap_amount", "sum"), os_amount=("os_amount", "sum"))
        )
        bucket_compare["difference"] = bucket_compare["sap_amount"] - bucket_compare["os_amount"]

        bucket_totals_sap = dict(zip(bucket_compare["bucket"], bucket_compare["sap_amount"]))
        bucket_totals_os = dict(zip(bucket_compare["bucket"], bucket_compare["os_amount"]))
        sap_summary = self._summary_from_buckets(bucket_totals_sap)
        os_summary = self._summary_from_buckets(bucket_totals_os)

        summary_rows = []
        classification_note = ""
        net_finance_diff = self._safe_number(sap_summary.get("Net Finance (Income / Expense)")) - self._safe_number(os_summary.get("Net Finance (Income / Expense)"))
        if abs(net_finance_diff) < 0.5 and (
            abs((self._safe_number(sap_summary.get("6020000 - Finance Income")) - self._safe_number(os_summary.get("6020000 - Finance Income"))) +
                (self._safe_number(sap_summary.get("6030000 - Finance Expense")) - self._safe_number(os_summary.get("6030000 - Finance Expense")))) < 0.5
        ):
            classification_note = "Classification difference only: Finance Income and Finance Expense offset at Net Finance."

        for name in SUMMARY_ORDER:
            s = float(sap_summary.get(name, 0.0))
            o = float(os_summary.get(name, 0.0))
            row = {"line_item": name, "sap_bfc": s, "onestream": o, "difference": s - o}
            if name == "Net Finance (Income / Expense)" and classification_note:
                row["note"] = classification_note
            summary_rows.append(row)

        drilldown_groups = []
        for _, bucket_row in bucket_compare.sort_values(["bucket", "line_item"]).iterrows():
            group_df = detail_compare[
                (detail_compare["bucket"] == bucket_row["bucket"]) &
                (detail_compare["line_item"] == bucket_row["line_item"])
            ].copy().sort_values(["display_code", "detail_id"])
            children = []
            for _, r in group_df.iterrows():
                children.append({
                    "display_code": r.get("display_code", ""),
                    "detail_key": r.get("detail_id", ""),
                    "description": "",
                    "currency": "",
                    "sap_bfc": float(r.get("sap_amount", 0.0)),
                    "onestream": float(r.get("os_amount", 0.0)),
                    "difference": float(r.get("difference", 0.0)),
                })
            drilldown_groups.append({
                "bucket": bucket_row["bucket"],
                "line_item": canonical_line_item(bucket_row["line_item"]),
                "sap_bfc": float(bucket_row["sap_amount"]),
                "onestream": float(bucket_row["os_amount"]),
                "difference": float(bucket_row["difference"]),
                "children": children,
            })

        payload = {
            "meta": {
                "sap_rows": int(len(sap_raw)),
                "os_rows": int(len(os_raw)),
                "gl_codes": int(sap_raw["gl_code"].nunique()),
                "unmapped_gl_codes": int(sap_unmapped["gl_code"].nunique()) if "gl_code" in sap_unmapped.columns else 0,
                "entity": entity or "2403",
            },
            "summary": summary_rows,
            "drilldown": drilldown_groups,
            "debug": {
                "warnings": warnings,
                "excluded_non_pl": {
                    "sap_rows": int(len(sap_non_pl)),
                    "os_rows": int(len(os_non_pl)),
                    "sap_line_items": sorted(sap_non_pl["line_item"].dropna().astype(str).unique().tolist())[:100],
                    "os_line_items": sorted(os_non_pl["line_item"].dropna().astype(str).unique().tolist())[:100],
                    "rule": "Only P&L buckets with first 3 digits >= 410 are included in visible summary and drilldown.",
                },
                "methodology": {
                    "sap_id": "OS Level 2 + '_' + GL Code",
                    "sap_join_key": "Raw SAP Mapping with suffix (exact match)",
                    "os_id": "OS Level 2 + '_' + Local COA code",
                    "os_level2_basis": "derived from first 3 digits of raw OS COA and expanded to full label",
                    "entity_model": "Entity-agnostic. The backend logic is reusable across entities as long as the uploaded raw files follow the same structure.",
                },
                "unmapped_sap": sap_unmapped[["gl_code", "sap_mapping_raw"]].head(200).to_dict(orient="records") if not sap_unmapped.empty else [],
                "unmapped_os": os_unmapped[["local_coa_code", "sap_code", "os_coa"]].head(200).to_dict(orient="records") if not os_unmapped.empty else [],
            }
        }
        return ReconResult(payload)
