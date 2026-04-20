from collections import OrderedDict

LEVEL2_LABELS = OrderedDict([
    ("410", "4100000 - Revenue"),
    ("501", "5010000 - Cost Of Raw Materials And Supplies"),
    ("502", "5020000 - Staff Costs"),
    ("503", "5030000 - Licence Fees"),
    ("504", "5040000 - Depreciation And Amortisation"),
    ("505", "5050000 - Company Premise Utilities And Maintenance"),
    ("506", "5060000 - Subcontracting Services"),
    ("507", "5070000 - Travel And Transport"),
    ("508", "5080000 - Other Costs"),
    ("601", "6010000 - Non Operating Gain Loss"),
    ("602", "6020000 - Finance Income"),
    ("603", "6030000 - Finance Expense"),
    ("607", "6070000 - Income Tax Expense"),
    ("699", "6990000 - Exceptional Items"),
    ("801", "8010000 - Share Of Results Of AJV"),
    ("861", "8610000 - Profit Or Loss From Discontinued Operation (Net Of Tax)"),
])

# High-level summary rows shown in the Summary tab.
SUMMARY_ORDER = [
    "Revenue",
    "Operating Expense (Ex-D And A)",
    "EBITDA",
    "EBIT",
    "6020000 - Finance Income",
    "6030000 - Finance Expense",
    "8010000 - Share Of Results Of AJV",
    "6010000 - Non Operating Gain Loss",
    "6990000 - Exceptional Items",
    "Profit Or Loss Before Tax",
    "6070000 - Income Tax Expense",
    "8610000 - Profit Or Loss From Discontinued Operation (Net Of Tax)",
    "Net Profit Or Loss (PAT)",
    "PL_MI - Minority Interest",
    "Profit Or Loss Attributable To Owners Of The Company",
]

OPERATING_EXPENSE_BUCKETS = ["501", "502", "503", "505", "506", "507", "508"]

REFERENCE_FILES = {
    "bfc_to_os": "data/mappings/BFC_To_OS_Mapping.xlsx",
    "hierarchy": "data/reference/hierarchy.xml",
}
