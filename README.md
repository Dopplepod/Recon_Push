# P&L Reconciliation V12

Run locally:

```bash
pip install -r requirements.txt
python app.py
```

Then open http://127.0.0.1:5000

## What this build does
- Reconciles only P&L buckets (`>= 4100000` at the first-3-digit level).
- Uses the exact raw SAP Mapping with suffix to join BFC to OS mapping.
- Derives OS Level 2 from the raw OS COA and the hierarchy reference.
- Uses entity-agnostic logic, so the same program can be reused across entities as long as the raw file structure is consistent.
- Keeps Finance Income and Finance Expense separate, but computes **Net Finance (Income / Expense)** as a derived summary row so classification offsets cancel cleanly at summary level.
- Uses canonical OS bucket labels to avoid case-only duplicate breaks.
- Shows fixed bold summary sections in drilldown and expands only the detailed reporting lines.

## Included reference files
- `data/mappings/BFC_To_OS_Mapping.xlsx`
- `data/reference/hierarchy.xml`

## Notes
- If the OS raw file has no usable `Amount` column, the app loads it in structure-only mode with zero OS amounts and a warning in Debug.


## Run with Docker

Build the image:

```bash
docker build -t reconciliation-app .
```

Run the container:

```bash
docker run --rm -p 5000:5000 reconciliation-app
```

Then open http://127.0.0.1:5000

### Optional: persistent secret
If you want to change the Flask secret for production, update `app.secret_key` in `app.py` before building.

### Notes for deployment
- Container listens on port `5000`.
- Works on Render or any Docker-capable host.
- Keep the `data/` folder inside the image because it contains the hierarchy and mapping references.
