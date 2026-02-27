# Touche Hair Caterham — Retail Statements (Streamlit)

## What it does
- Cleans **Retail Sales by Team Member.xls** (sheet: `Retail Sales by Team Memb`) using your Power Query logic
- Joins `Description` to **Inspired Hair Price List.xlsx** (select mapping columns in sidebar)
- Validates:
  - Missing/invalid costs
  - Unmatched product descriptions
- Outputs:
  - Optional Excel workbook (invoice output + optional cleaned tabs + validation tabs)
  - ZIP of per-stylist PDF statements (Product, Qty, Cost, Invoice Amount)

## Secrets (required)
Add a password in Streamlit Community Cloud Secrets:

```toml
[auth]
password = "your-strong-password"
```

## Deploy
- Main file: `app.py`
- Python: 3.12 recommended
