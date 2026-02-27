import numpy as np
import pandas as pd

# Retail report is .xls; pandas needs xlrd at runtime in Streamlit.
# We keep parsing robust but largely mirror your Power Query steps.

def _norm_text(s):
    if s is None:
        return None
    out = str(s).replace("\u00a0", " ").strip()
    return out if out != "" else None

def _find_col(columns, candidates):
    cols_lower = {c.lower(): c for c in columns}
    for cand in candidates:
        key = cand.lower()
        if key in cols_lower:
            return cols_lower[key]
    # partial match
    for c in columns:
        cl = c.lower()
        for cand in candidates:
            if cand.lower() in cl:
                return c
    return None

def load_price_list(price_path_or_file) -> pd.DataFrame:
    df = pd.read_excel(price_path_or_file)

    # Auto-detect likely columns
    desc_candidates = ["Description", "Product Description", "Item", "Product", "Name"]
    cost_candidates = ["Cost", "Cost Price", "CostPrice", "Unit Cost", "Unit Cost Price", "UnitCost", "Per Unit", "Per Product", "Trade"]

    desc_col = _find_col(df.columns, desc_candidates) or df.columns[0]
    cost_col = _find_col(df.columns, cost_candidates) or (df.columns[1] if len(df.columns) > 1 else df.columns[0])

    # store suggested indices for UI defaults
    df.attrs["desc_index"] = list(df.columns).index(desc_col)
    df.attrs["cost_index"] = list(df.columns).index(cost_col)

    return df

def convert_retail_sales(retail_path_or_file) -> pd.DataFrame:
    raw = pd.read_excel(retail_path_or_file, sheet_name="Retail Sales by Team Memb", header=None)

    # Keep same physical columns as your Power Query after removals:
    # Column2, Column7, Column11, Column13, Column15 (1-based)
    keep_idx = [1, 6, 10, 12, 14]  # 0-based
    df = raw.iloc[:, keep_idx].copy()
    df.columns = ["Description", "Qty", "Exc Vat", "Inc Vat", "Gross Profit"]

    # Filter rows where Description not null and not "Grand Total"
    df["Description"] = df["Description"].apply(_norm_text)
    df = df[df["Description"].notna()].copy()
    df = df[df["Description"] != "Grand Total"].copy()

    # Promote headers: first remaining row becomes headers
    headers = [str(h).replace("\u00a0", " ").strip() for h in df.iloc[0].tolist()]
    df2 = df.iloc[1:].copy()
    df2.columns = headers

    # --- Robust column mapping (headers vary across exports) ---
    # Find likely columns
    def pick(colnames, candidates):
        lower = {c.lower(): c for c in colnames}
        for cand in candidates:
            if cand.lower() in lower:
                return lower[cand.lower()]
        # partial
        for c in colnames:
            cl = c.lower()
            for cand in candidates:
                if cand.lower() in cl:
                    return c
        return None

    desc_col = pick(df2.columns, ["Description", "Product", "Item"])
    qty_col  = pick(df2.columns, ["Qty", "Quantity", "QTY"])
    inc_col  = pick(df2.columns, ["Inc Vat", "Inc VAT", "Inc. Vat", "Inc", "IncVat"])
    exc_col  = pick(df2.columns, ["Exc Vat", "Exc VAT", "Exc. Vat", "Exc", "ExcVat"])
    gp_col   = pick(df2.columns, ["Gross Profit", "Grossprofit", "Profit"])

    if desc_col is None:
        raise ValueError(f"Could not find a Description column after header promotion. Columns: {list(df2.columns)}")
    if qty_col is None:
        raise ValueError(f"Could not find a Qty/Quantity column after header promotion. Columns: {list(df2.columns)}")

    # Standardise to expected names
    rename_map = {desc_col: "Description", qty_col: "Qty"}
    if inc_col: rename_map[inc_col] = "Inc Vat"
    if exc_col: rename_map[exc_col] = "Exc Vat"
    if gp_col:  rename_map[gp_col]  = "Gross Profit"
    df2 = df2.rename(columns=rename_map)

    # Type conversions
    df2["Description"] = df2["Description"].apply(_norm_text)
    df2["Qty"] = pd.to_numeric(df2["Qty"], errors="coerce")
    for c in ["Inc Vat", "Exc Vat", "Gross Profit"]:
        if c in df2.columns:
            df2[c] = pd.to_numeric(df2[c], errors="coerce")

    # Remove "Inspired Hair Supplies" line
    df2 = df2[df2["Description"] != "Inspired Hair Supplies"].copy()

    # Add Stylist column: if Qty is null then Stylist name is Description
    df2["Stylist"] = np.where(df2["Qty"].isna(), df2["Description"], np.nan)

    # Fill Up (Power Query): fill missing stylist from the row below => reverse ffill then reverse back
    df2["Stylist"] = pd.Series(df2["Stylist"])[::-1].ffill()[::-1]

    # Keep only actual product lines (Qty not null)
    df2 = df2[df2["Qty"].notna()].copy()

    # Build final columns; drop unneeded if present
    keep_cols = ["Stylist", "Description", "Qty"]
    if "Inc Vat" in df2.columns:
        keep_cols.append("Inc Vat")
    # PQ removes Gross Profit and Exc Vat; keep only Inc Vat for reference
    df2 = df2[keep_cols].copy()

    # Normalise fields
    df2["Stylist"] = df2["Stylist"].apply(_norm_text)
    df2["Description"] = df2["Description"].apply(_norm_text)
    df2["Qty"] = df2["Qty"].astype(int)

    return df2.reset_index(drop=True)

def build_retail_output(retail_clean: pd.DataFrame, price_df: pd.DataFrame, price_desc_col: str, price_cost_col: str):
    price = price_df.copy()
    price_desc = price_desc_col
    price_cost = price_cost_col

    # Normalise keys
    price["_key"] = price[price_desc].apply(_norm_text)
    retail = retail_clean.copy()
    retail["_key"] = retail["Description"].apply(_norm_text)

    # Cost numeric
    price["_cost"] = pd.to_numeric(price[price_cost], errors="coerce")

    merged = retail.merge(
        price[["_key", "_cost"]],
        how="left",
        on="_key",
    )
    merged.rename(columns={"_cost": "Cost"}, inplace=True)

    merged["Invoice Amount"] = merged["Qty"] * merged["Cost"]

    out = merged[["Stylist", "Description", "Qty", "Cost", "Invoice Amount"]].copy()
    out = out.sort_values(["Stylist", "Description"]).reset_index(drop=True)

    # Validation tables
    missing_cost_df = out[out["Cost"].isna()].copy()
    unmatched_products_df = (
        out[out["Cost"].isna()][["Description"]]
        .drop_duplicates()
        .sort_values("Description")
        .reset_index(drop=True)
    )

    # Set a period label if the report has one (no dates in the PQ output); leave blank for now.
    out.attrs["period_label"] = ""

    return out, missing_cost_df, unmatched_products_df
