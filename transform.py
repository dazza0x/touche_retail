
import numpy as np
import pandas as pd

def _norm_text(x):
    if x is None:
        return None
    s = str(x).replace("\u00a0", " ").strip()
    if s == "" or s.lower() in ("nan", "none"):
        return None
    return s

def _pick(colnames, candidates):
    lower = {c.lower(): c for c in colnames}
    for cand in candidates:
        if cand.lower() in lower:
            return lower[cand.lower()]
    for c in colnames:
        cl = c.lower()
        for cand in candidates:
            if cand.lower() in cl:
                return c
    return None

def load_price_list(price_path_or_file) -> pd.DataFrame:
    df = pd.read_excel(price_path_or_file)

    desc_candidates = ["Description", "Product Description", "Item", "Product", "Name"]
    cost_candidates = ["Cost", "Cost Price", "CostPrice", "Unit Cost", "Unit Cost Price", "UnitCost", "Per Unit", "Trade"]

    desc_col = _pick(df.columns, desc_candidates) or df.columns[0]
    cost_col = _pick(df.columns, cost_candidates) or (df.columns[1] if len(df.columns) > 1 else df.columns[0])

    df.attrs["desc_index"] = list(df.columns).index(desc_col)
    df.attrs["cost_index"] = list(df.columns).index(cost_col)
    return df

def _find_header_row(df: pd.DataFrame) -> int:
    """
    Find the header row by scanning rows for 'Description' AND 'Qty/Quantity'.
    Works even if there are blank lines, stylist headers, etc.
    """
    for i in range(len(df)):
        row = df.iloc[i].tolist()
        cells = [(_norm_text(x) or "").lower() for x in row]
        if "description" in cells and any(c in ("qty", "quantity") for c in cells):
            return i
    return -1

def convert_retail_sales(retail_path_or_file) -> pd.DataFrame:
    raw = pd.read_excel(retail_path_or_file, sheet_name="Retail Sales by Team Memb", header=None)

    # PQ effectively keeps these columns (1-based): 2,7,11,13,15
    keep_idx = [1, 6, 10, 12, 14]
    kept = raw.iloc[:, keep_idx].copy()

    # Normalise text in first column to support header scan
    kept.iloc[:, 0] = kept.iloc[:, 0].apply(_norm_text)

    # Drop rows that are completely blank across kept columns
    kept = kept.dropna(how="all").reset_index(drop=True)

    # Find header row
    header_i = _find_header_row(kept)
    if header_i < 0:
        # Provide a helpful sample of what we saw
        sample = kept.head(20).values.tolist()
        raise ValueError(
            "Could not locate header row containing 'Description' and 'Qty/Quantity'. "
            f"First rows sample: {sample}"
        )

    headers = [(_norm_text(x) or "") for x in kept.iloc[header_i].tolist()]
    df2 = kept.iloc[header_i + 1:].copy()
    df2.columns = headers

    # Map columns robustly
    desc_col = _pick(df2.columns, ["Description", "Product", "Item"])
    qty_col = _pick(df2.columns, ["Qty", "Quantity", "QTY"])
    inc_col = _pick(df2.columns, ["Inc Vat", "Inc VAT", "Inc. Vat", "IncVat"])
    exc_col = _pick(df2.columns, ["Exc Vat", "Exc VAT", "Exc. Vat", "ExcVat"])
    gp_col  = _pick(df2.columns, ["Gross Profit", "Profit", "Grossprofit"])

    if desc_col is None or qty_col is None:
        raise ValueError(f"Header row found, but could not map Description/Qty columns. Columns: {list(df2.columns)}")

    rename_map = {desc_col: "Description", qty_col: "Qty"}
    if inc_col: rename_map[inc_col] = "Inc Vat"
    if exc_col: rename_map[exc_col] = "Exc Vat"
    if gp_col:  rename_map[gp_col]  = "Gross Profit"
    df2 = df2.rename(columns=rename_map)

    # Clean + filter like PQ
    df2["Description"] = df2["Description"].apply(_norm_text)
    df2 = df2[df2["Description"].notna()].copy()
    df2 = df2[df2["Description"] != "Grand Total"].copy()
    df2 = df2[df2["Description"] != "Inspired Hair Supplies"].copy()

    df2["Qty"] = pd.to_numeric(df2["Qty"], errors="coerce")

    # Stylist marker rows: Qty is null => stylist header
    stylist_marker = pd.Series(np.where(df2["Qty"].isna(), df2["Description"], np.nan), index=df2.index)

    # Some exports: stylist above => FillDown; others below => FillUp. Do both.
    stylist_up = stylist_marker[::-1].ffill()[::-1]   # FillUp
    stylist_down = stylist_marker.ffill()             # FillDown
    df2["Stylist"] = stylist_up.fillna(stylist_down)

    # Keep product lines only (Qty not null)
    df2 = df2[df2["Qty"].notna()].copy()

    # PQ drops Gross Profit and Exc Vat; we keep Inc Vat optional
    df2["Qty"] = df2["Qty"].astype(int)
    df2["Stylist"] = df2["Stylist"].apply(_norm_text)
    df2["Description"] = df2["Description"].apply(_norm_text)

    keep_cols = ["Stylist", "Description", "Qty"]
    if "Inc Vat" in df2.columns:
        df2["Inc Vat"] = pd.to_numeric(df2["Inc Vat"], errors="coerce")
        keep_cols.append("Inc Vat")

    out = df2[keep_cols].copy()
    return out.reset_index(drop=True)

def build_retail_output(retail_clean: pd.DataFrame, price_df: pd.DataFrame, price_desc_col: str, price_cost_col: str):
    price = price_df.copy()
    price["_key"] = price[price_desc_col].apply(_norm_text)
    price["_cost"] = pd.to_numeric(price[price_cost_col], errors="coerce")

    retail = retail_clean.copy()
    retail["_key"] = retail["Description"].apply(_norm_text)

    merged = retail.merge(price[["_key", "_cost"]], how="left", on="_key")
    merged.rename(columns={"_cost": "Cost"}, inplace=True)
    merged["Invoice Amount"] = merged["Qty"] * merged["Cost"]

    out = merged[["Stylist", "Description", "Qty", "Cost", "Invoice Amount"]].copy()
    out = out.sort_values(["Stylist", "Description"]).reset_index(drop=True)

    missing_cost_df = out[out["Cost"].isna()].copy()
    unmatched_products_df = (
        out[out["Cost"].isna()][["Description"]]
        .drop_duplicates()
        .sort_values("Description")
        .reset_index(drop=True)
    )

    out.attrs["period_label"] = ""
    return out, missing_cost_df, unmatched_products_df
