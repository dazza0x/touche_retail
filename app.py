import io
import zipfile
import hmac

import pandas as pd
import streamlit as st

from transform import convert_retail_sales, load_price_list, build_retail_output
from pdfs import build_retail_statement_pdf

st.set_page_config(page_title="Touche Retail Statements", page_icon="🛍️", layout="centered")

def _require_password():
    if "auth" not in st.secrets or "password" not in st.secrets["auth"]:
        st.error("Password protection is enabled but no password is configured in Streamlit Secrets.")
        st.info('Add this to your app Secrets:\n\n[auth]\npassword = "your-strong-password"')
        st.stop()

    if st.session_state.get("authenticated"):
        return

    st.sidebar.subheader("🔒 Access")
    pw = st.sidebar.text_input("Password", type="password")
    correct = st.secrets["auth"]["password"]

    if pw and hmac.compare_digest(pw, correct):
        st.session_state["authenticated"] = True
        st.sidebar.success("Access granted")
        return
    if pw:
        st.sidebar.error("Incorrect password")
    st.stop()

_require_password()

BRAND = "Touche Hair Caterham"

st.title("🛍️ Touche Hair Caterham — Retail Statements")
st.write(
    "Upload the Retail Sales report and the Inspired Hair Price List. "
    "The app will clean the report (matching your Power Query), join on product description, "
    "validate missing costs/unexpected lines, and generate per-stylist PDF statements (ZIP)."
)

with st.sidebar:
    st.header("Inputs (required)")
    retail_file = st.file_uploader("Retail Sales by Team Member (.xls)", type=["xls"])
    price_file = st.file_uploader("Inspired Hair Price List (.xlsx)", type=["xlsx"])

    st.divider()
    st.header("Options")
    include_excel = st.checkbox("Provide Excel output too", value=True)
    include_cleaned = st.checkbox("Include cleaned tabs in Excel output", value=True)

if retail_file is None or price_file is None:
    st.info("Upload both required files to begin.")
    st.stop()

try:
    price_df = load_price_list(price_file)

    # Let user choose mapping columns (auto-selected defaults)
    with st.sidebar:
        st.subheader("Price list mapping")
        desc_col = st.selectbox("Description column", options=list(price_df.columns), index=price_df.attrs.get("desc_index", 0))
        cost_col = st.selectbox("Cost column", options=list(price_df.columns), index=price_df.attrs.get("cost_index", 0))

    retail_clean = convert_retail_sales(retail_file)
    out_df, missing_cost_df, unmatched_products_df = build_retail_output(
        retail_clean,
        price_df,
        price_desc_col=desc_col,
        price_cost_col=cost_col,
    )

    # Summaries
    st.subheader("Summary")
    total_invoice = out_df["Invoice Amount"].fillna(0).sum()
    st.metric("Total invoice amount", f"£{total_invoice:,.2f}")
    st.metric("Rows", f"{len(out_df):,}")
    st.metric("Stylists", f"{out_df['Stylist'].nunique():,}")

    st.subheader("Preview — retail invoice output")
    st.dataframe(out_df.head(50), use_container_width=True)

    # Validation
    st.subheader("Validation checks")
    if len(missing_cost_df):
        st.warning(f"{len(missing_cost_df):,} row(s) have missing/invalid cost prices (cannot compute invoice amount).")
        with st.expander("Show missing cost rows"):
            st.dataframe(missing_cost_df, use_container_width=True)
    else:
        st.success("No missing cost prices detected.")

    if len(unmatched_products_df):
        st.warning(f"{len(unmatched_products_df):,} distinct product description(s) from the report were not found in the price list.")
        with st.expander("Show unmatched product descriptions"):
            st.dataframe(unmatched_products_df, use_container_width=True)
    else:
        st.success("All product descriptions matched to the price list.")

    # Excel output
    if include_excel:
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            out_df.to_excel(writer, index=False, sheet_name="Retail Invoice Output")
            if include_cleaned:
                retail_clean.to_excel(writer, index=False, sheet_name="Retail Cleaned")
                price_df.to_excel(writer, index=False, sheet_name="Price List")
            if len(missing_cost_df):
                missing_cost_df.to_excel(writer, index=False, sheet_name="Missing Costs")
            if len(unmatched_products_df):
                unmatched_products_df.to_excel(writer, index=False, sheet_name="Unmatched Products")
        excel_buf.seek(0)

        st.download_button(
            "Download Excel output (.xlsx)",
            data=excel_buf,
            file_name="Retail Statements Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # PDFs
    st.subheader("PDF statements")
    st.caption("One PDF per stylist, packaged into a single ZIP.")

    if st.button("Generate ZIP of stylist PDFs"):
        period_label = out_df.attrs.get("period_label", "")
        stylists = sorted(out_df["Stylist"].dropna().astype(str).unique())

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for stylist in stylists:
                s_df = out_df[out_df["Stylist"] == stylist].copy()
                pdf_bytes = build_retail_statement_pdf(
                    brand=BRAND,
                    stylist=stylist,
                    period_label=period_label,
                    retail_df=s_df,
                )
                safe = "".join(ch for ch in stylist if ch.isalnum() or ch in (" ", "-", "_")).strip().replace(" ", "_")
                z.writestr(f"{safe}.pdf", pdf_bytes)

        zip_buf.seek(0)
        st.download_button(
            "Download ZIP of PDFs",
            data=zip_buf,
            file_name="Retail Stylist Statements.zip",
            mime="application/zip",
        )

except Exception as e:
    st.error("Processing failed.")
    st.exception(e)
