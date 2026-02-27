import io
import pandas as pd
from typing import Optional
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

def _money(x) -> str:
    try:
        if pd.isna(x):
            return ""
        return f"£{float(x):,.2f}"
    except Exception:
        return ""

def build_retail_statement_pdf(
    brand: str,
    stylist: str,
    period_label: str,
    retail_df: pd.DataFrame,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18*mm,
        rightMargin=18*mm,
        topMargin=16*mm,
        bottomMargin=16*mm,
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"<b>{brand} — {stylist}</b>", styles["Title"]))
    story.append(Spacer(1, 6))
    if period_label:
        story.append(Paragraph(f"Statement period: {period_label}", styles["Normal"]))
    story.append(Spacer(1, 12))

    story.append(Paragraph(f"<b>{stylist} Retail Statement</b>", styles["Heading2"]))
    story.append(Spacer(1, 6))

    df = retail_df.copy()
    df = df[["Description", "Qty", "Cost", "Invoice Amount"]].copy()
    df["Cost"] = df["Cost"].apply(_money)
    df["Invoice Amount"] = df["Invoice Amount"].apply(_money)

    data = [["Product", "Qty", "Cost", "Invoice Amount"]] + df.values.tolist()

    t = Table(data, colWidths=[96*mm, 14*mm, 28*mm, 28*mm])
    t.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("ALIGN", (2,1), (3,-1), "RIGHT"),
    ]))
    story.append(t)

    total_qty = int(retail_df["Qty"].fillna(0).sum())
    total_inv = retail_df["Invoice Amount"].fillna(0).sum()
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<b>Totals:</b> Qty {total_qty} | Invoice {_money(total_inv)}", styles["Normal"]))

    doc.build(story)
    return buf.getvalue()
