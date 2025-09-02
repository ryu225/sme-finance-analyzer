import os
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt

# Minimal, professional comments only

st.set_page_config(page_title="SME Finance Analyzer", layout="wide")
st.title("📊 SME Finance Analyzer")
st.caption("Upload transactions, review monthly P&L, explore trends, and export a concise PowerPoint.")

# ---------- Helpers ----------
REQUIRED_COLS = ["date", "type", "category", "description", "amount"]
OPTIONAL_COLS = ["tax"]

def load_sample() -> pd.DataFrame:
    sample_path = "data/sample_transactions.csv"
    if os.path.exists(sample_path):
        df = pd.read_csv(sample_path)
    else:
        # tiny inline fallback if file is missing
        df = pd.DataFrame({
            "date": ["2025-01-01"], "type": ["sale"], "category": ["online"],
            "description": ["sample"], "amount": [1000.0], "tax": [50.0]
        })
    return df

def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # normalize columns
    df.columns = [c.strip().lower() for c in df.columns]
    # ensure required columns exist
    for col in REQUIRED_COLS:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")
    # parse date
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    # coerce numeric amounts
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df.dropna(subset=["amount"])
    # type normalization
    df["type"] = df["type"].str.lower().str.strip()
    # month key
    df["month"] = df["date"].dt.to_period("M").astype(str)
    # margin sign convention: positive sales, negative expenses
    return df

def summarize(df: pd.DataFrame) -> dict:
    revenue = df.loc[df["type"] == "sale", "amount"].sum()
    expenses = -df.loc[df["type"] == "expense", "amount"].sum()  # amounts are negative, invert
    profit = revenue - expenses
    margin = (profit / revenue * 100.0) if revenue != 0 else np.nan
    return dict(revenue=revenue, expenses=expenses, profit=profit, margin=margin)

def monthly_pnl(df: pd.DataFrame) -> pd.DataFrame:
    rev = df[df["type"] == "sale"].groupby("month")["amount"].sum().rename("revenue")
    exp = -df[df["type"] == "expense"].groupby("month")["amount"].sum().rename("expenses")
    pnl = pd.concat([rev, exp], axis=1).fillna(0.0)
    pnl["profit"] = pnl["revenue"] - pnl["expenses"]
    pnl["margin_%"] = np.where(pnl["revenue"] != 0, pnl["profit"] / pnl["revenue"] * 100, np.nan)
    pnl = pnl.reset_index()
    return pnl

def expense_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    exp = df[df["type"] == "expense"].copy()
    if exp.empty:
        return pd.DataFrame({"category": [], "total": []})
    out = (-exp.groupby("category")["amount"].sum()).reset_index().rename(columns={"amount": "total"})
    out = out.sort_values("total", ascending=False)
    return out

def fig_to_png_bytes(fig) -> bytes:
    # requires 'kaleido' in requirements
    buf = BytesIO()
    fig.write_image(buf, format="png", scale=2)
    return buf.getvalue()

def build_ppt(pnl_df: pd.DataFrame, kpis: dict, fig1, fig2) -> BytesIO:
    prs = Presentation()
    # Title slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = "SME Finance Analyzer"
    slide.placeholders[1].text = f"Auto-generated on {datetime.now():%Y-%m-%d}"

    # KPI slide
    layout = prs.slide_layouts[1]
    slide2 = prs.slides.add_slide(layout)
    slide2.shapes.title.text = "Key Metrics"
    body = slide2.placeholders[1].text_frame
    body.clear()
    for line in [
        f"Revenue: {kpis['revenue']:,.0f}",
        f"Expenses: {kpis['expenses']:,.0f}",
        f"Profit: {kpis['profit']:,.0f}",
        f"Margin: {kpis['margin']:.1f}%" if not np.isnan(kpis['margin']) else "Margin: n/a",
    ]:
        p = body.add_paragraph()
        p.text = line
        p.level = 0

    # Charts slide (monthly trends)
    chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
    chart_slide.shapes.title.text = "Monthly Trends"
    left, top = Inches(0.5), Inches(1.5)
    w, h = Inches(4.5), Inches(3)
    png1 = fig_to_png_bytes(fig1)
    chart_slide.shapes.add_picture(BytesIO(png1), left, top, width=w, height=h)

    left2 = Inches(5.2)
    png2 = fig_to_png_bytes(fig2)
    chart_slide.shapes.add_picture(BytesIO(png2), left2, top, width=w, height=h)

    # P&L table slide
    tbl_slide = prs.slides.add_slide(prs.slide_layouts[5])
    tbl_slide.shapes.title.text = "P&L by Month"
    rows, cols = len(pnl_df) + 1, 4
    left, top, width, height = Inches(0.5), Inches(1.5), Inches(9), Inches(0.8 + 0.3 * rows)
    shape = tbl_slide.shapes.add_table(rows, cols, left, top, width, height)
    table = shape.table
    headers = ["Month", "Revenue", "Expenses", "Profit"]
    for j, htxt in enumerate(headers):
        table.cell(0, j).text = htxt
    for i, (_, r) in enumerate(pnl_df.iterrows(), start=1):
        table.cell(i, 0).text = str(r["month"])
        table.cell(i, 1).text = f"{r['revenue']:,.0f}"
        table.cell(i, 2).text = f"{r['expenses']:,.0f}"
        table.cell(i, 3).text = f"{r['profit']:,.0f}"

    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# ---------- Data input ----------
st.sidebar.subheader("Data")
uploaded = st.sidebar.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
if uploaded:
    if uploaded.name.lower().endswith(".xlsx"):
        df = pd.read_excel(uploaded)
    else:
        df = pd.read_csv(uploaded)
else:
    df = load_sample()

try:
    df = clean_df(df)
except Exception as e:
    st.error(f"Data error: {e}")
    st.stop()

# ---------- KPIs ----------
kpis = summarize(df)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Revenue", f"{kpis['revenue']:,.0f}")
c2.metric("Expenses", f"{kpis['expenses']:,.0f}")
c3.metric("Profit", f"{kpis['profit']:,.0f}")
c4.metric("Margin", f"{kpis['margin']:.1f}%" if not np.isnan(kpis['margin']) else "n/a")

# ---------- Visuals ----------
pnl = monthly_pnl(df)
exp_cat = expense_breakdown(df)

st.subheader("Monthly Revenue vs Expenses")
fig_rev_exp = px.bar(
    pnl, x="month", y=["revenue", "expenses"], barmode="group",
    title="Revenue and Expenses by Month"
)
st.plotly_chart(fig_rev_exp, use_container_width=True)

st.subheader("Profit by Month")
fig_profit = px.bar(pnl, x="month", y="profit", title="Monthly Profit")
st.plotly_chart(fig_profit, use_container_width=True)

st.subheader("Expense Breakdown by Category")
if not exp_cat.empty:
    fig_exp = px.pie(exp_cat, names="category", values="total", title="Expenses by Category")
    st.plotly_chart(fig_exp, use_container_width=True)
else:
    st.info("No expenses found in the dataset.")

# ---------- Export ----------
st.subheader("Export")
col_a, col_b = st.columns(2)
with col_a:
    st.download_button(
        "Download Clean CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="transactions_clean.csv",
        mime="text/csv",
    )
with col_b:
    try:
        ppt_bytes = build_ppt(pnl, kpis, fig_rev_exp, fig_profit)
        st.download_button(
            "Export PowerPoint (P&L + Charts)",
            data=ppt_bytes,
            file_name="SME_Finance_Analyzer.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except Exception as e:
        st.warning("PowerPoint export unavailable (check 'kaleido' and 'python-pptx' installs).")

# ---------- Raw table ----------
with st.expander("Preview raw transactions"):
    st.dataframe(df.head(200))
