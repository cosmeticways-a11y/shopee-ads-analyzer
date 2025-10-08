import io
import streamlit as st
import pandas as pd
import numpy as np

from utils.compute import (
    load_shopee_ads_csv,
    load_costing_table,
    apply_v5_logic,
    split_deleted_and_active,
    compute_kpis,
    style_preview_df,
)
from utils.excel_writer import build_excel_bytes

st.set_page_config(page_title="Shopee Ads Analyzer", layout="wide")

st.title("Shopee Ads Analyzer")

# --- Upload ---
st.header("1) Upload")
col1, col2 = st.columns(2)
with col1:
    ads_file = st.file_uploader(
        "File 1: Shopee Ads CSV (original export)", type=["csv"], accept_multiple_files=False
    )
with col2:
    costing_file = st.file_uploader(
        "File 2: Product Costing (txt or csv)", type=["txt", "csv"], accept_multiple_files=False
    )

# --- Controls ---
st.header("2) Controls")
multiplier = st.number_input("Profit Multiplier (for Suggested ROAS)", min_value=0.1, value=1.25, step=0.05)

# --- Generate ---
st.header("3) Generate")
generate = st.button("Generate Report", type="primary", use_container_width=True)

if generate:
    if ads_file is None or costing_file is None:
        st.error("Please upload both files: Shopee Ads CSV and Product Costing (txt/csv).")
        st.stop()

    try:
        ads_raw, detected_delim = load_shopee_ads_csv(ads_file)
    except Exception as e:
        st.error(f"Failed to read Shopee Ads CSV: {e}")
        st.stop()

    try:
        costing_df = load_costing_table(costing_file)
    except Exception as e:
        st.error(f"Failed to read Product Costing file: {e}")
        st.stop()

    try:
        ads_with_logic = apply_v5_logic(ads_raw, costing_df, multiplier=multiplier)
    except Exception as e:
        st.error(f"Failed to apply computations: {e}")
        st.stop()

    df_active, df_deleted, df_unmatched = split_deleted_and_active(ads_with_logic)

    st.header("4) KPIs")
    k = compute_kpis(df_active)
    kpi1, kpi2, kpi3, kpi4, kpi5, kpi6 = st.columns(6)
    kpi1.metric("Total GMV", f"₱{k['total_gmv']:,.2f}")
    kpi2.metric("Total Expense", f"₱{k['total_expense']:,.2f}")
    kpi3.metric("Total Net Profit", f"₱{k['total_net_profit']:,.2f}")
    kpi4.metric("Winning Ads", f"{k['winning_cnt']}")
    kpi5.metric("Losing Ads", f"{k['losing_cnt']}")
    kpi6.metric("Average ROAS", f"{k['avg_roas']:.2f}")

    st.header("5) Preview (first 50 rows)")
    st.caption(f"Detected CSV delimiter: “{detected_delim}”")
    preview = df_active.head(50).copy()
    st.dataframe(style_preview_df(preview), use_container_width=True)

    st.header("6) Download")
    try:
        xlsx_bytes = build_excel_bytes(
            df_active=df_active,
            df_deleted=df_deleted,
            df_unmatched=df_unmatched,
            multiplier=multiplier,
            source_name=getattr(ads_file, "name", "shopee_ads.csv"),
        )
        st.download_button(
            label="Download Excel (.xlsx)",
            data=xlsx_bytes,
            file_name="Shopee_Ads_Analyzer_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Failed to build Excel: {e}")
