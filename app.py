import io
import streamlit as st
import pandas as pd
import numpy as np
from utils import compute, excel_writer


# -----------------------------------------------------------
# Streamlit Page Config
# -----------------------------------------------------------
st.set_page_config(page_title="Shopee Ads Analyzer", layout="wide")
st.title("Shopee Ads Analyzer")


# -----------------------------------------------------------
# Upload Section
# -----------------------------------------------------------
st.header("1) Upload")
col1, col2 = st.columns(2)

with col1:
    ads_file = st.file_uploader(
        "File 1: Shopee Ads CSV (original export)",
        type=["csv"],
        accept_multiple_files=False,
    )

with col2:
    costing_file = st.file_uploader(
        "File 2: Product Costing (txt or csv)",
        type=None,
        accept_multiple_files=False,
    )


# -----------------------------------------------------------
# Controls
# -----------------------------------------------------------
st.header("2) Controls")
multiplier = st.number_input(
    "Profit Multiplier (for Suggested ROAS)",
    min_value=0.1,
    value=1.25,
    step=0.05,
)


# -----------------------------------------------------------
# Generate Button
# -----------------------------------------------------------
st.header("3) Generate")
generate = st.button("Generate Report", type="primary", use_container_width=True)

if not generate:
    st.stop()

if ads_file is None or costing_file is None:
    st.error("Please upload both files: Shopee Ads CSV and Product Costing (txt/csv).")
    st.stop()


# -----------------------------------------------------------
# Load and Process Data
# -----------------------------------------------------------
try:
    ads_raw, detected_delim = compute.load_shopee_ads_csv(ads_file)
except Exception as e:
    st.error(f"Failed to read Shopee Ads CSV: {e}")
    st.stop()

try:
    costing_df = compute.load_costing_table(costing_file)
except Exception as e:
    st.error(f"Failed to read Product Costing file: {e}")
    st.stop()

try:
    ads_with_logic = compute.apply_v5_logic(
        ads_raw, costing_df, multiplier=multiplier
    )
except Exception as e:
    st.error(f"Failed to apply computations: {e}")
    st.stop()

df_active, df_deleted, df_unmatched = compute.split_deleted_and_active(ads_with_logic)


# -----------------------------------------------------------
# KPIs
# -----------------------------------------------------------
st.header("4) KPIs")
k = compute.compute_kpis(df_active)
cols = st.columns(6)
cols[0].metric("Total GMV", f"₱{k['total_gmv']:,.2f}")
cols[1].metric("Total Expense", f"₱{k['total_expense']:,.2f}")
cols[2].metric("Total Net Profit", f"₱{k['total_net_profit']:,.2f}")
cols[3].metric("Winning Ads", f"{k['winning_cnt']}")
cols[4].metric("Losing Ads", f"{k['losing_cnt']}")
cols[5].metric("Average ROAS", f"{k['avg_roas']:.2f}")


# -----------------------------------------------------------
# Preview Table
# -----------------------------------------------------------
st.header("5) Preview (first 50 rows)")
st.caption(f"Detected CSV delimiter: “{detected_delim}”")
preview = df_active.head(50).copy()
st.dataframe(compute.style_preview_df(preview), use_container_width=True)


# -----------------------------------------------------------
# Download Excel
# -----------------------------------------------------------
st.header("6) Download")
try:
    xlsx_bytes = excel_writer.build_excel_bytes(
        df_active=df_active,
        df_deleted=df_deleted,
        df_unmatched=df_unmatched,
        multiplier=multiplier,
        source_name=getattr(ads_file, "name", "Shopee-Ads-Report.csv"),
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
