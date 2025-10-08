import io
import os
import sys
import streamlit as st
import pandas as pd
import numpy as np

# Ensure utils is recognized as a package
sys.path.append(os.path.join(os.path.dirname(__file__), 'utils'))

from utils import compute, excel_writer

st.header("1) Upload")
col1, col2 = st.columns(2)
with col1:
    ads_file = st.file_uploader("File 1: Shopee Ads CSV (original export)", type=["csv"])
with col2:
    costing_file = st.file_uploader("File 2: Product Costing (txt or csv)", type=None)

st.header("2) Controls")
multiplier = st.number_input("Profit Multiplier (for Suggested ROAS)", min_value=0.1, value=1.25, step=0.05)

st.header("3) Generate")
if st.button("Generate Report", type="primary", use_container_width=True):
    if ads_file is None or costing_file is None:
        st.error("Please upload both files: Shopee Ads CSV and Product Costing.")
        st.stop()

    try:
        ads_raw = compute.load_ads(ads_file)
        costing_df = compute.load_costing(costing_file)
        result_df = compute.compute_logic(ads_raw, costing_df, multiplier)
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

    st.success("Computation completed successfully!")

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total GMV", f"₱{result_df['GMV'].sum():,.2f}")
    c2.metric("Total Expense", f"₱{result_df['Expense'].sum():,.2f}")
    c3.metric("Total Net Profit", f"₱{result_df['Net Profit'].sum():,.2f}")
    c4.metric("Avg ROAS", f"{result_df['ROAS'].replace([np.inf, -np.inf], np.nan).mean():.2f}")

    st.subheader("Preview (first 50 rows)")
    st.dataframe(result_df.head(50), use_container_width=True)

    st.subheader("Download Excel")
    excel_data = excel_writer.to_excel_bytes(result_df)
    st.download_button(
        label="Download Excel (.xlsx)",
        data=excel_data,
        file_name="Shopee_Ads_Analyzer_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.caption("Tip: You can test with the sample templates inside the repo's /templates folder.")
