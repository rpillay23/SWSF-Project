import streamlit as st
import pandas as pd

st.set_page_config(page_title="HNW Investment Matrix", layout="wide")

st.title("📊 HNW Investment Matrix (Editable)")

try:
    # Load the Excel file
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")

    st.subheader("🔧 Edit Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    st.divider()

    # Averages
    st.subheader("📈 Portfolio Averages")
    avg_return = edited_df["Expected Return (%)"].mean().round(2)
    avg_risk = edited_df["Risk Level (1-10)"].mean().round(2)
    avg_cap_rate = edited_df["Cap Rate (%)"].mean().round(2)

    col1, col2, col3 = st.columns(3)
    col1.metric("Avg Return (%)", f"{avg_return}%")
    col2.metric("Avg Risk (1-10)", f"{avg_risk}")
    col3.metric("Avg Cap Rate (%)", f"{avg_cap_rate}")

    st.divider()

    # Chart
    st.subheader("📊 Expected Return by Investment")
    st.bar_chart(edited_df.set_index("Investment Name")["Expected Return (%)"])

except Exception as e:
    st.error(f"⚠️ Error loading Excel file: {e}")
