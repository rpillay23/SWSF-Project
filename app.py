import streamlit as st
import pandas as pd

st.set_page_config(page_title="HNW Investment Matrix", layout="wide")
st.title("📊 HNW Investment Matrix (Comprehensive & Editable)")

try:
    # Load Excel file
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")

    # Editable Table
    st.subheader("🔧 Edit Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    st.divider()

    # Portfolio Metrics
    st.subheader("📈 Portfolio Averages & Totals")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Avg Return (%)", f"{edited_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk (1–10)", f"{edited_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate (%)", f"{edited_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{edited_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{edited_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees (%)", f"{edited_df['Fees (%)'].mean():.2f}%")

    col7, _ = st.columns(2)
    col7.metric("💰 Avg Min Investment", f"${edited_df['Minimum Investment ($)'].mean():,.0f}")

    st.divider()

    # Charts
    st.subheader("📊 Expected Return by Investment")
    st.bar_chart(edited_df.set_index("Investment Name")["Expected Return (%)"])

    st.subheader("📊 Liquidity vs. Volatility")
    st.scatter_chart(
        edited_df,
        x="Volatility (1–10)",
        y="Liquidity (1–10)",
        size="Expected Return (%)",
        color="Category"
    )

    st.divider()

    # Filters
    st.subheader("🎯 Filter by Time Horizon, Inflation Hedge, or Min Investment")
    time_options = ["All"] + sorted(edited_df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_options = ["All", "Yes", "No"]

    time_filter = st.selectbox("Select Time Horizon", time_options)
    hedge_filter = st.selectbox("Inflation Hedge?", hedge_options)
    min_inv_filter = st.slider("Minimum Investment ($)", 
                               int(edited_df["Minimum Investment ($)"].min()), 
                               int(edited_df["Minimum Investment ($)"].max()), 
                               int(edited_df["Minimum Investment ($)"].min()))

    # Apply Filters
    filtered_df = edited_df.copy()
    if time_filter != "All":
        filtered_df = filtered_df[filtered_df["Time Horizon (Short/Medium/Long)"] == time_filter]
    if hedge_filter != "All":
        filtered_df = filtered_df[filtered_df["Inflation Hedge (Yes/No)"] == hedge_filter]
    if min_inv_filter:
        filtered_df = filtered_df[filtered_df["Minimum Investment ($)"] >= min_inv_filter]

    st.subheader("📄 Filtered Investment Table")
    st.dataframe(filtered_df, use_container_width=True)

except Exception as e:
    st.error(f"⚠️ Error loading Excel file: {e}")
