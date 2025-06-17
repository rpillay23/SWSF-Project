import streamlit as st
import pandas as pd

# Load Excel file
df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")

# Title
st.title("HNW Investment Matrix Dashboard")

# Show full table
st.subheader("Investment Overview")
st.dataframe(df)

# Filters
category = st.selectbox("Filter by Category", ["All"] + list(df["Category"].unique()))
if category != "All":
    df = df[df["Category"] == category]

# Metrics
st.subheader("Averages")
st.write(df[["Expected Return (%)", "Risk Level (1-10)", "Cap Rate (%)"]].mean().round(2))

# Chart
st.subheader("Expected Return by Investment")
st.bar_chart(df.set_index("Investment Name")["Expected Return (%)"])
