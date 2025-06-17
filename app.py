import streamlit as st
import pandas as pd

st.title("📊 HNW Investment Matrix")

try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")

    st.subheader("Full Investment Table")
    st.dataframe(df)

    st.subheader("Expected Return Chart")
    st.bar_chart(df.set_index("Investment Name")["Expected Return (%)"])

except Exception as e:
    st.error(f"⚠️ Error loading Excel file: {e}")
