import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from fpdf import FPDF
import os

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

    st.divider()
    st.subheader("📥 Generate Reports")

    # === Utility to sanitize strings for PDF
    def sanitize_string(s):
        if isinstance(s, str):
            return (
                s.replace("–", "-")
                 .replace("’", "'")
                 .replace("“", '"')
                 .replace("”", '"')
                 .replace("•", "-")
                 .replace("©", "(c)")
            )
        return s

    # === PowerPoint Generator
    def create_ppt(df):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Comprehensive Investment Overview"
        slide.placeholders[1].text = "Alternative & Traditional Investments"

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Portfolio Averages"
        slide.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])

        chart_file = "streamlit_chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90)
        plt.tight_layout()
        plt.savefig(chart_file)
        plt.close()

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Expected Return Chart"
        slide.shapes.add_picture(chart_file, Inches(1), Inches(1.5), width=Inches(8))

        ppt_file = "HNW_Investment_Presentation.pptx"
        prs.save(ppt_file)
        return ppt_file

    # === PDF Generator
    def create_pdf(df):
        df = df.applymap(sanitize_string)
        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "HNW Investment Summary", ln=True, align="C")
        pdf.set_font("Arial", size=12)
        pdf.ln(10)

        pdf.cell(0, 10, "Portfolio Averages:", ln=True)
        for k, v in avg.items():
            pdf.cell(0, 10, f"{k}: {v}", ln=True)

        chart_file = "streamlit_chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90)
        plt.tight_layout()
        plt.savefig(chart_file)
        plt.close()

        pdf.image(chart_file, w=170)
        pdf_file = "HNW_Investment_Summary.pdf"
        pdf.output(pdf_file)
        return pdf_file

    # === Interactive Buttons
    col1, col2 = st.columns(2)

    with col1:
        if st.button("📽 Generate PowerPoint"):
            ppt_file = create_ppt(filtered_df)
            with open(ppt_file, "rb") as f:
                st.download_button("Download PowerPoint", f, file_name=ppt_file)

    with col2:
        if st.button("📄 Generate PDF Summary"):
            pdf_file = create_pdf(filtered_df)
            with open(pdf_file, "rb") as f:
                st.download_button("Download PDF", f, file_name=pdf_file)

except Exception as e:
    st.error(f"⚠️ Error loading Excel file: {e}")
