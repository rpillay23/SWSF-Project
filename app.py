import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import requests

# === Page Configuration ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# === Custom Styling ===
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, sans-serif;
        background-color: white;
        color: #003366;
    }

    .title-box {
        background-color: #B0C4DE;
        border: 5px solid #003366;
        border-radius: 10px;
        padding: 1em;
        text-align: center;
        margin-bottom: 20px;
    }

    .title-box h1 {
        font-size: 28px;
        font-weight: bold;
        margin: 0;
    }

    .subtitle {
        font-size: 16px;
        color: #003366;
        text-align: center;
        font-stretch: condensed;
        margin-bottom: 30px;
    }
    </style>
""", unsafe_allow_html=True)

# === Title ===
st.markdown('<div class="title-box"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Modular Platform for Traditional & Alternative Investment Evaluation</div>', unsafe_allow_html=True)

# === Helper: Sanitize Strings ===
def sanitize_string(s):
    if isinstance(s, str):
        return s.replace("–", "-").replace("’", "'").replace("“", '"').replace("”", '"').replace("•", "-").replace("©", "(c)")
    return s

# === Fetch Real-Time Market Data ===
def fetch_real_time_data():
    tickers = ["AAPL", "AMZN", "GOOGL", "META", "NVDA", "TSLA", "^GSPC", "^IXIC", "^DJI"]
    url = f"https://financialmodelingprep.com/api/v3/quote/{','.join(tickers)}?apikey=YOUR_API_KEY"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.json()
    except Exception as e:
        st.sidebar.error(f"API Error: {e}")
    return []

# === Sidebar: Live Market Tracker ===
st.sidebar.header("📈 Real-Time Market Data")
live_data = fetch_real_time_data()
if live_data:
    for item in live_data:
        name = item.get('name', item['symbol'])
        price = item.get('price', 0)
        change = item.get('change', 0)
        percent = item.get('changesPercentage', 0)
        st.sidebar.metric(label=name, value=f"${price:.2f}", delta=f"{change:+.2f} ({percent:+.2f}%)")
else:
    st.sidebar.warning("Unable to fetch market data.")

# === Main Application ===
try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    st.subheader("Investment Data Table")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    st.divider()

    # Portfolio Metrics
    st.subheader("Portfolio Metrics")
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("Avg Return (%)", f"{edited_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk", f"{edited_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate", f"{edited_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{edited_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{edited_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees", f"{edited_df['Fees (%)'].mean():.2f}%")
    col7.metric("Min Investment", f"${edited_df['Minimum Investment ($)'].mean():,.0f}")
    st.divider()

    # Charts
    st.subheader("Expected Return by Investment")
    st.bar_chart(edited_df.set_index("Investment Name")["Expected Return (%)"])

    st.subheader("Liquidity vs. Volatility")
    st.scatter_chart(edited_df, x="Volatility (1–10)", y="Liquidity (1–10)", size="Expected Return (%)", color="Category")
    st.divider()

    # Filters
    st.subheader("Filter Your Portfolio")
    time_options = ["All"] + sorted(edited_df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_options = ["All", "Yes", "No"]

    time_filter = st.selectbox("Select Time Horizon", time_options)
    hedge_filter = st.selectbox("Inflation Hedge?", hedge_options)
    min_inv_filter = st.slider("Minimum Investment ($)",
        int(edited_df["Minimum Investment ($)"].min()),
        int(edited_df["Minimum Investment ($)"].max()),
        int(edited_df["Minimum Investment ($)"].min()))

    filtered_df = edited_df.copy()
    if time_filter != "All":
        filtered_df = filtered_df[filtered_df["Time Horizon (Short/Medium/Long)"] == time_filter]
    if hedge_filter != "All":
        filtered_df = filtered_df[filtered_df["Inflation Hedge (Yes/No)"] == hedge_filter]
    filtered_df = filtered_df[filtered_df["Minimum Investment ($)"] >= min_inv_filter]

    st.subheader("Filtered Investment Table")
    st.dataframe(filtered_df, use_container_width=True)
    st.divider()

    # === Export Reports ===
    st.subheader("📤 Export Reports")

    def create_ppt(df):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Investment Overview"
        slide.placeholders[1].text = "Portfolio Summary"

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Averages"
        slide.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])

        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90)
        plt.tight_layout()
        chart_path = "ppt_chart.png"
        plt.savefig(chart_path)
        plt.close()

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Return Chart"
        slide.shapes.add_picture(chart_path, Inches(1), Inches(1.5), width=Inches(8))

        ppt_file = "Investment_Presentation.pptx"
        prs.save(ppt_file)
        return ppt_file

    def create_docx(df):
        doc = Document()
        doc.add_heading("Investment Summary", 0)

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        doc.add_heading("Portfolio Averages", level=1)
        for k, v in avg.items():
            doc.add_paragraph(f"{k}: {v}")

        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90)
        plt.tight_layout()
        chart_path = "docx_chart.png"
        plt.savefig(chart_path)
        plt.close()

        doc.add_heading("Return Chart", level=1)
        doc.add_picture(chart_path, width=DocxInches(6.5))

        docx_file = "Investment_Summary.docx"
        doc.save(docx_file)
        return docx_file

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Generate PowerPoint"):
            ppt_file = create_ppt(filtered_df)
            with open(ppt_file, "rb") as f:
                st.download_button("Download PowerPoint", f, file_name=ppt_file)

    with col2:
        if st.button("Generate Word Report"):
            docx_file = create_docx(filtered_df)
            with open(docx_file, "rb") as f:
                st.download_button("Download Word Report", f, file_name=docx_file)

except Exception as e:
    st.error(f"⚠️ An error occurred: {e}")
