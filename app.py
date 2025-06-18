import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import yfinance as yf

# === Page Config and Styling ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, sans-serif;
        background-color: white;
        color: #003366;
    }

    h1, h2, h3, .stMarkdown, p {
        color: #003366;
        font-weight: bold;
        border: none !important;
        outline: none !important;
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
        margin-bottom: 30px;
    }

    .stButton > button {
        background-color: #003366;
        color: white;
        border-radius: 6px;
        padding: 0.5em 1em;
        border: none;
        font-weight: bold;
    }

    .stButton > button:hover {
        background-color: #0055a5;
    }

    .stMetric {
        background-color: #f0f8ff;
        padding: 1em;
        border-radius: 8px;
        color: #003366;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# === Title & Subtitle ===
st.markdown('<div class="title-box"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automated Software for Traditional and Alternate Investment Analysis Designed for Portfolio Management and Building a Modular Sustainable Wealth Strategy Framework (SWSF)</div>', unsafe_allow_html=True)

# === Helper ===
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

@st.cache_data(ttl=600)
def get_index_data(ticker):
    index = yf.Ticker(ticker)
    hist = index.history(period="1mo")
    hist.reset_index(inplace=True)
    return hist

# === Sidebar: Real-Time Market Data ===
st.sidebar.header("Real-Time Market Data")

# S&P 500
sp500_data = get_index_data("^GSPC")
if not sp500_data.empty:
    latest_sp500 = sp500_data.iloc[-1]
    prev_sp500 = sp500_data.iloc[-2]
    latest_sp500_close = latest_sp500['Close']
    prev_sp500_close = prev_sp500['Close']
    change_sp500 = latest_sp500_close - prev_sp500_close
    pct_change_sp500 = (change_sp500 / prev_sp500_close) * 100

    st.sidebar.metric("S&P 500", f"${latest_sp500_close:.2f}", f"{change_sp500:+.2f} ({pct_change_sp500:+.2f}%)")

    fig, ax = plt.subplots(figsize=(4, 2.5))  # smaller figure
    ax.plot(sp500_data['Date'], sp500_data['Close'], label='S&P 500 Close', color='#003366')
    ax.set_xlabel("")
    ax.set_ylabel("Price ($)")
    ax.set_title("")
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    plt.tight_layout()
    st.sidebar.pyplot(fig)
else:
    st.sidebar.warning("Failed to fetch S&P 500 data.")

# Nasdaq
nasdaq_data = get_index_data("^IXIC")
if not nasdaq_data.empty:
    latest_nasdaq = nasdaq_data.iloc[-1]
    prev_nasdaq = nasdaq_data.iloc[-2]
    latest_nasdaq_close = latest_nasdaq['Close']
    prev_nasdaq_close = prev_nasdaq['Close']
    change_nasdaq = latest_nasdaq_close - prev_nasdaq_close
    pct_change_nasdaq = (change_nasdaq / prev_nasdaq_close) * 100

    st.sidebar.metric("Nasdaq", f"${latest_nasdaq_close:.2f}", f"{change_nasdaq:+.2f} ({pct_change_nasdaq:+.2f}%)")

    fig, ax = plt.subplots(figsize=(4, 2.5))  # smaller figure
    ax.plot(nasdaq_data['Date'], nasdaq_data['Close'], label='Nasdaq Close', color='#003366')
    ax.set_xlabel("")
    ax.set_ylabel("Price ($)")
    ax.set_title("")
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    plt.tight_layout()
    st.sidebar.pyplot(fig)
else:
    st.sidebar.warning("Failed to fetch Nasdaq data.")

# Dow Jones
dowjones_data = get_index_data("^DJI")
if not dowjones_data.empty:
    latest_dowjones = dowjones_data.iloc[-1]
    prev_dowjones = dowjones_data.iloc[-2]
    latest_dowjones_close = latest_dowjones['Close']
    prev_dowjones_close = prev_dowjones['Close']
    change_dowjones = latest_dowjones_close - prev_dowjones_close
    pct_change_dowjones = (change_dowjones / prev_dowjones_close) * 100

    st.sidebar.metric("Dow Jones", f"${latest_dowjones_close:.2f}", f"{change_dowjones:+.2f} ({pct_change_dowjones:+.2f}%)")

    fig, ax = plt.subplots(figsize=(4, 2.5))  # smaller figure
    ax.plot(dowjones_data['Date'], dowjones_data['Close'], label='Dow Jones Close', color='#003366')
    ax.set_xlabel("")
    ax.set_ylabel("Price ($)")
    ax.set_title("")
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.tick_params(axis='x', rotation=45, labelsize=8)
    ax.tick_params(axis='y', labelsize=8)
    plt.tight_layout()
    st.sidebar.pyplot(fig)
else:
    st.sidebar.warning("Failed to fetch Dow Jones data.")

st.divider()

try:
    # === Load Excel Data ===
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # === Editable Table ===
    st.subheader("Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    st.divider()

    # === Portfolio Metrics: 7 columns inline ===
    st.subheader("Portfolio Averages and Totals")
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("Avg Return (%)", f"{edited_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk (1–10)", f"{edited_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate (%)", f"{edited_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{edited_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{edited_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees (%)", f"{edited_df['Fees (%)'].mean():.2f}%")
    col7.metric("Avg Min Investment", f"${edited_df['Minimum Investment ($)'].mean():,.0f}")
    st.divider()

    # === Charts ===
    st.subheader("Expected Return by Investment")
    st.bar_chart(edited_df.set_index("Investment Name")["Expected Return (%)"])

    st.subheader("Liquidity vs. Volatility")
    st.scatter_chart(
        edited_df,
        x="Volatility (1–10)",
        y="Liquidity (1–10)",
        size="Expected Return (%)",
        color="Category"
    )
    st.divider()


try:
    # === Load Excel Data ===
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # === Editable Table ===
    st.subheader("Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    st.divider()

    # === Metrics ===
    st.subheader("Portfolio Averages and Totals")
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("Avg Return (%)", f"{edited_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk (1–10)", f"{edited_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate (%)", f"{edited_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{edited_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{edited_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees (%)", f"{edited_df['Fees (%)'].mean():.2f}%")
    col7.metric("Avg Min Investment", f"${edited_df['Minimum Investment ($)'].mean():,.0f}")
    st.divider()

    # === Charts ===
    st.subheader("Expected Return by Investment")
    st.bar_chart(edited_df.set_index("Investment Name")["Expected Return (%)"])

    st.subheader("Liquidity vs. Volatility")
    st.scatter_chart(
        edited_df,
        x="Volatility (1–10)",
        y="Liquidity (1–10)",
        size="Expected Return (%)",
        color="Category"
    )
    st.divider()

    # === Filters ===
    st.subheader("Filters")
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

    # === Reports ===
    st.subheader("Generate Reports")

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

    def create_docx(df):
        document = Document()
        document.add_heading("HNW Investment Summary", 0)

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        document.add_heading("Portfolio Averages", level=1)
        for k, v in avg.items():
            document.add_paragraph(f"{k}: {v}")

        chart_file = "streamlit_chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90)
        plt.tight_layout()
        plt.savefig(chart_file)
        plt.close()

        document.add_heading("Expected Return Chart", level=1)
        document.add_picture(chart_file, width=DocxInches(6.5))
        docx_file = "HNW_Investment_Summary.docx"
        document.save(docx_file)
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
    st.error(f"⚠️ Error loading Excel file: {e}")
