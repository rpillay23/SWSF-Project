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

# === Professional Black-Red-White Theme + Header ===
st.markdown("""
    <style>
    /* Reset body & fonts */
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', Helvetica, sans-serif;
        background-color: white;
        color: #111111;
    }

    /* Header box full width */
    .app-header {
        background-color: #111111;
        padding: 1rem 2rem;
        border-radius: 0;
        margin: 0;
        box-shadow: none;
        width: 100vw;
        position: relative;
        left: calc(-50vw + 50%);
        top: 0;
        user-select: none;
        -webkit-user-select: none;
        -moz-user-select: none;
        -ms-user-select: none;
    }
    .app-header h1 {
        color: white;
        font-size: 24px;
        margin: 0;
        font-weight: 700;
        line-height: 1.2;
    }
    .app-header p {
        color: #f44336;
        font-size: 13px;
        margin: 0.25rem 0 0 0;
        font-weight: 500;
    }

    /* Sidebar style */
    .sidebar .sidebar-content {
        background-color: #111111;
        color: white;
        padding: 1rem;
    }
    .sidebar .sidebar-content .stMetric {
        background-color: #222222;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        margin-bottom: 0.75rem;
        color: white;
        font-weight: 600;
    }
    /* Sidebar metric labels */
    .sidebar .sidebar-content .stMetric > div > div:nth-child(1) {
        color: #f44336;
        font-weight: 700;
    }
    /* Sidebar warnings */
    .sidebar .sidebar-content .stWarning {
        background-color: #330000;
        color: #ff6666;
        border-radius: 6px;
        padding: 0.5rem;
        margin-bottom: 0.5rem;
    }

    /* Buttons */
    .stButton > button {
        background-color: #111111;
        color: #f44336;
        border-radius: 6px;
        padding: 0.5em 1em;
        border: 2px solid #f44336;
        font-weight: 700;
        transition: background-color 0.3s ease, color 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #f44336;
        color: #111111;
        border: 2px solid #f44336;
    }

    /* Metrics box */
    .stMetric {
        background-color: #f9f9f9;
        padding: 1em;
        border-radius: 8px;
        color: #111111;
        font-weight: 700;
    }
    </style>

    <div class="app-header">
        <h1>Automated Investment Matrix</h1>
        <p>
            Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization and Financial Advisory.
            Designed for portfolio managers and finance professionals to generate, analyze, and optimize investment portfolios using real-time data.
        </p>
    </div>
""", unsafe_allow_html=True)


# === Helper Functions ===
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


# === Sidebar: Permanent Real-Time Market Data ===
st.sidebar.markdown("## Real-Time Market Indices")

def sidebar_metric(name, data):
    latest = data.iloc[-1]
    prev = data.iloc[-2]
    latest_close = latest['Close']
    prev_close = prev['Close']
    change = latest_close - prev_close
    pct_change = (change / prev_close) * 100

    st.sidebar.markdown(f"### {name}")
    st.sidebar.metric(label="", value=f"${latest_close:.2f}", delta=f"{change:+.2f} ({pct_change:+.2f}%)")

    fig, ax = plt.subplots(figsize=(3, 1.5))
    ax.plot(data['Date'], data['Close'], color='#f44336')
    ax.set_xticks([])
    ax.set_yticks([])
    ax.grid(False)
    plt.tight_layout()
    st.sidebar.pyplot(fig)

# Fetch & display indices data
try:
    sp500_data = get_index_data("^GSPC")
    if not sp500_data.empty:
        sidebar_metric("S&P 500", sp500_data)
    else:
        st.sidebar.warning("Failed to fetch S&P 500 data.")

    nasdaq_data = get_index_data("^IXIC")
    if not nasdaq_data.empty:
        sidebar_metric("Nasdaq", nasdaq_data)
    else:
        st.sidebar.warning("Failed to fetch Nasdaq data.")

    dowjones_data = get_index_data("^DJI")
    if not dowjones_data.empty:
        sidebar_metric("Dow Jones", dowjones_data)
    else:
        st.sidebar.warning("Failed to fetch Dow Jones data.")
except Exception as e:
    st.sidebar.error(f"Error loading market data: {e}")

st.sidebar.markdown("---")


# === Main App Content ===
try:
    # Load Excel Data
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # Editable Table
    st.subheader("Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    st.divider()

    # === Select Investment Types Filter ===
    st.subheader("Select Investment Types to Include")
    investment_types = sorted(edited_df["Category"].dropna().unique())
    selected_types = st.multiselect("Investment Types", investment_types, default=investment_types)
    filtered_df = edited_df[edited_df["Category"].isin(selected_types)]

    st.divider()

    # Metrics (on filtered data)
    st.subheader("Portfolio Averages and Totals")
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("Avg Return (%)", f"{filtered_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk (1–10)", f"{filtered_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate (%)", f"{filtered_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{filtered_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{filtered_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees (%)", f"{filtered_df['Fees (%)'].mean():.2f}%")
    col7.metric("Avg Min Investment", f"${filtered_df['Minimum Investment ($)'].mean():,.0f}")

    st.divider()

    # === Charts Section: 4 compact charts side by side ===
    st.subheader("Visual Insights")

    chart1, chart2, chart3, chart4 = st.columns(4)

    # 1) Expected Return by Investment (Bar chart)
    with chart1:
        st.markdown("**Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2))
        ax.bar(filtered_df["Investment Name"], filtered_df["Expected Return (%)"], color="#f44336")
        ax.tick_params(axis='x', rotation=45, labelsize=7)
        ax.tick_params(axis='y', labelsize=8)
        plt.tight_layout()
        st.pyplot(fig)

    # 2) Liquidity vs. Volatility (Scatterplot, red points)
    with chart2:
        st.markdown("**Liquidity vs Volatility**")
        fig, ax = plt.subplots(figsize=(3, 2))
        ax.scatter(filtered_df["Volatility (1–10)"], filtered_df["Liquidity (1–10)"], s=30,
                   c='red', alpha=0.7, edgecolors='black')
        ax.set_xlabel("Volatility (1–10)", fontsize=8)
        ax.set_ylabel("Liquidity (1–10)", fontsize=8)
        ax.tick_params(axis='x', labelsize=7)
        ax.tick_params(axis='y', labelsize=7)
        plt.tight_layout()
        st.pyplot(fig)

    # 3) Fees (%) vs Expected Return (%)
    with chart3:
        st.markdown("**Fees (%) vs Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2))
        ax.scatter(filtered_df["Fees (%)"], filtered_df["Expected Return (%)"], s=30,
                   c='red', alpha=0.7, edgecolors='black')
        ax.set_xlabel("Fees (%)", fontsize=8)
        ax.set_ylabel("Expected Return (%)", fontsize=8)
        ax.tick_params(axis='x', labelsize=7)
        ax.tick_params(axis='y', labelsize=7)
        plt.tight_layout()
        st.pyplot(fig)

    # 4) Risk Level Distribution (Histogram)
    with chart4:
        st.markdown("**Risk Level Distribution**")
        fig, ax = plt.subplots(figsize=(3, 2))
        ax.hist(filtered_df["Risk Level (1-10)"], bins=10, color="#f44336", alpha=0.8)
        ax.set_xlabel("Risk Level (1-10)", fontsize=8)
        ax.set_ylabel("Count", fontsize=8)
        ax.tick_params(axis='x', labelsize=7)
        ax.tick_params(axis='y', labelsize=7)
        plt.tight_layout()
        st.pyplot(fig)

    st.divider()

    # === Export Options ===
    st.subheader("Export Reports")
    export_format = st.radio("Select export format", ["PowerPoint", "Word Document"])

    if st.button("Generate Report"):
        if export_format == "PowerPoint":
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title_shape = slide.shapes.title
            title_shape.text = "Investment Portfolio Report"

            # Add a simple text box with summary
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(5)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            tf = textbox.text_frame
            tf.text = f"Average Expected Return: {filtered_df['Expected Return (%)'].mean():.2f}%\n"
            tf.add_paragraph().text = f"Average Risk Level: {filtered_df['Risk Level (1-10)'].mean():.2f}\n"
            tf.add_paragraph().text = f"Average Fees: {filtered_df['Fees (%)'].mean():.2f}%\n"

            prs.save("portfolio_report.pptx")
            with open("portfolio_report.pptx", "rb") as f:
                st.download_button("Download PowerPoint Report", f, "portfolio_report.pptx")

        elif export_format == "Word Document":
            doc = Document()
            doc.add_heading("Investment Portfolio Report", 0)
            doc.add_paragraph(f"Average Expected Return: {filtered_df['Expected Return (%)'].mean():.2f}%")
            doc.add_paragraph(f"Average Risk Level: {filtered_df['Risk Level (1-10)'].mean():.2f}")
            doc.add_paragraph(f"Average Fees: {filtered_df['Fees (%)'].mean():.2f}%")
            doc.save("portfolio_report.docx")
            with open("portfolio_report.docx", "rb") as f:
                st.download_button("Download Word Report", f, "portfolio_report.docx")

except FileNotFoundError:
    st.error("Error: 'Comprehensive_Investment_Matrix.xlsx' not found. Please upload it to the app directory.")
except Exception as e:
    st.error(f"Unexpected error: {e}")
