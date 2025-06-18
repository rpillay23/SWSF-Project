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

# === Sidebar: Real-Time Market Data (PERMANENT) ===
st.sidebar.header("Real-Time Market Data")

for label, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
    df_market = get_index_data(ticker)
    if not df_market.empty:
        last = df_market.iloc[-1]
        prev = df_market.iloc[-2]
        diff = last["Close"] - prev["Close"]
        pct = (diff / prev["Close"]) * 100
        st.sidebar.metric(label, f"${last['Close']:.2f}", f"{diff:+.2f} ({pct:+.2f}%)")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.plot(df_market["Date"], df_market["Close"], color="#003366")
        ax.tick_params(axis="x", labelsize=8, rotation=45)
        ax.tick_params(axis="y", labelsize=8)
        ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.sidebar.pyplot(fig)
    else:
        st.sidebar.warning(f"Failed to load {label} data.")
st.sidebar.divider()

try:
    # === Load Excel Data ===
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # === Select Investment Types to Include ===
    st.subheader("Select Investment Types to Include")
    investment_types = sorted(df["Category"].dropna().unique())
    selected_types = st.multiselect("Choose categories to include", investment_types, default=investment_types)
    filtered_investments = df[df["Category"].isin(selected_types)]

    # === Editable Table ===
    st.subheader("Investment Data")
    edited_df = st.data_editor(filtered_investments, use_container_width=True, num_rows="dynamic")

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

    # === Compact 4 Graphs Side-by-Side ===
    st.subheader("Investment Visualizations")
    g1, g2, g3, g4 = st.columns(4)

    # Expected Return Bar Chart
    with g1:
        st.markdown("**Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.bar(edited_df["Investment Name"], edited_df["Expected Return (%)"], color="teal")
        ax.set_xticklabels(edited_df["Investment Name"], rotation=45, ha="right", fontsize=7)
        ax.set_ylabel("Return (%)")
        ax.grid(axis="y", linestyle="--", alpha=0.4)
        plt.tight_layout()
        st.pyplot(fig)

    # Liquidity vs Volatility Scatter Plot
    with g2:
        st.markdown("**Liquidity vs. Volatility**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        categories = edited_df["Category"].astype(str).unique()
        colors = plt.cm.get_cmap('tab10', len(categories))
        for i, cat in enumerate(categories):
            subset = edited_df[edited_df["Category"] == cat]
            ax.scatter(subset["Volatility (1–10)"], subset["Liquidity (1–10)"],
                       s=subset["Expected Return (%)"]*10, label=cat,
                       alpha=0.7, color=colors(i))
        ax.set_xlabel("Volatility (1–10)")
        ax.set_ylabel("Liquidity (1–10)")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend(fontsize=7, loc="upper right")
        plt.tight_layout()
        st.pyplot(fig)

    # Risk Level Distribution Histogram
    with g3:
        st.markdown("**Risk Level Distribution**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.hist(edited_df["Risk Level (1-10)"].dropna(), bins=range(1,12), color="#003366", alpha=0.7, edgecolor="black")
        ax.set_xlabel("Risk Level (1-10)")
        ax.set_ylabel("Count")
        ax.grid(axis="y", linestyle="--", alpha=0.4)
        plt.tight_layout()
        st.pyplot(fig)

    # Fees vs Expected Return Scatter Plot
    with g4:
        st.markdown("**Fees (%) vs Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.scatter(edited_df["Fees (%)"], edited_df["Expected Return (%)"], alpha=0.7, color="#0055a5")
        ax.set_xlabel("Fees (%)")
        ax.set_ylabel("Expected Return (%)")
        ax.grid(True, linestyle="--", alpha=0.4)
        plt.tight_layout()
        st.pyplot(fig)

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
