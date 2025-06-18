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

st.markdown('<div class="title-box"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automated Software for Traditional and Alternate Investment Analysis Designed for Portfolio Management and Building a Modular Sustainable Wealth Strategy Framework (SWSF)</div>', unsafe_allow_html=True)

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

# === Sidebar Market Data ===
st.sidebar.header("Real-Time Market Data")
for label, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
    data = get_index_data(ticker)
    if not data.empty:
        latest = data.iloc[-1]
        prev = data.iloc[-2]
        diff = latest["Close"] - prev["Close"]
        pct = (diff / prev["Close"]) * 100
        st.sidebar.metric(label, f"${latest['Close']:.2f}", f"{diff:+.2f} ({pct:+.2f}%)")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.plot(data["Date"], data["Close"], color="#003366")
        ax.tick_params(axis='x', rotation=45, labelsize=8)
        ax.tick_params(axis='y', labelsize=8)
        ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.sidebar.pyplot(fig)
    else:
        st.sidebar.warning(f"Could not load {label} data.")

st.divider()

try:
    # === Load and Clean Excel Data ===
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # === Editable Table ===
    st.subheader("Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    st.divider()

    # === Filter by Investment Type ===
    st.subheader("Select Investment Types for Portfolio Analysis")
    all_types = sorted(edited_df["Category"].dropna().unique())
    selected_types = st.multiselect("Investment Types", all_types, default=all_types)
    filtered_by_type_df = edited_df[edited_df["Category"].isin(selected_types)]

    # === Portfolio Metrics ===
    st.subheader("Portfolio Averages and Totals")
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    col1.metric("Avg Return (%)", f"{filtered_by_type_df['Expected Return (%)'].mean():.2f}%")
    col2.metric("Avg Risk (1–10)", f"{filtered_by_type_df['Risk Level (1-10)'].mean():.2f}")
    col3.metric("Avg Cap Rate (%)", f"{filtered_by_type_df['Cap Rate (%)'].mean():.2f}%")
    col4.metric("Avg Liquidity", f"{filtered_by_type_df['Liquidity (1–10)'].mean():.2f}")
    col5.metric("Avg Volatility", f"{filtered_by_type_df['Volatility (1–10)'].mean():.2f}")
    col6.metric("Avg Fees (%)", f"{filtered_by_type_df['Fees (%)'].mean():.2f}%")
    col7.metric("Avg Min Investment", f"${filtered_by_type_df['Minimum Investment ($)'].mean():,.0f}")
    st.divider()

    # === Visual Charts ===
    st.subheader("Investment Visualizations")

    # Row 1
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Expected Return by Investment**")
        fig1, ax1 = plt.subplots(figsize=(5, 3))
        ax1.bar(filtered_by_type_df["Investment Name"], filtered_by_type_df["Expected Return (%)"], color="teal")
        ax1.set_xticklabels(filtered_by_type_df["Investment Name"], rotation=45, ha="right", fontsize=8)
        ax1.set_ylabel("Expected Return (%)")
        ax1.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig1)

    with col2:
        st.markdown("**Liquidity vs. Volatility**")
        fig2, ax2 = plt.subplots(figsize=(5, 3))
        ax2.scatter(
            filtered_by_type_df["Volatility (1–10)"],
            filtered_by_type_df["Liquidity (1–10)"],
            s=filtered_by_type_df["Expected Return (%)"] * 10,
            c=pd.factorize(filtered_by_type_df["Category"])[0],
            cmap="tab10",
            alpha=0.7
        )
        ax2.set_xlabel("Volatility (1–10)")
        ax2.set_ylabel("Liquidity (1–10)")
        ax2.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig2)

    # Row 2
    col3, col4 = st.columns(2)
    with col3:
        st.markdown("**Risk vs Expected Return**")
        fig3, ax3 = plt.subplots(figsize=(5, 3))
        ax3.scatter(
            filtered_by_type_df["Risk Level (1-10)"],
            filtered_by_type_df["Expected Return (%)"],
            c=pd.factorize(filtered_by_type_df["Category"])[0],
            cmap="tab10",
            alpha=0.7
        )
        ax3.set_xlabel("Risk Level (1–10)")
        ax3.set_ylabel("Expected Return (%)")
        ax3.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig3)

    with col4:
        st.markdown("**Minimum Investment Distribution**")
        fig4, ax4 = plt.subplots(figsize=(5, 3))
        ax4.hist(filtered_by_type_df["Minimum Investment ($)"], bins=10, color="#003366", edgecolor="white")
        ax4.set_xlabel("Minimum Investment ($)")
        ax4.set_ylabel("Number of Investments")
        ax4.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig4)

    st.divider()

    # === Filters ===
    st.subheader("Additional Filters")
    time_opts = ["All"] + sorted(edited_df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_opts = ["All", "Yes", "No"]

    time_filter = st.selectbox("Time Horizon", time_opts)
    hedge_filter = st.selectbox("Inflation Hedge?", hedge_opts)
    min_inv_filter = st.slider("Minimum Investment ($)",
                               int(edited_df["Minimum Investment ($)"].min()),
                               int(edited_df["Minimum Investment ($)"].max()),
                               int(edited_df["Minimum Investment ($)"].min()))

    filtered_df = filtered_by_type_df.copy()
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
