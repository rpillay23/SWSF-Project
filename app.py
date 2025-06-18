import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import yfinance as yf

# === Page Config & Force Sidebar Always Open ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
    <style>
    /* Hide the sidebar collapse toggle */
    [data-testid="collapsedControl"] {
        display: none;
    }
    </style>
""", unsafe_allow_html=True)

# === Custom Styling ===
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Helvetica Neue', Helvetica, sans-serif;
    background-color: white;
    color: #003366;
}
h1, h2, h3, .stMarkdown, p {
    color: #003366; font-weight: bold;
}
.title-box {
    background-color: #B0C4DE; border: 5px solid #003366;
    border-radius: 10px; padding: 1em; text-align: center;
    margin-bottom: 20px;
}
.subtitle {
    font-size: 16px; color: #003366;
    text-align: center; margin-bottom: 30px;
}
.stButton > button {
    background-color: #003366; color: white;
    border-radius: 6px; padding: 0.5em 1em; border: none; font-weight: bold;
}
.stButton > button:hover {
    background-color: #0055a5;
}
.stMetric {
    background-color: #f0f8ff; padding: 1em;
    border-radius: 8px; color: #003366; font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# === App Title & Subtitle ===
st.markdown('<div class="title-box"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automated Software for Traditional and Alternate Investment Analysis Designed for Portfolio Management and Building a Modular Sustainable Wealth Strategy Framework (SWSF)</div>', unsafe_allow_html=True)

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
    idx = yf.Ticker(ticker)
    hist = idx.history(period="1mo")
    hist.reset_index(inplace=True)
    return hist

# === Sidebar: Permanent Real-Time Market Indices ===
st.sidebar.markdown("### Real-Time Market Indices")
for label, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
    data = get_index_data(ticker)
    if not data.empty:
        last = data.iloc[-1]; prev = data.iloc[-2]
        diff = last["Close"] - prev["Close"]
        pct = (diff / prev["Close"]) * 100
        st.sidebar.metric(label, f"${last['Close']:.2f}", f"{diff:+.2f} ({pct:+.2f}%)")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.plot(data["Date"], data["Close"], color="#003366")
        ax.tick_params(axis="x", labelsize=8, rotation=45)
        ax.tick_params(axis="y", labelsize=8)
        ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.sidebar.pyplot(fig)
    else:
        st.sidebar.warning(f"Failed to load {label} data.")
st.sidebar.divider()

try:
    # === Load Excel & Sanitize ===
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    # === Choose Investment Categories to Include ===
    st.subheader("Select Investment Types to Include")
    types = sorted(df["Category"].dropna().unique())
    selected = st.multiselect("Investment Types", types, default=types)
    df_filtered = df[df["Category"].isin(selected)]

    # === Editable Data Table ===
    st.subheader("Investment Data")
    edited_df = st.data_editor(df_filtered, use_container_width=True, num_rows="dynamic")
    st.divider()

    # === Portfolio Metrics ===
    st.subheader("Portfolio Averages & Totals")
    cols = st.columns(7)
    metrics = [
        ("Avg Return (%)", edited_df['Expected Return (%)'].mean(), "%"),
        ("Avg Risk (1–10)", edited_df['Risk Level (1-10)'].mean(), ""),
        ("Avg Cap Rate (%)", edited_df['Cap Rate (%)'].mean(), "%"),
        ("Avg Liquidity", edited_df['Liquidity (1–10)'].mean(), ""),
        ("Avg Volatility", edited_df['Volatility (1–10)'].mean(), ""),
        ("Avg Fees (%)", edited_df['Fees (%)'].mean(), "%"),
        ("Avg Min Investment", edited_df['Minimum Investment ($)'].mean(), "$")
    ]
    for col, (title, val, symbol) in zip(cols, metrics):
        formatted = f"{val:,.2f}{symbol}"
        if title.startswith("Avg Min"):
            formatted = f"${val:,.0f}"
        col.metric(title, formatted)
    st.divider()

    # === Four Compact Side-by-Side Charts ===
    st.subheader("Investment Visualizations (Compact)")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("**Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.bar(edited_df["Investment Name"], edited_df["Expected Return (%)"], color="teal")
        ax.set_xticklabels(edited_df["Investment Name"], rotation=45, ha="right", fontsize=7)
        ax.set_ylabel("Return (%)")
        ax.grid(axis="y", linestyle="--", alpha=0.4)
        plt.tight_layout(); st.pyplot(fig)

    with c2:
        st.markdown("**Liquidity vs Volatility**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        cats = edited_df["Category"].unique()
        cmap = plt.cm.get_cmap('tab10', len(cats))
        for idx, cat in enumerate(cats):
            sub = edited_df[edited_df["Category"] == cat]
            ax.scatter(sub["Volatility (1–10)"], sub["Liquidity (1–10)"],
                       s=sub["Expected Return (%)"] * 10, label=cat,
                       alpha=0.7, color=cmap(idx))
        ax.set_xlabel("Volatility"); ax.set_ylabel("Liquidity")
        ax.grid(True, linestyle="--", alpha=0.4)
        ax.legend(fontsize=7); plt.tight_layout(); st.pyplot(fig)

    with c3:
        st.markdown("**Risk Level Distribution**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.hist(edited_df["Risk Level (1-10)"].dropna(), bins=range(1, 12),
                color="#003366", edgecolor="black", alpha=0.7)
        ax.set_xlabel("Risk Level"); ax.set_ylabel("Count")
        ax.grid(axis="y", linestyle="--", alpha=0.4)
        plt.tight_layout(); st.pyplot(fig)

    with c4:
        st.markdown("**Fees vs Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(3, 2.5))
        ax.scatter(edited_df["Fees (%)"], edited_df["Expected Return (%)"],
                   alpha=0.7, color="#0055a5")
        ax.set_xlabel("Fees (%)"); ax.set_ylabel("Return (%)")
        ax.grid(True, linestyle="--", alpha=0.4)
        plt.tight_layout(); st.pyplot(fig)

    st.divider()

    # === Filters below charts ===
    st.subheader("Additional Filters")
    time_opts = ["All"] + sorted(df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_opts = ["All", "Yes", "No"]
    tfilter = st.selectbox("Select Time Horizon", time_opts)
    hfilter = st.selectbox("Inflation Hedge?", hedge_opts)
    min_inv = st.slider("Minimum Investment ($)",
                        int(df["Minimum Investment ($)"].min()),
                        int(df["Minimum Investment ($)"].max()),
                        int(df["Minimum Investment ($)"].min()))

    df_final = edited_df.copy()
    if tfilter != "All":
        df_final = df_final[df_final["Time Horizon (Short/Medium/Long)"] == tfilter]
    if hfilter != "All":
        df_final = df_final[df_final["Inflation Hedge (Yes/No)"] == hfilter]
    df_final = df_final[df_final["Minimum Investment ($)"] >= min_inv]

    st.subheader("Filtered Investment Table")
    st.dataframe(df_final, use_container_width=True)
    st.divider()

    # === Report Generation ===
    st.subheader("Generate Reports")
    def create_ppt(df):
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = "Comprehensive Investment Overview"
        s.placeholders[1].text = "Alternative & Traditional Investments"
        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = "Portfolio Averages"
        s.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])
        chart_file = "streamlit_chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90); plt.tight_layout(); plt.savefig(chart_file); plt.close()
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = "Expected Return Chart"
        s.shapes.add_picture(chart_file, Inches(1), Inches(1.5), width=Inches(8))
        fname = "HNW_Investment_Presentation.pptx"
        prs.save(fname); return fname

    def create_docx(df):
        doc = Document()
        doc.add_heading("HNW Investment Summary", 0)
        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        doc.add_heading("Portfolio Averages", level=1)
        for k, v in avg.items():
            doc.add_paragraph(f"{k}: {v}")
        chart_file = "streamlit_chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90); plt.tight_layout(); plt.savefig(chart_file); plt.close()
        doc.add_heading("Expected Return Chart", level=1)
        doc.add_picture(chart_file, width=DocxInches(6.5))
        fname = "HNW_Investment_Summary.docx"
        doc.save(fname); return fname

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Generate PowerPoint"):
            ppt = create_ppt(df_final)
            with open(ppt, "rb") as f:
                st.download_button("Download PPT", f, file_name=ppt)
    with col2:
        if st.button("Generate Word Report"):
            docx = create_docx(df_final)
            with open(docx, "rb") as f:
                st.download_button("Download DOCX", f, file_name=docx)

except Exception as e:
    st.error(f"⚠️ Error loading data: {e}")
