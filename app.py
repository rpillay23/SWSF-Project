import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import yfinance as yf

# === Page Config & Styling ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Helvetica Neue', Helvetica, sans-serif;
    background-color: white; color: #003366;
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
    border-radius: 6px; padding: 0.5em 1em;
    border: none; font-weight: bold;
}
.stButton > button:hover {background-color: #0055a5;}
.stMetric {
    background-color: #f0f8ff; padding: 1em;
    border-radius: 8px; color: #003366; font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title-box"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automated Software for Traditional and Alternate Investment Analysis Designed for Portfolio Management and Building a Modular Sustainable Wealth Strategy Framework (SWSF)</div>', unsafe_allow_html=True)

def sanitize_string(val):
    if isinstance(val, str):
        return (val.replace("–", "-")
                   .replace("’", "'")
                   .replace("“", '"')
                   .replace("”", '"')
                   .replace("•", "-")
                   .replace("©", "(c)"))
    return val

@st.cache_data(ttl=600)
def get_index_data(ticker):
    tk = yf.Ticker(ticker)
    hist = tk.history(period="1mo")
    hist.reset_index(inplace=True)
    return hist

# === Sidebar: Real-Time Market Data ===
st.sidebar.header("Real‑Time Market Data")
for label, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
    df_market = get_index_data(ticker)
    if not df_market.empty:
        last = df_market.iloc[-1]; prev = df_market.iloc[-2]
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
st.divider()

try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)

    st.subheader("Investment Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    st.divider()

    st.subheader("Select Investment Types for Analysis")
    types = sorted(edited_df["Category"].dropna().unique())
    selected = st.multiselect("Investment Types", types, default=types)
    filtered_types_df = edited_df[edited_df["Category"].isin(selected)]

    st.subheader("Portfolio Averages & Totals")
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Avg Return (%)", f"{filtered_types_df['Expected Return (%)'].mean():.2f}%")
    c2.metric("Avg Risk (1–10)", f"{filtered_types_df['Risk Level (1-10)'].mean():.2f}")
    c3.metric("Avg Cap Rate (%)", f"{filtered_types_df['Cap Rate (%)'].mean():.2f}%")
    c4.metric("Avg Liquidity", f"{filtered_types_df['Liquidity (1–10)'].mean():.2f}")
    c5.metric("Avg Volatility", f"{filtered_types_df['Volatility (1–10)'].mean():.2f}")
    c6.metric("Avg Fees (%)", f"{filtered_types_df['Fees (%)'].mean():.2f}%")
    c7.metric("Avg Min Investment", f"${filtered_types_df['Minimum Investment ($)'].mean():,.0f}")
    st.divider()

    st.subheader("Investment Visualizations (Compact)")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("**Expected Return**")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.bar(filtered_types_df["Investment Name"], filtered_types_df["Expected Return (%)"], color="teal")
        ax.set_xticklabels(filtered_types_df["Investment Name"], rotation=45, ha="right", fontsize=6)
        ax.set_ylabel("%", fontsize=8)
        ax.tick_params(labelsize=6)
        ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout()
        st.pyplot(fig)

    with col2:
        st.markdown("**Liquidity vs Volatility**")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.scatter(
            filtered_types_df["Volatility (1–10)"],
            filtered_types_df["Liquidity (1–10)"],
            s=filtered_types_df["Expected Return (%)"] * 8,
            c=pd.factorize(filtered_types_df["Category"])[0],
            cmap="tab10", alpha=0.7
        )
        ax.set_xlabel("Volatility", fontsize=8); ax.set_ylabel("Liquidity", fontsize=8)
        ax.tick_params(labelsize=6); ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout(); st.pyplot(fig)

    with col3:
        st.markdown("**Risk vs Return**")
        fig, ax = plt.subplots(figsize=(4, 2.5))
        ax.scatter(
            filtered_types_df["Risk Level (1-10)"],
            filtered_types_df["Expected Return (%)"],
            c=pd.factorize(filtered_types_df["Category"])[0],
            cmap="tab10", alpha=0.7
        )
        ax.set_xlabel("Risk", fontsize=8); ax.set_ylabel("Return %", fontsize=8)
        ax.tick_params(labelsize=6); ax.grid(True, linestyle="--", alpha=0.3)
        plt.tight_layout(); st.pyplot(fig)

    with col4:
    st.markdown("**Fees vs Expected Return**")
    fig, ax = plt.subplots(figsize=(4, 2.5))
    ax.scatter(
        filtered_types_df["Fees (%)"],
        filtered_types_df["Expected Return (%)"],
        c=pd.factorize(filtered_types_df["Category"])[0],
        cmap="tab10", alpha=0.7
    )
    ax.set_xlabel("Fees (%)", fontsize=8)
    ax.set_ylabel("Return (%)", fontsize=8)
    ax.tick_params(labelsize=6)
    ax.grid(True, linestyle="--", alpha=0.3)
    plt.tight_layout()
    st.pyplot(fig)

    st.divider()

    # === Filtering Section ===
    st.subheader("Additional Filters")
    time_opts = ["All"] + sorted(edited_df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_opts = ["All", "Yes", "No"]
    tfilter = st.selectbox("Time Horizon", time_opts)
    hfilter = st.selectbox("Inflation Hedge?", hedge_opts)
    min_filter = st.slider("Minimum Investment ($)", 
                            int(edited_df["Minimum Investment ($)"].min()),
                            int(edited_df["Minimum Investment ($)"].max()),
                            int(edited_df["Minimum Investment ($)"].min()))

    df_final = filtered_types_df.copy()
    if tfilter != "All":
        df_final = df_final[df_final["Time Horizon (Short/Medium/Long)"] == tfilter]
    if hfilter != "All":
        df_final = df_final[df_final["Inflation Hedge (Yes/No)"] == hfilter]
    df_final = df_final[df_final["Minimum Investment ($)"] >= min_filter]

    st.subheader("Filtered Investment Table")
    st.dataframe(df_final, use_container_width=True)
    st.divider()

    # === Reporting ===
    st.subheader("Generate Reports")
    def create_ppt(dataframe):
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[0])
        s.shapes.title.text = "Comprehensive Investment Overview"
        s.placeholders[1].text = "Alternative & Traditional Investments"
        avg = dataframe.select_dtypes(include="number").mean(numeric_only=True).round(2)
        s = prs.slides.add_slide(prs.slide_layouts[1])
        s.shapes.title.text = "Portfolio Averages"
        s.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])
        chart_file = "chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(dataframe["Investment Name"], dataframe["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90); plt.tight_layout(); plt.savefig(chart_file); plt.close()
        s = prs.slides.add_slide(prs.slide_layouts[5])
        s.shapes.title.text = "Expected Return Chart"
        s.shapes.add_picture(chart_file, Inches(1), Inches(1.5), width=Inches(8))
        fname = "Investment_Presentation.pptx"
        prs.save(fname)
        return fname

    def create_doc(dataframe):
        doc = Document()
        doc.add_heading("HNW Investment Summary", 0)
        avg = dataframe.select_dtypes(include="number").mean(numeric_only=True).round(2)
        doc.add_heading("Portfolio Averages", level=1)
        for k, v in avg.items():
            doc.add_paragraph(f"{k}: {v}")
        chart_file = "chart.png"
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.bar(dataframe["Investment Name"], dataframe["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90); plt.tight_layout(); plt.savefig(chart_file); plt.close()
        doc.add_heading("Expected Return Chart", level=1)
        doc.add_picture(chart_file, width=DocxInches(6.5))
        fname = "Investment_Summary.docx"
        doc.save(fname)
        return fname

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Generate PowerPoint"):
            ppt = create_ppt(df_final)
            with open(ppt, "rb") as fp:
                st.download_button("Download PPT", fp, file_name=ppt)
    with c2:
        if st.button("Generate Word Report"):
            doc = create_doc(df_final)
            with open(doc, "rb") as fd:
                st.download_button("Download DOCX", fd, file_name=doc)

except Exception as err:
    st.error(f"⚠️ Error loading data: {err}")
