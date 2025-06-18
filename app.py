import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import yfinance as yf

# --- Page Config ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- Global Styles ---
st.markdown("""
    <style>
    body, html, [class*="css"] {
        font-family: 'Helvetica Neue', sans-serif;
        color: #111111;
        background-color: #ffffff;
    }
    h1 { font-size: 26px; font-weight: 700; }
    h2, h3 { font-size: 20px; font-weight: 600; }
    .stButton>button { background-color: #d32f2f; color: white; border: none; padding: 0.6em 1.2em; font-size: 14px; border-radius: 4px; }
    .stButton>button:hover { background-color: #b71c1c; }
    .stMetric { background-color: #f5f5f5; border-radius: 6px; padding: 0.6em; color: #111111; }
    .sidebar .stMetric .metric-label { color: #111111 !important; }
    .sidebar .stMetric .metric-value { color: #d32f2f !important; }
    .sidebar .stMetric .metric-delta { color: #111111 !important; }
    .sidebar .stMarkdownContainer { color: #111111; }
    .sidebar .css-1d391kg { background-color: #ffffff !important; } /* sidebar BG */
    </style>
""", unsafe_allow_html=True)

# --- Two-Column Layout ---
left_col, main_col = st.columns([1, 4], gap="medium")

# --- Left Panel: Market Indices (Professional) ---
with left_col:
    st.markdown("### 📈 Market Indices")
    @st.cache_data(ttl=600)
    def get_data(ticker):
        return yf.Ticker(ticker).history(period="1mo").reset_index()

    for name, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
        df_ind = get_data(ticker)
        if len(df_ind) > 1:
            cur, prev = df_ind.iloc[-1], df_ind.iloc[-2]
            diff = cur.Close - prev.Close
            pct = diff / prev.Close * 100
            st.metric(name, f"${cur.Close:.2f}", f"{diff:+.2f} ({pct:+.2f}%)")
            fig, ax = plt.subplots(figsize=(2.5, 0.8), dpi=100)
            ax.plot(df_ind["Date"], df_ind["Close"], color="#d32f2f", linewidth=1.5)
            ax.axis("off")
            st.pyplot(fig)
        else:
            st.error(f"{name} data unavailable")

# --- Main Panel: Core App ---
with main_col:
    st.markdown("<h1 style='text-align:center;'>Automated Investment Matrix</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;'>Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization and Financial Advisory.Designed for portfolio managers and finance professionals to generate, analyze, and optimize investment portfolios using real-time data.

</p>", unsafe_allow_html=True)

    # --- Load / Sanitize Data ---
    try:
        df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
        df = df.applymap(lambda x: x.replace("–", "-") if isinstance(x, str) else x)
    except Exception as e:
        st.error(f"⚠️ Data load error: {e}")
        st.stop()

    # --- Category Selector & Editable Table ---
    st.markdown("### Select Investment Types")
    cats = sorted(df["Category"].dropna().unique())
    sel = st.multiselect("Categories", cats, default=cats)
    df_sel = df[df["Category"].isin(sel)]
    st.markdown("### Investment Data")
    edited = st.data_editor(df_sel, num_rows="dynamic", use_container_width=True)

    # --- Metrics Row ---
    st.markdown("---")
    labels = {
        "Expected Return (%)": "Return (%)",
        "Risk Level (1-10)": "Risk (1–10)",
        "Cap Rate (%)": "Cap Rate (%)",
        "Liquidity (1–10)": "Liquidity (1–10)",
        "Volatility (1–10)": "Volatility (1–10)",
        "Fees (%)": "Fees (%)",
        "Minimum Investment ($)": "Min Invest ($)"
    }
    cols = st.columns(len(labels))
    for col, (col_name, label) in zip(cols, labels.items()):
        val = edited[col_name].mean() if col_name in edited else None
        if pd.notna(val):
            disp = f"${val:,.0f}" if "Investment" in label else f"{val:.2f}"
        else:
            disp = "N/A"
        col.metric(label, disp)

    # --- Visual Insights: 4 Plots ---
    st.markdown("### Visual Insights")
    c1, c2, c3, c4 = st.columns(4, gap="small")

    with c1:
        st.markdown("**Expected Return (%)**")
        fig, ax = plt.subplots(figsize=(2.2, 1.6), dpi=100)
        ax.bar(edited["Investment Name"], edited["Expected Return (%)"], color="#d32f2f")
        ax.set_xticklabels(edited["Investment Name"], rotation=45, fontsize=6)
        ax.tick_params(left=False, labelleft=False)
        plt.tight_layout(); st.pyplot(fig)

    with c2:
        st.markdown("**Liquidity vs Volatility**")
        fig, ax = plt.subplots(figsize=(2.2, 1.6), dpi=100)
        ax.scatter(
            edited["Volatility (1–10)"], edited["Liquidity (1–10)"],
            s=edited["Expected Return (%)"] * 5, c="red", alpha=0.7
        )
        ax.set_xlabel("Volatility", fontsize=6); ax.set_ylabel("Liquidity", fontsize=6)
        ax.tick_params(axis="both", which="major", labelsize=5)
        plt.tight_layout(); st.pyplot(fig)

    with c3:
        st.markdown("**Risk Distribution**")
        fig, ax = plt.subplots(figsize=(2.2, 1.6), dpi=100)
        ax.hist(edited["Risk Level (1-10)"].dropna(), bins=10, color="#222222", alpha=0.8)
        ax.set_xlabel("Risk Level", fontsize=6); ax.set_ylabel("Count", fontsize=6)
        ax.tick_params(labelsize=5); plt.tight_layout(); st.pyplot(fig)

    with c4:
        st.markdown("**Fees vs Expected Return**")
        fig, ax = plt.subplots(figsize=(2.2, 1.6), dpi=100)
        ax.scatter(
            edited["Fees (%)"], edited["Expected Return (%)"],
            c="red", alpha=0.7
        )
        ax.set_xlabel("Fees (%)", fontsize=6); ax.set_ylabel("Return (%)", fontsize=6)
        ax.tick_params(labelsize=5); plt.tight_layout(); st.pyplot(fig)

    # --- Data Filters ---
    st.markdown("---")
    time_opts = ["All"] + sorted(edited["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_opts = ["All", "Yes", "No"]
    time_f = st.selectbox("Time Horizon", time_opts)
    hedge_f = st.selectbox("Inflation Hedge?", hedge_opts)
    min_f = st.slider(
        "Minimum Investment ($)",
        int(edited["Minimum Investment ($)"].min()),
        int(edited["Minimum Investment ($)"].max()),
        int(edited["Minimum Investment ($)"].min())
    )
    filtered = edited.copy()
    if time_f != "All":
        filtered = filtered[filtered["Time Horizon (Short/Medium/Long)"] == time_f]
    if hedge_f != "All":
        filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == hedge_f]
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_f]

    st.markdown("### Filtered Results")
    st.dataframe(filtered, use_container_width=True)

    # --- Export Reports: PPTX & DOCX ---
    st.markdown("---")
    st.markdown("### Export Reports")
    def build_ppt(df_input):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Investment Overview"
        slide.placeholders[1].text = "Automated Investment Matrix"
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Portfolio Averages"
        avg = df_input.select_dtypes(include="number").mean(numeric_only=True).round(2)
        slide.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])
        chart_path = "ppt_chart.png"
        plt.figure(figsize=(6, 2)); plt.bar(df_input["Investment Name"], df_input["Expected Return (%)"], color="#d32f2f")
        plt.xticks(rotation=90, fontsize=6); plt.tight_layout(); plt.savefig(chart_path); plt.close()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Return Chart"
        slide.shapes.add_picture(chart_path, Inches(1), Inches(1.5), width=Inches(6))
        path = "Investment_Summary.pptx"; prs.save(path)
        return path

    def build_doc(df_input):
        doc = Document()
        doc.add_heading("Investment Summary", 0)
        avg = df_input.select_dtypes(include="number").mean(numeric_only=True).round(2)
        for k, v in avg.items():
            doc.add_paragraph(f"{k}: {v}")
        chart_path = "doc_chart.png"
        plt.figure(figsize=(6, 2)); plt.bar(df_input["Investment Name"], df_input["Expected Return (%)"], color="#d32f2f")
        plt.xticks(rotation=90, fontsize=6); plt.tight_layout(); plt.savefig(chart_path); plt.close()
        doc.add_picture(chart_path, width=DocxInches(6))
        path = "Investment_Summary.docx"; doc.save(path)
        return path

    cP, cD = st.columns(2)
    with cP:
        if st.button("Generate PowerPoint"):
            pptx_path = build_ppt(filtered)
            with open(pptx_path, "rb") as fp:
                st.download_button("Download PPTX", fp, file_name=pptx_path)
    with cD:
        if st.button("Generate Word Document"):
            docx_path = build_doc(filtered)
            with open(docx_path, "rb") as fd:
                st.download_button("Download DOCX", fd, file)
