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

# --- Styling ---
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Helvetica Neue', Helvetica, sans-serif;
    background-color: white; color: #003366;
}
h1 { font-size: 24px; }
h2, h3 { font-size: 18px; }
.stButton > button { font-size: 12px; padding: 0.3em 0.6em; }
.stMetric { font-size: 12px; padding: 0.6em; }
</style>
""", unsafe_allow_html=True)

# --- Layout Columns ---
left_col, main_col = st.columns([1, 4])

# --- Left Column: Market Indices ---
with left_col:
    st.markdown("### 📈 Market Indices")

    @st.cache_data(ttl=600)
    def get_index_data(ticker):
        return yf.Ticker(ticker).history(period="1mo").reset_index()

    for label, ticker in [("S&P 500", "^GSPC"), ("Nasdaq", "^IXIC"), ("Dow Jones", "^DJI")]:
        df_index = get_index_data(ticker)
        if not df_index.empty and len(df_index) > 1:
            cur, prev = df_index.iloc[-1], df_index.iloc[-2]
            diff = cur.Close - prev.Close
            pct = (diff / prev.Close) * 100
            st.metric(label, f"${cur.Close:.2f}", f"{diff:+.2f} ({pct:+.2f}%)")

            # Line chart
            fig, ax = plt.subplots(figsize=(2.5, 1.0))
            ax.plot(df_index["Date"], df_index["Close"], color="#003366", linewidth=1)
            ax.set_xticks([]); ax.set_yticks([]); ax.axis("off")
            st.pyplot(fig, bbox_inches="tight")
        else:
            st.warning(f"{label} data unavailable")

# --- Main Content ---
with main_col:
    st.markdown('<div style="text-align:center;"><h1>Automated Investment Matrix</h1></div>', unsafe_allow_html=True)
    st.markdown("Modular Analysis Tool for Building Smart Portfolios with Real-Time Data Designed for Portfolio Management for Finance Professionals")

    # --- Load Data ---
    try:
        df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
        df = df.applymap(lambda x: x.replace("–", "-") if isinstance(x, str) else x)
    except Exception as e:
        st.error(f"Error loading data: {e}")
        st.stop()

    # --- Select Investment Categories ---
    st.markdown("### Select Investment Types to Include")
    unique_types = sorted(df["Category"].dropna().unique())
    selected_types = st.multiselect("Categories", unique_types, default=unique_types)
    filtered_df = df[df["Category"].isin(selected_types)]

    # --- Editable Table ---
    st.markdown("### Investment Table")
    edited_df = st.data_editor(filtered_df, num_rows="dynamic", use_container_width=True)

    # --- Metrics ---
    st.markdown("### Portfolio Averages")
    metrics = {
        "Expected Return (%)": "Return (%)",
        "Risk Level (1-10)": "Risk",
        "Cap Rate (%)": "Cap Rate",
        "Liquidity (1–10)": "Liquidity",
        "Volatility (1–10)": "Volatility",
        "Fees (%)": "Fees",
        "Minimum Investment ($)": "Min Investment"
    }
    cols = st.columns(len(metrics))
    for col, (key, label) in zip(cols, metrics.items()):
        try:
            avg_val = edited_df[key].mean()
            fmt = f"${avg_val:,.0f}" if "Investment" in label else f"{avg_val:.2f}"
            col.metric(label, fmt)
        except:
            col.metric(label, "N/A")

    # --- Visual Insights ---
    st.markdown("### Visual Insights")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("**Expected Return (%)**")
        plt.figure(figsize=(2.5, 1.8))
        plt.bar(edited_df["Investment Name"], edited_df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=45, fontsize=6); plt.yticks(fontsize=6); plt.tight_layout()
        st.pyplot(plt)

    with c2:
        st.markdown("**Liquidity vs Volatility**")
        plt.figure(figsize=(2.5, 1.8))
        plt.scatter(
            edited_df["Volatility (1–10)"],
            edited_df["Liquidity (1–10)"],
            s=edited_df["Expected Return (%)"] * 8,
            alpha=0.7,
            color="red"
        )
        plt.xlabel("Volatility", fontsize=7); plt.ylabel("Liquidity", fontsize=7)
        plt.xticks(fontsize=6); plt.yticks(fontsize=6); plt.tight_layout()
        st.pyplot(plt)

    with c3:
        st.markdown("**Risk Distribution**")
        plt.figure(figsize=(2.5, 1.8))
        plt.hist(edited_df["Risk Level (1-10)"].dropna(), bins=10, color="#003366", alpha=0.7)
        plt.xlabel("Risk", fontsize=7); plt.ylabel("Count", fontsize=7)
        plt.xticks(fontsize=6); plt.yticks(fontsize=6); plt.tight_layout()
        st.pyplot(plt)

    with c4:
        st.markdown("**Fees vs Expected Return**")
        plt.figure(figsize=(2.5, 1.8))
        plt.scatter(
            edited_df["Fees (%)"],
            edited_df["Expected Return (%)"],
            color="red",
            alpha=0.7
        )
        plt.xlabel("Fees (%)", fontsize=7); plt.ylabel("Return (%)", fontsize=7)
        plt.xticks(fontsize=6); plt.yticks(fontsize=6); plt.tight_layout()
        st.pyplot(plt)

    # --- Filters ---
    st.markdown("### Filters")
    time_options = ["All"] + sorted(edited_df["Time Horizon (Short/Medium/Long)"].dropna().unique())
    hedge_options = ["All", "Yes", "No"]
    time_filter = st.selectbox("Select Time Horizon", time_options)
    hedge_filter = st.selectbox("Inflation Hedge?", hedge_options)
    min_inv_filter = st.slider("Minimum Investment ($)",
        int(edited_df["Minimum Investment ($)"].min()),
        int(edited_df["Minimum Investment ($)"].max()),
        int(edited_df["Minimum Investment ($)"].min())
    )

    filtered = edited_df.copy()
    if time_filter != "All":
        filtered = filtered[filtered["Time Horizon (Short/Medium/Long)"] == time_filter]
    if hedge_filter != "All":
        filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == hedge_filter]
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_inv_filter]

    st.markdown("### Filtered Table")
    st.dataframe(filtered, use_container_width=True)

    # --- Report Generation ---
    def create_ppt(df):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Investment Overview"
        slide.placeholders[1].text = "Automated Matrix Summary"

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Portfolio Averages"
        slide.placeholders[1].text = "\n".join([f"{k}: {v}" for k, v in avg.items()])

        chart_file = "ppt_chart.png"
        plt.figure(figsize=(6, 2))
        plt.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90, fontsize=6); plt.tight_layout(); plt.savefig(chart_file); plt.close()

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Expected Return Chart"
        slide.shapes.add_picture(chart_file, Inches(1), Inches(1.5), width=Inches(6))

        ppt_file = "Investment_Summary.pptx"
        prs.save(ppt_file)
        return ppt_file

    def create_docx(df):
        document = Document()
        document.add_heading("Investment Summary Report", 0)

        avg = df.select_dtypes(include='number').mean(numeric_only=True).round(2)
        document.add_heading("Portfolio Averages", level=1)
        for k, v in avg.items():
            document.add_paragraph(f"{k}: {v}")

        chart_file = "docx_chart.png"
        plt.figure(figsize=(6, 2))
        plt.bar(df["Investment Name"], df["Expected Return (%)"], color="teal")
        plt.xticks(rotation=90, fontsize=6); plt.tight_layout(); plt.savefig(chart_file); plt.close()

        document.add_picture(chart_file, width=DocxInches(6))
        docx_file = "Investment_Summary.docx"
        document.save(docx_file)
        return docx_file

    st.markdown("### Export Reports")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📤 Generate PowerPoint"):
            ppt_path = create_ppt(filtered)
            with open(ppt_path, "rb") as f:
                st.download_button("📥 Download PowerPoint", f, file_name=ppt_path)

    with col2:
        if st.button("📝 Generate Word Document"):
            docx_path = create_docx(filtered)
            with open(docx_path, "rb") as f:
                st.download_button("📥 Download Word Doc", f, file_name=docx_path)
