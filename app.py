import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches

# --- PAGE CONFIG ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- STYLING & HEADER ---
st.markdown("""
<style>
.app-header {
    background:#111;
    color:white;
    padding:12px 24px;
    position:fixed;
    top:0;
    left:0;
    width:100vw;
    z-index:1000;
    font-family:'Helvetica Neue', Helvetica, sans-serif;
}
.app-header h1 {
    margin:0;
    font-size:24px;
    font-weight:700;
}
.app-header p {
    margin:4px 0 0;
    color:#f44336;
    font-size:13px;
}
.market-indices {
    display:flex;
    gap:20px;
    margin-top:40px;
    margin-bottom:20px;
    justify-content:center;
}
.index-box {
    background:#111;
    color:white;
    padding:10px 16px;
    border-radius:4px;
    text-align:center;
    font-size:13px;
    min-width:120px;
}
.index-box .name { font-weight:700; margin-bottom:4px; }
.index-box .value { font-size:16px; margin-bottom:2px; }
.index-box .delta { font-size:14px; }
.index-box .positive { color:#4caf50; }
.index-box .negative { color:#f44336; }
[data-testid="stDataFrame"] { font-size:12px; }
.stButton > button {
    background:#111;
    color:#f44336;
    border:2px solid #f44336;
    padding:0.5em 1em;
    font-weight:700;
    border-radius:6px;
}
.stButton > button:hover {
    background:#f44336;
    color:#111;
}
</style>
<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Real‑Time Data for Portfolio Optimization and Advisory</p>
</div>
""", unsafe_allow_html=True)

# --- MARKET INDEXES ---
@st.cache_data(ttl=300)
def get_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        last, prev = hist.iloc[-1], hist.iloc[-2]
        delta = last - prev
        pct = delta / prev * 100
        return f"${last:,.2f}", f"{delta:+.2f} ({pct:+.2f}%)"
    return "N/A", ""

tickers = {"S&P 500": "^GSPC", "Nasdaq": "^IXIC", "Dow Jones": "^DJI"}
prices = {n: get_price(t) for n, t in tickers.items()}

market_html = '<div class="market-indices">'
for name, (price, delta) in prices.items():
    cls = "positive" if delta.startswith("+") else "negative"
    market_html += f'''
    <div class="index-box">
        <div class="name">{name}</div>
        <div class="value">{price}</div>
        <div class="delta {cls}">{delta}</div>
    </div>'''
market_html += '</div>'
st.markdown(market_html, unsafe_allow_html=True)

# --- DATA LOAD ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = [c.strip().replace("–", "-") for c in df.columns]  # ensure consistent dashes
    return df

df = load_data()

# --- 1. SELECT INVESTMENTS ---
st.subheader("1. Select Investment Types")
if "Category" not in df.columns:
    st.error("Missing Category column in data.")
    st.stop()
types = sorted(df["Category"].dropna().unique())
sel = st.multiselect("Choose investment categories:", types, default=types)

# --- 2. EDITABLE DATA ---
st.subheader("2. Editable Investment Data")
filtered_df = df[df["Category"].isin(sel)].copy()
edited = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")

# --- 3. PORTFOLIO METRICS ---
st.subheader("3. Portfolio Averages")
def mean_safe(df, col): return f"{df[col].mean():.2f}" if col in df.columns and not df.empty else "N/A"

metrics = [
    ("Avg Return (%)", "Expected Return (%)"),
    ("Avg Risk", "Risk Level (1-10)"),
    ("Avg Cap Rate (%)", "Cap Rate (%)"),
    ("Avg Liquidity", "Liquidity (1-10)"),
    ("Avg Volatility", "Volatility (1-10)"),
    ("Avg Fees (%)", "Fees (%)"),
    ("Avg Min Inv ($)", "Minimum Investment ($)")
]
cols = st.columns(len(metrics))
for (lbl, colname), c in zip(metrics, cols):
    val = mean_safe(edited, colname)
    suffix = "%" if "%" in lbl or "Rate" in lbl else "$" if "Inv" in lbl else ""
    c.metric(lbl, f"{val}{suffix}" if val != "N/A" else val)

# --- 4. VISUAL INSIGHTS ---
st.subheader("4. Visual Insights")
vc = st.columns(4)

def make_bar(x, y):
    fig, ax = plt.subplots(figsize=(2.6,1.8))
    ax.bar(x, y, color="#f44336")
    ax.tick_params(axis="x", rotation=45, labelsize=6)
    fig.tight_layout(); return fig

def make_scatter(x, y, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(2.6,1.8))
    ax.scatter(x, y, c='red', alpha=0.6, edgecolors='black')
    ax.set_xlabel(xlabel, fontsize=7); ax.set_ylabel(ylabel, fontsize=7)
    ax.tick_params(labelsize=6); fig.tight_layout(); return fig

vis = [
    ("Expected Return (%)", make_bar),
    (("Volatility (1-10)", "Liquidity (1-10)"), make_scatter),
    (("Fees (%)", "Expected Return (%)"), make_scatter),
    ("Risk Level (1-10)", None)
]
for slot, cfg in zip(vc, vis):
    if isinstance(cfg[0], tuple):
        x, y = cfg[0]
        slot.markdown(f"**{y} vs {x}**")
        if x in edited and y in edited and not edited.empty:
            slot.pyplot(cfg[1](edited[x], edited[y], x, y))
    else:
        col = cfg[0]
        slot.markdown(f"**{col}**")
        if col in edited and not edited.empty:
            if cfg[1]:
                slot.pyplot(cfg[1](edited["Investment Name"], edited[col]))
            else:
                fig, ax = plt.subplots(figsize=(2.6,1.8))
                ax.hist(edited[col], bins=7, color="#f44336", alpha=0.7)
                ax.tick_params(labelsize=6); fig.tight_layout(); slot.pyplot(fig)

# --- 5. FILTER -->
st.subheader("5. Portfolio Constraints")
min_i = st.slider("Min Investment ($)", 0, int(edited["Minimum Investment ($)"].max()), 0, step=1000) \
    if "Minimum Investment ($)" in edited else 0
min_r = st.slider("Min Return (%)", 0.0, float(edited["Expected Return (%)"].max()), 0.0, step=0.1) \
    if "Expected Return (%)" in edited else 0
max_r = st.slider("Max Risk level", 0, 10, 10)
time_h = st.selectbox("Time Horizon", ["Short", "Medium", "Long"])
hedge = st.checkbox("Inflation Hedge Only")

f = edited.copy()
if "Minimum Investment ($)" in f: f = f[f["Minimum Investment ($)"] >= min_i]
if "Expected Return (%)" in f: f = f[f["Expected Return (%)"] >= min_r]
if "Risk Level (1-10)" in f: f = f[f["Risk Level (1-10)"] <= max_r]
if hedge and "Inflation Hedge (Yes/No)" in f: f = f[f["Inflation Hedge (Yes/No)"] == "Yes"]

# --- 6. FILTERED TABLE ---
st.subheader(f"6. Filtered Investments ({len(f)})")
st.dataframe(f, height=220)

# --- 7. EXPORT REPORTS ---
st.subheader("7. Export Reports")
b1, b2 = st.columns(2)
with b1:
    if st.button("📤 Download PowerPoint"):
        st.success("PPT placeholder")
with b2:
    if st.button("📥 Download Word"):
        st.success("Word placeholder")
