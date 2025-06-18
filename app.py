import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf

# --- Page Config ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- Styling ---
st.markdown("""
<style>
/* Header */
header > div { background-color: #111 !important; color: white !important; }
/* Sidebar */
[data-testid="stSidebar"] { background-color: #111; color: white; padding: 1rem; }
/* Sidebar titles */
[data-testid="stSidebar"] h2 { color: #f44336; font-weight: 700; }
/* Sidebar text */
[data-testid="stSidebar"] p, label { color: white; font-size: 14px; }
/* Buttons */
.stButton > button {
  background-color: #111; color: #f44336; border: 2px solid #f44336;
  border-radius: 6px; padding: 0.4em 1em; font-weight: 700; font-size: 14px;
  transition: 0.2s;
}
.stButton > button:hover { background-color: #f44336; color: #111; }
/* Main content padding */
.css-1d391kg { padding-top: 110px !important; }
/* Header bar */
.app-header {
  background-color: #111; color: white; padding: 10px 25px;
  position: fixed; top: 0; left: 0; width: 100vw; z-index: 9999;
}
.app-header h1 { margin: 0; font-size: 26px; font-weight: 700; }
.app-header p { margin: 4px 0 0 0; font-size: 13px; color: #f44336; font-weight: 500; }
/* Metrics font */
.stMetric > div { font-size: 14px !important; }
.stMetric > div > div:first-child { font-weight: 700 !important; }
/* Table font */
[data-testid="stDataFrame"] { font-size: 13px; }
</style>

<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization and Financial Advisory.<br>
     Designed for finance professionals to generate, analyze, and optimize portfolios using real-time data.</p>
</div>
""", unsafe_allow_html=True)


# --- Helper Functions ---
@st.cache_data(ttl=600)
def get_index_data(ticker):
    try:
        df = yf.Ticker(ticker).history(period="1mo").reset_index()
        return df
    except:
        return pd.DataFrame()

def format_metric(df):
    if df.empty or len(df) < 2:
        return None, None
    cur, prev = df.iloc[-1], df.iloc[-2]
    diff = cur.Close - prev.Close
    pct = (diff / prev.Close) * 100
    return cur.Close, f"{diff:+.2f} ({pct:+.2f}%)"

def sanitize_string(s):
    if isinstance(s, str):
        return s.strip().replace("–","-").replace("’","'").replace("“",'"').replace("”",'"').replace("•","-").replace("©","(c)")
    return s

# --- Load Data ---
try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = [col.strip() for col in df.columns]
    df = df.applymap(sanitize_string)
except Exception as e:
    st.error(f"Error loading data file: {e}")
    st.stop()

# --- Sidebar: Real-Time Indices (numbers only) ---
st.sidebar.markdown("## Real-Time Market Indices")

for ticker, name in [("^GSPC", "S&P 500"), ("^IXIC", "Nasdaq"), ("^DJI", "Dow Jones")]:
    data = get_index_data(ticker)
    value, delta = format_metric(data)
    if value is None:
        st.sidebar.markdown(f"**{name}: Data unavailable**")
    else:
        st.sidebar.metric(name, f"${value:,.2f}", delta)

# --- Main App: Investment Generator ---
st.header("Investment Generator")

cols_map = {
    "Category": None,
    "Minimum Investment ($)": None,
    "Expected Return (%)": None,
    "Risk Level (1-10)": None
}

# Clean column names
for col in cols_map:
    if col not in df.columns:
        st.error(f"⚠️ Data column missing: {col}")
        st.stop()

# Investment column fallback
inv_col = next((c for c in df.columns if c not in cols_map), df.columns[0])

categories = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("Select Investment Categories:", categories, default=categories)

min_val = st.slider("Min Investment ($)", int(df["Minimum Investment ($)"].min()), int(df["Minimum Investment ($)"].max()), int(df["Minimum Investment ($)"].min()), step=1000)

ret_min = st.slider("Min Expected Return (%)", float(df["Expected Return (%)"].min()), float(df["Expected Return (%)"].max()), float(df["Expected Return (%)"].min()), step=0.1)

risk_max = st.slider("Max Risk Level (1-10)", int(df["Risk Level (1-10)"].min()), int(df["Risk Level (1-10)"].max()), int(df["Risk Level (1-10)"].max()), step=1)

filtered = df[
    (df["Category"].isin(sel_cats)) &
    (df["Minimum Investment ($)"] >= min_val) &
    (df["Expected Return (%)"] >= ret_min) &
    (df["Risk Level (1-10)"] <= risk_max)
]

st.write(f"### {len(filtered)} Investments Found")
st.dataframe(filtered, height=260)

st.divider()

# --- Portfolio Metrics ---
st.header("Portfolio Summary Metrics")
m1,m2,m3,m4,m5,m6,m7 = st.columns(7)
m1.metric("Avg Return (%)", f"{filtered['Expected Return (%)'].mean():.2f}%" if not filtered.empty else "N/A")
m2.metric("Avg Risk (1–10)", f"{filtered['Risk Level (1-10)'].mean():.2f}" if not filtered.empty else "N/A")
m3.metric("Avg Cap Rate (%)", f"{filtered['Cap Rate (%)'].mean():.2f}%" if 'Cap Rate (%)' in filtered.columns else "N/A")
m4.metric("Avg Liquidity", f"{filtered['Liquidity (1–10)'].mean():.2f}" if 'Liquidity (1–10)' in filtered.columns else "N/A")
m5.metric("Avg Volatility", f"{filtered['Volatility (1–10)'].mean():.2f}" if 'Volatility (1–10)' in filtered.columns else "N/A")
m6.metric("Avg Fees (%)", f"{filtered['Fees (%)'].mean():.2f}%" if 'Fees (%)' in filtered.columns else "N/A")
m7.metric("Avg Min Investment", f"${filtered['Minimum Investment ($)'].mean():,.0f}" if not filtered.empty else "N/A")

st.divider()

# --- Visual Insights ---
st.header("Visual Insights")
c1,c2,c3,c4 = st.columns(4)

def compact_fig():
    fig, ax = plt.subplots(figsize=(3,2))
    ax.tick_params(axis='both', labelsize=7)
    return fig, ax

with c1:
    st.markdown("**Expected Return (%)**")
    fig, ax = compact_fig()
    if not filtered.empty:
        ax.bar(filtered[inv_col], filtered["Expected Return (%)"], color='#f44336')
        ax.set_xticklabels(filtered[inv_col], rotation=45, ha='right')
    ax.set_ylabel("Return %", fontsize=8)
    st.pyplot(fig)

with c2:
    st.markdown("**Volatility vs Liquidity**")
    fig, ax = compact_fig()
    if not filtered.empty:
        ax.scatter(filtered["Volatility (1-10)"], filtered["Liquidity (1-10)"], c='red', alpha=0.7)
    ax.set_xlabel("Volatility", fontsize=8); ax.set_ylabel("Liquidity", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    st.pyplot(fig)

with c3:
    st.markdown("**Fees vs Expected Return**")
    fig, ax = compact_fig()
    if not filtered.empty and "Fees (%)" in filtered.columns:
        ax.scatter(filtered["Fees (%)"], filtered["Expected Return (%)"], c='red', alpha=0.7)
    ax.set_xlabel("Fees %", fontsize=8); ax.set_ylabel("Return %", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    st.pyplot(fig)

with c4:
    st.markdown("**Risk Distribution**")
    fig, ax = compact_fig()
    if not filtered.empty:
        ax.hist(filtered["Risk Level (1-10)"], bins=10, color='#f44336', alpha=0.8)
    ax.set_xlabel("Risk Level", fontsize=8); ax.set_ylabel("Count", fontsize=8)
    st.pyplot(fig)

st.divider()

# --- Export Reports ---
st.header("Export Reports")
d1, d2 = st.columns(2)
with d1:
    if st.button("📤 Download PowerPoint Report"):
        st.success("PowerPoint report ready (functionality to implement)")
with d2:
    if st.button("📥 Download Word Report"):
        st.success("Word report ready (functionality to implement)")
