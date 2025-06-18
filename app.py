import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import yfinance as yf

# Page config
st.set_page_config(page_title="Automated Investment Matrix", layout="wide", initial_sidebar_state="collapsed")

# === Styling & Header ===
st.markdown("""
<style>
body { font-family: 'Segoe UI', sans-serif; }
#header-box {
    position:fixed; top:0; left:0; width:100vw; height:70px;
    background:#111; color:white; display:flex;
    flex-direction:column; justify-content:center;
    padding-left:20px; border-bottom:3px solid red; z-index:9999;
}
#header-title { margin:0; font-size:22px; font-weight:700; }
#header-sub { margin:0; font-size:13px; color:#ddd; }
.main-container { padding-top:80px; margin-right:240px; }
#market-box {
    position:fixed; top:70px; right:0; width:220px;
    background:#111; color:white; padding:15px;
    border-left:3px solid red; height:calc(100vh-70px);
    overflow-y:auto; font-family:'Segoe UI',sans-serif;
}
#market-box h2 { text-align:center; color:red; margin-bottom:12px; }
.metric-label { font-size:13px; margin-top:8px; }
.metric-value { font-size:15px; font-weight:700; }
.metric-delta { font-size:13px; }
hr { border:0.5px solid #333; margin:8px 0; }
.stButton > button {
    background:#111; color:red; border:2px solid red;
    border-radius:4px; padding:0.3em 0.8em; font-weight:700;
}
.stButton > button:hover { background:red; color:#111; }
[data-testid="stDataFrame"] { font-size:12px; }
</style>
<div id="header-box">
  <p id="header-title">Automated Investment Matrix</p>
  <p id="header-sub">Portfolio Analysis Platform with Real-Time Data & Optimization Tools</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="main-container"></div>', unsafe_allow_html=True)

# === Market Indices Box ===
@st.cache_data(ttl=300)
def fetch_indices():
    out = {}
    for name, tick in [("S&P 500","^GSPC"), ("Nasdaq","^IXIC"), ("Dow Jones","^DJI")]:
        df = yf.Ticker(tick).history(period="2d")['Close']
        if len(df)>=2:
            curr, prev = df.iloc[-1], df.iloc[-2]
            delta = curr - prev
            pct = delta/prev*100
            out[name] = (curr, delta, pct)
    return out

indices = fetch_indices()
html = "<div id='market-box'><h2>Market Indices</h2>"
for name, (curr, delta, pct) in indices.items():
    clr = "#4caf50" if delta>0 else "#f44336"
    sign = "+" if delta>0 else ""
    html += f"<div class='metric-label'>{name}</div>"
    html += f"<div class='metric-value'>${curr:,.2f}</div>"
    html += f"<div class='metric-delta' style='color:{clr};'>{sign}{delta:.2f} ({sign}{pct:.2f}%)</div><hr>"
html += "</div>"
st.markdown(html, unsafe_allow_html=True)

# === Load & Sanitize Data ===
df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
for col in ["Expected Return (%)","Risk (%)","Volatility (%)","Liquidity (1–10)","Minimum Investment ($)","Fees (%)","Cap Rate (%)"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

# --- Investment Type Selection ---
st.markdown("### Select Investment Types")
if "Category" in df.columns:
    options = df["Category"].dropna().unique().tolist()
elif "Investment" in df.columns:
    options = df["Investment"].dropna().unique().tolist()
else:
    options = []
selected = st.multiselect("Include investments:", options, default=options)
filtered = df[df[df.keys()[0]].isin(selected)] if options else df.copy()

# Editable table
st.markdown("### Edit Investment Data")
edited = st.data_editor(filtered, use_container_width=True, num_rows="dynamic")

# Portfolio filters
st.markdown("### Portfolio Filters")
ret_s = st.slider("Min Expected Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()))
risk_s = st.slider("Max Risk (%)", float(edited["Risk (%)"].min()), float(edited["Risk (%)"].max()), float(edited["Risk (%)"].max()))
liq_s = st.slider("Min Liquidity (1–10)", int(edited["Liquidity (1–10)"].min()), int(edited["Liquidity (1–10)"].max()), int(edited["Liquidity (1–10)"].min()))
mininv_s = st.slider("Max Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].max()))
hedge = st.selectbox("Inflation Hedge", ["All", "Yes", "No"])
horizon = st.selectbox("Time Horizon", ["All"] + (df["Time Horizon (Short/Medium/Long)"].dropna().unique().tolist() if "Time Horizon (Short/Medium/Long)" in df.columns else []))

# Apply filters
f2 = edited.copy()
f2 = f2[f2["Expected Return (%)"] >= ret_s]
f2 = f2[f2["Risk (%)"] <= risk_s]
f2 = f2[f2["Liquidity (1–10)"] >= liq_s]
f2 = f2[f2["Minimum Investment ($)"] <= mininv_s]
if hedge != "All" and "Inflation Hedge (Yes/No)" in f2.columns:
    f2 = f2[f2["Inflation Hedge (Yes/No)"] == hedge]
if horizon != "All" and "Time Horizon (Short/Medium/Long)" in f2.columns:
    f2 = f2[f2["Time Horizon (Short/Medium/Long)"] == horizon]

# Averages
st.markdown("### Portfolio Averages")
cols = st.columns(7)
metrics = ["Expected Return (%)","Risk (%)","Volatility (%)","Liquidity (1–10)","Cap Rate (%)","Fees (%)","Minimum Investment ($)"]
for col, slot in zip(metrics, cols):
    if col in edited.columns:
        val = edited[col].mean()
        txt = f"{val:.2f}%" if "%" in col or "Cap Rate" in col else f"${val:,.0f}"
        slot.metric(col.replace(" (%)",""), txt)

# Visual Charts
st.markdown("### Visual Insights")
c1,c2 = st.columns(2)
with c1:
    fig, ax = plt.subplots(figsize=(5,3))
    ax.bar(edited.iloc[:,0], edited["Expected Return (%)"], color="red")
    ax.set_xticklabels(edited.iloc[:,0], rotation=45, ha='right', fontsize=8)
    ax.set_ylabel("Return (%)")
    st.pyplot(fig)
with c2:
    fig, ax = plt.subplots(figsize=(5,3))
    ax.scatter(edited["Volatility (%)"], edited["Liquidity (1–10)"], color="red")
    ax.set_xlabel("Volatility (%)"); ax.set_ylabel("Liquidity (1–10)")
    st.pyplot(fig)

# More plots
c3,c4 = st.columns(2)
with c3:
    fig, ax = plt.subplots(figsize=(5,3))
    ax.scatter(edited["Fees (%)"], edited["Expected Return (%)"], color="red")
    ax.set_xlabel("Fees (%)"); ax.set_ylabel("Expected Return (%)")
    st.pyplot(fig)
with c4:
    fig, ax = plt.subplots(figsize=(5,3))
    ax.scatter(edited["Risk (%)"], edited["Expected Return (%)"], color="red")
    ax.set_xlabel("Risk (%)"); ax.set_ylabel("Expected Return (%)")
    st.pyplot(fig)

# Filtered table
st.markdown(f"### Filtered {len(f2)} Investments Matching Criteria")
st.dataframe(f2, use_container_width=True)

# Report export buttons
st.markdown("### Generate Reports")
col1, col2 = st.columns(2)
with col1:
    if st.button("📤 Export PowerPoint"):
        st.success("PowerPoint Export Placeholder")
with col2:
    if st.button("📥 Export Word"):
        st.success("Word Export Placeholder")

