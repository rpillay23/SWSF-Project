import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- PAGE CONFIG & STYLING ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
/* Fixed header */
.css-1d391kg { padding-top: 110px !important; }
.app-header { background: #111; color: white; padding: 8px 20px;
    position: fixed; top: 0; left: 0; width: 100vw; z-index: 999; }
.app-header h1 { margin:0; font-size:20px; font-weight:700; }
.app-header p { margin:2px 0 0;color:#f44336;font-size:11px; }

/* Compact buttons */
.stButton > button {
    background:#111;color:#f44336;border:2px solid #f44336;
    border-radius:4px;padding:0.3em 0.8em;font-weight:700;font-size:12px;
}
.stButton > button:hover { background:#f44336;color:#111; }

/* Main content spacing */
main {
    margin-right: 300px !important;
    padding: 0 20px 20px 20px;
}

/* Right fixed panel */
#right-panel {
    position: fixed; top:80px; right:0;
    width: 280px; height: calc(100vh - 80px);
    background:#111; padding:15px; color:white;
    box-shadow: -2px 0 5px rgba(0,0,0,0.4);
    overflow-y: auto; z-index: 998;
}
#right-panel h2 { color:#f44336; font-size:18px; margin-bottom:10px; }
#right-panel .metric-label { font-size:14px; margin-top:10px; }
#right-panel .metric-value { font-size:18px; font-weight:700; }
#right-panel .metric-delta { font-size:14px; }

[data-testid="stDataFrame"] { font-size:12px; }

/* Compact plot labels */
.tick-label { font-size:7px !important; }
</style>

<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Modular Investment Analysis Platform with Real-Time Data</p>
</div>
""", unsafe_allow_html=True)

# --- INDEX DATA FUNCTION ---
@st.cache_data(ttl=300)
def get_idx_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        latest, prev = hist.iloc[-1], hist.iloc[-2]
        diff = latest - prev
        pct = diff / prev * 100
        return latest, f"{diff:+.2f} ({pct:+.2f}%)"
    return None, None

# --- FIXED RIGHT PANEL ---
prices = {}
for t,name in [("^GSPC","S&P 500"),("^IXIC","Nasdaq"),("^DJI","Dow Jones")]:
    p, d = get_idx_price(t)
    prices[name] = (p,d)

st.markdown('<div id="right-panel"><h2>Market Indices</h2></div>', unsafe_allow_html=True)
for name,(p,d) in prices.items():
    if p is not None:
        st.markdown(f"""
        <div id="right-panel">
          <div class="metric-label">{name}</div>
          <div class="metric-value">${p:,.2f}</div>
          <div class="metric-delta" style="color:{'#4caf50' if d.startswith('+') else '#f44336'}">{d}</div>
        </div>
        """, unsafe_allow_html=True)

# --- LOAD AND PREPARE DATA ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df
df = load_data()

# --- SELECTION, EDIT, FILTER, GRAPHS SEQUENCE ---
st.subheader("1. Select Investment Types")
cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("", cats, default=cats)

st.subheader("2. Editable Investment Data")
edited = st.data_editor(df[df["Category"].isin(sel_cats)], use_container_width=True, num_rows="fixed")

st.subheader("3. Portfolio Averages")
c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric("Return (%)", f"{edited['Expected Return (%)'].mean():.2f}%")
c2.metric("Risk (1–10)", f"{edited['Risk Level (1-10)'].mean():.2f}")
c3.metric("Cap Rate (%)", f"{edited['Cap Rate (%)'].mean():.2f}%")
c4.metric("Liquidity", f"{edited['Liquidity (1-10)'].mean():.2f}")
c5.metric("Volatility", f"{edited['Volatility (1-10)'].mean():.2f}")
c6.metric("Fees (%)", f"{edited['Fees (%)'].mean():.2f}%")
c7.metric("Min Invest ($)", f"${edited['Minimum Investment ($)'].mean():,.0f}")

st.subheader("4. Visual Insights")
cols = st.columns(4)
def plot_compact(x,y,kind):
    fig,ax = plt.subplots(figsize=(2.5,1.8))
    if kind=='bar':
        ax.bar(x,y,color="#f44336")
        ax.set_xticklabels(x, rotation=45, ha='right', fontsize=7)
    else:
        ax.scatter(x,y,c='red',alpha=0.7)
    ax.tick_params(axis='both', labelsize=7)
    ax.grid(True, linestyle='--', alpha=0.3)
    plt.tight_layout()
    return fig

with cols[0]:
    st.markdown("Expected Return")
    if not edited.empty: st.pyplot(plot_compact(edited["Investment Name"], edited["Expected Return (%)"], "bar"))
with cols[1]:
    st.markdown("Volatility vs Liquidity")
    if 'Volatility (1-10)' in edited and 'Liquidity (1-10)' in edited:
        st.pyplot(plot_compact(edited['Volatility (1-10)'], edited['Liquidity (1-10)'], "scatter"))
with cols[2]:
    st.markdown("Fees vs Return")
    if 'Fees (%)' in edited:
        st.pyplot(plot_compact(edited["Fees (%)"], edited["Expected Return (%)"], "scatter"))
with cols[3]:
    st.markdown("Risk Distribution")
    fig,ax = plt.subplots(figsize=(2.5,1.8))
    if not edited.empty:
        ax.hist(edited["Risk Level (1-10)"], bins=8, color="#f44336", alpha=0.7)
    ax.tick_params(axis='both', labelsize=7)
    plt.tight_layout()
    st.pyplot(fig)

st.subheader("5. Apply Portfolio Constraints")
min_inv = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].min()), step=1000)
min_ret = st.slider("Min Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()), step=0.1)
max_risk = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()), int(edited["Risk Level (1-10)"].max()), int(edited["Risk Level (1-10)"].max()), step=1)
horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"], index=1)
hedge = st.checkbox("Include inflation-hedge only")

filtered = edited[
    (edited["Minimum Investment ($)"] >= min_inv) &
    (edited["Expected Return (%)"] >= min_ret) &
    (edited["Risk Level (1-10)"] <= max_risk)
]
if hedge and "Inflation Hedge (Yes/No)" in filtered:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"]

st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=200)

st.subheader("7. Export Reports")
e1,e2 = st.columns(2)
with e1:
    if st.button("📤 Download PowerPoint Report"):
        st.success("PowerPoint export placeholder")
with e2:
    if st.button("📥 Download Word Report"):
        st.success("Word export placeholder")

