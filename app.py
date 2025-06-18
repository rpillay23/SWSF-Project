import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- Page Setup & Styling ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
.css-1d391kg { padding-top:110px !important; }
.app-header { background:#111; color:white; padding:8px 20px; position:fixed;top:0;left:0;width:100vw;z-index:999;}
.app-header h1{margin:0;font-size:20px;font-weight:700;}
.app-header p{margin:2px 0 0;color:#f44336;font-size:11px;}
.stButton > button{background:#111;color:#f44336;border:2px solid #f44336;border-radius:4px;padding:0.3em 0.8em;font-weight:700;font-size:12px;}
.stButton > button:hover{background:#f44336;color:#111;}
main { margin-right:300px !important; padding:0 20px 20px 20px; }
[data-testid="stDataFrame"] { font-size:12px; }
</style>
<div class="app-header"><h1>Automated Investment Matrix</h1><p>Portfolio Analysis Platform with Real-Time Data</p></div>
""", unsafe_allow_html=True)

# --- Right Fixed Index Panel ---
@st.cache_data(ttl=300)
def get_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        l, p = hist.iloc[-1], hist.iloc[-2]
        d = l - p
        pct = d / p * 100
        return f"${l:,.2f}", f"{d:+.2f} ({pct:+.2f}%)"
    return "N/A", ""

prices = {}
for t, name in [("^GSPC", "S&P 500"), ("^IXIC", "Nasdaq"), ("^DJI", "Dow Jones")]:
    prices[name] = get_price(t)

# Market indices container with black background on right fixed margin
st.markdown(
    """
    <div style="
        position: fixed;
        top: 80px;
        right: 0;
        width: 280px;
        background: #111;
        color: white;
        height: calc(100vh - 80px);
        padding: 15px;
        overflow-y: auto;
        box-shadow: -2px 0 5px rgba(0,0,0,0.4);
        z-index: 998;
        font-family: inherit;
    ">
    <h2 style="color:#f44336; margin-bottom:10px;">Market Indices</h2>
    """
    , unsafe_allow_html=True)

for name, (val, delta) in prices.items():
    color = "#4caf50" if delta.startswith("+") else "#f44336"
    st.markdown(
        f"""
        <div style="margin-bottom: 15px;">
            <div style="font-size:13px; font-weight:700;">{name}</div>
            <div style="font-size:18px; margin-top: 2px;">{val}</div>
            <div style="font-size:13px; color:{color}; margin-top: 2px;">{delta}</div>
            <hr style="border-color:#444; margin: 8px 0;">
        </div>
        """, unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)  # close container div

# --- Load Data ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# --- 1) Select Categories ---
st.subheader("1. Select Investment Types")
if "Category" not in df.columns:
    st.error("Missing 'Category' column in data.")
    st.stop()
cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("", cats, default=cats)

# --- 2) Editable Table ---
st.subheader("2. Editable Investment Data")
edited = st.data_editor(df[df["Category"].isin(sel_cats)], use_container_width=True, num_rows="fixed")

# --- Helper for safe column mean ---
def mean_or_na(df, col):
    return f"{df[col].mean():.2f}" if col in df.columns and not df.empty else "N/A"

# --- 3) Portfolio Averages ---
st.subheader("3. Portfolio Averages")
fields = [
    ("Avg Return (%)", "Expected Return (%)"),
    ("Avg Risk", "Risk Level (1-10)"),
    ("Avg Cap Rate (%)", "Cap Rate (%)"),
    ("Avg Liquidity", "Liquidity (1-10)"),
    ("Avg Volatility", "Volatility (1-10)"),
    ("Avg Fees (%)", "Fees (%)"),
    ("Avg Min Invest ($)", "Minimum Investment ($)")
]
cols = st.columns(len(fields))
for (label, col), panel in zip(fields, cols):
    val = mean_or_na(edited, col)
    panel.metric(label, f"{val}%" if "%" in label or "Rate" in label else (f"${val}" if "Invest" in label else val))

# --- 4) Visual Charts ---
st.subheader("4. Visual Insights")
charts = st.columns(4)
def chart_bar(x, y):
    fig, ax = plt.subplots(figsize=(2.7, 1.8))
    ax.bar(x, y, color="#f44336")
    ax.tick_params(axis="x", rotation=45, labelsize=6)
    fig.tight_layout()
    return fig

def chart_scatter(x, y, xl, yl):
    fig, ax = plt.subplots(figsize=(2.7, 1.8))
    ax.scatter(x, y, c="red", alpha=0.6)
    ax.set_xlabel(xl, fontsize=7)
    ax.set_ylabel(yl, fontsize=7)
    ax.tick_params(labelsize=6)
    fig.tight_layout()
    return fig

mapping = [
    ("Expected Return (%)", chart_bar),
    (("Volatility (1-10)", "Liquidity (1-10)"), chart_scatter),
    (("Fees (%)", "Expected Return (%)"), chart_scatter),
    ("Risk Level (1-10)", None)  # special histogram
]

for slot, cfg in zip(charts, mapping):
    if isinstance(cfg[0], tuple):
        x, y = cfg[0]
        label = f"{y} vs {x}"
        slot.markdown(f"**{label}**")
        if x in edited and y in edited and not edited.empty:
            func = cfg[1]
            slot.pyplot(func(edited[x], edited[y], x, y))
    else:
        col = cfg[0]
        slot.markdown(f"**{col}**")
        if col in edited and not edited.empty:
            if cfg[1]:
                slot.pyplot(cfg[1](edited["Investment Name"], edited[col]))
            else:
                fig, ax = plt.subplots(figsize=(2.7, 1.8))
                ax.hist(edited[col], bins=8, color="#f44336", alpha=0.7)
                ax.tick_params(labelsize=6)
                fig.tight_layout()
                slot.pyplot(fig)

# --- 5) Bottom Filters ---
st.subheader("5. Portfolio Constraints")
min_inv = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].min()), step=1000) if "Minimum Investment ($)" in edited.columns else 0
min_ret = st.slider("Min Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()), step=0.1) if "Expected Return (%)" in edited.columns else 0
max_risk = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()), int(edited["Risk Level (1-10)"].max()), int(edited["Risk Level (1-10)"].max()), step=1) if "Risk Level (1-10)" in edited.columns else 10
time_horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"], index=1)
hedge = st.checkbox("Inflation Hedge Only")

filtered = edited.copy()
if "Minimum Investment ($)" in filtered.columns:
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_inv]
if "Expected Return (%)" in filtered.columns:
    filtered = filtered[filtered["Expected Return (%)"] >= min_ret]
if "Risk Level (1-10)" in filtered.columns:
    filtered = filtered[filtered["Risk Level (1-10)"] <= max_risk]
if hedge and "Inflation Hedge (Yes/No)" in filtered.columns:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"]

# --- 6) Filtered Table & Export ---
st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=220)

st.subheader("7. Export Reports")
b1, b2 = st.columns(2)
with b1:
    if st.button("📤 Download PowerPoint"):
        st.success("PPT export placeholder")
with b2:
    if st.button("📥 Download Word"):
        st.success("Word export placeholder")
