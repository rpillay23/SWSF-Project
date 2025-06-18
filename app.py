import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# === Page Config & Styling ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
.css-1d391kg { padding-top: 110px !important; }
.app-header { background:#111;color:white;padding:8px 20px;position:fixed;top:0;left:0;width:100vw;z-index:999;}
.app-header h1{margin:0;font-size:20px;font-weight:700;}
.app-header p{margin:2px 0 0;color:#f44336;font-size:11px;}
main { margin-right:200px !important; padding:20px; }
.stButton > button{background:#111;color:#f44336;border:2px solid #f44336;border-radius:4px;padding:0.3em 0.8em;font-weight:700;font-size:12px;}
.stButton > button:hover{background:#f44336;color:#111;}
[data-testid="stDataFrame"] { font-size:12px; }
#right-panel {
    position: fixed; top:80px; right:0; width:180px;
    height:calc(100vh - 80px); background:#111; color:white;
    padding:12px 10px; z-index:998;
    box-shadow:-2px 0 5px rgba(0,0,0,0.3);
    overflow-y:auto;
}
#right-panel h2 { color:#f44336; font-size:16px; margin-bottom:8px; text-align:center; }
.metric-label { font-size:13px; margin-top:12px; }
.metric-value { font-size:16px; font-weight:700; }
.metric-delta { font-size:13px; }
</style>
<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Portfolio Analysis Platform with Real‑Time Data</p>
</div>
""", unsafe_allow_html=True)

# === Fetch Index Metrics ===
@st.cache_data(ttl=300)
def fetch_index(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        cur, prev = hist.iloc[-1], hist.iloc[-2]
        diff = cur - prev
        pct = diff / prev * 100
        return f"${cur:,.2f}", f"{diff:+.2f} ({pct:+.2f}%)"
    return None, None

indices = {
    "S&P 500": fetch_index("^GSPC"),
    "Nasdaq": fetch_index("^IXIC"),
    "Dow Jones": fetch_index("^DJI")
}

# === Right-side Panel ===
st.markdown('<div id="right-panel">', unsafe_allow_html=True)
st.markdown('<h2>Market Indices</h2>', unsafe_allow_html=True)
for name, (price, delta) in indices.items():
    if price:
        color = "#4caf50" if delta.startswith("+") else "#f44336"
        st.markdown(f'<div class="metric-label">{name}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{price}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-delta" style="color:{color};">{delta}</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# === Load & Prepare Investment Data ===
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# 1️⃣ Select Categories
st.subheader("1. Select Investment Types")
if "Category" not in df:
    st.error("⚠️ Missing 'Category' column.")
    st.stop()
all_cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("", all_cats, default=all_cats)

# 2️⃣ Editable Table
st.subheader("2. Editable Investment Data")
edited = st.data_editor(df[df["Category"].isin(sel_cats)], use_container_width=True, num_rows="fixed")

# Helper for safe averages
def safe_avg(df, c):
    return df[c].mean() if c in df.columns and not df.empty else None

# 3️⃣ Portfolio Averages
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
for (label, col_nm), slot in zip(fields, cols):
    m = safe_avg(edited, col_nm)
    if m is not None:
        val = f"{m:.2f}" if "%" in label or "Rate" in label else (f"${m:.0f}" if "Invest" in label else f"{m:.2f}")
    else:
        val = "N/A"
    suffix = "%" if "%" in label or "Rate" in label else ""
    slot.metric(label, f"{val}{suffix}")

# 4️⃣ Compact Charts
st.subheader("4. Visual Insights")
chart_cols = st.columns(4)

def bar_plot(xs, ys):
    fig, ax = plt.subplots(figsize=(2.5,1.8))
    ax.bar(xs, ys, color="#f44336")
    ax.tick_params(axis="x", rotation=45, labelsize=6)
    ax.tick_params(axis="y", labelsize=6)
    fig.tight_layout()
    return fig

def scatter_plot(xs, ys, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(2.5,1.8))
    ax.scatter(xs, ys, c="red", alpha=0.6)
    ax.set_xlabel(xlabel, fontsize=6)
    ax.set_ylabel(ylabel, fontsize=6)
    ax.tick_params(labelsize=6)
    ax.grid(True, linestyle="--", alpha=0.3)
    fig.tight_layout()
    return fig

charts = [
    ("Expected Return (%)", bar_plot),
    (("Volatility (1-10)", "Liquidity (1-10)"), scatter_plot),
    (("Fees (%)", "Expected Return (%)"), scatter_plot),
    ("Risk Level (1-10)", None)
]

for slot, cfg in zip(chart_cols, charts):
    slot.header("")  # clear default margin
    if isinstance(cfg[0], tuple):
        x_col, y_col = cfg[0]
        slot.markdown(f"**{y_col} vs {x_col}**")
        if x_col in edited and y_col in edited and not edited.empty:
            slot.pyplot(cfg[1](edited[x_col], edited[y_col], x_col, y_col))
    else:
        col_name = cfg[0]
        slot.markdown(f"**{col_name}**")
        if col_name in edited and not edited.empty:
            if cfg[1]:
                slot.pyplot(cfg[1](edited["Investment Name"], edited[col_name]))
            else:
                fig, ax = plt.subplots(figsize=(2.5,1.8))
                ax.hist(edited[col_name], bins=8, color="#f44336", alpha=0.7)
                ax.tick_params(labelsize=6)
                fig.tight_layout()
                slot.pyplot(fig)

# 5️⃣ Bottom Constraints
st.subheader("5. Portfolio Constraints")
min_i = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].min()), step=1000) if "Minimum Investment ($)" in edited else 0
min_r = st.slider("Min Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()), step=0.1) if "Expected Return (%)" in edited else 0
max_r = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()), int(edited["Risk Level (1-10)"].max()), int(edited["Risk Level (1-10)"].max()), step=1) if "Risk Level (1-10)" in edited else 10
horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"], index=1)
hedge = st.checkbox("Inflation Hedge Only")

filtered = edited.copy()
if "Minimum Investment ($)" in filtered:
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_i]
if "Expected Return (%)" in filtered:
    filtered = filtered[filtered["Expected Return (%)"] >= min_r]
if "Risk Level (1-10)" in filtered:
    filtered = filtered[filtered["Risk Level (1-10)"] <= max_r]
if hedge and "Inflation Hedge (Yes/No)" in filtered:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"]

# 6️⃣ Filtered Data Table
st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=200)

# 7️⃣ Export Buttons
st.subheader("7. Export Reports")
exp1, exp2 = st.columns(2)
with exp1:
    if st.button("📤 Export PowerPoint"):
        st.success("PowerPoint export placeholder")
with exp2:
    if st.button("📥 Export Word"):
        st.success("Word export placeholder")
