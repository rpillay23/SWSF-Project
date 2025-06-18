import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- Page Setup & Styling ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
.css-1d391kg { padding-top: 110px !important; }
.app-header { background:#111;color:white;padding:8px 20px;position:fixed;top:0;left:0;width:100vw;z-index:999;}
.app-header h1{margin:0;font-size:20px;font-weight:700;}
.app-header p{margin:2px 0 0;color:#f44336;font-size:11px;}
.stButton > button{background:#111;color:#f44336;border:2px solid #f44336;border-radius:4px;padding:0.3em 0.8em;font-weight:700;font-size:12px;}
.stButton > button:hover{background:#f44336;color:#111;}
main { margin-right:200px !important; padding:0 20px 20px 20px; }
[data-testid="stDataFrame"] { font-size:12px; }

/* Right fixed panel */
#right-panel {
    position: fixed; top:80px; right:0;
    width: 180px; height: calc(100vh - 80px);
    background:#111; padding:10px 8px; color:white;
    box-shadow:-2px 0 5px rgba(0,0,0,0.3);
    overflow-y:auto; z-index:998;
}
#right-panel h2 { color:#f44336; font-size:16px; margin-bottom:8px; text-align:center; }
.metric-label { font-size:13px; margin-top:10px; }
.metric-value { font-size:16px; font-weight:700; }
.metric-delta { font-size:13px; }
</style>
<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Portfolio Analysis Platform with Real-Time Data</p>
</div>
""", unsafe_allow_html=True)

# --- Fetch Index Prices ---
@st.cache_data(ttl=300)
def fetch_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        newest, previous = hist.iloc[-1], hist.iloc[-2]
        diff = newest - previous
        pct = diff / previous * 100
        return f"${newest:,.2f}", f"{diff:+.2f} ({pct:+.2f}%)"
    return None, None

indices = {
    "S&P 500": fetch_price("^GSPC"),
    "Nasdaq": fetch_price("^IXIC"),
    "Dow Jones": fetch_price("^DJI")
}

# --- Render Right Panel ---
st.markdown('<div id="right-panel">', unsafe_allow_html=True)
st.markdown("<h2>Market Indices</h2>", unsafe_allow_html=True)
for name, data in indices.items():
    price, delta = data
    if price:
        color = "#4caf50" if delta.startswith("+") else "#f44336"
        st.markdown(f"""
            <div class="metric-label">{name}</div>
            <div class="metric-value">{price}</div>
            <div class="metric-delta" style="color:{color};">{delta}</div>
        """, unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# --- Load Investment Data ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# --- 1) Select Investment Types ---
st.subheader("1. Select Investment Types")
if "Category" not in df.columns:
    st.error("⚠️ 'Category' column is missing.")
    st.stop()
choices = sorted(df["Category"].dropna().unique())
selected = st.multiselect("", choices, default=choices)

# --- 2) Editable Investment Table ---
st.subheader("2. Investment Data (Editable)")
edited = st.data_editor(df[df["Category"].isin(selected)], use_container_width=True, num_rows="fixed")

# --- Safe mean helper ---
def safe_mean(df, col):
    return df[col].mean() if col in df and not df.empty else None

# --- 3) Portfolio Averages ---
st.subheader("3. Portfolio Averages")
fields = [
    ("Avg Return (%)", "Expected Return (%)"),
    ("Avg Risk (1–10)", "Risk Level (1-10)"),
    ("Avg Cap Rate (%)", "Cap Rate (%)"),
    ("Avg Liquidity", "Liquidity (1-10)"),
    ("Avg Volatility", "Volatility (1-10)"),
    ("Avg Fees (%)", "Fees (%)"),
    ("Avg Min Invest ($)", "Minimum Investment ($)")
]
cols = st.columns(len(fields))
for (label, colname), col_area in zip(fields, cols):
    m = safe_mean(edited, colname)
    val_str = f"{m:.2f}{'%' if '%' in label or 'Rate' in label else ('$'+str(int(m)) if 'Invest' in label else '')}" if m is not None else "N/A"
    col_area.metric(label, val_str)

# --- 4) Compact Visual Charts ---
st.subheader("4. Visual Insights")
chart_cols = st.columns(4)
# Reuse previous chart code with smaller fonts
def bar_plot(x, y):
    fig, ax = plt.subplots(figsize=(2.5,1.8))
    ax.bar(x, y, color="#f44336")
    ax.tick_params(axis="x", labelrotation=45, labelsize=6)
    ax.tick_params(axis="y", labelsize=6)
    fig.tight_layout()
    return fig

def scatter_plot(ax_x, ax_y, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(2.5,1.8))
    ax.scatter(ax_x, ax_y, c="red", alpha=0.6)
    ax.set_xlabel(xlabel, fontsize=6); ax.set_ylabel(ylabel, fontsize=6)
    ax.tick_params(labelsize=6)
    ax.grid(True, linestyle="--", alpha=0.3)
    fig.tight_layout()
    return fig

plots = [
    ("Expected Return (%)", bar_plot),
    (("Volatility (1-10)", "Liquidity (1-10)"), scatter_plot),
    (("Fees (%)", "Expected Return (%)"), scatter_plot),
    ("Risk Level (1-10)", None)
]

for area, cfg in zip(chart_cols, plots):
    if isinstance(cfg[0], tuple):
        x_c, y_c = cfg[0]
        area.markdown(f"**{y_c} vs {x_c}**")
        if x_c in edited and y_c in edited and not edited.empty:
            area.pyplot(cfg[1](edited[x_c], edited[y_c], x_c, y_c))
    else:
        colname = cfg[0]
        area.markdown(f"**{colname}**")
        if colname in edited and not edited.empty:
            if cfg[1]:
                area.pyplot(cfg[1](edited["Investment Name"], edited[colname]))
            else:
                fig, ax = plt.subplots(figsize=(2.5,1.8))
                ax.hist(edited[colname], bins=8, color="#f44336", alpha=0.7)
                ax.tick_params(labelsize=6); fig.tight_layout()
                area.pyplot(fig)

# --- 5) Bottom Constraints ---
st.subheader("5. Portfolio Constraints")
min_i = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].min()), step=1000) if "Minimum Investment ($)" in edited else 0
min_r = st.slider("Min Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()), step=0.1) if "Expected Return (%)" in edited else 0
max_r = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()), int(edited["Risk Level (1-10)"].max()), int(edited["Risk Level (1-10)"].max()), step=1) if "Risk Level (1-10)" in edited else 10
horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"], index=1)
hedge = st.checkbox("Inflation Hedge Only")

filtered = edited.copy()
if "Minimum Investment ($)" in filtered: filtered = filtered[filtered["Minimum Investment ($)"]>=min_i]
if "Expected Return (%)" in filtered: filtered = filtered[filtered["Expected Return (%)"]>=min_r]
if "Risk Level (1-10)" in filtered: filtered = filtered[filtered["Risk Level (1-10)"]<=max_r]
if hedge and "Inflation Hedge (Yes/No)" in filtered: filtered = filtered[filtered["Inflation Hedge (Yes/No)"]=="Yes"]

# --- 6 & 7) Filtered Table + Exports ---
st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=200)

st.subheader("7. Export Reports")
ex1, ex2 = st.columns(2)
with ex1:
    if st.button("📤 Export PowerPoint"):
        st.success("PPT export placeholder")
with ex2:
    if st.button("📥 Export Word"):
        st.success("Word export placeholder")

