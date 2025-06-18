import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- Page Setup & Styling ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

st.markdown("""
<style>
/* Header Styling */
.app-header {
    background:#111; 
    color:white; 
    padding:10px 25px; 
    position:fixed;
    top:0;
    left:0;
    width:100%;
    z-index:1000;
    box-shadow: 0 2px 6px rgba(0,0,0,0.4);
}
.app-header h1 {
    margin:0;
    font-size:20px;
    font-weight:700;
}
.app-header p {
    margin:3px 0 0;
    color:#f44336;
    font-size:12px;
}

/* Index bar styling */
.index-bar {
    display:flex;
    justify-content:center;
    gap:30px;
    margin-top:70px;
    margin-bottom:-10px;
}
.index-box {
    background:#111;
    color:white;
    padding:10px 20px;
    border:1px solid #f44336;
    border-radius:5px;
    text-align:center;
    font-size:13px;
}
.index-delta-positive { color: #4caf50; }
.index-delta-negative { color: #f44336; }

[data-testid="stDataFrame"] { font-size:13px; }

</style>

<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Portfolio Analysis Platform with Real‑Time Data</p>
</div>
""", unsafe_allow_html=True)

# --- Fetch Market Indices ---
@st.cache_data(ttl=300)
def get_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")["Close"]
    if len(hist) >= 2:
        last, prev = hist.iloc[-1], hist.iloc[-2]
        delta = last - prev
        pct = delta / prev * 100
        return f"${last:,.2f}", f"{delta:+.2f} ({pct:+.2f}%)"
    return "N/A", ""

tickers = {"S&P 500": "^GSPC", "Nasdaq": "^IXIC", "Dow Jones": "^DJI"}
prices = {name: get_price(ticker) for name, ticker in tickers.items()}

# --- Show Market Indices Bar ---
index_html = '<div class="index-bar">'
for name, (price, delta) in prices.items():
    color_class = "index-delta-positive" if delta.startswith("+") else "index-delta-negative"
    index_html += f"""
    <div class="index-box">
        <div><strong>{name}</strong></div>
        <div style="font-size:16px;">{price}</div>
        <div class="{color_class}">{delta}</div>
    </div>
    """
index_html += "</div>"
st.markdown(index_html, unsafe_allow_html=True)

# --- Load Data ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# --- 1. Select Investment Types ---
st.subheader("1. Select Investment Types")
cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("Choose categories to include in portfolio:", cats, default=cats)

# --- 2. Editable Table ---
st.subheader("2. Editable Investment Data")
filtered_df = df[df["Category"].isin(sel_cats)].copy()
edited = st.data_editor(filtered_df, use_container_width=True, num_rows="fixed")

# --- 3. Portfolio Averages ---
st.subheader("3. Portfolio Averages")
def mean_safe(df, col): return f"{df[col].mean():.2f}" if col in df.columns and not df.empty else "N/A"

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
for (label, col), c in zip(fields, cols):
    val = mean_safe(edited, col)
    suffix = "%" if "%" in label or "Rate" in label else "$" if "$" in label else ""
    c.metric(label, f"{val}{suffix}" if val != "N/A" else val)

# --- 4. Visual Insights ---
st.subheader("4. Visual Insights")
chart_cols = st.columns(4)

def chart_bar(x, y):
    fig, ax = plt.subplots(figsize=(2.6, 1.8))
    ax.bar(x, y, color="#f44336")
    ax.tick_params(axis="x", rotation=45, labelsize=6)
    fig.tight_layout()
    return fig

def chart_scatter(x, y, xlabel, ylabel):
    fig, ax = plt.subplots(figsize=(2.6, 1.8))
    ax.scatter(x, y, color="red", alpha=0.6)
    ax.set_xlabel(xlabel, fontsize=7)
    ax.set_ylabel(ylabel, fontsize=7)
    ax.tick_params(labelsize=6)
    fig.tight_layout()
    return fig

visuals = [
    ("Expected Return (%)", chart_bar),
    (("Volatility (1-10)", "Liquidity (1-10)"), chart_scatter),
    (("Fees (%)", "Expected Return (%)"), chart_scatter),
    ("Risk Level (1-10)", None)
]

for slot, config in zip(chart_cols, visuals):
    if isinstance(config[0], tuple):
        x, y = config[0]
        if x in edited and y in edited and not edited.empty:
            slot.markdown(f"**{y} vs {x}**")
            slot.pyplot(config[1](edited[x], edited[y], x, y))
    else:
        col = config[0]
        slot.markdown(f"**{col}**")
        if col in edited and not edited.empty:
            if config[1]:
                slot.pyplot(config[1](edited["Investment Name"], edited[col]))
            else:
                fig, ax = plt.subplots(figsize=(2.6, 1.8))
                ax.hist(edited[col], bins=7, color="#f44336", alpha=0.7)
                ax.tick_params(labelsize=6)
                fig.tight_layout()
                slot.pyplot(fig)

# --- 5. Portfolio Constraints ---
st.subheader("5. Filter Your Portfolio")
min_inv = st.slider("Minimum Investment ($)", 0, int(edited["Minimum Investment ($)"].max()), 0, step=1000)
min_ret = st.slider("Minimum Expected Return (%)", 0.0, float(edited["Expected Return (%)"].max()), 0.0, step=0.1)
max_risk = st.slider("Maximum Risk Level", 0, 10, 10)
time_horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"])
hedge = st.checkbox("Inflation Hedge Only")

filtered = edited.copy()
if "Minimum Investment ($)" in filtered.columns: filtered = filtered[filtered["Minimum Investment ($)"] >= min_inv]
if "Expected Return (%)" in filtered.columns: filtered = filtered[filtered["Expected Return (%)"] >= min_ret]
if "Risk Level (1-10)" in filtered.columns: filtered = filtered[filtered["Risk Level (1-10)"] <= max_risk]
if hedge and "Inflation Hedge (Yes/No)" in filtered.columns:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"]

# --- 6. Show Filtered Table ---
st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=220)

# --- 7. Export Options ---
st.subheader("7. Export Reports")
b1, b2 = st.columns(2)
with b1:
    if st.button("📤 Download PowerPoint"):
        st.success("PPT export placeholder")
with b2:
    if st.button("📥 Download Word"):
        st.success("Word export placeholder")
