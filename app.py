import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import streamlit.components.v1 as components

# --- Page Config & Styling ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
.app-header {
    background-color: #111;
    color: white;
    padding: 8px 20px;
    position: fixed;
    top: 0; left: 0; width: 100%;
    z-index: 1000;
}
.app-header h1 { margin: 0; font-size: 20px; }
.app-header p { margin: 0; font-size: 12px; color: #f44336; }
main, .block-container {
    padding-top: 60px !important;
    margin: 0 200px 20px 20px; /* left, right margins */
}
.stButton > button {
    background: #111; color: #f44336;
    border: 2px solid #f44336;
    border-radius: 4px;
    padding: 0.3em 0.8em;
    font-weight: 700; font-size: 12px;
}
.stButton > button:hover { background: #f44336; color: #111; }
[data-testid="stDataFrame"] { font-size: 12px; }
</style>
<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Portfolio Analysis Platform with Real‑Time Data</p>
</div>
""", unsafe_allow_html=True)

# --- Market Indices Fetch & Display (Right Edge) ---
@st.cache_data(ttl=300)
def get_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        cur, prev = hist.iloc[-1], hist.iloc[-2]
        diff = cur - prev
        pct = diff / prev * 100
        return f"${cur:,.2f}", f"{diff:+.2f} ({pct:+.2f}%)"
    return None, None

indices = {
    "S&P 500": get_price("^GSPC"),
    "Nasdaq": get_price("^IXIC"),
    "Dow Jones": get_price("^DJI")
}

html = """
<div id="market-box">
  <h2>📈 Market Indices</h2>
"""
for name, (price, delta) in indices.items():
    if price:
        color = "#4caf50" if delta.startswith("+") else "#f44336"
        html += f"""
  <div class="metric-label">{name}</div>
  <div class="metric-value">{price}</div>
  <div class="metric-delta" style="color:{color};">{delta}</div>
"""
html += """
</div>
<style>
#market-box {
    position: fixed;
    top: 60px;
    right: 0;
    width: 180px;
    background: #111;
    color: #fff;
    padding: 15px;
    border-left: 2px solid #f44336;
    z-index: 999;
    height: calc(100vh - 60px);
    overflow-y: auto;
    font-family: system-ui, sans-serif;
}
#market-box h2 {
    text-align: center;
    color: #f44336;
    margin-bottom: 8px;
    font-size: 16px;
}
.metric-label { font-size: 13px; margin-top: 8px; }
.metric-value { font-size: 15px; font-weight: 700; }
.metric-delta { font-size: 13px; }
</style>
"""
components.html(html, height=600)

# --- Load Data ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# 1. Select Investment Types
st.subheader("1. Select Investment Types")
if "Category" not in df.columns:
    st.error("⚠️ Missing 'Category' column.")
    st.stop()
choices = sorted(df["Category"].dropna().unique())
selected = st.multiselect("", choices, default=choices)

# 2. Editable Investment Table
st.subheader("2. Editable Investment Data")
editable_df = st.data_editor(df[df["Category"].isin(selected)], num_rows="fixed", use_container_width=True)

# Helper for averages
def safe_mean(df, col):
    return df[col].mean() if col in df.columns and not df.empty else None

# 3. Portfolio Averages
st.subheader("3. Portfolio Averages")
metrics = [
    ("Expected Return (%)", "Avg Return (%)"),
    ("Risk Level (1-10)", "Avg Risk"),
    ("Cap Rate (%)", "Avg Cap Rate"),
    ("Liquidity (1-10)", "Avg Liquidity"),
    ("Volatility (1-10)", "Avg Volatility"),
    ("Fees (%)", "Avg Fees"),
    ("Minimum Investment ($)", "Avg Min Invest")
]
cols = st.columns(len(metrics))
for (col_name, label), slot in zip(metrics, cols):
    m = safe_mean(editable_df, col_name)
    display = f"${m:,.0f}" if col_name == "Minimum Investment ($)" else (f"{m:.2f}%" if m is not None else "N/A")
    slot.metric(label, display)

# 4. Visual Charts
st.subheader("4. Visual Insights")
chart_areas = st.columns(4)
# Bar plot
if "Expected Return (%)" in editable_df:
    fig, ax = plt.subplots(figsize=(2.5,2))
    ax.bar(editable_df["Investment Name"], editable_df["Expected Return (%)"], color="#f44336")
    ax.tick_params(axis="x", rotation=45, labelsize=6)
    ax.tick_params(labelsize=6)
    fig.tight_layout()
    chart_areas[0].pyplot(fig)
# Scatter volatility vs liquidity
if {"Volatility (1-10)", "Liquidity (1-10)"}.issubset(editable_df.columns):
    fig, ax = plt.subplots(figsize=(2.5,2))
    ax.scatter(editable_df["Volatility (1-10)"], editable_df["Liquidity (1-10)"], color="red", alpha=0.6)
    ax.set_xlabel("Volatility", fontsize=6)
    ax.set_ylabel("Liquidity", fontsize=6)
    ax.tick_params(labelsize=6)
    ax.grid(linestyle="--", alpha=0.3)
    fig.tight_layout()
    chart_areas[1].pyplot(fig)
# Fees vs return scatter
if {"Fees (%)", "Expected Return (%)"}.issubset(editable_df.columns):
    fig, ax = plt.subplots(figsize=(2.5,2))
    ax.scatter(editable_df["Fees (%)"], editable_df["Expected Return (%)"], color="red", alpha=0.6)
    ax.set_xlabel("Fees", fontsize=6)
    ax.set_ylabel("Return", fontsize=6)
    ax.tick_params(labelsize=6)
    ax.grid(linestyle="--", alpha=0.3)
    fig.tight_layout()
    chart_areas[2].pyplot(fig)
# Risk level histogram
if "Risk Level (1-10)" in editable_df:
    fig, ax = plt.subplots(figsize=(2.5,2))
    ax.hist(editable_df["Risk Level (1-10)"], bins=8, color="#f44336", alpha=0.7)
    ax.set_xlabel("Risk", fontsize=6)
    ax.set_ylabel("Count", fontsize=6)
    ax.tick_params(labelsize=6)
    fig.tight_layout()
    chart_areas[3].pyplot(fig)

# 5. Portfolio Constraints
st.subheader("5. Portfolio Constraints")
min_i = st.slider("Min Investment ($)", int(editable_df["Minimum Investment ($)"].min()), int(editable_df["Minimum Investment ($)"].max()), step=1000) if "Minimum Investment ($)" in editable_df else 0
min_r = st.slider("Min Return (%)", float(editable_df["Expected Return (%)"].min()), float(editable_df["Expected Return (%)"].max()), step=0.1) if "Expected Return (%)" in editable_df else 0
max_r = st.slider("Max Risk Level", int(editable_df["Risk Level (1-10)"].min()), int(editable_df["Risk Level (1-10)"].max()), step=1) if "Risk Level (1-10)" in editable_df else 10
horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"], index=1)
hedge = st.checkbox("Inflation Hedge Only")

filtered = editable_df.copy()
if "Minimum Investment ($)" in filtered:
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_i]
if "Expected Return (%)" in filtered:
    filtered = filtered[filtered["Expected Return (%)"] >= min_r]
if "Risk Level (1-10)" in filtered:
    filtered = filtered[filtered["Risk Level (1-10)"] <= max_r]
if hedge and "Inflation Hedge (Yes/No)" in filtered:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"}

# 6. Filtered Investments
st.subheader(f"6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=250)

# 7. Export Reports
st.subheader("7. Export Reports")
exp1, exp2 = st.columns(2)
with exp1:
    if st.button("📤 Export PowerPoint"):
        st.success("PowerPoint export placeholder")
with exp2:
    if st.button("📥 Export Word"):
        st.success("Word export placeholder")

