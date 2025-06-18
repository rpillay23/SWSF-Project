import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# === Page Config & Styling ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# === Custom Styling & Right Fixed Box ===
st.markdown("""
<style>
/* Main header */
.app-header {
    background-color: #111;
    color: white;
    padding: 10px 20px;
    position: fixed;
    top: 0; left: 0; width: 100%;
    z-index: 1000;
}
.app-header h1 {
    margin: 0;
    font-size: 20px;
}
.app-header p {
    margin: 0;
    font-size: 12px;
    color: #f44336;
}

/* Right market index panel */
#market-box {
    position: fixed;
    top: 70px;
    right: 0;
    width: 190px;
    background-color: #111;
    color: white;
    padding: 15px;
    border-left: 2px solid #f44336;
    z-index: 999;
    height: calc(100vh - 70px);
    overflow-y: auto;
}
.market-title {
    color: #f44336;
    font-weight: bold;
    font-size: 15px;
    text-align: center;
    margin-bottom: 10px;
}
.metric-label {
    font-size: 13px;
    margin-top: 10px;
}
.metric-value {
    font-size: 15px;
    font-weight: bold;
}
.metric-delta {
    font-size: 13px;
}
</style>

<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Portfolio Analysis Platform with Real‑Time Data</p>
</div>

<div id="market-box">
  <div class="market-title">📈 Market Indices</div>
""", unsafe_allow_html=True)

# === Fetch Index Data ===
@st.cache_data(ttl=300)
def fetch_index(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        cur, prev = hist.iloc[-1], hist.iloc[-2]
        diff = cur - prev
        pct = diff / prev * 100
        return f"${cur:,.2f}", f"{diff:+.2f} ({pct:+.2f}%)"
    return None, None

# Indices list
indices = {
    "S&P 500": fetch_index("^GSPC"),
    "Nasdaq": fetch_index("^IXIC"),
    "Dow Jones": fetch_index("^DJI")
}

# === Display in right-hand HTML box ===
for name, (price, delta) in indices.items():
    if price:
        color = "#4caf50" if delta.startswith("+") else "#f44336"
        st.markdown(f"""
        <div class="metric-label">{name}</div>
        <div class="metric-value">{price}</div>
        <div class="metric-delta" style="color:{color};">{delta}</div>
        """, unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)  # close market-box div

# === Load Investment Data ===
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# === App Sections ===
st.markdown("### 1. Select Investment Types")
all_cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("Categories", all_cats, default=all_cats)

st.markdown("### 2. Editable Investment Table")
editable_df = st.data_editor(df[df["Category"].isin(sel_cats)], use_container_width=True, num_rows="fixed")

# === Averages Section ===
st.markdown("### 3. Portfolio Averages")
metrics = {
    "Expected Return (%)": "Avg Return (%)",
    "Risk Level (1-10)": "Avg Risk",
    "Cap Rate (%)": "Avg Cap Rate",
    "Liquidity (1-10)": "Avg Liquidity",
    "Volatility (1-10)": "Avg Volatility",
    "Fees (%)": "Avg Fees",
    "Minimum Investment ($)": "Avg Minimum Invest"
}
avg_cols = st.columns(len(metrics))
for col, label in zip(metrics, avg_cols):
    if col in editable_df:
        val = editable_df[col].mean()
        if "Investment" in col:
            display = f"${val:,.0f}"
        else:
            display = f"{val:.2f}%"
        label.metric(metrics[col], display)

# === Visual Charts Section ===
st.markdown("### 4. Visual Charts")
charts = st.columns(4)

# Bar plot
if "Expected Return (%)" in editable_df:
    fig, ax = plt.subplots(figsize=(2.5,2))
    ax.bar(editable_df["Investment Name"], editable_df["Expected Return (%)"], color="red")
    ax.set_xticklabels(editable_df["Investment Name"], rotation=45, ha="right", fontsize=6)
    ax.set_ylabel("Expected Return (%)")
    charts[0].pyplot(fig)

# Scatter 1: Volatility vs Liquidity
if all(k in editable_df for k in ["Volatility (1-10)", "Liquidity (1-10)"]):
    fig2, ax2 = plt.subplots(figsize=(2.5,2))
    ax2.scatter(editable_df["Volatility (1-10)"], editable_df["Liquidity (1-10)"], color="red")
    ax2.set_xlabel("Volatility (1-10)", fontsize=6)
    ax2.set_ylabel("Liquidity (1-10)", fontsize=6)
    ax2.tick_params(labelsize=6)
    charts[1].pyplot(fig2)

# Scatter 2: Fees vs Return
if all(k in editable_df for k in ["Fees (%)", "Expected Return (%)"]):
    fig3, ax3 = plt.subplots(figsize=(2.5,2))
    ax3.scatter(editable_df["Fees (%)"], editable_df["Expected Return (%)"], color="red")
    ax3.set_xlabel("Fees (%)", fontsize=6)
    ax3.set_ylabel("Expected Return (%)", fontsize=6)
    ax3.tick_params(labelsize=6)
    charts[2].pyplot(fig3)

# Risk Histogram
if "Risk Level (1-10)" in editable_df:
    fig4, ax4 = plt.subplots(figsize=(2.5,2))
    ax4.hist(editable_df["Risk Level (1-10)"], bins=10, color="red", alpha=0.7)
    ax4.set_xlabel("Risk Level (1-10)", fontsize=6)
    ax4.set_ylabel("Count", fontsize=6)
    ax4.tick_params(labelsize=6)
    charts[3].pyplot(fig4)

# === Filters Section ===
st.markdown("### 5. Portfolio Constraints")
if "Minimum Investment ($)" in editable_df:
    min_inv = st.slider("Minimum Investment", int(editable_df["Minimum Investment ($)"].min()), int(editable_df["Minimum Investment ($)"].max()), step=1000)
else:
    min_inv = 0

if "Expected Return (%)" in editable_df:
    min_ret = st.slider("Minimum Expected Return (%)", float(editable_df["Expected Return (%)"].min()), float(editable_df["Expected Return (%)"].max()), step=0.1)
else:
    min_ret = 0.0

if "Risk Level (1-10)" in editable_df:
    max_risk = st.slider("Maximum Risk Level (1-10)", 1, 10, 10)
else:
    max_risk = 10

horizon = st.selectbox("Time Horizon", ["Short", "Medium", "Long"])
inflation = st.checkbox("Hedge Against Inflation")

# === Filtered Data Output ===
filtered = editable_df.copy()
if "Minimum Investment ($)" in filtered:
    filtered = filtered[filtered["Minimum Investment ($)"] >= min_inv]
if "Expected Return (%)" in filtered:
    filtered = filtered[filtered["Expected Return (%)"] >= min_ret]
if "Risk Level (1-10)" in filtered:
    filtered = filtered[filtered["Risk Level (1-10)"] <= max_risk]
if inflation and "Inflation Hedge (Yes/No)" in filtered:
    filtered = filtered[filtered["Inflation Hedge (Yes/No)"] == "Yes"]

st.markdown(f"### 6. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=250)

# === Export Buttons ===
st.markdown("### 7. Export Reports")
exp1, exp2 = st.columns(2)
with exp1:
    if st.button("📤 Export PowerPoint"):
        st.success("Exported PowerPoint (placeholder)")
with exp2:
    if st.button("📥 Export Word"):
        st.success("Exported Word Report (placeholder)")
