import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf

# --- PAGE CONFIG ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- CSS for layout and theme ---
st.markdown("""
<style>
/* Reset and base */
html, body, [class*="css"] {
    font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    background-color: white;
    color: #111;
    margin: 0; padding: 0;
}

/* HEADER fixed full width */
.app-header {
    background-color: #111;
    color: white;
    padding: 10px 20px;
    width: 100vw;
    position: fixed;
    top: 0; left: 0;
    z-index: 9999;
    box-sizing: border-box;
}
.app-header h1 {
    margin: 0;
    font-size: 24px;
    font-weight: 700;
}
.app-header p {
    margin: 4px 0 0 0;
    font-size: 12px;
    color: #f44336;
    font-weight: 500;
}

/* LEFT SIDEBAR fixed */
#market-indices {
    position: fixed;
    top: 80px; /* below header */
    left: 0;
    width: 280px;
    height: calc(100vh - 80px);
    background-color: #111;
    color: white;
    overflow-y: auto;
    padding: 15px 15px 30px 15px;
    border-right: 3px solid #f44336;
    font-size: 13px;
    box-sizing: border-box;
    z-index: 9998;
}
#market-indices h2 {
    color: #f44336;
    font-weight: 700;
    margin-bottom: 15px;
    font-size: 20px;
}
#market-indices .index-name {
    font-weight: 700;
    margin-top: 15px;
    font-size: 15px;
}
#market-indices .metric {
    font-weight: 700;
    margin-left: 5px;
    color: #f44336;
}
#market-indices .delta-positive {
    color: #4caf50;
    font-weight: 700;
}
#market-indices .delta-negative {
    color: #f44336;
    font-weight: 700;
}

/* Make mini charts fit container */
#market-indices .mini-chart {
    width: 100% !important;
    height: 90px !important;
    margin-top: 8px;
    margin-bottom: 12px;
}

/* MAIN CONTENT area */
#main-content {
    margin-left: 280px;
    margin-top: 110px; /* below header */
    padding: 15px 25px 40px 25px;
    max-width: calc(100vw - 280px);
    box-sizing: border-box;
}

/* Smaller fonts for metrics and data */
.stMetric > div {
    font-size: 14px !important;
}
.stMetric > div > div:first-child {
    font-weight: 700 !important;
}

/* Buttons */
.stButton > button {
    background-color: #111;
    color: #f44336;
    border-radius: 6px;
    padding: 0.4em 1em;
    border: 2px solid #f44336;
    font-weight: 700;
    font-size: 14px;
    transition: background-color 0.3s ease, color 0.3s ease;
}
.stButton > button:hover {
    background-color: #f44336;
    color: #111;
    border: 2px solid #f44336;
}

/* Visual Insights charts */
.visual-chart {
    width: 100%;
    max-width: 300px;
    height: 210px;
}
</style>

<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization and Financial Advisory.<br>
    Designed for portfolio managers and finance professionals to generate, analyze, and optimize investment portfolios using real-time data.</p>
</div>
""", unsafe_allow_html=True)


# --- FUNCTIONS ---

@st.cache_data(ttl=600)
def get_index_data(ticker):
    try:
        index = yf.Ticker(ticker)
        hist = index.history(period="1mo")
        hist.reset_index(inplace=True)
        return hist
    except Exception:
        return pd.DataFrame()

def format_market_metric(latest, prev):
    latest_close = latest['Close']
    prev_close = prev['Close']
    change = latest_close - prev_close
    pct_change = (change / prev_close) * 100
    delta_class = "delta-positive" if change >= 0 else "delta-negative"
    return latest_close, change, pct_change, delta_class

def plot_mini_line_chart(data):
    fig, ax = plt.subplots(figsize=(3, 1.2))
    ax.plot(data['Date'], data['Close'], color='#f44336', linewidth=1.7)
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)

def sanitize_string(s):
    if isinstance(s, str):
        return (
            s.replace("–", "-")
             .replace("’", "'")
             .replace("“", '"')
             .replace("”", '"')
             .replace("•", "-")
             .replace("©", "(c)")
        )
    return s

# --- LOAD DATA ---

try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df = df.applymap(sanitize_string)
except Exception as e:
    st.error(f"Error loading data file: {e}")
    st.stop()

# --- LEFT PERMANENT MARKET INDICES ---

st.markdown('<div id="market-indices">', unsafe_allow_html=True)
st.markdown("<h2>Real-Time Market Indices</h2>")

for ticker, name in [("^GSPC", "S&P 500"), ("^IXIC", "Nasdaq"), ("^DJI", "Dow Jones")]:
    data = get_index_data(ticker)
    if data.empty:
        st.markdown(f"<div class='index-name'>{name}</div><div>Data unavailable</div>", unsafe_allow_html=True)
        continue

    latest = data.iloc[-1]
    prev = data.iloc[-2]
    latest_close, change, pct_change, delta_class = format_market_metric(latest, prev)

    st.markdown(f"<div class='index-name'>{name}</div>", unsafe_allow_html=True)
    st.markdown(f"<span class='metric'>${latest_close:,.2f}</span> <span class='{delta_class}'>{change:+.2f} ({pct_change:+.2f}%)</span>", unsafe_allow_html=True)
    
    # plot mini chart
    plot_mini_line_chart(data)

st.markdown('</div>', unsafe_allow_html=True)

# --- MAIN CONTENT ---

st.markdown('<div id="main-content">', unsafe_allow_html=True)

# --- INVESTMENT GENERATOR ---

st.header("Investment Generator")

categories = sorted(df['Category'].dropna().unique())
selected_categories = st.multiselect("Select Investment Categories:", categories, default=categories)

min_inv_default = int(df['Minimum Investment ($)'].min())
max_inv_default = int(df['Minimum Investment ($)'].max())
min_inv = st.slider("Minimum Investment ($)", min_value=min_inv_default, max_value=max_inv_default, value=min_inv_default, step=1000)

exp_ret_min = float(df['Expected Return (%)'].min())
exp_ret_max = float(df['Expected Return (%)'].max())
exp_ret = st.slider("Minimum Expected Return (%)", min_value=exp_ret_min, max_value=exp_ret_max, value=exp_ret_min, step=0.1)

risk_min = int(df['Risk Level (1-10)'].min())
risk_max = int(df['Risk Level (1-10)'].max())
max_risk = st.slider("Maximum Risk Level (1-10)", min_value=risk_min, max_value=risk_max, value=risk_max, step=1)

filtered_df = df[
    (df['Category'].isin(selected_categories)) &
    (df['Minimum Investment ($)'] >= min_inv) &
    (df['Expected Return (%)'] >= exp_ret) &
    (df['Risk Level (1-10)'] <= max_risk)
]

st.write(f"### {len(filtered_df)} Investments Matching Your Criteria")
st.dataframe(filtered_df.reset_index(drop=True), height=250)

st.divider()

# --- PORTFOLIO SUMMARY METRICS ---

st.header("Portfolio Summary Metrics")
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("Avg Return (%)", f"{filtered_df['Expected Return (%)'].mean():.2f}" if not filtered_df.empty else "N/A")
c2.metric("Avg Risk (1–10)", f"{filtered_df['Risk Level (1-10)'].mean():.2f}" if not filtered_df.empty else "N/A")
c3.metric("Avg Cap Rate (%)", f"{filtered_df['Cap Rate (%)'].mean():.2f}" if not filtered_df.empty else "N/A")
c4.metric("Avg Liquidity", f"{filtered_df['Liquidity (1–10)'].mean():.2f}" if not filtered_df.empty else "N/A")
c5.metric("Avg Volatility", f"{filtered_df['Volatility (1–10)'].mean():.2f}" if not filtered_df.empty else "N/A")
c6.metric("Avg Fees (%)", f"{filtered_df['Fees (%)'].mean():.2f}" if not filtered_df.empty else "N/A")
c7.metric("Avg Min Investment", f"${filtered_df['Minimum Investment ($)'].mean():,.0f}" if not filtered_df.empty else "N/A")

st.divider()

# --- VISUAL INSIGHTS ---

st.header("Visual Insights")
ch1, ch2, ch3, ch4 = st.columns(4)

with ch1:
    st.markdown("**Expected Return (%)**")
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.bar(filtered_df['Investment'], filtered_df['Expected Return (%)'], color='#f44336')
    ax.set_xticklabels(filtered_df['Investment'], rotation=45, ha='right', fontsize=7)
    ax.set_ylabel("Return %", fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

with ch2:
    st.markdown("**Volatility vs Liquidity**")
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.scatter(filtered_df['Volatility (1–10)'], filtered_df['Liquidity (1–10)'], color='red', alpha=0.7)
    ax.set_xlabel("Volatility", fontsize=8)
    ax.set_ylabel("Liquidity", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    st.pyplot(fig)

with ch3:
    st.markdown("**Fees vs Expected Return**")
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.scatter(filtered_df['Fees (%)'], filtered_df['Expected Return (%)'], color='red', alpha=0.7)
    ax.set_xlabel("Fees %", fontsize=8)
    ax.set_ylabel("Expected Return %", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    st.pyplot(fig)

with ch4:
    st.markdown("**Risk Level Distribution**")
    fig, ax = plt.subplots(figsize=(3, 2))
    ax.hist(filtered_df['Risk Level (1-10)'], bins=10, color='#f44336', alpha=0.8)
    ax.set_xlabel("Risk Level", fontsize=8)
    ax.set_ylabel("Frequency", fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

st.divider()

# --- EXPORT REPORTS ---

st.header("Export Reports")
col1, col2 = st.columns(2)

with col1:
    if st.button("Download PowerPoint Report"):
        st.success("PowerPoint report generated (functionality to implement).")

with col2:
    if st.button("Download Word Report"):
        st.success("Word report generated (functionality to implement).")

st.markdown('</div>', unsafe_allow_html=True)
