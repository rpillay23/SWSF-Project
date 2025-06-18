import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# --- PAGE CONFIG ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- STYLE ---
st.markdown("""
<style>
/* Header styling */
header > div {
    background-color: #111 !important;
    color: white !important;
}

/* Main content area padding */
.css-1d391kg {
    padding-top: 110px !important;
}

/* Header bar */
.app-header {
    background-color: #111;
    color: white;
    padding: 10px 25px;
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    z-index: 9999;
    box-sizing: border-box;
}
.app-header h1 {
    margin: 0;
    font-size: 26px;
    font-weight: 700;
}
.app-header p {
    margin: 4px 0 0 0;
    font-size: 13px;
    color: #f44336;
    font-weight: 500;
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

/* Metrics font size */
.stMetric > div {
    font-size: 14px !important;
}
.stMetric > div > div:first-child {
    font-weight: 700 !important;
}

/* Smaller font for dataframes */
[data-testid="stDataFrame"] {
    font-size: 13px;
}

/* Compact graph containers */
.graph-container {
    padding: 0 10px;
}
</style>

<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization and Financial Advisory.<br>
    Designed for portfolio managers and finance professionals to generate, analyze, and optimize investment portfolios using real-time data.</p>
</div>
""", unsafe_allow_html=True)

# --- LOAD DATA ---

try:
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = [col.strip() for col in df.columns]
except Exception as e:
    st.error(f"Error loading data file: {e}")
    st.stop()

# --- MAIN CONTENT ---

st.header("Investment Generator")

cat_col = "Category"
min_inv_col = "Minimum Investment ($)"
exp_ret_col = "Expected Return (%)"
risk_col = "Risk Level (1-10)"
investment_col = "Investment"  # fallback if not found in df

if investment_col not in df.columns:
    for col in df.columns:
        if df[col].dtype == object:
            investment_col = col
            break

categories = sorted(df[cat_col].dropna().unique())
selected_categories = st.multiselect("Select Investment Categories:", categories, default=categories)

min_inv_default = int(df[min_inv_col].min())
max_inv_default = int(df[min_inv_col].max())
min_inv = st.slider("Minimum Investment ($)", min_value=min_inv_default, max_value=max_inv_default, value=min_inv_default, step=1000)

exp_ret_min = float(df[exp_ret_col].min())
exp_ret_max = float(df[exp_ret_col].max())
exp_ret = st.slider("Minimum Expected Return (%)", min_value=exp_ret_min, max_value=exp_ret_max, value=exp_ret_min, step=0.1)

risk_min = int(df[risk_col].min())
risk_max = int(df[risk_col].max())
max_risk = st.slider("Maximum Risk Level (1-10)", min_value=risk_min, max_value=risk_max, value=risk_max, step=1)

filtered_df = df[
    (df[cat_col].isin(selected_categories)) &
    (df[min_inv_col] >= min_inv) &
    (df[exp_ret_col] >= exp_ret) &
    (df[risk_col] <= max_risk)
]

st.write(f"### {len(filtered_df)} Investments Matching Your Criteria")
st.dataframe(filtered_df.reset_index(drop=True), height=250)

st.divider()

# --- PORTFOLIO SUMMARY METRICS ---

st.header("Portfolio Summary Metrics")
c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("Avg Return (%)", f"{filtered_df[exp_ret_col].mean():.2f}" if not filtered_df.empty else "N/A")
c2.metric("Avg Risk (1–10)", f"{filtered_df[risk_col].mean():.2f}" if not filtered_df.empty else "N/A")
c3.metric("Avg Cap Rate (%)", f"{filtered_df['Cap Rate (%)'].mean():.2f}" if 'Cap Rate (%)' in filtered_df.columns and not filtered_df.empty else "N/A")
c4.metric("Avg Liquidity", f"{filtered_df['Liquidity (1–10)'].mean():.2f}" if 'Liquidity (1–10)' in filtered_df.columns and not filtered_df.empty else "N/A")
c5.metric("Avg Volatility", f"{filtered_df['Volatility (1–10)'].mean():.2f}" if 'Volatility (1–10)' in filtered_df.columns and not filtered_df.empty else "N/A")
c6.metric("Avg Fees (%)", f"{filtered_df['Fees (%)'].mean():.2f}" if 'Fees (%)' in filtered_df.columns and not filtered_df.empty else "N/A")
c7.metric("Avg Min Investment", f"${filtered_df[min_inv_col].mean():,.0f}" if not filtered_df.empty else "N/A")

st.divider()

# --- VISUAL INSIGHTS ---

st.header("Visual Insights")
ch1, ch2, ch3, ch4 = st.columns(4)

with ch1:
    st.markdown("**Expected Return (%)**")
    fig, ax = plt.subplots(figsize=(3, 2))
    if not filtered_df.empty:
        ax.bar(filtered_df[investment_col], filtered_df[exp_ret_col], color='#f44336')
        ax.set_xticklabels(filtered_df[investment_col], rotation=45, ha='right', fontsize=7)
    ax.set_ylabel("Return %", fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

with ch2:
    st.markdown("**Volatility vs Liquidity**")
    fig, ax = plt.subplots(figsize=(3, 2))
    if not filtered_df.empty and 'Volatility (1–10)' in filtered_df.columns and 'Liquidity (1–10)' in filtered_df.columns:
        ax.scatter(filtered_df['Volatility (1–10)'], filtered_df['Liquidity (1–10)'], color='red', alpha=0.7)
    ax.set_xlabel("Volatility", fontsize=8)
    ax.set_ylabel("Liquidity", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    st.pyplot(fig)

with ch3:
    st.markdown("**Fees vs Expected Return**")
    fig, ax = plt.subplots(figsize=(3, 2))
    if not filtered_df.empty and 'Fees (%)' in filtered_df.columns:
        ax.scatter(filtered_df['Fees (%)'], filtered_df[exp_ret_col], color='red', alpha=0.7)
    ax.set_xlabel("Fees %", fontsize=8)
    ax.set_ylabel("Expected Return %", fontsize=8)
    ax.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    st.pyplot(fig)

with ch4:
    st.markdown("**Risk Level Distribution**")
    fig, ax = plt.subplots(figsize=(3, 2))
    if not filtered_df.empty:
        ax.hist(filtered_df[risk_col], bins=10, color='#f44336', alpha=0.8)
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
