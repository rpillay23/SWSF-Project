import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# ==== Top Header Bar ====
st.markdown("""
    <div style="
        background-color: #111; 
        color: white; 
        padding: 15px 30px; 
        border-bottom: 3px solid #f44336;
        margin: 0 -30px 25px -30px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    ">
        <h1 style="margin: 0; font-weight: 700;">Automated Investment Matrix</h1>
        <p style="margin: 5px 0 0 0; font-size: 18px; color: #f44336;">
            Portfolio Analysis Platform with Real-Time Data for Portfolio Optimization and Financial Advisory.
        </p>
    </div>
""", unsafe_allow_html=True)

# ==== Sample Investment Data ====
data = {
    "Investment": [
        "Bonds", "Commodities", "Direct Lending", "Equities", 
        "Infrastructure", "Life Settlements", "Real Estate"
    ],
    "Risk (%)": [3.5, 7, 4.5, 10, 6, 2, 5],
    "Expected Return (%)": [4.0, 8, 5, 12, 7, 3, 6],
    "Volatility (%)": [2, 6, 3, 9, 5, 1, 4],
    "Liquidity (1-10)": [9, 5, 6, 8, 7, 3, 4],
    "Inflation Hedge (Yes/No)": ["No", "Yes", "No", "No", "Yes", "No", "Yes"],
    "Minimum Investment ($k)": [10, 20, 15, 5, 25, 30, 20]
}
df = pd.DataFrame(data)

# ==== Investment Type Selector ====
st.markdown("### Select Investment Types to Include in Portfolio")
invest_types = st.multiselect(
    "Choose Investments:", options=df["Investment"].tolist(), default=df["Investment"].tolist()
)

# Filter df based on selection
filtered_df = df[df["Investment"].isin(invest_types)].reset_index(drop=True)

# ==== Editable Investment Data Table ====
st.markdown("### Editable Investment Data")
edited = st.experimental_data_editor(filtered_df, num_rows="dynamic")

# ==== Filter Portfolio Controls ====
st.markdown("### Filter Your Desired Portfolio")
col1, col2, col3, col4 = st.columns(4)
with col1:
    risk_s = st.slider(
        "Max Risk (%)", 
        float(edited["Risk (%)"].min()), 
        float(edited["Risk (%)"].max()), 
        float(edited["Risk (%)"].max())
    )
with col2:
    return_s = st.slider(
        "Min Expected Return (%)", 
        float(edited["Expected Return (%)"].min()), 
        float(edited["Expected Return (%)"].max()), 
        float(edited["Expected Return (%)"].min())
    )
with col3:
    liquidity_s = st.slider(
        "Min Liquidity (1-10)", 
        int(edited["Liquidity (1-10)"].min()), 
        int(edited["Liquidity (1-10)"].max()), 
        int(edited["Liquidity (1-10)"].min())
    )
with col4:
    inflation_hedge = st.checkbox("Inflation Hedge Only")

time_horizon = st.number_input("Investment Time Horizon (Years)", min_value=1, max_value=50, value=5)

# Apply filters
filtered_portfolio = edited[
    (edited["Risk (%)"] <= risk_s) &
    (edited["Expected Return (%)"] >= return_s) &
    (edited["Liquidity (1-10)"] >= liquidity_s)
]

if inflation_hedge:
    filtered_portfolio = filtered_portfolio[filtered_portfolio["Inflation Hedge (Yes/No)"] == "Yes"]

filtered_portfolio = filtered_portfolio.reset_index(drop=True)

# ==== Computed Averages ====
st.markdown("### Computed Portfolio Averages")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Risk (%)", f"{filtered_portfolio['Risk (%)'].mean():.2f}")
c2.metric("Expected Return (%)", f"{filtered_portfolio['Expected Return (%)'].mean():.2f}")
c3.metric("Volatility (%)", f"{filtered_portfolio['Volatility (%)'].mean():.2f}")
c4.metric("Liquidity", f"{filtered_portfolio['Liquidity (1-10)'].mean():.2f}")

# ==== 4 Compact Graphs Side by Side ====
st.markdown("### Portfolio Visual Insights")
fig, axs = plt.subplots(1, 4, figsize=(20, 4))

# 1. Expected Return Bar
axs[0].bar(filtered_portfolio["Investment"], filtered_portfolio["Expected Return (%)"], color='#f44336')
axs[0].set_title("Expected Return (%)")
axs[0].tick_params(axis='x', rotation=45)

# 2. Volatility vs Liquidity Scatterplot
axs[1].scatter(filtered_portfolio["Volatility (%)"], filtered_portfolio["Liquidity (1-10)"], color='red')
axs[1].set_title("Volatility vs Liquidity")
axs[1].set_xlabel("Volatility (%)")
axs[1].set_ylabel("Liquidity (1-10)")

# 3. Risk Bar
axs[2].bar(filtered_portfolio["Investment"], filtered_portfolio["Risk (%)"], color='#f44336')
axs[2].set_title("Risk (%)")
axs[2].tick_params(axis='x', rotation=45)

# 4. Minimum Investment Bar
axs[3].bar(filtered_portfolio["Investment"], filtered_portfolio["Minimum Investment ($k)"], color='#f44336')
axs[3].set_title("Minimum Investment ($k)")
axs[3].tick_params(axis='x', rotation=45)

plt.tight_layout()
st.pyplot(fig)

# ==== Other Toggles and Sliders at Bottom ====
st.markdown("---")
st.markdown("### Additional Portfolio Customization")

inflation_hedge_bottom = st.checkbox("Inflation Hedge Only (Bottom Section)", key="inflation_bottom", value=inflation_hedge)
time_horizon_bottom = st.number_input("Investment Time Horizon (Years, Bottom Section)", min_value=1, max_value=50, value=time_horizon, key="time_bottom")

min_investment_bottom = st.slider(
    "Minimum Investment ($k)", 
    int(filtered_portfolio["Minimum Investment ($k)"].min()), 
    int(filtered_portfolio["Minimum Investment ($k)"].max()), 
    int(filtered_portfolio["Minimum Investment ($k)"].min()), 
    key="min_inv_bottom"
)

# ==== Real-Time Market Indices Fixed Right Box ====

# Fake Market Data - Replace with API calls for real data
market_data = {
    "S&P 500": {"value": 5980.87, "change": -1.85, "percent": -0.03},
    "Nasdaq": {"value": 19546.27, "change": +25.18, "percent": +0.13},
    "Dow Jones": {"value": 42171.66, "change": -44.14, "percent": -0.10},
}

# Prepare HTML for market box
market_html = """
<div id="market-box">
  <h2>Market Indices</h2>
"""

for index, info in market_data.items():
    color = "#f44336" if info["change"] < 0 else "#4caf50"
    sign = "" if info["change"] < 0 else "+"
    market_html += f"""
    <div class="metric-label">{index}</div>
    <div class="metric-value">${info['value']:,}</div>
    <div class="metric-delta" style="color:{color};">{sign}{info['change']} ({sign}{info['percent']}%)</div>
    <hr>
    """

market_html += """
</div>
<style>
#market-box {{
    position: fixed;
    top: 70px;
    right: 0;
    width: 180px;
    background: #111;
    color: #fff;
    padding: 15px;
    border-left: 3px solid #f44336;
    height: calc(100vh - 70px);
    overflow-y: auto;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    z-index: 9999;
}}
#market-box h2 {{
    color: #f44336;
    font-size: 18px;
    text-align: center;
    margin-bottom: 12px;
}}
.metric-label {{
    font-size: 14px;
    margin-top: 12px;
    font-weight: 600;
}}
.metric-value {{
    font-size: 16px;
    font-weight: 700;
    margin-top: 2px;
}}
.metric-delta {{
    font-size: 14px;
    margin-top: 1px;
}}
hr {{
    border: 0.5px solid #444;
    margin: 8px 0;
}}
</style>
"""

st.markdown(market_html, unsafe_allow_html=True)
