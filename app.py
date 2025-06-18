import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- Sample Investment Data ---
data = {
    "Investment": [
        "Bonds", "Commodities", "Direct Lending", "Equities",
        "Infrastructure", "Life Settlements", "Real Estate"
    ],
    "Expected Return (%)": [5.0, 7.2, 6.5, 9.0, 7.0, 6.0, 6.5],
    "Risk (%)": [3.0, 10.0, 7.0, 15.0, 8.0, 6.0, 5.0],
    "Liquidity (1–10)": [8, 6, 4, 9, 5, 3, 7],
    "Minimum Investment ($)": [5000, 10000, 25000, 10000, 20000, 30000, 15000],
    "Inflation Hedge (Yes/No)": ["No", "Yes", "No", "No", "Yes", "No", "Yes"]
}

df = pd.DataFrame(data)

# --- Real-Time Market Indices Data (dummy example, replace with real API calls if desired) ---
market_indices = {
    "S&P 500": {"value": "$4,123.45", "change": "-12.34 (-0.30%)", "color": "#f44336"},
    "Nasdaq": {"value": "$13,234.56", "change": "+45.67 (+0.35%)", "color": "#4caf50"},
    "Dow Jones": {"value": "$33,987.65", "change": "-23.45 (-0.07%)", "color": "#f44336"},
}

# --- Custom CSS to style fonts and market box on right ---
st.markdown(
    """
    <style>
        /* General font */
        .streamlit-expanderHeader, .css-1d391kg, .css-18e3th9, .css-1lcbmhc, .css-1fv8s86 {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        /* Market indices box */
        #market-box {
            position: fixed;
            top: 70px;
            right: 0;
            width: 190px;
            background: #111;
            color: #fff;
            padding: 15px 10px 10px 10px;
            border-left: 3px solid #f44336;
            height: calc(100vh - 70px);
            overflow-y: auto;
            z-index: 9999;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 14px;
        }
        #market-box h2 {
            color: #f44336;
            font-size: 18px;
            text-align: center;
            margin-bottom: 15px;
            font-weight: 700;
        }
        .metric-label {
            font-weight: 600;
            margin-top: 12px;
            color: #ccc;
        }
        .metric-value {
            font-weight: 700;
            font-size: 16px;
            margin-top: 2px;
        }
        .metric-delta {
            font-size: 13px;
            margin-top: 2px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Market Indices Box (Right Margin) ---
html_market = "<div id='market-box'><h2>Market Indices</h2>"
for idx, data in market_indices.items():
    color = data["color"]
    html_market += f"""
    <div class='metric-label'>{idx}</div>
    <div class='metric-value'>{data['value']}</div>
    <div class='metric-delta' style='color:{color};'>{data['change']}</div>
    <hr style='border-color:#444; margin:6px 0;'>
    """
html_market += "</div>"

st.markdown(html_market, unsafe_allow_html=True)

# --- Title & Description ---
st.markdown(
    "<h1 style='margin-bottom:3px;'>Automated Investment Matrix</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='margin-top:0;margin-bottom:25px;font-size:16px;color:#444;'>"
    "Portfolio Analysis Platform with Real-Time Data for Portfolio Optimization and Financial Advisory."
    "</p>", unsafe_allow_html=True)

# --- Select Investment Types to Include (at very top) ---
st.subheader("Select Investment Types to Include in Portfolio")
investment_types = st.multiselect(
    "Choose one or more investment types:",
    options=df["Investment"].tolist(),
    default=df["Investment"].tolist()
)

filtered_df = df[df["Investment"].isin(investment_types)].reset_index(drop=True)

# --- Editable Investment Data Table ---
st.subheader("Edit Investment Parameters")
edited = st.experimental_data_editor(
    filtered_df,
    num_rows="dynamic",
    use_container_width=True
)

# --- Check Required Columns Exist ---
required_cols = [
    "Expected Return (%)", "Risk (%)", "Liquidity (1–10)", "Minimum Investment ($)", "Investment"
]
for col in required_cols:
    if col not in edited.columns:
        st.error(f"Error: Required column '{col}' missing from data. Please check your input.")
        st.stop()

# --- Sliders for Filtering Based on User Criteria (at bottom) ---
st.subheader("Filter Your Desired Portfolio Criteria")

def get_min_max(col, default_min=0.0, default_max=100.0):
    try:
        return float(edited[col].min()), float(edited[col].max())
    except Exception:
        return default_min, default_max

ret_min, ret_max = get_min_max("Expected Return (%)", 0.0, 20.0)
risk_min, risk_max = get_min_max("Risk (%)", 0.0, 50.0)
liq_min, liq_max = get_min_max("Liquidity (1–10)", 1, 10)
mininv_min, mininv_max = get_min_max("Minimum Investment ($)", 0, 1000000)

ret_s = st.slider("Min Expected Return (%)", ret_min, ret_max, ret_min, step=0.1)
risk_s = st.slider("Max Risk (%)", risk_min, risk_max, risk_max, step=0.1)
liq_s = st.slider("Min Liquidity (1–10)", liq_min, liq_max, liq_min, step=1)
mininv_s = st.slider("Max Minimum Investment ($)", mininv_min, mininv_max, mininv_max, step=1000)

# Time Horizon and Inflation Hedge toggles
time_horizon = st.number_input("Investment Time Horizon (years)", min_value=1, max_value=50, value=5, step=1)
inflation_hedge = st.checkbox("Include Inflation Hedge Only")

# --- Filter investments based on slider selections ---
filtered_criteria = edited[
    (edited["Expected Return (%)"] >= ret_s) &
    (edited["Risk (%)"] <= risk_s) &
    (edited["Liquidity (1–10)"] >= liq_s) &
    (edited["Minimum Investment ($)"] <= mininv_s)
]

if inflation_hedge:
    filtered_criteria = filtered_criteria[
        filtered_criteria["Inflation Hedge (Yes/No)"] == "Yes"
    ]

# --- Compute Averages ---
avg_return = filtered_criteria["Expected Return (%)"].mean() if not filtered_criteria.empty else 0
avg_risk = filtered_criteria["Risk (%)"].mean() if not filtered_criteria.empty else 0
avg_liquidity = filtered_criteria["Liquidity (1–10)"].mean() if not filtered_criteria.empty else 0
avg_min_inv = filtered_criteria["Minimum Investment ($)"].mean() if not filtered_criteria.empty else 0

# --- Display Averages ---
st.subheader("Portfolio Averages (Filtered)")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Expected Return (%)", f"{avg_return:.2f}")
c2.metric("Risk (%)", f"{avg_risk:.2f}")
c3.metric("Liquidity (1–10)", f"{avg_liquidity:.2f}")
c4.metric("Minimum Investment ($)", f"${avg_min_inv:,.0f}")

# --- Visualization Section ---
st.subheader("Portfolio Visual Insights")

fig, axs = plt.subplots(1, 4, figsize=(18, 4))

# Expected Return bar chart
axs[0].bar(edited["Investment"], edited["Expected Return (%)"], color="#f44336")
axs[0].set_title("Expected Return (%)")
axs[0].tick_params(axis='x', rotation=30)

# Risk bar chart
axs[1].bar(edited["Investment"], edited["Risk (%)"], color="#000000")
axs[1].set_title("Risk (%)")
axs[1].tick_params(axis='x', rotation=30)

# Liquidity bar chart
axs[2].bar(edited["Investment"], edited["Liquidity (1–10)"], color="#888888")
axs[2].set_title("Liquidity (1–10)")
axs[2].tick_params(axis='x', rotation=30)

# Minimum Investment bar chart
axs[3].bar(edited["Investment"], edited["Minimum Investment ($)"], color="#f44336")
axs[3].set_title("Minimum Investment ($)")
axs[3].tick_params(axis='x', rotation=30)

plt.tight_layout()
st.pyplot(fig)

# --- Investment Generator ---
st.subheader("Investment Generator")

invest_filter_types = st.multiselect(
    "Select investment types for generating portfolio:",
    options=edited["Investment"].tolist(),
    default=edited["Investment"].tolist()
)

est_return_slider = st.slider(
    "Desired Minimum Expected Return (%)",
    min_value=0.0,
    max_value=20.0,
    value=5.0,
    step=0.1
)

risk_slider = st.slider(
    "Maximum Acceptable Risk (%)",
    min_value=0.0,
    max_value=50.0,
    value=15.0,
    step=0.1
)

filtered_gen = edited[
    (edited["Investment"].isin(invest_filter_types)) &
    (edited["Expected Return (%)"] >= est_return_slider) &
    (edited["Risk (%)"] <= risk_slider)
]

st.write("### Suggested Investments")
st.dataframe(filtered_gen.reset_index(drop=True))

# --- Export buttons ---
st.markdown("---")
col_exp1, col_exp2 = st.columns(2)

with col_exp1:
    if st.button("Export Report as PowerPoint"):
        st.success("PowerPoint export feature coming soon!")

with col_exp2:
    if st.button("Export Report as Word Document"):
        st.success("Word document export feature coming soon!")
