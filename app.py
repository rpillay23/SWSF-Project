import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Page config
st.set_page_config(
    page_title="Automated Investment Matrix",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- HEADER ---
st.markdown("""
<style>
#header-box {
    position: fixed;
    top: 0;
    left: 0;
    width: 100vw;
    height: 70px;
    background-color: #111;
    color: white;
    display: flex;
    flex-direction: column;
    justify-content: center;
    padding-left: 20px;
    border-bottom: 2px solid #f44336;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    z-index: 9999;
}
#header-title {
    font-size: 22px;
    font-weight: 700;
    margin: 0;
}
#header-subtitle {
    font-size: 13px;
    color: #ddd;
    margin: 0;
}
</style>
<div id="header-box">
  <p id="header-title">Automated Investment Matrix</p>
  <p id="header-subtitle">Portfolio Analysis Platform with Real-Time Data</p>
</div>
""", unsafe_allow_html=True)

st.markdown("<br><br><br>", unsafe_allow_html=True)  # spacer below header

# --- Data Setup ---
# Sample investment data for 7 investment types
data = {
    "Investment": ["Bonds", "Commodities", "Direct Lending", "Equities", "Infrastructure", "Life Settlements", "Real Estate"],
    "Expected Return (%)": [5, 7, 9, 12, 8, 10, 6],
    "Risk (%)": [3, 10, 7, 15, 6, 8, 5],
    "Volatility (%)": [2.5, 12, 6, 18, 7, 9, 4],
    "Liquidity (1-10)": [9, 6, 4, 8, 7, 3, 5],
    "Inflation Hedge (Yes/No)": ["No", "Yes", "No", "No", "Yes", "No", "Yes"],
    "Minimum Investment": [1000, 2000, 5000, 1500, 3000, 4000, 2500]
}
df = pd.DataFrame(data)

# --- Right Fixed Market Index Box ---
# Sample real-time market indices (dummy data)
market_indices = {
    "S&P 500": {"price": 5980.87, "change": -1.85, "pct_change": -0.03},
    "Nasdaq": {"price": 19546.27, "change": 25.18, "pct_change": 0.13},
    "Dow Jones": {"price": 42171.66, "change": -44.14, "pct_change": -0.10},
}

def render_market_indices():
    html = """
    <div id="market-box">
      <h2>Market Indices</h2>
    """
    for name, data in market_indices.items():
        color = "#f44336" if data["change"] < 0 else "#4caf50"
        sign = "" if data["change"] < 0 else "+"
        html += f"""
        <div class="metric-label">{name}</div>
        <div class="metric-value">${data['price']:.2f}</div>
        <div class="metric-delta" style="color:{color};">{sign}{data['change']:.2f} ({sign}{data['pct_change']:.2f}%)</div>
        <hr>
        """
    html += "</div>"
    html += """
    <style>
    #market-box {
        position: fixed;
        top: 70px;
        right: 0;
        width: 160px;
        background: #111;
        color: #fff;
        padding: 15px 10px 15px 15px;
        border-left: 2px solid #f44336;
        height: calc(100vh - 70px);
        overflow-y: auto;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 13px;
        z-index: 9999;
    }
    #market-box h2 {
        color: #f44336;
        font-size: 18px;
        text-align: center;
        margin-bottom: 15px;
    }
    .metric-label {
        font-weight: 600;
        margin-top: 8px;
        margin-bottom: 2px;
    }
    .metric-value {
        font-weight: 700;
        font-size: 16px;
        margin-bottom: 2px;
    }
    .metric-delta {
        font-weight: 600;
        margin-bottom: 10px;
    }
    hr {
        border: 0.5px solid #333;
        margin: 8px 0;
    }
    </style>
    """
    st.markdown(html, unsafe_allow_html=True)

render_market_indices()

# --- Main Content Area ---
# Layout columns for main content to allow space on the right for market box
col_main, col_blank = st.columns([8, 1])

with col_main:
    # Investment types filter (checkboxes)
    st.markdown("### Select Investment Types to Include in Portfolio")
    selected_investments = st.multiselect(
        "Choose investments",
        options=df["Investment"],
        default=df["Investment"].tolist()
    )

    # Filter dataframe based on selection
    filtered_df = df[df["Investment"].isin(selected_investments)].copy()

    # Editable data table for investment parameters
    st.markdown("### Edit Investment Parameters")
    edited = st.experimental_data_editor(
        filtered_df.set_index("Investment"),
        num_rows="dynamic",
        key="editable_table"
    )
    edited.reset_index(inplace=True)

    # Filtering based on inflation hedge toggle and time horizon slider
    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        hedge = st.checkbox("Hedge Inflation Only?", value=False)
    with col2:
        time_horizon = st.slider("Time Horizon (Years)", min_value=1, max_value=30, value=5)
    with col3:
        pass  # empty for spacing or future use

    # Filter for inflation hedge if selected
    if hedge and "Inflation Hedge (Yes/No)" in edited.columns:
        edited = edited[edited["Inflation Hedge (Yes/No)"] == "Yes"]

    # Display computed averages
    st.markdown("### Computed Averages")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Expected Return (%)", f"{edited['Expected Return (%)'].mean():.2f}")
    c2.metric("Risk (%)", f"{edited['Risk (%)'].mean():.2f}")
    c3.metric("Volatility (%)", f"{edited['Volatility (%)'].mean():.2f}")
    c4.metric("Liquidity", f"{edited['Liquidity (1-10)'].mean():.2f}")

    # --- Visual Insights: 4 compact graphs side by side ---
    st.markdown("### Investment Visual Insights")
    fig, axs = plt.subplots(1, 4, figsize=(18, 4))

    # 1) Expected Return Bar Chart
    axs[0].bar(edited['Investment'], edited['Expected Return (%)'], color='#f44336')
    axs[0].set_title("Expected Return (%)")
    axs[0].set_xticklabels(edited['Investment'], rotation=45, ha='right', fontsize=8)
    axs[0].tick_params(axis='y', labelsize=8)

    # 2) Volatility vs Liquidity Scatterplot (points red)
    axs[1].scatter(edited['Volatility (%)'], edited['Liquidity (1-10)'], color='red')
    axs[1].set_xlabel("Volatility (%)")
    axs[1].set_ylabel("Liquidity (1-10)")
    axs[1].set_title("Volatility vs Liquidity")
    axs[1].tick_params(axis='both', labelsize=8)

    # 3) Risk vs Expected Return Scatterplot
    axs[2].scatter(edited['Risk (%)'], edited['Expected Return (%)'], color='red')
    axs[2].set_xlabel("Risk (%)")
    axs[2].set_ylabel("Expected Return (%)")
    axs[2].set_title("Risk vs Expected Return")
    axs[2].tick_params(axis='both', labelsize=8)

    # 4) Minimum Investment Bar Chart
    axs[3].bar(edited['Investment'], edited['Minimum Investment'], color='#f44336')
    axs[3].set_title("Minimum Investment")
    axs[3].set_xticklabels(edited['Investment'], rotation=45, ha='right', fontsize=8)
    axs[3].tick_params(axis='y', labelsize=8)

    plt.tight_layout()
    st.pyplot(fig)

    # --- Filter your desired portfolio ---
    st.markdown("### Filter Your Desired Portfolio")
    col1, col2, col3 = st.columns(3)

    with col1:
        risk_slider = st.slider("Max Risk (%)", 0, 30, 15)
    with col2:
        return_slider = st.slider("Min Expected Return (%)", 0, 20, 5)
    with col3:
        liquidity_slider = st.slider("Min Liquidity (1-10)", 1, 10, 5)

    # Filter investments based on these inputs
    portfolio_filtered = edited[
        (edited["Risk (%)"] <= risk_slider) &
        (edited["Expected Return (%)"] >= return_slider) &
        (edited["Liquidity (1-10)"] >= liquidity_slider)
    ]

    st.markdown("### Filtered Investments Matching Your Criteria")
    st.dataframe(portfolio_filtered.style.format({
        "Expected Return (%)": "{:.2f}",
        "Risk (%)": "{:.2f}",
        "Volatility (%)": "{:.2f}",
        "Liquidity (1-10)": "{:.0f}",
        "Minimum Investment": "${:,.0f}"
    }), height=250)

# END OF FILE

