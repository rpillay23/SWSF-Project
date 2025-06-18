import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Automated Investment Matrix", layout="wide", initial_sidebar_state="collapsed")

# Header box at top
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

st.markdown("<br><br><br>", unsafe_allow_html=True)

# Initial data
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

# Market indices data
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
        width: 180px;
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
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .metric-label {
        font-weight: 600;
        margin-top: 8px;
        margin-bottom: 2px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .metric-value {
        font-weight: 700;
        font-size: 16px;
        margin-bottom: 2px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .metric-delta {
        font-weight: 600;
        margin-bottom: 10px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    hr {
        border: 0.5px solid #333;
        margin: 8px 0;
    }
    </style>
    """
    st.markdown(html, unsafe_allow_html=True)

render_market_indices()

st.markdown("### Select Investment Types to Include in Portfolio")

# Investment type multiselect
selected_investments = st.multiselect(
    "Choose investments",
    options=df["Investment"],
    default=df["Investment"].tolist()
)

filtered_df = df[df["Investment"].isin(selected_investments)].copy()

# Editable inputs table replacement (per investment row)
st.markdown("### Edit Investment Values")

edited_data = {
    "Investment": [],
    "Expected Return (%)": [],
    "Risk (%)": [],
    "Volatility (%)": [],
    "Liquidity (1-10)": [],
    "Inflation Hedge (Yes/No)": [],
    "Minimum Investment": []
}

for i, row in filtered_df.iterrows():
    st.markdown(f"**{row['Investment']}**")
    expected_return = st.number_input(f"Expected Return (%) [{row['Investment']}]", value=row["Expected Return (%)"], key=f"er_{i}", step=0.1)
    risk = st.number_input(f"Risk (%) [{row['Investment']}]", value=row["Risk (%)"], key=f"risk_{i}", step=0.1)
    volatility = st.number_input(f"Volatility (%) [{row['Investment']}]", value=row["Volatility (%)"], key=f"vol_{i}", step=0.1)
    liquidity = st.slider(f"Liquidity (1-10) [{row['Investment']}]", min_value=1, max_value=10, value=int(row["Liquidity (1-10)"]), key=f"liq_{i}")
    inflation_hedge = st.selectbox(f"Inflation Hedge (Yes/No) [{row['Investment']}]", options=["Yes", "No"], index=0 if row["Inflation Hedge (Yes/No)"]=="Yes" else 1, key=f"inf_{i}")
    min_investment = st.number_input(f"Minimum Investment [{row['Investment']}]", value=row["Minimum Investment"], key=f"mininv_{i}", step=100)

    edited_data["Investment"].append(row["Investment"])
    edited_data["Expected Return (%)"].append(expected_return)
    edited_data["Risk (%)"].append(risk)
    edited_data["Volatility (%)"].append(volatility)
    edited_data["Liquidity (1-10)"].append(liquidity)
    edited_data["Inflation Hedge (Yes/No)"].append(inflation_hedge)
    edited_data["Minimum Investment"].append(min_investment)

edited_df = pd.DataFrame(edited_data)

# Filter toggles below the editable table
st.markdown("---")
st.markdown("### Filter Your Desired Portfolio")

hedge = st.checkbox("Inflation Hedge Only?", value=False)
time_horizon = st.slider("Time Horizon (Years)", 1, 30, 5)

if hedge:
    edited_df = edited_df[edited_df["Inflation Hedge (Yes/No)"] == "Yes"]

st.markdown("### Computed Averages")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Expected Return (%)", f"{edited_df['Expected Return (%)'].mean():.2f}")
c2.metric("Risk (%)", f"{edited_df['Risk (%)'].mean():.2f}")
c3.metric("Volatility (%)", f"{edited_df['Volatility (%)'].mean():.2f}")
c4.metric("Liquidity", f"{edited_df['Liquidity (1-10)'].mean():.2f}")

# Investment Visual Insights Graphs

st.markdown("### Investment Visual Insights")

fig, axs = plt.subplots(1, 4, figsize=(18, 4))

axs[0].bar(edited_df['Investment'], edited_df['Expected Return (%)'], color='red')
axs[0].set_title("Expected Return (%)")
axs[0].set_xticklabels(edited_df['Investment'], rotation=45, ha='right', fontsize=8)

axs[1].bar(edited_df['Investment'], edited_df['Risk (%)'], color='red')
axs[1].set_title("Risk (%)")
axs[1].set_xticklabels(edited_df['Investment'], rotation=45, ha='right', fontsize=8)

axs[2].bar(edited_df['Investment'], edited_df['Volatility (%)'], color='red')
axs[2].set_title("Volatility (%)")
axs[2].set_xticklabels(edited_df['Investment'], rotation=45, ha='right', fontsize=8)

axs[3].bar(edited_df['Investment'], edited_df['Liquidity (1-10)'], color='red')
axs[3].set_title("Liquidity (1-10)")
axs[3].set_xticklabels(edited_df['Investment'], rotation=45, ha='right', fontsize=8)

plt.tight_layout()
st.pyplot(fig)

# Sliders for portfolio filtering at bottom

st.markdown("---")
st.markdown("### Portfolio Filtering Controls")

col1, col2, col3 = st.columns(3)
with col1:
    max_risk = st.slider("Max Risk (%)", 0, 30, 15)
with col2:
    min_return = st.slider("Min Expected Return (%)", 0, 20, 5)
with col3:
    min_liquidity = st.slider("Min Liquidity (1-10)", 1, 10, 5)

portfolio_filtered = edited_df[
    (edited_df["Risk (%)"] <= max_risk) &
    (edited_df["Expected Return (%)"] >= min_return) &
    (edited_df["Liquidity (1-10)"] >= min_liquidity)
]

st.markdown("### Filtered Investments Matching Your Criteria")
st.dataframe(portfolio_filtered.style.format({
    "Expected Return (%)": "{:.2f}",
    "Risk (%)": "{:.2f}",
    "Volatility (%)": "{:.2f}",
    "Liquidity (1-10)": "{:.0f}",
    "Minimum Investment": "${:,.0f}"
}), height=250)
