import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- PAGE CONFIG & STYLING ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
header > div { background-color: #111 !important; color: white !important; }
.css-1d391kg { padding-top:110px !important; }

/* Header bar */
.app-header {
    background-color: #111; color: white;
    padding: 10px 20px; position: fixed;
    top: 0; left: 0; width: 100vw; z-index: 999;
}
.app-header h1 { margin:0; font-size:24px; font-weight:700; }
.app-header p { margin:4px 0 0;color:#f44336; font-size:13px; font-weight:500; }

/* Buttons */
.stButton > button {
    background:#111; color:#f44336;
    border:2px solid #f44336;
    border-radius:6px;
    padding:0.4em 1em;
    font-weight:700; font-size:14px;
    transition:0.2s;
}
.stButton > button:hover {
    background:#f44336; color:#111;
}

/* Sidebar styling */
[data-testid="stSidebar"][aria-expanded=true] {
    width: 300px;
}
[data-testid="stSidebar"][aria-expanded=true] .sidebar-content {
    padding: 1rem;
}
[data-testid="stSidebar"][aria-expanded=true] .stMetric {
    background-color: #111;
    color: white;
    border-radius: 6px;
    margin-bottom: 0.5rem;
}
</style>

<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Market Data for Portfolio Optimization</p>
</div>
""", unsafe_allow_html=True)

# --- FUNCTION FOR FETCHING INDEX DATA ---
@st.cache_data(ttl=300)
def get_index_price(ticker):
    try:
        hist = yf.Ticker(ticker).history(period="2d")
        latest = hist['Close'].iloc[-1]
        prev = hist['Close'].iloc[-2]
        diff = latest - prev
        pct = (diff / prev) * 100
        return latest, f"{diff:+.2f} ({pct:+.2f}%)"
    except:
        return None, None

# --- RIGHT SIDEBAR: MARKET INDICES ---
with st.sidebar:
    st.markdown("## Real-Time Market Indices")
    for tick, name in [("^GSPC", "S&P 500"), ("^IXIC", "Nasdaq"), ("^DJI", "Dow Jones")]:
        price, delta = get_index_price(tick)
        if price:
            st.metric(label=name, value=f"${price:,.2f}", delta=delta)

# --- LOAD INVESTMENT DATA ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# --- 1) Editable Table ---
st.subheader("1. Investment Data (Edit values directly)")
edited = st.data_editor(df, use_container_width=True, num_rows="fixed")

# --- 2) Filter Portfolio (Select Categories, Investment, Return & Risk) ---
st.subheader("2. Filter Your Desired Portfolio")
cats = sorted(edited["Category"].dropna().unique())
sel_cats = st.multiselect("Select Categories:", cats, default=cats)
min_inv = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()),
                    int(edited["Minimum Investment ($)"].max()),
                    int(edited["Minimum Investment ($)"].min()), step=1000)
min_ret = st.slider("Min Expected Return (%)", float(edited["Expected Return (%)"].min()),
                    float(edited["Expected Return (%)"].max()),
                    float(edited["Expected Return (%)"].min()), step=0.1)
max_risk = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()),
                     int(edited["Risk Level (1-10)"].max()),
                     int(edited["Risk Level (1-10)"].max()), step=1)

filtered = edited[
    (edited["Category"].isin(sel_cats)) &
    (edited["Minimum Investment ($)"] >= min_inv) &
    (edited["Expected Return (%)"] >= min_ret) &
    (edited["Risk Level (1-10)"] <= max_risk)
]

# --- 3) Computed Portfolio Averages ---
st.subheader("3. Computed Portfolio Averages")
c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric("Avg Return (%)", f"{filtered['Expected Return (%)'].mean():.2f}%")
c2.metric("Avg Risk (1–10)", f"{filtered['Risk Level (1-10)'].mean():.2f}")
c3.metric("Avg Cap Rate (%)", f"{filtered['Cap Rate (%)'].mean():.2f}%")
c4.metric("Avg Liquidity", f"{filtered['Liquidity (1–10)'].mean():.2f}")
c5.metric("Avg Volatility", f"{filtered['Volatility (1–10)'].mean():.2f}")
c6.metric("Avg Fees (%)", f"{filtered['Fees (%)'].mean():.2f}%")
c7.metric("Avg Min Investment", f"${filtered['Minimum Investment ($)'].mean():,.0f}")

# --- 4) Visual Insights: Four Graphs ---
st.subheader("4. Visual Insights")
g1,g2,g3,g4 = st.columns(4)

def plot_compact(x, y, xlabel, ylabel, chart_type='bar'):
    fig, ax = plt.subplots(figsize=(3,2))
    if chart_type == 'bar':
        ax.bar(x, y, color="#f44336")
        ax.tick_params(axis='x', rotation=45, labelsize=7)
    else:
        ax.scatter(x, y, c='red', alpha=0.7)
        ax.set_xlabel(xlabel, fontsize=8)
        ax.set_ylabel(ylabel, fontsize=8)
        ax.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    return fig

with g1:
    st.markdown("**Expected Return (%)**")
    if not filtered.empty:
        fig = plot_compact(filtered["Investment Name"], filtered["Expected Return (%)"], '', 'Return %')
        st.pyplot(fig)

with g2:
    st.markdown("**Volatility vs Liquidity**")
    if 'Volatility (1-10)' in filtered.columns and 'Liquidity (1-10)' in filtered.columns and not filtered.empty:
        fig = plot_compact(filtered['Volatility (1-10)'], filtered['Liquidity (1-10)'], 'Volatility', 'Liquidity', chart_type='scatter')
        st.pyplot(fig)

with g3:
    st.markdown("**Fees vs Return**")
    if 'Fees (%)' in filtered.columns and not filtered.empty:
        fig = plot_compact(filtered['Fees (%)'], filtered['Expected Return (%)'], 'Fees %', 'Return %', chart_type='scatter')
        st.pyplot(fig)

with g4:
    st.markdown("**Risk Distribution**")
    fig, ax = plt.subplots(figsize=(3,2))
    if not filtered.empty:
        ax.hist(filtered["Risk Level (1-10)"], bins=10, color="#f44336", alpha=0.8)
    ax.set_xlabel("Risk", fontsize=8)
    ax.set_ylabel("Count", fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

# --- 5) Filtered Investments Table ---
st.subheader(f"5. Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=250)

# --- Export Report Buttons ---
st.subheader("Export Reports")
b1, b2 = st.columns(2)
with b1:
    if st.button("📤 Download PowerPoint Report"):
        st.success("PowerPoint report ready (to implement)")
with b2:
    if st.button("📥 Download Word Report"):
        st.success("Word report ready (to implement)")

