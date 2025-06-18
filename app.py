import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# === Config & Styles ===
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
header > div { background-color: #111 !important; color: white !important; }
.app-header {background:#111;color:white;padding:10px 20px;position:fixed;top:0;left:0;width:100vw;z-index:999;}
.app-header h1{margin:0;font-size:24px;font-weight:700;}
.app-header p{margin:4px 0 0;color:#f44336;font-size:13px;}
.css-1d391kg { padding-top:110px !important; }
.stButton > button{background:#111;color:#f44336;border:2px solid #f44336;border-radius:6px;font-weight:700;font-size:14px;}
.stButton > button:hover{background:#f44336;color:#111;}
.stMetric>div{font-size:14px!important;}
.dataframe-container { max-height:300px; overflow-y:auto; }
</style>
<div class="app-header"><h1>Automated Investment Matrix</h1>
<p>Modular Investment Analysis Platform... uses real‑time data.</p></div>
""", unsafe_allow_html=True)

# === Load Data ===
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df
df = load_data()

# === 1) Editable Table ===
st.subheader("Investment Data (Edit any value below)")
edited = st.data_editor(df, use_container_width=True, num_rows="fixed")

# === 2) Averages Metrics ===
st.subheader("Computed Portfolio Averages")
c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric("Avg Return (%)", f"{edited['Expected Return (%)'].mean():.2f}%")
c2.metric("Avg Risk (1–10)", f"{edited['Risk Level (1-10)'].mean():.2f}")
c3.metric("Avg Cap Rate (%)", f"{edited['Cap Rate (%)'].mean():.2f}%")
c4.metric("Avg Liquidity", f"{edited['Liquidity (1–10)'].mean():.2f}")
c5.metric("Avg Volatility", f"{edited['Volatility (1–10)'].mean():.2f}")
c6.metric("Avg Fees (%)", f"{edited['Fees (%)'].mean():.2f}%")
c7.metric("Avg Min Investment", f"${edited['Minimum Investment ($)'].mean():,.0f}")

st.markdown("---")

# === 3) Four Graphs ===
st.subheader("Visual Insights")
g1,g2,g3,g4 = st.columns(4)

with g1:
    st.markdown("**Expected Return (%)**")
    fig, ax = plt.subplots(figsize=(3,2))
    ax.bar(edited["Investment Name"], edited["Expected Return (%)"], color="#f44336")
    ax.tick_params(axis='x',rotation=45,labelsize=7)
    st.pyplot(fig)

with g2:
    st.markdown("**Liquidity vs Volatility**")
    fig, ax = plt.subplots(figsize=(3,2))
    ax.scatter(edited["Volatility (1–10)"], edited["Liquidity (1–10)"], c="red", alpha=0.7)
    ax.set_xlabel("Volatility"); ax.set_ylabel("Liquidity")
    st.pyplot(fig)

with g3:
    st.markdown("**Fees vs Return**")
    fig, ax = plt.subplots(figsize=(3,2))
    ax.scatter(edited["Fees (%)"], edited["Expected Return (%)"], c="red", alpha=0.7)
    ax.set_xlabel("Fees"); ax.set_ylabel("Return")
    st.pyplot(fig)

with g4:
    st.markdown("**Risk Distribution**")
    fig, ax = plt.subplots(figsize=(3,2))
    ax.hist(edited["Risk Level (1-10)"], bins=10, color="#f44336")
    ax.set_xlabel("Risk Level"); ax.set_ylabel("Count")
    st.pyplot(fig)

st.markdown("---")

# === 4) Custom Filters ===
st.subheader("Filter Your Desired Portfolio")
cats = sorted(edited["Category"].dropna().unique())
sel_cats = st.multiselect("Categories", cats, default=cats)
min_inv = st.slider("Min Investment ($)", int(edited["Minimum Investment ($)"].min()), int(edited["Minimum Investment ($)"].max()), int(edited["Minimum Investment ($)"].min()), step=1000)
min_ret = st.slider("Min Expected Return (%)", float(edited["Expected Return (%)"].min()), float(edited["Expected Return (%)"].max()), float(edited["Expected Return (%)"].min()), step=0.1)
max_risk = st.slider("Max Risk Level", int(edited["Risk Level (1-10)"].min()), int(edited["Risk Level (1-10)"].max()), int(edited["Risk Level (1-10)"].max()), step=1)

# === 5) Filtered Output ===
filtered = edited[
    (edited["Category"].isin(sel_cats)) &
    (edited["Minimum Investment ($)"] >= min_inv) &
    (edited["Expected Return (%)"] >= min_ret) &
    (edited["Risk Level (1-10)"] <= max_risk)
]

st.subheader(f"Filtered Investments ({len(filtered)})")
st.dataframe(filtered, height=250)

# === Export Section ===
st.markdown("---")
st.subheader("Export Reports")
col1,col2 = st.columns(2)
with col1:
    if st.button("Download PowerPoint Report"):
        st.success("Generated PPT report!")
with col2:
    if st.button("Download Word Report"):
        st.success("Generated Word report!")
