import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt

# --- Setup ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")
st.markdown("""
<style>
/* Header */
header > div { background-color: #111 !important; color: white !important; }
.css-1d391kg { padding-top:110px !important; }
/* App header */
.app-header { background:#111;color:white;padding:10px 20px;position:fixed;top:0;left:0;width:100vw;z-index:999;}
.app-header h1{margin:0;font-size:24px;font-weight:700;}
.app-header p{margin:4px 0 0;color:#f44336;font-size:13px;}
/* Buttons */
.stButton > button{background:#111;color:#f44336;border:2px solid #f44336;border-radius:6px;padding:0.4em 1em;font-weight:700;font-size:14px;}
.stButton > button:hover{background:#f44336;color:#111;}
/* Right sidebar width*/
[data-testid="stSidebar"][aria-expanded=true] { width: 300px; }
[data-testid="stSidebar"] .stMetric { background:#111;color:white;border-radius:6px;margin-bottom:0.5rem;}
</style>
<div class="app-header">
  <h1>Automated Investment Matrix</h1>
  <p>Modular Investment Analysis Platform with Real-Time Data</p>
</div>
""", unsafe_allow_html=True)

@st.cache_data(ttl=300)
def get_index_price(ticker):
    df = yf.Ticker(ticker).history(period="2d")['Close']
    if len(df)>=2:
        latest, prev = df.iloc[-1], df.iloc[-2]
        diff = latest - prev; pct= diff/prev*100
        return f"${latest:,.2f}", f"{diff:+.2f} ({pct:+.2f}%)"
    return "N/A",""

with st.sidebar:
    st.markdown("## Real-Time Market Indices")
    for t,name in [("^GSPC","S&P 500"),("^IXIC","Nasdaq"),("^DJI","Dow Jones")]:
        val, delta = get_index_price(t)
        st.metric(label=name, value=val, delta=delta)

# --- Data ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = df.columns.str.strip()
    return df

df = load_data()

# 1) Select categories
st.subheader("🔧 1. Select Investment Types")
cats = sorted(df["Category"].dropna().unique())
sel_cats = st.multiselect("Choose categories:", cats, default=cats)

# 2) Editable table
st.subheader("✏️ 2. Investment Data (editable)")
edited = st.data_editor(df[df["Category"].isin(sel_cats)], use_container_width=True, num_rows="fixed")

# 3) Compute averages
st.subheader("📊 3. Portfolio Averages")
c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
c1.metric("Avg Return (%)", f"{edited['Expected Return (%)'].mean():.2f}%" if not edited.empty else "N/A")
c2.metric("Avg Risk", f"{edited['Risk Level (1-10)'].mean():.2f}")
c3.metric("Avg Cap Rate (%)", f"{edited['Cap Rate (%)'].mean():.2f}%" if 'Cap Rate (%)' in edited else "N/A")
c4.metric("Avg Liquidity", f"{edited['Liquidity (1-10)'].mean():.2f}" if 'Liquidity (1-10)' in edited else "N/A")
c5.metric("Avg Volatility", f"{edited['Volatility (1-10)'].mean():.2f}" if 'Volatility (1-10)' in edited else "N/A")
c6.metric("Avg Fees (%)", f"{edited['Fees (%)'].mean():.2f}%" if 'Fees (%)' in edited else "N/A")
c7.metric("Avg Min Inv ($)", f"${edited['Minimum Investment ($)'].mean():,.0f}" if not edited.empty else "N/A")

# 4) Charts
st.subheader("📈 4. Visual Insights")
cols = st.columns(4)
def compact(x,y,kind):
    fig,ax=plt.subplots(figsize=(3,2))
    if kind=="bar": ax.bar(x,y,color="#f44336"); ax.set_xticklabels(x,rotation=45,fontsize=7)
    else:
        ax.scatter(x,y,c='red',alpha=0.7)
        ax.set_xlabel(kind.split(" vs ")[0],fontsize=8)
        ax.set_ylabel(kind.split(" vs ")[1],fontsize=8)
        ax.grid(True,linestyle="--",alpha=0.5)
    plt.tight_layout()
    return fig

with cols[0]:
    st.markdown("**Expected Return (%)**")
    if not edited.empty: st.pyplot(compact(edited["Investment Name"], edited["Expected Return (%)"], "bar"))
with cols[1]:
    st.markdown("**Volatility vs Liquidity**")
    if not edited.empty and 'Volatility (1-10)' in edited and 'Liquidity (1-10)' in edited:
        st.pyplot(compact(edited['Volatility (1-10)'], edited['Liquidity (1-10)'], "Volatility vs Liquidity"))
with cols[2]:
    st.markdown("**Fees vs Return**")
    if not edited.empty and 'Fees (%)' in edited:
        st.pyplot(compact(edited['Fees (%)'], edited['Expected Return (%)'], "Fees vs Return"))
with cols[3]:
    st.markdown("**Risk Distribution**")
    fig,ax=plt.subplots(figsize=(3,2))
    if not edited.empty: ax.hist(edited["Risk Level (1-10)"],bins=10,color="#f44336",alpha=0.8)
    ax.set_xlabel("Risk",fontsize=8); ax.set_ylabel("Count",fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

# 5) Bottom Filters
st.subheader("🎯 5. Portfolio Constraints")
min_inv = st.slider("Min Investment ($)",int(edited["Minimum Investment ($)"].min()),int(edited["Minimum Investment ($)"].max()),int(edited["Minimum Investment ($)"].min()),step=1000)
min_ret = st.slider("Min Expected Return (%)",float(edited["Expected Return (%)"].min()),float(edited["Expected Return (%)"].max()),float(edited["Expected Return (%)"].min()),step=0.1)
max_risk = st.slider("Max Risk Level",int(edited["Risk Level (1-10)"].min()),int(edited["Risk Level (1-10)"].max()),int(edited["Risk Level (1-10)"].max()),step=1)
time_horizon = st.selectbox("Time Horizon", ["Short","Medium","Long"], index=1)
inflation_hedge = st.checkbox("Include only inflation-hedge investments")

# 6) Filter results
filt = edited[
    (edited["Minimum Investment ($)"] >= min_inv) &
    (edited["Expected Return (%)"] >= min_ret) &
    (edited["Risk Level (1-10)"] <= max_risk)
]
if inflation_hedge:
    filt = filt[filt["Inflation Hedge (Yes/No)"]=="Yes"]

st.subheader(f"Filtered Investments ({len(filt)})")
st.dataframe(filt, height=250)

# 7) Export
st.subheader("🧾 Export Reports")
b1,b2 = st.columns(2)
with b1:
    if st.button("📤 Download PowerPoint Report"):
        st.success("PPT export ready (to code)")
with b2:
    if st.button("📥 Download Word Report"):
        st.success("Word export ready (to code)")

