import streamlit as st
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches

# --- PAGE CONFIG ---
st.set_page_config(page_title="Automated Investment Matrix", layout="wide")

# --- LIGHT/DARK MODE TOGGLE ---
dark_mode = st.sidebar.checkbox("Dark Mode", value=True)

# --- CSS STYLES FOR DARK AND LIGHT ---
dark_css = """
<style>
.app-header {
    background:#111;
    color:white;
    padding:12px 24px;
    position:fixed;
    top:0;
    left:0;
    width:100vw;
    z-index:1000;
    font-family:'Helvetica Neue', Helvetica, sans-serif;
    height: 72px;
}
.app-header h1 {
    margin: 0 0 0 0;
    font-size: 24px;
    font-weight: 700;
    line-height: 1.1;
}
.app-header p {
    margin: 0;
    color: #f44336;
    font-size: 13px;
    line-height: 1.1;
}
.market-indices {
    display:flex;
    gap:20px;
    margin-top:20px;
    margin-bottom:1px;
    justify-content:right;
}
.index-box {
    background:#111;
    color:white;
    padding:10px 16px;
    border-radius:4px;
    text-align:center;
    font-size:13px;
    min-width:120px;
}
.index-box .name { font-weight:700; margin-bottom:4px; }
.index-box .value { font-size:16px; margin-bottom:2px; }
.index-box .delta { font-size:14px; }
.index-box .positive { color:#4caf50; }
.index-box .negative { color:#f44336; }
[data-testid="stDataFrame"] { font-size:12px; color: white; background-color:#222; }
.stButton > button {
    background:#111;
    color:#f44336;
    border:2px solid #f44336;
    padding:0.5em 1em;
    font-weight:700;
    border-radius:6px;
}
.stButton > button:hover {
    background:#f44336;
    color:#111;
}
body {
    background-color: #121212;
    color: white;
}
.main > div {
    padding-top: 72px !important;
}
</style>
<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Real‑Time Data for Portfolio Optimization and Advisory</p>
</div>
"""

light_css = """
<style>
.app-header {
    background:#f0f0f0;
    color:#222;
    padding:12px 24px;
    position:fixed;
    top:0;
    left:0;
    width:100vw;
    z-index:1000;
    font-family:'Helvetica Neue', Helvetica, sans-serif;
    height: 72px;
    border-bottom: 1px solid #ccc;
}
.app-header h1 {
    margin: 0 0 0 0;
    font-size: 24px;
    font-weight: 700;
    line-height: 1.1;
}
.app-header p {
    margin: 0;
    color: #a30000;
    font-size: 13px;
    line-height: 1.1;
}
.market-indices {
    display:flex;
    gap:20px;
    margin-top:20px;
    margin-bottom:1px;
    justify-content:right;
}
.index-box {
    background:#f0f0f0;
    color:#222;
    padding:10px 16px;
    border-radius:4px;
    text-align:center;
    font-size:13px;
    min-width:120px;
    border: 1px solid #ccc;
}
.index-box .name { font-weight:700; margin-bottom:4px; }
.index-box .value { font-size:16px; margin-bottom:2px; }
.index-box .delta { font-size:14px; }
.index-box .positive { color:#4caf50; }
.index-box .negative { color:#a30000; }
[data-testid="stDataFrame"] { font-size:12px; color: #222; background-color:#fff; }
.stButton > button {
    background:#f0f0f0;
    color:#a30000;
    border:2px solid #a30000;
    padding:0.5em 1em;
    font-weight:700;
    border-radius:6px;
}
.stButton > button:hover {
    background:#a30000;
    color:#f0f0f0;
}
body {
    background-color: #fff;
    color: #222;
}
.main > div {
    padding-top: 72px !important;
}
</style>
<div class="app-header">
    <h1>Automated Investment Matrix</h1>
    <p>Modular Investment Analysis Platform with Real Real‑Time Data for Portfolio Optimization and Advisory</p>
</div>
"""

# --- APPLY CSS ---
if dark_mode:
    st.markdown(dark_css, unsafe_allow_html=True)
else:
    st.markdown(light_css, unsafe_allow_html=True)

# --- MARKET INDEXES ---
@st.cache_data(ttl=300)
def get_price(ticker):
    hist = yf.Ticker(ticker).history(period="2d")['Close']
    if len(hist) >= 2:
        last, prev = hist.iloc[-1], hist.iloc[-2]
        delta = last - prev
        pct = delta / prev * 100
        return f"${last:,.2f}", f"{delta:+.2f} ({pct:+.2f}%)"
    return "N/A", ""

tickers = {"S&P 500": "^GSPC", "Nasdaq": "^IXIC", "Dow Jones": "^DJI"}
prices = {n: get_price(t) for n, t in tickers.items()}

market_html = '<div class="market-indices">'
for name, (price, delta) in prices.items():
    cls = "positive" if delta.startswith("+") else "negative"
    market_html += f'''
    <div class="index-box">
        <div class="name">{name}</div>
        <div class="value">{price}</div>
        <div class="delta {cls}">{delta}</div>
    </div>'''
market_html += '</div>'
st.markdown(market_html, unsafe_allow_html=True)

# (Continue with your existing code below this line)

# --- DATA LOAD ---
@st.cache_data(ttl=600)
def load_data():
    df = pd.read_excel("Comprehensive_Investment_Matrix.xlsx")
    df.columns = [c.strip().replace("–", "-") for c in df.columns]  # ensure consistent dashes
    return df

df = load_data()

# ... rest of your code unchanged ...
