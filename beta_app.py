"""
Equity Beta Calculator — Streamlit Web App
Real-time search dropdown · All NSE + BSE companies
Run: streamlit run beta_app.py
"""

import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import requests
from datetime import date, timedelta, datetime
import io

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="Equity Beta Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Exchange / Index config ───────────────────────────────────────────────────
INDEX_CFG = {
    "NSE": {"yf": "^NSEI",  "name": "NIFTY 50"},
    "BSE": {"yf": "^BSESN", "name": "SENSEX"},
}
NSE_CODES = {"NSI", "NSE", "NIM"}
BSE_CODES = {"BSE", "BOM"}

def classify_exchange(symbol: str, exch_code: str = "") -> str:
    if symbol.endswith(".NS"):  return "NSE"
    if symbol.endswith(".BO"):  return "BSE"
    if exch_code.upper() in NSE_CODES: return "NSE"
    if exch_code.upper() in BSE_CODES: return "BSE"
    return "NSE"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #0d1117; color: #e6edf3; }
[data-testid="stSidebar"] { background: #161b22 !important; border-right: 1px solid #21262d; }

.main-title { font-family: 'IBM Plex Mono', monospace; font-size: 2rem; font-weight: 600; color: #f0f6fc; }
.main-sub   { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; color: #3fb950; letter-spacing: 3px; text-transform: uppercase; margin-top: 4px; }
.sec-hdr    { font-family: 'IBM Plex Mono', monospace; font-size: 0.68rem; color: #3fb950; letter-spacing: 3px; text-transform: uppercase; border-bottom: 1px solid #21262d; padding-bottom: 8px; margin-bottom: 14px; }

/* Search result rows */
.sr-row {
    display: flex; align-items: center; justify-content: space-between;
    background: #161b22; border: 1px solid #21262d; border-radius: 6px;
    padding: 10px 14px; margin-bottom: 5px;
    transition: border-color 0.15s;
}
.sr-row:hover { border-color: #3fb950; }
.sr-name  { font-size: 0.87rem; color: #e6edf3; }
.sr-meta  { font-family: 'IBM Plex Mono', monospace; font-size: 0.72rem; }
.nse-tag  { color: #58a6ff; }
.bse-tag  { color: #f0883e; }
.idx-tag  { color: #484f58; margin-left: 6px; }
.added-tag { font-family: 'IBM Plex Mono', monospace; font-size: 0.72rem; color: #3fb950; }

/* Selected list */
.sel-row  { background: #111820; border: 1px solid #21262d; border-radius: 6px; padding: 8px 14px; margin-bottom: 5px; display: flex; align-items: center; justify-content: space-between; }
.sel-name { font-size: 0.85rem; color: #e6edf3; }

.tag-nse { display: inline-block; background: #0d1f38; border: 1px solid #1f3a5f; border-radius: 4px; padding: 1px 8px; font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; color: #58a6ff; margin-right: 4px; }
.tag-bse { display: inline-block; background: #2a1a0a; border: 1px solid #5a3010; border-radius: 4px; padding: 1px 8px; font-family: 'IBM Plex Mono', monospace; font-size: 0.7rem; color: #f0883e; margin-right: 4px; }

.hint-box { background: #111820; border: 1px solid #21262d; border-radius: 6px; padding: 10px 16px; font-size: 0.8rem; color: #8b949e; margin-bottom: 10px; }
.err-box  { background: #1c1010; border: 1px solid #f85149; border-radius: 6px; padding: 10px 16px; color: #f85149; font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem; margin-bottom: 6px; }
.ok-box   { background: #0d1a0f; border: 1px solid #3fb950; border-radius: 6px; padding: 10px 16px; color: #3fb950; font-family: 'IBM Plex Mono', monospace; font-size: 0.82rem; }
.empty-state { background: #161b22; border: 1px solid #21262d; border-radius: 12px; padding: 52px; text-align: center; margin-top: 24px; }

.metric-card { background: #161b22; border: 1px solid #21262d; border-radius: 8px; padding: 20px 16px; text-align: center; }
.m-ticker { font-family: 'IBM Plex Mono', monospace; font-size: 0.68rem; color: #8b949e; letter-spacing: 1px; }
.m-value  { font-family: 'IBM Plex Mono', monospace; font-size: 2rem; font-weight: 600; }
.m-label  { font-size: 0.72rem; color: #8b949e; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
.m-r2     { font-family: 'IBM Plex Mono', monospace; font-size: 0.72rem; color: #3fb950; margin-top: 6px; }
.m-exch-nse { font-family: 'IBM Plex Mono', monospace; font-size: 0.65rem; color: #58a6ff; margin-top: 4px; font-weight: 600; }
.m-exch-bse { font-family: 'IBM Plex Mono', monospace; font-size: 0.65rem; color: #f0883e; margin-top: 4px; font-weight: 600; }

.no-results { font-size: 0.82rem; color: #484f58; padding: 10px 0; font-style: italic; }
.searching  { font-size: 0.82rem; color: #8b949e; padding: 10px 0; font-family: 'IBM Plex Mono', monospace; }

.stButton > button { background: #238636 !important; color: #fff !important; border: 1px solid #2ea043 !important; border-radius: 6px !important; font-weight: 500 !important; width: 100%; }
.stButton > button[kind="secondary"] { background: #1c2128 !important; color: #8b949e !important; border: 1px solid #30363d !important; }
.stDownloadButton > button { background: #1c2128 !important; color: #58a6ff !important; border: 1px solid #30363d !important; border-radius: 6px !important; width: 100%; }
.stProgress > div > div > div { background-color: #3fb950 !important; }
div[data-testid="stTextInput"] input { background: #161b22 !important; border: 1px solid #30363d !important; color: #e6edf3 !important; font-size: 1rem !important; border-radius: 8px !important; }
div[data-testid="stTextInput"] input:focus { border-color: #3fb950 !important; box-shadow: 0 0 0 2px rgba(63,185,80,0.15) !important; }
#MainMenu{visibility:hidden;} footer{visibility:hidden;} header{visibility:hidden;}
</style>
""", unsafe_allow_html=True)

# ── Excel helpers ─────────────────────────────────────────────────────────────
_CLR = {"h_bg":"1F3864","s_bg":"2F5597","alt":"EEF2F7","wht":"FFFFFF",
        "grn":"C6EFCE","amb":"FFEB9C","bdr":"B8CCE4"}
_T   = Side(style="thin", color=_CLR["bdr"])
_BDR = Border(left=_T, right=_T, top=_T, bottom=_T)

def _xhdr(cell, bg=None, sz=10):
    cell.font      = Font(name="Arial", size=sz, bold=True, color="FFFFFF")
    cell.fill      = PatternFill("solid", start_color=bg or _CLR["h_bg"])
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = _BDR

def _xcell(cell, value, bg="FFFFFF", fmt=None, bold=False, right=False):
    cell.value     = value
    cell.font      = Font(name="Arial", size=10 if bold else 9, bold=bold)
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.border    = _BDR
    cell.alignment = Alignment(horizontal="right" if right else "center", vertical="center")
    if fmt:
        cell.number_format = fmt

# ── Cloud-compatible data fetching ───────────────────────────────────────────
# curl_cffi impersonates Chrome at TLS level — works on Streamlit Cloud (AWS).
# Session is created OUTSIDE cache functions (cache can't serialize sessions).
try:
    from curl_cffi import requests as _curl
    _SESSION = _curl.Session(impersonate="chrome110")
    _HAS_CURL = True
except Exception:
    _SESSION = None
    _HAS_CURL = False

def _get_ticker(symbol: str):
    """Return a yf.Ticker with cloud-safe session."""
    if _HAS_CURL and _SESSION is not None:
        return yf.Ticker(symbol, session=_SESSION)
    return yf.Ticker(symbol)

# ── Core functions ────────────────────────────────────────────────────────────
@st.cache_data(ttl=120, show_spinner=False)
def live_search(query: str) -> list:
    """Live search Yahoo Finance — covers ALL NSE + BSE listed companies."""
    if not query or len(query.strip()) < 2:
        return []
    try:
        if _HAS_CURL and _SESSION is not None:
            results = yf.Search(query.strip(), max_results=12,
                                news_count=0, session=_SESSION)
        else:
            results = yf.Search(query.strip(), max_results=12, news_count=0)
        quotes = results.quotes if hasattr(results, "quotes") else []
        found  = []
        for q in quotes:
            sym = q.get("symbol", "")
            if not (sym.endswith(".NS") or sym.endswith(".BO")):
                continue
            exch = classify_exchange(sym, q.get("exchange", ""))
            name = q.get("longname") or q.get("shortname") or sym
            found.append({
                "name":     name,
                "symbol":   sym,
                "exchange": exch,
                "ticker":   sym.replace(".NS","").replace(".BO",""),
            })
        return found
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_prices(yf_symbol: str, start: str, end: str) -> pd.Series:
    try:
        tkr = _get_ticker(yf_symbol)
        raw = tkr.history(start=start, end=end, auto_adjust=True)
        if raw.empty:
            return pd.Series(dtype=float)
        close = raw["Close"]
        if isinstance(close, pd.DataFrame):
            close = close.iloc[:, 0]
        close.index = pd.to_datetime(close.index).normalize()
        return close.astype(float)
    except Exception:
        return pd.Series(dtype=float)


def calc_beta(comp_ret, idx_ret):
    df = pd.concat([comp_ret, idx_ret], axis=1).dropna()
    df.columns = ["c","i"]
    n = len(df)
    if n < 20:
        return None, None, None, n
    x, y     = df["i"].values, df["c"].values
    slope, _ = np.polyfit(x, y, 1)
    corr     = float(np.corrcoef(x, y)[0,1])
    return round(float(slope),4), round(corr**2,4), round(corr,4), n


def beta_style(b):
    if b is None: return "#8b949e","N/A"
    if b < 0:     return "#f85149","Negative"
    if b < 0.8:   return "#58a6ff","Defensive"
    if b <= 1.2:  return "#3fb950","Market-like"
    if b <= 1.5:  return "#f0883e","Moderate"
    return "#ff6b6b","Aggressive"


# ── Excel export (same as before) ─────────────────────────────────────────────
def build_excel(results: list, start_str: str, end_str: str) -> bytes:
    wb = Workbook()
    s_ws = wb.active; s_ws.title = "Summary"
    m_ws = wb.create_sheet("Methodology")

    s_ws.sheet_view.showGridLines = False
    s_ws.row_dimensions[1].height = 38
    s_ws.row_dimensions[3].height = 28

    s_ws.merge_cells("A1:L1")
    c = s_ws["A1"]
    c.value = "EQUITY BETA ANALYSIS — SUMMARY  (NSE + BSE)"
    c.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color=_CLR["h_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center")

    s_ws.merge_cells("A2:L2")
    c = s_ws["A2"]
    c.value = "Period: " + start_str + " to " + end_str + "  |  Run: " + datetime.today().strftime("%d-%b-%Y") + "  |  NSE → Nifty 50  |  BSE → Sensex"
    c.font = Font(name="Arial", size=9, italic=True, color="FFFFFF")
    c.fill = PatternFill("solid", start_color=_CLR["s_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center")

    hdrs = ["S.No.","Company Name","Ticker","Exchange","Benchmark Index",
            "Start Date","End Date","Observations","Beta","R2","Correlation","Remarks"]
    for j, h in enumerate(hdrs):
        _xhdr(s_ws.cell(row=3, column=j+1, value=h))

    for i, r in enumerate(results):
        row = 4+i; bg = _CLR["alt"] if i%2 else _CLR["wht"]
        vals = [i+1, r["name"], r["ticker"], r["exchange"], r["index_name"],
                r.get("start_date","—"), r.get("end_date","—"), r.get("n_obs",0),
                r.get("beta"), r.get("r2"), r.get("corr"), r.get("error") or ""]
        for j, v in enumerate(vals):
            _xcell(s_ws.cell(row=row, column=j+1), v, bg=bg)
        ec = s_ws.cell(row=row, column=4)
        ec.font = Font(name="Arial", size=9, bold=True,
                       color="1F5297" if r["exchange"]=="NSE" else "8B4513")
        bc = s_ws.cell(row=row, column=9)
        bc.number_format = "0.00"; bc.font = Font(name="Arial", size=10, bold=True)
        if r.get("beta") is not None:
            b = r["beta"]
            fc = _CLR["grn"] if 0.8<=b<=1.5 else (_CLR["amb"] if (b>1.5 or b<0) else _CLR["wht"])
            bc.fill = PatternFill("solid", start_color=fc)
        s_ws.cell(row=row, column=8).number_format  = "#,##0"
        s_ws.cell(row=row, column=10).number_format = "0.0000"
        s_ws.cell(row=row, column=11).number_format = "0.0000"

    note = 5+len(results)
    s_ws.merge_cells("A"+str(note)+":L"+str(note))
    s_ws["A"+str(note)].value = "Green=Beta 0.80-1.50  |  Amber=>1.50 or Negative  |  NSE→Nifty 50  |  BSE→Sensex"
    s_ws["A"+str(note)].font = Font(name="Arial", size=8, italic=True)
    s_ws["A"+str(note)].alignment = Alignment(horizontal="left")

    for col,w in [("A",5),("B",28),("C",14),("D",10),("E",16),
                  ("F",13),("G",13),("H",10),("I",8),("J",8),("K",12),("L",22)]:
        s_ws.column_dimensions[col].width = w
    s_ws.freeze_panes = "A4"

    for r in results:
        if r.get("aligned") is None: continue
        aligned = r["aligned"]
        idx_label = r["index_name"]
        dws = wb.create_sheet((r["ticker"]+"_"+r["exchange"])[:31])
        dws.sheet_view.showGridLines = False
        last_r = 3+len(aligned)

        dws.merge_cells("A1:I1")
        c = dws["A1"]
        c.value = "Historical Data  —  "+r["name"]+"  vs  "+idx_label
        c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=_CLR["h_bg"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        dws.row_dimensions[1].height = 26

        dws.merge_cells("A2:D2")
        c = dws["A2"]
        c.value = r["exchange"]+" stock  |  Benchmark: "+idx_label+"  |  Beta = SLOPE(I4:I"+str(last_r)+", H4:H"+str(last_r)+")"
        c.font = Font(name="Arial", size=9, italic=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=_CLR["s_bg"])
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        dws.merge_cells("E2:I2")
        c = dws["E2"]
        c.value = "=SLOPE(I4:I"+str(last_r)+",H4:H"+str(last_r)+")"
        c.font = Font(name="Arial", size=11, bold=True)
        c.fill = PatternFill("solid", start_color=_CLR["grn"])
        c.number_format = "0.0000"
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _BDR

        h2 = ["Date", idx_label+" Price", idx_label+" Return", "",
               r["name"]+" Price", r["name"]+" Return", "",
               "Index Return (x)", "Stock Return (y)"]
        for j,h in enumerate(h2):
            _xhdr(dws.cell(row=3, column=j+1, value=h), bg=_CLR["s_bg"])

        fmt_map = {1:"#,##0.00",2:"0.00%",4:"#,##0.00",5:"0.00%",7:"0.00%",8:"0.00%"}
        for ii,(dt,row_d) in enumerate(aligned.iterrows()):
            rr = 4+ii; bg = _CLR["alt"] if ii%2 else _CLR["wht"]
            rv = [dt.strftime("%d-%b-%Y"),
                  row_d["index_price"], row_d["index_return"], "",
                  row_d["comp_price"],  row_d["comp_return"],  "",
                  row_d["index_return"],row_d["comp_return"]]
            for j,v in enumerate(rv):
                _xcell(dws.cell(row=rr, column=j+1), v, bg=bg,
                       fmt=fmt_map.get(j), right=(j>0 and j not in (3,6)))

        for col,w in [("A",14),("B",18),("C",18),("D",4),
                      ("E",18),("F",20),("G",4),("H",22),("I",22)]:
            dws.column_dimensions[col].width = w
        dws.freeze_panes = "A4"

    m_ws.sheet_view.showGridLines = False
    m_ws.column_dimensions["A"].width = 100
    mrows = [
        ("BETA CALCULATION — METHODOLOGY (NSE + BSE, ALL LISTED COMPANIES)", True, _CLR["h_bg"], "FFFFFF", 13),
        ("", False, None, None, 10),
        ("1.  DATA SOURCE & EXCHANGE ROUTING", True, _CLR["s_bg"], "FFFFFF", 11),
        ("Live search via Yahoo Finance API — covers ALL NSE (~2,000) and BSE (~5,000+) listed companies.", False, None, None, 10),
        ("NSE stocks: symbol.NS  →  benchmarked vs Nifty 50 (^NSEI).", False, None, None, 10),
        ("BSE stocks: symbol.BO  →  benchmarked vs Sensex (^BSESN).", False, None, None, 10),
        ("Prices: Adjusted Closing Prices — auto-adjusted for splits and dividends.", False, None, None, 10),
        ("", False, None, None, 10),
        ("2.  RETURN CALCULATION", True, _CLR["s_bg"], "FFFFFF", 11),
        ("Daily Return(t) = [Close(t) - Close(t-1)] / Close(t-1) — identical to Excel =(B2-B1)/B1.", False, None, None, 10),
        ("Date alignment: inner join — only days where BOTH the stock and its index have data.", False, None, None, 10),
        ("", False, None, None, 10),
        ("3.  BETA FORMULA", True, _CLR["s_bg"], "FFFFFF", 11),
        ("Beta = SLOPE(Stock Daily Returns, Benchmark Index Daily Returns)", False, None, None, 10),
        ("Python: numpy.polyfit(x, y, 1)[0] — mathematically identical to Excel SLOPE.", False, None, None, 10),
        ("Cell E2 on each sheet has a live =SLOPE() formula for independent audit.", False, None, None, 10),
        ("", False, None, None, 10),
        ("4.  CAPM CAVEATS", True, _CLR["s_bg"], "FFFFFF", 11),
        ("This is LEVERED (equity) beta. Unlever using Hamada: bu = bL / [1 + (1-t)(D/E)].", False, None, None, 10),
        ("Re-lever with target capital structure before applying in CAPM / WACC.", False, None, None, 10),
        ("Ensure consistent benchmark index when comparing betas across a peer group.", False, None, None, 10),
        ("Daily returns produce higher betas than weekly/monthly due to microstructure noise.", False, None, None, 10),
    ]
    for i,(txt,bold,bg,fg,sz) in enumerate(mrows):
        ri = i+1; c = m_ws.cell(row=ri, column=1, value=txt)
        c.font = Font(name="Arial", size=sz, bold=bold, color=fg if fg else "000000")
        if bg: c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        m_ws.row_dimensions[ri].height = 22 if bold else 17

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════
if "selected" not in st.session_state:
    st.session_state.selected = []          # list of {name, symbol, exchange, ticker}
if "search_results" not in st.session_state:
    st.session_state.search_results = []
if "last_query" not in st.session_state:
    st.session_state.last_query = ""


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        "<div style='padding:16px 0 8px;'>"
        "<div style='font-family:IBM Plex Mono,monospace;font-size:1.1rem;font-weight:600;color:#f0f6fc;'>β Calculator</div>"
        "<div style='font-family:IBM Plex Mono,monospace;font-size:0.62rem;color:#3fb950;letter-spacing:2px;margin-top:2px;'>ALL NSE + BSE · AUTO INDEX</div>"
        "</div><hr style='border-color:#21262d;margin:10px 0;'>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='sec-hdr'>DATE RANGE</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("From", value=date(2022,1,1),
                                   min_value=date(2010,1,1),
                                   max_value=date.today()-timedelta(days=30))
    with c2:
        end_date = st.date_input("To", value=date.today(),
                                 min_value=date(2010,1,1),
                                 max_value=date.today())
    if (end_date - start_date).days < 90:
        st.warning("Select at least 90 days.")

    st.markdown("<hr style='border-color:#21262d;margin:14px 0;'>", unsafe_allow_html=True)
    st.markdown("<div class='sec-hdr'>PRESETS</div>", unsafe_allow_html=True)

    PRESETS = {
        "-- None --":  [],
        "Nifty IT":    [("TCS.NS","NSE","Tata Consultancy Services"),("INFY.NS","NSE","Infosys"),("WIPRO.NS","NSE","Wipro"),("HCLTECH.NS","NSE","HCL Technologies"),("TECHM.NS","NSE","Tech Mahindra")],
        "Nifty Bank":  [("HDFCBANK.NS","NSE","HDFC Bank"),("ICICIBANK.NS","NSE","ICICI Bank"),("KOTAKBANK.NS","NSE","Kotak Mahindra Bank"),("AXISBANK.NS","NSE","Axis Bank"),("SBIN.NS","NSE","State Bank of India")],
        "Nifty Top 5": [("RELIANCE.NS","NSE","Reliance Industries"),("TCS.NS","NSE","TCS"),("HDFCBANK.NS","NSE","HDFC Bank"),("INFY.NS","NSE","Infosys"),("ICICIBANK.NS","NSE","ICICI Bank")],
        "Pharma":      [("SUNPHARMA.NS","NSE","Sun Pharmaceutical"),("DRREDDY.NS","NSE","Dr. Reddy's"),("CIPLA.NS","NSE","Cipla"),("DIVISLAB.NS","NSE","Divi's Laboratories"),("LUPIN.NS","NSE","Lupin")],
        "Hospitals":   [("APOLLOHOSP.NS","NSE","Apollo Hospitals"),("FORTIS.NS","NSE","Fortis Healthcare"),("NH.NS","NSE","Narayana Hrudayalaya"),("MAXHEALTH.NS","NSE","Max Healthcare"),("MEDANTA.NS","NSE","Global Health / Medanta")],
        "Auto":        [("MARUTI.NS","NSE","Maruti Suzuki"),("TATAMOTORS.NS","NSE","Tata Motors"),("M&M.NS","NSE","Mahindra & Mahindra"),("BAJAJ-AUTO.NS","NSE","Bajaj Auto"),("HEROMOTOCO.NS","NSE","Hero MotoCorp")],
    }
    preset_choice = st.selectbox("Preset", list(PRESETS.keys()), label_visibility="collapsed")
    if st.button("Load Preset", use_container_width=True):
        if preset_choice != "-- None --":
            existing = {c["symbol"] for c in st.session_state.selected}
            for sym, exch, name in PRESETS[preset_choice]:
                if sym not in existing:
                    st.session_state.selected.append({
                        "name": name, "symbol": sym,
                        "exchange": exch,
                        "ticker": sym.replace(".NS","").replace(".BO",""),
                    })
            st.rerun()

    st.markdown("<hr style='border-color:#21262d;margin:8px 0 14px;'>", unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:0.7rem;color:#484f58;line-height:1.9;'>"
        "<span style='color:#58a6ff;font-weight:600;'>NSE .NS</span> → Nifty 50<br>"
        "<span style='color:#f0883e;font-weight:600;'>BSE .BO</span> → Sensex<br>"
        "Live search · all listed cos.<br>"
        "Beta = SLOPE(Ri, Rm)"
        "</div>",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(
    "<div style='padding:28px 0 6px;'>"
    "<div class='main-title'>Equity Beta Calculator</div>"
    "<div class='main-sub'>All NSE + BSE Companies · Auto Index · SLOPE Methodology</div>"
    "</div><hr style='border-color:#21262d;margin:10px 0 22px;'>",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
#  SEARCH — real-time as user types
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='sec-hdr'>SEARCH COMPANIES</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='hint-box'>"
    "Start typing a company name or ticker — results appear instantly. "
    "Click <strong>＋ Add</strong> to build your list. "
    "Works for <span style='color:#58a6ff;'>all NSE</span> and "
    "<span style='color:#f0883e;'>all BSE</span> listed companies."
    "</div>",
    unsafe_allow_html=True,
)

query = st.text_input(
    "search",
    placeholder="Type company name or ticker  —  e.g.  Dr Agarwals  /  Reliance  /  HDFCBANK",
    label_visibility="collapsed",
    key="live_search_input",
)

# ── Trigger live search whenever query changes ────────────────────────────────
if query != st.session_state.last_query:
    st.session_state.last_query = query
    if len(query.strip()) >= 2:
        st.session_state.search_results = live_search(query)
    else:
        st.session_state.search_results = []

# ── Render dropdown-style results immediately below the input ─────────────────
existing_syms = {c["symbol"] for c in st.session_state.selected}

if len(query.strip()) >= 2:
    results_list = st.session_state.search_results

    if results_list:
        for res in results_list:
            col_name, col_idx, col_btn = st.columns([5, 2, 1])

            exch_class = "nse-tag" if res["exchange"] == "NSE" else "bse-tag"
            idx_name   = INDEX_CFG[res["exchange"]]["name"]
            already    = res["symbol"] in existing_syms

            with col_name:
                st.markdown(
                    "<div class='sr-row'>"
                    "<span class='sr-name'>" + res["name"] + "</span>"
                    "&nbsp;&nbsp;"
                    "<span class='sr-meta " + exch_class + "'>"
                    + res["exchange"] + " · " + res["ticker"] +
                    "</span>"
                    "<span class='sr-meta idx-tag'>→ " + idx_name + "</span>"
                    "</div>",
                    unsafe_allow_html=True,
                )
            with col_idx:
                st.markdown("<div style='height:46px;'></div>", unsafe_allow_html=True)
            with col_btn:
                if already:
                    st.markdown(
                        "<div style='padding:12px 0;'><span class='added-tag'>✓ Added</span></div>",
                        unsafe_allow_html=True,
                    )
                else:
                    if st.button("＋ Add", key="add_" + res["symbol"], use_container_width=True):
                        st.session_state.selected.append(res)
                        existing_syms.add(res["symbol"])
                        st.rerun()

    elif len(query.strip()) >= 2:
        st.markdown(
            "<div class='no-results'>No results for \"" + query + "\" — try a different name or ticker.</div>",
            unsafe_allow_html=True,
        )

elif len(query.strip()) == 1:
    st.markdown(
        "<div class='searching'>Type at least 2 characters to search...</div>",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  SELECTED COMPANIES LIST
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<hr style='border-color:#21262d;margin:22px 0 16px;'>", unsafe_allow_html=True)

count = len(st.session_state.selected)
hdr_text = "SELECTED COMPANIES  (" + str(count) + ")" if count else "SELECTED COMPANIES"
st.markdown("<div class='sec-hdr'>" + hdr_text + "</div>", unsafe_allow_html=True)

if st.session_state.selected:
    for idx, comp in enumerate(st.session_state.selected):
        col_name, col_tag, col_remove = st.columns([5, 2, 1])
        tag_class  = "tag-nse" if comp["exchange"] == "NSE" else "tag-bse"
        idx_name   = INDEX_CFG[comp["exchange"]]["name"]

        with col_name:
            st.markdown(
                "<div style='padding:8px 0;font-size:0.86rem;color:#e6edf3;'>"
                + comp["name"] + "</div>",
                unsafe_allow_html=True,
            )
        with col_tag:
            st.markdown(
                "<div style='padding:8px 0;'>"
                "<span class='" + tag_class + "'>" + comp["ticker"] + "</span>"
                "<span style='font-size:0.7rem;color:#484f58;font-family:IBM Plex Mono,monospace;'>"
                "→ " + idx_name + "</span></div>",
                unsafe_allow_html=True,
            )
        with col_remove:
            if st.button("✕", key="rm_" + str(idx) + comp["symbol"], use_container_width=True):
                st.session_state.selected.pop(idx)
                st.rerun()

    st.markdown("<div style='margin-top:4px;'>", unsafe_allow_html=True)
    if st.button("Clear All", use_container_width=False):
        st.session_state.selected = []
        st.rerun()
else:
    st.markdown(
        "<div style='font-size:0.82rem;color:#484f58;padding:6px 0;'>"
        "No companies selected yet — search above to add.</div>",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════════
#  CALCULATE BUTTON
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<hr style='border-color:#21262d;margin:22px 0 16px;'>", unsafe_allow_html=True)
run_col, _ = st.columns([1, 4])
with run_col:
    run = st.button(
        "⚡  Calculate Beta  (" + str(len(st.session_state.selected)) + " companies)",
        use_container_width=True,
        disabled=(len(st.session_state.selected) == 0),
    )


# ══════════════════════════════════════════════════════════════════════════════
#  CALCULATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════
if run and st.session_state.selected:
    if (end_date - start_date).days < 30:
        st.markdown("<div class='err-box'>Date range too short — minimum 30 days.</div>",
                    unsafe_allow_html=True)
        st.stop()

    start_str = start_date.strftime("%Y-%m-%d")
    end_str   = end_date.strftime("%Y-%m-%d")
    companies = st.session_state.selected

    st.markdown("<hr style='border-color:#21262d;margin:24px 0 16px;'>", unsafe_allow_html=True)
    st.markdown("<div class='sec-hdr'>FETCHING DATA</div>", unsafe_allow_html=True)

    progress = st.progress(0)
    status   = st.empty()

    # Fetch indices
    exchanges_needed = {c["exchange"] for c in companies}
    idx_cache = {}
    for step, exch in enumerate(exchanges_needed):
        cfg = INDEX_CFG[exch]
        status.markdown(
            "<div style='font-family:IBM Plex Mono,monospace;font-size:0.8rem;"
            "color:#8b949e;'>Fetching " + cfg["name"] + "...</div>",
            unsafe_allow_html=True,
        )
        px = fetch_prices(cfg["yf"], start_str, end_str)
        if not px.empty:
            idx_cache[exch] = {"prices":px, "returns":px.pct_change().dropna(), "name":cfg["name"]}
        progress.progress(int(8*(step+1)/max(len(exchanges_needed),1)))

    if not idx_cache:
        st.markdown("<div class='err-box'>Could not fetch index data. Check connection.</div>",
                    unsafe_allow_html=True)
        st.stop()

    results = []
    total   = len(companies)

    for i, comp in enumerate(companies):
        exch = comp["exchange"]
        if exch not in idx_cache:
            results.append({"name":comp["name"],"ticker":comp["ticker"],"exchange":exch,
                             "index_name":INDEX_CFG[exch]["name"],
                             "beta":None,"r2":None,"corr":None,"n_obs":0,
                             "error":"Index unavailable","aligned":None})
            continue

        status.markdown(
            "<div style='font-family:IBM Plex Mono,monospace;font-size:0.8rem;color:#8b949e;'>"
            "Fetching " + comp["ticker"] + " [" + exch + "]  (" + str(i+1) + "/" + str(total) + ")...</div>",
            unsafe_allow_html=True,
        )
        prices = fetch_prices(comp["symbol"], start_str, end_str)
        progress.progress(10 + int(85*(i+1)/total))

        if prices.empty:
            results.append({"name":comp["name"],"ticker":comp["ticker"],"exchange":exch,
                             "index_name":idx_cache[exch]["name"],
                             "beta":None,"r2":None,"corr":None,"n_obs":0,
                             "error":"No price data","aligned":None})
            continue

        comp_ret = prices.pct_change().dropna()
        idx_data = idx_cache[exch]

        aligned = pd.concat({
            "index_price":  idx_data["prices"],
            "index_return": idx_data["returns"],
            "comp_price":   prices,
            "comp_return":  comp_ret,
        }, axis=1).dropna()
        aligned.columns = ["index_price","index_return","comp_price","comp_return"]

        beta, r2, corr, n = calc_beta(aligned["comp_return"], aligned["index_return"])
        results.append({
            "name":       comp["name"],
            "ticker":     comp["ticker"],
            "exchange":   exch,
            "index_name": idx_data["name"],
            "beta":beta,"r2":r2,"corr":corr,"n_obs":n,
            "start_date": aligned.index.min().strftime("%d-%b-%Y") if len(aligned) else "—",
            "end_date":   aligned.index.max().strftime("%d-%b-%Y") if len(aligned) else "—",
            "error":      None if beta is not None else "Insufficient data",
            "aligned":    aligned if beta is not None else None,
        })

    progress.progress(100)
    ok_count = sum(1 for r in results if r["beta"] is not None)
    status.markdown(
        "<div class='ok-box'>Done — " + str(ok_count) + " of " + str(total) + " companies processed.</div>",
        unsafe_allow_html=True,
    )

    # Beta cards
    valid = [r for r in results if r["beta"] is not None]
    if valid:
        st.markdown("<hr style='border-color:#21262d;margin:22px 0 16px;'>", unsafe_allow_html=True)
        st.markdown("<div class='sec-hdr'>BETA RESULTS</div>", unsafe_allow_html=True)
        cols = st.columns(min(len(valid), 5))
        for i, r in enumerate(valid):
            color, label = beta_style(r["beta"])
            exch_cls = "m-exch-nse" if r["exchange"]=="NSE" else "m-exch-bse"
            with cols[i % len(cols)]:
                st.markdown(
                    "<div class='metric-card'>"
                    "<div class='m-ticker'>" + r["ticker"] + "</div>"
                    "<div class='m-value' style='color:"+color+";'>" + str(round(r["beta"],2)) + "</div>"
                    "<div class='m-label'>" + label + "</div>"
                    "<div class='m-r2'>R² = " + str(round(r["r2"],3)) + "</div>"
                    "<div class='" + exch_cls + "'>" + r["exchange"] + " · " + r["index_name"] + "</div>"
                    "</div>",
                    unsafe_allow_html=True,
                )

    # Results table
    st.markdown("<div style='margin-top:22px;'>", unsafe_allow_html=True)
    rows = []
    for r in results:
        _, label = beta_style(r["beta"])
        rows.append({
            "Ticker":    r["ticker"],
            "Exchange":  r["exchange"],
            "Company":   r["name"][:38],
            "Benchmark": r["index_name"],
            "Beta":      str(round(r["beta"],4)) if r["beta"] is not None else "—",
            "Category":  label,
            "R²":        str(round(r["r2"],4)) if r["r2"] else "—",
            "Obs":       r["n_obs"],
            "Period":    r.get("start_date","—") + " to " + r.get("end_date","—"),
            "Status":    r["error"] or "OK",
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True,
                 hide_index=True, height=min(420, 56+len(results)*36))

    # Download
    st.markdown("<hr style='border-color:#21262d;margin:22px 0 16px;'>", unsafe_allow_html=True)
    st.markdown("<div class='sec-hdr'>EXPORT</div>", unsafe_allow_html=True)
    excel_bytes = build_excel(results, start_date.strftime("%d-%b-%Y"), end_date.strftime("%d-%b-%Y"))
    fname = "Beta_NSE_BSE_" + datetime.today().strftime("%Y%m%d") + ".xlsx"

    dl_col, info_col = st.columns([1,3])
    with dl_col:
        st.download_button("⬇ Download Excel Report", data=excel_bytes, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with info_col:
        st.markdown(
            "<div style='font-size:0.78rem;color:#8b949e;padding-top:10px;font-family:IBM Plex Mono,monospace;'>"
            + fname + "<br><span style='color:#484f58;'>Summary · Historical Data · Methodology</span></div>",
            unsafe_allow_html=True,
        )

    bad = [r for r in results if r["error"]]
    if bad:
        st.markdown("<hr style='border-color:#21262d;margin:20px 0 12px;'>", unsafe_allow_html=True)
        st.markdown("<div class='sec-hdr'>WARNINGS</div>", unsafe_allow_html=True)
        for r in bad:
            st.markdown(
                "<div class='err-box'>" + r["ticker"] + " [" + r["exchange"] + "] — " + r["error"] + "</div>",
                unsafe_allow_html=True,
            )

# Empty state
elif not st.session_state.selected:
    st.markdown(
        "<div class='empty-state'>"
        "<div style='font-size:3rem;margin-bottom:16px;'>📊</div>"
        "<div style='font-family:IBM Plex Mono,monospace;font-size:1.1rem;color:#f0f6fc;font-weight:600;margin-bottom:10px;'>Ready to calculate</div>"
        "<div style='font-size:0.84rem;color:#484f58;max-width:460px;margin:0 auto;line-height:1.9;'>"
        "Search companies · select from live results · calculate"
        "</div>"
        "<div style='margin-top:28px;display:flex;justify-content:center;gap:48px;'>"
        "<div style='text-align:center;'><div style='font-family:IBM Plex Mono,monospace;font-size:1.2rem;color:#58a6ff;font-weight:600;'>~2,000</div><div style='font-size:0.68rem;color:#484f58;margin-top:4px;'>NSE companies</div></div>"
        "<div style='text-align:center;'><div style='font-family:IBM Plex Mono,monospace;font-size:1.2rem;color:#f0883e;font-weight:600;'>~5,000+</div><div style='font-size:0.68rem;color:#484f58;margin-top:4px;'>BSE companies</div></div>"
        "<div style='text-align:center;'><div style='font-family:IBM Plex Mono,monospace;font-size:1.2rem;color:#3fb950;font-weight:600;'>AUTO</div><div style='font-size:0.68rem;color:#484f58;margin-top:4px;'>index routing</div></div>"
        "</div></div>",
        unsafe_allow_html=True,
    )
