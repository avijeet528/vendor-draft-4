# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  PwC Brand | No sidebar | Full-width | Orange/Black/White/Grey
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import defaultdict
import openpyxl
import os, re, io, zipfile

try:
    import requests
    REQUESTS_OK = True
except ImportError:
    REQUESTS_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

st.set_page_config(
    page_title="IT Procurement Intelligence",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ════════════════════════════════════════════════════════════
# COLOURS — Orange · Black · White · Grey ONLY
# ════════════════════════════════════════════════════════════
C_ORANGE      = "#D04A02"
C_ORANGE_DARK = "#A33A00"
C_ORANGE_MID  = "#E8703A"
C_ORANGE_LITE = "#FAD4C0"
C_BLACK       = "#1A1A1A"
C_DARK        = "#2D2D2D"
C_MID         = "#4A4A4A"
C_GREY_DARK   = "#7D7D7D"
C_GREY        = "#B0B0B0"
C_GREY_LITE   = "#E0E0E0"
C_GREY_BG     = "#F3F3F3"
C_WHITE       = "#FFFFFF"

CHART_SEQ = [C_ORANGE, C_DARK, C_ORANGE_MID, C_GREY_DARK,
             C_ORANGE_DARK, C_GREY, C_MID, C_BLACK]

CFONT = dict(family="Georgia,'Source Sans Pro',Arial",
             size=11, color=C_DARK)
CBG   = C_GREY_BG
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# CSS — no sidebar, full width, all fixes
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

html, body, [class*="css"], div, p, span, td, th,
label, button, .stMarkdown {
    font-family: 'Source Sans Pro','Helvetica Neue',
                 Arial, sans-serif !important;
}
h1,h2,h3 {
    font-family: Georgia, 'ITC Charter', serif !important;
    font-weight: 700 !important;
}

/* ── Hide sidebar completely ── */
section[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"]  { display: none !important; }
#MainMenu  { visibility: hidden; }
footer     { visibility: hidden; }
header     { visibility: hidden; }

/* ── Full-width layout ── */
.main .block-container {
    background: #F3F3F3 !important;
    max-width: 100% !important;
    padding: 1.2rem 2.5rem !important;
}

/* ── Tabs ── */
button[data-baseweb="tab"] {
    font-weight: 600 !important;
    font-size: 0.90em !important;
    color: #7D7D7D !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #D04A02 !important;
    border-bottom: 3px solid #D04A02 !important;
    background: transparent !important;
}

/* ── KPI boxes ── */
.kpi-box {
    border-radius: 4px; padding: 20px 12px;
    text-align: center; color: white;
    border-left: 5px solid rgba(255,255,255,0.2);
}
.kpi-value {
    font-size: 2.4em; font-weight: 700; margin: 0;
    line-height: 1.1;
    font-family: Georgia, serif !important;
}
.kpi-label {
    font-size: 0.72em; font-weight: 700;
    opacity: 0.9; margin-top: 6px;
    letter-spacing: 1.2px; text-transform: uppercase;
}

/* ── Section heading ── */
.sec-head {
    font-size: 0.73em; font-weight: 700;
    letter-spacing: 1.4px; text-transform: uppercase;
    color: #D04A02; margin: 22px 0 10px;
    border-bottom: 2px solid #D04A02;
    padding-bottom: 5px; display: block;
}

/* ── Tables ── */
.comp-table {
    width: 100%; border-collapse: collapse;
    font-size: 0.81em; border: 1px solid #E0E0E0;
}
.comp-table thead tr { background: #2D2D2D; }
.comp-table thead th {
    padding: 10px 12px; text-align: left;
    font-weight: 700; font-size: 0.78em;
    letter-spacing: 0.5px; text-transform: uppercase;
    color: white !important; border: none;
}
.comp-table tbody tr:nth-child(even) { background: #F8F8F8; }
.comp-table tbody tr:hover { background: #FAD4C0; }
.comp-table tbody td {
    padding: 9px 12px; border-bottom: 1px solid #EBEBEB;
    vertical-align: middle; word-break: break-word;
    color: #2D2D2D;
}

/* ── Vendor badge ── */
.vbadge {
    display: inline-block; padding: 3px 9px;
    border-radius: 2px; color: white;
    font-size: 0.76em; font-weight: 700;
    white-space: nowrap;
}

/* ── Score cards ── */
.scard {
    border-radius: 4px; padding: 16px;
    margin-bottom: 10px; border-left: 5px solid #D04A02;
    background: white;
}
.scard-orange { border-color: #D04A02; background: #FFF5F0; }
.scard-dark   { border-color: #2D2D2D; background: #F5F5F5; }
.scard-grey   { border-color: #7D7D7D; background: #FAFAFA; }

/* ── Verdict ── */
.verdict-good {
    background: #FFF5F0; border: 2px solid #D04A02;
    border-radius: 4px; padding: 14px 18px;
    color: #D04A02; font-weight: 700;
}
.verdict-mid {
    background: #F5F5F5; border: 2px solid #4A4A4A;
    border-radius: 4px; padding: 14px 18px;
    color: #4A4A4A; font-weight: 700;
}
.verdict-bad {
    background: #F0F0F0; border: 2px solid #7D7D7D;
    border-radius: 4px; padding: 14px 18px;
    color: #2D2D2D; font-weight: 700;
}

/* ── CHAT ── */
.chat-header {
    background: #2D2D2D; color: white;
    padding: 12px 18px;
    border-radius: 6px 6px 0 0;
}
.chat-outer {
    background: #F8F8F8; border: 1px solid #E0E0E0;
    border-top: none;
    border-radius: 0 0 0 0;
    padding: 14px 14px 6px;
    min-height: 320px; max-height: 400px;
    overflow-y: auto;
}
.msg-user {
    background: #D04A02; color: white;
    border-radius: 14px 14px 3px 14px;
    padding: 9px 14px; margin: 5px 0 5px auto;
    max-width: 74%; font-size: 0.86em;
    display: inline-block; float: right; clear: both;
}
.msg-bot {
    background: white; color: #2D2D2D;
    border: 1px solid #E0E0E0;
    border-left: 4px solid #D04A02;
    border-radius: 14px 14px 14px 3px;
    padding: 9px 14px; margin: 5px 0;
    max-width: 84%; font-size: 0.86em;
    display: inline-block; float: left; clear: both;
}
.chat-wrap { overflow: hidden; margin-bottom: 3px; }

/* ── Quick-question chips INSIDE chat ── */
.chip-row {
    overflow: hidden; padding: 8px 0 4px;
    border-top: 1px solid #E0E0E0; margin-top: 4px;
}
.chip {
    display: inline-block;
    background: white;
    border: 1.5px solid #D04A02;
    color: #D04A02 !important;
    border-radius: 20px;
    padding: 4px 12px;
    font-size: 0.78em; font-weight: 600;
    margin: 3px 4px; cursor: pointer;
    white-space: nowrap;
}
.chip:hover { background: #FFF5F0; }

/* ── Chat input bar ── */
.chat-input-bar {
    background: white;
    border: 1px solid #E0E0E0;
    border-top: none;
    border-radius: 0 0 6px 6px;
    padding: 10px 12px;
}

/* ── Filter bar ── */
.filter-bar {
    background: #2D2D2D; padding: 14px 20px;
    border-radius: 4px; margin-bottom: 18px;
    border-left: 4px solid #D04A02;
}

/* ── Insight box ── */
.insight-box {
    background: #FFF5F0;
    border-left: 4px solid #D04A02;
    border-radius: 0 4px 4px 0;
    padding: 11px 16px; margin: 10px 0;
    font-size: 0.86em; color: #2D2D2D;
}

/* ── Analysis panel ── */
.analysis-panel {
    background: white; border: 1px solid #E0E0E0;
    border-radius: 6px; padding: 20px 24px;
    margin-bottom: 16px;
    border-top: 4px solid #D04A02;
}

/* ── Global nav header strip ── */
.global-nav {
    background: white;
    border-bottom: 3px solid #D04A02;
    padding: 10px 0 8px;
    margin-bottom: 18px;
    display: flex; align-items: center;
    gap: 28px;
}
.nav-item {
    font-size: 0.84em; font-weight: 600;
    color: #7D7D7D; letter-spacing: 0.3px;
    cursor: default;
}
.nav-item.active { color: #D04A02; }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"(?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}"
    r"(?:[,]\d{3})*(?:\.\d{1,2})?"
    r"|(?:[\$\€\£]\s?)\d{1,3}"
    r"(?:[,\s]\d{3})*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?",
    re.IGNORECASE)
TOTAL_KW = ["grand total","total amount","total price",
            "amount due","net total","total cost",
            "total value","subtotal","total"]

def _parse_num(s):
    try:    return float(re.sub(r"[^\d.]","",str(s)) or "0")
    except: return 0.0

def _fmt(val):
    try:
        v = float(re.sub(r"[^\d.]","",str(val)) or "0")
        return "—" if v <= 0 else "${:,.2f}".format(v)
    except: return str(val)

def _best_price(text):
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx == -1: continue
        snip  = text[max(0,idx-20):idx+300]
        hits  = PRICE_RE.findall(snip)
        valid = [h.strip() for h in hits if _parse_num(h)>=50]
        if valid: return max(valid, key=_parse_num)
    all_h = PRICE_RE.findall(text)
    valid = [h.strip() for h in all_h if _parse_num(h)>=100]
    return max(valid, key=_parse_num) if valid else ""

def _text_from_bytes(content, ext):
    text = ""; ext = ext.lower().strip(".")
    try:
        if ext == "pdf":
            if not PDF_OK: return ""
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for p in pdf.pages:
                    t = p.extract_text()
                    if t: text += t + "\n"
        elif ext in ("xlsx","xls"):
            wb = openpyxl.load_workbook(
                io.BytesIO(content), data_only=True, read_only=True)
            rows_t = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs = "  ".join(
                        str(c) for c in row if c is not None)
                    if rs.strip(): rows_t.append(rs)
            text = "\n".join(rows_t); wb.close()
        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml  = z.read("word/document.xml"
                                  ).decode("utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>"," ",xml)
                    text = re.sub(r"\s{2,}","\n",text)
    except: pass
    return text

def extract_price_from_bytes(content, ext):
    text  = _text_from_bytes(content, ext)
    price = _best_price(text)
    if not price or _parse_num(price) <= 0:
        all_n = PRICE_RE.findall(text)
        valid = [h.strip() for h in all_n if _parse_num(h)>=1000]
        if valid: price = max(valid, key=_parse_num)
    return {"price": price,
            "price_num": _parse_num(price) if price else 0.0,
            "text": text[:5000]}

def extract_price_from_file(fp):
    try:
        with open(fp,"rb") as f: c = f.read()
        return extract_price_from_bytes(c, fp.rsplit(".",1)[-1])
    except: return {"price":"","price_num":0.0,"text":""}

# ════════════════════════════════════════════════════════════
# SCORING
# ════════════════════════════════════════════════════════════
def price_score(new_price, hist_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return None,"No comparison data",0,0,0
    mn = min(valid); mx = max(valid); avg = sum(valid)/len(valid)
    if mx == mn: return 50,"Same as historical average",avg,mn,mx
    score = round((1-(new_price-mn)/(mx-mn))*100,1)
    score = max(0,min(100,score))
    pct   = round((new_price-avg)/avg*100,1)
    lbl   = ("{}% BELOW average — COMPETITIVE".format(abs(pct))
             if new_price < avg
             else "{}% ABOVE average — REVIEW NEEDED".format(abs(pct)))
    return score, lbl, avg, mn, mx

def score_color(s):
    if s is None: return C_GREY_DARK
    if s >= 70:   return C_ORANGE
    if s >= 40:   return C_GREY_DARK
    return C_MID

def get_verdict(ps):
    if ps is None:
        return "⚪ No Data","No comparison data.",C_GREY_DARK
    if ps >= 70:
        return "✅ COMPETITIVE","Priced competitively.",C_ORANGE
    if ps >= 40:
        return "🟡 AVERAGE","Within range. Negotiate.",C_GREY_DARK
    return "🔴 HIGH","Above average. Recommend negotiating.",C_MID

# ════════════════════════════════════════════════════════════
# DATA HELPERS
# ════════════════════════════════════════════════════════════
def _clean_df(df):
    safe = [c for c in df.columns if c != "Services List"]
    df   = df[safe].copy()
    for col in df.columns:
        try:    df[col] = df[col].fillna("").apply(lambda x: str(x).strip())
        except: df[col] = ""
    mask = (df["Category"].apply(lambda x: x in ["","nan"]) &
            df["Vendor"].apply(lambda x: x in ["","nan"]))
    df   = df[~mask].copy()
    df.reset_index(drop=True, inplace=True)
    return df

def _parse_services(v):
    if not v or str(v).strip() in ["","nan","None"]:
        return ["(unspecified)"]
    s = str(v).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts = [p.strip() for p in s.split("\n")
             if p.strip() and p.strip() != "nan"]
    if not parts: parts = [p.strip() for p in s.split(";") if p.strip()]
    if not parts and len(s) < 300:
        parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts if parts else ["(unspecified)"]

def _explode(df):
    df2 = df.copy()
    df2["Services List"] = df2["Comments"].apply(_parse_services)
    dfe = df2.explode("Services List").copy()
    dfe.rename(columns={"Services List":"Service"}, inplace=True)
    dfe["Service"] = dfe["Service"].apply(lambda x: str(x).strip())
    dfe = dfe[~dfe["Service"].isin(
        ["","(unspecified)","nan","None"])].reset_index(drop=True)
    return df2, dfe

def _norm_cols(df):
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category"                              and "Category"     not in col_map: col_map["Category"]     = c
        elif any(k in cl for k in ["vendor","supplier"])and "Vendor"       not in col_map: col_map["Vendor"]       = c
        elif ("file name" in cl or cl=="filename")      and "File Name"    not in col_map: col_map["File Name"]    = c
        elif any(k in cl for k in ["file link","file url"])and "File Link" not in col_map: col_map["File Link"]    = c
        elif any(k in cl for k in ["comment","service","description","scope"])and "Comments" not in col_map: col_map["Comments"]    = c
        elif any(k in cl for k in ["price","cost","amount","quoted"])      and "Quoted Price" not in col_map: col_map["Quoted Price"] = c
    df.rename(columns={v:k for k,v in col_map.items()}, inplace=True)
    for req in ["Category","Vendor","File Name","Comments"]:
        if req not in df.columns: df[req] = ""
    keep = ["Category","Vendor","File Name","Comments"]
    for e in ["File Link","Quoted Price"]:
        if e in df.columns: keep.append(e)
    return df[[c for c in keep if c in df.columns]].copy()

# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    XLS = "Master Catalog.xlsx"; CSV = "master_catalog.csv"
    df  = None
    if os.path.exists(XLS):
        try:
            raw = pd.read_excel(XLS, engine="openpyxl", header=None)
            hr  = 0
            for i, row in raw.iterrows():
                vals = [str(v).strip().lower()
                        for v in row.values if pd.notna(v)]
                if (any("category" in v for v in vals)
                        and any("vendor" in v for v in vals)):
                    hr = i; break
            df = pd.read_excel(XLS, engine="openpyxl", header=hr)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("Excel: {}".format(e)); df = None
    if df is None and os.path.exists(CSV):
        try:
            df = pd.read_csv(CSV)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("CSV: {}".format(e)); df = None
    if df is None: return None, None
    df = _norm_cols(df); df = _clean_df(df)
    df["Hyperlink"] = ""
    if "File Link" in df.columns:
        df["Hyperlink"] = df["File Link"].apply(
            lambda x: "" if x in ["","nan"] else x)
    df, dfe = _explode(df)
    return df, dfe

def process_uploaded_catalog(file_bytes, filename):
    try:
        ext = filename.rsplit(".",1)[-1].lower()
        raw = (pd.read_excel(io.BytesIO(file_bytes),
                              engine="openpyxl", header=None)
               if ext in ("xlsx","xls")
               else pd.read_csv(io.BytesIO(file_bytes), header=None)
               if ext == "csv"
               else None)
        if raw is None: return None, None, "Unsupported type."
        hr = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower()
                    for v in row.values if pd.notna(v)]
            j = " ".join(vals)
            if (any(k in j for k in ["vendor","supplier"])
                    and any(k in j for k in ["file","document"])):
                hr = i; break
        df = (pd.read_excel(io.BytesIO(file_bytes),
                             engine="openpyxl", header=hr)
              if ext in ("xlsx","xls")
              else pd.read_csv(io.BytesIO(file_bytes), header=hr))
        df = df.loc[:,df.columns.notna()]
        df.columns = [str(c).strip() for c in df.columns]
        df.dropna(how="all", inplace=True)
        df = _norm_cols(df); df = _clean_df(df)
        hmap = {}
        if ext in ("xlsx","xls"):
            try:
                wb  = openpyxl.load_workbook(io.BytesIO(file_bytes))
                ws  = wb.active; fc = None; hr2 = None
                for row in ws.iter_rows():
                    for cell in row:
                        if (cell.value
                                and str(cell.value).strip().lower()
                                == "file name"):
                            fc = cell.column; hr2 = cell.row; break
                    if fc: break
                if fc and hr2:
                    for row in ws.iter_rows(min_row=hr2+1):
                        for cell in row:
                            if (cell.column==fc
                                    and cell.value
                                    and cell.hyperlink):
                                hmap[str(cell.value).strip()] = str(
                                    cell.hyperlink.target).strip()
                wb.close()
            except: pass
        df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
        df, dfe = _explode(df)
        return df, dfe, None
    except Exception as e:
        return None, None, str(e)

# ════════════════════════════════════════════════════════════
# SUBCATEGORY
# ════════════════════════════════════════════════════════════
def infer_sub(cat, comments, fname):
    txt = (str(comments)+" "+str(fname)).lower()
    c   = str(cat).lower().strip()
    if "cybersecurity" in c:
        if any(k in txt for k in ["trendmicro","endpoint","antivirus"]): return "Endpoint Protection"
        if any(k in txt for k in ["cyberark","privileged","pam"]):        return "Privileged Access"
        if any(k in txt for k in ["knowbe4","awareness","phishing"]):     return "Security Awareness"
        if any(k in txt for k in ["forescout","nac"]):                    return "Network Access Control"
        if any(k in txt for k in ["siem","splunk","monitor"]):            return "SIEM / Monitoring"
        return "General Security"
    if "network" in c or "telecom" in c:
        if "meraki"    in txt: return "Cisco Meraki"
        if "palo alto" in txt: return "Palo Alto NGFW"
        if "equinix"   in txt: return "Equinix"
        if "cisco"     in txt: return "Cisco Networking"
        return "General Network"
    if "hosting" in c:
        if any(k in txt for k in ["vmware","vcf"]): return "VMware"
        if "oracle" in txt: return "Oracle DB"
        if "netapp" in txt: return "NetApp"
        if any(k in txt for k in ["colo","colocation"]): return "Colocation"
        return "General Hosting"
    if "m365" in c:    return "M365 Licensing"
    if "idam" in c:    return "Identity Migration"
    if "snow" in c:    return "ServiceNow ITSM"
    if "summary" in c: return "Reporting"
    return str(cat).strip().title()

# ════════════════════════════════════════════════════════════
# CHATBOT — with full catalog knowledge + price analysis
# ════════════════════════════════════════════════════════════
def build_catalog_context(df_master, df_exploded):
    """Build a rich text context from the master catalog."""
    if df_master is None or df_exploded is None:
        return ""
    lines = []
    # Category summary
    for cat in sorted(df_master["Category"].unique()):
        d_cat = df_master[df_master["Category"]==cat]
        n_v   = d_cat["Vendor"].nunique()
        n_q   = len(d_cat)
        svcs  = sorted(df_exploded[
            df_exploded["Category"]==cat
        ]["Service"].unique().tolist())[:10]
        lines.append(
            "Category: {} | {} vendors | {} quotes | "
            "Services: {}".format(
                cat, n_v, n_q, ", ".join(svcs)))
    # Vendor summary
    for v in sorted(df_master["Vendor"].unique()):
        d_v  = df_master[df_master["Vendor"]==v]
        cats = ", ".join(sorted(d_v["Category"].unique()))
        svcs = sorted(df_exploded[
            df_exploded["Vendor"]==v
        ]["Service"].unique().tolist())[:8]
        prices = []
        if "Quoted Price" in d_v.columns:
            for p in d_v["Quoted Price"]:
                n = _parse_num(str(p))
                if n > 0: prices.append(n)
        price_info = ""
        if prices:
            price_info = "| avg price: {} | min: {} | max: {}".format(
                _fmt(sum(prices)/len(prices)),
                _fmt(min(prices)), _fmt(max(prices)))
        lines.append(
            "Vendor: {} | categories: {} {} | "
            "services: {}".format(
                v, cats, price_info,
                ", ".join(svcs)))
    return "\n".join(lines)

def chatbot_response(user_msg, df_master, df_exploded,
                     uploaded_file_bytes=None,
                     uploaded_file_name=None):
    msg = user_msg.lower().strip()

    if df_master is None or df_exploded is None:
        return {"type":"text",
                "text":"No catalog loaded yet."}

    all_services = sorted(df_exploded["Service"].unique().tolist())
    all_vendors  = sorted(df_master["Vendor"].unique().tolist())
    all_cats     = sorted(df_master["Category"].unique().tolist())

    # ── Uploaded file ──
    if uploaded_file_bytes is not None and uploaded_file_name:
        ext    = uploaded_file_name.rsplit(".",1)[-1].lower()
        result = extract_price_from_bytes(
            uploaded_file_bytes, ext)
        price  = result["price_num"]
        st.session_state["tab2_upload_price"]  = price
        st.session_state["tab2_upload_fname"]  = uploaded_file_name
        st.session_state["tab2_file_bytes"]    = uploaded_file_bytes
        st.session_state["tab2_file_ext"]      = ext
        st.session_state["chat_redirect_upload"] = True
        if price > 0:
            # Find similar services
            text_lower = result["text"].lower()
            matched_svcs = [s for s in all_services
                            if any(w in text_lower
                                   for w in s.lower().split()
                                   if len(w)>3)][:5]
            hist = []
            if matched_svcs:
                for s in matched_svcs:
                    for _, r in df_exploded[
                            df_exploded["Service"]==s
                    ].iterrows():
                        qp = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp > 0: hist.append(qp)
            verdict_txt = ""
            if hist:
                ps,lbl,avg,mn,mx = price_score(price, hist)
                verdict_txt = (
                    "\n\n📊 **Quick verdict vs "
                    "{} similar quotes:**\n"
                    "Market avg: {} | Min: {} | Max: {}\n"
                    "**{}**".format(
                        len(hist),
                        _fmt(avg),_fmt(mn),_fmt(mx),lbl))
            return {
                "type": "redirect_upload",
                "text": (
                    "📄 Analysed **{}** — "
                    "found price **{}**{}\n\n"
                    "🔄 **Redirecting to Upload & "
                    "Score** for full comparison…".format(
                        uploaded_file_name,
                        _fmt(price), verdict_txt)),
                "price": price,
            }
        return {
            "type": "redirect_upload",
            "text": (
                "📄 Uploaded **{}** — "
                "price not found automatically.\n\n"
                "🔄 Going to **Upload & Score** — "
                "enter price manually there.".format(
                    uploaded_file_name)),
            "price": 0,
        }

    # ── Greetings ──
    if any(w in msg for w in ["hello","hi","hey"]):
        return {
            "type":"text",
            "text":(
                "👋 Hello! I'm your **PwC Procurement "
                "Assistant**.\n\n"
                "I have full knowledge of the master "
                "catalog — **{} quotes**, **{} vendors**, "
                "**{} services**, **{} categories**.\n\n"
                "Ask me:\n"
                "• *Who quoted for Cisco Catalyst?*\n"
                "• *Compare Palo Alto prices*\n"
                "• *What does TrendMicro offer?*\n"
                "• *Which vendor is cheapest for "
                "Cybersecurity?*\n"
                "• Upload a quote file for instant scoring"
                .format(
                    len(df_master),
                    df_master["Vendor"].nunique(),
                    df_exploded["Service"].nunique(),
                    df_master["Category"].nunique()))
        }

    # ── Summary / stats ──
    if any(w in msg for w in [
            "summary","overview","how many",
            "total","stats","catalog"]):
        lines_cat = []
        for cat in all_cats:
            d_c  = df_master[df_master["Category"]==cat]
            n_v  = d_c["Vendor"].nunique()
            n_q  = len(d_c)
            n_s  = df_exploded[
                df_exploded["Category"]==cat
            ]["Service"].nunique()
            prices = []
            if "Quoted Price" in d_c.columns:
                for p in d_c["Quoted Price"]:
                    n = _parse_num(str(p))
                    if n > 0: prices.append(n)
            p_txt = ("avg {}".format(
                _fmt(sum(prices)/len(prices)))
                      if prices else "no prices")
            lines_cat.append(
                "• **{}**: {} quotes · {} vendors · "
                "{} services · {}".format(
                    cat, n_q, n_v, n_s, p_txt))
        return {
            "type":"text",
            "text":(
                "📊 **Catalog Overview**\n\n"
                "{} total quotes · {} vendors · "
                "{} services · {} categories\n\n"
                "**By Category:**\n{}".format(
                    len(df_master),
                    df_master["Vendor"].nunique(),
                    df_exploded["Service"].nunique(),
                    df_master["Category"].nunique(),
                    "\n".join(lines_cat)))
        }

    # ── Price analysis for a category ──
    if any(w in msg for w in [
            "cheapest","expensive","best price",
            "most competitive","price analysis",
            "which vendor"]):
        matched_cat = next(
            (c for c in all_cats
             if c.lower() in msg
             or any(w in msg
                    for w in c.lower().split()
                    if len(w)>3)),
            None)
        scope_df = (
            df_master[df_master["Category"]==matched_cat]
            if matched_cat else df_master)
        vendor_avgs = {}
        for v in scope_df["Vendor"].unique():
            d_v = scope_df[scope_df["Vendor"]==v]
            ps  = []
            if "Quoted Price" in d_v.columns:
                for p in d_v["Quoted Price"]:
                    n = _parse_num(str(p))
                    if n > 0: ps.append(n)
            if ps:
                vendor_avgs[v] = sum(ps)/len(ps)
        if vendor_avgs:
            best_v  = min(vendor_avgs,
                           key=vendor_avgs.get)
            worst_v = max(vendor_avgs,
                           key=vendor_avgs.get)
            ov_avg  = sum(vendor_avgs.values()
                          )/len(vendor_avgs)
            lines_v = []
            for v,avg in sorted(
                    vendor_avgs.items(),
                    key=lambda x:x[1]):
                pct = round((avg-ov_avg)/ov_avg*100,1)
                tag = ("🟠 Cheapest"
                        if v==best_v
                        else "⚫ Most expensive"
                        if v==worst_v
                        else "⚪ Mid-range")
                lines_v.append(
                    "**{}**: {} ({}% vs avg) {}".format(
                        v, _fmt(avg),
                        "+" if pct>0 else ""
                        + str(pct), tag))
            scope_label = (
                "**{}**".format(matched_cat)
                if matched_cat
                else "all categories")
            return {
                "type":"price_analysis",
                "text":(
                    "💰 **Price Analysis — {}**\n\n"
                    "Market average: **{}**\n\n"
                    "{}\n\n"
                    "✅ **{}** is the most competitive "
                    "at **{}**".format(
                        scope_label, _fmt(ov_avg),
                        "\n".join(lines_v),
                        best_v,
                        _fmt(vendor_avgs[best_v]))),
                "vendor_avgs": vendor_avgs,
                "avg": ov_avg,
                "best_vendor": best_v,
            }
        return {
            "type":"text",
            "text":"No price data available for analysis."
        }

    # ── Vendor match ──
    matched_vendor = next(
        (v for v in all_vendors if v.lower() in msg),
        None)

    # ── Service match ──
    matched_services = [
        svc for svc in all_services
        if any(w in msg for w in svc.lower().split()
               if len(w)>3)
        or svc.lower() in msg][:8]

    # ── Category match ──
    matched_cat = next(
        (c for c in all_cats
         if c.lower() in msg
         or any(w in msg
                for w in c.lower().split()
                if len(w)>4)),
        None)

    # ── Vendor + service ──
    if matched_vendor and matched_services:
        svc = matched_services[0]
        d   = df_exploded[
            (df_exploded["Vendor"]==matched_vendor)
            & (df_exploded["Service"]==svc)]
        if not d.empty:
            files = d["File Name"].unique().tolist()
            prices = []
            for f in files:
                rows = df_master[
                    df_master["File Name"]==f]
                if len(rows)>0 and "Quoted Price" in rows.columns:
                    qp = _parse_num(str(rows["Quoted Price"].values[0]))
                    if qp > 0: prices.append(qp)
            p_txt = ("Price: **{}**".format(
                _fmt(prices[0])) if prices
                      else "No price on record.")
            # compare vs others
            all_p_svc = []
            for _, r in df_exploded[
                    df_exploded["Service"]==svc
            ].iterrows():
                qp = _parse_num(str(r.get("Quoted Price","")).strip())
                if qp > 0: all_p_svc.append(qp)
            verdict_txt = ""
            if prices and all_p_svc:
                ps,lbl,avg,mn,mx = price_score(
                    prices[0], all_p_svc)
                verdict_txt = "\n📊 vs market: **{}**".format(lbl)
            return {
                "type":"vendor_service",
                "text":(
                    "✅ **{}** has quoted for "
                    "**{}**.\n\n{}{}\n📄 Files: {}".format(
                        matched_vendor, svc,
                        p_txt, verdict_txt,
                        ", ".join(files[:3]))),
            }
        return {
            "type":"text",
            "text":(
                "❌ **{}** has **not** quoted for "
                "**{}** in our catalog.\n\n"
                "💡 Ask *'Who quoted for {}?'*".format(
                    matched_vendor, svc, svc))
        }

    # ── Price comparison for service ──
    if matched_services and any(w in msg for w in [
            "compare","price","cost","expensive",
            "cheap","competitive","vs","versus"]):
        svc = matched_services[0]
        d   = df_exploded[
            df_exploded["Service"]==svc
        ].drop_duplicates(subset=["Vendor","File Name"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No quotes for **{}**.".format(svc)}
        vp = {}
        for _, r in d.iterrows():
            v  = r["Vendor"]
            qp = _parse_num(str(r.get("Quoted Price","")).strip())
            ck = "px_{}".format(str(r.get("File Name","")).strip())
            ca = st.session_state.get(ck)
            ep = ca["price_num"] if ca else 0.0
            p  = ep if ep>0 else qp
            if p>0: vp[v] = min(vp.get(v,p), p)
        if not vp:
            vendors = d["Vendor"].unique().tolist()
            return {"type":"text",
                    "text":"📋 **{}** quoted by: {}\n"
                           "No price data — upload files "
                           "to score.".format(
                               svc, ", ".join(vendors))}
        avg_p  = sum(vp.values())/len(vp)
        best_v = min(vp, key=vp.get)
        spread = round((max(vp.values())-min(vp.values()))
                       /min(vp.values())*100,1) \
            if min(vp.values())>0 else 0
        lines = []
        for v,p in sorted(vp.items(), key=lambda x:x[1]):
            pct = round((p-avg_p)/avg_p*100,1) if avg_p>0 else 0
            tag = ("🟠 Best" if v==best_v
                   else "⚫ Most exp."
                   if p==max(vp.values())
                   else "⚪ Mid")
            lines.append(
                "**{}**: {} ({:+.1f}% vs avg) {}".format(
                    v, _fmt(p), pct, tag))
        return {
            "type":"comparison",
            "text":(
                "📊 **Price comparison — {}**\n\n"
                "{}\n\nSpread: **{}%** · "
                "Best: **{}** at **{}**".format(
                    svc, "\n".join(lines),
                    spread, best_v,
                    _fmt(min(vp.values())))),
            "service":svc,
            "vendor_prices":vp,
            "avg":avg_p,
            "best_vendor":best_v,
        }

    # ── Who quoted ──
    if matched_services and any(w in msg for w in [
            "who","vendor","quoted","available"]):
        svc = matched_services[0]
        d   = df_exploded[
            df_exploded["Service"]==svc
        ].drop_duplicates(subset=["Vendor"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No vendor quoted for **{}**.".format(svc)}
        vendors = d["Vendor"].unique().tolist()
        n_files = df_exploded[
            df_exploded["Service"]==svc
        ]["File Name"].nunique()
        return {
            "type":"who_quoted",
            "text":(
                "✅ **{}** vendor(s) quoted for "
                "**{}**:\n\n{}\n\n📄 {} files".format(
                    len(vendors), svc,
                    "\n".join(["• **{}**".format(v)
                                for v in vendors]),
                    n_files)),
            "service":svc, "vendors":vendors,
        }

    # ── Vendor profile ──
    if matched_vendor:
        d    = df_exploded[df_exploded["Vendor"]==matched_vendor]
        svcs = sorted(d["Service"].unique().tolist())
        cats = sorted(d["Category"].unique().tolist())
        n_q  = len(df_master[df_master["Vendor"]==matched_vendor])
        prices = []
        dv = df_master[df_master["Vendor"]==matched_vendor]
        if "Quoted Price" in dv.columns:
            for p in dv["Quoted Price"]:
                n = _parse_num(str(p))
                if n>0: prices.append(n)
        p_summary = ""
        if prices:
            p_summary = (
                "\n\n💰 **Prices:** avg {} · "
                "min {} · max {}".format(
                    _fmt(sum(prices)/len(prices)),
                    _fmt(min(prices)),
                    _fmt(max(prices))))
        return {
            "type":"vendor_profile",
            "text":(
                "🏢 **{}**\n\n"
                "📂 Categories: {}\n"
                "📄 {} quotes{}\n"
                "🛠 Services ({}):\n{}".format(
                    matched_vendor,
                    ", ".join(cats), n_q,
                    p_summary, len(svcs),
                    "\n".join(["• {}".format(s)
                                for s in svcs[:12]])
                    + ("\n…+{} more".format(len(svcs)-12)
                       if len(svcs)>12 else ""))),
        }

    # ── Category ──
    if matched_cat:
        d_c   = df_master[df_master["Category"]==matched_cat]
        n_v   = d_c["Vendor"].nunique()
        n_q   = len(d_c)
        n_svc = df_exploded[
            df_exploded["Category"]==matched_cat
        ]["Service"].nunique()
        vendors = sorted(d_c["Vendor"].unique().tolist())
        prices = []
        if "Quoted Price" in d_c.columns:
            for p in d_c["Quoted Price"]:
                n = _parse_num(str(p))
                if n>0: prices.append(n)
        p_txt = ("avg {} · min {} · max {}".format(
            _fmt(sum(prices)/len(prices)),
            _fmt(min(prices)), _fmt(max(prices)))
                  if prices else "no price data")
        return {
            "type":"text",
            "text":(
                "📂 **{}**\n\n"
                "📄 {} quotes · 🏢 {} vendors · "
                "🛠 {} services\n"
                "💰 Prices: {}\n\n"
                "Vendors: {}".format(
                    matched_cat, n_q, n_v, n_svc,
                    p_txt, ", ".join(vendors))),
        }

    # ── Combination ──
    if len(matched_services)>=2 and any(
            w in msg for w in [
                "together","both","combination","and","all"]):
        svc_set = set(matched_services[:3])
        vsmap   = defaultdict(set)
        for _,r in df_exploded.iterrows():
            vsmap[r["Vendor"]].add(r["Service"])
        full    = [v for v,s in vsmap.items()
                   if svc_set.issubset(s)]
        partial = [v for v,s in vsmap.items()
                   if svc_set & s and v not in full]
        return {
            "type":"text",
            "text":(
                ("✅ **{}** vendor(s) cover ALL of: {}\n\n"
                 "Full: {} | Partial: {}".format(
                     len(full),
                     " · ".join(
                         ["**{}**".format(s)
                          for s in svc_set]),
                     ", ".join(full),
                     ", ".join(partial) if partial
                     else "None"))
                if full else
                ("⚠️ No single vendor covers all of: {}\n\n"
                 "Partial: {}".format(
                     " · ".join(
                         ["**{}**".format(s)
                          for s in svc_set]),
                     ", ".join(partial) if partial
                     else "None")))
        }

    # ── Service info ──
    if matched_services:
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc]
        if not d.empty:
            vendors  = d["Vendor"].unique().tolist()
            n_q2     = d["File Name"].nunique()
            prices   = []
            for _,r in d.drop_duplicates(
                    subset=["File Name"]).iterrows():
                qp = _parse_num(str(r.get(
                    "Quoted Price","")).strip())
                if qp>0: prices.append(qp)
            p_txt = ""
            if prices:
                p_txt = (
                    "\n\n💰 Range: {} – {} "
                    "(avg: {})".format(
                        _fmt(min(prices)),
                        _fmt(max(prices)),
                        _fmt(sum(prices)/len(prices))))
            return {
                "type":"service_info",
                "text":(
                    "📋 **{}**\n\n"
                    "🏢 {} vendor(s): {}\n"
                    "📄 {} files{}\n\n"
                    "💡 Ask *'compare {} prices'*".format(
                        svc, len(vendors),
                        ", ".join(vendors[:5]),
                        n_q2, p_txt, svc)),
            }

    # ── List vendors ──
    if any(p in msg for p in [
            "list vendor","all vendor","show vendor"]):
        lines = []
        for v in all_vendors:
            n_q = len(df_master[df_master["Vendor"]==v])
            n_s = df_exploded[
                df_exploded["Vendor"]==v
            ]["Service"].nunique()
            prices = []
            dv = df_master[df_master["Vendor"]==v]
            if "Quoted Price" in dv.columns:
                for p in dv["Quoted Price"]:
                    n = _parse_num(str(p))
                    if n>0: prices.append(n)
            p_txt = ("avg {}".format(
                _fmt(sum(prices)/len(prices)))
                      if prices else "no prices")
            lines.append(
                "**{}** — {} quotes · {} services · {}"
                .format(v,n_q,n_s,p_txt))
        return {"type":"text",
                "text":"**Vendors:**\n\n"+"\n".join(lines)}

    # ── Fallback ──
    sugg = []
    if matched_services:
        sugg.append("*'Compare {}?'*".format(
            matched_services[0]))
    if matched_vendor:
        sugg.append("*'Profile of {}?'*".format(
            matched_vendor))
    fb = ("I couldn't find specific data. "
          "Try asking about a vendor, service, "
          "category or price.")
    if sugg: fb += "\n\n💡 Try: " + " or ".join(sugg)
    return {"type":"text","text":fb}

def render_chat_chart(resp):
    if resp is None: return
    vp  = resp.get("vendor_prices") or resp.get("vendor_avgs")
    avg = resp.get("avg",0)
    bv  = resp.get("best_vendor","")
    if not vp: return
    sv  = sorted(vp.items(), key=lambda x:x[1])
    bc  = [C_ORANGE if v==bv
           else C_DARK if p==max(vp.values())
           else C_GREY_DARK
           for v,p in sv]
    fig = go.Figure(go.Bar(
        x=[v for v,_ in sv], y=[p for _,p in sv],
        marker_color=bc, marker_line_width=0,
        text=[_fmt(p) for _,p in sv],
        textposition="outside",
        textfont=dict(size=10,color=C_DARK)))
    if avg>0:
        fig.add_hline(y=avg,line_dash="dash",
                       line_color=C_DARK,line_width=1.5,
                       annotation_text="Avg: {}".format(_fmt(avg)),
                       annotation_position="top right")
    fig.update_layout(
        height=220,plot_bgcolor=CBG,paper_bgcolor=CBG,
        margin=dict(l=5,r=10,t=16,b=5),font=CFONT,
        yaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,
                   zeroline=False),
        xaxis=dict(tickangle=-10,tickfont=dict(size=10)),
        bargap=0.45,showlegend=False)
    st.plotly_chart(fig,use_container_width=True)

# ════════════════════════════════════════════════════════════
# UI HELPERS
# ════════════════════════════════════════════════════════════
def vpill(v, color=None):
    return ("<span class='vbadge' style='background:{}'>"
            "{}</span>".format(color or C_DARK, v))

def sec(txt, caption=""):
    st.markdown(
        "<span class='sec-head'>{}</span>".format(txt),
        unsafe_allow_html=True)
    if caption:
        st.markdown(
            "<div style='font-size:0.81em;"
            "color:#7D7D7D;margin:-6px 0 10px'>"
            "{}</div>".format(caption),
            unsafe_allow_html=True)

def kpi_box(col, val, lbl, bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(bg,val,lbl),
        unsafe_allow_html=True)

def mini_kpi(col, val, lbl, bg, icon=""):
    col.markdown(
        "<div class='kpi-box' "
        "style='background:{};min-height:82px'>"
        "<div style='font-size:1.3em;margin-bottom:2px'>"
        "{}</div>"
        "<div class='kpi-value' "
        "style='font-size:1.3em'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(bg,icon,val,lbl),
        unsafe_allow_html=True)

def pwc_header(title, subtitle=""):
    st.markdown(
        "<div style='background:#2D2D2D;color:white;"
        "padding:22px 28px;border-radius:4px;"
        "border-left:6px solid #D04A02;"
        "margin-bottom:16px'>"
        "<div style='font-size:0.68em;font-weight:700;"
        "letter-spacing:2px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:6px'>"
        "IT PROCUREMENT</div>"
        "<h1 style='margin:0;font-size:1.5em;"
        "font-weight:700;color:white;"
        "font-family:Georgia,serif'>{}</h1>"
        "{}</div>".format(
            title,
            "<p style='margin:7px 0 0;opacity:0.55;"
            "font-size:0.84em'>{}</p>".format(subtitle)
            if subtitle else ""),
        unsafe_allow_html=True)

def insight(text):
    st.markdown(
        "<div class='insight-box'>💡 {}</div>".format(text),
        unsafe_allow_html=True)

def resolve_url(row):
    fname = str(row.get("File Name","")).strip()
    if fname:
        local = os.path.join(DEMO_DIR, fname)
        if os.path.exists(local): return local
    for col in ["Hyperlink","File Link"]:
        val = str(row.get(col,"")).strip()
        if val and val not in ["","nan"] and val.startswith("http"):
            return val
    return ""

def pwc_bar(fig_obj, title="", height=320):
    fig_obj.update_layout(
        height=height, plot_bgcolor=CBG,
        paper_bgcolor=CBG,
        margin=dict(l=5,r=10,
                    t=36 if title else 16,b=10),
        font=CFONT,
        yaxis=dict(showgrid=True,
                   gridcolor=C_GREY_LITE,
                   zeroline=False),
        bargap=0.35, showlegend=False)
    if title:
        fig_obj.update_layout(title=dict(
            text=title,
            font=dict(size=12,color=C_DARK,
                      family="Georgia,serif"),
            x=0,xanchor="left"))

# ════════════════════════════════════════════════════════════
# SESSION STATE
# ════════════════════════════════════════════════════════════
for k,v in [
    ("chat_history",[]),
    ("gh_prices_loaded",False),
    ("real_analysis_done",False),
    ("chat_redirect_upload",False),
    ("tab2_upload_price",0.0),
    ("tab2_upload_fname",""),
    ("tab2_file_bytes",None),
    ("tab2_file_ext",""),
    ("last_chat_file",None),
]:
    if k not in st.session_state:
        st.session_state[k] = v

# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state
        and st.session_state["uploaded_catalog_df"]
        is not None):
    df_master   = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "file"

NO_DATA = (df_master is None
           or df_exploded is None
           or df_master.empty)

vendor_color_map = {}
if not NO_DATA:
    for i,v in enumerate(
            sorted(df_master["Vendor"].unique())):
        vendor_color_map[v] = (
            CHART_SEQ[i % len(CHART_SEQ)])

# ════════════════════════════════════════════════════════════
# MAIN HEADER  (full-width, no sidebar)
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div style='background:#2D2D2D;color:white;"
    "padding:24px 32px;border-radius:4px;"
    "border-left:8px solid #D04A02;"
    "margin-bottom:20px'>"
    "<div style='font-size:0.68em;font-weight:700;"
    "letter-spacing:2.5px;text-transform:uppercase;"
    "color:#D04A02;margin-bottom:8px'>"
    "IT PROCUREMENT · INTELLIGENCE DASHBOARD</div>"
    "<h1 style='margin:0;font-size:1.85em;"
    "font-weight:700;color:white;"
    "font-family:Georgia,serif'>"
    "Procurement Intelligence Dashboard</h1>"
    "<p style='margin:8px 0 0;opacity:0.5;"
    "font-size:0.85em'>"
    "Catalog overview · Browse &amp; chat · "
    "Upload &amp; score · Vendor analysis · "
    "Real file analysis</p>"
    "</div>",
    unsafe_allow_html=True)

# ── Global KPI row ──
if not NO_DATA:
    k1,k2,k3,k4 = st.columns(4)
    kpi_box(k1, df_master["File Name"].nunique(),
            "Total Quotes",    C_ORANGE)
    kpi_box(k2, df_exploded["Service"].nunique(),
            "Unique Services", C_DARK)
    kpi_box(k3, df_master["Vendor"].nunique(),
            "Vendors",         C_MID)
    kpi_box(k4, df_master["Category"].nunique(),
            "Categories",      C_GREY_DARK)
    st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# GLOBAL NAV HEADER — shown on Catalog & Browse tabs
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    cats_list    = sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["","nan"]])
    vendors_list = sorted([
        v for v in df_master["Vendor"].unique()
        if str(v).strip() not in ["","nan"]])

    nav_cols = st.columns(len(cats_list) + 1)
    nav_cols[0].markdown(
        "<div style='font-size:0.72em;font-weight:700;"
        "color:#D04A02;letter-spacing:1px;"
        "text-transform:uppercase;padding-top:4px'>"
        "CATEGORIES</div>",
        unsafe_allow_html=True)
    for i, cat in enumerate(cats_list):
        cat_count = len(df_master[
            df_master["Category"]==cat])
        nav_cols[i+1].markdown(
            "<div style='background:white;"
            "border:1px solid #E0E0E0;"
            "border-top:3px solid #D04A02;"
            "border-radius:4px;"
            "padding:8px 10px;text-align:center'>"
            "<div style='font-size:0.75em;"
            "font-weight:700;color:#2D2D2D;"
            "white-space:nowrap;overflow:hidden;"
            "text-overflow:ellipsis'>{}</div>"
            "<div style='font-size:1.1em;"
            "font-weight:800;color:#D04A02;"
            "font-family:Georgia,serif'>{}</div>"
            "<div style='font-size:0.65em;"
            "color:#7D7D7D;text-transform:uppercase;"
            "letter-spacing:0.5px'>quotes</div>"
            "</div>".format(cat, cat_count),
            unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
(tab0, tab1, tab2,
 tab3, tab4, tab5, tab6) = st.tabs([
    "🗂️ Catalog Overview",
    "💬 Browse & Verdict",
    "📤 Upload & Score",
    "📄 Data Table",
    "🗂 Upload Catalog",
    "🔍 Vendor Analysis",
    "📂 Real Analysis",
])

# ════════════════════════════════════════════════════════════
# TAB 0 — CATALOG OVERVIEW
# ════════════════════════════════════════════════════════════
with tab0:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        df_ov = df_master.copy()
        df_ov["Subcategory"] = df_ov.apply(
            lambda r: infer_sub(
                r.get("Category",""),
                r.get("Comments",""),
                r.get("File Name","")), axis=1)
        all_cats_ov = sorted([
            c for c in df_ov["Category"].unique()
            if str(c).strip() not in ["","nan"]])

        pwc_header("Catalog Overview",
                   "Full picture of all categories, "
                   "vendors, subcategories and services")

        k0a,k0b,k0c,k0d,k0e = st.columns(5)
        mini_kpi(k0a, len(df_ov),
                 "Total Quotations", C_ORANGE,"📄")
        mini_kpi(k0b, df_ov["Vendor"].nunique(),
                 "Unique Vendors",   C_DARK,  "🏢")
        mini_kpi(k0c, df_ov["Category"].nunique(),
                 "Categories",       C_MID,   "📂")
        mini_kpi(k0d, df_ov["Subcategory"].nunique(),
                 "Subcategories",    C_GREY_DARK,"🏷️")
        mini_kpi(k0e,
                 df_exploded["Service"].nunique()
                 if df_exploded is not None else "—",
                 "Unique Services",  C_BLACK, "🛠")
        st.markdown("<br>", unsafe_allow_html=True)

        # Category cards
        sec("CATEGORIES AT A GLANCE")
        cat_stats = []
        for cat in all_cats_ov:
            d_cat = df_ov[df_ov["Category"]==cat]
            n_s   = (df_exploded[
                df_exploded["Category"]==cat
            ]["Service"].nunique()
                     if df_exploded is not None else 0)
            cat_stats.append({
                "Category":cat,
                "Quotations":len(d_cat),
                "Vendors":d_cat["Vendor"].nunique(),
                "Subcategories":d_cat["Subcategory"].nunique(),
                "Services":n_s})
        cat_stats_df = pd.DataFrame(cat_stats
            ).sort_values("Quotations",ascending=False)

        CAT_ICONS = {
            "Cybersecurity":"🛡️",
            "Network & Telecom":"🌐",
            "Hosting":"🖥️",
            "M365 & Power Platform":"☁️",
            "IdAM":"🔑",
            "Service Management (SNow)":"⚙️",
            "Summary & Reporting":"📊"}

        rows3 = [cat_stats_df.iloc[i:i+3]
                 for i in range(0,len(cat_stats_df),3)]
        for chunk in rows3:
            cols = st.columns(len(chunk), gap="medium")
            for ci,(_,rs) in enumerate(chunk.iterrows()):
                cn    = rs["Category"]
                icon  = CAT_ICONS.get(cn,"📁")
                cidx  = all_cats_ov.index(cn) \
                    if cn in all_cats_ov else 0
                color = CHART_SEQ[cidx % len(CHART_SEQ)]
                cols[ci].markdown(
                    "<div style='background:white;"
                    "border:1px solid #E0E0E0;"
                    "border-radius:6px;"
                    "padding:18px 16px;"
                    "border-top:4px solid {}'>"
                    "<div style='font-size:1.5em;"
                    "margin-bottom:6px'>{}</div>"
                    "<div style='font-size:0.92em;"
                    "font-weight:700;color:#2D2D2D;"
                    "margin-bottom:12px;"
                    "font-family:Georgia,serif'>{}</div>"
                    "<div style='display:flex;"
                    "gap:6px;flex-wrap:wrap'>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.73em;font-weight:700;"
                    "color:#2D2D2D'>📄 {} quotes</span>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.73em;font-weight:700;"
                    "color:#4A4A4A'>🏢 {} vendors</span>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.73em;font-weight:700;"
                    "color:#7D7D7D'>🛠 {} services</span>"
                    "</div></div>".format(
                        color,icon,cn,
                        rs["Quotations"],
                        rs["Vendors"],
                        rs["Services"]),
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Charts
        sec("DISTRIBUTION — QUOTATIONS & VENDORS")
        ch1,ch2 = st.columns(2,gap="large")
        with ch1:
            fig1 = go.Figure(go.Bar(
                x=cat_stats_df["Category"],
                y=cat_stats_df["Quotations"],
                marker_color=C_ORANGE,
                marker_line_width=0,
                text=cat_stats_df["Quotations"],
                textposition="outside"))
            pwc_bar(fig1,"Quotations per Category")
            fig1.update_xaxes(tickangle=-30,
                               tickfont=dict(size=9.5))
            st.plotly_chart(fig1,use_container_width=True)
        with ch2:
            fig2 = go.Figure(go.Bar(
                x=cat_stats_df["Category"],
                y=cat_stats_df["Vendors"],
                marker_color=C_DARK,
                marker_line_width=0,
                text=cat_stats_df["Vendors"],
                textposition="outside"))
            pwc_bar(fig2,"Vendors per Category")
            fig2.update_xaxes(tickangle=-30,
                               tickfont=dict(size=9.5))
            st.plotly_chart(fig2,use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        sec("CATALOG COMPOSITION")
        dl,dr = st.columns([1,2],gap="large")
        with dl:
            fig_d = px.pie(
                cat_stats_df,
                values="Quotations",names="Category",
                hole=0.55,
                color_discrete_sequence=CHART_SEQ)
            fig_d.update_traces(
                textposition="outside",
                textinfo="percent+label",
                textfont_size=10,
                pull=[0.03]*len(cat_stats_df))
            fig_d.update_layout(
                height=360,
                margin=dict(l=10,r=10,t=10,b=10),
                paper_bgcolor=CBG,font=CFONT,
                showlegend=False)
            st.plotly_chart(fig_d,use_container_width=True)
        with dr:
            tbl = [
                "<table class='comp-table'>"
                "<thead><tr>"
                "<th>Category</th>"
                "<th style='text-align:center'>Quotes</th>"
                "<th style='text-align:center'>Vendors</th>"
                "<th style='text-align:center'>Services</th>"
                "<th style='text-align:center'>Subcats</th>"
                "</tr></thead><tbody>"]
            for _,rs in cat_stats_df.iterrows():
                cn    = rs["Category"]
                cidx  = all_cats_ov.index(cn) \
                    if cn in all_cats_ov else 0
                color = CHART_SEQ[cidx%len(CHART_SEQ)]
                icon  = CAT_ICONS.get(cn,"📁")
                tbl.append(
                    "<tr><td>"
                    "<span style='border-left:"
                    "4px solid {};padding-left:8px;"
                    "font-weight:600'>"
                    "{} {}</span></td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#D04A02'>"
                    "{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#4A4A4A'>"
                    "{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#7D7D7D'>"
                    "{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#2D2D2D'>"
                    "{}</td>"
                    "</tr>".format(
                        color,icon,cn,
                        rs["Quotations"],rs["Vendors"],
                        rs["Services"],rs["Subcategories"]))
            tbl.append("</tbody></table>")
            st.markdown("".join(tbl),
                         unsafe_allow_html=True)

        # Drill-down
        st.markdown("<br>", unsafe_allow_html=True)
        sec("DRILL-DOWN BY CATEGORY")
        ordered = (["Cybersecurity"]
                   + [c for c in all_cats_ov
                      if c != "Cybersecurity"])
        cat_tabs = st.tabs([
            "🛡️ {}".format(c)
            if c=="Cybersecurity" else c
            for c in ordered])

        for tidx,cn in enumerate(ordered):
            with cat_tabs[tidx]:
                dc = df_ov[df_ov["Category"]==cn].copy()
                if dc.empty:
                    st.info("No data."); continue
                ck1,ck2,ck3,ck4 = st.columns(4)
                nq  = len(dc); nv = dc["Vendor"].nunique()
                ns  = dc["Subcategory"].nunique()
                nsv = (df_exploded[
                    df_exploded["Category"]==cn
                ]["Service"].nunique()
                        if df_exploded is not None else 0)
                mini_kpi(ck1,nq, "Quotations",C_ORANGE,"📄")
                mini_kpi(ck2,nv, "Vendors",   C_DARK,  "🏢")
                mini_kpi(ck3,nsv,"Services",  C_MID,   "🛠")
                mini_kpi(ck4,ns, "Subcats",   C_GREY_DARK,"🏷️")
                st.markdown("<br>",unsafe_allow_html=True)

                ss = (dc.groupby("Subcategory")
                      .agg(Quotations=("File Name","count"),
                           Vendors=("Vendor","nunique"))
                      .reset_index()
                      .sort_values("Quotations",ascending=False))

                sc1,sc2 = st.columns([1,1],gap="medium")
                with sc1:
                    st.markdown(
                        "<div style='font-size:0.77em;"
                        "font-weight:700;text-transform:"
                        "uppercase;color:#2D2D2D;"
                        "margin-bottom:8px'>"
                        "Subcategories</div>",
                        unsafe_allow_html=True)
                    fs = go.Figure(go.Bar(
                        x=ss["Quotations"],
                        y=ss["Subcategory"],
                        orientation="h",
                        marker_color=C_ORANGE,
                        marker_line_width=0,
                        text=ss["Quotations"],
                        textposition="outside"))
                    fs.update_layout(
                        height=max(200,len(ss)*36),
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=8,b=8),
                        font=CFONT,
                        xaxis=dict(showgrid=True,
                                   gridcolor=C_GREY_LITE,
                                   zeroline=False),
                        yaxis=dict(autorange="reversed",
                                   tickfont=dict(size=9.5)),
                        bargap=0.3)
                    st.plotly_chart(fs,use_container_width=True)
                with sc2:
                    st.markdown(
                        "<div style='font-size:0.77em;"
                        "font-weight:700;text-transform:"
                        "uppercase;color:#2D2D2D;"
                        "margin-bottom:8px'>"
                        "Vendors</div>",
                        unsafe_allow_html=True)
                    vs2 = (dc.groupby("Vendor")
                           .agg(Quotations=(
                               "File Name","count"))
                           .reset_index()
                           .sort_values("Quotations",
                                         ascending=False))
                    fv2 = go.Figure(go.Bar(
                        x=vs2["Quotations"],
                        y=vs2["Vendor"],
                        orientation="h",
                        marker_color=C_DARK,
                        marker_line_width=0,
                        text=vs2["Quotations"],
                        textposition="outside"))
                    fv2.update_layout(
                        height=max(200,len(vs2)*36),
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=8,b=8),
                        font=CFONT,
                        xaxis=dict(showgrid=True,
                                   gridcolor=C_GREY_LITE,
                                   zeroline=False),
                        yaxis=dict(autorange="reversed",
                                   tickfont=dict(size=9.5)),
                        bargap=0.3)
                    st.plotly_chart(fv2,use_container_width=True)

                st.markdown(
                    "<div style='font-size:0.77em;"
                    "font-weight:700;text-transform:uppercase;"
                    "color:#2D2D2D;margin:12px 0 8px'>"
                    "📄 All {} files</div>".format(len(dc)),
                    unsafe_allow_html=True)
                ft = ["<table class='comp-table'>"
                      "<thead><tr><th>File Name</th>"
                      "<th>Vendor</th><th>Subcategory</th>"
                      "<th>Services</th>"
                      "</tr></thead><tbody>"]
                for fi,(_,fr) in enumerate(
                        dc.sort_values("Subcategory"
                                        ).iterrows()):
                    bg  = "white" if fi%2==0 else "#F8F8F8"
                    vc  = vendor_color_map.get(fr["Vendor"],C_DARK)
                    cmt = str(fr.get("Comments","")
                               ).replace("\n"," · ")[:80]
                    ft.append(
                        "<tr style='background:{}'>"
                        "<td style='font-family:monospace;"
                        "font-size:0.77em;word-break:"
                        "break-all'>{}</td>"
                        "<td>{}</td>"
                        "<td style='color:#7D7D7D;"
                        "font-size:0.81em'>{}</td>"
                        "<td style='font-size:0.79em'>"
                        "{}</td></tr>".format(
                            bg,fr.get("File Name",""),
                            vpill(fr["Vendor"],vc),
                            fr["Subcategory"],cmt))
                ft.append("</tbody></table>")
                st.markdown("".join(ft),
                             unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# TAB 1 — BROWSE & VERDICT
# Left = Chatbot | Right = Manual browser
# ════════════════════════════════════════════════════════════
with tab1:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header(
            "Browse & Verdict",
            "Chat with the assistant · "
            "or manually browse services for deep analysis")

        # ── Top filter bar ──
        st.markdown("<div class='filter-bar'>",
                     unsafe_allow_html=True)
        fb1,fb2,fb3 = st.columns(3)
        all_cats_f = (["All"] + sorted([
            c for c in df_master["Category"].unique()
            if str(c).strip() not in ["","nan"]]))
        with fb1:
            st.markdown(
                "<p style='color:#D04A02;font-size:"
                "0.77em;font-weight:700;margin-bottom:"
                "4px;letter-spacing:0.6px;"
                "text-transform:uppercase'>"
                "📂 CATEGORY</p>",
                unsafe_allow_html=True)
            sel_cat = st.selectbox(
                "cat",all_cats_f,
                label_visibility="collapsed",
                key="bv_cat")
        vpool = (df_master if sel_cat=="All"
                  else df_master[
                      df_master["Category"]==sel_cat])
        all_vend_f = (["All"] + sorted([
            v for v in vpool["Vendor"].unique()
            if str(v).strip() not in ["","nan"]]))
        with fb2:
            st.markdown(
                "<p style='color:#D04A02;font-size:"
                "0.77em;font-weight:700;margin-bottom:"
                "4px;letter-spacing:0.6px;"
                "text-transform:uppercase'>"
                "🏢 VENDOR</p>",
                unsafe_allow_html=True)
            sel_vendor = st.selectbox(
                "ven",all_vend_f,
                label_visibility="collapsed",
                key="bv_ven")
        with fb3:
            st.markdown(
                "<p style='color:#D04A02;font-size:"
                "0.77em;font-weight:700;margin-bottom:"
                "4px;letter-spacing:0.6px;"
                "text-transform:uppercase'>"
                "🔍 SEARCH SERVICE</p>",
                unsafe_allow_html=True)
            svc_filter = st.text_input(
                "svc",
                placeholder="Type to filter…",
                label_visibility="collapsed",
                key="bv_svc_search")
        st.markdown("</div>",unsafe_allow_html=True)

        # Apply filters
        d_filt_bv = df_exploded.copy()
        if sel_cat    != "All":
            d_filt_bv = d_filt_bv[
                d_filt_bv["Category"]==sel_cat]
        if sel_vendor != "All":
            d_filt_bv = d_filt_bv[
                d_filt_bv["Vendor"]==sel_vendor]
        avail_svcs = sorted([
            s for s in d_filt_bv["Service"].unique()
            if str(s).strip() not in ["","nan"]])
        if svc_filter:
            avail_svcs = [s for s in avail_svcs
                           if svc_filter.lower()
                           in s.lower()]

        # ── Side-by-side ──
        left_col, right_col = st.columns(
            [1,1], gap="large")

        # ══════════════════════════
        # LEFT — CHATBOT
        # ══════════════════════════
        with left_col:
            # Header
            st.markdown(
                "<div class='chat-header'>"
                "<span style='font-weight:700;"
                "font-size:0.96em;"
                "font-family:Georgia,serif'>"
                "💬 Procurement Assistant</span>"
                "<span style='font-size:0.74em;"
                "opacity:0.6;margin-left:10px'>"
                "Full catalog knowledge · "
                "Upload a quote for instant scoring"
                "</span></div>",
                unsafe_allow_html=True)

            # Chat history window
            chat_html = "<div class='chat-outer'>"
            if not st.session_state["chat_history"]:
                chat_html += (
                    "<div class='chat-wrap'>"
                    "<div class='msg-bot'>"
                    "👋 Hello! I have full knowledge "
                    "of the master catalog.<br><br>"
                    "<b>Try asking:</b><br>"
                    "• <i>Who quoted for "
                    "Cisco Catalyst?</i><br>"
                    "• <i>Compare Palo Alto prices</i>"
                    "<br>"
                    "• <i>Which vendor is cheapest "
                    "for Cybersecurity?</i><br>"
                    "• <i>What does TrendMicro offer?</i>"
                    "<br><br>"
                    "Or upload a quote file below "
                    "for instant scoring."
                    "</div></div>")
            else:
                for turn in st.session_state[
                        "chat_history"]:
                    chat_html += (
                        "<div class='chat-wrap'>"
                        "<div class='msg-user'>{}</div>"
                        "</div>".format(turn["user"]))
                    bot = (turn["bot_text"]
                           .replace("\n","<br>"))
                    # bold handling
                    import re as _re
                    bot = _re.sub(
                        r'\*\*(.+?)\*\*',
                        r'<b>\1</b>', bot)
                    chat_html += (
                        "<div class='chat-wrap'>"
                        "<div class='msg-bot'>{}</div>"
                        "</div>".format(bot))

            # Quick-question chips INSIDE chat window
            chips = [
                "Who quoted Cisco Catalyst?",
                "Compare Palo Alto prices",
                "Cheapest Cybersecurity vendor?",
                "What does TrendMicro offer?",
                "List all vendors",
                "Show catalog summary",
            ]
            chat_html += (
                "<div class='chip-row'>"
                + "".join(
                    "<span class='chip' "
                    "title='{}'>{}</span>".format(
                        c, c[:28]+"…"
                        if len(c)>28 else c)
                    for c in chips)
                + "</div>")
            chat_html += "</div>"  # close chat-outer
            st.markdown(chat_html,
                         unsafe_allow_html=True)

            # Render last chart
            if st.session_state["chat_history"]:
                last_resp = st.session_state[
                    "chat_history"][-1].get("bot_resp")
                if last_resp:
                    render_chat_chart(last_resp)

            # File upload
            chat_file = st.file_uploader(
                "📎 Upload quote file for instant scoring",
                type=["pdf","xlsx","xls","docx"],
                key="chat_file_upload",
                label_visibility="visible")

            # Text input + send
            st.markdown(
                "<div class='chat-input-bar'>",
                unsafe_allow_html=True)
            with st.form("chat_form",
                          clear_on_submit=True):
                ci1,ci2 = st.columns([5,1])
                with ci1:
                    user_input = st.text_input(
                        "msg",
                        placeholder=
                        "Ask a question or click a chip above…",
                        label_visibility="collapsed")
                with ci2:
                    sent = st.form_submit_button(
                        "Send", type="primary",
                        use_container_width=True)
            st.markdown("</div>",unsafe_allow_html=True)

            # Chip buttons (functional) — below chat
            st.markdown(
                "<div style='font-size:0.72em;"
                "color:#7D7D7D;font-weight:600;"
                "letter-spacing:0.6px;"
                "margin:6px 0 4px'>"
                "CLICK A QUICK QUESTION:</div>",
                unsafe_allow_html=True)

            chip_cols = st.columns(3)
            chip_pairs = [
                ("Who quoted Cisco Catalyst?", 0),
                ("Compare Palo Alto prices",   1),
                ("Cheapest Cybersecurity?",    2),
                ("TrendMicro profile",         0),
                ("List all vendors",           1),
                ("Catalog summary",            2),
            ]
            for chip_txt, col_idx in chip_pairs:
                if chip_cols[col_idx].button(
                        chip_txt,
                        key="chip_{}".format(chip_txt),
                        use_container_width=True):
                    resp = chatbot_response(
                        chip_txt,
                        df_master, df_exploded)
                    st.session_state[
                        "chat_history"].append({
                        "user":     chip_txt,
                        "bot_text": resp["text"],
                        "bot_resp": resp,
                    })
                    st.rerun()

            # Handle send
            if sent and user_input.strip():
                resp = chatbot_response(
                    user_input.strip(),
                    df_master, df_exploded)
                st.session_state[
                    "chat_history"].append({
                    "user":     user_input.strip(),
                    "bot_text": resp["text"],
                    "bot_resp": resp,
                })
                st.rerun()

            # Handle file upload
            if (chat_file is not None
                    and st.session_state.get(
                        "last_chat_file")
                    != chat_file.name):
                st.session_state[
                    "last_chat_file"] = chat_file.name
                file_bytes = chat_file.read()
                resp = chatbot_response(
                    "uploaded file",
                    df_master, df_exploded,
                    uploaded_file_bytes=file_bytes,
                    uploaded_file_name=chat_file.name)
                st.session_state[
                    "chat_history"].append({
                    "user": "📎 {}".format(
                        chat_file.name),
                    "bot_text": resp["text"],
                    "bot_resp": resp,
                })
                st.rerun()

            if st.button("🗑 Clear chat",
                          key="clr_chat"):
                st.session_state["chat_history"] = []
                st.rerun()

        # ══════════════════════════
        # RIGHT — MANUAL BROWSER
        # ══════════════════════════
        with right_col:
            sec("SERVICE BROWSER")
            selected_svcs = st.multiselect(
                "Select services to analyse",
                options=avail_svcs,
                default=[],
                label_visibility="visible")

            if not selected_svcs:
                # All-services overview
                sec("SERVICE COMPETITIVENESS MAP",
                    "Orange = multiple vendors "
                    "· Grey = single vendor only")
                svc_sum = (
                    df_exploded.groupby("Service")[
                        "Vendor"].nunique()
                    .reset_index()
                    .sort_values("Vendor",
                                  ascending=False))
                svc_sum.columns = ["Service",
                                    "Vendor Count"]
                top20 = svc_sum.head(20)
                fig_sv = go.Figure(go.Bar(
                    x=top20["Vendor Count"],
                    y=top20["Service"].apply(
                        lambda x: x[:48]),
                    orientation="h",
                    marker_color=[
                        C_ORANGE if v>1 else C_GREY
                        for v in top20["Vendor Count"]],
                    marker_line_width=0,
                    text=top20["Vendor Count"],
                    textposition="outside",
                    textfont=dict(size=10)))
                fig_sv.update_layout(
                    height=520,
                    plot_bgcolor=CBG,
                    paper_bgcolor=CBG,
                    margin=dict(l=5,r=40,t=10,b=8),
                    font=CFONT,
                    xaxis=dict(
                        title="Vendors",
                        showgrid=True,
                        gridcolor=C_GREY_LITE,
                        zeroline=False),
                    yaxis=dict(
                        autorange="reversed",
                        tickfont=dict(size=9.2)),
                    bargap=0.28,showlegend=False)
                st.plotly_chart(fig_sv,
                                 use_container_width=True)

                n_multi  = svc_sum[
                    svc_sum["Vendor Count"]>1
                ].shape[0]
                n_single = svc_sum[
                    svc_sum["Vendor Count"]==1
                ].shape[0]
                sm1,sm2,sm3 = st.columns(3)
                kpi_box(sm1, len(svc_sum),
                         "Total Services", C_DARK)
                kpi_box(sm2, n_multi,
                         "Competitive (2+ vendors)",
                         C_ORANGE)
                kpi_box(sm3, n_single,
                         "Single Vendor Only",
                         C_GREY_DARK)

            else:
                # Deep analysis
                d_sel = (d_filt_bv[
                    d_filt_bv["Service"].isin(
                        selected_svcs)].copy())
                if d_sel.empty:
                    st.warning(
                        "No quotations found.")
                else:
                    has_price = (
                        "Quoted Price" in d_sel.columns)

                    # Collect vendor prices
                    vpm = {}
                    for _, r in d_sel.drop_duplicates(
                            subset=["Vendor","File Name"]
                    ).iterrows():
                        v  = r["Vendor"]
                        qp = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        ck = "px_{}".format(str(r.get(
                            "File Name","")).strip())
                        ca = st.session_state.get(ck)
                        ep = ca["price_num"] if ca else 0.0
                        ref = ep if ep>0 else qp
                        if ref>0:
                            vpm[v] = min(
                                vpm.get(v,ref),ref)
                    if not vpm:
                        for _,r in d_sel.drop_duplicates(
                                subset=["Vendor"]
                        ).iterrows():
                            v  = r["Vendor"]
                            qp = _parse_num(str(r.get(
                                "Quoted Price","")).strip())
                            if qp>0 and v not in vpm:
                                vpm[v] = qp

                    # Coverage verdict
                    vsmap = defaultdict(set)
                    for _,r in d_sel.iterrows():
                        vsmap[r["Vendor"]].add(
                            r["Service"])
                    full_cov = [
                        v for v,s in vsmap.items()
                        if set(selected_svcs).issubset(s)]
                    partial = [
                        v for v,s in vsmap.items()
                        if set(selected_svcs) & s
                        and v not in full_cov]

                    if len(selected_svcs) == 1:
                        nv = d_sel["Vendor"].nunique()
                        if nv > 1:
                            st.markdown(
                                "<div class='verdict-good'>"
                                "✅ <b>{}</b> — quoted by "
                                "<b>{} vendors</b>. "
                                "Competitive benchmarking "
                                "available.</div>".format(
                                    selected_svcs[0], nv),
                                unsafe_allow_html=True)
                        else:
                            st.markdown(
                                "<div class='verdict-mid'>"
                                "⚠️ <b>{}</b> — only "
                                "<b>1 vendor</b> quoted. "
                                "Limited benchmarking."
                                "</div>".format(
                                    selected_svcs[0]),
                                unsafe_allow_html=True)
                    else:
                        if full_cov:
                            st.markdown(
                                "<div class='verdict-good'>"
                                "✅ <b>{}</b> vendor(s) "
                                "cover ALL {} selected: "
                                "<b>{}</b></div>".format(
                                    len(full_cov),
                                    len(selected_svcs),
                                    ", ".join(full_cov)),
                                unsafe_allow_html=True)
                        else:
                            st.markdown(
                                "<div class='verdict-mid'>"
                                "⚠️ No single vendor covers "
                                "all {} services. Partial: "
                                "<b>{}</b></div>".format(
                                    len(selected_svcs),
                                    ", ".join(partial)
                                    if partial else "None"),
                                unsafe_allow_html=True)
                    st.markdown("<br>",
                                 unsafe_allow_html=True)

                    # Price verdict
                    if vpm:
                        avg_p  = sum(vpm.values()
                                     )/len(vpm)
                        best_v = min(vpm,key=vpm.get)
                        worst_v= max(vpm,key=vpm.get)
                        spread = round(
                            (max(vpm.values())
                             -min(vpm.values()))
                            /min(vpm.values())*100,1
                        ) if min(vpm.values())>0 else 0

                        sec("PRICE VERDICT")
                        pv1,pv2,pv3 = st.columns(3)
                        pv1.markdown(
                            "<div class='scard "
                            "scard-orange'>"
                            "<div style='font-size:0.67em;"
                            "font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#D04A02'>"
                            "Best Price</div>"
                            "<div style='font-size:1.5em;"
                            "font-weight:800;color:#D04A02;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:0.73em;"
                            "color:#7D7D7D;margin-top:2px'>"
                            "{}</div></div>".format(
                                _fmt(min(vpm.values())),
                                best_v),
                            unsafe_allow_html=True)
                        pv2.markdown(
                            "<div class='scard scard-dark'>"
                            "<div style='font-size:0.67em;"
                            "font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#2D2D2D'>"
                            "Market Average</div>"
                            "<div style='font-size:1.5em;"
                            "font-weight:800;color:#2D2D2D;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:0.73em;"
                            "color:#7D7D7D;margin-top:2px'>"
                            "{} vendors</div>"
                            "</div>".format(
                                _fmt(avg_p),len(vpm)),
                            unsafe_allow_html=True)
                        pv3.markdown(
                            "<div class='scard scard-grey'>"
                            "<div style='font-size:0.67em;"
                            "font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#7D7D7D'>"
                            "Price Spread</div>"
                            "<div style='font-size:1.5em;"
                            "font-weight:800;color:#4A4A4A;"
                            "font-family:Georgia,serif'>"
                            "{}%</div>"
                            "<div style='font-size:0.73em;"
                            "color:#7D7D7D;margin-top:2px'>"
                            "negotiation room</div>"
                            "</div>".format(spread),
                            unsafe_allow_html=True)
                        st.markdown("<br>",
                                     unsafe_allow_html=True)

                        # Price chart
                        sec("PRICE COMPARISON",
                            "Orange = best · Dark = highest"
                            " · Dashed = market average")
                        sv = sorted(vpm.items(),
                                     key=lambda x:x[1])
                        bc = [C_ORANGE if v==best_v
                               else C_DARK if v==worst_v
                               else C_GREY_DARK
                               for v,_ in sv]
                        fig_c = go.Figure(go.Bar(
                            x=[v for v,_ in sv],
                            y=[p for _,p in sv],
                            marker_color=bc,
                            marker_line_width=0,
                            text=[_fmt(p) for _,p in sv],
                            textposition="outside",
                            textfont=dict(size=11,
                                          color=C_DARK)))
                        fig_c.add_hline(
                            y=avg_p,
                            line_dash="dash",
                            line_color=C_MID,
                            line_width=2,
                            annotation_text=
                            "Avg: {}".format(_fmt(avg_p)),
                            annotation_position=
                            "top right")
                        fig_c.update_layout(
                            height=290,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(l=5,r=10,
                                        t=24,b=8),
                            font=CFONT,
                            yaxis=dict(
                                title="Price (USD)",
                                showgrid=True,
                                gridcolor=C_GREY_LITE,
                                zeroline=False),
                            xaxis=dict(
                                tickangle=-10,
                                tickfont=dict(size=10.5)),
                            bargap=0.4,showlegend=False)
                        st.plotly_chart(
                            fig_c,
                            use_container_width=True)

                        # Score table
                        sec("VENDOR SCORE CARD")
                        all_pv = list(vpm.values())
                        tbl = [
                            "<table class='comp-table'>"
                            "<thead><tr><th>Rank</th>"
                            "<th>Vendor</th>"
                            "<th>Price</th>"
                            "<th>vs Average</th>"
                            "<th>Score</th>"
                            "<th>Verdict</th>"
                            "</tr></thead><tbody>"]
                        for rank,(v,p) in enumerate(
                                sorted(vpm.items(),
                                       key=lambda x:x[1]),
                                1):
                            bg  = ("white" if rank%2==0
                                    else "#F8F8F8")
                            vc  = vendor_color_map.get(
                                v,C_DARK)
                            oth = [x for x in all_pv
                                   if x != p]
                            ps  = None
                            if p>0 and oth:
                                ps,_,_,_,_ = price_score(
                                    p,oth)
                            sc  = score_color(ps)
                            pct = round(
                                (p-avg_p)/avg_p*100,1
                            ) if avg_p>0 else 0
                            vs  = (
                                "{}% below ✅".format(
                                    abs(pct))
                                if pct<0
                                else "{}% above ⚠️".format(
                                    abs(pct))
                                if pct>0
                                else "At average")
                            vc2 = (C_ORANGE if pct<0
                                    else C_DARK
                                    if pct>10
                                    else C_GREY_DARK)
                            vt,_,_ = get_verdict(ps)
                            medal = (
                                "🥇" if rank==1
                                else "🥈" if rank==2
                                else "🥉" if rank==3
                                else str(rank))
                            tbl.append(
                                "<tr style='background:{}'>"
                                "<td style='text-align:"
                                "center;font-size:1.05em'>"
                                "{}</td>"
                                "<td>{}</td>"
                                "<td style='font-family:"
                                "monospace;font-weight:700'>"
                                "{}</td>"
                                "<td style='color:{}'>"
                                "{}</td>"
                                "<td style='text-align:"
                                "center'>"
                                "<span style='font-weight:"
                                "800;font-size:1.1em;"
                                "color:{}'>{}</span>"
                                "</td>"
                                "<td style='font-weight:700;"
                                "color:{}'>{}</td>"
                                "</tr>".format(
                                    bg,medal,
                                    vpill(v,vc),
                                    _fmt(p),
                                    vc2,vs,sc,
                                    ps if ps is not None
                                    else "—",
                                    sc,vt))
                        tbl.append("</tbody></table>")
                        st.markdown("".join(tbl),
                                     unsafe_allow_html=True)

                        # Insight
                        pct_b = round(
                            (avg_p-min(vpm.values()))
                            /avg_p*100,1
                        ) if avg_p>0 else 0
                        insight(
                            "<b>{}</b> is most competitive "
                            "at <b>{}</b> — "
                            "<b>{}% below</b> market avg. "
                            "Spread of <b>{}%</b> → "
                            "<b>{}</b>.".format(
                                best_v,
                                _fmt(min(vpm.values())),
                                pct_b, spread,
                                "strong negotiation "
                                "potential"
                                if spread>20
                                else "moderate room"
                                if spread>10
                                else "competitive market"))

                    # Per-service file detail
                    st.markdown("<br>",
                                 unsafe_allow_html=True)
                    sec("QUOTATION FILE DETAILS")
                    for svc in selected_svcs:
                        d_svc = (
                            d_sel[d_sel["Service"]==svc]
                            .drop_duplicates(
                                subset=["Vendor",
                                         "File Name"])
                            .sort_values("Vendor"))
                        nv2 = d_svc["Vendor"].nunique()
                        st.markdown(
                            "<div style='background:white;"
                            "border-left:4px solid {};"
                            "padding:10px 14px;"
                            "border-radius:2px;"
                            "margin:8px 0;"
                            "font-weight:700;"
                            "font-size:0.88em'>"
                            "{}  ·  {} vendor(s)  ·  {}"
                            "</div>".format(
                                C_ORANGE if nv2>1
                                else C_GREY_DARK,
                                svc, nv2,
                                "✅ COMPETITIVE"
                                if nv2>1
                                else "⚠️ SINGLE VENDOR"),
                            unsafe_allow_html=True)

                        all_p2 = []
                        for _,r in d_svc.iterrows():
                            qp = _parse_num(str(r.get(
                                "Quoted Price","")).strip())
                            if qp>0: all_p2.append(qp)

                        rt = [
                            "<table class='comp-table'>"
                            "<thead><tr>"
                            "<th>Vendor</th>"
                            "<th>File</th>"]
                        if has_price:
                            rt.append(
                                "<th>Quoted Price</th>")
                        rt.append(
                            "<th>Score</th>"
                            "<th>Verdict</th>"
                            "<th>Open</th>"
                            "</tr></thead><tbody>")

                        for i,(_,row) in enumerate(
                                d_svc.iterrows()):
                            bg2   = ("white" if i%2==0
                                      else "#F8F8F8")
                            vc3   = vendor_color_map.get(
                                row["Vendor"],C_DARK)
                            fname = str(row.get(
                                "File Name","")).strip()
                            url   = resolve_url(row)
                            qpn   = _parse_num(str(row.get(
                                "Quoted Price","")).strip())
                            ck    = "px_{}".format(fname)
                            ca    = st.session_state.get(ck)
                            ref   = (ca["price_num"]
                                      if ca and ca.get(
                                          "price_num",0)>0
                                      else qpn
                                      if qpn>0 else 0)
                            oth   = [p for p in all_p2
                                     if p!=ref]
                            ps2   = None; vt2="—"
                            vc4   = C_GREY_DARK
                            if ref>0 and oth:
                                ps2,_,_,_,_ = price_score(
                                    ref,oth)
                                vt2,_,vc4 = get_verdict(ps2)
                            sc2 = score_color(ps2)
                            lnk = (
                                "<a href='{}' "
                                "target='_blank' "
                                "style='color:#D04A02;"
                                "font-weight:600;"
                                "text-decoration:none'>"
                                "📂 Open</a>".format(url)
                                if url else "—")
                            rt.append(
                                "<tr style='background:{}'>"
                                "<td>{}</td>"
                                "<td style='font-family:"
                                "monospace;font-size:0.77em;"
                                "word-break:break-all'>"
                                "{}</td>".format(
                                    bg2,
                                    vpill(row["Vendor"],vc3),
                                    fname))
                            if has_price:
                                rt.append(
                                    "<td style='font-family:"
                                    "monospace;font-weight:"
                                    "700;color:#D04A02'>{}"
                                    "</td>".format(
                                        _fmt(qpn)
                                        if qpn>0 else "—"))
                            rt.append(
                                "<td style='text-align:"
                                "center'>"
                                "<span style='font-weight:"
                                "800;color:{}'>{}</span>"
                                "</td>"
                                "<td style='font-weight:700;"
                                "color:{}'>{}</td>"
                                "<td>{}</td>"
                                "</tr>".format(
                                    sc2,
                                    "{}/100".format(ps2)
                                    if ps2 is not None
                                    else "—",
                                    vc4,vt2,lnk))
                        rt.append("</tbody></table>")
                        st.markdown("".join(rt),
                                     unsafe_allow_html=True)

                        st.markdown("<br>",
                                     unsafe_allow_html=True)
                        if st.button(
                                "🔍 Extract Prices — "
                                "{}".format(svc[:38]),
                                key="ep_{}".format(
                                    svc[:32]),
                                type="primary"):
                            prog = st.progress(0)
                            nn   = len(d_svc)
                            for ki,(_,r2) in enumerate(
                                    d_svc.iterrows()):
                                f2  = str(r2.get(
                                    "File Name","")).strip()
                                ck2 = "px_{}".format(f2)
                                if not st.session_state.get(
                                        ck2):
                                    loc = os.path.join(
                                        DEMO_DIR,f2)
                                    if os.path.exists(loc):
                                        st.session_state[
                                            ck2] = extract_price_from_file(
                                            loc)
                                    else:
                                        u2 = resolve_url(r2)
                                        if (u2
                                                and u2.startswith("http")
                                                and REQUESTS_OK):
                                            try:
                                                rr = requests.get(
                                                    u2,timeout=20)
                                                ee = u2.split(
                                                    "?"
                                                )[0].rsplit(
                                                    ".",1
                                                )[-1].lower()
                                                st.session_state[
                                                    ck2] = extract_price_from_bytes(
                                                    rr.content,ee)
                                            except: pass
                                prog.progress((ki+1)/nn)
                            prog.empty(); st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab2:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header(
            "Upload & Score",
            "Upload a quote → auto-extract price "
            "→ compare vs history → get verdict")

        prefill_price = st.session_state.get(
            "tab2_upload_price", 0.0)
        prefill_fname = st.session_state.get(
            "tab2_upload_fname", "")
        prefill_bytes = st.session_state.get(
            "tab2_file_bytes", None)

        if prefill_fname:
            st.markdown(
                "<div class='insight-box'>"
                "📎 File from chat: <b>{}</b> — "
                "Extracted price: <b>{}</b>"
                "</div>".format(
                    prefill_fname,
                    _fmt(prefill_price)
                    if prefill_price>0
                    else "not found"),
                unsafe_allow_html=True)

        sec("STEP 1 — UPLOAD QUOTE FILE")
        uploaded = st.file_uploader(
            "Upload",
            type=["pdf","xlsx","xls","docx"],
            label_visibility="collapsed")

        new_price = 0.0; fname_up = ""
        if uploaded is not None:
            content  = uploaded.read()
            ext_up = uploaded.name.rsplit(".", 1)[-1]
            fname_up = uploaded.name
            st.success("Uploaded: **{}** ({} KB)".format(
                fname_up,
                round(len(content)/1024,1)))
            sec("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting…"):
                res       = extract_price_from_bytes(
                    content,ext_up)
                new_price = res["price_num"]
            if new_price>0:
                st.markdown(
                    "<div class='scard scard-orange'>"
                    "<div style='font-size:0.70em;"
                    "font-weight:700;"
                    "text-transform:uppercase;"
                    "color:#D04A02'>"
                    "Extracted Price</div>"
                    "<div style='font-size:2.1em;"
                    "font-weight:800;color:#D04A02;"
                    "font-family:Georgia,serif'>"
                    "{}</div></div>".format(
                        _fmt(new_price)),
                    unsafe_allow_html=True)
            else:
                st.warning(
                    "Price not found automatically.")
                manual = st.number_input(
                    "Enter price manually (USD)",
                    min_value=0.0,step=100.0,
                    value=0.0,key="manual_price")
                if manual>0: new_price = manual
        elif prefill_price>0:
            new_price = prefill_price
            fname_up  = prefill_fname

        sec("STEP 3 — SELECT SERVICES & FILTERS")
        up1,up2,up3 = st.columns(3)
        with up1:
            cat_up = st.selectbox(
                "📂 Filter by Category",
                ["All"]+sorted([
                    c for c in df_master[
                        "Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="cat_up2")
        with up2:
            svcs_up = sorted([
                s for s in df_exploded[
                    "Service"].unique()
                if str(s).strip()
                not in ["","nan"]])
            svc_srch = st.text_input(
                "🔍 Filter services",
                placeholder="Search…",
                key="svc_srch_up")
            if svc_srch:
                svcs_up = [s for s in svcs_up
                            if svc_srch.lower()
                            in s.lower()]
        with up3:
            new_svcs = st.multiselect(
                "🛠 Select Services",
                options=svcs_up,
                key="new_svcs2")

        sec("STEP 4 — COMPARISON & VERDICT")
        if new_price<=0 and not new_svcs:
            st.info(
                "Upload a file and select services "
                "to compare.")
        else:
            cands = (
                df_exploded[
                    df_exploded["Service"].isin(
                        new_svcs)].copy()
                if new_svcs else df_exploded.copy())
            if cat_up != "All":
                cands = cands[
                    cands["Category"]==cat_up]
            cf = (cands.drop_duplicates(
                subset=["File Name","Vendor"])
                  [["File Name","Vendor","Category",
                    "Hyperlink","Quoted Price"]].copy())

            if cf.empty:
                st.warning("No historical quotes found.")
            else:
                hist = []; vpm2 = {}
                for _,r in cf.iterrows():
                    qp = _parse_num(str(r.get(
                        "Quoted Price","")).strip())
                    if qp>0:
                        hist.append(qp)
                        vpm2[r["Vendor"]] = qp
                    ck = "px_{}".format(str(r.get(
                        "File Name","")).strip())
                    ca = st.session_state.get(ck)
                    if ca and ca.get("price_num",0)>0:
                        hist.append(ca["price_num"])
                        vpm2[r["Vendor"]] = ca["price_num"]

                if new_price>0 and hist:
                    ps,lbl,avh,mnh,mxh = price_score(
                        new_price,hist)
                    vt,vd,vc5 = get_verdict(ps)
                    css = ("orange" if (ps or 0)>=70
                            else "mid" if (ps or 0)>=40
                            else "dark")
                    bgs = {"orange":"#FFF5F0",
                            "mid":"#F5F5F5",
                            "dark":"#F0F0F0"}
                    bds = {"orange":C_ORANGE,
                            "mid":C_GREY_DARK,
                            "dark":C_DARK}
                    st.markdown(
                        "<div style='background:{};"
                        "border:2px solid {};"
                        "border-radius:4px;"
                        "padding:16px 20px;"
                        "margin-bottom:16px'>"
                        "<div style='font-size:1.2em;"
                        "font-weight:700;color:{};"
                        "font-family:Georgia,serif'>"
                        "{}</div>"
                        "<div style='font-size:0.87em;"
                        "color:#4A4A4A;margin-top:5px'>"
                        "{}</div></div>".format(
                            bgs[css],bds[css],
                            bds[css],vt,vd),
                        unsafe_allow_html=True)

                    sv1,sv2,sv3,sv4 = st.columns(4)
                    sv1.markdown(
                        "<div class='scard scard-orange'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#D04A02'>Score</div>"
                        "<div style='font-size:2.1em;"
                        "font-weight:800;color:#D04A02;"
                        "font-family:Georgia,serif'>"
                        "{}/100</div>"
                        "<div style='font-size:0.73em;"
                        "color:#7D7D7D;margin-top:3px'>"
                        "vs {} historical</div>"
                        "</div>".format(
                            ps if ps is not None
                            else "N/A",len(hist)),
                        unsafe_allow_html=True)
                    sv2.markdown(
                        "<div class='scard scard-dark'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#2D2D2D'>"
                        "Your Price</div>"
                        "<div style='font-size:2.1em;"
                        "font-weight:800;color:#D04A02;"
                        "font-family:Georgia,serif'>"
                        "{}</div></div>".format(
                            _fmt(new_price)),
                        unsafe_allow_html=True)
                    sv3.markdown(
                        "<div class='scard scard-dark'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#2D2D2D'>"
                        "Market Average</div>"
                        "<div style='font-size:2.1em;"
                        "font-weight:800;color:#4A4A4A;"
                        "font-family:Georgia,serif'>"
                        "{}</div>"
                        "<div style='font-size:0.72em;"
                        "color:#7D7D7D;margin-top:3px'>"
                        "min {} · max {}</div>"
                        "</div>".format(
                            _fmt(avh),
                            _fmt(mnh),_fmt(mxh)),
                        unsafe_allow_html=True)
                    sv4.markdown(
                        "<div class='scard scard-grey'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#7D7D7D'>"
                        "vs Average</div>"
                        "<div style='font-size:0.95em;"
                        "font-weight:800;color:#4A4A4A;"
                        "margin-top:8px'>"
                        "{}</div></div>".format(lbl),
                        unsafe_allow_html=True)

                    st.markdown("<br>",
                                 unsafe_allow_html=True)
                    sec("PRICE POSITIONING CHART",
                        "Orange = your quote · "
                        "Dark = historical · "
                        "Dashed = market average")
                    cd2 = []
                    for _,r in cf.iterrows():
                        fn  = str(r.get(
                            "File Name","")).strip()
                        qp2 = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        ck  = "px_{}".format(fn)
                        ca  = st.session_state.get(ck)
                        ep  = ca["price_num"] if ca else 0.0
                        pv  = ep if ep>0 else qp2
                        if pv>0:
                            cd2.append({
                                "Label":"{}/{}".format(
                                    r["Vendor"],fn[:10]),
                                "Price":pv,
                                "Type":"Historical"})
                    cd2.append({
                        "Label":"★ YOUR QUOTE",
                        "Price":new_price,
                        "Type":"New"})
                    cdf2 = pd.DataFrame(cd2
                        ).sort_values("Price")
                    bc2  = [C_ORANGE
                             if t=="New" else C_DARK
                             for t in cdf2["Type"]]
                    fig_up = go.Figure(go.Bar(
                        x=cdf2["Label"],
                        y=cdf2["Price"],
                        marker_color=bc2,
                        marker_line_width=0,
                        text=cdf2["Price"].apply(_fmt),
                        textposition="outside"))
                    fig_up.add_hline(
                        y=avh,line_dash="dash",
                        line_color=C_GREY_DARK,
                        line_width=2,
                        annotation_text=
                        "Avg: {}".format(_fmt(avh)),
                        annotation_position="top right")
                    fig_up.update_layout(
                        height=380,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,
                                    t=20,b=10),
                        font=CFONT,
                        yaxis=dict(
                            title="Price (USD)",
                            showgrid=True,
                            gridcolor=C_GREY_LITE,
                            zeroline=False),
                        xaxis=dict(tickangle=-25),
                        bargap=0.3,showlegend=False)
                    st.plotly_chart(
                        fig_up,use_container_width=True)

                    pct_vs = round(
                        (new_price-avh)/avh*100,1
                    ) if avh>0 else 0
                    insight(
                        "Your quote of <b>{}</b> is "
                        "<b>{}% {}</b> the market avg "
                        "of <b>{}</b>. Range: "
                        "<b>{}</b> – <b>{}</b>.".format(
                            _fmt(new_price),
                            abs(pct_vs),
                            "below" if pct_vs<0
                            else "above",
                            _fmt(avh),
                            _fmt(mnh),_fmt(mxh)))
                else:
                    st.info(
                        "No historical price data. "
                        "Try selecting more services.")


# ════════════════════════════════════════════════════════════
# TAB 3 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab3:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header("Data Table",
                   "Full catalog — filterable")
        tf1,tf2,tf3 = st.columns(3)
        with tf1:
            dt_cat = st.selectbox(
                "📂 Category",
                ["All"]+sorted([
                    c for c in df_master[
                        "Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="dt_cat")

        with tf2:
            vp_dt = (df_master if dt_cat=="All"
                      else df_master[
                          df_master["Category"]==dt_cat])
            dt_ven = st.selectbox(
                "🏢 Vendor",
                ["All"]+sorted([
                    v for v in vp_dt["Vendor"].unique()
                    if str(v).strip()
                    not in ["","nan"]]),
                key="dt_ven")
        with tf3:
            dt_srch = st.text_input(
                "🔍 Search",
                placeholder="File name or comments…",
                key="dt_srch")

        dm = df_master.copy()
        if dt_cat  != "All":
            dm = dm[dm["Category"]==dt_cat]
        if dt_ven  != "All":
            dm = dm[dm["Vendor"]==dt_ven]
        if dt_srch:
            mask = (
                dm["File Name"].str.contains(
                    dt_srch,case=False,na=False)
                | dm["Comments"].str.contains(
                    dt_srch,case=False,na=False))
            dm = dm[mask]

        st.markdown(
            "<div style='font-size:0.82em;"
            "color:#7D7D7D;margin:8px 0'>"
            "Showing <b>{}</b> of <b>{}</b> "
            "records</div>".format(
                len(dm),len(df_master)),
            unsafe_allow_html=True)
        st.dataframe(
            dm.drop(
                columns=["Services List","Hyperlink"],
                errors="ignore"),
            use_container_width=True,
            height=520)


# ════════════════════════════════════════════════════════════
# TAB 4 — UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab4:
    pwc_header(
        "Upload Master Catalog",
        "Upload Excel or CSV — "
        "AI auto-detects columns")

    if DATA_SOURCE=="uploaded" and not NO_DATA:
        st.success(
            "✅ Using uploaded catalog: "
            "**{}** rows · **{}** vendors · "
            "**{}** services".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique()))
    elif not NO_DATA:
        insight(
            "Currently using: <b>Master Catalog.xlsx"
            "</b> — {} vendors · {} services · "
            "{} categories".format(
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique(),
                df_master["Category"].nunique()))

    if DATA_SOURCE=="uploaded":
        if st.button("🔄 Reset to Default Catalog",
                      type="primary"):
            st.session_state[
                "uploaded_catalog_df"]  = None
            st.session_state[
                "uploaded_catalog_exp"] = None
            st.rerun()

    sec("UPLOAD A NEW CATALOG")
    cat_file = st.file_uploader(
        "Upload",
        type=["xlsx","xls","csv"],
        label_visibility="collapsed",
        key="catalog_upload")

    if cat_file is not None:
        fb = cat_file.read()
        fn = cat_file.name
        with st.spinner("Analysing…"):
            df_n,dfe_n,err = process_uploaded_catalog(
                fb,fn)
        if err:
            st.error("❌ {}".format(err))
        elif df_n is None:
            st.error("❌ Could not process file.")
        else:
            st.success(
                "✅ **{}** rows · **{}** vendors · "
                "**{}** categories".format(
                    len(df_n),
                    df_n["Vendor"].nunique(),
                    df_n["Category"].nunique()))
            pc1,pc2 = st.columns(2,gap="large")
            spv = (dfe_n.groupby("Vendor")[
                "Service"].nunique()
                    .sort_values(ascending=False)
                    .reset_index())
            spv.columns = ["Vendor","Services"]
            with pc1:
                pf1 = go.Figure(go.Bar(
                    x=spv["Vendor"],
                    y=spv["Services"],
                    marker_color=C_ORANGE,
                    marker_line_width=0,
                    text=spv["Services"],
                    textposition="outside"))
                pwc_bar(pf1,
                         "Services per Vendor",
                         height=300)
                pf1.update_xaxes(
                    tickangle=-30,
                    tickfont=dict(size=9.5))
                st.plotly_chart(
                    pf1,use_container_width=True)
            cn2 = (df_n.drop_duplicates(
                subset=["Category","File Name"])
                    .groupby("Category").size()
                    .reset_index())
            cn2.columns = ["Category","Count"]
            with pc2:
                if not cn2.empty:

                    pf2 = px.pie(
                        cn2,
                        names="Category",
                        values="Count",
                        hole=0.45,
                        color_discrete_sequence=
                        CHART_SEQ)
                    pf2.update_traces(
                        textposition="outside",
                        textinfo="label+percent",
                        textfont_size=10)
                    pf2.update_layout(
                        title=dict(
                            text="Category Distribution",
                            font=dict(
                                size=12,color=C_DARK,
                                family="Georgia,serif"),
                            x=0,xanchor="left"),
                        height=300,
                        margin=dict(l=10,r=10,
                                    t=40,b=10),
                        paper_bgcolor=CBG,
                        font=CFONT)
                    st.plotly_chart(
                        pf2,
                        use_container_width=True)

            st.dataframe(
                df_n.drop(
                    columns=["Services List",
                              "Hyperlink"],
                    errors="ignore").head(20),
                use_container_width=True,
                height=280)

            if st.button(
                    "✅ Apply This Catalog "
                    "to Dashboard",
                    type="primary"):
                st.session_state[
                    "uploaded_catalog_df"]  = df_n
                st.session_state[
                    "uploaded_catalog_exp"] = dfe_n
                st.success("✅ Applied!")
                st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 5 — VENDOR ANALYSIS
# ════════════════════════════════════════════════════════════
with tab5:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header(
            "Vendor Price Analysis",
            "Prices from GitHub · "
            "Per-service benchmarking")

        GITHUB_RAW = (
            "https://raw.githubusercontent.com/"
            "avijeet528/vendor-draft-2/main/"
            "demo_quotes/{}")

        @st.cache_data(show_spinner=False)
        def _gh_price(filename):
            try:
                resp = requests.get(
                    GITHUB_RAW.format(filename),
                    timeout=30)
                if resp.status_code != 200:
                    return {"price":"",
                            "price_num":0.0,
                            "status":"Not found"}
                res = extract_price_from_bytes(
                    resp.content,
                    filename.rsplit(".",1)[-1].lower())
                res["status"] = (
                    "✅ Extracted"
                    if res["price_num"]>0
                    else "⚠️ No price")
                return res
            except Exception as e:
                return {"price":"","price_num":0.0,
                        "status":"❌ {}".format(
                            str(e)[:40])}

        @st.cache_data(show_spinner=False)
        def _gh_all(fnames_tuple):
            return {f:_gh_price(f)
                    for f in fnames_tuple}

        all_fnames = tuple(
            str(f).strip()
            for f in df_master["File Name"].unique()
            if str(f).strip() not in ["","nan"])

        # Filters + load button
        vf1,vf2,vf3 = st.columns([2,1,1])
        with vf1:
            run_gh = st.button(
                "🔄 Load All Prices from GitHub",
                type="primary",
                use_container_width=True,
                key="run_gh")
        with vf2:
            va_cat = st.selectbox(
                "📂 Category",
                ["All"]+sorted([
                    c for c in df_master[
                        "Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="va_cat")
        with vf3:
            vap = (df_master if va_cat=="All"
                    else df_master[
                        df_master["Category"]==va_cat])
            va_ven = st.selectbox(
                "🏢 Vendor",
                ["All"]+sorted([
                    v for v in vap["Vendor"].unique()
                    if str(v).strip()
                    not in ["","nan"]]),
                key="va_ven")

        insight(
            "Will read <b>{}</b> files from GitHub. "
            "Catalog prices used as fallback.".format(
                len(all_fnames)))

        if run_gh:
            st.session_state[
                "gh_prices_loaded"] = True
            _gh_all.clear()
            _gh_price.clear()

        if st.session_state.get(
                "gh_prices_loaded",False):
            with st.spinner("Loading prices…"):
                gh = _gh_all(all_fnames)

            dm5 = df_master.copy()
            if va_cat != "All":
                dm5 = dm5[dm5["Category"]==va_cat]
            if va_ven != "All":
                dm5 = dm5[dm5["Vendor"]==va_ven]

            rows5 = []
            for _,r in dm5.iterrows():
                fn  = str(r.get(
                    "File Name","")).strip()
                qp  = _parse_num(str(r.get(
                    "Quoted Price","")).strip())
                gi  = gh.get(fn,{})
                gp  = gi.get("price_num",0.0)
                bp  = gp if gp>0 else qp
                src = ("extracted" if gp>0
                        else "catalog" if qp>0
                        else "none")
                rows5.append({
                    "Vendor":   r["Vendor"],
                    "Category": r["Category"],
                    "File Name":fn,
                    "Quoted Price":qp,
                    "Extracted Price":gp,
                    "Best Price":bp,
                    "Source":src,
                    "Status":gi.get("status","—"),
                    "Services":r.get("Comments","")})
            df5 = pd.DataFrame(rows5)
            df5 = df5[df5["Best Price"]>0]

            if df5.empty:
                st.warning("No prices found.")
            else:
                sec("SUMMARY")
                k1,k2,k3,k4 = st.columns(4)
                kpi_box(k1,len(df5),
                         "Files with Prices",C_ORANGE)
                kpi_box(k2,
                         len(df5[df5["Source"]==
                                  "extracted"]),
                         "Extracted",C_DARK)
                kpi_box(k3,
                         len(df5[df5["Source"]==
                                  "catalog"]),
                         "From Catalog",C_MID)
                kpi_box(k4,
                         _fmt(df5["Best Price"].mean()),
                         "Avg Quote",C_GREY_DARK)
                st.markdown("<br>",
                             unsafe_allow_html=True)

                vt5 = (df5.groupby("Vendor")[
                    "Best Price"]
                        .agg(["mean","sum",
                               "min","max","count"])
                        .reset_index())
                vt5.columns = [
                    "Vendor","Average","Total",
                    "Min","Max","Quotes"]
                vt5 = vt5.sort_values("Average")
                oa5 = df5["Best Price"].mean()

                sec("VENDOR COMPARISON",
                    "Orange = cheapest · "
                    "Dark = most expensive · "
                    "Dashed = overall average")
                bc5 = [C_ORANGE if i==0
                        else C_DARK
                        if i==len(vt5)-1
                        else C_GREY_DARK
                        for i in range(len(vt5))]
                fig5 = go.Figure(go.Bar(
                    x=vt5["Vendor"],
                    y=vt5["Average"],
                    marker_color=bc5,
                    marker_line_width=0,
                    text=vt5["Average"].apply(_fmt),
                    textposition="outside"))
                fig5.add_hline(
                    y=oa5,line_dash="dash",
                    line_color=C_MID,line_width=2,
                    annotation_text=
                    "Avg: {}".format(_fmt(oa5)),
                    annotation_position="top right")
                pwc_bar(fig5,
                         "Average Quote per Vendor",
                         height=360)
                fig5.update_xaxes(tickangle=-20)
                st.plotly_chart(
                    fig5,use_container_width=True)

                sec("VENDOR RANKING TABLE")
                vtbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>Rank</th><th>Vendor</th>"
                    "<th>Quotes</th><th>Avg</th>"
                    "<th>Min</th><th>Max</th>"
                    "<th>vs Avg</th><th>Verdict</th>"
                    "</tr></thead><tbody>"]
                for rank,(_,vr) in enumerate(
                        vt5.iterrows(),start=1):
                    bg  = ("white" if rank%2==0
                            else "#F8F8F8")
                    vc  = vendor_color_map.get(
                        vr["Vendor"],C_DARK)
                    pct = round(
                        (vr["Average"]-oa5)/oa5*100,1
                    ) if oa5>0 else 0
                    pc  = (C_ORANGE if pct<-5
                            else C_DARK if pct>5
                            else C_GREY_DARK)
                    pt  = ("{}% below".format(abs(pct))
                            if pct<0
                            else "{}% above".format(
                                abs(pct))
                            if pct>0 else "At avg")
                    ov  = (("✅ COMPETITIVE",C_ORANGE)
                            if pct<-10
                            else ("🔴 EXPENSIVE",C_DARK)
                            if pct>10
                            else ("🟡 AVERAGE",
                                   C_GREY_DARK))
                    medal = ("🥇" if rank==1
                              else "🥈" if rank==2
                              else "🥉" if rank==3
                              else str(rank))
                    vtbl.append(
                        "<tr style='background:{}'>"
                        "<td style='text-align:center;"
                        "font-size:1.05em'>{}</td>"
                        "<td>{}</td>"
                        "<td style='text-align:center;"
                        "font-weight:700'>{}</td>"
                        "<td style='font-family:"
                        "monospace;font-weight:700;"
                        "color:#D04A02'>{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#D04A02'>"
                        "{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#2D2D2D'>"
                        "{}</td>"
                        "<td style='color:{}'>{}</td>"
                        "<td style='color:{};"
                        "font-weight:700'>{}</td>"
                        "</tr>".format(
                            bg,medal,
                            vpill(vr["Vendor"],vc),
                            int(vr["Quotes"]),
                            _fmt(vr["Average"]),
                            _fmt(vr["Min"]),
                            _fmt(vr["Max"]),
                            pc,pt,ov[1],ov[0]))
                vtbl.append("</tbody></table>")
                st.markdown("".join(vtbl),
                             unsafe_allow_html=True)

                # Per-service benchmarking
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("PER-SERVICE BENCHMARKING",
                    "Services with multiple vendor "
                    "quotes — full comparison")
                svc_rows5 = []
                for _,r in df5.iterrows():
                    raw = str(r["Services"]
                               ).replace(
                        "\\n","\n"
                    ).replace("\r\n","\n"
                               ).replace("\r","\n")
                    ss  = [s.strip()
                            for s in raw.split("\n")
                            if s.strip()
                            and s.strip()
                            not in ["nan","None",""]]
                    if not ss: ss = [raw.strip()]
                    for s in ss:
                        svc_rows5.append({
                            "Service":s,
                            "Vendor":r["Vendor"],
                            "Price":r["Best Price"],
                            "Source":r["Source"]})
                df_sv5 = pd.DataFrame(svc_rows5)
                svc_vc5 = (df_sv5.groupby(
                    "Service")["Vendor"].nunique())
                multi5  = (svc_vc5[svc_vc5>1]
                            .index.tolist())

                if not multi5:
                    st.info(
                        "No services with multiple "
                        "vendor quotes in current filter.")
                else:
                    insight(
                        "<b>{}</b> service(s) have "
                        "multi-vendor quotes.".format(
                            len(multi5)))
                    for svc5 in sorted(multi5):
                        ds5  = df_sv5[
                            df_sv5["Service"]==svc5
                        ].sort_values("Price")
                        mn5  = ds5["Price"].min()
                        mx5  = ds5["Price"].max()
                        av5  = ds5["Price"].mean()
                        bv5  = ds5.loc[
                            ds5["Price"].idxmin(),
                            "Vendor"]
                        wv5  = ds5.loc[
                            ds5["Price"].idxmax(),
                            "Vendor"]
                        sp5  = round(
                            (mx5-mn5)/mn5*100,1
                        ) if mn5>0 else 0

                        st.markdown(
                            "<div style='background:"
                            "white;border-left:4px "
                            "solid {};padding:10px 14px;"
                            "border-radius:2px;"
                            "margin:10px 0;"
                            "font-weight:700;"
                            "font-size:0.88em'>"
                            "{} · {} vendors · "
                            "spread {}% · best: {} @ {}"
                            "</div>".format(
                                C_ORANGE,svc5,
                                ds5["Vendor"].nunique(),
                                sp5,bv5,_fmt(mn5)),
                            unsafe_allow_html=True)

                        sc5a,sc5b,sc5c = st.columns(3)
                        sc5a.markdown(
                            "<div class='scard "
                            "scard-orange'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#D04A02'>"
                            "Cheapest</div>"
                            "<div style='font-size:"
                            "1.4em;font-weight:800;"
                            "color:#D04A02;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.73em;color:#7D7D7D'>"
                            "{}</div></div>".format(
                                _fmt(mn5),bv5),
                            unsafe_allow_html=True)
                        sc5b.markdown(
                            "<div class='scard scard-dark'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#2D2D2D'>"
                            "Average</div>"
                            "<div style='font-size:"
                            "1.4em;font-weight:800;"
                            "color:#2D2D2D;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.73em;color:#7D7D7D'>"
                            "{} vendors</div>"
                            "</div>".format(
                                _fmt(av5),
                                ds5["Vendor"].nunique()),
                            unsafe_allow_html=True)
                        sc5c.markdown(
                            "<div class='scard scard-grey'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#7D7D7D'>"
                            "Most Expensive</div>"
                            "<div style='font-size:"
                            "1.4em;font-weight:800;"
                            "color:#4A4A4A;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.73em;color:#7D7D7D'>"
                            "{}</div></div>".format(
                                _fmt(mx5),wv5),
                            unsafe_allow_html=True)

                        bc5s = [
                            C_ORANGE if v==bv5
                            else C_DARK if v==wv5
                            else C_GREY_DARK
                            for v in ds5["Vendor"]]
                        fig5s = go.Figure(go.Bar(
                            x=ds5["Vendor"],
                            y=ds5["Price"],
                            marker_color=bc5s,
                            marker_line_width=0,
                            text=ds5["Price"].apply(
                                _fmt),
                            textposition="outside"))
                        fig5s.add_hline(
                            y=av5,line_dash="dash",
                            line_color=C_MID,
                            line_width=1.5,
                            annotation_text=
                            "Avg: {}".format(_fmt(av5)),
                            annotation_position=
                            "top right")
                        fig5s.update_layout(
                            height=240,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(l=5,r=10,
                                        t=12,b=8),
                            font=CFONT,
                            yaxis=dict(
                                showgrid=True,
                                gridcolor=C_GREY_LITE,
                                zeroline=False),
                            bargap=0.4,
                            showlegend=False)
                        st.plotly_chart(
                            fig5s,
                            use_container_width=True)

                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("DOWNLOAD")
                st.download_button(
                    "📥 Download Analysis CSV",
                    data=df5[[
                        "Vendor","Category",
                        "File Name","Best Price",
                        "Source"]].to_csv(index=False),
                    file_name="vendor_analysis.csv",
                    mime="text/csv",
                    type="primary")
        else:
            st.markdown(
                "<div style='background:white;"
                "border:1px solid #E0E0E0;"
                "border-radius:6px;"
                "padding:32px;text-align:center;"
                "margin-top:24px'>"
                "<div style='font-size:2.2em;"
                "margin-bottom:12px'>🔍</div>"
                "<div style='font-size:1.1em;"
                "font-weight:700;color:#2D2D2D;"
                "margin-bottom:8px;"
                "font-family:Georgia,serif'>"
                "Click the button above to start"
                "</div>"
                "<div style='font-size:0.84em;"
                "color:#7D7D7D'>"
                "Fetches <b>{}</b> files from GitHub "
                "and extracts prices."
                "</div></div>".format(len(all_fnames)),
                unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# TAB 6 — REAL ANALYSIS
# ════════════════════════════════════════════════════════════
with tab6:
    pwc_header(
        "Real Quotation Analysis",
        "Upload catalog → follow file links → "
        "extract prices → full vendor analysis")

    sec("STEP 1 — UPLOAD MASTER CATALOG")
    insight(
        "Your Excel must have: <b>Category</b> · "
        "<b>Vendor</b> · <b>File Name</b> · "
        "<b>Quoted Price</b> · <b>Comments</b> · "
        "<b>File Link</b> (URL to each quote file)")

    up_real = st.file_uploader(
        "Upload",
        type=["xlsx","xls","csv"],
        key="tab6_real",
        label_visibility="collapsed")

    df_rc = None
    if up_real is None:
        if not NO_DATA and not df_master.empty:
            use_cur = st.checkbox(
                "Use currently loaded catalog "
                "({} files)".format(len(df_master)),
                key="use_cur_real")
            if use_cur:
                df_rc = df_master.copy()
                st.success(
                    "Using current catalog — "
                    "**{}** vendors · **{}** files"
                    .format(
                        df_rc["Vendor"].nunique(),
                        len(df_rc)))
            else:
                st.info(
                    "👆 Upload Master Catalog "
                    "to start.")
        else:
            st.info("👆 Upload Master Catalog.")
    else:
        rb = up_real.read()
        df_rc,_,re2 = process_uploaded_catalog(
            rb, up_real.name)
        if re2 or df_rc is None:
            st.error("❌ {}".format(re2))
            df_rc = None
        else:
            st.success(
                "✅ **{}** vendors · **{}** files · "
                "**{}** categories".format(
                    df_rc["Vendor"].nunique(),
                    len(df_rc),
                    df_rc["Category"].nunique()))

    if df_rc is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        sec("STEP 2 — RUN ANALYSIS")
        run_real = st.button(
            "🚀 Run Full Analysis",
            type="primary",
            key="run_real")

        if run_real:
            st.session_state[
                "real_analysis_done"] = False
            st.session_state[
                "real_analysis_df"]   = None

        if (run_real or st.session_state.get(
                "real_analysis_done", False)):
            if run_real:
                results_r = []
                prog_r    = st.progress(0)
                tot_r     = len(df_rc)
                stxt_r    = st.empty()

                for i,(_, row) in enumerate(
                        df_rc.iterrows()):
                    fn  = str(row.get(
                        "File Name","")).strip()
                    ven = str(row.get(
                        "Vendor","")).strip()
                    cat = str(row.get(
                        "Category","")).strip()
                    qp  = _parse_num(str(row.get(
                        "Quoted Price","")).strip())
                    url = ""
                    for col in [
                            "Hyperlink","File Link",
                            "File URL","URL"]:
                        val = str(row.get(
                            col,"")).strip()
                        if (val
                                and val not in
                                ["","nan","None"]
                                and val.startswith(
                                    "http")):
                            url = val; break

                    stxt_r.markdown(
                        "<div style='font-size:0.80em;"
                        "color:#7D7D7D'>"
                        "Processing {}/{}: "
                        "<b>{}</b></div>".format(
                            i+1, tot_r, fn),
                        unsafe_allow_html=True)

                    if url and REQUESTS_OK:
                        try:
                            resp_r = requests.get(
                                url, timeout=30,
                                headers={
                                    "User-Agent":
                                    "Mozilla/5.0"})
                            if resp_r.status_code==200:
                                ext_r = (
                                    fn.rsplit(".",1)[-1]
                                    .lower()
                                    if "." in fn
                                    else "xlsx")
                                er = extract_price_from_bytes(
                                    resp_r.content,
                                    ext_r)
                                ep  = er["price_num"]
                                sts = (
                                    "✅ Extracted"
                                    if ep>0
                                    else "⚠️ No price")
                            else:
                                ep  = 0.0
                                sts = "❌ HTTP {}".format(
                                    resp_r.status_code)
                        except Exception as ex:
                            ep  = 0.0
                            sts = "❌ {}".format(
                                str(ex)[:40])
                    else:
                        ep  = 0.0
                        sts = "⚪ No link"

                    bp  = ep if ep>0 else qp
                    src = ("extracted" if ep>0
                            else "catalog" if qp>0
                            else "none")
                    results_r.append({
                        "Vendor":   ven,
                        "Category": cat,
                        "File Name":fn,
                        "Quoted Price":qp,
                        "Extracted Price":ep,
                        "Best Price":bp,
                        "Source":   src,
                        "Status":   sts})
                    prog_r.progress((i+1)/tot_r)

                prog_r.empty()
                stxt_r.empty()
                df_res = pd.DataFrame(results_r)
                st.session_state[
                    "real_analysis_done"] = True
                st.session_state[
                    "real_analysis_df"] = (
                    df_res.to_dict("records"))
            else:
                df_res = pd.DataFrame(
                    st.session_state[
                        "real_analysis_df"])

            df_pr = df_res[
                df_res["Best Price"]>0].copy()

            # KPIs
            rk1,rk2,rk3,rk4,rk5 = st.columns(5)
            kpi_box(rk1, len(df_res),
                     "Total Files",  C_ORANGE)
            kpi_box(rk2, len(df_pr),
                     "Prices Found", C_DARK)
            kpi_box(rk3,
                     len(df_res)-len(df_pr),
                     "No Price",     C_GREY_DARK)
            kpi_box(rk4,
                     df_res["Vendor"].nunique(),
                     "Vendors",      C_MID)
            kpi_box(rk5,
                     _fmt(df_pr["Best Price"].mean())
                     if not df_pr.empty else "—",
                     "Avg Quote",    C_BLACK)

            if df_pr.empty:
                st.warning(
                    "No prices found. "
                    "Check File Link URLs.")
            else:
                # Vendor ranking chart
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("VENDOR RANKING — "
                    "CHEAPEST TO MOST EXPENSIVE",
                    "Orange = cheapest · "
                    "Dark = most expensive · "
                    "Dashed = market average")
                vs_r = (df_pr.groupby("Vendor")[
                    "Best Price"]
                         .agg(["mean","sum","min",
                                "max","count"])
                         .reset_index())
                vs_r.columns = [
                    "Vendor","Average","Total",
                    "Min","Max","Quotes"]
                vs_r = vs_r.sort_values("Average")
                oa_r = df_pr["Best Price"].mean()

                bc_r = [
                    C_ORANGE if i==0
                    else C_DARK
                    if i==len(vs_r)-1
                    else C_GREY_DARK
                    for i in range(len(vs_r))]
                fig_r = go.Figure(go.Bar(
                    x=vs_r["Vendor"],
                    y=vs_r["Average"],
                    marker_color=bc_r,
                    marker_line_width=0,
                    text=vs_r["Average"].apply(_fmt),
                    textposition="outside"))
                fig_r.add_hline(
                    y=oa_r,
                    line_dash="dash",
                    line_color=C_MID,
                    line_width=2,
                    annotation_text=
                    "Avg: {}".format(_fmt(oa_r)),
                    annotation_position="top right")
                pwc_bar(fig_r,
                         "Average Quote per Vendor",
                         height=380)
                fig_r.update_xaxes(tickangle=-20)
                st.plotly_chart(
                    fig_r, use_container_width=True)

                # Vendor ranking table
                sec("VENDOR RANKING TABLE")
                rtbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>Rank</th><th>Vendor</th>"
                    "<th>Quotes</th><th>Avg</th>"
                    "<th>Min</th><th>Max</th>"
                    "<th>vs Avg</th><th>Verdict</th>"
                    "</tr></thead><tbody>"]
                for rnk,(_,vr) in enumerate(
                        vs_r.iterrows(), start=1):
                    bg_r = ("white" if rnk%2==0
                             else "#F8F8F8")
                    vc_r = vendor_color_map.get(
                        vr["Vendor"], C_DARK)
                    pct_r = round(
                        (vr["Average"]-oa_r)
                        /oa_r*100, 1
                    ) if oa_r>0 else 0
                    pc_r = (C_ORANGE if pct_r<-5
                             else C_DARK if pct_r>5
                             else C_GREY_DARK)
                    pt_r = (
                        "{}% below".format(abs(pct_r))
                        if pct_r<0
                        else "{}% above".format(
                            abs(pct_r))
                        if pct_r>0
                        else "At avg")
                    ov_r = (
                        ("✅ COMPETITIVE", C_ORANGE)
                        if pct_r<-10
                        else ("🔴 EXPENSIVE", C_DARK)
                        if pct_r>10
                        else ("🟡 AVERAGE",
                               C_GREY_DARK))
                    md_r = (
                        "🥇" if rnk==1
                        else "🥈" if rnk==2
                        else "🥉" if rnk==3
                        else str(rnk))
                    rtbl.append(
                        "<tr style='background:{}'>"
                        "<td style='text-align:center;"
                        "font-size:1.05em'>{}</td>"
                        "<td>{}</td>"
                        "<td style='text-align:center;"
                        "font-weight:700'>{}</td>"
                        "<td style='font-family:"
                        "monospace;font-weight:700;"
                        "color:#D04A02'>{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#D04A02'>"
                        "{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#2D2D2D'>"
                        "{}</td>"
                        "<td style='color:{}'>{}</td>"
                        "<td style='color:{};"
                        "font-weight:700'>{}</td>"
                        "</tr>".format(
                            bg_r, md_r,
                            vpill(vr["Vendor"],vc_r),
                            int(vr["Quotes"]),
                            _fmt(vr["Average"]),
                            _fmt(vr["Min"]),
                            _fmt(vr["Max"]),
                            pc_r, pt_r,
                            ov_r[1], ov_r[0]))
                rtbl.append("</tbody></table>")
                st.markdown("".join(rtbl),
                             unsafe_allow_html=True)

                # File extraction status
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("FILE EXTRACTION STATUS")
                stbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>File Name</th>"
                    "<th>Vendor</th>"
                    "<th>Quoted</th>"
                    "<th>Extracted</th>"
                    "<th>Used</th>"
                    "<th>Source</th>"
                    "<th>Status</th>"
                    "</tr></thead><tbody>"]
                for ii, rr in df_res.iterrows():
                    bg_s = ("white" if ii%2==0
                             else "#F8F8F8")
                    vc_s = vendor_color_map.get(
                        rr["Vendor"], C_DARK)
                    sc_s = (
                        C_ORANGE
                        if rr["Source"]=="extracted"
                        else C_GREY_DARK
                        if rr["Source"]=="catalog"
                        else C_GREY)
                    stbl.append(
                        "<tr style='background:{}'>"
                        "<td style='font-family:"
                        "monospace;font-size:0.76em;"
                        "word-break:break-all'>"
                        "{}</td>"
                        "<td>{}</td>"
                        "<td style='font-family:"
                        "monospace'>{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#D04A02'>"
                        "{}</td>"
                        "<td style='font-family:"
                        "monospace;font-weight:700;"
                        "color:#D04A02'>{}</td>"
                        "<td><span style='background:"
                        "{};color:white;padding:"
                        "2px 6px;border-radius:2px;"
                        "font-size:0.72em;"
                        "font-weight:700'>"
                        "{}</span></td>"
                        "<td style='font-size:"
                        "0.79em'>{}</td>"
                        "</tr>".format(
                            bg_s,
                            rr["File Name"],
                            vpill(rr["Vendor"],vc_s),
                            _fmt(rr["Quoted Price"])
                            if rr["Quoted Price"]>0
                            else "—",
                            _fmt(rr["Extracted Price"])
                            if rr["Extracted Price"]>0
                            else "—",
                            _fmt(rr["Best Price"])
                            if rr["Best Price"]>0
                            else "—",
                            sc_s,
                            rr["Source"].upper(),
                            rr["Status"]))
                stbl.append("</tbody></table>")
                st.markdown("".join(stbl),
                             unsafe_allow_html=True)

                # Download
                st.markdown("<br>",
                             unsafe_allow_html=True)
                st.download_button(
                    "📥 Download Analysis CSV",
                    data=df_pr[[
                        "Vendor","Category",
                        "File Name",
                        "Quoted Price",
                        "Extracted Price",
                        "Best Price",
                        "Source",
                        "Status",
                    ]].to_csv(index=False),
                    file_name="real_analysis.csv",
                    mime="text/csv",
                    type="primary")
