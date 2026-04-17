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
            "font-size:0.84
