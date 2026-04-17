# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  Reads Master Catalog.xlsx (Excel first, CSV fallback)
#  PwC Brand | Source Sans Pro | Streamlit
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

# ════════════════════════════════════════════════════════════
# PAGE CONFIG
# ════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="IT Procurement Intelligence",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# PwC COLOUR PALETTE
# ════════════════════════════════════════════════════════════
PwC_ORANGE   = "#D04A02"
PwC_RED      = "#E0301E"
PwC_YELLOW   = "#FFB600"
PwC_GREEN    = "#22992E"
PwC_NAVY     = "#003078"
PwC_BLUE     = "#295477"
PwC_TEAL     = "#299D8F"
PwC_PURPLE   = "#6E2585"
PwC_DARK     = "#2D2D2D"
PwC_GREY     = "#8C8C8C"
PwC_LTGREY   = "#F3F3F3"
PwC_AMBER    = "#EB8C00"

COLORS = [PwC_ORANGE, PwC_BLUE, PwC_TEAL, PwC_YELLOW,
          PwC_GREEN, PwC_RED, PwC_AMBER, PwC_PURPLE,
          PwC_GREY, PwC_NAVY]

CAT_COLORS = [PwC_ORANGE, PwC_BLUE, PwC_TEAL, PwC_AMBER,
              PwC_PURPLE, PwC_GREEN, PwC_RED, PwC_NAVY,
              PwC_GREY, PwC_YELLOW, PwC_TEAL, PwC_ORANGE]

def get_color(i): return COLORS[i % len(COLORS)]
CFONT = dict(family="Source Sans Pro,Helvetica Neue,Arial", size=11, color=PwC_DARK)
CBG   = PwC_LTGREY
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# CSS  — PwC brand + arrow/expander fixes
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');
html,body,[class*="css"],div,p,span,td,th,label,button,.stMarkdown{
    font-family:'Source Sans Pro','Helvetica Neue',Arial,sans-serif !important;}
.main .block-container{
    background-color:#F3F3F3 !important;
    padding-top:1.5rem;max-width:100% !important;
    padding-left:2rem !important;padding-right:2rem !important;}
#MainMenu{visibility:hidden;}footer{visibility:hidden;}header{visibility:hidden;}
[data-testid="collapsedControl"]{display:none !important;}

/* ── Sidebar ── */
section[data-testid="stSidebar"]{
    background-color:#2D2D2D !important;
    border-right:3px solid #D04A02;
    min-width:300px !important;max-width:300px !important;}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div{color:#F0F0F0 !important;}
section[data-testid="stSidebar"] div[data-baseweb="select"]{
    background-color:#FFFFFF !important;border-radius:2px !important;border:1px solid #999 !important;}
section[data-testid="stSidebar"] div[data-baseweb="select"] *{color:#2D2D2D !important;}
section[data-testid="stSidebar"] span[data-baseweb="tag"]{
    background-color:#D04A02 !important;border-radius:2px !important;}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span{color:white !important;}

/* ── KPI boxes ── */
.kpi-box{border-radius:4px;padding:18px 10px;text-align:center;color:white;
    border-left:5px solid rgba(255,255,255,0.25);}
.kpi-value{font-size:2.1em;font-weight:700;margin:0;line-height:1.1;}
.kpi-label{font-size:0.78em;font-weight:700;opacity:0.9;margin-top:5px;
    letter-spacing:0.8px;text-transform:uppercase;}

/* ── Tabs ── */
button[data-baseweb="tab"]{font-weight:600 !important;font-size:0.92em !important;color:#7D7D7D !important;}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;border-bottom:3px solid #D04A02 !important;}

/* ── Nuclear expander arrow removal ── */
div[data-testid="stExpander"] details summary{
    list-style:none !important;pointer-events:none !important;cursor:default !important;}
div[data-testid="stExpander"] details summary::-webkit-details-marker,
div[data-testid="stExpander"] details summary::marker,
div[data-testid="stExpander"] details summary::before,
div[data-testid="stExpander"] details summary::after{
    display:none !important;content:"" !important;width:0 !important;height:0 !important;}
div[data-testid="stExpander"] details summary svg{display:none !important;}
div[data-testid="stExpander"] details summary button,
div[data-testid="stExpander"] details summary [data-testid="stExpanderToggleIcon"]{display:none !important;}
div[data-testid="stExpander"] details summary div[data-testid="stMarkdownContainer"] p{
    font-weight:700;font-size:0.94em;color:#2D2D2D !important;
    padding-left:0 !important;margin-left:0 !important;cursor:default !important;}
div[data-testid="stExpander"] details{
    border:1px solid #ddd;border-radius:4px;margin-bottom:10px;padding:2px 0;}

/* ── Tables ── */
.comp-table{width:100%;border-collapse:collapse;table-layout:fixed;
    font-size:0.83em;border:1px solid #e0e0e0;}
.comp-table thead tr{background:#2D2D2D;}
.comp-table thead th{padding:10px;text-align:left;font-weight:700;font-size:0.80em;
    letter-spacing:0.4px;text-transform:uppercase;color:white !important;
    border:none;word-break:break-word;}
.comp-table tbody tr:nth-child(even){background:#F3F3F3;}
.comp-table tbody tr:hover{background:#FCE8DC;}
.comp-table tbody td{padding:8px 10px;border-bottom:1px solid #e8e8e8;
    vertical-align:middle;word-break:break-word;font-size:0.82em;color:#2D2D2D;}
.comp-table th:nth-child(1),.comp-table td:nth-child(1){width:13%;}
.comp-table th:nth-child(2),.comp-table td:nth-child(2){width:12%;}
.comp-table th:nth-child(3),.comp-table td:nth-child(3){width:22%;}
.comp-table th:nth-child(4),.comp-table td:nth-child(4){width:10%;}
.comp-table th:nth-child(5),.comp-table td:nth-child(5){width:10%;}
.comp-table th:nth-child(6),.comp-table td:nth-child(6){width:10%;}
.comp-table th:nth-child(7),.comp-table td:nth-child(7){width:13%;}
.comp-table th:nth-child(8),.comp-table td:nth-child(8){width:10%;}

/* ── Vendor badge ── */
.vendor-badge{display:inline-block;padding:3px 8px;border-radius:2px;
    color:white;font-size:0.78em;font-weight:700;white-space:nowrap;
    overflow:hidden;text-overflow:ellipsis;max-width:100%;}

/* ── Score cards ── */
.score-card{border-radius:4px;padding:14px 16px;margin-bottom:10px;border-left:5px solid #D04A02;}
.score-card.green{background:#F0FFF4;border-color:#22992E;}
.score-card.yellow{background:#FFF8E1;border-color:#FFB600;}
.score-card.red{background:#FFF3F0;border-color:#E0301E;}

/* ── AI box ── */
.ai-box{background:#F8F0FF;border-left:5px solid #6E2585;
    border-radius:4px;padding:14px 18px;margin:10px 0;}

/* ── Verdict boxes ── */
.verdict-green{background:#F0FFF4;border:2px solid #22992E;border-radius:4px;
    padding:12px 16px;color:#22992E;font-weight:700;}
.verdict-yellow{background:#FFF8E1;border:2px solid #FFB600;border-radius:4px;
    padding:12px 16px;color:#856404;font-weight:700;}
.verdict-red{background:#FFF3F0;border:2px solid #E0301E;border-radius:4px;
    padding:12px 16px;color:#E0301E;font-weight:700;}

/* ── Demo banner ── */
.demo-banner{background:linear-gradient(135deg,#D04A02 0%,#B83D00 100%);
    color:white;padding:10px 18px;border-radius:4px;
    margin-bottom:16px;font-size:0.88em;font-weight:600;}

/* ── Section heading — no arrow overlap ── */
.sec-heading{display:block;font-size:0.78em;font-weight:700;
    letter-spacing:1px;text-transform:uppercase;
    color:#D04A02;margin:18px 0 8px;line-height:1.4;}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"(?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?"
    r"|(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?",
    re.IGNORECASE)
TOTAL_KW = ["grand total","total amount","total price","amount due",
            "net total","total cost","total value","quote total",
            "subtotal","estimated total","total"]

def _parse_num(s):
    try: return float(re.sub(r"[^\d.]","",str(s)) or "0")
    except: return 0.0

def _fmt(val):
    try:
        v = float(re.sub(r"[^\d.]","",str(val)) or "0")
        if v <= 0: return "—"
        return "${:,.2f}".format(v)
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
    if valid: return max(valid, key=_parse_num)
    return ""

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
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            rows_text = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs = "  ".join(str(c) for c in row if c is not None)
                    if rs.strip(): rows_text.append(rs)
            text = "\n".join(rows_text); wb.close()
        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>"," ",xml)
                    text = re.sub(r"\s{2,}","\n",text)
    except: pass
    return text

def extract_price_from_bytes(content, ext):
    text  = _text_from_bytes(content, ext)
    price = _best_price(text)
    if not price or _parse_num(price) <= 0:
        all_nums = PRICE_RE.findall(text)
        valid = [h.strip() for h in all_nums if _parse_num(h) >= 1000]
        if valid: price = max(valid, key=_parse_num)
    if (not price or _parse_num(price) <= 0) and ext.lower() in ("xlsx","xls"):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
            all_vals = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell is None: continue
                        try:
                            n = float(str(cell).replace(",","").strip())
                            if n >= 1000: all_vals.append(n)
                        except: pass
            wb.close()
            if all_vals: price = str(max(all_vals))
        except: pass
    return {"price": price, "price_num": _parse_num(price) if price else 0.0, "text": text[:5000]}

def extract_price_from_file(filepath):
    try:
        with open(filepath,"rb") as f: content = f.read()
        ext = filepath.rsplit(".",1)[-1].lower()
        return extract_price_from_bytes(content, ext)
    except: return {"price":"","price_num":0.0,"text":""}

# ════════════════════════════════════════════════════════════
# SCORING & VERDICT
# ════════════════════════════════════════════════════════════
def price_score(new_price, hist_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0: return None,"No comparison data",0,0,0
    mn = min(valid); mx = max(valid); avg = sum(valid)/len(valid)
    if mx == mn: return 50,"Same as historical average",avg,mn,mx
    score = round((1-(new_price-mn)/(mx-mn))*100,1)
    score = max(0,min(100,score))
    pct   = round((new_price-avg)/avg*100,1)
    label = ("{}% BELOW average — COMPETITIVE".format(abs(pct)) if new_price < avg
             else "{}% ABOVE average — REVIEW NEEDED".format(abs(pct)) if new_price > avg
             else "Matches historical average")
    return score,label,avg,mn,mx

def score_color(s):
    if s is None: return PwC_GREY
    if s >= 70: return PwC_GREEN
    if s >= 40: return PwC_YELLOW
    return PwC_RED

def score_css(s):
    if s is None: return "yellow"
    if s >= 70: return "green"
    if s >= 40: return "yellow"
    return "red"

def get_verdict(ps):
    if ps is None: return "⚪ No Data","No comparison data available.",PwC_GREY
    if ps >= 70: return "✅ COMPETITIVE","This quote is priced competitively. Proceed with confidence.",PwC_GREEN
    if ps >= 40: return "🟡 AVERAGE","This quote is within average range. Negotiate for a small discount.","#856404"
    return "🔴 HIGH — NEGOTIATE","This quote is above the historical average. Strongly recommend negotiating.",PwC_RED

# ════════════════════════════════════════════════════════════
# AI INSIGHTS
# ════════════════════════════════════════════════════════════
def ai_service_summary(df_master, df_exploded):
    if df_master.empty: return "No vendor data."
    svc_by_v = {v: list(df_exploded[df_exploded["Vendor"]==v]["Service"].unique())
                for v in df_master["Vendor"].unique()}
    if not svc_by_v: return "No vendor data."
    best   = max(svc_by_v, key=lambda v: len(svc_by_v[v]))
    n_best = len(svc_by_v[best])
    total  = len(set(s for svcs in svc_by_v.values() for s in svcs))
    shared = [s for s in set(s for svcs in svc_by_v.values() for s in svcs)
              if sum(1 for svcs in svc_by_v.values() if s in svcs) > 1]
    lines  = ["**{}** covers the most services ({} of {} total).".format(best, n_best, total)]
    if shared: lines.append("**{}** service(s) offered by multiple vendors — ideal for competitive benchmarking.".format(len(shared)))
    return " ".join(lines)

def ai_price_insight(new_price, hist_prices, vendor_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0: return "Insufficient data for price analysis."
    avg = sum(valid)/len(valid)
    pct = round((new_price-avg)/avg*100,1)
    lines = []
    if new_price <= min(valid): lines.append("This quote is the **lowest price** — excellent value.")
    elif new_price >= max(valid): lines.append("This quote is **above all historical prices** — negotiate strongly.")
    elif pct > 15: lines.append("Quote is **{}% above** average. Request a revised quote.".format(abs(pct)))
    elif pct < -15: lines.append("Quote is **{}% below** average — very competitive.".format(abs(pct)))
    else: lines.append("Quote is **within normal range** ({}% vs average).".format(pct))
    if vendor_prices:
        best_v = min(vendor_prices, key=vendor_prices.get)
        lines.append("**{}** has historically offered the lowest prices.".format(best_v))
    return " ".join(lines)

def generate_selection_verdict(selected_svcs, d_sel, vendor_prices_map, df_exploded):
    if d_sel.empty or not selected_svcs: return None
    vsmap = defaultdict(set)
    for _, r in d_sel.iterrows(): vsmap[r["Vendor"]].add(r["Service"])
    full_cover = [v for v,s in vsmap.items() if set(selected_svcs).issubset(s)]
    all_prices = [p for p in vendor_prices_map.values() if p > 0]
    if not all_prices:
        return {"title":"📊 Vendor Coverage","vendors_all":full_cover,
                "vendors_some":sorted([v for v in vsmap if v not in full_cover]),
                "has_prices":False,"lines":[]}
    avg_p  = sum(all_prices)/len(all_prices)
    min_p  = min(all_prices); max_p = max(all_prices)
    best_v = min(vendor_prices_map, key=vendor_prices_map.get)
    worst_v= max(vendor_prices_map, key=vendor_prices_map.get)
    spread = round((max_p-min_p)/min_p*100,1) if min_p>0 else 0
    lines  = ["**{}** vendor(s) quoted for the selected service(s).".format(len(vendor_prices_map)),
              "Price range: **{}** — **{}** (spread: **{}%**).".format(_fmt(min_p),_fmt(max_p),spread),
              "Average quoted price: **{}**.".format(_fmt(avg_p))]
    if full_cover: lines.append("**{}** offer(s) ALL selected services.".format(", ".join(full_cover)))
    else: lines.append("No single vendor covers all selected services — consider multi-vendor approach.")
    lines.append("**Best price:** {} at {} — {}% below average.".format(
        best_v, _fmt(min_p), round((avg_p-min_p)/avg_p*100,1) if avg_p>0 else 0))
    if spread > 20: lines.append("⚠️ Large price spread ({}%) — significant negotiation opportunity.".format(spread))
    return {"title":"📊 Procurement Verdict","vendors_all":full_cover,
            "vendors_some":sorted([v for v in vsmap if v not in full_cover]),
            "has_prices":True,"lines":lines,"best_vendor":best_v,"best_price":min_p,
            "worst_vendor":worst_v,"worst_price":max_p,"avg_price":avg_p,
            "spread":spread,"vendor_prices":vendor_prices_map}

# ════════════════════════════════════════════════════════════
# SUBCATEGORY INFERENCE
# ════════════════════════════════════════════════════════════
def infer_subcategory(category, comments, file_name):
    txt = (str(comments) + " " + str(file_name)).lower()
    cat = str(category).lower().strip()
    if "cybersecurity" in cat:
        if any(k in txt for k in ["trendmicro","trend micro","endpoint","antivirus"]): return "Endpoint Protection"
        if any(k in txt for k in ["cyberark","privileged","pam"]): return "Privileged Access Mgmt"
        if any(k in txt for k in ["knowbe4","awareness","phishing","training"]): return "Security Awareness"
        if any(k in txt for k in ["forescout","nac","network access"]): return "Network Access Control"
        if any(k in txt for k in ["siem","splunk","monitor","log"]): return "SIEM / Monitoring"
        if any(k in txt for k in ["vulnerability","scan","qualys"]): return "Vulnerability Mgmt"
        return "General Security"
    if "network" in cat or "telecom" in cat:
        if any(k in txt for k in ["cisco ise","ise"]): return "Cisco ISE"
        if "meraki" in txt: return "Cisco Meraki"
        if "palo alto" in txt: return "Palo Alto NGFW"
        if "solarwinds" in txt: return "SolarWinds"
        if "equinix" in txt: return "Equinix Interconnect"
        if "cisco" in txt: return "Cisco Networking"
        return "General Network"
    if "hosting" in cat:
        if any(k in txt for k in ["vmware","vcf"]): return "VMware"
        if "oracle" in txt: return "Oracle DB"
        if "netapp" in txt: return "NetApp Storage"
        if any(k in txt for k in ["colocation","colo"]): return "Colocation Build"
        if "windows server" in txt: return "Windows Server"
        if "ibm" in txt: return "IBM Power"
        return "General Hosting"
    if "m365" in cat or "power platform" in cat:
        if any(k in txt for k in ["sharegate","migrate"]): return "Migration Tools"
        if "copilot" in txt: return "Copilot / AI"
        if any(k in txt for k in ["power bi","powerbi"]): return "Power BI"
        if "teams" in txt: return "Teams"
        return "M365 Licensing"
    if "idam" in cat or "iam" in cat:
        return "AD Migration Consulting" if "consulting" in txt else "Identity Migration"
    if "snow" in cat or "servicenow" in cat: return "ServiceNow ITSM"
    if "summary" in cat or "reporting" in cat: return "Reporting & Tracking"
    return str(category).strip().title()

# ════════════════════════════════════════════════════════════
# DATA LOADING  — Excel first, CSV fallback
# ════════════════════════════════════════════════════════════
def _clean_df(df):
    """
    Safely clean a DataFrame:
    1. Keep only plain string/numeric columns
    2. Clean each column with apply(str.strip)
    3. Drop empty rows
    4. Reset index
    """
    # Step 1 — only string-safe columns (no list columns yet)
    safe_cols = [c for c in df.columns if c != "Services List"]
    df = df[safe_cols].copy()

    # Step 2 — clean each column individually using apply
    for col in df.columns:
        try:
            df[col] = df[col].fillna("").apply(lambda x: str(x).strip())
        except Exception:
            df[col] = ""

    # Step 3 — drop empty rows
    mask = (df["Category"].apply(lambda x: x in ["","nan"]) &
            df["Vendor"].apply(lambda x: x in ["","nan"]))
    df = df[~mask].copy()

    # Step 4 — reset
    df.reset_index(drop=True, inplace=True)
    return df

def _parse_services(v):
    if not v or str(v).strip() in ["","nan","None"]:
        return ["(unspecified)"]
    s = str(v).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts = [p.strip() for p in s.split("\n") if p.strip() and p.strip() != "nan"]
    if not parts: parts = [p.strip() for p in s.split(";") if p.strip()]
    if not parts and len(s) < 300: parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts if parts else ["(unspecified)"]

def _explode_services(df):
    df2 = df.copy()
    df2["Services List"] = df2["Comments"].apply(_parse_services)
    df_exp = df2.explode("Services List").copy()
    df_exp.rename(columns={"Services List": "Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].apply(lambda x: str(x).strip())
    df_exp = df_exp[~df_exp["Service"].isin(
        ["","(unspecified)","nan","None"])].reset_index(drop=True)
    return df2, df_exp

def _normalise_columns(df):
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category" and "Category" not in col_map: col_map["Category"] = c
        elif any(k in cl for k in ["vendor","supplier"]) and "Vendor" not in col_map: col_map["Vendor"] = c
        elif ("file name" in cl or cl == "filename") and "File Name" not in col_map: col_map["File Name"] = c
        elif any(k in cl for k in ["file link","file url"]) and "File Link" not in col_map: col_map["File Link"] = c
        elif any(k in cl for k in ["comment","service","description","scope"]) and "Comments" not in col_map: col_map["Comments"] = c
        elif any(k in cl for k in ["price","cost","amount","quoted"]) and "Quoted Price" not in col_map: col_map["Quoted Price"] = c
    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)
    for req in ["Category","Vendor","File Name","Comments"]:
        if req not in df.columns: df[req] = ""
    keep = ["Category","Vendor","File Name","Comments"]
    for e in ["File Link","Quoted Price"]:
        if e in df.columns: keep.append(e)
    return df[[c for c in keep if c in df.columns]].copy()

@st.cache_data
def load_data():
    CSV_PATH = "master_catalog.csv"
    XLS_PATH = "Master Catalog.xlsx"
    df = None

    # Excel first
    if os.path.exists(XLS_PATH):
        try:
            raw = pd.read_excel(XLS_PATH, engine="openpyxl", header=None)
            header_row = 0
            for i, row in raw.iterrows():
                vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if any("category" in v for v in vals) and any("vendor" in v for v in vals):
                    header_row = i; break
            df = pd.read_excel(XLS_PATH, engine="openpyxl", header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("Excel load error: {}".format(e)); df = None

    # CSV fallback
    if df is None and os.path.exists(CSV_PATH):
        try:
            df = pd.read_csv(CSV_PATH)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("CSV load error: {}".format(e)); df = None

    if df is None: return None, None

    df = _normalise_columns(df)
    df = _clean_df(df)

    # Add hyperlink column
    df["Hyperlink"] = ""
    if "File Link" in df.columns:
        df["Hyperlink"] = df["File Link"].apply(lambda x: "" if x in ["","nan"] else x)

    df, df_exp = _explode_services(df)
    return df, df_exp

# ════════════════════════════════════════════════════════════
# PROCESS UPLOADED CATALOG
# ════════════════════════════════════════════════════════════
def process_uploaded_catalog(file_bytes, filename):
    try:
        ext = filename.rsplit(".",1)[-1].lower()
        if ext in ("xlsx","xls"):
            raw = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", header=None)
        elif ext == "csv":
            raw = pd.read_csv(io.BytesIO(file_bytes), header=None)
        else:
            return None, None, "Unsupported file type."

        header_row = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            joined = " ".join(vals)
            if any(k in joined for k in ["vendor","supplier"]) and any(k in joined for k in ["file","document"]):
                header_row = i; break

        if ext in ("xlsx","xls"):
            df = pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl", header=header_row)
        else:
            df = pd.read_csv(io.BytesIO(file_bytes), header=header_row)

        df = df.loc[:, df.columns.notna()]
        df.columns = [str(c).strip() for c in df.columns]
        df.dropna(how="all", inplace=True)
        df = _normalise_columns(df)
        df = _clean_df(df)

        # Extract hyperlinks
        hmap = {}
        if ext in ("xlsx","xls"):
            try:
                wb_h = openpyxl.load_workbook(io.BytesIO(file_bytes))
                ws_h = wb_h.active
                fn_col = None; hdr_row = None
                for row in ws_h.iter_rows():
                    for cell in row:
                        if cell.value and str(cell.value).strip().lower() == "file name":
                            fn_col = cell.column; hdr_row = cell.row; break
                    if fn_col: break
                if fn_col and hdr_row:
                    for row in ws_h.iter_rows(min_row=hdr_row+1):
                        for cell in row:
                            if cell.column == fn_col and cell.value and cell.hyperlink:
                                hmap[str(cell.value).strip()] = str(cell.hyperlink.target).strip()
                wb_h.close()
            except Exception: pass

        df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
        df, df_exp = _explode_services(df)
        return df, df_exp, None
    except Exception as e:
        return None, None, str(e)

# ════════════════════════════════════════════════════════════
# REAL CATALOG ANALYZER
# ════════════════════════════════════════════════════════════
def parse_services_from_cell(val):
    if not val or str(val).strip() in ["","nan","None"]: return []
    s = str(val).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts = [p.strip() for p in s.split("\n") if p.strip() and p.strip() not in ["nan","None"]]
    if not parts: parts = [p.strip() for p in s.split(";") if p.strip()]
    return parts

def get_file_link_from_row(row):
    for col in ["Hyperlink","File Link","File URL","URL","hyperlink","link"]:
        val = str(row.get(col,"")).strip()
        if val and val not in ["","nan","None","#N/A"] and val.startswith("http"):
            return val
    fname = str(row.get("File Name","")).strip()
    if fname.startswith("http"): return fname
    return ""

@st.cache_data(ttl=600, show_spinner=False)
def fetch_file_and_extract(url, filename):
    if not url or str(url).strip() in ["","nan","None"]:
        return {"price_num":0.0,"price":"","status":"⚪ No link"}
    url = str(url).strip()
    if "drive.google.com" in url:
        match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        if match: url = "https://drive.google.com/uc?export=download&id={}".format(match.group(1))
    if "1drv.ms" in url or "onedrive" in url:
        if not url.endswith("download"): url = url.rstrip("/") + "?download=1"
    try:
        headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        resp = requests.get(url, timeout=30, headers=headers, allow_redirects=True)
        if resp.status_code != 200:
            return {"price_num":0.0,"price":"","status":"❌ HTTP {}".format(resp.status_code)}
        ext = filename.rsplit(".",1)[-1].lower() if "." in filename else url.split("?")[0].rsplit(".",1)[-1].lower()
        if ext not in ("xlsx","xls","pdf","docx","csv"): ext = "xlsx"
        result = extract_price_from_bytes(resp.content, ext)
        result["status"] = "✅ Extracted" if result["price_num"] > 0 else "⚠️ No price found in file"
        return result
    except requests.exceptions.Timeout:
        return {"price_num":0.0,"price":"","status":"❌ Timeout"}
    except Exception as e:
        return {"price_num":0.0,"price":"","status":"❌ {}".format(str(e)[:60])}

def analyze_real_catalog(df_catalog):
    results = []; total = len(df_catalog)
    prog_bar = st.progress(0); status_txt = st.empty()
    for i, (_, row) in enumerate(df_catalog.iterrows()):
        fname    = str(row.get("File Name","")).strip()
        vendor   = str(row.get("Vendor","")).strip()
        cat      = str(row.get("Category","")).strip()
        comments = str(row.get("Comments","")).strip()
        qp       = _parse_num(str(row.get("Quoted Price","")).strip())
        url      = get_file_link_from_row(row)
        status_txt.markdown("<div style='font-size:0.82em;color:#555'>Processing {}/{}: <b>{}</b></div>".format(i+1,total,fname), unsafe_allow_html=True)
        extracted = fetch_file_and_extract(url, fname) if url else {"price_num":0.0,"price":"","status":"⚪ No link provided"}
        ex_p   = extracted.get("price_num", 0.0)
        status = extracted.get("status", "—")
        best_p = ex_p if ex_p > 0 else qp
        source = "extracted" if ex_p > 0 else "catalog" if qp > 0 else "none"
        svcs   = parse_services_from_cell(comments)
        results.append({"Vendor":vendor,"Category":cat,"File Name":fname,"File Link":url,
                        "Quoted Price":qp,"Extracted Price":ex_p,"Best Price":best_p,
                        "Source":source,"Status":status,"Services":svcs})
        prog_bar.progress((i+1)/total)
    prog_bar.empty(); status_txt.empty()
    return pd.DataFrame(results)

# ════════════════════════════════════════════════════════════
# GITHUB LOADER
# ════════════════════════════════════════════════════════════
GITHUB_RAW = "https://raw.githubusercontent.com/avijeet528/vendor-draft-2/main/demo_quotes/{}"

@st.cache_data(show_spinner=False)
def load_price_from_github(filename):
    url = GITHUB_RAW.format(filename)
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code != 200:
            return {"price":"","price_num":0.0,"status":"Not found ({})".format(resp.status_code)}
        ext    = filename.rsplit(".",1)[-1].lower()
        result = extract_price_from_bytes(resp.content, ext)
        result["status"] = "✅ Extracted" if result["price_num"] > 0 else "⚠️ No price found"
        return result
    except Exception as e:
        return {"price":"","price_num":0.0,"status":"❌ Error: {}".format(str(e))}

@st.cache_data(show_spinner=False)
def load_all_prices_from_github(filenames_tuple):
    results = {}
    for fname in filenames_tuple:
        results[fname] = load_price_from_github(fname)
    return results

def get_price_for_row(row, gh_prices):
    fname = str(row.get("File Name","")).strip()
    qp    = _parse_num(str(row.get("Quoted Price","")).strip())
    gh    = gh_prices.get(fname,{})
    gh_p  = gh.get("price_num",0.0)
    if gh_p > 0: return gh_p, "extracted"
    if qp  > 0:  return qp,   "catalog"
    return 0.0, "none"

# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state and st.session_state["uploaded_catalog_df"] is not None):
    df_master   = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "file"

NO_DATA = (df_master is None or df_exploded is None or df_master.empty)

vendor_color_map = ({v: get_color(i) for i,v in enumerate(sorted(df_master["Vendor"].unique()))}
                    if not NO_DATA else {})

# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════
def sb_label(txt):
    st.markdown("<p style='color:#F0F0F0;font-weight:700;font-size:0.85em;margin:12px 0 4px;"
                "letter-spacing:0.5px;text-transform:uppercase'>{}</p>".format(txt), unsafe_allow_html=True)

def section_title(txt, caption=""):
    st.markdown("<div class='sec-heading'>{}</div>".format(txt), unsafe_allow_html=True)
    if caption: st.caption(caption)

def vendor_pill(v, color):
    return "<span class='vendor-badge' style='background:{}'>{}</span>".format(color, v)

def ai_box(content):
    st.markdown("<div class='ai-box'><div style='font-size:0.68em;font-weight:700;letter-spacing:1px;"
                "text-transform:uppercase;color:#6E2585;margin-bottom:6px'>AI Insight</div>"
                "{}</div>".format(content), unsafe_allow_html=True)

def resolve_url(row):
    fname = str(row.get("File Name","")).strip()
    if fname:
        local = os.path.join(DEMO_DIR, fname)
        if os.path.exists(local): return local
    for col in ["Hyperlink","File Link"]:
        val = str(row.get(col,"")).strip()
        if val and val not in ["","nan"]:
            if val.startswith("http"): return val
            local2 = os.path.join(DEMO_DIR, os.path.basename(val.replace("\\","/")))
            if os.path.exists(local2): return local2
    return ""

def kpi(col, val, lbl, bg):
    col.markdown("<div class='kpi-box' style='background:{}'>"
                 "<div class='kpi-value'>{}</div>"
                 "<div class='kpi-label'>{}</div></div>".format(bg,val,lbl), unsafe_allow_html=True)

def mini_kpi(col, val, lbl, bg, icon=""):
    col.markdown("<div class='kpi-box' style='background:{};min-height:80px'>"
                 "<div style='font-size:1.5em;margin-bottom:2px'>{}</div>"
                 "<div class='kpi-value' style='font-size:1.4em'>{}</div>"
                 "<div class='kpi-label'>{}</div></div>".format(bg,icon,val,lbl), unsafe_allow_html=True)

def pwc_header(title, subtitle="", eyebrow="IT Procurement Analytics"):
    st.markdown("<div style='background:#2D2D2D;color:white;padding:20px 28px;border-radius:4px;"
                "border-left:6px solid #D04A02;margin-bottom:22px'>"
                "<div style='font-size:0.72em;font-weight:700;letter-spacing:2px;"
                "text-transform:uppercase;color:#D04A02;margin-bottom:5px'>{}</div>"
                "<h1 style='margin:0;font-size:1.4em;font-weight:700;color:white'>{}</h1>"
                "{}</div>".format(eyebrow, title,
                "<p style='margin:6px 0 0;opacity:0.6;font-size:0.85em'>{}</p>".format(subtitle) if subtitle else ""),
                unsafe_allow_html=True)

def static_section_header(title, border_color=PwC_ORANGE):
    st.markdown("<div style='border:1px solid #e0e0e0;border-radius:4px;padding:10px 16px;"
                "margin-bottom:4px;background:white;border-left:4px solid {}'>"
                "<span style='font-weight:700;font-size:0.94em;color:#2D2D2D;"
                "display:block;line-height:1.4'>{}</span></div>".format(border_color, title),
                unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
selected_svcs = []; selected_cat = "All"; selected_vendor = "All"; d_filt = pd.DataFrame()

if not NO_DATA:
    with st.sidebar:
        st.markdown("<div style='text-align:center;padding:20px 0 14px'>"
                    "<div style='font-size:2em'>📋</div>"
                    "<div style='font-size:1.05em;font-weight:700;color:white;margin:5px 0 2px'>IT Procurement</div>"
                    "<div style='font-size:0.70em;color:#aaa;letter-spacing:1px;text-transform:uppercase'>"
                    "Intelligence Dashboard</div></div>"
                    "<hr style='border-color:#D04A02;border-width:2px;margin:0 0 14px'>", unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown("<div style='background:#D04A02;color:white;padding:6px 10px;border-radius:2px;"
                        "font-size:0.75em;font-weight:700;text-align:center;margin-bottom:10px'>"
                        "📤 USING UPLOADED CATALOG</div>", unsafe_allow_html=True)

        sb_label("📂 Category")
        all_cats = ["All"] + sorted([c for c in df_master["Category"].unique() if str(c).strip() not in ["","nan"]])
        selected_cat = st.selectbox("Category", all_cats, label_visibility="collapsed")

        sb_label("🏢 Vendor")
        vpool = df_master if selected_cat == "All" else df_master[df_master["Category"]==selected_cat]
        all_vendors = ["All"] + sorted([v for v in vpool["Vendor"].unique() if str(v).strip() not in ["","nan"]])
        selected_vendor = st.selectbox("Vendor", all_vendors, label_visibility="collapsed")

        st.markdown("<hr style='border-color:#555;margin:12px 0'>", unsafe_allow_html=True)

        d_filt = df_exploded.copy()
        if selected_cat != "All":    d_filt = d_filt[d_filt["Category"]==selected_cat]
        if selected_vendor != "All": d_filt = d_filt[d_filt["Vendor"]==selected_vendor]

        sb_label("🔍 Search Services")
        svc_search = st.text_input("Search", placeholder="e.g. Cisco, Azure…", label_visibility="collapsed")
        avail = sorted([s for s in d_filt["Service"].unique() if str(s).strip() not in ["","nan"]])
        if svc_search: avail = [s for s in avail if svc_search.lower() in s.lower()]

        sb_label("🛠 Select Services ({})".format(len(avail)))
        selected_svcs = st.multiselect("Services", options=avail, default=[], label_visibility="collapsed",
                                        help="Select services to compare vendors")

        st.markdown("<hr style='border-color:#555;margin:12px 0'>", unsafe_allow_html=True)
        st.markdown("<p style='color:#888;font-size:0.78em;margin:2px 0'>"
                    "📄 {} quotes | 🛠 {} services | 🏢 {} vendors</p>".format(
                        len(df_master), df_exploded["Service"].nunique(), df_master["Vendor"].nunique()),
                    unsafe_allow_html=True)

        if selected_svcs:
            st.markdown("<p style='color:#D04A02;font-size:0.80em;font-weight:700;margin:4px 0'>"
                        "✅ {} service(s) selected</p>".format(len(selected_svcs)), unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown("<hr style='border-color:#555;margin:12px 0'>", unsafe_allow_html=True)
            if st.button("🔄 Reset to Default Catalog", use_container_width=True):
                st.session_state["uploaded_catalog_df"]  = None
                st.session_state["uploaded_catalog_exp"] = None
                st.rerun()

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
pwc_header("Procurement Intelligence Dashboard",
           "Browse quotations · Compare prices · Upload & score · AI insights · Verdict system")

if os.path.exists(DEMO_DIR):
    st.markdown("<div class='demo-banner'>🎯 <b>DEMO MODE</b> — Sample quote files detected in demo_quotes/.</div>",
                unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    use_filt = d_filt if not d_filt.empty else df_exploded
    k1,k2,k3,k4 = st.columns(4)
    kpi(k1, use_filt["File Name"].nunique(), "Total Quotes",    PwC_ORANGE)
    kpi(k2, use_filt["Service"].nunique(),   "Unique Services", PwC_BLUE)
    kpi(k3, use_filt["Vendor"].nunique(),    "Vendors",         PwC_TEAL)
    kpi(k4, use_filt["Category"].nunique(),  "Categories",      PwC_DARK)
    st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
tab0,tab1,tab2,tab3,tab4,tab5,tab6,tab7 = st.tabs([
    "🗂️ Catalog Overview",
    "📊 Analytics",
    "📋 Browse & Verdict",
    "📤 Upload & Score",
    "📄 Data Table",
    "🗂 Upload Catalog",
    "🔍 Vendor Analysis",
    "📂 Real Analysis",
])

# ════════════════════════════════════════════════════════════
# TAB 0 — CATALOG OVERVIEW  (reads real Master Catalog.xlsx)
# ════════════════════════════════════════════════════════════
with tab0:
    if NO_DATA:
        st.info("No catalog loaded. Go to 🗂 Upload Catalog tab.")
    else:
        df_ov = df_master.copy()
        df_ov["Subcategory"] = df_ov.apply(
            lambda r: infer_subcategory(r.get("Category",""), r.get("Comments",""), r.get("File Name","")), axis=1)

        all_cats_ov = sorted([c for c in df_ov["Category"].unique()
                               if str(c).strip() not in ["","nan"]])

        pwc_header("Catalog Overview",
                   "Categories · Subcategories · Vendors · Services · Quotations — from Master Catalog.xlsx")

        # ── Top KPIs ──
        k0a,k0b,k0c,k0d,k0e = st.columns(5)
        mini_kpi(k0a, len(df_ov),                        "Total Quotations", PwC_ORANGE, "📄")
        mini_kpi(k0b, df_ov["Vendor"].nunique(),          "Unique Vendors",   PwC_BLUE,   "🏢")
        mini_kpi(k0c, df_ov["Category"].nunique(),        "Categories",       PwC_TEAL,   "📂")
        mini_kpi(k0d, df_ov["Subcategory"].nunique(),     "Subcategories",    PwC_AMBER,  "🏷️")
        mini_kpi(k0e, df_exploded["Service"].nunique() if df_exploded is not None else "—",
                 "Unique Services", PwC_DARK, "🛠")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Category stats ──
        section_title("CATEGORIES AT A GLANCE")
        cat_stats = []
        for cat in all_cats_ov:
            d_cat = df_ov[df_ov["Category"]==cat]
            n_svcs = df_exploded[df_exploded["Category"]==cat]["Service"].nunique() if df_exploded is not None else 0
            cat_stats.append({"Category":cat,"Quotations":len(d_cat),
                               "Vendors":d_cat["Vendor"].nunique(),
                               "Subcategories":d_cat["Subcategory"].nunique(),"Services":n_svcs})
        cat_stats_df = pd.DataFrame(cat_stats).sort_values("Quotations", ascending=False)

        CAT_ICONS = {
            "Cybersecurity":"🛡️","Network & Telecom":"🌐","Hosting":"🖥️",
            "M365 & Power Platform":"☁️","IdAM":"🔑","Service Management (SNow)":"⚙️",
            "Summary & Reporting":"📊","Fabric":"🔌","Internet access":"🌍","Power":"⚡","PDU Hosting":"🔋",
        }

        rows_of_3 = [cat_stats_df.iloc[i:i+3] for i in range(0, len(cat_stats_df), 3)]
        for row_chunk in rows_of_3:
            cols = st.columns(len(row_chunk), gap="medium")
            for col_idx,(_, row_s) in enumerate(row_chunk.iterrows()):
                cat_name = row_s["Category"]
                icon  = CAT_ICONS.get(cat_name,"📁")
                cidx  = all_cats_ov.index(cat_name) if cat_name in all_cats_ov else 0
                color = CAT_COLORS[cidx % len(CAT_COLORS)]
                cols[col_idx].markdown(
                    "<div style='background:white;border:1px solid #e0e0e0;border-radius:6px;"
                    "padding:16px 18px;border-top:4px solid {};height:100%'>"
                    "<div style='font-size:1.4em;margin-bottom:6px'>{}</div>"
                    "<div style='font-size:0.92em;font-weight:700;color:#2D2D2D;margin-bottom:10px;line-height:1.3'>{}</div>"
                    "<div style='display:flex;gap:8px;flex-wrap:wrap'>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;font-size:0.76em;font-weight:700;color:#2D2D2D'>📄 {} quotes</span>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;font-size:0.76em;font-weight:700;color:#295477'>🏢 {} vendors</span>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;font-size:0.76em;font-weight:700;color:#299D8F'>🛠 {} services</span>"
                    "</div></div>".format(color,icon,cat_name,
                                          row_s["Quotations"],row_s["Vendors"],row_s["Services"]),
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Charts ──
        section_title("DISTRIBUTION — QUOTATIONS & VENDORS BY CATEGORY")
        ch_a, ch_b = st.columns(2, gap="large")
        with ch_a:
            fig_cq = go.Figure(go.Bar(
                x=cat_stats_df["Category"], y=cat_stats_df["Quotations"],
                marker_color=[CAT_COLORS[all_cats_ov.index(c)%len(CAT_COLORS)] if c in all_cats_ov else PwC_GREY
                              for c in cat_stats_df["Category"]],
                marker_line_width=0, text=cat_stats_df["Quotations"], textposition="outside"))
            fig_cq.update_layout(height=340, plot_bgcolor=CBG, paper_bgcolor=CBG,
                                  margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                                  title=dict(text="Quotations per Category",font=dict(size=12,color=PwC_DARK),x=0),
                                  yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                                  xaxis=dict(tickangle=-30,tickfont=dict(size=9.5)),bargap=0.35,showlegend=False)
            st.plotly_chart(fig_cq, use_container_width=True)
        with ch_b:
            fig_cv = go.Figure(go.Bar(
                x=cat_stats_df["Category"], y=cat_stats_df["Vendors"],
                marker_color=[CAT_COLORS[all_cats_ov.index(c)%len(CAT_COLORS)] if c in all_cats_ov else PwC_GREY
                              for c in cat_stats_df["Category"]],
                marker_line_width=0, text=cat_stats_df["Vendors"], textposition="outside"))
            fig_cv.update_layout(height=340, plot_bgcolor=CBG, paper_bgcolor=CBG,
                                  margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                                  title=dict(text="Vendors per Category",font=dict(size=12,color=PwC_DARK),x=0),
                                  yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                                  xaxis=dict(tickangle=-30,tickfont=dict(size=9.5)),bargap=0.35,showlegend=False)
            st.plotly_chart(fig_cv, use_container_width=True)

        # ── Donut + summary table ──
        st.markdown("<br>", unsafe_allow_html=True)
        section_title("CATALOG COMPOSITION — QUOTE SHARE BY CATEGORY")
        d_left, d_right = st.columns([1,2], gap="large")
        with d_left:
            fig_donut = px.pie(cat_stats_df, values="Quotations", names="Category",
                               hole=0.55, color_discrete_sequence=CAT_COLORS)
            fig_donut.update_traces(textposition="outside", textinfo="percent+label",
                                     textfont_size=10, pull=[0.03]*len(cat_stats_df))
            fig_donut.update_layout(height=360, margin=dict(l=10,r=10,t=10,b=10),
                                     paper_bgcolor=CBG, font=CFONT, showlegend=False)
            st.plotly_chart(fig_donut, use_container_width=True)
        with d_right:
            tbl_rows = ["<table class='comp-table'><thead><tr>"
                        "<th>Category</th><th style='text-align:center'>Quotes</th>"
                        "<th style='text-align:center'>Vendors</th><th style='text-align:center'>Services</th>"
                        "<th style='text-align:center'>Subcategories</th></tr></thead><tbody>"]
            for _, row_s in cat_stats_df.iterrows():
                cat_n = row_s["Category"]
                cidx  = all_cats_ov.index(cat_n) if cat_n in all_cats_ov else 0
                color = CAT_COLORS[cidx % len(CAT_COLORS)]
                icon  = CAT_ICONS.get(cat_n,"📁")
                tbl_rows.append(
                    "<tr><td><span style='border-left:4px solid {};padding-left:8px;font-weight:600'>"
                    "{} {}</span></td>"
                    "<td style='text-align:center;font-weight:700;color:#D04A02'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#295477'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#299D8F'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#EB8C00'>{}</td>"
                    "</tr>".format(color,icon,cat_n,row_s["Quotations"],
                                   row_s["Vendors"],row_s["Services"],row_s["Subcategories"]))
            tbl_rows.append("</tbody></table>")
            st.markdown("".join(tbl_rows), unsafe_allow_html=True)

        # ── Per-category drill-down ──
        st.markdown("<br>", unsafe_allow_html=True)
        section_title("DRILL-DOWN BY CATEGORY — SUBCATEGORIES & VENDORS")
        ordered_cats = (["Cybersecurity"] + [c for c in all_cats_ov if c != "Cybersecurity"])
        cat_subtabs  = st.tabs(["🛡️ {}".format(c) if c=="Cybersecurity" else c for c in ordered_cats])

        for tab_idx, cat_name in enumerate(ordered_cats):
            with cat_subtabs[tab_idx]:
                d_cat = df_ov[df_ov["Category"]==cat_name].copy()
                if d_cat.empty: st.info("No data for {}.".format(cat_name)); continue

                ck1,ck2,ck3,ck4 = st.columns(4)
                n_q   = len(d_cat); n_v = d_cat["Vendor"].nunique()
                n_sub = d_cat["Subcategory"].nunique()
                n_svc = df_exploded[df_exploded["Category"]==cat_name]["Service"].nunique() if df_exploded is not None else 0
                mini_kpi(ck1,n_q,  "Quotations",   PwC_ORANGE,"📄")
                mini_kpi(ck2,n_v,  "Vendors",       PwC_BLUE,  "🏢")
                mini_kpi(ck3,n_svc,"Services",       PwC_TEAL,  "🛠")
                mini_kpi(ck4,n_sub,"Subcategories",  PwC_AMBER, "🏷️")
                st.markdown("<br>", unsafe_allow_html=True)

                sub_stats = (d_cat.groupby("Subcategory")
                             .agg(Quotations=("File Name","count"), Vendors=("Vendor","nunique"))
                             .reset_index().sort_values("Quotations", ascending=False))

                sc_col1, sc_col2 = st.columns([1,1], gap="medium")
                with sc_col1:
                    st.markdown("<div style='font-size:0.78em;font-weight:700;letter-spacing:0.8px;"
                                "text-transform:uppercase;color:#2D2D2D;margin-bottom:8px'>Subcategories</div>",
                                unsafe_allow_html=True)
                    fig_sub = go.Figure(go.Bar(
                        x=sub_stats["Quotations"], y=sub_stats["Subcategory"], orientation="h",
                        marker_color=PwC_ORANGE, marker_line_width=0,
                        text=sub_stats["Quotations"], textposition="outside", textfont=dict(size=10)))
                    fig_sub.update_layout(
                        height=max(220, len(sub_stats)*38), plot_bgcolor=CBG, paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=10,b=10), font=CFONT,
                        xaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False,title="Quotations"),
                        yaxis=dict(autorange="reversed",tickfont=dict(size=10)), bargap=0.3)
                    st.plotly_chart(fig_sub, use_container_width=True)

                with sc_col2:
                    st.markdown("<div style='font-size:0.78em;font-weight:700;letter-spacing:0.8px;"
                                "text-transform:uppercase;color:#2D2D2D;margin-bottom:8px'>Vendors</div>",
                                unsafe_allow_html=True)
                    vnd_stats = (d_cat.groupby("Vendor").agg(Quotations=("File Name","count"))
                                 .reset_index().sort_values("Quotations", ascending=False))
                    fig_vnd = go.Figure(go.Bar(
                        x=vnd_stats["Quotations"], y=vnd_stats["Vendor"], orientation="h",
                        marker_color=[vendor_color_map.get(v,PwC_GREY) for v in vnd_stats["Vendor"]],
                        marker_line_width=0, text=vnd_stats["Quotations"], textposition="outside", textfont=dict(size=10)))
                    fig_vnd.update_layout(
                        height=max(220, len(vnd_stats)*38), plot_bgcolor=CBG, paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=10,b=10), font=CFONT,
                        xaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False,title="Quotations"),
                        yaxis=dict(autorange="reversed",tickfont=dict(size=10)), bargap=0.3)
                    st.plotly_chart(fig_vnd, use_container_width=True)

                # Subcategory detail table
                st.markdown("<div style='font-size:0.78em;font-weight:700;letter-spacing:0.8px;"
                            "text-transform:uppercase;color:#2D2D2D;margin:10px 0 8px'>Subcategory Detail</div>",
                            unsafe_allow_html=True)
                sub_tbl = ["<table class='comp-table'><thead><tr><th>Subcategory</th><th>Vendors</th>"
                           "<th style='text-align:center'>Quotations</th><th>Files</th></tr></thead><tbody>"]
                for si,(_, sr) in enumerate(sub_stats.iterrows()):
                    bg       = "white" if si%2==0 else "#F3F3F3"
                    sub_name = sr["Subcategory"]
                    d_sub    = d_cat[d_cat["Subcategory"]==sub_name]
                    pills    = " ".join(["<span class='vendor-badge' style='background:{}'>{}</span>".format(
                                        vendor_color_map.get(v,PwC_GREY),v)
                                        for v in sorted(d_sub["Vendor"].unique())])
                    files_list = ", ".join(d_sub["File Name"].apply(lambda x: x[:30]).tolist()[:3])
                    if len(d_sub) > 3: files_list += " (+{} more)".format(len(d_sub)-3)
                    sub_tbl.append(
                        "<tr style='background:{}'><td style='font-weight:600'>{}</td>"
                        "<td>{}</td><td style='text-align:center;font-weight:700;color:#D04A02'>{}</td>"
                        "<td style='font-size:0.78em;color:#555;font-family:monospace'>{}</td></tr>".format(
                            bg,sub_name,pills,sr["Quotations"],files_list))
                sub_tbl.append("</tbody></table>")
                st.markdown("".join(sub_tbl), unsafe_allow_html=True)

                # Cybersecurity callout
                if cat_name == "Cybersecurity":
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown(
                        "<div style='background:#F8F0FF;border-left:5px solid #6E2585;border-radius:4px;"
                        "padding:14px 18px;margin:6px 0'>"
                        "<div style='font-size:0.70em;font-weight:700;letter-spacing:1px;"
                        "text-transform:uppercase;color:#6E2585;margin-bottom:6px'>Cybersecurity Coverage</div>"
                        "<div style='font-size:0.87em;color:#2D2D2D'>"
                        "This category covers <b>{}</b> security subcategories across <b>{}</b> vendors "
                        "with <b>{}</b> total quotation files. Vendors include: <b>{}</b>."
                        "</div></div>".format(n_sub,n_v,n_q,", ".join(sorted(d_cat["Vendor"].unique()))),
                        unsafe_allow_html=True)

                # File list
                static_section_header("📄 All {} quotation files in {}".format(len(d_cat), cat_name), P

