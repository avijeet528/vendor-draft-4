# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
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
# PwC COLOURS — strictly standardised
# ════════════════════════════════════════════════════════════
PwC_ORANGE  = "#D04A02"
PwC_RED     = "#E0301E"
PwC_YELLOW  = "#FFB600"
PwC_GREEN   = "#22992E"
PwC_NAVY    = "#003078"
PwC_BLUE    = "#295477"
PwC_TEAL    = "#299D8F"
PwC_PURPLE  = "#6E2585"
PwC_DARK    = "#2D2D2D"
PwC_GREY    = "#7D7D7D"
PwC_LTGREY  = "#F3F3F3"
PwC_AMBER   = "#EB8C00"
PwC_WHITE   = "#FFFFFF"

# Standardised chart palette — PwC approved sequence
CHART_SEQ = [PwC_ORANGE, PwC_NAVY, PwC_TEAL, PwC_AMBER,
             PwC_PURPLE, PwC_GREEN, PwC_RED, PwC_BLUE]

def chart_color(i): return CHART_SEQ[i % len(CHART_SEQ)]

CFONT = dict(family="ITC Charter, Georgia, 'Source Sans Pro', Arial", size=11, color=PwC_DARK)
CBG   = PwC_LTGREY
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# CSS — PwC brand + ITC Charter / Source Sans Pro fonts
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

/* ── Global font — PwC uses ITC Charter for headings, Source Sans Pro for body ── */
html,body,[class*="css"],div,p,span,td,th,label,button,.stMarkdown{
    font-family:'Source Sans Pro','Helvetica Neue',Arial,sans-serif !important;}
h1,h2,h3,.pwc-heading{
    font-family:Georgia,'ITC Charter','Source Sans Pro',serif !important;
    font-weight:700 !important;}

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
.kpi-value{font-size:2.1em;font-weight:700;margin:0;line-height:1.1;
    font-family:Georgia,serif !important;}
.kpi-label{font-size:0.78em;font-weight:700;opacity:0.9;margin-top:5px;
    letter-spacing:0.8px;text-transform:uppercase;}

/* ── Tabs ── */
button[data-baseweb="tab"]{font-weight:600 !important;font-size:0.92em !important;color:#7D7D7D !important;}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;border-bottom:3px solid #D04A02 !important;}

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

/* ── Vendor badge ── */
.vendor-badge{display:inline-block;padding:3px 8px;border-radius:2px;
    color:white;font-size:0.78em;font-weight:700;white-space:nowrap;}

/* ── Score cards ── */
.score-card{border-radius:4px;padding:14px 16px;margin-bottom:10px;border-left:5px solid #D04A02;}
.score-card.green{background:#F0FFF4;border-color:#22992E;}
.score-card.yellow{background:#FFF8E1;border-color:#FFB600;}
.score-card.red{background:#FFF3F0;border-color:#E0301E;}
.score-card.navy{background:#EEF2FF;border-color:#003078;}

/* ── Verdict boxes ── */
.verdict-green{background:#F0FFF4;border:2px solid #22992E;border-radius:4px;
    padding:12px 16px;color:#22992E;font-weight:700;}
.verdict-yellow{background:#FFF8E1;border:2px solid #FFB600;border-radius:4px;
    padding:12px 16px;color:#856404;font-weight:700;}
.verdict-red{background:#FFF3F0;border:2px solid #E0301E;border-radius:4px;
    padding:12px 16px;color:#E0301E;font-weight:700;}

/* ── AI / Chat boxes ── */
.ai-box{background:#EEF2FF;border-left:5px solid #003078;
    border-radius:4px;padding:14px 18px;margin:10px 0;}
.chat-user{background:#D04A02;color:white;border-radius:12px 12px 2px 12px;
    padding:10px 16px;margin:8px 0 8px auto;max-width:75%;
    font-size:0.88em;display:inline-block;float:right;clear:both;}
.chat-bot{background:white;color:#2D2D2D;border-radius:12px 12px 12px 2px;
    border:1px solid #e0e0e0;border-left:4px solid #003078;
    padding:10px 16px;margin:8px 0;max-width:85%;
    font-size:0.88em;display:inline-block;float:left;clear:both;}
.chat-wrap{overflow:hidden;margin-bottom:8px;}
.chat-container{background:#F9F9F9;border:1px solid #e0e0e0;
    border-radius:6px;padding:16px;max-height:420px;
    overflow-y:auto;margin-bottom:12px;}

/* ── Section heading ── */
.sec-heading{display:block;font-size:0.78em;font-weight:700;
    letter-spacing:1px;text-transform:uppercase;
    color:#D04A02;margin:18px 0 8px;line-height:1.4;}

/* ── Price tag ── */
.price-tag-green{background:#22992E;color:white;padding:4px 10px;
    border-radius:3px;font-weight:700;font-size:0.85em;}
.price-tag-yellow{background:#FFB600;color:white;padding:4px 10px;
    border-radius:3px;font-weight:700;font-size:0.85em;}
.price-tag-red{background:#E0301E;color:white;padding:4px 10px;
    border-radius:3px;font-weight:700;font-size:0.85em;}
.price-tag-grey{background:#7D7D7D;color:white;padding:4px 10px;
    border-radius:3px;font-weight:700;font-size:0.85em;}
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
    if s >= 40: return PwC_AMBER
    return PwC_RED

def score_css(s):
    if s is None: return "yellow"
    if s >= 70: return "green"
    if s >= 40: return "yellow"
    return "red"

def get_verdict(ps):
    if ps is None: return "⚪ No Data","No comparison data available.",PwC_GREY
    if ps >= 70:   return "✅ COMPETITIVE","This quote is priced competitively.",PwC_GREEN
    if ps >= 40:   return "🟡 AVERAGE","Within average range. Negotiate for discount.","#856404"
    return "🔴 HIGH — NEGOTIATE","Above historical average. Strongly recommend negotiating.",PwC_RED

# ════════════════════════════════════════════════════════════
# DATA HELPERS — safe clean + explode
# ════════════════════════════════════════════════════════════
def _clean_df(df):
    safe_cols = [c for c in df.columns if c != "Services List"]
    df = df[safe_cols].copy()
    for col in df.columns:
        try: df[col] = df[col].fillna("").apply(lambda x: str(x).strip())
        except: df[col] = ""
    mask = (df["Category"].apply(lambda x: x in ["","nan"]) &
            df["Vendor"].apply(lambda x: x in ["","nan"]))
    df = df[~mask].copy()
    df.reset_index(drop=True, inplace=True)
    return df

def _parse_services(v):
    if not v or str(v).strip() in ["","nan","None"]: return ["(unspecified)"]
    s = str(v).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts = [p.strip() for p in s.split("\n") if p.strip() and p.strip() != "nan"]
    if not parts: parts = [p.strip() for p in s.split(";") if p.strip()]
    if not parts and len(s) < 300: parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts if parts else ["(unspecified)"]

def _explode_services(df):
    df2 = df.copy()
    df2["Services List"] = df2["Comments"].apply(_parse_services)
    df_exp = df2.explode("Services List").copy()
    df_exp.rename(columns={"Services List":"Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].apply(lambda x: str(x).strip())
    df_exp = df_exp[~df_exp["Service"].isin(["","(unspecified)","nan","None"])].reset_index(drop=True)
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
    CSV_PATH = "master_catalog.csv"
    XLS_PATH = "Master Catalog.xlsx"
    df = None
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
    if df is None and os.path.exists(CSV_PATH):
        try:
            df = pd.read_csv(CSV_PATH)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("CSV load error: {}".format(e)); df = None
    if df is None: return None, None
    df = _normalise_columns(df)
    df = _clean_df(df)
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
    for col in ["Hyperlink","File Link","File URL","URL"]:
        val = str(row.get(col,"")).strip()
        if val and val not in ["","nan","None","#N/A"] and val.startswith("http"):
            return val
    return ""

@st.cache_data(ttl=600, show_spinner=False)
def fetch_file_and_extract(url, filename):
    if not url or str(url).strip() in ["","nan","None"]:
        return {"price_num":0.0,"price":"","status":"⚪ No link"}
    url = str(url).strip()
    if "drive.google.com" in url:
        match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        if match: url = "https://drive.google.com/uc?export=download&id={}".format(match.group(1))
    try:
        headers = {"User-Agent":"Mozilla/5.0"}
        resp = requests.get(url, timeout=30, headers=headers, allow_redirects=True)
        if resp.status_code != 200:
            return {"price_num":0.0,"price":"","status":"❌ HTTP {}".format(resp.status_code)}
        ext = filename.rsplit(".",1)[-1].lower() if "." in filename else "xlsx"
        if ext not in ("xlsx","xls","pdf","docx","csv"): ext = "xlsx"
        result = extract_price_from_bytes(resp.content, ext)
        result["status"] = "✅ Extracted" if result["price_num"] > 0 else "⚠️ No price found"
        return result
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
        extracted = fetch_file_and_extract(url, fname) if url else {"price_num":0.0,"price":"","status":"⚪ No link"}
        ex_p   = extracted.get("price_num",0.0)
        best_p = ex_p if ex_p > 0 else qp
        source = "extracted" if ex_p > 0 else "catalog" if qp > 0 else "none"
        results.append({"Vendor":vendor,"Category":cat,"File Name":fname,"File Link":url,
                        "Quoted Price":qp,"Extracted Price":ex_p,"Best Price":best_p,
                        "Source":source,"Status":extracted.get("status","—"),
                        "Services":parse_services_from_cell(comments)})
        prog_bar.progress((i+1)/total)
    prog_bar.empty(); status_txt.empty()
    return pd.DataFrame(results)

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
        return "General Security"
    if "network" in cat or "telecom" in cat:
        if "meraki" in txt: return "Cisco Meraki"
        if "palo alto" in txt: return "Palo Alto NGFW"
        if "equinix" in txt: return "Equinix Interconnect"
        if "cisco" in txt: return "Cisco Networking"
        return "General Network"
    if "hosting" in cat:
        if any(k in txt for k in ["vmware","vcf"]): return "VMware"
        if "oracle" in txt: return "Oracle DB"
        if "netapp" in txt: return "NetApp Storage"
        if any(k in txt for k in ["colocation","colo"]): return "Colocation Build"
        return "General Hosting"
    if "m365" in cat or "power platform" in cat: return "M365 Licensing"
    if "idam" in cat or "iam" in cat: return "Identity Migration"
    if "snow" in cat or "servicenow" in cat: return "ServiceNow ITSM"
    if "summary" in cat or "reporting" in cat: return "Reporting & Tracking"
    return str(category).strip().title()

# ════════════════════════════════════════════════════════════
# CHATBOT ENGINE
# ════════════════════════════════════════════════════════════
def chatbot_response(user_msg, df_master, df_exploded):
    """
    Conversational procurement assistant.
    Understands natural language queries about services,
    vendors, prices and comparisons.
    """
    msg = user_msg.lower().strip()
    all_services = sorted(df_exploded["Service"].unique().tolist()) if df_exploded is not None else []
    all_vendors  = sorted(df_master["Vendor"].unique().tolist())    if df_master  is not None else []
    all_cats     = sorted(df_master["Category"].unique().tolist())  if df_master  is not None else []

    # ── Greeting ──
    if any(w in msg for w in ["hello","hi","hey","good morning","good afternoon"]):
        return {
            "type": "text",
            "text": ("👋 Hello! I'm your **PwC Procurement Assistant**.\n\n"
                     "I can help you:\n"
                     "- 🔍 Find which vendors quoted for a specific service\n"
                     "- 💰 Compare prices across vendors\n"
                     "- 📋 Check if a service combination has been quoted before\n"
                     "- 📊 Tell you if a price is competitive\n\n"
                     "Try asking: *'Who quoted for Cisco Catalyst?'* or "
                     "*'Compare prices for Palo Alto'*")
        }

    # ── Help ──
    if any(w in msg for w in ["help","what can you","capabilities","options"]):
        return {
            "type": "text",
            "text": ("**I can answer questions like:**\n\n"
                     "🔍 *'Who has quoted for [service]?'*\n"
                     "💰 *'What is the price for [service]?'*\n"
                     "📊 *'Compare [service] across vendors'*\n"
                     "🏢 *'What services does [vendor] offer?'*\n"
                     "📋 *'Show me all Network & Telecom quotes'*\n"
                     "⚖️ *'Is [vendor] competitive for [service]?'*\n"
                     "🔄 *'Have we quoted [service A] and [service B] together?'*")
        }

    # ── List all services ──
    if any(p in msg for p in ["list all services","all services","show services","what services"]):
        by_cat = defaultdict(list)
        for _, r in df_exploded.iterrows():
            by_cat[r["Category"]].append(r["Service"])
        lines = []
        for cat, svcs in sorted(by_cat.items()):
            unique_svcs = sorted(set(svcs))[:8]
            lines.append("**{}** ({}): {}".format(cat, len(set(svcs)), " · ".join(unique_svcs)))
        return {"type":"text","text":"**All services by category:**\n\n" + "\n\n".join(lines)}

    # ── List all vendors ──
    if any(p in msg for p in ["list vendors","all vendors","show vendors","who are the vendors"]):
        vendor_info = []
        for v in all_vendors:
            n_quotes = len(df_master[df_master["Vendor"]==v])
            n_svcs   = df_exploded[df_exploded["Vendor"]==v]["Service"].nunique()
            vendor_info.append("**{}** — {} quotes · {} services".format(v, n_quotes, n_svcs))
        return {"type":"text","text":"**Vendors in catalog:**\n\n" + "\n".join(vendor_info)}

    # ── Category query ──
    matched_cat = None
    for cat in all_cats:
        if cat.lower() in msg or any(w in msg for w in cat.lower().split()):
            matched_cat = cat; break

    # ── Vendor match ──
    matched_vendor = None
    for v in all_vendors:
        if v.lower() in msg:
            matched_vendor = v; break

    # ── Service match (fuzzy keyword) ──
    matched_services = []
    for svc in all_services:
        svc_words = [w for w in svc.lower().split() if len(w) > 3]
        if any(w in msg for w in svc_words) or svc.lower() in msg:
            matched_services.append(svc)
    matched_services = matched_services[:10]

    # ── Vendor + service query ──
    if matched_vendor and matched_services:
        svc = matched_services[0]
        d   = df_exploded[(df_exploded["Vendor"]==matched_vendor) & (df_exploded["Service"]==svc)]
        if not d.empty:
            files = d["File Name"].unique().tolist()
            prices = []
            for f in files:
                qp = _parse_num(str(df_master[df_master["File Name"]==f]["Quoted Price"].values[0])
                                if len(df_master[df_master["File Name"]==f]) > 0 else "0")
                if qp > 0: prices.append(qp)
            price_txt = "Price: **{}**".format(_fmt(prices[0])) if prices else "No price data available."
            return {
                "type": "vendor_service",
                "text": ("✅ **{}** has quoted for **{}**.\n\n"
                         "{}\n\n📄 Files: {}".format(
                             matched_vendor, svc, price_txt,
                             ", ".join(files[:3]))),
                "vendor": matched_vendor,
                "service": svc,
                "has_quote": True,
                "prices": prices,
            }
        else:
            return {
                "type": "text",
                "text": ("❌ **{}** has **not** quoted for **{}** in our catalog.\n\n"
                         "💡 Try asking *'Who has quoted for {}?'* to see all vendors.".format(
                             matched_vendor, svc, svc))
            }

    # ── Compare service across vendors ──
    if matched_services and any(w in msg for w in ["compare","vs","versus","price","cost","expensive","cheap","competitive"]):
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc].drop_duplicates(subset=["Vendor","File Name"])
        if d.empty:
            return {"type":"text","text":"❌ No quotes found for **{}**.".format(svc)}

        vendor_prices = {}
        for _, r in d.iterrows():
            v  = r["Vendor"]
            qp = _parse_num(str(r.get("Quoted Price","")).strip())
            ck = "px_{}".format(str(r.get("File Name","")).strip())
            ca = st.session_state.get(ck)
            ep = ca["price_num"] if ca else 0.0
            p  = ep if ep > 0 else qp
            if p > 0: vendor_prices[v] = min(vendor_prices.get(v,p), p)

        if not vendor_prices:
            lines = ["**{}** — {} quotes (no price data)".format(
                v, len(d[d["Vendor"]==v])) for v in d["Vendor"].unique()]
            return {
                "type": "service_no_price",
                "text": ("📋 **{}** has been quoted by **{}** vendor(s):\n\n{}".format(
                    svc, len(d["Vendor"].unique()), "\n".join(lines))),
                "service": svc,
                "vendors": d["Vendor"].unique().tolist(),
            }

        avg_p   = sum(vendor_prices.values()) / len(vendor_prices)
        best_v  = min(vendor_prices, key=vendor_prices.get)
        worst_v = max(vendor_prices, key=vendor_prices.get)
        spread  = round((max(vendor_prices.values())-min(vendor_prices.values()))
                        /min(vendor_prices.values())*100, 1) if min(vendor_prices.values()) > 0 else 0

        lines = []
        for v, p in sorted(vendor_prices.items(), key=lambda x: x[1]):
            pct = round((p-avg_p)/avg_p*100,1) if avg_p > 0 else 0
            tag = "🟢 Cheapest" if v==best_v else "🔴 Most expensive" if v==worst_v else "🟡 Mid-range"
            lines.append("**{}**: {} ({})".format(v, _fmt(p), tag))

        return {
            "type": "comparison",
            "text": ("📊 **Price comparison for {}**:\n\n{}\n\n"
                     "💡 Price spread: **{}%** — "
                     "{} is cheapest at **{}**.".format(
                         svc, "\n".join(lines), spread, best_v, _fmt(min(vendor_prices.values())))),
            "service": svc,
            "vendor_prices": vendor_prices,
            "avg": avg_p,
            "spread": spread,
            "best_vendor": best_v,
        }

    # ── Who quoted for service ──
    if matched_services and any(w in msg for w in ["who","vendor","quoted","available","offered"]):
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc].drop_duplicates(subset=["Vendor"])
        if d.empty:
            return {"type":"text","text":"❌ No vendor has quoted for **{}** in our catalog.".format(svc)}
        vendors = d["Vendor"].unique().tolist()
        n_files = df_exploded[df_exploded["Service"]==svc]["File Name"].nunique()
        return {
            "type": "who_quoted",
            "text": ("✅ **{}** vendor(s) have quoted for **{}**:\n\n"
                     "{}\n\n📄 Total files: {}".format(
                         len(vendors), svc,
                         "\n".join(["• **{}**".format(v) for v in vendors]), n_files)),
            "service": svc,
            "vendors": vendors,
        }

    # ── What does vendor offer ──
    if matched_vendor and any(w in msg for w in ["offer","service","provide","quote","what does","capability"]):
        d    = df_exploded[df_exploded["Vendor"]==matched_vendor]
        svcs = sorted(d["Service"].unique().tolist())
        cats = sorted(d["Category"].unique().tolist())
        n_q  = len(df_master[df_master["Vendor"]==matched_vendor])
        return {
            "type": "vendor_profile",
            "text": ("🏢 **{}** — Vendor Profile\n\n"
                     "📂 **Categories:** {}\n"
                     "📄 **Total quotes:** {}\n"
                     "🛠 **Services ({}):**\n{}".format(
                         matched_vendor, ", ".join(cats), n_q,
                         len(svcs), "\n".join(["• {}".format(s) for s in svcs[:15]])
                         + ("\n...and {} more".format(len(svcs)-15) if len(svcs) > 15 else ""))),
            "vendor": matched_vendor,
            "services": svcs,
        }

    # ── Category overview ──
    if matched_cat:
        d_cat  = df_master[df_master["Category"]==matched_cat]
        n_v    = d_cat["Vendor"].nunique()
        n_q    = len(d_cat)
        n_svc  = df_exploded[df_exploded["Category"]==matched_cat]["Service"].nunique()
        vendors= sorted(d_cat["Vendor"].unique().tolist())
        return {
            "type": "category_overview",
            "text": ("📂 **{}** — Category Overview\n\n"
                     "📄 **{}** quotations · 🏢 **{}** vendors · 🛠 **{}** services\n\n"
                     "**Vendors:** {}".format(
                         matched_cat, n_q, n_v, n_svc, ", ".join(vendors))),
            "category": matched_cat,
        }

    # ── Combination check ──
    if len(matched_services) >= 2 and any(w in msg for w in ["together","combination","both","and","all of"]):
        svc_set = set(matched_services[:3])
        vsmap   = defaultdict(set)
        for _, r in df_exploded.iterrows():
            vsmap[r["Vendor"]].add(r["Service"])
        full_cover = [v for v,s in vsmap.items() if svc_set.issubset(s)]
        partial    = [v for v,s in vsmap.items() if svc_set & s and v not in full_cover]
        if full_cover:
            return {
                "type": "combination",
                "text": ("✅ **{}** vendor(s) have quoted for ALL of: {}\n\n"
                         "**Full coverage:** {}\n\n"
                         "**Partial coverage:** {}".format(
                             len(full_cover),
                             " · ".join(["**{}**".format(s) for s in svc_set]),
                             ", ".join(full_cover),
                             ", ".join(partial) if partial else "None")),
                "services": list(svc_set),
                "full_cover": full_cover,
                "partial": partial,
            }
        return {
            "type": "combination",
            "text": ("⚠️ **No single vendor** has quoted for all of: {}\n\n"
                     "**Partial coverage:** {}".format(
                         " · ".join(["**{}**".format(s) for s in svc_set]),
                         ", ".join(partial) if partial else "None")),
            "services": list(svc_set),
            "full_cover": [],
            "partial": partial,
        }

    # ── Service info (without compare keyword) ──
    if matched_services:
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc]
        if not d.empty:
            vendors  = d["Vendor"].unique().tolist()
            n_quotes = d["File Name"].nunique()
            prices   = []
            for _, r in d.drop_duplicates(subset=["File Name"]).iterrows():
                qp = _parse_num(str(r.get("Quoted Price","")).strip())
                if qp > 0: prices.append(qp)
            price_summary = ""
            if prices:
                price_summary = "\n\n💰 **Price range:** {} — {} (avg: {})".format(
                    _fmt(min(prices)), _fmt(max(prices)), _fmt(sum(prices)/len(prices)))
            return {
                "type": "service_info",
                "text": ("📋 **{}**\n\n"
                         "🏢 **{}** vendor(s) have quoted: {}\n"
                         "📄 **{}** total quote files{}\n\n"
                         "💡 Ask *'compare {} across vendors'* for price details.".format(
                             svc, len(vendors), ", ".join(vendors[:5]),
                             n_quotes, price_summary, svc)),
                "service": svc,
                "vendors": vendors,
            }

    # ── Fallback ──
    suggestions = []
    if matched_services: suggestions.append("*'Compare {}?'*".format(matched_services[0]))
    if matched_vendor:   suggestions.append("*'What does {} offer?'*".format(matched_vendor))
    if matched_cat:      suggestions.append("*'Show {} quotes'*".format(matched_cat))
    fallback = ("I couldn't find specific data for your query. "
                "Try rephrasing or ask about a specific vendor, service, or category.")
    if suggestions: fallback += "\n\n💡 Did you mean: " + " or ".join(suggestions)
    return {"type":"text","text":fallback}

def render_chat_response(resp):
    """Render a chatbot response with optional chart."""
    st.markdown(resp["text"])

    if resp["type"] == "comparison" and "vendor_prices" in resp:
        vp  = resp["vendor_prices"]
        avg = resp.get("avg", 0)
        df_chart = pd.DataFrame([
            {"Vendor":v, "Price":p,
             "Color": PwC_GREEN if v==resp.get("best_vendor") else PwC_ORANGE}
            for v,p in sorted(vp.items(), key=lambda x:x[1])])
        if not df_chart.empty:
            fig = go.Figure(go.Bar(
                x=df_chart["Vendor"], y=df_chart["Price"],
                marker_color=df_chart["Color"],
                marker_line_width=0,
                text=df_chart["Price"].apply(_fmt),
                textposition="outside"))
            if avg > 0:
                fig.add_hline(y=avg, line_dash="dash", line_color=PwC_NAVY, line_width=2,
                               annotation_text="Avg: {}".format(_fmt(avg)),
                               annotation_position="top right")
            fig.update_layout(
                height=280, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5,r=10,t=20,b=10), font=CFONT,
                yaxis=dict(title="Price (USD)", showgrid=True, gridcolor="#E0E0E0", zeroline=False),
                xaxis=dict(tickangle=-15), bargap=0.4, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════
# UI HELPERS
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
                "text-transform:uppercase;color:#003078;margin-bottom:6px'>Procurement Insight</div>"
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
                "<h1 style='margin:0;font-size:1.4em;font-weight:700;color:white;"
                "font-family:Georgia,serif'>{}</h1>"
                "{}</div>".format(eyebrow, title,
                "<p style='margin:6px 0 0;opacity:0.6;font-size:0.85em'>{}</p>".format(subtitle) if subtitle else ""),
                unsafe_allow_html=True)

def static_section_header(title, border_color=PwC_ORANGE):
    st.markdown("<div style='border:1px solid #e0e0e0;border-radius:4px;padding:10px 16px;"
                "margin-bottom:4px;background:white;border-left:4px solid {}'>"
                "<span style='font-weight:700;font-size:0.94em;color:#2D2D2D'>{}</span>"
                "</div>".format(border_color, title), unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# INIT SESSION STATE
# ════════════════════════════════════════════════════════════
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []
if "gh_prices_loaded" not in st.session_state:
    st.session_state["gh_prices_loaded"] = False
if "real_analysis_done" not in st.session_state:
    st.session_state["real_analysis_done"] = False

# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state and
        st.session_state["uploaded_catalog_df"] is not None):
    df_master   = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "file"

NO_DATA = (df_master is None or df_exploded is None or df_master.empty)

vendor_color_map = {}
if not NO_DATA:
    for i,v in enumerate(sorted(df_master["Vendor"].unique())):
        vendor_color_map[v] = CHART_SEQ[i % len(CHART_SEQ)]

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
selected_cat = "All"; selected_vendor = "All"; d_filt = pd.DataFrame()

if not NO_DATA:
    with st.sidebar:
        st.markdown("<div style='text-align:center;padding:20px 0 14px'>"
                    "<div style='font-size:2em'>📋</div>"
                    "<div style='font-size:1.05em;font-weight:700;color:white;margin:5px 0 2px;"
                    "font-family:Georgia,serif'>IT Procurement</div>"
                    "<div style='font-size:0.70em;color:#aaa;letter-spacing:1px;text-transform:uppercase'>"
                    "Intelligence Dashboard</div></div>"
                    "<hr style='border-color:#D04A02;border-width:2px;margin:0 0 14px'>",
                    unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown("<div style='background:#D04A02;color:white;padding:6px 10px;border-radius:2px;"
                        "font-size:0.75em;font-weight:700;text-align:center;margin-bottom:10px'>"
                        "📤 USING UPLOADED CATALOG</div>", unsafe_allow_html=True)

        sb_label("📂 Category")
        all_cats = ["All"] + sorted([c for c in df_master["Category"].unique() if str(c).strip() not in ["","nan"]])
        selected_cat = st.selectbox("Category", all_cats, label_visibility="collapsed")

        sb_label("🏢 Vendor")
        vpool = df_master if selected_cat=="All" else df_master[df_master["Category"]==selected_cat]
        all_vendors = ["All"] + sorted([v for v in vpool["Vendor"].unique() if str(v).strip() not in ["","nan"]])
        selected_vendor = st.selectbox("Vendor", all_vendors, label_visibility="collapsed")

        st.markdown("<hr style='border-color:#555;margin:12px 0'>", unsafe_allow_html=True)
        st.markdown("<p style='color:#888;font-size:0.78em;margin:2px 0'>"
                    "📄 {} quotes | 🛠 {} services | 🏢 {} vendors</p>".format(
                        len(df_master), df_exploded["Service"].nunique(), df_master["Vendor"].nunique()),
                    unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown("<hr style='border-color:#555;margin:12px 0'>", unsafe_allow_html=True)
            if st.button("🔄 Reset to Default Catalog", use_container_width=True):
                st.session_state["uploaded_catalog_df"]  = None
                st.session_state["uploaded_catalog_exp"] = None
                st.rerun()

    d_filt = df_exploded.copy()
    if selected_cat    != "All": d_filt = d_filt[d_filt["Category"]==selected_cat]
    if selected_vendor != "All": d_filt = d_filt[d_filt["Vendor"]==selected_vendor]

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
pwc_header("Procurement Intelligence Dashboard",
           "Catalog overview · Browse & chat · Upload & score · Vendor analysis")

# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    use_filt = d_filt if not d_filt.empty else df_exploded
    k1,k2,k3,k4 = st.columns(4)
    kpi(k1, df_master["File Name"].nunique(), "Total Quotes",    PwC_ORANGE)
    kpi(k2, df_exploded["Service"].nunique(), "Unique Services", PwC_NAVY)
    kpi(k3, df_master["Vendor"].nunique(),    "Vendors",         PwC_TEAL)
    kpi(k4, df_master["Category"].nunique(),  "Categories",      PwC_DARK)
    st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS  — Analytics tab REMOVED
# ════════════════════════════════════════════════════════════
tab0,tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
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
        st.info("No catalog loaded. Go to 🗂 Upload Catalog tab.")
    else:
        df_ov = df_master.copy()
        df_ov["Subcategory"] = df_ov.apply(
            lambda r: infer_subcategory(r.get("Category",""),r.get("Comments",""),r.get("File Name","")), axis=1)
        all_cats_ov = sorted([c for c in df_ov["Category"].unique() if str(c).strip() not in ["","nan"]])

        pwc_header("Catalog Overview",
                   "Categories · Subcategories · Vendors · Services · Quotations")

        k0a,k0b,k0c,k0d,k0e = st.columns(5)
        mini_kpi(k0a, len(df_ov),                    "Total Quotations", PwC_ORANGE,"📄")
        mini_kpi(k0b, df_ov["Vendor"].nunique(),      "Unique Vendors",   PwC_NAVY,  "🏢")
        mini_kpi(k0c, df_ov["Category"].nunique(),    "Categories",       PwC_TEAL,  "📂")
        mini_kpi(k0d, df_ov["Subcategory"].nunique(), "Subcategories",    PwC_AMBER, "🏷️")
        mini_kpi(k0e, df_exploded["Service"].nunique() if df_exploded is not None else "—",
                 "Unique Services", PwC_DARK, "🛠")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Category cards ──
        section_title("CATEGORIES AT A GLANCE")
        cat_stats = []
        for cat in all_cats_ov:
            d_cat = df_ov[df_ov["Category"]==cat]
            n_svcs = df_exploded[df_exploded["Category"]==cat]["Service"].nunique() if df_exploded is not None else 0
            cat_stats.append({"Category":cat,"Quotations":len(d_cat),
                               "Vendors":d_cat["Vendor"].nunique(),
                               "Subcategories":d_cat["Subcategory"].nunique(),"Services":n_svcs})
        cat_stats_df = pd.DataFrame(cat_stats).sort_values("Quotations", ascending=False)

        CAT_ICONS = {"Cybersecurity":"🛡️","Network & Telecom":"🌐","Hosting":"🖥️",
                     "M365 & Power Platform":"☁️","IdAM":"🔑","Service Management (SNow)":"⚙️",
                     "Summary & Reporting":"📊"}

        rows_of_3 = [cat_stats_df.iloc[i:i+3] for i in range(0, len(cat_stats_df), 3)]
        for row_chunk in rows_of_3:
            cols = st.columns(len(row_chunk), gap="medium")
            for col_idx,(_, row_s) in enumerate(row_chunk.iterrows()):
                cat_name = row_s["Category"]
                icon  = CAT_ICONS.get(cat_name,"📁")
                color = CHART_SEQ[all_cats_ov.index(cat_name) % len(CHART_SEQ)] if cat_name in all_cats_ov else PwC_GREY
                cols[col_idx].markdown(
                    "<div style='background:white;border:1px solid #e0e0e0;border-radius:6px;"
                    "padding:16px 18px;border-top:4px solid {};height:100%'>"
                    "<div style='font-size:1.4em;margin-bottom:6px'>{}</div>"
                    "<div style='font-size:0.92em;font-weight:700;color:#2D2D2D;"
                    "margin-bottom:10px;font-family:Georgia,serif'>{}</div>"
                    "<div style='display:flex;gap:8px;flex-wrap:wrap'>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;"
                    "font-size:0.76em;font-weight:700;color:#2D2D2D'>📄 {} quotes</span>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;"
                    "font-size:0.76em;font-weight:700;color:#295477'>🏢 {} vendors</span>"
                    "<span style='background:#F3F3F3;border-radius:3px;padding:3px 8px;"
                    "font-size:0.76em;font-weight:700;color:#299D8F'>🛠 {} services</span>"
                    "</div></div>".format(color,icon,cat_name,
                                          row_s["Quotations"],row_s["Vendors"],row_s["Services"]),
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Charts — PwC single colour sequences ──
        section_title("DISTRIBUTION — QUOTATIONS & VENDORS BY CATEGORY")
        ch_a, ch_b = st.columns(2, gap="large")
        with ch_a:
            # Single colour — PwC Orange for quotations bar
            fig_cq = go.Figure(go.Bar(
                x=cat_stats_df["Category"], y=cat_stats_df["Quotations"],
                marker_color=PwC_ORANGE, marker_line_width=0,
                text=cat_stats_df["Quotations"], textposition="outside",
                textfont=dict(size=11, color=PwC_DARK)))
            fig_cq.update_layout(
                height=340, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                title=dict(text="Quotations per Category",
                           font=dict(size=13,color=PwC_DARK,family="Georgia,serif"),x=0),
                yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False,
                           tickfont=dict(color=PwC_DARK)),
                xaxis=dict(tickangle=-30,tickfont=dict(size=9.5,color=PwC_DARK)),
                bargap=0.35, showlegend=False)
            st.plotly_chart(fig_cq, use_container_width=True)
        with ch_b:
            # PwC Navy for vendors
            fig_cv = go.Figure(go.Bar(
                x=cat_stats_df["Category"], y=cat_stats_df["Vendors"],
                marker_color=PwC_NAVY, marker_line_width=0,
                text=cat_stats_df["Vendors"], textposition="outside",
                textfont=dict(size=11, color=PwC_DARK)))
            fig_cv.update_layout(
                height=340, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                title=dict(text="Vendors per Category",
                           font=dict(size=13,color=PwC_DARK,family="Georgia,serif"),x=0),
                yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                xaxis=dict(tickangle=-30,tickfont=dict(size=9.5)),
                bargap=0.35, showlegend=False)
            st.plotly_chart(fig_cv, use_container_width=True)

        # ── Donut — PwC standardised palette ──
        st.markdown("<br>", unsafe_allow_html=True)
        section_title("CATALOG COMPOSITION — QUOTE SHARE BY CATEGORY")
        d_left, d_right = st.columns([1,2], gap="large")
        with d_left:
            fig_donut = px.pie(cat_stats_df, values="Quotations", names="Category",
                               hole=0.55, color_discrete_sequence=CHART_SEQ)
            fig_donut.update_traces(textposition="outside", textinfo="percent+label",
                                     textfont_size=10, pull=[0.03]*len(cat_stats_df))
            fig_donut.update_layout(height=360, margin=dict(l=10,r=10,t=10,b=10),
                                     paper_bgcolor=CBG, font=CFONT, showlegend=False)
            st.plotly_chart(fig_donut, use_container_width=True)
        with d_right:
            tbl_rows = ["<table class='comp-table'><thead><tr>"
                        "<th>Category</th><th style='text-align:center'>Quotes</th>"
                        "<th style='text-align:center'>Vendors</th>"
                        "<th style='text-align:center'>Services</th>"
                        "<th style='text-align:center'>Subcategories</th>"
                        "</tr></thead><tbody>"]
            for _, row_s in cat_stats_df.iterrows():
                cat_n = row_s["Category"]
                color = CHART_SEQ[all_cats_ov.index(cat_n)%len(CHART_SEQ)] if cat_n in all_cats_ov else PwC_GREY
                icon  = CAT_ICONS.get(cat_n,"📁")
                tbl_rows.append(
                    "<tr><td><span style='border-left:4px solid {};padding-left:8px;font-weight:600'>"
                    "{} {}</span></td>"
                    "<td style='text-align:center;font-weight:700;color:#D04A02'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#003078'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#299D8F'>{}</td>"
                    "<td style='text-align:center;font-weight:700;color:#EB8C00'>{}</td>"
                    "</tr>".format(color,icon,cat_n,
                                   row_s["Quotations"],row_s["Vendors"],
                                   row_s["Services"],row_s["Subcategories"]))
            tbl_rows.append("</tbody></table>")
            st.markdown("".join(tbl_rows), unsafe_allow_html=True)

        # ── Per-category drill-down ──
        st.markdown("<br>", unsafe_allow_html=True)
        section_title("DRILL-DOWN BY CATEGORY")
        ordered_cats = ["Cybersecurity"] + [c for c in all_cats_ov if c != "Cybersecurity"]
        cat_subtabs  = st.tabs(["🛡️ {}".format(c) if c=="Cybersecurity" else c for c in ordered_cats])

        for tab_idx, cat_name in enumerate(ordered_cats):
            with cat_subtabs[tab_idx]:
                d_cat = df_ov[df_ov["Category"]==cat_name].copy()
                if d_cat.empty: st.info("No data for {}.".format(cat_name)); continue

                ck1,ck2,ck3,ck4 = st.columns(4)
                n_q   = len(d_cat); n_v = d_cat["Vendor"].nunique()
                n_sub = d_cat["Subcategory"].nunique()
                n_svc = df_exploded[df_exploded["Category"]==cat_name]["Service"].nunique() if df_exploded is not None else 0
                mini_kpi(ck1,n_q,  "Quotations",  PwC_ORANGE,"📄")
                mini_kpi(ck2,n_v,  "Vendors",      PwC_NAVY,  "🏢")
                mini_kpi(ck3,n_svc,"Services",      PwC_TEAL,  "🛠")
                mini_kpi(ck4,n_sub,"Subcategories", PwC_AMBER, "🏷️")
                st.markdown("<br>", unsafe_allow_html=True)

                sub_stats = (d_cat.groupby("Subcategory")
                             .agg(Quotations=("File Name","count"), Vendors=("Vendor","nunique"))
                             .reset_index().sort_values("Quotations", ascending=False))

                sc1, sc2 = st.columns([1,1], gap="medium")
                with sc1:
                    st.markdown("<div style='font-size:0.78em;font-weight:700;text-transform:uppercase;"
                                "color:#2D2D2D;margin-bottom:8px'>Subcategories</div>", unsafe_allow_html=True)
                    fig_sub = go.Figure(go.Bar(
                        x=sub_stats["Quotations"], y=sub_stats["Subcategory"], orientation="h",
                        marker_color=PwC_ORANGE, marker_line_width=0,
                        text=sub_stats["Quotations"], textposition="outside"))
                    fig_sub.update_layout(
                        height=max(220,len(sub_stats)*38), plot_bgcolor=CBG, paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=10,b=10), font=CFONT,
                        xaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False,title="Quotations"),
                        yaxis=dict(autorange="reversed",tickfont=dict(size=10)), bargap=0.3)
                    st.plotly_chart(fig_sub, use_container_width=True)
                with sc2:
                    st.markdown("<div style='font-size:0.78em;font-weight:700;text-transform:uppercase;"
                                "color:#2D2D2D;margin-bottom:8px'>Vendors</div>", unsafe_allow_html=True)
                    vnd_stats = (d_cat.groupby("Vendor").agg(Quotations=("File Name","count"))
                                 .reset_index().sort_values("Quotations", ascending=False))
                    fig_vnd = go.Figure(go.Bar(
                        x=vnd_stats["Quotations"], y=vnd_stats["Vendor"], orientation="h",
                        marker_color=PwC_NAVY, marker_line_width=0,
                        text=vnd_stats["Quotations"], textposition="outside"))
                    fig_vnd.update_layout(
                        height=max(220,len(vnd_stats)*38), plot_bgcolor=CBG, paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=10,b=10), font=CFONT,
                        xaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False,title="Quotations"),
                        yaxis=dict(autorange="reversed",tickfont=dict(size=10)), bargap=0.3)
                    st.plotly_chart(fig_vnd, use_container_width=True)

                # File list
                static_section_header("📄 All {} quotation files in {}".format(len(d_cat), cat_name), PwC_NAVY)
                with st.container():
                    file_tbl = ["<table class='comp-table'><thead><tr>"
                                "<th>File Name</th><th>Vendor</th><th>Subcategory</th>"
                                "<th>Services / Comments</th></tr></thead><tbody>"]
                    for fi,(_, fr) in enumerate(d_cat.sort_values("Subcategory").iterrows()):
                        bg  = "white" if fi%2==0 else "#F3F3F3"
                        vc  = vendor_color_map.get(fr["Vendor"], PwC_GREY)
                        cmt = str(fr.get("Comments","")).replace("\n"," · ")[:80]
                        file_tbl.append(
                            "<tr style='background:{}'>"
                            "<td style='font-family:monospace;font-size:0.78em;word-break:break-all'>{}</td>"
                            "<td>{}</td>"
                            "<td style='color:#555;font-size:0.82em'>{}</td>"
                            "<td style='font-size:0.80em'>{}</td></tr>".format(
                                bg, fr.get("File Name",""),
                                vendor_pill(fr["Vendor"],vc), fr["Subcategory"], cmt))
                    file_tbl.append("</tbody></table>")
                    st.markdown("".join(file_tbl), unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# TAB 1 — BROWSE & VERDICT  (with conversational chatbot)
# ════════════════════════════════════════════════════════════
with tab1:
    if NO_DATA:
        st.info("No catalog loaded. Go to 🗂 Upload Catalog tab.")
    else:
        pwc_header("Browse & Verdict",
                   "Ask about any service, vendor or price — or explore manually below",
                   "Procurement Intelligence")

        # ════════════════════════════════════════
        # CHATBOT SECTION
        # ════════════════════════════════════════
        st.markdown("<div style='background:#003078;color:white;padding:12px 18px;"
                    "border-radius:4px 4px 0 0;margin-bottom:0'>"
                    "<span style='font-weight:700;font-size:0.95em;font-family:Georgia,serif'>"
                    "💬 Procurement Assistant</span>"
                    "<span style='font-size:0.78em;opacity:0.7;margin-left:10px'>"
                    "Ask me anything about services, vendors or prices</span>"
                    "</div>", unsafe_allow_html=True)

        # Chat history display
        chat_html = "<div class='chat-container'>"
        if not st.session_state["chat_history"]:
            chat_html += ("<div class='chat-wrap'><div class='chat-bot'>"
                          "👋 Hello! I'm your <b>PwC Procurement Assistant</b>.<br><br>"
                          "Ask me things like:<br>"
                          "• <i>'Who has quoted for Cisco Catalyst?'</i><br>"
                          "• <i>'Compare Palo Alto prices across vendors'</i><br>"
                          "• <i>'What does NTT Data offer?'</i><br>"
                          "• <i>'Is there a quote for VMware and NetApp together?'</i>"
                          "</div></div>")
        else:
            for turn in st.session_state["chat_history"]:
                chat_html += "<div class='chat-wrap'><div class='chat-user'>{}</div></div>".format(
                    turn["user"])
                chat_html += "<div class='chat-wrap'><div class='chat-bot'>{}</div></div>".format(
                    turn["bot_text"].replace("\n","<br>").replace("**","<b>").replace("**","</b>"))
        chat_html += "</div>"
        st.markdown(chat_html, unsafe_allow_html=True)

        # Render last chart if any
        if st.session_state["chat_history"]:
            last = st.session_state["chat_history"][-1]
            if last.get("bot_resp"):
                render_chat_response(last["bot_resp"])

        # Input
        with st.form("chat_form", clear_on_submit=True):
            col_inp, col_btn = st.columns([5,1])
            with col_inp:
                user_input = st.text_input("Ask a question…", placeholder="e.g. Who quoted for Cisco Catalyst C8300?",
                                            label_visibility="collapsed")
            with col_btn:
                submitted = st.form_submit_button("Send", type="primary", use_container_width=True)

        if submitted and user_input.strip():
            resp = chatbot_response(user_input.strip(), df_master, df_exploded)
            st.session_state["chat_history"].append({
                "user": user_input.strip(),
                "bot_text": resp["text"],
                "bot_resp": resp,
            })
            st.rerun()

        # Quick suggestion chips
        st.markdown("<div style='margin:8px 0 4px;font-size:0.78em;color:#7D7D7D;font-weight:600;"
                    "letter-spacing:0.5px'>QUICK QUESTIONS</div>", unsafe_allow_html=True)
        suggestions = [
            "Who quoted for Cisco Catalyst?",
            "Compare Palo Alto prices",
            "What does NTT Data offer?",
            "List all vendors",
            "Show Network & Telecom quotes",
            "Cisco and VMware together?",
        ]
        chip_cols = st.columns(len(suggestions))
        for i, sugg in enumerate(suggestions):
            if chip_cols[i].button(sugg, key="chip_{}".format(i), use_container_width=True):
                resp = chatbot_response(sugg, df_master, df_exploded)
                st.session_state["chat_history"].append({
                    "user": sugg,
                    "bot_text": resp["text"],
                    "bot_resp": resp,
                })
                st.rerun()

        if st.button("🗑 Clear conversation", key="clear_chat"):
            st.session_state["chat_history"] = []
            st.rerun()

        # ════════════════════════════════════════
        # MANUAL BROWSE SECTION
        # ════════════════════════════════════════
        st.markdown("<hr style='border:none;border-top:2px solid #D04A02;margin:28px 0 16px'>",
                    unsafe_allow_html=True)
        section_title("MANUAL SERVICE BROWSER",
                      "Select services below to see detailed quotation data and price comparisons")

        # Service selector
        avail_svcs = sorted([s for s in df_exploded["Service"].unique()
                              if str(s).strip() not in ["","nan"]])
        svc_search = st.text_input("🔍 Filter services", placeholder="Type to filter…",
                                    label_visibility="collapsed", key="manual_svc_search")
        if svc_search:
            avail_svcs = [s for s in avail_svcs if svc_search.lower() in s.lower()]

        selected_svcs = st.multiselect("Select services to compare",
                                        options=avail_svcs, default=[],
                                        label_visibility="collapsed",
                                        help="Select one or more services to see vendor quotes")

        if not selected_svcs:
            # Show summary of all services with vendor count
            section_title("ALL SERVICES AT A GLANCE")
            svc_summary = (df_exploded.groupby("Service")["Vendor"].nunique()
                           .reset_index().sort_values("Vendor", ascending=False))
            svc_summary.columns = ["Service","Vendor Count"]
            svc_summary["Has Multiple Vendors"] = svc_summary["Vendor Count"] > 1

            # Chart — PwC Orange for competitive, PwC Grey for single vendor
            svc_top = svc_summary.head(25)
            fig_svc = go.Figure(go.Bar(
                x=svc_top["Vendor Count"],
                y=svc_top["Service"].apply(lambda x: x[:50]),
                orientation="h",
                marker_color=[PwC_ORANGE if v>1 else PwC_GREY for v in svc_top["Vendor Count"]],
                marker_line_width=0,
                text=svc_top["Vendor Count"],
                textposition="outside",
                textfont=dict(size=10)))
            fig_svc.update_layout(
                height=600, plot_bgcolor=CBG, paper_bgcolor=CBG,
                margin=dict(l=5,r=40,t=30,b=10), font=CFONT,
                title=dict(text="Services by Number of Vendors (Orange = Multiple Vendors)",
                           font=dict(size=12,color=PwC_DARK,family="Georgia,serif"),x=0),
                xaxis=dict(title="Number of Vendors",showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                yaxis=dict(autorange="reversed",tickfont=dict(size=9)),
                bargap=0.3, showlegend=False)
            st.plotly_chart(fig_svc, use_container_width=True)

            # Stats
            n_multi  = svc_summary[svc_summary["Vendor Count"]>1].shape[0]
            n_single = svc_summary[svc_summary["Vendor Count"]==1].shape[0]
            ms1,ms2,ms3 = st.columns(3)
            kpi(ms1, len(svc_summary),  "Total Services",          PwC_DARK)
            kpi(ms2, n_multi,           "Competitive (2+ vendors)", PwC_ORANGE)
            kpi(ms3, n_single,          "Single Vendor",            PwC_GREY)

        else:
            # ── Service selected — show verdict ──
            use_filt2 = d_filt if not d_filt.empty else df_exploded
            d_sel = use_filt2[use_filt2["Service"].isin(selected_svcs)].copy()

            if d_sel.empty:
                st.warning("No quotations found for the selected service(s).")
            else:
                # Collect prices
                vendor_prices_map = {}
                for _, r in d_sel.drop_duplicates(subset=["Vendor","File Name"]).iterrows():
                    v  = r["Vendor"]
                    qp = _parse_num(str(r.get("Quoted Price","")).strip())
                    ck = "px_{}".format(str(r.get("File Name","")).strip())
                    ca = st.session_state.get(ck)
                    ep = ca["price_num"] if ca else 0.0
                    ref = ep if ep > 0 else qp
                    if ref > 0:
                        vendor_prices_map[v] = min(vendor_prices_map.get(v,ref), ref)
                if not vendor_prices_map:
                    for _, r in d_sel.drop_duplicates(subset=["Vendor"]).iterrows():
                        v  = r["Vendor"]
                        qp = _parse_num(str(r.get("Quoted Price","")).strip())
                        if qp > 0 and v not in vendor_prices_map:
                            vendor_prices_map[v] = qp

                # ── Coverage verdict banner ──
                vsmap = defaultdict(set)
                for _, r in d_sel.iterrows(): vsmap[r["Vendor"]].add(r["Service"])
                full_cover = [v for v,s in vsmap.items() if set(selected_svcs).issubset(s)]
                partial    = [v for v,s in vsmap.items() if set(selected_svcs) & s and v not in full_cover]

                if len(selected_svcs) == 1:
                    n_vendors = d_sel["Vendor"].nunique()
                    if n_vendors > 1:
                        st.markdown("<div class='verdict-green'>"
                                    "✅ <b>{}</b> has been previously quoted by <b>{} vendors</b> — "
                                    "competitive benchmarking available."
                                    "</div>".format(selected_svcs[0], n_vendors), unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='verdict-yellow'>"
                                    "⚠️ <b>{}</b> has only <b>1 vendor</b> quotation — "
                                    "limited benchmarking. Consider sourcing additional quotes."
                                    "</div>".format(selected_svcs[0]), unsafe_allow_html=True)
                else:
                    if full_cover:
                        st.markdown("<div class='verdict-green'>"
                                    "✅ <b>{} vendor(s)</b> cover ALL {} selected services: <b>{}</b>"
                                    "</div>".format(len(full_cover), len(selected_svcs),
                                                    ", ".join(full_cover)), unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='verdict-yellow'>"
                                    "⚠️ No single vendor covers all {} selected services. "
                                    "Partial coverage: <b>{}</b>"
                                    "</div>".format(len(selected_svcs),
                                                    ", ".join(partial) if partial else "None"),
                                    unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)

                # ── Price verdict cards ──
                if vendor_prices_map:
                    avg_p   = sum(vendor_prices_map.values()) / len(vendor_prices_map)
                    best_v  = min(vendor_prices_map, key=vendor_prices_map.get)
                    worst_v = max(vendor_prices_map, key=vendor_prices_map.get)
                    spread  = round((max(vendor_prices_map.values())-min(vendor_prices_map.values()))
                                    /min(vendor_prices_map.values())*100, 1) if min(vendor_prices_map.values()) > 0 else 0

                    section_title("PRICE VERDICT")
                    pv1,pv2,pv3,pv4 = st.columns(4)
                    pv1.markdown("<div class='score-card green'>"
                                 "<div style='font-size:0.68em;font-weight:700;text-transform:uppercase;"
                                 "color:#22992E'>Best Price</div>"
                                 "<div style='font-size:1.6em;font-weight:800;color:#22992E;"
                                 "font-family:Georgia,serif'>{}</div>"
                                 "<div style='font-size:0.75em;color:#555;margin-top:2px'>{}</div>"
                                 "</div>".format(_fmt(min(vendor_prices_map.values())),best_v),
                                 unsafe_allow_html=True)
                    pv2.markdown("<div class='score-card navy'>"
                                 "<div style='font-size:0.68em;font-weight:700;text-transform:uppercase;"
                                 "color:#003078'>Market Average</div>"
                                 "<div style='font-size:1.6em;font-weight:800;color:#003078;"
                                 "font-family:Georgia,serif'>{}</div>"
                                 "<div style='font-size:0.75em;color:#555;margin-top:2px'>"
                                 "{} vendor(s)</div>"
                                 "</div>".format(_fmt(avg_p), len(vendor_prices_map)),
                                 unsafe_allow_html=True)
                    pv3.markdown("<div class='score-card red'>"
                                 "<div style='font-size:0.68em;font-weight:700;text-transform:uppercase;"
                                 "color:#E0301E'>Highest Quote</div>"
                                 "<div style='font-size:1.6em;font-weight:800;color:#E0301E;"
                                 "font-family:Georgia,serif'>{}</div>"
                                 "<div style='font-size:0.75em;color:#555;margin-top:2px'>{}</div>"
                                 "</div>".format(_fmt(max(vendor_prices_map.values())),worst_v),
                                 unsafe_allow_html=True)
                    spread_css = "red" if spread>20 else "yellow" if spread>10 else "green"
                    pv4.markdown("<div class='score-card {}'>"
                                 "<div style='font-size:0.68em;font-weight:700;text-transform:uppercase;"
                                 "color:{}'>Price Spread</div>"
                                 "<div style='font-size:1.6em;font-weight:800;font-family:Georgia,serif;"
                                 "color:{}'>{}%</div>"
                                 "<div style='font-size:0.75em;color:#555;margin-top:2px'>"
                                 "negotiation room</div>"
                                 "</div>".format(
                                     spread_css,
                                     PwC_RED if spread>20 else PwC_AMBER if spread>10 else PwC_GREEN,
                                     PwC_RED if spread>20 else PwC_AMBER if spread>10 else PwC_GREEN,
                                     spread), unsafe_allow_html=True)

                    st.markdown("<br>", unsafe_allow_html=True)

                    # ── Price comparison chart — PwC colours ──
                    section_title("PRICE COMPARISON CHART",
                                  "Green = most competitive · Red = most expensive · Navy line = average")
                    chart_data = sorted(vendor_prices_map.items(), key=lambda x: x[1])
                    bar_colors = []
                    for v,p in chart_data:
                        if v == best_v:   bar_colors.append(PwC_GREEN)
                        elif v == worst_v: bar_colors.append(PwC_RED)
                        else:              bar_colors.append(PwC_ORANGE)

                    fig_cmp = go.Figure()
                    fig_cmp.add_trace(go.Bar(
                        x=[v for v,_ in chart_data],
                        y=[p for _,p in chart_data],
                        marker_color=bar_colors,
                        marker_line_width=0,
                        text=[_fmt(p) for _,p in chart_data],
                        textposition="outside",
                        textfont=dict(size=11, color=PwC_DARK)))
                    fig_cmp.add_hline(y=avg_p, line_dash="dash", line_color=PwC_NAVY, line_width=2,
                                       annotation_text="Market Avg: {}".format(_fmt(avg_p)),
                                       annotation_font_color=PwC_NAVY,
                                       annotation_position="top right")
                    fig_cmp.update_layout(
                        height=340, plot_bgcolor=CBG, paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                        yaxis=dict(title="Quoted Price (USD)", showgrid=True,
                                   gridcolor="#E0E0E0", zeroline=False),
                        xaxis=dict(tickangle=-15, tickfont=dict(size=11)),
                        bargap=0.4, showlegend=False)
                    st.plotly_chart(fig_cmp, use_container_width=True)

                    # ── Vendor score table ──
                    section_title("VENDOR SCORE CARD")
                    all_p_vals = list(vendor_prices_map.values())
                    rows = ["<table class='comp-table'><thead><tr>"
                            "<th>Rank</th><th>Vendor</th><th>Price</th>"
                            "<th>vs Market Avg</th><th>Score</th>"
                            "<th>Verdict</th></tr></thead><tbody>"]
                    for rank,(v,p) in enumerate(sorted(vendor_prices_map.items(), key=lambda x:x[1]),1):
                        bg  = "white" if rank%2==0 else "#F3F3F3"
                        vc  = vendor_color_map.get(v, PwC_GREY)
                        others = [px for px in all_p_vals if px != p]
                        ps  = None
                        if p > 0 and others: ps,_,_,_,_ = price_score(p, others)
                        sc  = score_color(ps)
                        pct = round((p-avg_p)/avg_p*100,1) if avg_p > 0 else 0
                        vs_txt = ("{}% below avg ✅".format(abs(pct)) if pct<0
                                  else "{}% above avg ⚠️".format(abs(pct)) if pct>0 else "At average")
                        vs_col = PwC_GREEN if pct<0 else PwC_RED if pct>10 else PwC_AMBER
                        vt,_,vc_col = get_verdict(ps)
                        medal = "🥇" if rank==1 else "🥈" if rank==2 else "🥉" if rank==3 else str(rank)
                        rows.append("<tr style='background:{}'>"
                                    "<td style='text-align:center;font-size:1.1em'>{}</td>"
                                    "<td>{}</td>"
                                    "<td style='font-family:monospace;font-weight:700;"
                                    "font-size:1.05em'>{}</td>"
                                    "<td style='color:{};font-weight:600'>{}</td>"
                                    "<td style='text-align:center'>"
                                    "<span style='font-weight:800;font-size:1.15em;color:{}'>"
                                    "{}</span></td>"
                                    "<td><span style='color:{};font-weight:700'>{}</span></td>"
                                    "</tr>".format(bg, medal, vendor_pill(v,vc), _fmt(p),
                                                   vs_col, vs_txt, sc,
                                                   ps if ps is not None else "—",
                                                   vc_col, vt))
                    rows.append("</tbody></table>")
                    st.markdown("".join(rows), unsafe_allow_html=True)

                # ── Per-service file details ──
                st.markdown("<br>", unsafe_allow_html=True)
                section_title("QUOTATION FILE DETAILS")
                for svc in selected_svcs:
                    d_svc = (d_sel[d_sel["Service"]==svc]
                             .drop_duplicates(subset=["Vendor","File Name"])
                             .sort_values("Vendor"))
                    n_v   = d_svc["Vendor"].nunique()
                    has_price = "Quoted Price" in d_svc.columns

                    static_section_header(
                        "{}  ·  {} vendor(s)  ·  {} file(s)  ·  {}".format(
                            svc, n_v, len(d_svc),
                            "✅ COMPETITIVE" if n_v > 1 else "⚠️ SINGLE VENDOR"),
                        PwC_ORANGE if n_v > 1 else PwC_AMBER)

                    with st.container():
                        all_prices_svc = []
                        for _,r in d_svc.iterrows():
                            qp = _parse_num(str(r.get("Quoted Price","")).strip())
                            if qp > 0: all_prices_svc.append(qp)

                        tbl_rows = ["<table class='comp-table'><thead><tr>"
                                    "<th>Vendor</th><th>Category</th><th>File Name</th>"]
                        if has_price: tbl_rows.append("<th>Quoted Price</th>")
                        tbl_rows.append("<th>Extracted</th><th>Score</th><th>Verdict</th>"
                                        "<th>Open</th></tr></thead><tbody>")

                        for i,(_,row) in enumerate(d_svc.iterrows()):
                            bg    = "white" if i%2==0 else "#F3F3F3"
                            vc2   = vendor_color_map.get(row["Vendor"],PwC_GREY)
                            fname = str(row.get("File Name","")).strip()
                            url   = resolve_url(row)
                            qp_str= str(row.get("Quoted Price","")).strip()
                            qp_num= _parse_num(qp_str)
                            ck    = "px_{}".format(fname)
                            cached= st.session_state.get(ck)
                            ep_fmt= _fmt(cached["price"]) if cached and cached.get("price_num",0)>0 else "—"
                            ref   = (cached["price_num"] if cached and cached.get("price_num",0)>0
                                     else qp_num if qp_num > 0 else 0)
                            others= [p for p in all_prices_svc if p != ref]
                            ps    = None; vt="—"; vc_col="#bbb"
                            if ref > 0 and others:
                                ps,_,_,_,_ = price_score(ref, others)
                                vt,_,vc_col = get_verdict(ps)
                            sc = score_color(ps)
                            link_cell = ("<a href='{}' target='_blank' style='color:#D04A02;"
                                         "font-weight:600;text-decoration:none'>📂 Open</a>".format(url)
                                         if url else "—")
                            tbl_rows.append("<tr style='background:{}'>"
                                            "<td>{}</td><td style='color:#555'>{}</td>"
                                            "<td style='font-family:monospace;font-size:0.79em;"
                                            "word-break:break-all'>{}</td>".format(
                                                bg, vendor_pill(row["Vendor"],vc2),
                                                row.get("Category",""), fname))
                            if has_price:
                                tbl_rows.append("<td style='font-family:monospace;font-weight:700;"
                                                "color:#22992E'>{}</td>".format(
                                                    _fmt(qp_str) if qp_num>0 else "—"))
                            tbl_rows.append("<td style='font-family:monospace;color:#295477;"
                                            "font-weight:700'>{}</td>"
                                            "<td style='text-align:center'>"
                                            "<span style='font-weight:800;color:{}'>{}</span></td>"
                                            "<td><span style='color:{};font-weight:700;"
                                            "font-size:0.82em'>{}</span></td>"
                                            "<td>{}</td></tr>".format(
                                                ep_fmt, sc,
                                                "{}/100".format(ps) if ps is not None else "—",
                                                vc_col, vt, link_cell))
                        tbl_rows.append("</tbody></table>")
                        st.markdown("".join(tbl_rows), unsafe_allow_html=True)

                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("🔍 Extract Prices — {}".format(svc[:40]),
                                      key="ep_{}".format(svc[:35]), type="primary"):
                            prog = st.progress(0); n = len(d_svc)
                            for ki,(_,row2) in enumerate(d_svc.iterrows()):
                                fname2 = str(row2.get("File Name","")).strip()
                                ck2    = "px_{}".format(fname2)
                                if st.session_state.get(ck2) is None:
                                    local = os.path.join(DEMO_DIR, fname2)
                                    if os.path.exists(local):
                                        st.session_state[ck2] = extract_price_from_file(local)
                                    else:
                                        url2 = resolve_url(row2)
                                        if url2 and url2.startswith("http") and REQUESTS_OK:
                                            try:
                                                resp2 = requests.get(url2, timeout=20)
                                                ext2  = url2.split("?")[0].rsplit(".",1)[-1].lower()
                                                st.session_state[ck2] = extract_price_from_bytes(resp2.content, ext2)
                                            except Exception: pass
                                prog.progress((ki+1)/n)
                            prog.empty(); st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab2:
    if NO_DATA:
        st.info("No catalog loaded. Go to 🗂 Upload Catalog tab.")
    else:
        pwc_header("New Quotation Analysis","Upload → extract price → score vs history")
        section_title("STEP 1 — UPLOAD QUOTE FILE")
        uploaded = st.file_uploader("Upload", type=["pdf","xlsx","xls","docx"],
                                     label_visibility="collapsed")
        if uploaded is not None:
            content  = uploaded.read()
            ext      = uploaded.name.rsplit(".",1)[-1]
            fname_up = uploaded.name
            st.success("Uploaded: **{}** ({} KB)".format(fname_up, round(len(content)/1024,1)))

            section_title("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting price…"):
                result    = extract_price_from_bytes(content, ext)
                new_price = result["price_num"]

            if new_price > 0:
                st.markdown("<div class='score-card green'>"
                            "<div style='font-size:0.72em;font-weight:700;text-transform:uppercase;"
                            "color:#22992E'>Extracted Price</div>"
                            "<div style='font-size:2.0em;font-weight:800;color:#22992E;"
                            "font-family:Georgia,serif'>{}</div>"
                            "</div>".format(_fmt(new_price)), unsafe_allow_html=True)
            else:
                st.warning("Price not found automatically.")
                manual    = st.number_input("Enter price manually (USD)",
                                             min_value=0.0, step=100.0, value=0.0,
                                             key="manual_price_input")
                if manual > 0: new_price = manual

            section_title("STEP 3 — SELECT SERVICES")
            all_svcs_up = sorted([s for s in df_exploded["Service"].unique()
                                   if str(s).strip() not in ["","nan"]])
            svc_search_up = st.text_input("Filter", placeholder="Search services…",
                                           key="svc_up", label_visibility="collapsed")
            if svc_search_up:
                all_svcs_up = [s for s in all_svcs_up if svc_search_up.lower() in s.lower()]
            new_services = st.multiselect("Services in this quotation",
                                           options=all_svcs_up, key="new_svcs",
                                           label_visibility="collapsed")
            cat_filter_up = st.selectbox(
                "Filter historical by category",
                options=["All"]+sorted([c for c in df_master["Category"].unique()
                                        if str(c).strip() not in ["","nan"]]),
                key="cat_up")

            section_title("STEP 4 — COMPARISON & VERDICT")
            manual_val = st.session_state.get("manual_price_input",0.0)
            if new_price <= 0 and manual_val > 0: new_price = float(manual_val)

            if new_services or new_price > 0:
                candidates = (df_exploded[df_exploded["Service"].isin(new_services)].copy()
                              if new_services else df_exploded.copy())
                if cat_filter_up != "All":
                    candidates = candidates[candidates["Category"]==cat_filter_up]
                cand_files = (candidates.drop_duplicates(subset=["File Name","Vendor"])
                              [["File Name","Vendor","Category","Hyperlink","Quoted Price"]].copy())
                if cand_files.empty:
                    st.warning("No historical quotes found.")
                else:
                    hist_prices = []; vendor_p_map = {}
                    for _,r in cand_files.iterrows():
                        qp = _parse_num(str(r.get("Quoted Price","")).strip())
                        if qp > 0: hist_prices.append(qp); vendor_p_map[r["Vendor"]] = qp
                        ck = "px_{}".format(str(r.get("File Name","")).strip())
                        ca = st.session_state.get(ck)
                        if ca and ca.get("price_num",0)>0:
                            hist_prices.append(ca["price_num"]); vendor_p_map[r["Vendor"]]=ca["price_num"]

                    if new_price > 0 and hist_prices:
                        ps,ps_lbl,avg_h,mn_h,mx_h = price_score(new_price, hist_prices)
                        vt,vt_desc,vt_col = get_verdict(ps)
                        vt_css = score_css(ps)
                        st.markdown("<div class='verdict-{}'>"
                                    "<div style='font-size:1.2em;margin-bottom:6px;"
                                    "font-family:Georgia,serif'>{}</div>"
                                    "<div style='font-size:0.88em;font-weight:400'>{}</div>"
                                    "</div>".format(vt_css,vt,vt_desc), unsafe_allow_html=True)
                        st.markdown("<br>", unsafe_allow_html=True)

                        sc1,sc2,sc3,sc4 = st.columns(4)
                        sc1.markdown("<div class='score-card {}'>"
                                     "<div style='font-size:0.68em;font-weight:700;"
                                     "text-transform:uppercase;color:{}'>Price Score</div>"
                                     "<div style='font-size:2.0em;font-weight:800;color:{};"
                                     "font-family:Georgia,serif'>{}/100</div>"
                                     "<div style='font-size:0.75em;color:#555;margin-top:4px'>"
                                     "vs {} historical quotes</div></div>".format(
                                         vt_css,vt_col,vt_col,
                                         ps if ps is not None else "N/A",
                                         len(hist_prices)), unsafe_allow_html=True)
                        sc2.markdown("<div class='score-card navy'>"
                                     "<div style='font-size:0.68em;font-weight:700;"
                                     "text-transform:uppercase;color:#003078'>Your Price</div>"
                                     "<div style='font-size:1.8em;font-weight:800;color:#D04A02;"
                                     "font-family:Georgia,serif'>{}</div>"
                                     "</div>".format(_fmt(new_price)), unsafe_allow_html=True)
                        sc3.markdown("<div class='score-card navy'>"
                                     "<div style='font-size:0.68em;font-weight:700;"
                                     "text-transform:uppercase;color:#003078'>Market Average</div>"
                                     "<div style='font-size:1.8em;font-weight:800;color:#003078;"
                                     "font-family:Georgia,serif'>{}</div>"
                                     "<div style='font-size:0.72em;color:#555;margin-top:4px'>"
                                     "min {} · max {}</div>"
                                     "</div>".format(_fmt(avg_h),_fmt(mn_h),_fmt(mx_h)),
                                     unsafe_allow_html=True)
                        sc4.markdown("<div class='score-card {}'>"
                                     "<div style='font-size:0.68em;font-weight:700;"
                                     "text-transform:uppercase;color:{}'>vs Average</div>"
                                     "<div style='font-size:1.0em;font-weight:800;color:{};"
                                     "margin-top:6px'>{}</div>"
                                     "</div>".format(vt_css,vt_col,vt_col,ps_lbl),
                                     unsafe_allow_html=True)

                        # Positioning chart — PwC colours
                        st.markdown("<br>", unsafe_allow_html=True)
                        section_title("PRICE POSITIONING")
                        chart_data = []
                        for _,r in cand_files.iterrows():
                            fn  = str(r.get("File Name","")).strip()
                            qp  = _parse_num(str(r.get("Quoted Price","")).strip())
                            ck  = "px_{}".format(fn)
                            ca  = st.session_state.get(ck)
                            ep  = ca["price_num"] if ca else 0.0
                            pval= ep if ep>0 else qp
                            if pval>0:
                                chart_data.append({"Label":"{} / {}".format(r["Vendor"],fn[:10]),
                                                   "Price":pval,"Type":"Historical"})
                        chart_data.append({"Label":"★ YOUR QUOTE","Price":new_price,"Type":"New"})
                        cdf = pd.DataFrame(chart_data).sort_values("Price")
                        colors = [PwC_ORANGE if t=="New" else PwC_NAVY for t in cdf["Type"]]
                        cf = go.Figure(go.Bar(
                            x=cdf["Label"], y=cdf["Price"],
                            marker_color=colors, marker_line_width=0,
                            text=cdf["Price"].apply(_fmt), textposition="outside"))
                        cf.add_hline(y=avg_h, line_dash="dash", line_color=PwC_TEAL, line_width=2,
                                      annotation_text="Market Avg: {}".format(_fmt(avg_h)),
                                      annotation_position="top right")
                        cf.update_layout(
                            height=360, plot_bgcolor=CBG, paper_bgcolor=CBG,
                            margin=dict(l=5,r=10,t=20,b=10), font=CFONT,
                            yaxis=dict(title="Price",showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                            xaxis=dict(tickangle=-25), bargap=0.3, showlegend=False)
                        st.plotly_chart(cf, use_container_width=True)


# ════════════════════════════════════════════════════════════
# TAB 3 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab3:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        dm = df_master.copy()
        if selected_cat    != "All": dm = dm[dm["Category"]==selected_cat]
        if selected_vendor != "All": dm = dm[dm["Vendor"]==selected_vendor]
        st.dataframe(dm.drop(columns=["Services List","Hyperlink"], errors="ignore"),
                     use_container_width=True, height=500)


# ════════════════════════════════════════════════════════════
# TAB 4 — UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab4:
    pwc_header("Upload Master Catalog","Upload any Excel or CSV catalog — AI auto-detects columns.")
    if DATA_SOURCE=="uploaded":
        st.success("✅ Using uploaded catalog: **{}** rows · **{}** vendors · **{}** services".format(
            len(df_master), df_master["Vendor"].nunique(), df_exploded["Service"].nunique()))
    elif not NO_DATA:
        st.info("📂 Currently using: **Master Catalog.xlsx** — {} vendors · {} services".format(
            df_master["Vendor"].nunique(), df_exploded["Service"].nunique()))

    section_title("UPLOAD A DIFFERENT CATALOG")
    catalog_file = st.file_uploader("Upload", type=["xlsx","xls","csv"],
                                     label_visibility="collapsed", key="catalog_upload")
    if catalog_file is not None:
        file_bytes = catalog_file.read(); fname_cat = catalog_file.name
        with st.spinner("Analyzing catalog…"):
            df_new, df_exp_new, err = process_uploaded_catalog(file_bytes, fname_cat)
        if err: st.error("❌ {}".format(err))
        elif df_new is None: st.error("❌ Could not process file.")
        else:
            st.success("✅ **{}** rows · **{}** vendors · **{}** categories".format(
                len(df_new), df_new["Vendor"].nunique(), df_new["Category"].nunique()))
            pc1,pc2 = st.columns(2)
            spv_new = (df_exp_new.groupby("Vendor")["Service"].nunique()
                       .sort_values(ascending=False).reset_index())
            spv_new.columns = ["Vendor","Services"]
            with pc1:
                pf1 = go.Figure(go.Bar(
                    x=spv_new["Vendor"], y=spv_new["Services"],
                    marker_color=PwC_ORANGE, marker_line_width=0,
                    text=spv_new["Services"], textposition="outside"))
                pf1.update_layout(title="Services per Vendor",height=300,
                                   plot_bgcolor=CBG,paper_bgcolor=CBG,
                                   margin=dict(l=5,r=10,t=40,b=10),font=CFONT,
                                   yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                                   xaxis=dict(tickangle=-30),bargap=0.35)
                st.plotly_chart(pf1, use_container_width=True)
            cat_new = (df_new.drop_duplicates(subset=["Category","File Name"])
                       .groupby("Category").size().reset_index())
            cat_new.columns = ["Category","Count"]
            with pc2:
                if not cat_new.empty:
                    pf2 = px.pie(cat_new,names="Category",values="Count",
                                  hole=0.45,color_discrete_sequence=CHART_SEQ)
                    pf2.update_traces(textposition="outside",textinfo="label+percent",textfont_size=10)
                    pf2.update_layout(title="Category Distribution",height=300,
                                       margin=dict(l=10,r=10,t=40,b=10),
                                       paper_bgcolor=CBG,font=CFONT)
                    st.plotly_chart(pf2, use_container_width=True)
            st.dataframe(df_new.drop(columns=["Services List","Hyperlink"],
                                      errors="ignore").head(20),
                         use_container_width=True, height=280)
            if st.button("✅ Apply This Catalog to Dashboard", type="primary"):
                st.session_state["uploaded_catalog_df"]  = df_new
                st.session_state["uploaded_catalog_exp"] = df_exp_new
                st.success("✅ Applied!"); st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 5 — VENDOR ANALYSIS
# ════════════════════════════════════════════════════════════
with tab5:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header("Vendor Price Analysis","Prices from GitHub quote files · Per-service benchmarking")

        GITHUB_RAW = "https://raw.githubusercontent.com/avijeet528/vendor-draft-2/main/demo_quotes/{}"

        @st.cache_data(show_spinner=False)
        def load_price_from_github(filename):
            url = GITHUB_RAW.format(filename)
            try:
                resp = requests.get(url, timeout=30)
                if resp.status_code != 200:
                    return {"price":"","price_num":0.0,"status":"Not found"}
                ext    = filename.rsplit(".",1)[-1].lower()
                result = extract_price_from_bytes(resp.content, ext)
                result["status"] = "✅ Extracted" if result["price_num"]>0 else "⚠️ No price"
                return result
            except Exception as e:
                return {"price":"","price_num":0.0,"status":"❌ {}".format(str(e)[:40])}

        @st.cache_data(show_spinner=False)
        def load_all_prices_from_github(filenames_tuple):
            return {fname: load_price_from_github(fname) for fname in filenames_tuple}

        all_fnames = tuple(str(f).strip() for f in df_master["File Name"].unique()
                           if str(f).strip() not in ["","nan"])

        col_load, col_info = st.columns([2,3])
        with col_load:
            run_analysis = st.button("🔄 Load All Prices from GitHub",
                                      type="primary", use_container_width=True, key="run_gh_analysis")
        with col_info:
            st.markdown("<div style='background:#EEF2FF;border:1px solid #003078;border-radius:4px;"
                        "padding:8px 14px;font-size:0.83em'>📁 Will read <b>{}</b> files from GitHub. "
                        "Catalog prices used as fallback.</div>".format(len(all_fnames)),
                        unsafe_allow_html=True)

        if run_analysis:
            st.session_state["gh_prices_loaded"] = True
            load_all_prices_from_github.clear(); load_price_from_github.clear()

        if st.session_state.get("gh_prices_loaded", False):
            with st.spinner("Loading prices…"):
                gh_prices = load_all_prices_from_github(all_fnames)

            rows_data = []
            for _, r in df_master.iterrows():
                fname   = str(r.get("File Name","")).strip()
                qp      = _parse_num(str(r.get("Quoted Price","")).strip())
                gh_info = gh_prices.get(fname,{})
                gh_p    = gh_info.get("price_num",0.0)
                best_p  = gh_p if gh_p > 0 else qp
                source  = "extracted" if gh_p>0 else "catalog" if qp>0 else "none"
                rows_data.append({"Vendor":r["Vendor"],"Category":r["Category"],"File Name":fname,
                                   "Quoted Price":qp,"Extracted Price":gh_p,
                                   "Best Price":best_p,"Source":source,
                                   "Status":gh_info.get("status","—"),
                                   "Services":r.get("Comments","")})
            df_analysis = pd.DataFrame(rows_data)
            df_analysis = df_analysis[df_analysis["Best Price"]>0]

            if df_analysis.empty:
                st.warning("No prices found.")
            else:
                section_title("SUMMARY")
                k1,k2,k3,k4 = st.columns(4)
                kpi(k1, len(df_analysis),    "Files with Prices",    PwC_ORANGE)
                kpi(k2, len(df_analysis[df_analysis["Source"]=="extracted"]), "Extracted", PwC_NAVY)
                kpi(k3, len(df_analysis[df_analysis["Source"]=="catalog"]),   "From Catalog", PwC_TEAL)
                kpi(k4, _fmt(df_analysis["Best Price"].mean()), "Avg Quote",  PwC_DARK)
                st.markdown("<br>", unsafe_allow_html=True)

                # Vendor totals
                vendor_totals = (df_analysis.groupby("Vendor")["Best Price"]
                                 .agg(["sum","mean","count"]).reset_index())
                vendor_totals.columns = ["Vendor","Total","Average","Quotes"]
                vendor_totals = vendor_totals.sort_values("Average")
                overall_avg   = df_analysis["Best Price"].mean()

                section_title("VENDOR COMPARISON — AVERAGE QUOTE VALUE")
                fig_v = go.Figure()
                for i,(_, vr) in enumerate(vendor_totals.iterrows()):
                    color = PwC_GREEN if i==0 else PwC_RED if i==len(vendor_totals)-1 else PwC_ORANGE
                    fig_v.add_trace(go.Bar(
                        x=[vr["Vendor"]], y=[vr["Average"]],
                        marker_color=color, marker_line_width=0,
                        name=vr["Vendor"],
                        text=[_fmt(vr["Average"])], textposition="outside"))
                fig_v.add_hline(y=overall_avg, line_dash="dash", line_color=PwC_NAVY, line_width=2,
                                 annotation_text="Avg: {}".format(_fmt(overall_avg)),
                                 annotation_position="top right")
                fig_v.update_layout(height=360, plot_bgcolor=CBG, paper_bgcolor=CBG,
                                     margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
                                     yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                                     xaxis=dict(tickangle=-20), bargap=0.35, showlegend=False)
                st.plotly_chart(fig_v, use_container_width=True)

                # Per-service
                section_title("PER-SERVICE BENCHMARKING")
                df_svc_rows = []
                for _, r in df_analysis.iterrows():
                    svcs_raw = str(r["Services"]).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
                    svcs = [s.strip() for s in svcs_raw.split("\n")
                            if s.strip() and s.strip() not in ["nan","None",""]]
                    if not svcs: svcs = [svcs_raw.strip()]
                    for svc in svcs:
                        df_svc_rows.append({"Service":svc,"Vendor":r["Vendor"],
                                             "Price":r["Best Price"],"Source":r["Source"]})
                df_svc_df = pd.DataFrame(df_svc_rows)
                svc_vc    = df_svc_df.groupby("Service")["Vendor"].nunique()
                multi_svcs= svc_vc[svc_vc>1].index.tolist()

                if multi_svcs:
                    st.markdown("<div style='background:#F0FFF4;border-left:4px solid #22992E;"
                                "padding:8px 14px;border-radius:2px;margin-bottom:12px;font-size:0.87em'>"
                                "✅ <b>{}</b> services with multiple vendor quotes available."
                                "</div>".format(len(multi_svcs)), unsafe_allow_html=True)

                    for svc in sorted(multi_svcs):
                        d_s   = df_svc_df[df_svc_df["Service"]==svc].sort_values("Price")
                        min_p = d_s["Price"].min(); max_p = d_s["Price"].max()
                        avg_p = d_s["Price"].mean()
                        best_v= d_s.loc[d_s["Price"].idxmin(),"Vendor"]
                        spread= round((max_p-min_p)/min_p*100,1) if min_p>0 else 0

                        static_section_header(
                            "{}  ·  {} vendors  ·  spread {}%  ·  cheapest: {} @ {}".format(
                                svc, d_s["Vendor"].nunique(), spread, best_v, _fmt(min_p)),
                            PwC_ORANGE)
                        with st.container():
                            sc1,sc2,sc3 = st.columns(3)
                            sc1.markdown("<div class='score-card green'>"
                                         "<div style='font-size:0.68em;font-weight:700;"
                                         "text-transform:uppercase;color:#22992E'>Cheapest</div>"
                                         "<div style='font-size:1.5em;font-weight:800;color:#22992E;"
                                         "font-family:Georgia,serif'>{}</div>"
                                         "<div style='font-size:0.78em;color:#555'>{}</div>"
                                         "</div>".format(_fmt(min_p),best_v), unsafe_allow_html=True)
                            sc2.markdown("<div class='score-card navy'>"
                                         "<div style='font-size:0.68em;font-weight:700;"
                                         "text-transform:uppercase;color:#003078'>Average</div>"
                                         "<div style='font-size:1.5em;font-weight:800;color:#003078;"
                                         "font-family:Georgia,serif'>{}</div>"
                                         "<div style='font-size:0.78em;color:#555'>"
                                         "{} vendors</div></div>".format(
                                             _fmt(avg_p), d_s["Vendor"].nunique()), unsafe_allow_html=True)
                            sc3.markdown("<div class='score-card red'>"
                                         "<div style='font-size:0.68em;font-weight:700;"
                                         "text-transform:uppercase;color:#E0301E'>Most Expensive</div>"
                                         "<div style='font-size:1.5em;font-weight:800;color:#E0301E;"
                                         "font-family:Georgia,serif'>{}</div>"
                                         "<div style='font-size:0.78em;color:#555'>"
                                         "{}</div></div>".format(
                                             _fmt(max_p),
                                             d_s.loc[d_s["Price"].idxmax(),"Vendor"]),
                                         unsafe_allow_html=True)
                            st.markdown("<br>", unsafe_allow_html=True)

                            all_p = d_s["Price"].tolist()
                            bar_c = []
                            for idx,(_,vr) in enumerate(d_s.iterrows()):
                                if vr["Vendor"]==best_v: bar_c.append(PwC_GREEN)
                                elif vr["Vendor"]==d_s.loc[d_s["Price"].idxmax(),"Vendor"]: bar_c.append(PwC_RED)
                                else: bar_c.append(PwC_ORANGE)
                            fig_s = go.Figure(go.Bar(
                                x=d_s["Vendor"], y=d_s["Price"],
                                marker_color=bar_c, marker_line_width=0,
                                text=d_s["Price"].apply(_fmt), textposition="outside"))
                            fig_s.add_hline(y=avg_p, line_dash="dash", line_color=PwC_NAVY, line_width=2,
                                             annotation_text="Avg: {}".format(_fmt(avg_p)),
                                             annotation_position="top right")
                            fig_s.update_layout(
                                height=260, plot_bgcolor=CBG, paper_bgcolor=CBG,
                                margin=dict(l=5,r=10,t=10,b=10), font=CFONT,
                                yaxis=dict(showgrid=True,gridcolor="#E0E0E0",zeroline=False),
                                bargap=0.4, showlegend=False)
                            st.plotly_chart(fig_s, use_container_width=True)

                # Download
                st.markdown("<br>", unsafe_allow_html=True)
                section_title("DOWNLOAD")
                st.download_button("📥 Download Analysis CSV",
                                    data=df_analysis[["Vendor","Category","File Name",
                                                       "Best Price","Source"]].to_csv(index=False),
                                    file_name="vendor_analysis.csv", mime="text/csv", type="primary")
        else:
            st.markdown("<div style='background:white;border:1px solid #e0e0e0;border-radius:4px;"
                        "padding:24px;text-align:center;margin-top:20px'>"
                        "<div style='font-size:2em'>🔍</div>"
                        "<div style='font-size:1.1em;font-weight:700;color:#2D2D2D;margin:8px 0;"
                        "font-family:Georgia,serif'>Click the button above to start</div>"
                        "<div style='font-size:0.85em;color:#7D7D7D'>"
                        "Fetches <b>{}</b> files from GitHub and extracts prices."
                        "</div></div>".format(len(all_fnames)), unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# TAB 6 — REAL ANALYSIS
# ════════════════════════════════════════════════════════════
with tab6:
    pwc_header("Real Quotation Analysis",
               "Upload Master Catalog → follow File Links → extract prices → full analysis")

    section_title("STEP 1 — UPLOAD YOUR MASTER CATALOG")
    uploaded_real = st.file_uploader("Upload", type=["xlsx","xls","csv"],
                                      key="tab6_real_upload", label_visibility="collapsed")

    if uploaded_real is None:
        if not NO_DATA and not df_master.empty:
            use_real = st.checkbox("Use currently loaded catalog ({} files)".format(len(df_master)),
                                    key="use_current_for_real")
            if use_real:
                df_real_cat = df_master.copy()
                st.success("Using current catalog — **{}** vendors · **{}** files".format(
                    df_real_cat["Vendor"].nunique(), len(df_real_cat)))
            else:
                st.info("👆 Upload your Master Catalog to start."); st.stop()
        else:
            st.info("👆 Upload your Master Catalog to start."); st.stop()
    else:
        real_bytes = uploaded_real.read()
        df_real_cat,_,real_err = process_uploaded_catalog(real_bytes, uploaded_real.name)
        if real_err or df_real_cat is None:
            st.error("❌ {}".format(real_err)); st.stop()
        st.success("✅ **{}** vendors · **{}** files · **{}** categories".format(
            df_real_cat["Vendor"].nunique(), len(df_real_cat), df_real_cat["Category"].nunique()))

    st.markdown("<br>", unsafe_allow_html=True)
    section_title("STEP 2 — RUN ANALYSIS")
    run_real = st.button("🚀 Run Full Analysis", type="primary", key="run_real_analysis")

    if run_real:
        st.session_state["real_analysis_done"] = False
        st.session_state["real_analysis_df"]   = None

    if run_real or st.session_state.get("real_analysis_done", False):
        if run_real:
            with st.spinner("Analyzing {} files…".format(len(df_real_cat))):
                df_result = analyze_real_catalog(df_real_cat)
            st.session_state["real_analysis_done"] = True
            st.session_state["real_analysis_df"]   = df_result.to_dict("records")
        else:
            df_result = pd.DataFrame(st.session_state["real_analysis_df"])

        df_priced = df_result[df_result["Best Price"]>0].copy()

        k1,k2,k3,k4,k5 = st.columns(5)
        kpi(k1, len(df_result),                    "Total Files",  PwC_ORANGE)
        kpi(k2, len(df_priced),                    "Prices Found", PwC_GREEN)
        kpi(k3, len(df_result)-len(df_priced),     "No Price",     PwC_GREY)
        kpi(k4, df_result["Vendor"].nunique(),      "Vendors",      PwC_NAVY)
        kpi(k5, _fmt(df_priced["Best Price"].mean()) if not df_priced.empty else "—",
            "Avg Quote", PwC_TEAL)

        if df_priced.empty:
            st.warning("No prices found. Check File Link URLs are accessible."); st.stop()

        # Vendor ranking
        st.markdown("<br>", unsafe_allow_html=True)
        section_title("VENDOR RANKING — CHEAPEST TO MOST EXPENSIVE")
        v_sum = (df_priced.groupby("Vendor")["Best Price"]
                 .agg(["mean","sum","min","max","count"]).reset_index())
        v_sum.columns = ["Vendor","Average","Total","Min","Max","Quotes"]
        v_sum = v_sum.sort_values("Average")
        ov_avg = df_priced["Best Price"].mean()

        bar_c = []
        for i in range(len(v_sum)):
            if i==0: bar_c.append(PwC_GREEN)
            elif i==len(v_sum)-1: bar_c.append(PwC_RED)
            else: bar_c.append(PwC_ORANGE)

        fig_rank = go.Figure(go.Bar(
            x=v_sum["Vendor"], y=v_sum["Average"],
            marker_color=bar_c, marker_line_width=0,
            text=v_sum["Average"].apply(_fmt), textposition="outside"))
        fig_rank.add_hline(y=ov_avg, line_dash="dash", line_color=PwC_NAVY, line_width=2,
                            annotation_text="Avg: {}".format(_fmt(ov_avg)),
                            annotation_position="top right")
        fig_rank.update_layout(
            height=360, plot_bgcolor=CBG, paper_bgcolor=CBG,
            margin=dict(l=5,r=10,t=30,b=10), font=CFONT,
            yaxis=dict(title="Average Quote (USD)",showgrid=True,gridcolor="#E0E0E0",zeroline=False),
            xaxis=dict(tickangle=-20), bargap=0.35, showlegend=False)
        st.plotly_chart(fig_rank, use_container_width=True)

        # Download
        st.download_button("📥 Download Analysis CSV",
                            data=df_priced[["Vendor","Category","File Name",
                                            "Quoted Price","Extracted Price",
                                            "Best Price","Source","Status"]].to_csv(index=False),
                            file_name="real_analysis.csv", mime="text/csv", type="primary")
