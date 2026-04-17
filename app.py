# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  PwC Brand | Source Sans Pro + Georgia | Streamlit
#  Colour palette: Orange · Black · White · Grey ONLY
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
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# PwC COLOUR PALETTE — Orange · Black · White · Grey ONLY
# ════════════════════════════════════════════════════════════
C_ORANGE      = "#D04A02"   # PwC primary orange
C_ORANGE_DARK = "#A33A00"   # deep orange
C_ORANGE_MID  = "#E8703A"   # mid orange
C_ORANGE_LITE = "#FAD4C0"   # light orange tint
C_BLACK       = "#1A1A1A"   # near black
C_DARK        = "#2D2D2D"   # dark charcoal
C_MID         = "#4A4A4A"   # mid grey-black
C_GREY_DARK   = "#7D7D7D"   # dark grey
C_GREY        = "#B0B0B0"   # mid grey
C_GREY_LITE   = "#D8D8D8"   # light grey
C_GREY_BG     = "#F3F3F3"   # background grey
C_WHITE       = "#FFFFFF"   # white

# Chart colour sequence — shades of orange + grey
CHART_SEQ = [C_ORANGE, C_DARK, C_ORANGE_MID, C_GREY_DARK,
             C_ORANGE_DARK, C_GREY, C_ORANGE_LITE, C_MID]

def chart_color(i): return CHART_SEQ[i % len(CHART_SEQ)]

CFONT = dict(
    family="Georgia,'ITC Charter','Source Sans Pro',Arial",
    size=11, color=C_DARK)
CBG = C_GREY_BG
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');

html,body,[class*="css"],div,p,span,td,th,label,button,.stMarkdown{
    font-family:'Source Sans Pro','Helvetica Neue',Arial,sans-serif !important;}
h1,h2,h3,.pwc-heading{
    font-family:Georgia,'ITC Charter',serif !important; font-weight:700 !important;}

.main .block-container{
    background:#F3F3F3 !important; padding-top:1.2rem;
    max-width:100% !important;
    padding-left:1.8rem !important; padding-right:1.8rem !important;}
#MainMenu{visibility:hidden;} footer{visibility:hidden;} header{visibility:hidden;}
[data-testid="collapsedControl"]{display:none !important;}

/* ── Sidebar — minimal, dark ── */
section[data-testid="stSidebar"]{
    background:#1A1A1A !important;
    border-right:3px solid #D04A02;
    min-width:260px !important; max-width:260px !important;}
section[data-testid="stSidebar"] *{color:#F0F0F0 !important;}
section[data-testid="stSidebar"] .stButton button{
    background:#D04A02 !important; color:white !important;
    border:none !important; border-radius:2px !important;}

/* ── Top filter bar ── */
.filter-bar{background:#2D2D2D;padding:12px 20px;border-radius:4px;
    margin-bottom:18px;border-left:4px solid #D04A02;}
.filter-bar label, .filter-bar p, .filter-bar span{color:#F0F0F0 !important;}

/* ── Tabs ── */
button[data-baseweb="tab"]{
    font-weight:600 !important; font-size:0.92em !important;
    color:#7D7D7D !important; border-radius:0 !important;}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;
    border-bottom:3px solid #D04A02 !important;
    background:transparent !important;}

/* ── KPI boxes ── */
.kpi-box{border-radius:4px;padding:20px 12px;text-align:center;
    color:white;border-left:5px solid rgba(255,255,255,0.2);}
.kpi-value{font-size:2.2em;font-weight:700;margin:0;line-height:1.1;
    font-family:Georgia,serif !important;}
.kpi-label{font-size:0.75em;font-weight:700;opacity:0.9;margin-top:6px;
    letter-spacing:1px;text-transform:uppercase;}

/* ── Section heading ── */
.sec-head{font-size:0.75em;font-weight:700;letter-spacing:1.2px;
    text-transform:uppercase;color:#D04A02;
    margin:20px 0 10px;border-bottom:2px solid #D04A02;
    padding-bottom:4px;display:block;}

/* ── Tables ── */
.comp-table{width:100%;border-collapse:collapse;
    font-size:0.82em;border:1px solid #E0E0E0;}
.comp-table thead tr{background:#2D2D2D;}
.comp-table thead th{padding:10px 12px;text-align:left;font-weight:700;
    font-size:0.79em;letter-spacing:0.5px;text-transform:uppercase;
    color:white !important;border:none;}
.comp-table tbody tr:nth-child(even){background:#F8F8F8;}
.comp-table tbody tr:hover{background:#FAD4C0;}
.comp-table tbody td{padding:9px 12px;border-bottom:1px solid #EBEBEB;
    vertical-align:middle;word-break:break-word;color:#2D2D2D;}

/* ── Vendor badge ── */
.vbadge{display:inline-block;padding:3px 9px;border-radius:2px;
    color:white;font-size:0.77em;font-weight:700;white-space:nowrap;}

/* ── Score / verdict cards ── */
.scard{border-radius:4px;padding:16px;margin-bottom:10px;
    border-left:5px solid #D04A02;background:white;}
.scard-orange{border-color:#D04A02;background:#FFF5F0;}
.scard-dark{border-color:#2D2D2D;background:#F5F5F5;}
.scard-grey{border-color:#7D7D7D;background:#FAFAFA;}
.scard-lite{border-color:#E8703A;background:#FFF8F5;}

.verdict-good{background:#FFF5F0;border:2px solid #D04A02;
    border-radius:4px;padding:14px 18px;color:#D04A02;font-weight:700;}
.verdict-mid{background:#F5F5F5;border:2px solid #4A4A4A;
    border-radius:4px;padding:14px 18px;color:#4A4A4A;font-weight:700;}
.verdict-bad{background:#F0F0F0;border:2px solid #7D7D7D;
    border-radius:4px;padding:14px 18px;color:#2D2D2D;font-weight:700;}

/* ── Chat ── */
.chat-outer{background:#F8F8F8;border:1px solid #E0E0E0;
    border-radius:0 0 6px 6px;padding:14px;
    max-height:480px;overflow-y:auto;margin-bottom:0;}
.chat-header{background:#2D2D2D;color:white;padding:12px 18px;
    border-radius:6px 6px 0 0;margin-bottom:0;}
.msg-user{background:#D04A02;color:white;border-radius:14px 14px 3px 14px;
    padding:10px 15px;margin:6px 0 6px auto;max-width:72%;
    font-size:0.87em;display:inline-block;float:right;clear:both;}
.msg-bot{background:white;color:#2D2D2D;
    border:1px solid #E0E0E0;border-left:4px solid #D04A02;
    border-radius:14px 14px 14px 3px;padding:10px 15px;
    margin:6px 0;max-width:82%;font-size:0.87em;
    display:inline-block;float:left;clear:both;}
.chat-wrap{overflow:hidden;margin-bottom:4px;}

/* ── Price tags ── */
.ptag-good{background:#D04A02;color:white;padding:3px 9px;
    border-radius:3px;font-weight:700;font-size:0.83em;}
.ptag-mid{background:#4A4A4A;color:white;padding:3px 9px;
    border-radius:3px;font-weight:700;font-size:0.83em;}
.ptag-bad{background:#7D7D7D;color:white;padding:3px 9px;
    border-radius:3px;font-weight:700;font-size:0.83em;}

/* ── Analysis panel ── */
.analysis-panel{background:white;border:1px solid #E0E0E0;
    border-radius:6px;padding:20px 24px;margin-bottom:16px;
    border-top:4px solid #D04A02;}

/* ── Insight box ── */
.insight-box{background:#FFF5F0;border-left:4px solid #D04A02;
    border-radius:0 4px 4px 0;padding:12px 16px;
    margin:10px 0;font-size:0.87em;color:#2D2D2D;}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"(?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?"
    r"|(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?"
    r"|\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?", re.IGNORECASE)
TOTAL_KW = ["grand total","total amount","total price","amount due",
            "net total","total cost","total value","subtotal","total"]

def _parse_num(s):
    try: return float(re.sub(r"[^\d.]","",str(s)) or "0")
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
        valid = [h.strip() for h in hits if _parse_num(h) >= 50]
        if valid: return max(valid, key=_parse_num)
    all_h = PRICE_RE.findall(text)
    valid = [h.strip() for h in all_h if _parse_num(h) >= 100]
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
            wb = openpyxl.load_workbook(io.BytesIO(content),
                                         data_only=True, read_only=True)
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
    return {"price": price,
            "price_num": _parse_num(price) if price else 0.0,
            "text": text[:5000]}

def extract_price_from_file(filepath):
    try:
        with open(filepath,"rb") as f: content = f.read()
        ext = filepath.rsplit(".",1)[-1].lower()
        return extract_price_from_bytes(content, ext)
    except: return {"price":"","price_num":0.0,"text":""}

# ════════════════════════════════════════════════════════════
# SCORING
# ════════════════════════════════════════════════════════════
def price_score(new_price, hist_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return None,"No comparison data",0,0,0
    mn = min(valid); mx = max(valid)
    avg = sum(valid)/len(valid)
    if mx == mn: return 50,"Same as historical average",avg,mn,mx
    score = round((1-(new_price-mn)/(mx-mn))*100,1)
    score = max(0,min(100,score))
    pct   = round((new_price-avg)/avg*100,1)
    label = ("{}% BELOW average — COMPETITIVE".format(abs(pct)) if new_price < avg
             else "{}% ABOVE average — REVIEW NEEDED".format(abs(pct)))
    return score,label,avg,mn,mx

def score_color(s):
    if s is None: return C_GREY
    if s >= 70:   return C_ORANGE
    if s >= 40:   return C_GREY_DARK
    return C_MID

def score_css(s):
    if s is None: return "grey"
    if s >= 70:   return "orange"
    if s >= 40:   return "mid"
    return "dark"

def get_verdict(ps):
    if ps is None:
        return "⚪ No Data","No comparison data.",C_GREY
    if ps >= 70:
        return "✅ COMPETITIVE","Priced competitively. Proceed with confidence.",C_ORANGE
    if ps >= 40:
        return "🟡 AVERAGE","Within range. Negotiate for small discount.",C_GREY_DARK
    return "🔴 HIGH — NEGOTIATE","Above historical average. Recommend negotiating.",C_MID

# ════════════════════════════════════════════════════════════
# DATA HELPERS
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
    if not v or str(v).strip() in ["","nan","None"]:
        return ["(unspecified)"]
    s = str(v).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts = [p.strip() for p in s.split("\n")
             if p.strip() and p.strip() != "nan"]
    if not parts:
        parts = [p.strip() for p in s.split(";") if p.strip()]
    if not parts and len(s) < 300:
        parts = [p.strip() for p in s.split(",") if p.strip()]
    return parts if parts else ["(unspecified)"]

def _explode_services(df):
    df2 = df.copy()
    df2["Services List"] = df2["Comments"].apply(_parse_services)
    df_exp = df2.explode("Services List").copy()
    df_exp.rename(columns={"Services List":"Service"}, inplace=True)
    df_exp["Service"] = df_exp["Service"].apply(lambda x: str(x).strip())
    df_exp = df_exp[~df_exp["Service"].isin(
        ["","(unspecified)","nan","None"])].reset_index(drop=True)
    return df2, df_exp

def _normalise_columns(df):
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category" and "Category" not in col_map:
            col_map["Category"] = c
        elif any(k in cl for k in ["vendor","supplier"]) and "Vendor" not in col_map:
            col_map["Vendor"] = c
        elif ("file name" in cl or cl == "filename") and "File Name" not in col_map:
            col_map["File Name"] = c
        elif any(k in cl for k in ["file link","file url"]) and "File Link" not in col_map:
            col_map["File Link"] = c
        elif any(k in cl for k in ["comment","service","description","scope"]) and "Comments" not in col_map:
            col_map["Comments"] = c
        elif any(k in cl for k in ["price","cost","amount","quoted"]) and "Quoted Price" not in col_map:
            col_map["Quoted Price"] = c
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
                vals = [str(v).strip().lower()
                        for v in row.values if pd.notna(v)]
                if (any("category" in v for v in vals)
                        and any("vendor" in v for v in vals)):
                    header_row = i; break
            df = pd.read_excel(XLS_PATH, engine="openpyxl",
                                header=header_row)
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
        df["Hyperlink"] = df["File Link"].apply(
            lambda x: "" if x in ["","nan"] else x)
    df, df_exp = _explode_services(df)
    return df, df_exp

def process_uploaded_catalog(file_bytes, filename):
    try:
        ext = filename.rsplit(".",1)[-1].lower()
        if ext in ("xlsx","xls"):
            raw = pd.read_excel(io.BytesIO(file_bytes),
                                 engine="openpyxl", header=None)
        elif ext == "csv":
            raw = pd.read_csv(io.BytesIO(file_bytes), header=None)
        else:
            return None, None, "Unsupported file type."
        header_row = 0
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower()
                    for v in row.values if pd.notna(v)]
            joined = " ".join(vals)
            if (any(k in joined for k in ["vendor","supplier"])
                    and any(k in joined for k in ["file","document"])):
                header_row = i; break
        if ext in ("xlsx","xls"):
            df = pd.read_excel(io.BytesIO(file_bytes),
                                engine="openpyxl", header=header_row)
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
                        if (cell.value and
                                str(cell.value).strip().lower() == "file name"):
                            fn_col = cell.column
                            hdr_row = cell.row; break
                    if fn_col: break
                if fn_col and hdr_row:
                    for row in ws_h.iter_rows(min_row=hdr_row+1):
                        for cell in row:
                            if (cell.column == fn_col
                                    and cell.value and cell.hyperlink):
                                hmap[str(cell.value).strip()] = str(
                                    cell.hyperlink.target).strip()
                wb_h.close()
            except Exception: pass
        df["Hyperlink"] = df["File Name"].map(hmap).fillna("")
        df, df_exp = _explode_services(df)
        return df, df_exp, None
    except Exception as e:
        return None, None, str(e)

# ════════════════════════════════════════════════════════════
# SUBCATEGORY INFERENCE
# ════════════════════════════════════════════════════════════
def infer_subcategory(category, comments, file_name):
    txt = (str(comments) + " " + str(file_name)).lower()
    cat = str(category).lower().strip()
    if "cybersecurity" in cat:
        if any(k in txt for k in ["trendmicro","endpoint","antivirus"]):
            return "Endpoint Protection"
        if any(k in txt for k in ["cyberark","privileged","pam"]):
            return "Privileged Access Mgmt"
        if any(k in txt for k in ["knowbe4","awareness","phishing"]):
            return "Security Awareness"
        if any(k in txt for k in ["forescout","nac","network access"]):
            return "Network Access Control"
        if any(k in txt for k in ["siem","splunk","monitor"]):
            return "SIEM / Monitoring"
        return "General Security"
    if "network" in cat or "telecom" in cat:
        if "meraki" in txt: return "Cisco Meraki"
        if "palo alto" in txt: return "Palo Alto NGFW"
        if "equinix" in txt: return "Equinix"
        if "cisco" in txt: return "Cisco Networking"
        return "General Network"
    if "hosting" in cat:
        if any(k in txt for k in ["vmware","vcf"]): return "VMware"
        if "oracle" in txt: return "Oracle DB"
        if "netapp" in txt: return "NetApp Storage"
        if any(k in txt for k in ["colocation","colo"]): return "Colocation"
        return "General Hosting"
    if "m365" in cat: return "M365 Licensing"
    if "idam" in cat or "iam" in cat: return "Identity Migration"
    if "snow" in cat: return "ServiceNow ITSM"
    if "summary" in cat: return "Reporting & Tracking"
    return str(category).strip().title()

# ════════════════════════════════════════════════════════════
# CHATBOT ENGINE
# ════════════════════════════════════════════════════════════
def chatbot_response(user_msg, df_master, df_exploded,
                     uploaded_file_bytes=None,
                     uploaded_file_name=None):
    msg = user_msg.lower().strip()
    all_services = (sorted(df_exploded["Service"].unique().tolist())
                    if df_exploded is not None else [])
    all_vendors  = (sorted(df_master["Vendor"].unique().tolist())
                    if df_master  is not None else [])
    all_cats     = (sorted(df_master["Category"].unique().tolist())
                    if df_master  is not None else [])

    # ── File uploaded via chat ──
    if uploaded_file_bytes is not None and uploaded_file_name is not None:
        ext    = uploaded_file_name.rsplit(".",1)[-1].lower()
        result = extract_price_from_bytes(uploaded_file_bytes, ext)
        price  = result["price_num"]
        st.session_state["chat_upload_price"]    = price
        st.session_state["chat_upload_filename"] = uploaded_file_name
        st.session_state["chat_redirect_upload"] = True
        if price > 0:
            return {
                "type": "redirect_upload",
                "text": ("📄 I've analysed **{}** and found a price of **{}**.\n\n"
                         "🔄 **Redirecting you to Upload & Score** tab to compare this "
                         "against historical quotes...\n\n"
                         "You'll see the full price verdict there.".format(
                             uploaded_file_name, _fmt(price))),
                "price": price,
                "filename": uploaded_file_name,
            }
        else:
            return {
                "type": "redirect_upload",
                "text": ("📄 Uploaded **{}** but could not extract a price automatically.\n\n"
                         "🔄 **Redirecting to Upload & Score** — "
                         "you can enter the price manually there.".format(
                             uploaded_file_name)),
                "price": 0,
                "filename": uploaded_file_name,
            }

    # ── Greeting ──
    if any(w in msg for w in ["hello","hi","hey","good morning"]):
        return {
            "type": "text",
            "text": ("👋 Hello! I'm your **PwC Procurement Assistant**.\n\n"
                     "I can help you:\n"
                     "• 🔍 Find which vendors quoted for a service\n"
                     "• 💰 Compare prices across vendors\n"
                     "• 📋 Check if a service combination was quoted\n"
                     "• 📤 Upload a quote file — I'll score it automatically\n\n"
                     "Try: *'Who quoted for Cisco Catalyst?'* or "
                     "upload a quote file in the box below.")
        }

    # ── Help ──
    if any(w in msg for w in ["help","what can","capabilities"]):
        return {
            "type": "text",
            "text": ("**I can answer:**\n\n"
                     "🔍 *'Who has quoted for [service]?'*\n"
                     "💰 *'Compare [service] prices'*\n"
                     "🏢 *'What does [vendor] offer?'*\n"
                     "📋 *'Show [category] quotes'*\n"
                     "⚖️ *'Is [vendor] competitive?'*\n"
                     "🔄 *'Have [A] and [B] been quoted together?'*\n"
                     "📤 Upload a quote file to auto-score it")
        }

    # ── Stats / summary ──
    if any(w in msg for w in ["how many","total","count","summary","overview"]):
        n_quotes = len(df_master)
        n_vendors= df_master["Vendor"].nunique()
        n_svcs   = df_exploded["Service"].nunique()
        n_cats   = df_master["Category"].nunique()
        return {
            "type": "text",
            "text": ("📊 **Catalog Summary**\n\n"
                     "📄 **{}** total quotation files\n"
                     "🏢 **{}** unique vendors\n"
                     "🛠 **{}** unique services\n"
                     "📂 **{}** categories\n\n"
                     "Categories: {}".format(
                         n_quotes, n_vendors, n_svcs, n_cats,
                         " · ".join(all_cats)))
        }

    # ── List vendors ──
    if any(p in msg for p in ["list vendors","all vendors","show vendors","who are"]):
        lines = []
        for v in all_vendors:
            n_q  = len(df_master[df_master["Vendor"]==v])
            n_s  = df_exploded[df_exploded["Vendor"]==v]["Service"].nunique()
            lines.append("**{}** — {} quotes · {} services".format(v,n_q,n_s))
        return {"type":"text","text":"**Vendors in catalog:**\n\n" + "\n".join(lines)}

    # ── Vendor match ──
    matched_vendor = None
    for v in all_vendors:
        if v.lower() in msg:
            matched_vendor = v; break

    # ── Service match ──
    matched_services = []
    for svc in all_services:
        svc_words = [w for w in svc.lower().split() if len(w) > 3]
        if any(w in msg for w in svc_words) or svc.lower() in msg:
            matched_services.append(svc)
    matched_services = matched_services[:8]

    # ── Category match ──
    matched_cat = None
    for cat in all_cats:
        if cat.lower() in msg or any(w in msg for w in cat.lower().split() if len(w)>3):
            matched_cat = cat; break

    # ── Vendor + service ──
    if matched_vendor and matched_services:
        svc = matched_services[0]
        d   = df_exploded[(df_exploded["Vendor"]==matched_vendor)
                          & (df_exploded["Service"]==svc)]
        if not d.empty:
            files  = d["File Name"].unique().tolist()
            prices = []
            for f in files:
                rows = df_master[df_master["File Name"]==f]
                if len(rows) > 0:
                    qp = _parse_num(str(rows["Quoted Price"].values[0])
                                    if "Quoted Price" in rows.columns else "0")
                    if qp > 0: prices.append(qp)
            price_txt = ("Price on record: **{}**".format(_fmt(prices[0]))
                         if prices else "No price data on record.")
            return {
                "type": "vendor_service",
                "text": ("✅ **{}** has quoted for **{}**.\n\n"
                         "{}\n📄 Files: {}".format(
                             matched_vendor, svc, price_txt,
                             ", ".join(files[:3]))),
            }
        else:
            return {
                "type": "text",
                "text": ("❌ **{}** has **not** quoted for **{}**.\n\n"
                         "💡 Ask *'Who has quoted for {}?'*".format(
                             matched_vendor, svc, svc))
            }

    # ── Compare prices ──
    if matched_services and any(w in msg for w in [
            "compare","price","cost","expensive","cheap","competitive","vs","versus"]):
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc].drop_duplicates(
            subset=["Vendor","File Name"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No quotes found for **{}**.".format(svc)}
        vendor_prices = {}
        for _, r in d.iterrows():
            v  = r["Vendor"]
            qp = _parse_num(str(r.get("Quoted Price","")).strip())
            ck = "px_{}".format(str(r.get("File Name","")).strip())
            ca = st.session_state.get(ck)
            ep = ca["price_num"] if ca else 0.0
            p  = ep if ep > 0 else qp
            if p > 0:
                vendor_prices[v] = min(vendor_prices.get(v,p), p)
        if not vendor_prices:
            vendors = d["Vendor"].unique().tolist()
            return {
                "type": "text",
                "text": ("📋 **{}** quoted by **{}** vendor(s): {}\n\n"
                         "No price data — upload quote files to score.".format(
                             svc, len(vendors), ", ".join(vendors)))
            }
        avg_p  = sum(vendor_prices.values())/len(vendor_prices)
        best_v = min(vendor_prices, key=vendor_prices.get)
        spread = round((max(vendor_prices.values())-min(vendor_prices.values()))
                       /min(vendor_prices.values())*100,1) \
            if min(vendor_prices.values()) > 0 else 0
        lines = []
        for v,p in sorted(vendor_prices.items(), key=lambda x:x[1]):
            pct = round((p-avg_p)/avg_p*100,1) if avg_p>0 else 0
            tag = ("🟠 Cheapest" if v==best_v
                   else "⚫ Most Exp." if p==max(vendor_prices.values())
                   else "⚪ Mid-range")
            lines.append("**{}**: {} ({})".format(v,_fmt(p),tag))
        return {
            "type": "comparison",
            "text": ("📊 **Price comparison — {}**\n\n{}\n\n"
                     "Spread: **{}%** · Cheapest: **{}** at **{}**".format(
                         svc, "\n".join(lines), spread,
                         best_v, _fmt(min(vendor_prices.values())))),
            "service": svc,
            "vendor_prices": vendor_prices,
            "avg": avg_p,
            "best_vendor": best_v,
        }

    # ── Who quoted ──
    if matched_services and any(w in msg for w in [
            "who","vendor","quoted","available","offered"]):
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc].drop_duplicates(
            subset=["Vendor"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No vendor quoted for **{}**.".format(svc)}
        vendors = d["Vendor"].unique().tolist()
        n_files = df_exploded[df_exploded["Service"]==svc]["File Name"].nunique()
        return {
            "type": "who_quoted",
            "text": ("✅ **{}** vendor(s) quoted for **{}**:\n\n"
                     "{}\n\n📄 {} total files".format(
                         len(vendors), svc,
                         "\n".join(["• **{}**".format(v) for v in vendors]),
                         n_files)),
            "service": svc,
            "vendors": vendors,
        }

    # ── Vendor profile ──
    if matched_vendor and any(w in msg for w in [
            "offer","service","provide","what does","capability","profile"]):
        d    = df_exploded[df_exploded["Vendor"]==matched_vendor]
        svcs = sorted(d["Service"].unique().tolist())
        cats = sorted(d["Category"].unique().tolist())
        n_q  = len(df_master[df_master["Vendor"]==matched_vendor])
        return {
            "type": "vendor_profile",
            "text": ("🏢 **{}**\n\n"
                     "📂 Categories: {}\n"
                     "📄 Total quotes: {}\n"
                     "🛠 Services ({}):\n{}".format(
                         matched_vendor,
                         ", ".join(cats), n_q, len(svcs),
                         "\n".join(["• {}".format(s)
                                    for s in svcs[:12]])
                         + ("\n…+{} more".format(len(svcs)-12)
                            if len(svcs) > 12 else ""))),
        }

    # ── Category overview ──
    if matched_cat:
        d_cat  = df_master[df_master["Category"]==matched_cat]
        n_v    = d_cat["Vendor"].nunique()
        n_q    = len(d_cat)
        n_svc  = df_exploded[df_exploded["Category"]==matched_cat]["Service"].nunique()
        vendors= sorted(d_cat["Vendor"].unique().tolist())
        return {
            "type": "text",
            "text": ("📂 **{}**\n\n"
                     "📄 {} quotations · 🏢 {} vendors · 🛠 {} services\n\n"
                     "Vendors: {}".format(
                         matched_cat, n_q, n_v, n_svc,
                         ", ".join(vendors))),
        }

    # ── Combination check ──
    if len(matched_services) >= 2 and any(w in msg for w in [
            "together","combination","both","and","all"]):
        svc_set  = set(matched_services[:3])
        vsmap    = defaultdict(set)
        for _, r in df_exploded.iterrows():
            vsmap[r["Vendor"]].add(r["Service"])
        full   = [v for v,s in vsmap.items() if svc_set.issubset(s)]
        partial= [v for v,s in vsmap.items()
                  if svc_set & s and v not in full]
        if full:
            return {
                "type": "text",
                "text": ("✅ **{}** vendor(s) quoted ALL of: {}\n\n"
                         "Full: {} | Partial: {}".format(
                             len(full),
                             " · ".join(["**{}**".format(s)
                                         for s in svc_set]),
                             ", ".join(full),
                             ", ".join(partial) if partial else "None"))
            }
        return {
            "type": "text",
            "text": ("⚠️ No single vendor covers all of: {}\n\n"
                     "Partial: {}".format(
                         " · ".join(["**{}**".format(s)
                                     for s in svc_set]),
                         ", ".join(partial) if partial else "None"))
        }

    # ── Service info ──
    if matched_services:
        svc = matched_services[0]
        d   = df_exploded[df_exploded["Service"]==svc]
        if not d.empty:
            vendors  = d["Vendor"].unique().tolist()
            n_quotes = d["File Name"].nunique()
            return {
                "type": "service_info",
                "text": ("📋 **{}**\n\n"
                         "🏢 {} vendor(s): {}\n"
                         "📄 {} quote files\n\n"
                         "💡 Ask *'compare {} prices'* for breakdown.".format(
                             svc, len(vendors),
                             ", ".join(vendors[:5]),
                             n_quotes, svc)),
            }

    # ── Fallback ──
    suggestions = []
    if matched_services:
        suggestions.append("*'Compare {}?'*".format(matched_services[0]))
    if matched_vendor:
        suggestions.append("*'What does {} offer?'*".format(matched_vendor))
    fallback = ("I couldn't find specific data for that query. "
                "Try asking about a vendor, service, or category.")
    if suggestions:
        fallback += "\n\n💡 Try: " + " or ".join(suggestions)
    return {"type":"text","text":fallback}

def render_chat_chart(resp):
    """Render inline chart for comparison responses."""
    if resp.get("type") == "comparison" and "vendor_prices" in resp:
        vp     = resp["vendor_prices"]
        avg    = resp.get("avg",0)
        best_v = resp.get("best_vendor","")
        sorted_vp = sorted(vp.items(), key=lambda x:x[1])
        bar_c  = [C_ORANGE if v==best_v
                  else C_DARK if p==max(vp.values())
                  else C_GREY_DARK
                  for v,p in sorted_vp]
        fig = go.Figure(go.Bar(
            x=[v for v,_ in sorted_vp],
            y=[p for _,p in sorted_vp],
            marker_color=bar_c, marker_line_width=0,
            text=[_fmt(p) for _,p in sorted_vp],
            textposition="outside",
            textfont=dict(size=10,color=C_DARK)))
        if avg > 0:
            fig.add_hline(y=avg, line_dash="dash",
                           line_color=C_DARK, line_width=1.5,
                           annotation_text="Avg: {}".format(_fmt(avg)),
                           annotation_position="top right")
        fig.update_layout(
            height=240, plot_bgcolor=CBG, paper_bgcolor=CBG,
            margin=dict(l=5,r=10,t=20,b=5), font=CFONT,
            yaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,
                       zeroline=False),
            xaxis=dict(tickangle=-10,tickfont=dict(size=10)),
            bargap=0.45, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════
# UI HELPERS
# ════════════════════════════════════════════════════════════
def vpill(v, color=None):
    c = color or C_DARK
    return ("<span class='vbadge' style='background:{}'>{}</span>"
            .format(c,v))

def sec(txt, caption=""):
    st.markdown("<span class='sec-head'>{}</span>".format(txt),
                unsafe_allow_html=True)
    if caption:
        st.markdown("<div style='font-size:0.82em;color:#7D7D7D;"
                    "margin:-8px 0 10px'>{}</div>".format(caption),
                    unsafe_allow_html=True)

def kpi_box(col, val, lbl, bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div></div>".format(bg,val,lbl),
        unsafe_allow_html=True)

def mini_kpi(col, val, lbl, bg, icon=""):
    col.markdown(
        "<div class='kpi-box' style='background:{};min-height:82px'>"
        "<div style='font-size:1.4em;margin-bottom:2px'>{}</div>"
        "<div class='kpi-value' style='font-size:1.35em'>{}</div>"
        "<div class='kpi-label'>{}</div></div>".format(bg,icon,val,lbl),
        unsafe_allow_html=True)

def pwc_header(title, subtitle="", eyebrow="IT Procurement"):
    st.markdown(
        "<div style='background:#2D2D2D;color:white;padding:22px 28px;"
        "border-radius:4px;border-left:6px solid #D04A02;"
        "margin-bottom:20px'>"
        "<div style='font-size:0.70em;font-weight:700;letter-spacing:2px;"
        "text-transform:uppercase;color:#D04A02;margin-bottom:6px'>{}</div>"
        "<h1 style='margin:0;font-size:1.45em;font-weight:700;color:white;"
        "font-family:Georgia,serif'>{}</h1>"
        "{}</div>".format(
            eyebrow, title,
            "<p style='margin:7px 0 0;opacity:0.55;font-size:0.84em'>"
            "{}</p>".format(subtitle) if subtitle else ""),
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
        height=height, plot_bgcolor=CBG, paper_bgcolor=CBG,
        margin=dict(l=5,r=10,t=36 if title else 16,b=10),
        font=CFONT,
        yaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,zeroline=False),
        bargap=0.35, showlegend=False)
    if title:
        fig_obj.update_layout(
            title=dict(text=title,
                       font=dict(size=12,color=C_DARK,
                                 family="Georgia,serif"),
                       x=0, xanchor="left"))

# ════════════════════════════════════════════════════════════
# SESSION STATE
# ════════════════════════════════════════════════════════════
for k,v in [("chat_history",[]),
             ("gh_prices_loaded",False),
             ("real_analysis_done",False),
             ("chat_redirect_upload",False),
             ("chat_upload_price",0.0),
             ("chat_upload_filename",""),
             ("active_tab",0)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ════════════════════════════════════════════════════════════
# DATA
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state
        and st.session_state["uploaded_catalog_df"] is not None):
    df_master   = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "file"

NO_DATA = (df_master is None or df_exploded is None
           or df_master.empty)

vendor_color_map = {}
if not NO_DATA:
    vendors_sorted = sorted(df_master["Vendor"].unique())
    for i,v in enumerate(vendors_sorted):
        vendor_color_map[v] = CHART_SEQ[i % len(CHART_SEQ)]

# ════════════════════════════════════════════════════════════
# SIDEBAR — minimal, navigation only
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        "<div style='text-align:center;padding:22px 0 16px'>"
        "<div style='font-size:2.2em'>📋</div>"
        "<div style='font-size:1.05em;font-weight:700;color:white;"
        "margin:6px 0 2px;font-family:Georgia,serif'>"
        "IT Procurement</div>"
        "<div style='font-size:0.68em;color:#B0B0B0;"
        "letter-spacing:1.2px;text-transform:uppercase'>"
        "Intelligence Dashboard</div></div>"
        "<hr style='border-color:#D04A02;border-width:2px;"
        "margin:0 0 16px'>",
        unsafe_allow_html=True)

    if not NO_DATA:
        st.markdown(
            "<p style='color:#B0B0B0;font-size:0.75em;margin:4px 0'>"
            "📄 {} quotes</p>"
            "<p style='color:#B0B0B0;font-size:0.75em;margin:4px 0'>"
            "🏢 {} vendors</p>"
            "<p style='color:#B0B0B0;font-size:0.75em;margin:4px 0'>"
            "🛠 {} services</p>"
            "<p style='color:#B0B0B0;font-size:0.75em;margin:4px 0'>"
            "📂 {} categories</p>".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique(),
                df_master["Category"].nunique()),
            unsafe_allow_html=True)

    if DATA_SOURCE == "uploaded":
        st.markdown(
            "<div style='background:#D04A02;color:white;"
            "padding:6px 10px;border-radius:2px;"
            "font-size:0.74em;font-weight:700;"
            "text-align:center;margin:10px 0'>"
            "📤 UPLOADED CATALOG</div>",
            unsafe_allow_html=True)
        if st.button("🔄 Reset to Default",
                      use_container_width=True):
            st.session_state["uploaded_catalog_df"]  = None
            st.session_state["uploaded_catalog_exp"] = None
            st.rerun()

    st.markdown(
        "<hr style='border-color:#4A4A4A;margin:14px 0'>"
        "<p style='color:#7D7D7D;font-size:0.72em;"
        "margin:4px 0;text-align:center'>"
        "Filters available on each tab</p>",
        unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
pwc_header(
    "Procurement Intelligence Dashboard",
    "Catalog overview · Browse & chat · Upload & score · "
    "Vendor analysis · Real file analysis")

if not NO_DATA:
    k1,k2,k3,k4 = st.columns(4)
    kpi_box(k1, df_master["File Name"].nunique(),
            "Total Quotes", C_ORANGE)
    kpi_box(k2, df_exploded["Service"].nunique(),
            "Unique Services", C_DARK)
    kpi_box(k3, df_master["Vendor"].nunique(),
            "Vendors", C_MID)
    kpi_box(k4, df_master["Category"].nunique(),
            "Categories", C_GREY_DARK)
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
            lambda r: infer_subcategory(
                r.get("Category",""),
                r.get("Comments",""),
                r.get("File Name","")), axis=1)
        all_cats_ov = sorted([
            c for c in df_ov["Category"].unique()
            if str(c).strip() not in ["","nan"]])

        pwc_header("Catalog Overview",
                   "Full picture of all categories, vendors, "
                   "subcategories and services")

        # ── KPIs ──
        k0a,k0b,k0c,k0d,k0e = st.columns(5)
        mini_kpi(k0a, len(df_ov),
                 "Total Quotations", C_ORANGE, "📄")
        mini_kpi(k0b, df_ov["Vendor"].nunique(),
                 "Unique Vendors", C_DARK, "🏢")
        mini_kpi(k0c, df_ov["Category"].nunique(),
                 "Categories", C_MID, "📂")
        mini_kpi(k0d, df_ov["Subcategory"].nunique(),
                 "Subcategories", C_GREY_DARK, "🏷️")
        mini_kpi(k0e,
                 df_exploded["Service"].nunique()
                 if df_exploded is not None else "—",
                 "Unique Services", C_BLACK, "🛠")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Category cards ──
        sec("CATEGORIES AT A GLANCE")
        cat_stats = []
        for cat in all_cats_ov:
            d_cat = df_ov[df_ov["Category"]==cat]
            n_svcs = (df_exploded[
                df_exploded["Category"]==cat
            ]["Service"].nunique()
                      if df_exploded is not None else 0)
            cat_stats.append({
                "Category":cat,
                "Quotations":len(d_cat),
                "Vendors":d_cat["Vendor"].nunique(),
                "Subcategories":d_cat["Subcategory"].nunique(),
                "Services":n_svcs})
        cat_stats_df = pd.DataFrame(cat_stats).sort_values(
            "Quotations", ascending=False)

        CAT_ICONS = {
            "Cybersecurity":"🛡️",
            "Network & Telecom":"🌐",
            "Hosting":"🖥️",
            "M365 & Power Platform":"☁️",
            "IdAM":"🔑",
            "Service Management (SNow)":"⚙️",
            "Summary & Reporting":"📊",
        }
        rows_of_3 = [cat_stats_df.iloc[i:i+3]
                     for i in range(0,len(cat_stats_df),3)]
        for row_chunk in rows_of_3:
            cols = st.columns(len(row_chunk), gap="medium")
            for ci,(_, rs) in enumerate(row_chunk.iterrows()):
                cat_name = rs["Category"]
                icon  = CAT_ICONS.get(cat_name,"📁")
                cidx  = all_cats_ov.index(cat_name) \
                    if cat_name in all_cats_ov else 0
                color = CHART_SEQ[cidx % len(CHART_SEQ)]
                cols[ci].markdown(
                    "<div style='background:white;"
                    "border:1px solid #E0E0E0;"
                    "border-radius:6px;"
                    "padding:18px 16px;"
                    "border-top:4px solid {};height:100%'>"
                    "<div style='font-size:1.5em;"
                    "margin-bottom:6px'>{}</div>"
                    "<div style='font-size:0.93em;"
                    "font-weight:700;color:#2D2D2D;"
                    "margin-bottom:12px;"
                    "font-family:Georgia,serif'>{}</div>"
                    "<div style='display:flex;"
                    "gap:6px;flex-wrap:wrap'>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.74em;font-weight:700;"
                    "color:#2D2D2D'>📄 {} quotes</span>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.74em;font-weight:700;"
                    "color:#4A4A4A'>🏢 {} vendors</span>"
                    "<span style='background:#F3F3F3;"
                    "border-radius:3px;padding:3px 8px;"
                    "font-size:0.74em;font-weight:700;"
                    "color:#7D7D7D'>🛠 {} services</span>"
                    "</div></div>".format(
                        color, icon, cat_name,
                        rs["Quotations"],
                        rs["Vendors"],
                        rs["Services"]),
                    unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Charts ──
        sec("DISTRIBUTION — QUOTATIONS & VENDORS BY CATEGORY")
        ch_a, ch_b = st.columns(2, gap="large")
        with ch_a:
            fig_cq = go.Figure(go.Bar(
                x=cat_stats_df["Category"],
                y=cat_stats_df["Quotations"],
                marker_color=C_ORANGE,
                marker_line_width=0,
                text=cat_stats_df["Quotations"],
                textposition="outside",
                textfont=dict(size=11)))
            pwc_bar(fig_cq, "Quotations per Category")
            fig_cq.update_xaxes(
                tickangle=-30,tickfont=dict(size=9.5))
            st.plotly_chart(fig_cq,
                             use_container_width=True)
        with ch_b:
            fig_cv = go.Figure(go.Bar(
                x=cat_stats_df["Category"],
                y=cat_stats_df["Vendors"],
                marker_color=C_DARK,
                marker_line_width=0,
                text=cat_stats_df["Vendors"],
                textposition="outside",
                textfont=dict(size=11)))
            pwc_bar(fig_cv, "Vendors per Category")
            fig_cv.update_xaxes(
                tickangle=-30,tickfont=dict(size=9.5))
            st.plotly_chart(fig_cv,
                             use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        sec("CATALOG COMPOSITION")
        d_left, d_right = st.columns([1,2], gap="large")
        with d_left:
            fig_d = px.pie(
                cat_stats_df,
                values="Quotations",
                names="Category",
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
                paper_bgcolor=CBG,
                font=CFONT,
                showlegend=False)
            st.plotly_chart(fig_d,
                             use_container_width=True)
        with d_right:
            tbl = ["<table class='comp-table'>"
                   "<thead><tr>"
                   "<th>Category</th>"
                   "<th style='text-align:center'>Quotes</th>"
                   "<th style='text-align:center'>Vendors</th>"
                   "<th style='text-align:center'>Services</th>"
                   "<th style='text-align:center'>Subcats</th>"
                   "</tr></thead><tbody>"]
            for _, rs in cat_stats_df.iterrows():
                cat_n = rs["Category"]
                cidx  = (all_cats_ov.index(cat_n)
                         if cat_n in all_cats_ov else 0)
                color = CHART_SEQ[cidx % len(CHART_SEQ)]
                icon  = CAT_ICONS.get(cat_n,"📁")
                tbl.append(
                    "<tr><td>"
                    "<span style='border-left:4px solid {};"
                    "padding-left:8px;font-weight:600'>"
                    "{} {}</span></td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#D04A02'>{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#4A4A4A'>{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#7D7D7D'>{}</td>"
                    "<td style='text-align:center;"
                    "font-weight:700;color:#2D2D2D'>{}</td>"
                    "</tr>".format(
                        color, icon, cat_n,
                        rs["Quotations"], rs["Vendors"],
                        rs["Services"],
                        rs["Subcategories"]))
            tbl.append("</tbody></table>")
            st.markdown("".join(tbl),
                         unsafe_allow_html=True)

        # ── Per-category drill-down ──
        st.markdown("<br>", unsafe_allow_html=True)
        sec("DRILL-DOWN BY CATEGORY")
        ordered_cats = (
            ["Cybersecurity"]
            + [c for c in all_cats_ov
               if c != "Cybersecurity"])
        cat_subtabs = st.tabs([
            "🛡️ {}".format(c)
            if c=="Cybersecurity" else c
            for c in ordered_cats])

        for tidx, cat_name in enumerate(ordered_cats):
            with cat_subtabs[tidx]:
                d_cat = df_ov[
                    df_ov["Category"]==cat_name].copy()
                if d_cat.empty:
                    st.info("No data."); continue

                ck1,ck2,ck3,ck4 = st.columns(4)
                n_q   = len(d_cat)
                n_v   = d_cat["Vendor"].nunique()
                n_sub = d_cat["Subcategory"].nunique()
                n_svc = (df_exploded[
                    df_exploded["Category"]==cat_name
                ]["Service"].nunique()
                          if df_exploded is not None else 0)
                mini_kpi(ck1,n_q, "Quotations",
                          C_ORANGE,"📄")
                mini_kpi(ck2,n_v, "Vendors",
                          C_DARK,"🏢")
                mini_kpi(ck3,n_svc,"Services",
                          C_MID,"🛠")
                mini_kpi(ck4,n_sub,"Subcategories",
                          C_GREY_DARK,"🏷️")
                st.markdown("<br>",
                             unsafe_allow_html=True)

                sub_stats = (
                    d_cat.groupby("Subcategory")
                    .agg(
                        Quotations=("File Name","count"),
                        Vendors=("Vendor","nunique"))
                    .reset_index()
                    .sort_values("Quotations",
                                  ascending=False))

                sc1,sc2 = st.columns([1,1], gap="medium")
                with sc1:
                    st.markdown(
                        "<div style='font-size:0.77em;"
                        "font-weight:700;letter-spacing:0.8px;"
                        "text-transform:uppercase;"
                        "color:#2D2D2D;margin-bottom:8px'>"
                        "Subcategories</div>",
                        unsafe_allow_html=True)
                    fig_s = go.Figure(go.Bar(
                        x=sub_stats["Quotations"],
                        y=sub_stats["Subcategory"],
                        orientation="h",
                        marker_color=C_ORANGE,
                        marker_line_width=0,
                        text=sub_stats["Quotations"],
                        textposition="outside"))
                    fig_s.update_layout(
                        height=max(200,
                                   len(sub_stats)*36),
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=8,b=8),
                        font=CFONT,
                        xaxis=dict(
                            showgrid=True,
                            gridcolor=C_GREY_LITE,
                            zeroline=False),
                        yaxis=dict(
                            autorange="reversed",
                            tickfont=dict(size=9.5)),
                        bargap=0.3)
                    st.plotly_chart(
                        fig_s,
                        use_container_width=True)
                with sc2:
                    st.markdown(
                        "<div style='font-size:0.77em;"
                        "font-weight:700;letter-spacing:0.8px;"
                        "text-transform:uppercase;"
                        "color:#2D2D2D;margin-bottom:8px'>"
                        "Vendors</div>",
                        unsafe_allow_html=True)
                    vnd_stats = (
                        d_cat.groupby("Vendor")
                        .agg(Quotations=(
                            "File Name","count"))
                        .reset_index()
                        .sort_values("Quotations",
                                      ascending=False))
                    fig_v = go.Figure(go.Bar(
                        x=vnd_stats["Quotations"],
                        y=vnd_stats["Vendor"],
                        orientation="h",
                        marker_color=C_DARK,
                        marker_line_width=0,
                        text=vnd_stats["Quotations"],
                        textposition="outside"))
                    fig_v.update_layout(
                        height=max(200,
                                   len(vnd_stats)*36),
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=40,t=8,b=8),
                        font=CFONT,
                        xaxis=dict(
                            showgrid=True,
                            gridcolor=C_GREY_LITE,
                            zeroline=False),
                        yaxis=dict(
                            autorange="reversed",
                            tickfont=dict(size=9.5)),
                        bargap=0.3)
                    st.plotly_chart(
                        fig_v,
                        use_container_width=True)

                # File list
                st.markdown(
                    "<div style='font-size:0.77em;"
                    "font-weight:700;text-transform:uppercase;"
                    "color:#2D2D2D;margin:12px 0 8px'>"
                    "📄 All {} files</div>".format(
                        len(d_cat)),
                    unsafe_allow_html=True)
                file_tbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>File Name</th><th>Vendor</th>"
                    "<th>Subcategory</th>"
                    "<th>Services</th>"
                    "</tr></thead><tbody>"]
                for fi,(_, fr) in enumerate(
                        d_cat.sort_values(
                            "Subcategory").iterrows()):
                    bg  = ("white" if fi%2==0
                           else "#F8F8F8")
                    vc  = vendor_color_map.get(
                        fr["Vendor"], C_DARK)
                    cmt = str(
                        fr.get("Comments","")
                    ).replace("\n"," · ")[:80]
                    file_tbl.append(
                        "<tr style='background:{}'>"
                        "<td style='font-family:monospace;"
                        "font-size:0.77em;"
                        "word-break:break-all'>{}</td>"
                        "<td>{}</td>"
                        "<td style='color:#7D7D7D;"
                        "font-size:0.81em'>{}</td>"
                        "<td style='font-size:0.79em'>"
                        "{}</td>"
                        "</tr>".format(
                            bg,
                            fr.get("File Name",""),
                            vpill(fr["Vendor"],vc),
                            fr["Subcategory"],cmt))
                file_tbl.append("</tbody></table>")
                st.markdown("".join(file_tbl),
                             unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# TAB 1 — BROWSE & VERDICT
# Side-by-side: LEFT = Chatbot · RIGHT = Manual browser
# ════════════════════════════════════════════════════════════
with tab1:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        # ── Redirect if file uploaded in chat ──
        if st.session_state.get("chat_redirect_upload"):
            st.session_state["chat_redirect_upload"] = False
            st.info("📤 Redirecting to Upload & Score tab…")

        pwc_header(
            "Browse & Verdict",
            "Chat with the assistant · or manually "
            "browse services for deep analysis")

        # ── TOP FILTER BAR (replaces sidebar filters) ──
        st.markdown(
            "<div class='filter-bar'>",
            unsafe_allow_html=True)
        fb1,fb2,fb3 = st.columns(3)
        all_cats_f = (
            ["All"]
            + sorted([c for c in df_master[
                "Category"].unique()
                       if str(c).strip()
                       not in ["","nan"]]))
        with fb1:
            st.markdown(
                "<p style='color:#D04A02;"
                "font-size:0.78em;font-weight:700;"
                "margin-bottom:4px;"
                "letter-spacing:0.5px;"
                "text-transform:uppercase'>"
                "📂 CATEGORY</p>",
                unsafe_allow_html=True)
            sel_cat = st.selectbox(
                "cat", all_cats_f,
                label_visibility="collapsed",
                key="bv_cat")
        vpool = (df_master
                  if sel_cat=="All"
                  else df_master[
                      df_master["Category"]==sel_cat])
        all_vendors_f = (
            ["All"]
            + sorted([v for v in vpool[
                "Vendor"].unique()
                       if str(v).strip()
                       not in ["","nan"]]))
        with fb2:
            st.markdown(
                "<p style='color:#D04A02;"
                "font-size:0.78em;font-weight:700;"
                "margin-bottom:4px;"
                "letter-spacing:0.5px;"
                "text-transform:uppercase'>"
                "🏢 VENDOR</p>",
                unsafe_allow_html=True)
            sel_vendor = st.selectbox(
                "ven", all_vendors_f,
                label_visibility="collapsed",
                key="bv_ven")
        with fb3:
            st.markdown(
                "<p style='color:#D04A02;"
                "font-size:0.78em;font-weight:700;"
                "margin-bottom:4px;"
                "letter-spacing:0.5px;"
                "text-transform:uppercase'>"
                "🔍 SEARCH SERVICE</p>",
                unsafe_allow_html=True)
            svc_filter = st.text_input(
                "svc", placeholder="Type to filter…",
                label_visibility="collapsed",
                key="bv_svc_search")
        st.markdown("</div>", unsafe_allow_html=True)

        # Apply filters
        d_filt_bv = df_exploded.copy()
        if sel_cat != "All":
            d_filt_bv = d_filt_bv[
                d_filt_bv["Category"]==sel_cat]
        if sel_vendor != "All":
            d_filt_bv = d_filt_bv[
                d_filt_bv["Vendor"]==sel_vendor]
        avail_svcs = sorted([
            s for s in d_filt_bv["Service"].unique()
            if str(s).strip() not in ["","nan"]])
        if svc_filter:
            avail_svcs = [
                s for s in avail_svcs
                if svc_filter.lower() in s.lower()]

        # ── Side-by-side layout ──
        left_col, right_col = st.columns(
            [1, 1], gap="large")

        # ════════════════════════════════
        # LEFT — CHATBOT
        # ════════════════════════════════
        with left_col:
            st.markdown(
                "<div class='chat-header'>"
                "<span style='font-weight:700;"
                "font-size:0.96em;"
                "font-family:Georgia,serif'>"
                "💬 Procurement Assistant</span>"
                "<br><span style='font-size:0.76em;"
                "opacity:0.65'>"
                "Ask anything · Upload a quote file "
                "for instant scoring</span>"
                "</div>",
                unsafe_allow_html=True)

            # Chat history
            chat_html = "<div class='chat-outer'>"
            if not st.session_state["chat_history"]:
                chat_html += (
                    "<div class='chat-wrap'>"
                    "<div class='msg-bot'>"
                    "👋 Hello! Ask me about any "
                    "<b>service, vendor or price</b>."
                    "<br><br>"
                    "Or <b>upload a quote file</b> "
                    "below — I'll auto-score it and "
                    "redirect you to the full verdict."
                    "<br><br>"
                    "<b>Try:</b><br>"
                    "• <i>Who quoted for Cisco Catalyst?</i>"
                    "<br>"
                    "• <i>Compare Palo Alto prices</i><br>"
                    "• <i>What does NTT Data offer?</i>"
                    "</div></div>")
            else:
                for turn in st.session_state[
                        "chat_history"]:
                    chat_html += (
                        "<div class='chat-wrap'>"
                        "<div class='msg-user'>{}</div>"
                        "</div>".format(turn["user"]))
                    bot_txt = (
                        turn["bot_text"]
                        .replace("\n","<br>")
                        .replace("**","<b>",1)
                        .replace("**","</b>",1))
                    chat_html += (
                        "<div class='chat-wrap'>"
                        "<div class='msg-bot'>{}</div>"
                        "</div>".format(bot_txt))
            chat_html += "</div>"
            st.markdown(chat_html,
                         unsafe_allow_html=True)

            # Render last chart
            if st.session_state["chat_history"]:
                last_resp = st.session_state[
                    "chat_history"][-1].get(
                    "bot_resp")
                if last_resp:
                    render_chat_chart(last_resp)

            # File upload in chat
            chat_file = st.file_uploader(
                "📎 Upload quote file for instant scoring",
                type=["pdf","xlsx","xls","docx"],
                key="chat_file_upload",
                label_visibility="visible")

            # Text input
            with st.form("chat_form",
                          clear_on_submit=True):
                ci1,ci2 = st.columns([5,1])
                with ci1:
                    user_input = st.text_input(
                        "msg",
                        placeholder="Ask a question…",
                        label_visibility="collapsed")
                with ci2:
                    sent = st.form_submit_button(
                        "Send", type="primary",
                        use_container_width=True)

            if sent or (
                    chat_file is not None
                    and st.session_state.get(
                        "last_chat_file") != (
                        chat_file.name
                        if chat_file else None)):

                if chat_file is not None:
                    st.session_state[
                        "last_chat_file"] = (
                        chat_file.name)
                    file_bytes = chat_file.read()
                    resp = chatbot_response(
                        "Analysing uploaded file",
                        df_master, df_exploded,
                        uploaded_file_bytes=file_bytes,
                        uploaded_file_name=chat_file.name)
                    # Store upload data for tab2
                    st.session_state[
                        "tab2_upload_price"] = (
                        resp.get("price",0))
                    st.session_state[
                        "tab2_upload_fname"] = (
                        resp.get("filename",""))
                    st.session_state[
                        "tab2_file_bytes"] = file_bytes
                    st.session_state[
                        "tab2_file_ext"] = (
                        chat_file.name.rsplit(
                            ".",1)[-1])
                elif sent and user_input.strip():
                    resp = chatbot_response(
                        user_input.strip(),
                        df_master, df_exploded)
                else:
                    resp = None

                if resp:
                    st.session_state[
                        "chat_history"].append({
                        "user": (
                            user_input.strip()
                            if sent and user_input.strip()
                            else "📎 {}".format(
                                chat_file.name
                                if chat_file else "")),
                        "bot_text": resp["text"],
                        "bot_resp": resp,
                    })
                    st.rerun()

            # Quick chips
            st.markdown(
                "<div style='margin:8px 0 4px;"
                "font-size:0.74em;color:#7D7D7D;"
                "font-weight:600;letter-spacing:0.6px'>"
                "QUICK QUESTIONS</div>",
                unsafe_allow_html=True)
            chips = [
                "Who quoted for Cisco Catalyst?",
                "Compare Palo Alto prices",
                "What does NTT Data offer?",
                "List all vendors",
                "Show Cybersecurity quotes",
            ]
            chip_cols = st.columns(len(chips))
            for i,chip in enumerate(chips):
                if chip_cols[i].button(
                        chip, key="chip_{}".format(i),
                        use_container_width=True):
                    resp = chatbot_response(
                        chip, df_master, df_exploded)
                    st.session_state[
                        "chat_history"].append({
                        "user": chip,
                        "bot_text": resp["text"],
                        "bot_resp": resp,
                    })
                    st.rerun()

            if st.button("🗑 Clear chat",
                          key="clr_chat"):
                st.session_state["chat_history"] = []
                st.rerun()

        # ════════════════════════════════
        # RIGHT — MANUAL BROWSER + ANALYSIS
        # ════════════════════════════════
        with right_col:
            sec("SERVICE BROWSER")
            selected_svcs = st.multiselect(
                "Select services to analyse",
                options=avail_svcs,
                default=[],
                label_visibility="visible",
                help="Select one or more services")

            if not selected_svcs:
                # ── All-services overview chart ──
                sec("SERVICE COMPETITIVENESS MAP",
                    "Orange = multiple vendors · Grey = single vendor")
                svc_summary = (
                    df_exploded.groupby("Service")[
                        "Vendor"].nunique()
                    .reset_index()
                    .sort_values("Vendor",
                                  ascending=False))
                svc_summary.columns = [
                    "Service","Vendor Count"]
                svc_top = svc_summary.head(20)
                fig_sv = go.Figure(go.Bar(
                    x=svc_top["Vendor Count"],
                    y=svc_top["Service"].apply(
                        lambda x: x[:48]),
                    orientation="h",
                    marker_color=[
                        C_ORANGE if v > 1
                        else C_GREY
                        for v in svc_top[
                            "Vendor Count"]],
                    marker_line_width=0,
                    text=svc_top["Vendor Count"],
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
                    bargap=0.28,
                    showlegend=False)
                st.plotly_chart(
                    fig_sv,
                    use_container_width=True)

                n_multi  = svc_summary[
                    svc_summary["Vendor Count"]>1
                ].shape[0]
                n_single = svc_summary[
                    svc_summary["Vendor Count"]==1
                ].shape[0]
                sm1,sm2,sm3 = st.columns(3)
                kpi_box(sm1,
                         len(svc_summary),
                         "Total Services", C_DARK)
                kpi_box(sm2, n_multi,
                         "Competitive (2+ vendors)",
                         C_ORANGE)
                kpi_box(sm3, n_single,
                         "Single Vendor Only",
                         C_GREY_DARK)

            else:
                # ════════════════════════════
                # DEEP ANALYSIS
                # ════════════════════════════
                d_sel = (
                    d_filt_bv[
                        d_filt_bv["Service"].isin(
                            selected_svcs)]
                    .copy())
                if d_sel.empty:
                    st.warning(
                        "No quotations found for "
                        "selected service(s).")
                else:
                    has_price = (
                        "Quoted Price" in d_sel.columns)

                    # Collect vendor prices
                    vendor_prices_map = {}
                    for _, r in d_sel.drop_duplicates(
                            subset=["Vendor",
                                     "File Name"]
                    ).iterrows():
                        v  = r["Vendor"]
                        qp = _parse_num(
                            str(r.get(
                                "Quoted Price","")
                            ).strip())
                        ck = "px_{}".format(
                            str(r.get(
                                "File Name","")
                            ).strip())
                        ca = st.session_state.get(ck)
                        ep = (ca["price_num"]
                               if ca else 0.0)
                        ref = ep if ep > 0 else qp
                        if ref > 0:
                            vendor_prices_map[v] = min(
                                vendor_prices_map.get(
                                    v, ref), ref)
                    if not vendor_prices_map:
                        for _, r in d_sel.drop_duplicates(
                                subset=["Vendor"]
                        ).iterrows():
                            v  = r["Vendor"]
                            qp = _parse_num(
                                str(r.get(
                                    "Quoted Price",""
                                )).strip())
                            if (qp > 0
                                    and v not in
                                    vendor_prices_map):
                                vendor_prices_map[v] = qp

                    # ── Coverage verdict ──
                    vsmap = defaultdict(set)
                    for _, r in d_sel.iterrows():
                        vsmap[r["Vendor"]].add(
                            r["Service"])
                    full_cov = [
                        v for v,s in vsmap.items()
                        if set(selected_svcs
                                ).issubset(s)]
                    partial  = [
                        v for v,s in vsmap.items()
                        if set(selected_svcs) & s
                        and v not in full_cov]

                    if len(selected_svcs) == 1:
                        n_v = d_sel["Vendor"].nunique()
                        if n_v > 1:
                            st.markdown(
                                "<div class='verdict-good'>"
                                "✅ <b>{}</b> — previously "
                                "quoted by <b>{} vendors</b>. "
                                "Competitive benchmarking "
                                "available."
                                "</div>".format(
                                    selected_svcs[0],
                                    n_v),
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
                                "cover ALL {} services: "
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
                                    if partial
                                    else "None"),
                                unsafe_allow_html=True)

                    st.markdown("<br>",
                                 unsafe_allow_html=True)

                    # ── Price verdict KPIs ──
                    if vendor_prices_map:
                        vps   = vendor_prices_map
                        avg_p = (sum(vps.values())
                                 / len(vps))
                        best_v = min(
                            vps, key=vps.get)
                        worst_v= max(
                            vps, key=vps.get)
                        spread = round(
                            (max(vps.values())
                             - min(vps.values()))
                            / min(vps.values())
                            * 100, 1
                        ) if min(vps.values()) > 0 else 0

                        sec("PRICE VERDICT")
                        pv1,pv2,pv3 = st.columns(3)
                        pv1.markdown(
                            "<div class='scard "
                            "scard-orange'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#D04A02'>"
                            "Best Price</div>"
                            "<div style='font-size:"
                            "1.55em;font-weight:800;"
                            "color:#D04A02;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.74em;color:#7D7D7D;"
                            "margin-top:2px'>{}</div>"
                            "</div>".format(
                                _fmt(min(vps.values())),
                                best_v),
                            unsafe_allow_html=True)
                        pv2.markdown(
                            "<div class='scard "
                            "scard-dark'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#2D2D2D'>"
                            "Market Average</div>"
                            "<div style='font-size:"
                            "1.55em;font-weight:800;"
                            "color:#2D2D2D;"
                            "font-family:Georgia,serif'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.74em;color:#7D7D7D;"
                            "margin-top:2px'>"
                            "{} vendor(s)</div>"
                            "</div>".format(
                                _fmt(avg_p),
                                len(vps)),
                            unsafe_allow_html=True)
                        pv3.markdown(
                            "<div class='scard "
                            "scard-grey'>"
                            "<div style='font-size:"
                            "0.67em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#7D7D7D'>"
                            "Price Spread</div>"
                            "<div style='font-size:"
                            "1.55em;font-weight:800;"
                            "color:#4A4A4A;"
                            "font-family:Georgia,serif'>"
                            "{}%</div>"
                            "<div style='font-size:"
                            "0.74em;color:#7D7D7D;"
                            "margin-top:2px'>"
                            "negotiation room</div>"
                            "</div>".format(spread),
                            unsafe_allow_html=True)

                        st.markdown("<br>",
                                     unsafe_allow_html=True)

                        # ── Price comparison chart ──
                        sec("PRICE COMPARISON CHART",
                            "Orange = best · Dark = highest · "
                            "Dashed = market average")
                        sorted_vp = sorted(
                            vps.items(),
                            key=lambda x:x[1])
                        bar_c = [
                            C_ORANGE
                            if v==best_v
                            else C_DARK
                            if v==worst_v
                            else C_GREY_DARK
                            for v,_ in sorted_vp]
                        fig_cmp = go.Figure(go.Bar(
                            x=[v for v,_ in sorted_vp],
                            y=[p for _,p in sorted_vp],
                            marker_color=bar_c,
                            marker_line_width=0,
                            text=[_fmt(p)
                                  for _,p in sorted_vp],
                            textposition="outside",
                            textfont=dict(
                                size=11,
                                color=C_DARK)))
                        fig_cmp.add_hline(
                            y=avg_p,
                            line_dash="dash",
                            line_color=C_MID,
                            line_width=2,
                            annotation_text=
                            "Avg: {}".format(
                                _fmt(avg_p)),
                            annotation_font_color=C_MID,
                            annotation_position=
                            "top right")
                        fig_cmp.update_layout(
                            height=300,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=28,b=8),
                            font=CFONT,
                            yaxis=dict(
                                title="Price (USD)",
                                showgrid=True,
                                gridcolor=C_GREY_LITE,
                                zeroline=False),
                            xaxis=dict(
                                tickangle=-10,
                                tickfont=dict(
                                    size=10.5)),
                            bargap=0.4,
                            showlegend=False)
                        st.plotly_chart(
                            fig_cmp,
                            use_container_width=True)

                        # ── Score table ──
                        sec("VENDOR SCORE CARD")
                        all_pv = list(vps.values())
                        tbl = [
                            "<table class='comp-table'>"
                            "<thead><tr>"
                            "<th>Rank</th>"
                            "<th>Vendor</th>"
                            "<th>Price</th>"
                            "<th>vs Average</th>"
                            "<th>Score</th>"
                            "<th>Verdict</th>"
                            "</tr></thead><tbody>"]
                        for rank,(v,p) in enumerate(
                                sorted(
                                    vps.items(),
                                    key=lambda x:x[1]),
                                1):
                            bg  = ("white"
                                    if rank%2==0
                                    else "#F8F8F8")
                            vc  = vendor_color_map.get(
                                v, C_DARK)
                            oth = [x for x in all_pv
                                   if x != p]
                            ps  = None
                            if p > 0 and oth:
                                ps,_,_,_,_ = price_score(
                                    p,oth)
                            sc     = score_color(ps)
                            pct    = (round(
                                (p-avg_p)/avg_p*100,1)
                                       if avg_p>0 else 0)
                            vs_txt = (
                                "{}% below ✅".format(
                                    abs(pct))
                                if pct < 0
                                else "{}% above ⚠️".format(
                                    abs(pct))
                                if pct > 0
                                else "At average")
                            vs_c = (
                                C_ORANGE if pct < 0
                                else C_DARK if pct > 10
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
                                "<td style='font-weight:"
                                "700;color:{}'>{}</td>"
                                "</tr>".format(
                                    bg, medal,
                                    vpill(v,vc),
                                    _fmt(p),
                                    vs_c, vs_txt,
                                    sc,
                                    ps if ps is not None
                                    else "—",
                                    sc, vt))
                        tbl.append("</tbody></table>")
                        st.markdown("".join(tbl),
                                     unsafe_allow_html=True)

                        # ── AI insight ──
                        pct_best = round(
                            (avg_p-min(vps.values()))
                            /avg_p*100,1) if avg_p>0 else 0
                        insight(
                            "<b>{}</b> is the most "
                            "competitive vendor at "
                            "<b>{}</b> — "
                            "<b>{}% below</b> the "
                            "market average. "
                            "Price spread of <b>{}%</b> "
                            "indicates "
                            "<b>{}</b>.".format(
                                best_v,
                                _fmt(min(vps.values())),
                                pct_best, spread,
                                "strong negotiation "
                                "potential"
                                if spread > 20
                                else "moderate "
                                "negotiation room"
                                if spread > 10
                                else "a competitive "
                                "market"))

                    # ── Per-service file detail ──
                    st.markdown("<br>",
                                 unsafe_allow_html=True)
                    sec("QUOTATION FILE DETAILS")
                    for svc in selected_svcs:
                        d_svc = (
                            d_sel[
                                d_sel["Service"]==svc]
                            .drop_duplicates(
                                subset=["Vendor",
                                         "File Name"])
                            .sort_values("Vendor"))
                        n_v = d_svc["Vendor"].nunique()
                        st.markdown(
                            "<div style='background:"
                            "white;border-left:4px solid "
                            "{};padding:10px 14px;"
                            "border-radius:2px;"
                            "margin:8px 0;font-weight:"
                            "700;font-size:0.88em'>"
                            "{}  ·  {} vendor(s)  ·  "
                            "{}</div>".format(
                                C_ORANGE
                                if n_v > 1
                                else C_GREY_DARK,
                                svc, n_v,
                                "✅ COMPETITIVE"
                                if n_v > 1
                                else "⚠️ SINGLE VENDOR"),
                            unsafe_allow_html=True)

                        all_p_svc = []
                        for _,r in d_svc.iterrows():
                            qp = _parse_num(str(r.get(
                                "Quoted Price",""
                            )).strip())
                            if qp > 0:
                                all_p_svc.append(qp)

                        rows_tbl = [
                            "<table class='comp-table'>"
                            "<thead><tr>"
                            "<th>Vendor</th>"
                            "<th>File</th>"]
                        if has_price:
                            rows_tbl.append(
                                "<th>Quoted Price</th>")
                        rows_tbl.append(
                            "<th>Score</th>"
                            "<th>Verdict</th>"
                            "<th>Open</th>"
                            "</tr></thead><tbody>")

                        for i,(_,row) in enumerate(
                                d_svc.iterrows()):
                            bg    = ("white"
                                      if i%2==0
                                      else "#F8F8F8")
                            vc2   = vendor_color_map.get(
                                row["Vendor"],C_DARK)
                            fname = str(row.get(
                                "File Name","")).strip()
                            url   = resolve_url(row)
                            qp_num= _parse_num(str(row.get(
                                "Quoted Price",""
                            )).strip())
                            ck    = "px_{}".format(fname)
                            ca    = st.session_state.get(ck)
                            ref   = (
                                ca["price_num"]
                                if ca and ca.get(
                                    "price_num",0) > 0
                                else qp_num
                                if qp_num > 0
                                else 0)
                            oth   = [p for p in all_p_svc
                                     if p != ref]
                            ps    = None; vt = "—"
                            vt_c  = C_GREY
                            if ref > 0 and oth:
                                ps,_,_,_,_ = price_score(
                                    ref, oth)
                                vt,_,vt_c = get_verdict(ps)
                            sc_c = score_color(ps)
                            link = (
                                "<a href='{}' "
                                "target='_blank' "
                                "style='color:#D04A02;"
                                "font-weight:600;"
                                "text-decoration:none'>"
                                "📂 Open</a>".format(url)
                                if url else "—")
                            rows_tbl.append(
                                "<tr style='background:"
                                "{}'>"
                                "<td>{}</td>"
                                "<td style='font-family:"
                                "monospace;font-size:"
                                "0.77em;word-break:"
                                "break-all'>{}</td>".format(
                                    bg,
                                    vpill(row["Vendor"],
                                           vc2),
                                    fname))
                            if has_price:
                                rows_tbl.append(
                                    "<td style='font-family:"
                                    "monospace;font-weight:"
                                    "700;color:#D04A02'>"
                                    "{}</td>".format(
                                        _fmt(qp_num)
                                        if qp_num > 0
                                        else "—"))
                            rows_tbl.append(
                                "<td style='text-align:"
                                "center'>"
                                "<span style='font-weight:"
                                "800;color:{}'>"
                                "{}</span></td>"
                                "<td style='font-weight:"
                                "700;color:{}'>{}</td>"
                                "<td>{}</td>"
                                "</tr>".format(
                                    sc_c,
                                    "{}/100".format(ps)
                                    if ps is not None
                                    else "—",
                                    vt_c, vt, link))
                        rows_tbl.append(
                            "</tbody></table>")
                        st.markdown(
                            "".join(rows_tbl),
                            unsafe_allow_html=True)

                        st.markdown("<br>",
                                     unsafe_allow_html=True)
                        if st.button(
                                "🔍 Extract Prices — "
                                "{}".format(svc[:38]),
                                key="ep2_{}".format(
                                    svc[:32]),
                                type="primary"):
                            prog = st.progress(0)
                            n    = len(d_svc)
                            for ki,(_,row2) in enumerate(
                                    d_svc.iterrows()):
                                f2 = str(row2.get(
                                    "File Name","")
                                ).strip()
                                ck2 = "px_{}".format(f2)
                                if st.session_state.get(
                                        ck2) is None:
                                    local = os.path.join(
                                        DEMO_DIR, f2)
                                    if os.path.exists(
                                            local):
                                        st.session_state[
                                            ck2] = extract_price_from_file(
                                            local)
                                    else:
                                        u2 = resolve_url(
                                            row2)
                                        if (u2
                                                and u2.startswith(
                                                    "http")
                                                and REQUESTS_OK):
                                            try:
                                                r2 = requests.get(
                                                    u2,
                                                    timeout=20)
                                                e2 = u2.split(
                                                    "?"
                                                )[0].rsplit(
                                                    ".",1
                                                )[-1].lower()
                                                st.session_state[
                                                    ck2] = extract_price_from_bytes(
                                                    r2.content,
                                                    e2)
                                            except Exception:
                                                pass
                                prog.progress((ki+1)/n)
                            prog.empty()
                            st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab2:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header(
            "Upload & Score",
            "Upload a new quote → auto-extract price "
            "→ compare vs historical → get verdict")

        # ── Check if redirected from chatbot ──
        prefill_price = st.session_state.get(
            "tab2_upload_price", 0.0)
        prefill_fname = st.session_state.get(
            "tab2_upload_fname", "")
        prefill_bytes = st.session_state.get(
            "tab2_file_bytes", None)
        prefill_ext   = st.session_state.get(
            "tab2_file_ext", "")

        if prefill_fname:
            st.success(
                "📎 File from chat: **{}** — "
                "Price: **{}**".format(
                    prefill_fname,
                    _fmt(prefill_price)
                    if prefill_price > 0
                    else "not found"))

        # ── Upload ──
        sec("STEP 1 — UPLOAD QUOTE FILE")
        uploaded = st.file_uploader(
            "Upload",
            type=["pdf","xlsx","xls","docx"],
            label_visibility="collapsed")

        new_price = 0.0; fname_up = ""
        if uploaded is not None:
            content   = uploaded.read()
            ext_up    = uploaded.name.rsplit(".",1)[-1]
            fname_up  = uploaded.name
            st.success(
                "Uploaded: **{}** ({} KB)".format(
                    fname_up,
                    round(len(content)/1024,1)))
            sec("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting…"):
                result    = extract_price_from_bytes(
                    content, ext_up)
                new_price = result["price_num"]
            if new_price > 0:
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
                    min_value=0.0, step=100.0,
                    value=0.0,
                    key="manual_price")
                if manual > 0:
                    new_price = manual
        elif prefill_price > 0:
            new_price = prefill_price
            fname_up  = prefill_fname

# ── Service selection + filter bar ──
        sec("STEP 3 — SELECT SERVICES & FILTERS")
        up1, up2, up3 = st.columns(3)
        with up1:
            all_cats_up = (
                ["All"]
                + sorted([c for c in df_master[
                    "Category"].unique()
                           if str(c).strip()
                           not in ["","nan"]]))
            cat_filter_up = st.selectbox(
                "📂 Filter by Category",
                options=all_cats_up,
                key="cat_up2")
        with up2:
            all_svcs_up = sorted([
                s for s in df_exploded[
                    "Service"].unique()
                if str(s).strip()
                not in ["","nan"]])
            svc_search_up = st.text_input(
                "🔍 Filter services",
                placeholder="Search…",
                key="svc_up2")
            if svc_search_up:
                all_svcs_up = [
                    s for s in all_svcs_up
                    if svc_search_up.lower()
                    in s.lower()]
        with up3:
            new_services = st.multiselect(
                "🛠 Select Services",
                options=all_svcs_up,
                key="new_svcs2")

        sec("STEP 4 — COMPARISON & VERDICT")
        if new_price <= 0 and not new_services:
            st.info(
                "Upload a file or enter a price, "
                "then select services to compare.")
        else:
            candidates = (
                df_exploded[
                    df_exploded["Service"].isin(
                        new_services)].copy()
                if new_services
                else df_exploded.copy())
            if cat_filter_up != "All":
                candidates = candidates[
                    candidates["Category"]
                    == cat_filter_up]
            cand_files = (
                candidates
                .drop_duplicates(
                    subset=["File Name","Vendor"])
                [["File Name","Vendor","Category",
                  "Hyperlink","Quoted Price"]]
                .copy())

            if cand_files.empty:
                st.warning(
                    "No historical quotes found "
                    "for selected services.")
            else:
                hist_prices = []; vendor_p_map = {}
                for _, r in cand_files.iterrows():
                    qp = _parse_num(str(r.get(
                        "Quoted Price","")).strip())
                    if qp > 0:
                        hist_prices.append(qp)
                        vendor_p_map[r["Vendor"]] = qp
                    ck = "px_{}".format(str(r.get(
                        "File Name","")).strip())
                    ca = st.session_state.get(ck)
                    if ca and ca.get("price_num",0) > 0:
                        hist_prices.append(
                            ca["price_num"])
                        vendor_p_map[r["Vendor"]] = (
                            ca["price_num"])

                if new_price > 0 and hist_prices:
                    ps, ps_lbl, avg_h, mn_h, mx_h = (
                        price_score(
                            new_price, hist_prices))
                    vt, vt_desc, vt_col = (
                        get_verdict(ps))
                    vt_css = score_css(ps)

                    # Verdict banner
                    verdict_bg = {
                        "orange": "#FFF5F0",
                        "mid":    "#F5F5F5",
                        "dark":   "#F0F0F0",
                        "grey":   "#FAFAFA",
                    }.get(vt_css, "#F5F5F5")
                    verdict_border = {
                        "orange": C_ORANGE,
                        "mid":    C_GREY_DARK,
                        "dark":   C_DARK,
                        "grey":   C_GREY,
                    }.get(vt_css, C_GREY_DARK)
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
                        "{}</div>"
                        "</div>".format(
                            verdict_bg,
                            verdict_border,
                            verdict_border,
                            vt, vt_desc),
                        unsafe_allow_html=True)

                    # Score KPIs
                    sv1,sv2,sv3,sv4 = st.columns(4)
                    sv1.markdown(
                        "<div class='scard "
                        "scard-orange'>"
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
                            else "N/A",
                            len(hist_prices)),
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
                        "{}</div>"
                        "</div>".format(
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
                            _fmt(avg_h),
                            _fmt(mn_h), _fmt(mx_h)),
                        unsafe_allow_html=True)
                    sv4.markdown(
                        "<div class='scard scard-grey'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#7D7D7D'>"
                        "vs Average</div>"
                        "<div style='font-size:0.98em;"
                        "font-weight:800;color:#4A4A4A;"
                        "margin-top:8px'>"
                        "{}</div>"
                        "</div>".format(ps_lbl),
                        unsafe_allow_html=True)

                    st.markdown("<br>",
                                 unsafe_allow_html=True)

                    # Positioning chart
                    sec("PRICE POSITIONING CHART",
                        "Orange = your quote · "
                        "Dark = historical · "
                        "Dashed = market average")
                    chart_data = []
                    for _, r in cand_files.iterrows():
                        fn  = str(r.get(
                            "File Name","")).strip()
                        qp  = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        ck  = "px_{}".format(fn)
                        ca  = st.session_state.get(ck)
                        ep  = ca["price_num"] if ca else 0.0
                        pval= ep if ep > 0 else qp
                        if pval > 0:
                            chart_data.append({
                                "Label":"{} / {}".format(
                                    r["Vendor"],fn[:10]),
                                "Price": pval,
                                "Type": "Historical"})
                    chart_data.append({
                        "Label": "★ YOUR QUOTE",
                        "Price": new_price,
                        "Type":  "New"})
                    cdf = pd.DataFrame(
                        chart_data
                    ).sort_values("Price")
                    bar_colors_up = [
                        C_ORANGE
                        if t == "New"
                        else C_DARK
                        for t in cdf["Type"]]
                    fig_up = go.Figure(go.Bar(
                        x=cdf["Label"],
                        y=cdf["Price"],
                        marker_color=bar_colors_up,
                        marker_line_width=0,
                        text=cdf["Price"].apply(_fmt),
                        textposition="outside",
                        textfont=dict(size=10)))
                    fig_up.add_hline(
                        y=avg_h,
                        line_dash="dash",
                        line_color=C_GREY_DARK,
                        line_width=2,
                        annotation_text=
                        "Market Avg: {}".format(
                            _fmt(avg_h)),
                        annotation_position=
                        "top right")
                    fig_up.update_layout(
                        height=380,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(
                            l=5,r=10,t=20,b=10),
                        font=CFONT,
                        yaxis=dict(
                            title="Price (USD)",
                            showgrid=True,
                            gridcolor=C_GREY_LITE,
                            zeroline=False),
                        xaxis=dict(tickangle=-25),
                        bargap=0.3,
                        showlegend=False)
                    st.plotly_chart(
                        fig_up,
                        use_container_width=True)

                    # Insight
                    pct_vs = round(
                        (new_price - avg_h)
                        / avg_h * 100, 1
                    ) if avg_h > 0 else 0
                    insight(
                        "Your quote of <b>{}</b> is "
                        "<b>{}% {}</b> the market "
                        "average of <b>{}</b>. "
                        "Historical range: "
                        "<b>{}</b> – <b>{}</b>.".format(
                            _fmt(new_price),
                            abs(pct_vs),
                            "below" if pct_vs < 0
                            else "above",
                            _fmt(avg_h),
                            _fmt(mn_h), _fmt(mx_h)))

                    # Historical table
                    st.markdown("<br>",
                                 unsafe_allow_html=True)
                    sec("HISTORICAL QUOTES USED "
                        "FOR COMPARISON")
                    ht = [
                        "<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>Vendor</th>"
                        "<th>Category</th>"
                        "<th>File Name</th>"
                        "<th>Price</th>"
                        "<th>Source</th>"
                        "</tr></thead><tbody>"]
                    for i,(_,r) in enumerate(
                            cand_files.iterrows()):
                        bg  = ("white" if i%2==0
                                else "#F8F8F8")
                        vc  = vendor_color_map.get(
                            r["Vendor"],C_DARK)
                        fn  = str(r.get(
                            "File Name","")).strip()
                        qp  = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        ck  = "px_{}".format(fn)
                        ca  = st.session_state.get(ck)
                        ep  = (ca["price_num"]
                                if ca else 0.0)
                        pval= ep if ep>0 else qp
                        src = ("Extracted"
                                if ep>0 else "Catalog"
                                if qp>0 else "—")
                        src_c = (C_ORANGE
                                  if ep>0
                                  else C_GREY_DARK)
                        if pval <= 0: continue
                        ht.append(
                            "<tr style='background:{}'>"
                            "<td>{}</td>"
                            "<td style='color:#7D7D7D;"
                            "font-size:0.82em'>{}</td>"
                            "<td style='font-family:"
                            "monospace;font-size:0.78em;"
                            "word-break:break-all'>"
                            "{}</td>"
                            "<td style='font-family:"
                            "monospace;font-weight:700;"
                            "color:#D04A02'>{}</td>"
                            "<td><span style='background:"
                            "{};color:white;padding:"
                            "2px 7px;border-radius:2px;"
                            "font-size:0.75em;"
                            "font-weight:700'>"
                            "{}</span></td>"
                            "</tr>".format(
                                bg,
                                vpill(r["Vendor"],vc),
                                r["Category"],fn,
                                _fmt(pval),
                                src_c,src))
                    ht.append("</tbody></table>")
                    st.markdown(
                        "".join(ht),
                        unsafe_allow_html=True)

                elif not hist_prices:
                    st.info(
                        "No historical price data "
                        "for selected services. "
                        "Try selecting more services "
                        "or uploading quote files.")


# ════════════════════════════════════════════════════════════
# TAB 3 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab3:
    if NO_DATA:
        st.info("No catalog loaded.")
    else:
        pwc_header("Data Table",
                   "Full catalog data — filterable")

        # Top filter bar
        tf1,tf2,tf3 = st.columns(3)
        with tf1:
            all_cats_dt = (
                ["All"]
                + sorted([c for c in df_master[
                    "Category"].unique()
                           if str(c).strip()
                           not in ["","nan"]]))
            dt_cat = st.selectbox(
                "📂 Category",
                all_cats_dt,
                key="dt_cat")
        with tf2:
            vpool_dt = (
                df_master
                if dt_cat == "All"
                else df_master[
                    df_master["Category"]==dt_cat])
            all_vend_dt = (
                ["All"]
                + sorted([v for v in vpool_dt[
                    "Vendor"].unique()
                           if str(v).strip()
                           not in ["","nan"]]))
            dt_ven = st.selectbox(
                "🏢 Vendor",
                all_vend_dt,
                key="dt_ven")
        with tf3:
            dt_search = st.text_input(
                "🔍 Search file name / comments",
                placeholder="Type to search…",
                key="dt_search")

        dm = df_master.copy()
        if dt_cat  != "All":
            dm = dm[dm["Category"]==dt_cat]
        if dt_ven  != "All":
            dm = dm[dm["Vendor"]==dt_ven]
        if dt_search:
            mask = (
                dm["File Name"].str.contains(
                    dt_search, case=False,
                    na=False)
                | dm["Comments"].str.contains(
                    dt_search, case=False,
                    na=False))
            dm = dm[mask]

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            "<div style='font-size:0.82em;"
            "color:#7D7D7D;margin-bottom:8px'>"
            "Showing <b>{}</b> of <b>{}</b> "
            "records</div>".format(
                len(dm), len(df_master)),
            unsafe_allow_html=True)

        st.dataframe(
            dm.drop(
                columns=["Services List",
                          "Hyperlink"],
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
        "AI auto-detects columns and structure")

    if DATA_SOURCE == "uploaded":
        st.success(
            "✅ Using uploaded catalog: "
            "**{}** rows · **{}** vendors · "
            "**{}** services".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique()
                if not NO_DATA else 0))
    elif not NO_DATA:
        st.markdown(
            "<div class='insight-box'>"
            "📂 Currently using: "
            "<b>Master Catalog.xlsx</b> — "
            "{} vendors · {} services · "
            "{} categories</div>".format(
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique(),
                df_master["Category"].nunique()),
            unsafe_allow_html=True)

    sec("UPLOAD A NEW CATALOG")
    catalog_file = st.file_uploader(
        "Upload",
        type=["xlsx","xls","csv"],
        label_visibility="collapsed",
        key="catalog_upload")

    if catalog_file is not None:
        file_bytes = catalog_file.read()
        fname_cat  = catalog_file.name
        with st.spinner("Analysing catalog…"):
            df_new, df_exp_new, err = (
                process_uploaded_catalog(
                    file_bytes, fname_cat))
        if err:
            st.error("❌ {}".format(err))
        elif df_new is None:
            st.error("❌ Could not process file.")
        else:
            st.success(
                "✅ Detected **{}** rows · "
                "**{}** vendors · "
                "**{}** categories".format(
                    len(df_new),
                    df_new["Vendor"].nunique(),
                    df_new["Category"].nunique()))

            pc1,pc2 = st.columns(2, gap="large")
            spv_new = (
                df_exp_new.groupby("Vendor")[
                    "Service"].nunique()
                .sort_values(ascending=False)
                .reset_index())
            spv_new.columns = ["Vendor","Services"]
            with pc1:
                pf1 = go.Figure(go.Bar(
                    x=spv_new["Vendor"],
                    y=spv_new["Services"],
                    marker_color=C_ORANGE,
                    marker_line_width=0,
                    text=spv_new["Services"],
                    textposition="outside"))
                pwc_bar(pf1, "Services per Vendor",
                         height=300)
                pf1.update_xaxes(
                    tickangle=-30,
                    tickfont=dict(size=9.5))
                st.plotly_chart(
                    pf1, use_container_width=True)
            cat_new = (
                df_new.drop_duplicates(
                    subset=["Category",
                             "File Name"])
                .groupby("Category").size()
                .reset_index())
            cat_new.columns = ["Category","Count"]
            with pc2:
                if not cat_new.empty:
                    pf2 = px.pie(
                        cat_new,
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
                                size=12,
                                color=C_DARK,
                                family=
                                "Georgia,serif"),
                            x=0),
                        height=300,
                        margin=dict(
                            l=10,r=10,
                            t=40,b=10),
                        paper_bgcolor=CBG,
                        font=CFONT)
                    st.plotly_chart(
                        pf2,
                        use_container_width=True)

            st.dataframe(
                df_new.drop(
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
                    "uploaded_catalog_df"] = df_new
                st.session_state[
                    "uploaded_catalog_exp"] = (
                    df_exp_new)
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
            "Prices from GitHub quote files · "
            "Per-service benchmarking")

        GITHUB_RAW = (
            "https://raw.githubusercontent.com/"
            "avijeet528/vendor-draft-2/main/"
            "demo_quotes/{}")

        @st.cache_data(show_spinner=False)
        def load_price_from_github(filename):
            url = GITHUB_RAW.format(filename)
            try:
                resp = requests.get(url, timeout=30)
                if resp.status_code != 200:
                    return {
                        "price":"",
                        "price_num":0.0,
                        "status":"Not found"}
                ext    = filename.rsplit(".",1)[
                    -1].lower()
                result = extract_price_from_bytes(
                    resp.content, ext)
                result["status"] = (
                    "✅ Extracted"
                    if result["price_num"] > 0
                    else "⚠️ No price")
                return result
            except Exception as e:
                return {
                    "price":"","price_num":0.0,
                    "status":"❌ {}".format(
                        str(e)[:40])}

        @st.cache_data(show_spinner=False)
        def load_all_prices_from_github(
                filenames_tuple):
            return {
                fname: load_price_from_github(fname)
                for fname in filenames_tuple}

        all_fnames = tuple(
            str(f).strip()
            for f in df_master["File Name"].unique()
            if str(f).strip() not in ["","nan"])

        # Top filter bar
        va1,va2 = st.columns([2,3])
        with va1:
            run_gh = st.button(
                "🔄 Load All Prices from GitHub",
                type="primary",
                use_container_width=True,
                key="run_gh_analysis")
        with va2:
            st.markdown(
                "<div class='insight-box'>"
                "📁 Will read <b>{}</b> files from "
                "GitHub. Catalog prices used as "
                "fallback.</div>".format(
                    len(all_fnames)),
                unsafe_allow_html=True)

        # Category / vendor filter
        vf1,vf2 = st.columns(2)
        with vf1:
            all_cats_va = (
                ["All"]
                + sorted([c for c in df_master[
                    "Category"].unique()
                           if str(c).strip()
                           not in ["","nan"]]))
            va_cat = st.selectbox(
                "📂 Filter by Category",
                all_cats_va,
                key="va_cat")
        with vf2:
            va_vpool = (
                df_master
                if va_cat == "All"
                else df_master[
                    df_master["Category"]==va_cat])
            all_vend_va = (
                ["All"]
                + sorted([v for v in va_vpool[
                    "Vendor"].unique()
                           if str(v).strip()
                           not in ["","nan"]]))
            va_ven = st.selectbox(
                "🏢 Filter by Vendor",
                all_vend_va,
                key="va_ven")

        if run_gh:
            st.session_state[
                "gh_prices_loaded"] = True
            load_all_prices_from_github.clear()
            load_price_from_github.clear()

        if st.session_state.get(
                "gh_prices_loaded", False):
            with st.spinner("Loading prices…"):
                gh_prices = (
                    load_all_prices_from_github(
                        all_fnames))

            # Build analysis df
            dm_va = df_master.copy()
            if va_cat != "All":
                dm_va = dm_va[
                    dm_va["Category"]==va_cat]
            if va_ven != "All":
                dm_va = dm_va[
                    dm_va["Vendor"]==va_ven]

            rows_data = []
            for _, r in dm_va.iterrows():
                fname   = str(r.get(
                    "File Name","")).strip()
                qp      = _parse_num(str(r.get(
                    "Quoted Price","")).strip())
                gh_info = gh_prices.get(fname,{})
                gh_p    = gh_info.get(
                    "price_num",0.0)
                best_p  = gh_p if gh_p>0 else qp
                source  = (
                    "extracted" if gh_p>0
                    else "catalog" if qp>0
                    else "none")
                rows_data.append({
                    "Vendor": r["Vendor"],
                    "Category": r["Category"],
                    "File Name": fname,
                    "Quoted Price": qp,
                    "Extracted Price": gh_p,
                    "Best Price": best_p,
                    "Source": source,
                    "Status": gh_info.get(
                        "status","—"),
                    "Services": r.get(
                        "Comments","")})
            df_analysis = pd.DataFrame(rows_data)
            df_analysis = df_analysis[
                df_analysis["Best Price"]>0]

            if df_analysis.empty:
                st.warning("No prices found.")
            else:
                # Summary KPIs
                sec("PRICE EXTRACTION SUMMARY")
                k1,k2,k3,k4 = st.columns(4)
                kpi_box(k1, len(df_analysis),
                         "Files with Prices",
                         C_ORANGE)
                kpi_box(k2,
                         len(df_analysis[
                             df_analysis["Source"]
                             =="extracted"]),
                         "Extracted from Files",
                         C_DARK)
                kpi_box(k3,
                         len(df_analysis[
                             df_analysis["Source"]
                             =="catalog"]),
                         "From Catalog",
                         C_MID)
                kpi_box(k4,
                         _fmt(df_analysis[
                             "Best Price"].mean()),
                         "Avg Quote Value",
                         C_GREY_DARK)
                st.markdown("<br>",
                             unsafe_allow_html=True)

                # Vendor comparison
                vendor_totals = (
                    df_analysis.groupby("Vendor")[
                        "Best Price"]
                    .agg(["mean","sum",
                           "min","max","count"])
                    .reset_index())
                vendor_totals.columns = [
                    "Vendor","Average","Total",
                    "Min","Max","Quotes"]
                vendor_totals = vendor_totals.sort_values(
                    "Average")
                ov_avg = df_analysis["Best Price"].mean()

                sec("VENDOR COMPARISON — "
                    "AVERAGE QUOTE VALUE",
                    "Orange = cheapest · "
                    "Dark = most expensive · "
                    "Dashed = overall average")
                bar_c_va = []
                for i in range(len(vendor_totals)):
                    if i == 0:
                        bar_c_va.append(C_ORANGE)
                    elif i == len(vendor_totals)-1:
                        bar_c_va.append(C_DARK)
                    else:
                        bar_c_va.append(C_GREY_DARK)
                fig_va = go.Figure(go.Bar(
                    x=vendor_totals["Vendor"],
                    y=vendor_totals["Average"],
                    marker_color=bar_c_va,
                    marker_line_width=0,
                    text=vendor_totals[
                        "Average"].apply(_fmt),
                    textposition="outside"))
                fig_va.add_hline(
                    y=ov_avg,
                    line_dash="dash",
                    line_color=C_MID,
                    line_width=2,
                    annotation_text=
                    "Avg: {}".format(_fmt(ov_avg)),
                    annotation_position="top right")
                pwc_bar(fig_va,
                         "Average Quote per Vendor",
                         height=360)
                fig_va.update_xaxes(
                    tickangle=-20)
                st.plotly_chart(
                    fig_va,
                    use_container_width=True)

                # Ranking table
                sec("VENDOR RANKING TABLE")
                vt_tbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>Rank</th><th>Vendor</th>"
                    "<th>Quotes</th><th>Avg</th>"
                    "<th>Min</th><th>Max</th>"
                    "<th>vs Avg</th><th>Verdict</th>"
                    "</tr></thead><tbody>"]
                for rank,(_, vr) in enumerate(
                        vendor_totals.iterrows(),
                        start=1):
                    bg  = ("white" if rank%2==0
                            else "#F8F8F8")
                    vc  = vendor_color_map.get(
                        vr["Vendor"], C_DARK)
                    pct = round(
                        (vr["Average"]-ov_avg)
                        /ov_avg*100,1
                    ) if ov_avg>0 else 0
                    pc  = (C_ORANGE if pct<-5
                            else C_DARK if pct>5
                            else C_GREY_DARK)
                    pt  = (
                        "{}% below".format(abs(pct))
                        if pct<0
                        else "{}% above".format(
                            abs(pct))
                        if pct>0
                        else "At avg")
                    ov  = (
                        ("✅ COMPETITIVE",C_ORANGE)
                        if pct<-10
                        else ("🔴 EXPENSIVE",C_DARK)
                        if pct>10
                        else ("🟡 AVERAGE",
                               C_GREY_DARK))
                    medal = (
                        "🥇" if rank==1
                        else "🥈" if rank==2
                        else "🥉" if rank==3
                        else str(rank))
                    vt_tbl.append(
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
                            bg, medal,
                            vpill(vr["Vendor"],vc),
                            int(vr["Quotes"]),
                            _fmt(vr["Average"]),
                            _fmt(vr["Min"]),
                            _fmt(vr["Max"]),
                            pc,pt,ov[1],ov[0]))
                vt_tbl.append("</tbody></table>")
                st.markdown(
                    "".join(vt_tbl),
                    unsafe_allow_html=True)

                # Per-service
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("PER-SERVICE BENCHMARKING",
                    "Services quoted by multiple "
                    "vendors — full comparison")
                df_svc_rows = []
                for _, r in df_analysis.iterrows():
                    svcs_raw = str(
                        r["Services"]
                    ).replace("\\n","\n"
                              ).replace("\r\n","\n"
                                        ).replace(
                        "\r","\n")
                    svcs = [s.strip()
                             for s in svcs_raw.split(
                                 "\n")
                             if s.strip()
                             and s.strip()
                             not in ["nan","None",""]]
                    if not svcs:
                        svcs = [svcs_raw.strip()]
                    for svc in svcs:
                        df_svc_rows.append({
                            "Service": svc,
                            "Vendor":  r["Vendor"],
                            "Price":   r["Best Price"],
                            "Source":  r["Source"]})
                df_svc_df = pd.DataFrame(df_svc_rows)
                svc_vc    = (df_svc_df.groupby(
                    "Service")["Vendor"].nunique())
                multi_svcs= (svc_vc[svc_vc>1]
                              .index.tolist())

                if not multi_svcs:
                    st.info(
                        "No services with multiple "
                        "vendor quotes in current "
                        "filter.")
                else:
                    insight(
                        "<b>{}</b> service(s) have "
                        "quotes from multiple vendors "
                        "— competitive benchmarking "
                        "available.".format(
                            len(multi_svcs)))
                    for svc in sorted(multi_svcs):
                        d_s = df_svc_df[
                            df_svc_df["Service"]==svc
                        ].sort_values("Price")
                        min_p = d_s["Price"].min()
                        max_p = d_s["Price"].max()
                        avg_p = d_s["Price"].mean()
                        best_v= d_s.loc[
                            d_s["Price"].idxmin(),
                            "Vendor"]
                        worst_v=d_s.loc[
                            d_s["Price"].idxmax(),
                            "Vendor"]
                        spread= round(
                            (max_p-min_p)
                            /min_p*100,1
                        ) if min_p>0 else 0

                        st.markdown(
                            "<div style='background:"
                            "white;border-left:4px "
                            "solid {};padding:"
                            "10px 14px;border-radius:"
                            "2px;margin:10px 0;"
                            "font-weight:700;"
                            "font-size:0.88em'>"
                            "{} · {} vendors · "
                            "spread {}% · cheapest: "
                            "{} @ {}"
                            "</div>".format(
                                C_ORANGE, svc,
                                d_s["Vendor"].nunique(),
                                spread, best_v,
                                _fmt(min_p)),
                            unsafe_allow_html=True)

                        bc_s = [
                            C_ORANGE
                            if v==best_v
                            else C_DARK
                            if v==worst_v
                            else C_GREY_DARK
                            for v in d_s["Vendor"]]
                        fig_s = go.Figure(go.Bar(
                            x=d_s["Vendor"],
                            y=d_s["Price"],
                            marker_color=bc_s,
                            marker_line_width=0,
                            text=d_s["Price"].apply(
                                _fmt),
                            textposition="outside"))
                        fig_s.add_hline(
                            y=avg_p,
                            line_dash="dash",
                            line_color=C_MID,
                            line_width=1.5,
                            annotation_text=
                            "Avg: {}".format(
                                _fmt(avg_p)),
                            annotation_position=
                            "top right")
                        fig_s.update_layout(
                            height=240,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=12,b=8),
                            font=CFONT,
                            yaxis=dict(
                                showgrid=True,
                                gridcolor=C_GREY_LITE,
                                zeroline=False),
                            bargap=0.4,
                            showlegend=False)
                        st.plotly_chart(
                            fig_s,
                            use_container_width=True)

                # Download
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("DOWNLOAD ANALYSIS")
                st.download_button(
                    "📥 Download Analysis CSV",
                    data=df_analysis[[
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
                "and extracts prices automatically."
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
    st.markdown(
        "<div class='insight-box'>"
        "Your Excel must have: <b>Category</b> · "
        "<b>Vendor</b> · <b>File Name</b> · "
        "<b>Quoted Price</b> · <b>Comments</b> · "
        "<b>File Link</b> (URL to each quote file)"
        "</div>",
        unsafe_allow_html=True)

    uploaded_real = st.file_uploader(
        "Upload",
        type=["xlsx","xls","csv"],
        key="tab6_real_upload",
        label_visibility="collapsed")

    df_real_cat = None
    if uploaded_real is None:
        if not NO_DATA and not df_master.empty:
            use_real = st.checkbox(
                "Use currently loaded catalog "
                "({} files)".format(len(df_master)),
                key="use_current_real")
            if use_real:
                df_real_cat = df_master.copy()
                st.success(
                    "Using current catalog — "
                    "**{}** vendors · **{}** files"
                    .format(
                        df_real_cat["Vendor"].nunique(),
                        len(df_real_cat)))
            else:
                st.info(
                    "👆 Upload your Master Catalog "
                    "to start analysis.")
        else:
            st.info("👆 Upload your Master Catalog.")
    else:
        real_bytes = uploaded_real.read()
        df_real_cat, _, real_err = (
            process_uploaded_catalog(
                real_bytes, uploaded_real.name))
        if real_err or df_real_cat is None:
            st.error("❌ {}".format(real_err))
            df_real_cat = None
        else:
            st.success(
                "✅ **{}** vendors · **{}** files · "
                "**{}** categories".format(
                    df_real_cat["Vendor"].nunique(),
                    len(df_real_cat),
                    df_real_cat["Category"].nunique()))

    if df_real_cat is not None:
        st.markdown("<br>", unsafe_allow_html=True)
        sec("STEP 2 — RUN ANALYSIS")
        run_real = st.button(
            "🚀 Run Full Analysis",
            type="primary",
            key="run_real_analysis")

        if run_real:
            st.session_state[
                "real_analysis_done"] = False
            st.session_state[
                "real_analysis_df"] = None

        if (run_real
                or st.session_state.get(
                    "real_analysis_done", False)):
            if run_real:
                with st.spinner(
                        "Analysing {} files…".format(
                            len(df_real_cat))):

                    def _fetch_and_extract(url, fname):
                        if not url:
                            return {
                                "price_num":0.0,
                                "price":"",
                                "status":"⚪ No link"}
                        try:
                            resp = requests.get(
                                url, timeout=30,
                                headers={
                                    "User-Agent":
                                    "Mozilla/5.0"})
                            if resp.status_code != 200:
                                return {
                                    "price_num":0.0,
                                    "price":"",
                                    "status":"❌ HTTP {}".format(
                                        resp.status_code)}
                            ext = (
                                fname.rsplit(".",1)[-1]
                                .lower()
                                if "." in fname
                                else "xlsx")
                            r = extract_price_from_bytes(
                                resp.content, ext)
                            r["status"] = (
                                "✅ Extracted"
                                if r["price_num"]>0
                                else "⚠️ No price")
                            return r
                        except Exception as e:
                            return {
                                "price_num":0.0,
                                "price":"",
                                "status":"❌ {}".format(
                                    str(e)[:50])}

                    results_real = []
                    prog = st.progress(0)
                    total_r = len(df_real_cat)
                    stxt = st.empty()
                    for i,(_, row) in enumerate(
                            df_real_cat.iterrows()):
                        fname   = str(row.get(
                            "File Name","")).strip()
                        vendor  = str(row.get(
                            "Vendor","")).strip()
                        cat     = str(row.get(
                            "Category","")).strip()
                        qp      = _parse_num(str(row.get(
                            "Quoted Price","")).strip())
                        url     = ""
                        for col in [
                                "Hyperlink",
                                "File Link",
                                "File URL","URL"]:
                            val = str(row.get(
                                col,"")).strip()
                            if (val
                                    and val not in
                                    ["","nan","None"]
                                    and val.startswith(
                                        "http")):
                                url = val; break
                        stxt.markdown(
                            "<div style='font-size:"
                            "0.81em;color:#7D7D7D'>"
                            "Processing {}/{}: "
                            "<b>{}</b></div>".format(
                                i+1, total_r, fname),
                            unsafe_allow_html=True)
                        ext_r = (
                            _fetch_and_extract(
                                url, fname)
                            if url
                            else {
                                "price_num":0.0,
                                "price":"",
                                "status":"⚪ No link"})
                        ep     = ext_r.get(
                            "price_num",0.0)
                        best_p = ep if ep>0 else qp
                        source = (
                            "extracted" if ep>0
                            else "catalog" if qp>0
                            else "none")
                        results_real.append({
                            "Vendor":   vendor,
                            "Category": cat,
                            "File Name":fname,
                            "Quoted Price":qp,
                            "Extracted Price":ep,
                            "Best Price":best_p,
                            "Source":   source,
                            "Status":   ext_r.get(
                                "status","—")})
                        prog.progress((i+1)/total_r)
                    prog.empty(); stxt.empty()
                    df_result = pd.DataFrame(
                        results_real)
                    st.session_state[
                        "real_analysis_done"] = True
                    st.session_state[
                        "real_analysis_df"] = (
                        df_result.to_dict("records"))
            else:
                df_result = pd.DataFrame(
                    st.session_state[
                        "real_analysis_df"])

            df_priced = df_result[
                df_result["Best Price"]>0].copy()

            # KPIs
            k1,k2,k3,k4,k5 = st.columns(5)
            kpi_box(k1, len(df_result),
                     "Total Files", C_ORANGE)
            kpi_box(k2, len(df_priced),
                     "Prices Found", C_DARK)
            kpi_box(k3,
                     len(df_result)-len(df_priced),
                     "No Price", C_GREY_DARK)
            kpi_box(k4,
                     df_result["Vendor"].nunique(),
                     "Vendors", C_MID)
            kpi_box(k5,
                     _fmt(df_priced[
                         "Best Price"].mean())
                     if not df_priced.empty else "—",
                     "Avg Quote", C_BLACK)

            if df_priced.empty:
                st.warning(
                    "No prices found. "
                    "Check File Link URLs.")
            else:
                # Vendor ranking
                st.markdown("<br>",
                             unsafe_allow_html=True)
                sec("VENDOR RANKING — "
                    "CHEAPEST TO MOST EXPENSIVE")
                v_sum = (
                    df_priced.groupby("Vendor")[
                        "Best Price"]
                    .agg(["mean","sum","min",
                           "max","count"])
                    .reset_index())
                v_sum.columns = [
                    "Vendor","Average","Total",
                    "Min","Max","Quotes"]
                v_sum = v_sum.sort_values("Average")
                ov_avg_r = (
                    df_priced["Best Price"].mean())

                bar_c_r = [
                    C_ORANGE if i==0
                    else C_DARK
                    if i==len(v_sum)-1
                    else C_GREY_DARK
                    for i in range(len(v_sum))]
                fig_r = go.Figure(go.Bar(
                    x=v_sum["Vendor"],
                    y=v_sum["Average"],
                    marker_color=bar_c_r,
                    marker_line_width=0,
                    text=v_sum["Average"].apply(_fmt),
                    textposition="outside"))
                fig_r.add_hline(
                    y=ov_avg_r,
                    line_dash="dash",
                    line_color=C_MID,
                    line_width=2,
                    annotation_text=
                    "Avg: {}".format(
                        _fmt(ov_avg_r)),
                    annotation_position="top right")
                pwc_bar(fig_r,
                         "Average Quote per Vendor",
                         height=360)
                fig_r.update_xaxes(tickangle=-20)
                st.plotly_chart(
                    fig_r,
                    use_container_width=True)

                # Status table
                sec("FILE EXTRACTION STATUS")
                s_tbl = [
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
                for i,r in df_result.iterrows():
                    bg  = ("white" if i%2==0
                            else "#F8F8F8")
                    vc  = vendor_color_map.get(
                        r["Vendor"],C_DARK)
                    sc  = (C_ORANGE
                            if r["Source"]==
                            "extracted"
                            else C_GREY_DARK
                            if r["Source"]==
                            "catalog"
                            else C_GREY)
                    s_tbl.append(
                        "<tr style='background:{}'>"
                        "<td style='font-family:"
                        "monospace;font-size:0.77em;"
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
                        "font-size:0.73em;"
                        "font-weight:700'>"
                        "{}</span></td>"
                        "<td style='font-size:0.80em'>"
                        "{}</td>"
                        "</tr>".format(
                            bg,
                            r["File Name"],
                            vpill(r["Vendor"],vc),
                            _fmt(r["Quoted Price"])
                            if r["Quoted Price"]>0
                            else "—",
                            _fmt(r["Extracted Price"])
                            if r["Extracted Price"]>0
                            else "—",
                            _fmt(r["Best Price"])
                            if r["Best Price"]>0
                            else "—",
                            sc,
                            r["Source"].upper(),
                            r["Status"]))
                s_tbl.append("</tbody></table>")
                st.markdown(
                    "".join(s_tbl),
                    unsafe_allow_html=True)

                # Download
                st.markdown("<br>",
                             unsafe_allow_html=True)
                st.download_button(
                    "📥 Download Analysis CSV",
                    data=df_priced[[
                        "Vendor","Category",
                        "File Name",
                        "Quoted Price",
                        "Extracted Price",
                        "Best Price",
                        "Source","Status"
                    ]].to_csv(index=False),
                    file_name="real_analysis.csv",
                    mime="text/csv",
                    type="primary")
