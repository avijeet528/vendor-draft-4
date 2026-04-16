# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  Reads from master_catalog.csv (GitHub-friendly)
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
# CSS
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');
html,body,[class*="css"],div,p,span,td,th,
label,button,.stMarkdown{
    font-family:'Source Sans Pro','Helvetica Neue',
    Arial,sans-serif !important;}
.main .block-container{
    background-color:#F3F3F3 !important;
    padding-top:1.5rem;max-width:100% !important;
    padding-left:2rem !important;
    padding-right:2rem !important;}
#MainMenu{visibility:hidden;}
footer{visibility:hidden;}
header{visibility:hidden;}
[data-testid="collapsedControl"]{display:none !important;}
section[data-testid="stSidebar"]{
    background-color:#2D2D2D !important;
    border-right:3px solid #D04A02;
    min-width:300px !important;
    max-width:300px !important;}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div{
    color:#F0F0F0 !important;
    font-family:'Source Sans Pro',sans-serif !important;}
section[data-testid="stSidebar"] div[data-baseweb="select"]{
    background-color:#FFFFFF !important;
    border-radius:2px !important;border:1px solid #999 !important;}
section[data-testid="stSidebar"] div[data-baseweb="select"] *{
    color:#2D2D2D !important;}
section[data-testid="stSidebar"] span[data-baseweb="tag"]{
    background-color:#D04A02 !important;border-radius:2px !important;}
section[data-testid="stSidebar"] span[data-baseweb="tag"] span{
    color:white !important;}
.kpi-box{border-radius:4px;padding:18px 10px;
    text-align:center;color:white;
    border-left:5px solid rgba(255,255,255,0.25);}
.kpi-value{font-size:2.1em;font-weight:700;
    margin:0;line-height:1.1;}
.kpi-label{font-size:0.78em;font-weight:700;
    opacity:0.9;margin-top:5px;
    letter-spacing:0.8px;text-transform:uppercase;}
button[data-baseweb="tab"]{
    font-weight:600 !important;
    font-size:0.92em !important;
    color:#7D7D7D !important;}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;
    border-bottom:3px solid #D04A02 !important;}
div[data-testid="stExpander"] details>summary{
    list-style:none !important;padding-left:12px !important;}
div[data-testid="stExpander"] details>summary::before,
div[data-testid="stExpander"] details>summary::after,
div[data-testid="stExpander"] details>summary::-webkit-details-marker,
div[data-testid="stExpander"] details>summary::marker{
    display:none !important;content:"" !important;width:0 !important;}
div[data-testid="stExpander"] details summary p{
    font-weight:700;font-size:0.94em;
    color:#2D2D2D !important;
    padding-left:0 !important;margin-left:0 !important;}
div[data-testid="stExpander"] details{
    border:1px solid #ddd;border-radius:4px;
    margin-bottom:10px;padding:2px 0;}
.comp-table{width:100%;border-collapse:collapse;
    table-layout:fixed;font-size:0.83em;
    border:1px solid #e0e0e0;}
.comp-table thead tr{background:#2D2D2D;}
.comp-table thead th{padding:10px;text-align:left;
    font-weight:700;font-size:0.80em;
    letter-spacing:0.4px;text-transform:uppercase;
    color:white !important;border:none;word-break:break-word;}
.comp-table tbody tr:nth-child(even){background:#F3F3F3;}
.comp-table tbody tr:hover{background:#FCE8DC;}
.comp-table tbody td{padding:8px 10px;
    border-bottom:1px solid #e8e8e8;
    vertical-align:middle;word-break:break-word;
    font-size:0.82em;color:#2D2D2D;}
.comp-table th:nth-child(1),.comp-table td:nth-child(1){width:13%;}
.comp-table th:nth-child(2),.comp-table td:nth-child(2){width:12%;}
.comp-table th:nth-child(3),.comp-table td:nth-child(3){width:22%;}
.comp-table th:nth-child(4),.comp-table td:nth-child(4){width:10%;}
.comp-table th:nth-child(5),.comp-table td:nth-child(5){width:10%;}
.comp-table th:nth-child(6),.comp-table td:nth-child(6){width:10%;}
.comp-table th:nth-child(7),.comp-table td:nth-child(7){width:13%;}
.comp-table th:nth-child(8),.comp-table td:nth-child(8){width:10%;}
.vendor-badge{display:inline-block;padding:3px 8px;
    border-radius:2px;color:white;font-size:0.78em;
    font-weight:700;white-space:nowrap;overflow:hidden;
    text-overflow:ellipsis;max-width:100%;}
.score-card{border-radius:4px;padding:14px 16px;
    margin-bottom:10px;border-left:5px solid #D04A02;}
.score-card.green{background:#F0FFF4;border-color:#22992E;}
.score-card.yellow{background:#FFF8E1;border-color:#FFB600;}
.score-card.red{background:#FFF3F0;border-color:#E0301E;}
.ai-box{background:#F8F0FF;border-left:5px solid #6E2585;
    border-radius:4px;padding:14px 18px;margin:10px 0;}
.verdict-green{background:#F0FFF4;border:2px solid #22992E;
    border-radius:4px;padding:12px 16px;
    color:#22992E;font-weight:700;}
.verdict-yellow{background:#FFF8E1;border:2px solid #FFB600;
    border-radius:4px;padding:12px 16px;
    color:#856404;font-weight:700;}
.verdict-red{background:#FFF3F0;border:2px solid #E0301E;
    border-radius:4px;padding:12px 16px;
    color:#E0301E;font-weight:700;}
.demo-banner{background:linear-gradient(
    135deg,#D04A02 0%,#B83D00 100%);
    color:white;padding:10px 18px;border-radius:4px;
    margin-bottom:16px;font-size:0.88em;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# COLOURS
# ════════════════════════════════════════════════════════════
COLORS = [
    "#D04A02","#295477","#299D8F","#FFB600",
    "#22992E","#E0301E","#EB8C00","#6E2585",
    "#8C8C8C","#004F9F",
]
def get_color(i): return COLORS[i % len(COLORS)]
CFONT = dict(
    family="Source Sans Pro,Helvetica Neue,Arial",
    size=11, color="#2D2D2D")
CBG = "#F3F3F3"
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# PRICE EXTRACTION
# ════════════════════════════════════════════════════════════
PRICE_RE = re.compile(
    r"""
    (?:USD|EUR|GBP|SGD|MYR|AUD|CAD)\s?\d{1,3}(?:[,]\d{3})*(?:\.\d{1,2})?
    |(?:[\$\€\£]\s?)\d{1,3}(?:[,\s]\d{3})*(?:\.\d{1,2})?
    |\d{1,3}(?:[,]\d{3})+(?:\.\d{1,2})?
    """, re.VERBOSE|re.IGNORECASE)
TOTAL_KW = [
    "grand total","total amount","total price","amount due",
    "net total","total cost","total value","quote total",
    "subtotal","estimated total","total",
]

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
            wb = openpyxl.load_workbook(
                io.BytesIO(content),
                data_only=True, read_only=True)
            rows_text = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs = "  ".join(
                        str(c) for c in row if c is not None)
                    if rs.strip(): rows_text.append(rs)
            text = "\n".join(rows_text); wb.close()
        elif ext == "docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml = z.read("word/document.xml").decode(
                        "utf-8", errors="ignore")
                    text = re.sub(r"<[^>]+>"," ",xml)
                    text = re.sub(r"\s{2,}","\n",text)
    except: pass
    return text

def extract_price_from_bytes(content, ext):
    text  = _text_from_bytes(content, ext)
    price = _best_price(text)

    # ── FALLBACK: scan all numbers if keyword search failed ──
    if not price or _parse_num(price) <= 0:
        # Find the LARGEST number in the document
        all_nums = PRICE_RE.findall(text)
        valid    = [
            h.strip() for h in all_nums
            if _parse_num(h) >= 1000]
        if valid:
            price = max(valid, key=_parse_num)

    # ── FALLBACK 2: read Excel cells directly ──
    if (not price or _parse_num(price) <= 0) \
            and ext.lower() in ("xlsx","xls"):
        try:
            wb = openpyxl.load_workbook(
                io.BytesIO(content),
                data_only=True, read_only=True)
            all_vals = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(
                        values_only=True):
                    for cell in row:
                        if cell is None:
                            continue
                        # Check cell label for total keywords
                        cell_str = str(cell).lower()
                        if any(k in cell_str for k in
                               ["total","grand","amount",
                                "subtotal"]):
                            continue
                        # Collect numeric values
                        try:
                            n = float(str(cell).replace(
                                ",","").strip())
                            if n >= 1000:
                                all_vals.append(n)
                        except Exception:
                            pass
            wb.close()
            if all_vals:
                price = str(max(all_vals))
        except Exception:
            pass

    # ── FALLBACK 3: look for GRAND TOTAL row in Excel ──
    if (not price or _parse_num(price) <= 0) \
            and ext.lower() in ("xlsx","xls"):
        try:
            wb = openpyxl.load_workbook(
                io.BytesIO(content),
                data_only=True, read_only=True)
            for ws in wb.worksheets:
                prev_was_total = False
                for row in ws.iter_rows(
                        values_only=True):
                    row_text = " ".join(
                        str(c).lower()
                        for c in row if c is not None)
                    if any(k in row_text for k in
                           ["grand total","grand",
                            "total amount","total"]):
                        # Find numeric in this row
                        for cell in row:
                            try:
                                n = float(
                                    str(cell)
                                    .replace(",","")
                                    .strip())
                                if n >= 1000:
                                    price = str(n)
                                    break
                            except Exception:
                                pass
                    if price and _parse_num(price) > 0:
                        break
            wb.close()
        except Exception:
            pass

    return {
        "price"    : price,
        "price_num": _parse_num(price) if price else 0.0,
        "text"     : text[:5000],
    }

def extract_price_from_file(filepath):
    try:
        with open(filepath,"rb") as f:
            content = f.read()
        ext = filepath.rsplit(".",1)[-1].lower()
        return extract_price_from_bytes(content, ext)
    except:
        return {"price":"","price_num":0.0,"text":""}

# ════════════════════════════════════════════════════════════
# SCORING
# ════════════════════════════════════════════════════════════
def price_score(new_price, hist_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return None,"No comparison data",0,0,0
    mn  = min(valid); mx = max(valid)
    avg = sum(valid)/len(valid)
    if mx == mn: return 50,"Same as historical average",avg,mn,mx
    score = round((1-(new_price-mn)/(mx-mn))*100,1)
    score = max(0,min(100,score))
    pct   = round((new_price-avg)/avg*100,1)
    if new_price < avg:
        label = "{}% BELOW average — COMPETITIVE".format(abs(pct))
    elif new_price > avg:
        label = "{}% ABOVE average — REVIEW NEEDED".format(abs(pct))
    else:
        label = "Matches historical average"
    return score,label,avg,mn,mx

def score_color(s):
    if s is None: return "#8C8C8C"
    if s >= 70: return "#22992E"
    if s >= 40: return "#FFB600"
    return "#E0301E"

def score_css(s):
    if s is None: return "yellow"
    if s >= 70: return "green"
    if s >= 40: return "yellow"
    return "red"

def get_verdict(ps):
    if ps is None:
        return "⚪ No Data","No comparison data available.","#8C8C8C"
    if ps >= 70:
        return (
            "✅ COMPETITIVE",
            "This quote is priced competitively. "
            "Proceed with confidence.",
            "#22992E")
    if ps >= 40:
        return (
            "🟡 AVERAGE",
            "This quote is within average range. "
            "Negotiate for a small discount.",
            "#856404")
    return (
        "🔴 HIGH — NEGOTIATE",
        "This quote is above the historical average. "
        "Strongly recommend negotiating or seeking "
        "alternative vendors.",
        "#E0301E")

# ════════════════════════════════════════════════════════════
# AI INSIGHTS
# ════════════════════════════════════════════════════════════
def ai_service_summary(df_master, df_exploded):
    if df_master.empty: return "No vendor data."
    svc_by_v = {}
    for v in df_master["Vendor"].unique():
        svc_by_v[v] = list(
            df_exploded[df_exploded["Vendor"]==v
                       ]["Service"].unique())
    if not svc_by_v: return "No vendor data."
    best   = max(svc_by_v, key=lambda v:len(svc_by_v[v]))
    n_best = len(svc_by_v[best])
    total  = len(set(
        s for svcs in svc_by_v.values() for s in svcs))
    shared = [
        s for s in set(
            s for svcs in svc_by_v.values() for s in svcs)
        if sum(1 for svcs in svc_by_v.values()
               if s in svcs) > 1]
    lines  = ["**{}** covers the most services "
              "({} of {} total).".format(
                  best, n_best, total)]
    if shared:
        lines.append(
            "**{}** service(s) offered by multiple "
            "vendors — ideal for competitive "
            "benchmarking.".format(len(shared)))
    return " ".join(lines)

def ai_price_insight(new_price, hist_prices, vendor_prices):
    valid = [p for p in hist_prices if p > 0]
    if not valid or new_price <= 0:
        return "Insufficient data for price analysis."
    avg = sum(valid)/len(valid)
    pct = round((new_price-avg)/avg*100,1)
    lines = []
    if new_price <= min(valid):
        lines.append(
            "This quote is the **lowest price** — "
            "excellent value.")
    elif new_price >= max(valid):
        lines.append(
            "This quote is **above all historical "
            "prices** — negotiate strongly.")
    elif pct > 15:
        lines.append(
            "Quote is **{}% above** average. "
            "Request a revised quote.".format(abs(pct)))
    elif pct < -15:
        lines.append(
            "Quote is **{}% below** average — "
            "very competitive.".format(abs(pct)))
    else:
        lines.append(
            "Quote is **within normal range** "
            "({}% vs average).".format(pct))
    if vendor_prices:
        best_v = min(vendor_prices, key=vendor_prices.get)
        lines.append(
            "**{}** has historically offered the "
            "lowest prices.".format(best_v))
    return " ".join(lines)

def generate_selection_verdict(
        selected_svcs, d_sel,
        vendor_prices_map, df_exploded):
    if d_sel.empty or not selected_svcs:
        return None
    vsmap = defaultdict(set)
    for _, r in d_sel.iterrows():
        vsmap[r["Vendor"]].add(r["Service"])
    full_cover = [
        v for v,s in vsmap.items()
        if set(selected_svcs).issubset(s)]
    all_prices = [
        p for p in vendor_prices_map.values() if p > 0]
    if not all_prices:
        return {
            "title"       : "📊 Vendor Coverage",
            "vendors_all" : full_cover,
            "vendors_some": sorted([
                v for v in vsmap
                if v not in full_cover]),
            "has_prices"  : False,
            "lines"       : [],
        }
    avg_p   = sum(all_prices)/len(all_prices)
    min_p   = min(all_prices)
    max_p   = max(all_prices)
    best_v  = min(vendor_prices_map,
                  key=vendor_prices_map.get)
    worst_v = max(vendor_prices_map,
                  key=vendor_prices_map.get)
    spread  = round(
        (max_p-min_p)/min_p*100,1) if min_p>0 else 0
    lines = [
        "**{}** vendor(s) quoted for the selected "
        "service(s).".format(len(vendor_prices_map)),
        "Price range: **{}** — **{}** "
        "(spread: **{}%**).".format(
            _fmt(min_p),_fmt(max_p),spread),
        "Average quoted price: **{}**.".format(
            _fmt(avg_p)),
    ]
    if full_cover:
        lines.append(
            "**{}** offer(s) ALL selected services.".format(
                ", ".join(full_cover)))
    else:
        lines.append(
            "No single vendor covers all selected "
            "services — consider multi-vendor approach.")
    lines.append(
        "**Best price:** {} at {} — "
        "{}% below average.".format(
            best_v, _fmt(min_p),
            round((avg_p-min_p)/avg_p*100,1)
            if avg_p > 0 else 0))
    if spread > 20:
        lines.append(
            "⚠️ Large price spread ({}%) — "
            "significant negotiation opportunity.".format(
                spread))
    return {
        "title"        : "📊 Procurement Verdict",
        "vendors_all"  : full_cover,
        "vendors_some" : sorted([
            v for v in vsmap
            if v not in full_cover]),
        "has_prices"   : True,
        "lines"        : lines,
        "best_vendor"  : best_v,
        "best_price"   : min_p,
        "worst_vendor" : worst_v,
        "worst_price"  : max_p,
        "avg_price"    : avg_p,
        "spread"       : spread,
        "vendor_prices": vendor_prices_map,
    }

# ════════════════════════════════════════════════════════════
# DATA LOADING — reads CSV (GitHub-friendly)
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_data():
    # ── Try CSV first (most reliable) ──
    CSV_PATH = "master_catalog.csv"
    XLS_PATH = "Master Catalog.xlsx"

    df = None

    if os.path.exists(CSV_PATH):
        try:
            df = pd.read_csv(CSV_PATH)
            df.columns = [
                str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("CSV load error: {}".format(e))

    elif os.path.exists(XLS_PATH):
        try:
            raw = pd.read_excel(
                XLS_PATH, engine="openpyxl",
                header=None)
            header_row = 0
            for i, row in raw.iterrows():
                vals = [
                    str(v).strip().lower()
                    for v in row.values if pd.notna(v)]
                if (any("category" in v for v in vals)
                        and any("vendor" in v
                                for v in vals)):
                    header_row = i; break
            df = pd.read_excel(
                XLS_PATH, engine="openpyxl",
                header=header_row)
            df.columns = [
                str(c).strip() for c in df.columns]
        except Exception as e:
            st.warning("Excel load error: {}".format(e))

    if df is None:
        return None, None

    # ── Normalise column names ──
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == "category":
            col_map["Category"] = c
        elif any(k in cl for k in
                 ["vendor","supplier"]):
            col_map["Vendor"] = c
        elif "file name" in cl or cl == "filename":
            col_map["File Name"] = c
        elif any(k in cl for k in
                 ["file link","file url","url","link"]):
            col_map["File Link"] = c
        elif any(k in cl for k in
                 ["comment","service","description",
                  "scope"]):
            col_map["Comments"] = c
        elif any(k in cl for k in
                 ["price","cost","amount","quoted"]):
            col_map["Quoted Price"] = c

    df.rename(
        columns={v:k for k,v in col_map.items()},
        inplace=True)

    # Ensure all required columns exist
    for req in ["Category","Vendor","File Name",
                "Comments"]:
        if req not in df.columns:
            df[req] = ""

    keep = ["Category","Vendor","File Name","Comments"]
    for e in ["File Link","Quoted Price"]:
        if e in df.columns: keep.append(e)
    df = df[[c for c in keep
             if c in df.columns]].copy()

    # Clean
    df = df[~(
        df["Category"].astype(str).str.strip()
        .isin(["","nan"]) &
        df["Vendor"].astype(str).str.strip()
        .isin(["","nan"]))].copy()

    for col in df.columns:
        df[col] = (df[col].fillna("")
                   .astype(str).str.strip())
    df.reset_index(drop=True, inplace=True)

    # File link column
    df["Hyperlink"] = ""
    if "File Link" in df.columns:
        df["Hyperlink"] = df["File Link"].apply(
            lambda x: "" if x in ["","nan"] else x)

    # ── Parse services — handles ALL newline types ──
    def parse_svc(v):
        if not v or str(v).strip() in ["","nan","None"]:
            return ["(unspecified)"]
        s = str(v)
        # Normalise all newline variants
        s = s.replace("\\n", "\n")
        s = s.replace("\r\n", "\n")
        s = s.replace("\r", "\n")
        parts = [
            p.strip() for p in s.split("\n")
            if p.strip() and p.strip() != "nan"]
        if not parts:
            # Try semicolon
            parts = [
                p.strip() for p in s.split(";")
                if p.strip()]
        if not parts:
            # Try comma (short values only)
            if len(s) < 300:
                parts = [
                    p.strip() for p in s.split(",")
                    if p.strip()]
        return parts if parts else ["(unspecified)"]

    df["Services List"] = df["Comments"].apply(parse_svc)

    # Explode
    df_exp = df.explode("Services List").copy()
    df_exp.rename(
        columns={"Services List":"Service"},
        inplace=True)
    df_exp["Service"] = (
        df_exp["Service"].astype(str).str.strip())
    df_exp = df_exp[
        ~df_exp["Service"].isin(
            ["","(unspecified)","nan","None"])
    ].reset_index(drop=True)

    return df, df_exp

# ════════════════════════════════════════════════════════════
# REAL CATALOG ANALYZER
# Reads Master Catalog.xlsx → follows File Links
# → extracts prices → builds full analysis
# ════════════════════════════════════════════════════════════

@st.cache_data(ttl=600, show_spinner=False)
def fetch_file_and_extract(url, filename):
    """
    Downloads a file from any URL and extracts price.
    Supports SharePoint, OneDrive, Google Drive,
    direct HTTP links, GitHub raw links.
    """
    if not url or str(url).strip() in [
            "","nan","None"]:
        return {
            "price_num": 0.0,
            "price"    : "",
            "status"   : "⚪ No link",
        }

    url = str(url).strip()

    # Fix Google Drive links
    if "drive.google.com" in url:
        match = re.search(
            r"/d/([a-zA-Z0-9_-]+)", url)
        if match:
            url = (
                "https://drive.google.com/"
                "uc?export=download&id={}"
                .format(match.group(1)))

    # Fix OneDrive links
    if "1drv.ms" in url or "onedrive" in url:
        if not url.endswith("download"):
            url = url.rstrip("/") + \
                  "?download=1"

    try:
        headers = {
            "User-Agent":
                "Mozilla/5.0 (Windows NT 10.0;"
                " Win64; x64) AppleWebKit/537.36"
        }
        resp = requests.get(
            url, timeout=30,
            headers=headers,
            allow_redirects=True)

        if resp.status_code != 200:
            return {
                "price_num": 0.0,
                "price"    : "",
                "status"   : "❌ HTTP {}".format(
                    resp.status_code),
            }

        # Detect file type from URL or content
        ext = filename.rsplit(
            ".",1)[-1].lower() if "." in filename \
            else url.split("?")[0].rsplit(
                ".",1)[-1].lower()
        if ext not in (
                "xlsx","xls","pdf","docx","csv"):
            ext = "xlsx"  # default

        result = extract_price_from_bytes(
            resp.content, ext)
        result["status"] = (
            "✅ Extracted"
            if result["price_num"] > 0
            else "⚠️ No price found in file")
        return result

    except requests.exceptions.Timeout:
        return {
            "price_num": 0.0,
            "price"    : "",
            "status"   : "❌ Timeout",
        }
    except Exception as e:
        return {
            "price_num": 0.0,
            "price"    : "",
            "status"   : "❌ {}".format(
                str(e)[:60]),
        }


def get_file_link_from_row(row):
    """
    Extracts file URL from a catalog row.
    Checks multiple possible column names.
    """
    for col in [
        "File Link","File URL","URL",
        "Hyperlink","Link","file_link",
        "file link","hyperlink"
    ]:
        val = str(row.get(col,"")).strip()
        if val and val not in [
                "","nan","None","#N/A"]:
            return val
    return ""


def analyze_real_catalog(df_catalog):
    """
    Takes the master catalog DataFrame,
    follows all File Links,
    extracts prices,
    returns enriched analysis DataFrame.
    """
    results = []
    total = len(df_catalog)

    prog_bar = st.progress(0)
    status_txt = st.empty()

    for i, (_, row) in enumerate(
            df_catalog.iterrows()):

        fname   = str(
            row.get("File Name","")).strip()
        vendor  = str(
            row.get("Vendor","")).strip()
        cat     = str(
            row.get("Category","")).strip()
        comments= str(
            row.get("Comments","")).strip()
        qp      = _parse_num(str(
            row.get("Quoted Price","")).strip())
        url     = get_file_link_from_row(row)

        status_txt.markdown(
            "<div style='font-size:0.82em;"
            "color:#555'>"
            "Processing {}/{}: "
            "<b>{}</b></div>".format(
                i+1, total, fname),
            unsafe_allow_html=True)

        # Extract price from linked file
        extracted = {}
        if url:
            extracted = fetch_file_and_extract(
                url, fname)
        else:
            extracted = {
                "price_num": 0.0,
                "price"    : "",
                "status"   : "⚪ No link provided",
            }

        ex_p   = extracted.get("price_num", 0.0)
        status = extracted.get("status", "—")

        # Best price: extracted > quoted price
        best_p = ex_p if ex_p > 0 else qp
        source = (
            "extracted" if ex_p > 0
            else "catalog"  if qp  > 0
            else "none")

        # Parse services
        svcs = parse_services_from_cell(comments)

        results.append({
            "Vendor"          : vendor,
            "Category"        : cat,
            "File Name"       : fname,
            "File Link"       : url,
            "Quoted Price"    : qp,
            "Extracted Price" : ex_p,
            "Best Price"      : best_p,
            "Source"          : source,
            "Status"          : status,
            "Services"        : svcs,
        })

        prog_bar.progress((i+1)/total)

    prog_bar.empty()
    status_txt.empty()

    return pd.DataFrame(results)
# ════════════════════════════════════════════════════════════
# PROCESS UPLOADED CATALOG
# ════════════════════════════════════════════════════════════
def process_uploaded_catalog(file_bytes, filename):
    try:
        ext = filename.rsplit(".",1)[-1].lower()
        if ext in ("xlsx","xls"):
            raw = pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl", header=None)
        elif ext == "csv":
            raw = pd.read_csv(
                io.BytesIO(file_bytes), header=None)
        else:
            return None,None,"Unsupported file type."

        header_row = 0
        for i, row in raw.iterrows():
            vals = [
                str(v).strip().lower()
                for v in row.values if pd.notna(v)]
            joined = " ".join(vals)
            if (any(k in joined for k in
                    ["vendor","supplier"]) and
                    any(k in joined for k in
                        ["file","document"])):
                header_row = i; break

        if ext in ("xlsx","xls"):
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                engine="openpyxl",
                header=header_row)
        else:
            df = pd.read_csv(
                io.BytesIO(file_bytes),
                header=header_row)

        df = df.loc[:, df.columns.notna()]
        df.columns = [str(c).strip() for c in df.columns]
        df.dropna(how="all", inplace=True)

        col_map = {}
        for c in df.columns:
            cl = str(c).lower().strip()
            if any(k in cl for k in
                   ["category","type","domain"]):
                if "Category" not in col_map:
                    col_map["Category"] = c
            elif any(k in cl for k in
                     ["vendor","supplier","company"]):
                if "Vendor" not in col_map:
                    col_map["Vendor"] = c
            elif any(k in cl for k in
                     ["file name","filename"]):
                if "File Name" not in col_map:
                    col_map["File Name"] = c
            elif any(k in cl for k in
                     ["link","url","hyperlink"]):
                if "File Link" not in col_map:
                    col_map["File Link"] = c
            elif any(k in cl for k in
                     ["comment","service",
                      "description","scope"]):
                if "Comments" not in col_map:
                    col_map["Comments"] = c
            elif any(k in cl for k in
                     ["price","cost","amount"]):
                if "Quoted Price" not in col_map:
                    col_map["Quoted Price"] = c

        df.rename(
            columns={v:k for k,v in col_map.items()},
            inplace=True)
        for req in ["Category","Vendor",
                    "File Name","Comments"]:
            if req not in df.columns: df[req] = ""

        keep = ["Category","Vendor","File Name",
                "Comments"]
        for e in ["File Link","Quoted Price"]:
            if e in df.columns: keep.append(e)
        df = df[[c for c in keep
                 if c in df.columns]].copy()
        df = df[~(
            df["Category"].astype(str).str.strip()
            .isin(["","nan"]) &
            df["Vendor"].astype(str).str.strip()
            .isin(["","nan"]))].copy()
        for col in df.columns:
            df[col] = (df[col].fillna("")
                       .astype(str).str.strip())
        df.reset_index(drop=True, inplace=True)
        df["Hyperlink"] = ""

        def parse_svc(v):
            if not v or str(v).strip() in [
                    "","nan","None"]:
                return ["(unspecified)"]
            s = str(v).replace("\\n","\n")
            s = s.replace("\r\n","\n").replace("\r","\n")
            parts = [
                p.strip() for p in s.split("\n")
                if p.strip()]
            if not parts:
                parts = [
                    p.strip() for p in s.split(";")
                    if p.strip()]
            return parts if parts else ["(unspecified)"]

        df["Services List"] = df["Comments"].apply(
            parse_svc)
        df_exp = df.explode("Services List").copy()
        df_exp.rename(
            columns={"Services List":"Service"},
            inplace=True)
        df_exp["Service"] = (
            df_exp["Service"].astype(str).str.strip())
        df_exp = df_exp[
            ~df_exp["Service"].isin(
                ["","(unspecified)","nan","None"])
        ].reset_index(drop=True)

        return df, df_exp, None
    except Exception as e:
        return None, None, str(e)

# ════════════════════════════════════════════════════════════
# LOAD DATA
# ════════════════════════════════════════════════════════════
if ("uploaded_catalog_df" in st.session_state and
        st.session_state["uploaded_catalog_df"]
        is not None):
    df_master   = st.session_state["uploaded_catalog_df"]
    df_exploded = st.session_state["uploaded_catalog_exp"]
    DATA_SOURCE = "uploaded"
else:
    df_master, df_exploded = load_data()
    DATA_SOURCE = "file"

NO_DATA = (df_master is None or
           df_exploded is None or
           df_master.empty)

if not NO_DATA:
    vendor_color_map = {
        v: get_color(i)
        for i,v in enumerate(
            sorted(df_master["Vendor"].unique()))}
else:
    vendor_color_map = {}

# ════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════
def sb_label(txt):
    st.markdown(
        "<p style='color:#F0F0F0;font-weight:700;"
        "font-size:0.85em;margin:12px 0 4px;"
        "letter-spacing:0.5px;text-transform:uppercase'>"
        "{}</p>".format(txt),
        unsafe_allow_html=True)

def section_title(txt, caption=""):
    st.markdown(
        "<div style='font-size:0.78em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:4px'>"
        "{}</div>".format(txt),
        unsafe_allow_html=True)
    if caption: st.caption(caption)

def vendor_pill(v, color):
    return ("<span class='vendor-badge' "
            "style='background:{}'>{}</span>".format(
                color, v))

def ai_box(content):
    st.markdown(
        "<div class='ai-box'>"
        "<div style='font-size:0.68em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#6E2585;margin-bottom:6px'>"
        "AI Insight</div>{}</div>".format(content),
        unsafe_allow_html=True)

def resolve_url(row):
    """Check local demo_quotes folder first."""
    fname = str(row.get("File Name","")).strip()
    if fname:
        local = os.path.join(DEMO_DIR, fname)
        if os.path.exists(local):
            return local
    for col in ["Hyperlink","File Link"]:
        val = str(row.get(col,"")).strip()
        if val and val not in ["","nan"]:
            if val.startswith("http"):
                return val
            # Relative path
            base = os.path.basename(
                val.replace("\\","/"))
            local2 = os.path.join(DEMO_DIR, base)
            if os.path.exists(local2):
                return local2
    return ""

def kpi(col, val, lbl, bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div>"
        "</div>".format(bg, val, lbl),
        unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
selected_svcs   = []
selected_cat    = "All"
selected_vendor = "All"
d_filt          = pd.DataFrame()

if not NO_DATA:
    with st.sidebar:
        st.markdown(
            "<div style='text-align:center;"
            "padding:20px 0 14px'>"
            "<div style='font-size:2em'>📋</div>"
            "<div style='font-size:1.05em;"
            "font-weight:700;color:white;"
            "margin:5px 0 2px'>IT Procurement</div>"
            "<div style='font-size:0.70em;color:#aaa;"
            "letter-spacing:1px;"
            "text-transform:uppercase'>"
            "Intelligence Dashboard</div></div>"
            "<hr style='border-color:#D04A02;"
            "border-width:2px;margin:0 0 14px'>",
            unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown(
                "<div style='background:#D04A02;"
                "color:white;padding:6px 10px;"
                "border-radius:2px;font-size:0.75em;"
                "font-weight:700;text-align:center;"
                "margin-bottom:10px'>"
                "📤 USING UPLOADED CATALOG</div>",
                unsafe_allow_html=True)

        sb_label("📂 Category")
        all_cats = ["All"] + sorted([
            c for c in df_master["Category"].unique()
            if str(c).strip() not in ["","nan"]])
        selected_cat = st.selectbox(
            "Category", all_cats,
            label_visibility="collapsed")

        sb_label("🏢 Vendor")
        vpool = (df_master if selected_cat == "All"
                 else df_master[
                     df_master["Category"]==selected_cat])
        all_vendors = ["All"] + sorted([
            v for v in vpool["Vendor"].unique()
            if str(v).strip() not in ["","nan"]])
        selected_vendor = st.selectbox(
            "Vendor", all_vendors,
            label_visibility="collapsed")

        st.markdown(
            "<hr style='border-color:#555;"
            "margin:12px 0'>",
            unsafe_allow_html=True)

        d_filt = df_exploded.copy()
        if selected_cat != "All":
            d_filt = d_filt[
                d_filt["Category"]==selected_cat]
        if selected_vendor != "All":
            d_filt = d_filt[
                d_filt["Vendor"]==selected_vendor]

        sb_label("🔍 Search Services")
        svc_search = st.text_input(
            "Search",
            placeholder="e.g. Cisco, Azure…",
            label_visibility="collapsed")

        avail = sorted([
            s for s in d_filt["Service"].unique()
            if str(s).strip() not in ["","nan"]])
        if svc_search:
            avail = [
                s for s in avail
                if svc_search.lower() in s.lower()]

        sb_label(
            "🛠 Select Services ({})".format(
                len(avail)))
        selected_svcs = st.multiselect(
            "Services",
            options=avail,
            default=[],
            label_visibility="collapsed",
            help="Select services to compare vendors")

        st.markdown(
            "<hr style='border-color:#555;"
            "margin:12px 0'>",
            unsafe_allow_html=True)

        st.markdown(
            "<p style='color:#888;font-size:0.78em;"
            "margin:2px 0'>"
            "📄 {} quotes | 🛠 {} services | "
            "🏢 {} vendors</p>".format(
                len(df_master),
                df_exploded["Service"].nunique(),
                df_master["Vendor"].nunique()),
            unsafe_allow_html=True)

        if selected_svcs:
            st.markdown(
                "<p style='color:#D04A02;"
                "font-size:0.80em;font-weight:700;"
                "margin:4px 0'>✅ {} service(s) "
                "selected</p>".format(
                    len(selected_svcs)),
                unsafe_allow_html=True)

        if DATA_SOURCE == "uploaded":
            st.markdown(
                "<hr style='border-color:#555;"
                "margin:12px 0'>",
                unsafe_allow_html=True)
            if st.button(
                    "🔄 Reset to Default Catalog",
                    use_container_width=True):
                st.session_state[
                    "uploaded_catalog_df"] = None
                st.session_state[
                    "uploaded_catalog_exp"] = None
                st.rerun()

# ════════════════════════════════════════════════════════════
# MAIN HEADER
# ════════════════════════════════════════════════════════════
st.markdown(
    "<div style='background:#2D2D2D;color:white;"
    "padding:20px 28px;border-radius:4px;"
    "border-left:6px solid #D04A02;"
    "margin-bottom:22px'>"
    "<div style='font-size:0.72em;font-weight:700;"
    "letter-spacing:2px;text-transform:uppercase;"
    "color:#D04A02;margin-bottom:5px'>"
    "IT Procurement Analytics</div>"
    "<h1 style='margin:0;font-size:1.45em;"
    "font-weight:700;color:white'>"
    "Procurement Intelligence Dashboard</h1>"
    "<p style='margin:6px 0 0;opacity:0.6;"
    "font-size:0.85em'>"
    "Browse quotations · Compare prices · "
    "Upload &amp; score · AI insights · Verdict system"
    "</p></div>",
    unsafe_allow_html=True)

if os.path.exists(DEMO_DIR):
    st.markdown(
        "<div class='demo-banner'>"
        "🎯 <b>DEMO MODE</b> — Sample data loaded. "
        "Select <b>Cisco Catalyst C9300</b> from the "
        "sidebar → see verdict. "
        "Upload <b>DEMO_New_Quotation_AlphaNetworks.xlsx"
        "</b> in Upload tab to score a new quote."
        "</div>",
        unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# KPI ROW
# ════════════════════════════════════════════════════════════
if not NO_DATA:
    use_filt = (d_filt if not d_filt.empty
                else df_exploded)
    k1,k2,k3,k4 = st.columns(4)
    kpi(k1,
        use_filt["File Name"].nunique()
        if not use_filt.empty
        else df_master["File Name"].nunique(),
        "Total Quotes","#D04A02")
    kpi(k2,
        use_filt["Service"].nunique()
        if not use_filt.empty
        else df_exploded["Service"].nunique(),
        "Unique Services","#295477")
    kpi(k3,
        use_filt["Vendor"].nunique()
        if not use_filt.empty
        else df_master["Vendor"].nunique(),
        "Vendors","#299D8F")
    kpi(k4,
        use_filt["Category"].nunique()
        if not use_filt.empty
        else df_master["Category"].nunique(),
        "Categories","#2D2D2D")
    st.markdown("<br>", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS
# ════════════════════════════════════════════════════════════
tab1,tab2,tab3,tab4,tab5,tab6,tab7 = st.tabs([
    "📊 Analytics",
    "📋 Browse & Verdict",
    "📤 Upload & Score",
    "📄 Data Table",
    "🗂 Upload Catalog",
    "🔍 Vendor Analysis",
    "📂 Real Analysis",      
])
# ════════════════════════════════════════════════════════════
# GITHUB FILE LOADER
# ════════════════════════════════════════════════════════════
GITHUB_RAW = (
    "https://raw.githubusercontent.com/"
    "avijeet528/vendor-draft-2/main/demo_quotes/{}")

@st.cache_data(show_spinner=False)
def load_price_from_github(filename):
    """
    Downloads a quote file directly from GitHub
    and extracts price. Cached so only runs once.
    """
    url = GITHUB_RAW.format(filename)
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code != 200:
            return {
                "price"    : "",
                "price_num": 0.0,
                "status"   : "Not found ({})"
                             .format(resp.status_code),
            }
        ext = filename.rsplit(".",1)[-1].lower()
        result = extract_price_from_bytes(
            resp.content, ext)
        result["status"] = (
            "✅ Extracted"
            if result["price_num"] > 0
            else "⚠️ No price found")
        return result
    except Exception as e:
        return {
            "price"    : "",
            "price_num": 0.0,
            "status"   : "❌ Error: {}".format(str(e)),
        }


@st.cache_data(show_spinner=False)
def load_all_prices_from_github(filenames_tuple):
    """
    Loads prices for all files from GitHub.
    Takes a tuple (hashable) for caching.
    """
    results = {}
    for fname in filenames_tuple:
        results[fname] = load_price_from_github(fname)
    return results


def get_price_for_row(row, gh_prices):
    """
    Returns best available price for a catalog row.
    Priority: GitHub extracted > Quoted Price in CSV
    """
    fname  = str(row.get("File Name","")).strip()
    qp     = _parse_num(
        str(row.get("Quoted Price","")).strip())
    gh     = gh_prices.get(fname,{})
    gh_p   = gh.get("price_num", 0.0)
    if gh_p > 0:
        return gh_p, "extracted"
    if qp > 0:
        return qp,  "catalog"
    return 0.0, "none"
# ════════════════════════════════════════════════════════════
# TAB 1 — ANALYTICS
# ════════════════════════════════════════════════════════════
with tab1:
    if NO_DATA:
        st.info(
            "No catalog loaded. "
            "Go to 🗂 Upload Catalog tab.")
    else:
        use_df = (d_filt if not d_filt.empty
                  else df_exploded)
        col_l, col_r = st.columns(2, gap="large")

        with col_l:
            section_title(
                "SERVICE OVERLAP ANALYSIS",
                "Orange = quoted by multiple vendors.")
            shared = (
                use_df.groupby("Service")[
                    "Vendor"].nunique()
                .sort_values(ascending=False)
                .head(20).reset_index())
            shared.columns = ["Service","Vendor Count"]
            shared["Color"] = shared[
                "Vendor Count"].apply(
                lambda x:"#D04A02" if x>1
                else "#C0C0C0")
            fig1 = go.Figure(go.Bar(
                x=shared["Vendor Count"],
                y=shared["Service"].str[:44],
                orientation="h",
                marker_color=shared["Color"],
                marker_line_width=0,
                text=shared["Vendor Count"],
                textposition="outside",
                textfont=dict(size=10)))
            fig1.update_layout(
                height=480,
                plot_bgcolor=CBG,
                paper_bgcolor=CBG,
                margin=dict(l=5,r=40,t=20,b=10),
                font=CFONT,
                xaxis=dict(
                    title="Vendors",showgrid=True,
                    gridcolor="#E0E0E0",
                    zeroline=False),
                yaxis=dict(
                    autorange="reversed",
                    tickfont=dict(size=9.5)),
                bargap=0.35)
            st.plotly_chart(
                fig1, use_container_width=True)

        with col_r:
            section_title(
                "VENDOR SERVICE COVERAGE",
                "Higher = broader capability.")
            spv = (
                use_df.groupby("Vendor")[
                    "Service"].nunique()
                .sort_values(ascending=False)
                .reset_index())
            spv.columns = ["Vendor","Count"]
            spv["Color"] = [
                vendor_color_map.get(v,"#8C8C8C")
                for v in spv["Vendor"]]
            fig2 = go.Figure(go.Bar(
                x=spv["Vendor"],
                y=spv["Count"],
                marker_color=spv["Color"],
                marker_line_width=0,
                text=spv["Count"],
                textposition="outside",
                textfont=dict(size=10)))
            fig2.update_layout(
                height=480,
                plot_bgcolor=CBG,
                paper_bgcolor=CBG,
                margin=dict(l=5,r=10,t=20,b=10),
                font=CFONT,
                yaxis=dict(
                    title="Unique Services",
                    showgrid=True,
                    gridcolor="#E0E0E0",
                    zeroline=False),
                xaxis=dict(
                    tickangle=-35,
                    tickfont=dict(size=9.5)),
                bargap=0.35)
            st.plotly_chart(
                fig2, use_container_width=True)

        section_title(
            "CATEGORY DISTRIBUTION",
            "Share of quote files across categories.")
        cat_c = (
            use_df.drop_duplicates(
                subset=["Category","File Name"])
            .groupby("Category").size()
            .reset_index())
        cat_c.columns = ["Category","Count"]
        if not cat_c.empty:
            fig3 = px.pie(
                cat_c,names="Category",
                values="Count",hole=0.50,
                color_discrete_sequence=COLORS)
            fig3.update_traces(
                textposition="outside",
                textinfo="label+percent",
                textfont_size=11,
                pull=[0.03]*len(cat_c))
            fig3.update_layout(
                height=380,
                margin=dict(l=20,r=20,t=20,b=20),
                paper_bgcolor=CBG,font=CFONT,
                legend=dict(
                    orientation="v",
                    x=1.02,y=0.5,
                    font=dict(size=10)))
            st.plotly_chart(
                fig3, use_container_width=True)

        st.markdown("<br>",unsafe_allow_html=True)
        section_title("AI VENDOR SUMMARY")
        ai_box(ai_service_summary(
            df_master, df_exploded))

# ════════════════════════════════════════════════════════════
# TAB 2 — BROWSE & VERDICT
# ════════════════════════════════════════════════════════════
with tab2:
    if not NO_DATA:
        with st.expander(
                "🔧 Debug — click to verify data",
                expanded=False):
            st.write("**Vendors:**",
                df_master["Vendor"].unique().tolist())
            st.write("**Services sample:**",
                df_exploded["Service"].head(15).tolist())
            st.write("**Has Quoted Price column:**",
                "Quoted Price" in df_master.columns)
            if "Quoted Price" in df_master.columns:
                st.write("**Sample prices:**",
                    df_master["Quoted Price"].head(10).tolist())
    # ── END DEBUG ──
    if NO_DATA:
        st.info(
            "No catalog loaded. "
            "Go to 🗂 Upload Catalog tab.")
    elif not selected_svcs:
        st.info(
            "👈 Select services from the sidebar "
            "to browse quotations and see the verdict.")
        if os.path.exists(DEMO_DIR):
            st.markdown(
                "<div style='background:white;"
                "border:1px solid #e0e0e0;"
                "border-radius:4px;"
                "padding:16px 20px;margin-top:12px'>"
                "<div style='font-size:0.72em;"
                "font-weight:700;letter-spacing:1px;"
                "text-transform:uppercase;"
                "color:#D04A02;margin-bottom:8px'>"
                "💡 Try These Demo Services</div>"
                "<p style='font-size:0.87em;"
                "color:#555;margin-bottom:6px'>"
                "Select any of these from the sidebar:"
                "</p>"
                "<p style='font-size:0.87em;"
                "color:#2D2D2D;margin:0'>"
                "• <b>Cisco Catalyst C9300</b> "
                "→ 4 vendors quoted<br>"
                "• <b>Palo Alto Firewall</b> "
                "→ 2 vendors quoted<br>"
                "• <b>Azure Virtual Machines</b> "
                "→ 2 vendors quoted<br>"
                "• <b>Dell Latitude 5540</b> "
                "→ 2 vendors quoted"
                "</p></div>",
                unsafe_allow_html=True)
    else:
        use_filt2 = (d_filt if not d_filt.empty
                     else df_exploded)
        d_sel = use_filt2[
            use_filt2["Service"].isin(
                selected_svcs)].copy()

        if d_sel.empty:
            st.warning("No results found.")
        else:
            has_price = "Quoted Price" in d_sel.columns

            # Collect vendor prices
            # KEY FIX: always use Quoted Price from catalog first
            vendor_prices_map = {}
            for _, r in d_sel.drop_duplicates(
                    subset=["Vendor","File Name"]).iterrows():
            
                v  = r["Vendor"]
                qp = _parse_num(
                    str(r.get("Quoted Price","")).strip())
                ck = "px_{}".format(
                    str(r.get("File Name","")).strip())
                ca = st.session_state.get(ck)
                ep = ca["price_num"] if ca else 0.0
            
                # Priority: extracted price > quoted price
                ref = ep if ep > 0 else qp
            
                if ref > 0:
                    if v not in vendor_prices_map:
                        vendor_prices_map[v] = ref
                    else:
                        # Keep lowest price per vendor
                        vendor_prices_map[v] = min(
                            vendor_prices_map[v], ref)
            
            # ── If still no prices, use Quoted Price directly ──
            if not vendor_prices_map:
                for _, r in d_sel.drop_duplicates(
                        subset=["Vendor"]).iterrows():
                    v  = r["Vendor"]
                    qp = _parse_num(
                        str(r.get("Quoted Price","")).strip())
                    if qp > 0 and v not in vendor_prices_map:
                        vendor_prices_map[v] = qp

            verdict = generate_selection_verdict(
                selected_svcs, d_sel,
                vendor_prices_map, df_exploded)

            # ── VERDICT BANNER ──────────────────────
            if verdict and verdict["has_prices"]:
                st.markdown(
                    "<div style='font-size:0.78em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#D04A02;margin-bottom:10px'>"
                    "PROCUREMENT VERDICT</div>",
                    unsafe_allow_html=True)

                best_p  = verdict["best_price"]
                worst_p = verdict["worst_price"]
                avg_p   = verdict["avg_price"]
                spread  = verdict["spread"]
                best_v  = verdict["best_vendor"]

                vc1,vc2,vc3,vc4 = st.columns(4)
                vc1.markdown(
                    "<div class='score-card green'>"
                    "<div style='font-size:0.68em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#22992E'>Best Price</div>"
                    "<div style='font-size:1.6em;"
                    "font-weight:800;color:#22992E'>"
                    "{}</div>"
                    "<div style='font-size:0.75em;"
                    "color:#555;margin-top:2px'>"
                    "by {}</div>"
                    "</div>".format(
                        _fmt(best_p), best_v),
                    unsafe_allow_html=True)

                vc2.markdown(
                    "<div class='score-card yellow'>"
                    "<div style='font-size:0.68em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#856404'>Avg Price</div>"
                    "<div style='font-size:1.6em;"
                    "font-weight:800;color:#856404'>"
                    "{}</div>"
                    "<div style='font-size:0.75em;"
                    "color:#555;margin-top:2px'>"
                    "{} vendors</div>"
                    "</div>".format(
                        _fmt(avg_p),
                        len(vendor_prices_map)),
                    unsafe_allow_html=True)

                vc3.markdown(
                    "<div class='score-card red'>"
                    "<div style='font-size:0.68em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#E0301E'>Highest Price</div>"
                    "<div style='font-size:1.6em;"
                    "font-weight:800;color:#E0301E'>"
                    "{}</div>"
                    "<div style='font-size:0.75em;"
                    "color:#555;margin-top:2px'>"
                    "by {}</div>"
                    "</div>".format(
                        _fmt(worst_p),
                        verdict["worst_vendor"]),
                    unsafe_allow_html=True)

                vc4.markdown(
                    "<div class='score-card {}'>"
                    "<div style='font-size:0.68em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:{}'>Price Spread</div>"
                    "<div style='font-size:1.6em;"
                    "font-weight:800;color:{}'>"
                    "{}%</div>"
                    "<div style='font-size:0.75em;"
                    "color:#555;margin-top:2px'>"
                    "negotiation room</div>"
                    "</div>".format(
                        "red" if spread>20
                        else "yellow" if spread>10
                        else "green",
                        score_color(
                            100-min(spread*2,100)),
                        score_color(
                            100-min(spread*2,100)),
                        spread),
                    unsafe_allow_html=True)

                st.markdown(
                    "<br>", unsafe_allow_html=True)

                for line in verdict["lines"]:
                    st.markdown(
                        "<div style='background:white;"
                        "border-left:4px solid #D04A02;"
                        "padding:8px 14px;"
                        "border-radius:2px;"
                        "margin-bottom:6px;"
                        "font-size:0.87em'>"
                        "▸ {}</div>".format(line),
                        unsafe_allow_html=True)

                # Price comparison chart
                st.markdown(
                    "<br>", unsafe_allow_html=True)
                section_title(
                    "PRICE COMPARISON — ALL VENDORS")
                chart_rows = [
                    {
                        "Vendor": v,
                        "Price" : p,
                        "Color" : vendor_color_map.get(
                            v,"#8C8C8C"),
                    }
                    for v,p in sorted(
                        vendor_prices_map.items(),
                        key=lambda x:x[1])
                ]
                if chart_rows:
                    cdf = pd.DataFrame(chart_rows)
                    cf  = go.Figure(go.Bar(
                        x=cdf["Vendor"],
                        y=cdf["Price"],
                        marker_color=cdf["Color"],
                        marker_line_width=0,
                        text=cdf["Price"].apply(_fmt),
                        textposition="outside"))
                    cf.add_hline(
                        y=avg_p,
                        line_dash="dash",
                        line_color="#FFB600",
                        line_width=2,
                        annotation_text="Avg: {}".format(
                            _fmt(avg_p)),
                        annotation_position="top right")
                    cf.update_layout(
                        height=320,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=20,b=10),
                        font=CFONT,
                        yaxis=dict(
                            title="Price (USD)",
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-20),
                        bargap=0.4)
                    st.plotly_chart(
                        cf,use_container_width=True)

                # Vendor score card table
                st.markdown(
                    "<br>", unsafe_allow_html=True)
                section_title(
                    "VENDOR PRICE SCORE CARD")
                all_p_vals = list(
                    vendor_prices_map.values())
                rows = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>Vendor</th>"
                    "<th>Price</th>"
                    "<th>vs Average</th>"
                    "<th>Price Score</th>"
                    "<th>Verdict</th>"
                    "</tr></thead><tbody>"]

                for i,(v,p) in enumerate(sorted(
                        vendor_prices_map.items(),
                        key=lambda x:x[1])):
                    bg  = ("white" if i%2==0
                           else "#F3F3F3")
                    vc  = vendor_color_map.get(
                        v,"#8C8C8C")
                    others = [
                        px for px in all_p_vals
                        if px != p]
                    ps = None
                    if p > 0 and others:
                        ps,_,_,_,_ = price_score(
                            p, others)
                    sc = score_color(ps)
                    pct_vs = (
                        round((p-avg_p)/avg_p*100,1)
                        if avg_p > 0 else 0)
                    vs_txt = (
                        "{}% below avg".format(
                            abs(pct_vs))
                        if pct_vs < 0
                        else "{}% above avg".format(
                            abs(pct_vs))
                        if pct_vs > 0
                        else "At average")
                    vs_col = (
                        "#22992E" if pct_vs < 0
                        else "#E0301E" if pct_vs > 5
                        else "#856404")
                    vt, _, vc_col = get_verdict(ps)
                    rows.append(
                        "<tr style='background:{}'>"
                        "<td>{}</td>"
                        "<td style='font-family:"
                        "monospace;font-weight:700'>"
                        "{}</td>"
                        "<td style='color:{}'>{}</td>"
                        "<td style='text-align:center'>"
                        "<span style='font-weight:800;"
                        "font-size:1.1em;color:{}'>"
                        "{}/100</span></td>"
                        "<td><span style='color:{};"
                        "font-weight:700;"
                        "font-size:0.85em'>"
                        "{}</span></td>"
                        "</tr>".format(
                            bg,
                            vendor_pill(v,vc),
                            _fmt(p),
                            vs_col,vs_txt,
                            sc,
                            ps if ps is not None
                            else "—",
                            vc_col,vt))
                rows.append("</tbody></table>")
                st.markdown(
                    "".join(rows),
                    unsafe_allow_html=True)

                st.markdown(
                    "<hr style='border:none;"
                    "border-top:2px solid #D04A02;"
                    "margin:20px 0'>",
                    unsafe_allow_html=True)

            elif verdict and not verdict["has_prices"]:
                st.info(
                    "Services found but no price data. "
                    "Click **Extract Prices** in the "
                    "tables below.")

            # ── Per-service detail tables ───────────
            section_title(
                "QUOTATION FILES — PER SERVICE")

            vsmap = defaultdict(set)
            for _, r in d_sel.iterrows():
                vsmap[r["Vendor"]].add(r["Service"])
            vendors_all = sorted([
                v for v,s in vsmap.items()
                if set(selected_svcs).issubset(s)])

            if len(selected_svcs) > 1:
                if vendors_all:
                    st.success(
                        "✅ {} vendor(s) cover ALL "
                        "{} services: {}".format(
                            len(vendors_all),
                            len(selected_svcs),
                            " · ".join([
                                "**{}**".format(v)
                                for v in vendors_all])))
                else:
                    st.warning(
                        "No single vendor covers "
                        "all {} selected services."
                        .format(len(selected_svcs)))

            for svc in selected_svcs:
                d_svc = (
                    d_sel[d_sel["Service"]==svc]
                    .drop_duplicates(
                        subset=["Vendor","File Name"])
                    .sort_values("Vendor"))
                vc    = d_svc["Vendor"].nunique()
                s_tag = ("SHARED" if vc>1
                          else "SINGLE VENDOR")

                with st.expander(
                    "{}  —  {} vendor(s) · "
                    "{} file(s) · {}".format(
                        svc,vc,len(d_svc),s_tag),
                    expanded=True):

                    pills = " ".join([
                        vendor_pill(
                            v,vendor_color_map.get(
                                v,"#8C8C8C"))
                        for v in sorted(
                            d_svc["Vendor"].unique())])
                    st.markdown(
                        "<div style='margin-bottom:"
                        "12px'><b style='font-size:"
                        "0.87em'>Vendors:</b>"
                        "&nbsp;&nbsp;{}</div>".format(
                            pills),
                        unsafe_allow_html=True)

                    all_prices_svc = []
                    for _,r in d_svc.iterrows():
                        qp = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp > 0:
                            all_prices_svc.append(qp)
                        ck = "px_{}".format(
                            str(r.get(
                                "File Name","")).strip())
                        ca = st.session_state.get(ck)
                        if ca and ca.get(
                                "price_num",0) > 0:
                            all_prices_svc.append(
                                ca["price_num"])

                    rows = [
                        "<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>Vendor</th>"
                        "<th>Category</th>"
                        "<th>File Name</th>"]
                    if has_price:
                        rows.append(
                            "<th>Quoted Price</th>")
                    rows.append(
                        "<th>Extracted</th>"
                        "<th>Score</th>"
                        "<th>Verdict</th>"
                        "<th>Open</th>"
                        "</tr></thead><tbody>")

                    for i,(_,row) in enumerate(
                            d_svc.iterrows()):
                        bg  = ("white" if i%2==0
                               else "#F3F3F3")
                        vc2 = vendor_color_map.get(
                            row["Vendor"],"#8C8C8C")
                        fname = str(row.get(
                            "File Name","")).strip()
                        url   = resolve_url(row)

                        qp_str = str(row.get(
                            "Quoted Price","")).strip()
                        qp_num = _parse_num(qp_str)

                        ck     = "px_{}".format(fname)
                        cached = st.session_state.get(
                            ck)
                        if (cached and
                                cached.get(
                                    "price_num",0)>0):
                            ep_num = cached["price_num"]
                            ep_fmt = _fmt(
                                cached["price"])
                            ref    = ep_num
                        elif qp_num > 0:
                            ep_fmt = "—"
                            ref    = qp_num
                        else:
                            ep_fmt = "—"; ref = 0

                        others = [
                            p for p in all_prices_svc
                            if p != ref]
                        ps = None; vt="—"; vc_col="#bbb"
                        if ref > 0 and others:
                            ps,_,_,_,_ = price_score(
                                ref, others)
                            vt,_,vc_col = get_verdict(
                                ps)

                        sc = score_color(ps)
                        link_cell = (
                            "<a href='{}' "
                            "target='_blank' "
                            "style='color:#D04A02;"
                            "font-weight:600;"
                            "text-decoration:none'>"
                            "📂 Open</a>".format(url)
                            if url else "—")

                        rows.append(
                            "<tr style='background:{}'>"
                            "<td>{}</td>"
                            "<td style='color:#555'>"
                            "{}</td>"
                            "<td style='font-family:"
                            "monospace;font-size:0.79em;"
                            "word-break:break-all'>"
                            "{}</td>".format(
                                bg,
                                vendor_pill(
                                    row["Vendor"],vc2),
                                row.get("Category",""),
                                fname))
                        if has_price:
                            rows.append(
                                "<td style='color:"
                                "#22992E;font-weight:700;"
                                "font-family:monospace'>"
                                "{}</td>".format(
                                    _fmt(qp_str)
                                    if qp_num > 0
                                    else "—"))
                        rows.append(
                            "<td style='color:#295477;"
                            "font-weight:700;"
                            "font-family:monospace'>"
                            "{}</td>"
                            "<td style='text-align:"
                            "center'>"
                            "<span style='font-weight:"
                            "800;color:{}'>{}</span>"
                            "</td>"
                            "<td><span style='color:{};"
                            "font-weight:700;"
                            "font-size:0.82em'>{}</span>"
                            "</td>"
                            "<td>{}</td>"
                            "</tr>".format(
                                ep_fmt,
                                sc,
                                "{}/100".format(ps)
                                if ps is not None
                                else "—",
                                vc_col, vt,
                                link_cell))

                    rows.append("</tbody></table>")
                    st.markdown(
                        "".join(rows),
                        unsafe_allow_html=True)

                    if all_prices_svc:
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        ai_box(ai_price_insight(
                            0, all_prices_svc,
                            {v:p for v,p in
                             vendor_prices_map.items()
                             if p > 0}))

                    st.markdown(
                        "<br>", unsafe_allow_html=True)

                    if st.button(
                        "Extract Prices — {}".format(
                            svc[:40]),
                        key="ep_{}".format(svc[:35]),
                        type="primary"):
                        prog = st.progress(0)
                        n    = len(d_svc)
                        for ki,(_,row2) in enumerate(
                                d_svc.iterrows()):
                            fname2 = str(row2.get(
                                "File Name","")).strip()
                            ck2    = "px_{}".format(
                                fname2)
                            if st.session_state.get(
                                    ck2) is None:
                                local = os.path.join(
                                    DEMO_DIR, fname2)
                                if os.path.exists(local):
                                    res = extract_price_from_file(
                                        local)
                                    st.session_state[
                                        ck2] = res
                                else:
                                    url2 = resolve_url(
                                        row2)
                                    if (url2 and
                                            url2.startswith(
                                                "http") and
                                            REQUESTS_OK):
                                        try:
                                            resp = requests.get(
                                                url2,
                                                timeout=20)
                                            ext2 = (
                                                url2
                                                .split("?")[0]
                                                .rsplit(".",1)[-1]
                                                .lower())
                                            res = extract_price_from_bytes(
                                                resp.content,
                                                ext2)
                                            st.session_state[
                                                ck2] = res
                                        except Exception:
                                            pass
                            prog.progress((ki+1)/n)
                        prog.empty()
                        st.rerun()

# ════════════════════════════════════════════════════════════
# TAB 3 — UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab3:
    if NO_DATA:
        st.info(
            "No catalog loaded. "
            "Go to 🗂 Upload Catalog tab.")
    else:
        st.markdown(
            "<div style='background:#2D2D2D;"
            "color:white;padding:14px 20px;"
            "border-radius:4px;"
            "border-left:6px solid #D04A02;"
            "margin-bottom:16px'>"
            "<div style='font-size:0.72em;"
            "font-weight:700;letter-spacing:2px;"
            "text-transform:uppercase;"
            "color:#D04A02;margin-bottom:4px'>"
            "New Quotation Analysis</div>"
            "<div style='font-size:1.0em;"
            "font-weight:700'>"
            "Upload → auto-extract price → "
            "score &amp; verdict vs history</div>"
            "</div>",
            unsafe_allow_html=True)

        st.markdown(
            "<div style='background:#FFF3F0;"
            "border:1px solid #D04A02;"
            "border-radius:4px;"
            "padding:10px 16px;margin-bottom:12px;"
            "font-size:0.85em'>"
            "🎯 <b>Demo:</b> Upload "
            "<code>demo_quotes/"
            "DEMO_New_Quotation_AlphaNetworks.xlsx"
            "</code> "
            "to see a live price verdict vs "
            "historical data."
            "</div>",
            unsafe_allow_html=True)

        section_title("STEP 1 — UPLOAD FILE")
        uploaded = st.file_uploader(
            "Upload",
            type=["pdf","xlsx","xls","docx"],
            label_visibility="collapsed")

        if uploaded is not None:
            content  = uploaded.read()
            ext      = uploaded.name.rsplit(".",1)[-1]
            fname_up = uploaded.name
            st.success(
                "Uploaded: **{}** ({} KB)".format(
                    fname_up,
                    round(len(content)/1024,1)))

            section_title(
                "STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting price…"):
                result    = extract_price_from_bytes(
                    content, ext)
                new_price = result["price_num"]
                new_text  = result["text"]

            if new_price > 0:
                st.markdown(
                    "<div class='score-card green'>"
                    "<div style='font-size:0.72em;"
                    "font-weight:700;letter-spacing:1px;"
                    "text-transform:uppercase;"
                    "color:#22992E'>Extracted Price"
                    "</div>"
                    "<div style='font-size:2.0em;"
                    "font-weight:800;color:#22992E'>"
                    "{}</div>"
                    "</div>".format(_fmt(new_price)),
                    unsafe_allow_html=True)
            else:
                st.warning(
                    "Price not found automatically.")
                manual = st.number_input(
                    "Enter price manually (USD)",
                    min_value=0.0, step=100.0,
                    value=0.0,
                    key="manual_price_input")
                if manual > 0:
                    new_price = manual
                    st.success(
                        "Using manual price: **{}**".format(
                            _fmt(manual)))

            section_title(
                "STEP 3 — SELECT SERVICES")
            all_svcs_up = sorted([
                s for s in df_exploded[
                    "Service"].unique()
                if str(s).strip()
                not in ["","nan"]])
            svc_search_up = st.text_input(
                "Filter",
                placeholder="Search services…",
                key="svc_up",
                label_visibility="collapsed")
            if svc_search_up:
                all_svcs_up = [
                    s for s in all_svcs_up
                    if svc_search_up.lower()
                    in s.lower()]
            new_services = st.multiselect(
                "Services in this quotation",
                options=all_svcs_up,
                key="new_svcs",
                label_visibility="collapsed")

            cat_filter_up = st.selectbox(
                "Filter historical by category",
                options=["All"]+sorted([
                    c for c in
                    df_master["Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="cat_up")

            section_title(
                "STEP 4 — COMPARISON & VERDICT")

            # Auto-set price from manual input
            manual_val = st.session_state.get(
                "manual_price_input", 0.0)
            if new_price <= 0 and manual_val > 0:
                new_price = float(manual_val)
            if not new_services and new_price <= 0:
                st.info(
                    "Select services to compare.")
            else:
                candidates = (
                    df_exploded[
                        df_exploded["Service"].isin(
                            new_services)].copy()
                    if new_services
                    else df_exploded.copy())

                if cat_filter_up != "All":
                    candidates = candidates[
                        candidates["Category"]==
                        cat_filter_up]

                cand_files = (
                    candidates
                    .drop_duplicates(
                        subset=["File Name","Vendor"])
                    [["File Name","Vendor","Category",
                      "Hyperlink","Quoted Price"]]
                    .copy())

                if cand_files.empty:
                    st.warning(
                        "No historical quotes found.")
                else:
                    hist_prices = []
                    vendor_p_map = {}
                    for _,r in cand_files.iterrows():
                        qp = _parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp > 0:
                            hist_prices.append(qp)
                            vendor_p_map[
                                r["Vendor"]] = qp
                        ck = "px_{}".format(str(r.get(
                            "File Name","")).strip())
                        ca = st.session_state.get(ck)
                        if ca and ca.get(
                                "price_num",0) > 0:
                            hist_prices.append(
                                ca["price_num"])
                            vendor_p_map[
                                r["Vendor"]] = \
                                ca["price_num"]

                    if new_price>0 and hist_prices:
                        ps,ps_lbl,avg_h,mn_h,mx_h = \
                            price_score(
                                new_price,hist_prices)
                        vt,vt_desc,vt_col = \
                            get_verdict(ps)
                        vt_css = score_css(ps)

                        st.markdown(
                            "<div class='verdict-{}'>"
                            "<div style='font-size:"
                            "1.2em;margin-bottom:6px'>"
                            "{}</div>"
                            "<div style='font-size:"
                            "0.88em;font-weight:400'>"
                            "{}</div>"
                            "</div>".format(
                                vt_css,vt,vt_desc),
                            unsafe_allow_html=True)

                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)

                        sc1,sc2,sc3,sc4 = st.columns(4)
                        sc1.markdown(
                            "<div class='score-card {}'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:{}'>Price Score"
                            "</div>"
                            "<div style='font-size:2.0em;"
                            "font-weight:800;color:{}'>"
                            "{}/100</div>"
                            "<div style='font-size:0.75em;"
                            "color:#555;margin-top:4px'>"
                            "vs {} historical"
                            "</div></div>".format(
                                vt_css,vt_col,vt_col,
                                ps if ps is not None
                                else "N/A",
                                len(hist_prices)),
                            unsafe_allow_html=True)

                        sc2.markdown(
                            "<div class='score-card "
                            "yellow'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:#856404'>Your Price"
                            "</div>"
                            "<div style='font-size:1.8em;"
                            "font-weight:800;"
                            "color:#D04A02'>{}</div>"
                            "</div>".format(
                                _fmt(new_price)),
                            unsafe_allow_html=True)

                        sc3.markdown(
                            "<div class='score-card "
                            "yellow'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:#856404'>"
                            "Historical Avg</div>"
                            "<div style='font-size:1.8em;"
                            "font-weight:800;"
                            "color:#295477'>{}</div>"
                            "<div style='font-size:0.72em;"
                            "color:#555;margin-top:4px'>"
                            "min {} · max {}</div>"
                            "</div>".format(
                                _fmt(avg_h),
                                _fmt(mn_h),_fmt(mx_h)),
                            unsafe_allow_html=True)

                        sc4.markdown(
                            "<div class='score-card {}'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "letter-spacing:1px;"
                            "text-transform:uppercase;"
                            "color:{}'>vs Average</div>"
                            "<div style='font-size:1.1em;"
                            "font-weight:800;color:{};"
                            "margin-top:6px'>{}</div>"
                            "</div>".format(
                                vt_css,vt_col,
                                vt_col,ps_lbl),
                            unsafe_allow_html=True)

                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        ai_box(ai_price_insight(
                            new_price,hist_prices,
                            vendor_p_map))

                        # Price chart
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        section_title(
                            "PRICE POSITIONING CHART")

                        chart_data = []
                        for _,r in cand_files.iterrows():
                            fn     = str(r.get(
                                "File Name","")).strip()
                            ck     = "px_{}".format(fn)
                            cached = st.session_state\
                                .get(ck)
                            qp     = _parse_num(str(
                                r.get("Quoted Price",
                                      "")).strip())
                            ep     = (
                                cached["price_num"]
                                if cached else 0.0)
                            pval   = (ep if ep > 0
                                      else qp)
                            if pval > 0:
                                chart_data.append({
                                    "Label": "{}/{}".format(
                                        r["Vendor"],
                                        fn[:12]),
                                    "Price": pval,
                                    "Type" : "Historical",
                                    "Color": vendor_color_map
                                    .get(r["Vendor"],
                                         "#8C8C8C"),
                                })

                        chart_data.append({
                            "Label": "★ NEW: {}".format(
                                fname_up[:15]),
                            "Price": new_price,
                            "Type" : "New Upload",
                            "Color": "#D04A02",
                        })

                        cdf = pd.DataFrame(
                            chart_data).sort_values(
                            "Price")

                        cf = go.Figure()
                        hist_df = cdf[
                            cdf["Type"]=="Historical"]
                        new_df  = cdf[
                            cdf["Type"]=="New Upload"]

                        if not hist_df.empty:
                            cf.add_trace(go.Bar(
                                x=hist_df["Label"],
                                y=hist_df["Price"],
                                marker_color=hist_df[
                                    "Color"],
                                marker_line_width=0,
                                name="Historical",
                                text=hist_df["Price"]
                                .apply(_fmt),
                                textposition="outside"))

                        if not new_df.empty:
                            cf.add_trace(go.Bar(
                                x=new_df["Label"],
                                y=new_df["Price"],
                                marker_color="#D04A02",
                                marker_line_width=0,
                                name="Your Upload",
                                text=new_df["Price"]
                                .apply(_fmt),
                                textposition="outside"))

                        cf.add_hline(
                            y=avg_h,
                            line_dash="dash",
                            line_color="#FFB600",
                            line_width=2,
                            annotation_text="Avg: {}".format(
                                _fmt(avg_h)),
                            annotation_position=
                            "top right")

                        cf.update_layout(
                            height=360,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=20,b=10),
                            font=CFONT,
                            barmode="group",
                            yaxis=dict(
                                title="Price",
                                showgrid=True,
                                gridcolor="#E0E0E0",
                                zeroline=False),
                            xaxis=dict(tickangle=-25),
                            legend=dict(
                                orientation="h",
                                x=0,y=1.05),
                            bargap=0.25)

                        st.plotly_chart(
                            cf,
                            use_container_width=True)


# ════════════════════════════════════════════════════════════
# TAB 4 — DATA TABLE
# ════════════════════════════════════════════════════════════
with tab4:
    if NO_DATA:
        st.info(
            "No catalog loaded. "
            "Go to 🗂 Upload Catalog tab.")
    else:
        dm = df_master.copy()
        if selected_cat != "All":
            dm = dm[dm["Category"]==selected_cat]
        if selected_vendor != "All":
            dm = dm[dm["Vendor"]==selected_vendor]

        st.dataframe(
            dm.drop(
                columns=["Services List","Hyperlink"],
                errors="ignore"),
            use_container_width=True,
            height=500)


# ════════════════════════════════════════════════════════════
# TAB 5 — UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab5:
    st.markdown(
        "<div style='background:#2D2D2D;color:white;"
        "padding:20px 28px;border-radius:4px;"
        "border-left:6px solid #D04A02;"
        "margin-bottom:20px'>"
        "<div style='font-size:0.72em;font-weight:700;"
        "letter-spacing:2px;text-transform:uppercase;"
        "color:#D04A02;margin-bottom:5px'>"
        "Catalog Management</div>"
        "<h1 style='margin:0;font-size:1.2em;"
        "font-weight:700;color:white'>"
        "Upload Master Catalog</h1>"
        "<p style='margin:6px 0 0;opacity:0.6;"
        "font-size:0.85em'>"
        "Upload any Excel or CSV catalog — "
        "AI auto-detects columns and builds "
        "the dashboard automatically."
        "</p></div>",
        unsafe_allow_html=True)

    if DATA_SOURCE == "uploaded":
        st.success(
            "✅ Using uploaded catalog: "
            "**{}** rows · **{}** vendors · "
            "**{}** services".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique()))

    # Show what is currently loaded
    if not NO_DATA and DATA_SOURCE == "file":
        st.info(
            "📂 Currently using: **master_catalog.csv** "
            "— {} vendors · {} services · {} categories"
            .format(
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique(),
                df_master["Category"].nunique()))

    st.markdown(
        "<div style='font-size:0.78em;font-weight:700;"
        "letter-spacing:1px;text-transform:uppercase;"
        "color:#D04A02;margin:16px 0 10px'>"
        "UPLOAD A DIFFERENT CATALOG</div>",
        unsafe_allow_html=True)

    catalog_file = st.file_uploader(
        "Upload Master Catalog",
        type=["xlsx","xls","csv"],
        label_visibility="collapsed",
        key="catalog_upload")

    if catalog_file is not None:
        file_bytes = catalog_file.read()
        fname_cat  = catalog_file.name

        with st.spinner(
                "AI analyzing catalog…"):
            df_new,df_exp_new,err = \
                process_uploaded_catalog(
                    file_bytes, fname_cat)

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

            # Preview charts
            pc1,pc2 = st.columns(2)

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
                    marker_color=[
                        get_color(i)
                        for i in range(len(spv_new))],
                    marker_line_width=0,
                    text=spv_new["Services"],
                    textposition="outside"))
                pf1.update_layout(
                    title="Services per Vendor",
                    height=300,
                    plot_bgcolor=CBG,
                    paper_bgcolor=CBG,
                    margin=dict(
                        l=5,r=10,t=40,b=10),
                    font=CFONT,
                    yaxis=dict(
                        showgrid=True,
                        gridcolor="#E0E0E0",
                        zeroline=False),
                    xaxis=dict(tickangle=-30),
                    bargap=0.35)
                st.plotly_chart(
                    pf1,use_container_width=True)

            cat_new = (
                df_new.drop_duplicates(
                    subset=["Category","File Name"])
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
                        color_discrete_sequence=COLORS)
                    pf2.update_traces(
                        textposition="outside",
                        textinfo="label+percent",
                        textfont_size=10)
                    pf2.update_layout(
                        title="Category Distribution",
                        height=300,
                        margin=dict(
                            l=10,r=10,t=40,b=10),
                        paper_bgcolor=CBG,
                        font=CFONT)
                    st.plotly_chart(
                        pf2,use_container_width=True)

            # Data preview
            st.markdown(
                "<div style='font-size:0.78em;"
                "font-weight:700;letter-spacing:1px;"
                "text-transform:uppercase;"
                "color:#D04A02;margin:12px 0 6px'>"
                "DATA PREVIEW (first 20 rows)</div>",
                unsafe_allow_html=True)
            st.dataframe(
                df_new.drop(
                    columns=["Services List",
                             "Hyperlink"],
                    errors="ignore").head(20),
                use_container_width=True,
                height=280)

            if st.button(
                "✅ Apply This Catalog to Dashboard",
                type="primary"):
                st.session_state[
                    "uploaded_catalog_df"] = df_new
                st.session_state[
                    "uploaded_catalog_exp"] = df_exp_new
                st.success(
                    "✅ Catalog applied! "
                    "Dashboard updated.")
                st.rerun()


# ════════════════════════════════════════════════════════════
# TAB 6 — VENDOR ANALYSIS
# ════════════════════════════════════════════════════════════
with tab6:
    if NO_DATA:
        st.info(
            "No catalog loaded. "
            "Go to 🗂 Upload Catalog tab first.")
    else:
        # ── Header ──────────────────────────────────
        st.markdown(
            "<div style='background:#2D2D2D;"
            "color:white;padding:20px 28px;"
            "border-radius:4px;"
            "border-left:6px solid #D04A02;"
            "margin-bottom:20px'>"
            "<div style='font-size:0.72em;"
            "font-weight:700;letter-spacing:2px;"
            "text-transform:uppercase;"
            "color:#D04A02;margin-bottom:5px'>"
            "Master Catalog Intelligence</div>"
            "<h1 style='margin:0;font-size:1.3em;"
            "font-weight:700;color:white'>"
            "Vendor Price Analysis</h1>"
            "<p style='margin:6px 0 0;opacity:0.6;"
            "font-size:0.85em'>"
            "Prices extracted directly from quote "
            "files in GitHub repo · "
            "Cheap vs expensive analysis per service"
            "</p></div>",
            unsafe_allow_html=True)

        # ── Load all prices from GitHub ──────────────
        all_fnames = tuple(
            str(f).strip()
            for f in df_master["File Name"].unique()
            if str(f).strip() not in ["","nan"])

        col_load, col_info = st.columns([2,3])
        with col_load:
            run_analysis = st.button(
                "🔄 Load All Prices from GitHub",
                type="primary",
                use_container_width=True,
                key="run_gh_analysis")

        with col_info:
            st.markdown(
                "<div style='background:#FFF3F0;"
                "border:1px solid #D04A02;"
                "border-radius:4px;"
                "padding:8px 14px;font-size:0.83em'>"
                "📁 Reads <b>{}</b> quote files "
                "directly from your GitHub repo. "
                "Uses <b>Quoted Price</b> from catalog "
                "as fallback if extraction fails."
                "</div>".format(len(all_fnames)),
                unsafe_allow_html=True)

        # Cache key in session state
        if run_analysis:
            st.session_state[
                "gh_prices_loaded"] = True
            # Clear cache to re-fetch
            load_all_prices_from_github.clear()
            load_price_from_github.clear()

        # Auto-load if already loaded before
        if st.session_state.get(
                "gh_prices_loaded", False):

            with st.spinner(
                    "Loading prices from GitHub…"):
                gh_prices = load_all_prices_from_github(
                    all_fnames)

            # Build enriched dataframe
            rows = []
            for _, r in df_master.iterrows():
                fname = str(
                    r.get("File Name","")).strip()
                price, source = get_price_for_row(
                    r, gh_prices)
                gh_info = gh_prices.get(fname, {})
                rows.append({
                    "Vendor"  : r["Vendor"],
                    "Category": r["Category"],
                    "File Name": fname,
                    "Quoted Price": _parse_num(str(
                        r.get("Quoted Price",""))),
                    "Extracted Price": gh_info.get(
                        "price_num", 0.0),
                    "Best Price": price,
                    "Source"  : source,
                    "Status"  : gh_info.get(
                        "status","—"),
                    "Services": r.get("Comments",""),
                })
            df_analysis = pd.DataFrame(rows)
            df_analysis = df_analysis[
                df_analysis["Best Price"] > 0]

            if df_analysis.empty:
                st.warning(
                    "No prices found. "
                    "Check that files exist in "
                    "demo_quotes/ on GitHub.")
            else:
                # ── KPI summary ─────────────────────
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "PRICE EXTRACTION SUMMARY")

                k1,k2,k3,k4 = st.columns(4)
                extracted = df_analysis[
                    df_analysis["Source"]
                    =="extracted"]
                from_cat  = df_analysis[
                    df_analysis["Source"]
                    =="catalog"]

                kpi(k1,
                    len(df_analysis),
                    "Files with Prices",
                    "#D04A02")
                kpi(k2,
                    len(extracted),
                    "Extracted from Files",
                    "#295477")
                kpi(k3,
                    len(from_cat),
                    "From Catalog",
                    "#299D8F")
                kpi(k4,
                    _fmt(df_analysis[
                        "Best Price"].mean()),
                    "Avg Quote Value",
                    "#2D2D2D")

                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)

                # ── Extraction status table ──────────
                with st.expander(
                        "📋 Extraction Status per File",
                        expanded=False):
                    status_rows = [
                        "<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>File Name</th>"
                        "<th>Vendor</th>"
                        "<th>Quoted Price</th>"
                        "<th>Extracted Price</th>"
                        "<th>Used Price</th>"
                        "<th>Source</th>"
                        "<th>Status</th>"
                        "</tr></thead><tbody>"]
                    for i,r in df_analysis.iterrows():
                        bg = ("white" if i%2==0
                              else "#F3F3F3")
                        vc = vendor_color_map.get(
                            r["Vendor"],"#8C8C8C")
                        src_color = (
                            "#22992E"
                            if r["Source"]=="extracted"
                            else "#FFB600")
                        status_rows.append(
                            "<tr style='background:{}'>"
                            "<td style='font-family:"
                            "monospace;font-size:0.78em'>"
                            "{}</td>"
                            "<td>{}</td>"
                            "<td style='font-family:"
                            "monospace'>{}</td>"
                            "<td style='font-family:"
                            "monospace;color:#295477'>"
                            "{}</td>"
                            "<td style='font-family:"
                            "monospace;font-weight:700;"
                            "color:#D04A02'>{}</td>"
                            "<td style='color:{};"
                            "font-weight:700;"
                            "font-size:0.80em'>{}</td>"
                            "<td style='font-size:"
                            "0.80em'>{}</td>"
                            "</tr>".format(
                                bg,
                                r["File Name"],
                                vendor_pill(
                                    r["Vendor"],vc),
                                _fmt(r["Quoted Price"])
                                if r["Quoted Price"]>0
                                else "—",
                                _fmt(r["Extracted Price"])
                                if r["Extracted Price"]>0
                                else "—",
                                _fmt(r["Best Price"]),
                                src_color,
                                r["Source"].upper(),
                                r["Status"]))
                    status_rows.append(
                        "</tbody></table>")
                    st.markdown(
                        "".join(status_rows),
                        unsafe_allow_html=True)

                # ════════════════════════════════════
                # SECTION A — Per-Service Analysis
                # ════════════════════════════════════
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "PER-SERVICE VENDOR ANALYSIS",
                    "For each service: which vendor "
                    "was cheapest and most expensive.")

                # Explode services for analysis
                df_svc_analysis = []
                for _, r in df_analysis.iterrows():
                    svcs_raw = str(
                        r["Services"]).replace(
                        "\\n","\n").replace(
                        "\r\n","\n").replace(
                        "\r","\n")
                    svcs = [
                        s.strip()
                        for s in svcs_raw.split("\n")
                        if s.strip() and
                        s.strip() not in
                        ["nan","None",""]]
                    if not svcs:
                        svcs = [svcs_raw.strip()]
                    for svc in svcs:
                        df_svc_analysis.append({
                            "Service" : svc,
                            "Vendor"  : r["Vendor"],
                            "Category": r["Category"],
                            "Price"   : r["Best Price"],
                            "File"    : r["File Name"],
                            "Source"  : r["Source"],
                        })

                df_svc_df = pd.DataFrame(
                    df_svc_analysis)

                # Filter to services with
                # multiple vendor quotes
                svc_vendor_counts = (
                    df_svc_df.groupby("Service")[
                        "Vendor"].nunique())
                multi_vendor_svcs = svc_vendor_counts[
                    svc_vendor_counts > 1].index.tolist()
                single_vendor_svcs = svc_vendor_counts[
                    svc_vendor_counts == 1].index.tolist()

                # ── Multi-vendor services ────────────
                if multi_vendor_svcs:
                    st.markdown(
                        "<div style='background:#F0FFF4;"
                        "border-left:4px solid #22992E;"
                        "padding:8px 14px;"
                        "border-radius:2px;"
                        "margin-bottom:12px;"
                        "font-size:0.87em'>"
                        "✅ <b>{} service(s)</b> quoted "
                        "by multiple vendors — "
                        "full price comparison available."
                        "</div>".format(
                            len(multi_vendor_svcs)),
                        unsafe_allow_html=True)

                    for svc in sorted(
                            multi_vendor_svcs):
                        d_svc = df_svc_df[
                            df_svc_df["Service"]==svc
                        ].sort_values("Price")

                        min_p   = d_svc["Price"].min()
                        max_p   = d_svc["Price"].max()
                        avg_p   = d_svc["Price"].mean()
                        spread  = round(
                            (max_p-min_p)/min_p*100,1
                        ) if min_p > 0 else 0
                        best_v  = d_svc.loc[
                            d_svc["Price"].idxmin(),
                            "Vendor"]
                        worst_v = d_svc.loc[
                            d_svc["Price"].idxmax(),
                            "Vendor"]

                        with st.expander(
                            "{}  ·  {} vendors  ·  "
                            "spread {}%  ·  "
                            "best: {} @ {}".format(
                                svc,
                                d_svc["Vendor"].nunique(),
                                spread,
                                best_v,
                                _fmt(min_p)),
                            expanded=False):

                            # Score cards
                            sc1,sc2,sc3 = st.columns(3)
                            sc1.markdown(
                                "<div class='score-card "
                                "green'>"
                                "<div style='font-size:"
                                "0.68em;font-weight:700;"
                                "text-transform:uppercase;"
                                "color:#22992E'>"
                                "Cheapest</div>"
                                "<div style='font-size:"
                                "1.6em;font-weight:800;"
                                "color:#22992E'>{}</div>"
                                "<div style='font-size:"
                                "0.78em;color:#555'>"
                                "{}</div>"
                                "</div>".format(
                                    _fmt(min_p),best_v),
                                unsafe_allow_html=True)
                            sc2.markdown(
                                "<div class='score-card "
                                "yellow'>"
                                "<div style='font-size:"
                                "0.68em;font-weight:700;"
                                "text-transform:uppercase;"
                                "color:#856404'>"
                                "Average</div>"
                                "<div style='font-size:"
                                "1.6em;font-weight:800;"
                                "color:#856404'>{}</div>"
                                "<div style='font-size:"
                                "0.78em;color:#555'>"
                                "{} vendors</div>"
                                "</div>".format(
                                    _fmt(avg_p),
                                    d_svc["Vendor"]
                                    .nunique()),
                                unsafe_allow_html=True)
                            sc3.markdown(
                                "<div class='score-card "
                                "red'>"
                                "<div style='font-size:"
                                "0.68em;font-weight:700;"
                                "text-transform:uppercase;"
                                "color:#E0301E'>"
                                "Most Expensive</div>"
                                "<div style='font-size:"
                                "1.6em;font-weight:800;"
                                "color:#E0301E'>{}</div>"
                                "<div style='font-size:"
                                "0.78em;color:#555'>"
                                "{}</div>"
                                "</div>".format(
                                    _fmt(max_p),
                                    worst_v),
                                unsafe_allow_html=True)

                            st.markdown(
                                "<br>",
                                unsafe_allow_html=True)

                            # Per-vendor table
                            tbl = [
                                "<table class='"
                                "comp-table'>"
                                "<thead><tr>"
                                "<th>Vendor</th>"
                                "<th>Price</th>"
                                "<th>vs Average</th>"
                                "<th>Price Score</th>"
                                "<th>Verdict</th>"
                                "<th>Source</th>"
                                "</tr></thead><tbody>"]

                            all_p = d_svc[
                                "Price"].tolist()
                            for j,(_,vr) in enumerate(
                                    d_svc.iterrows()):
                                bg  = ("white"
                                       if j%2==0
                                       else "#F3F3F3")
                                vc  = vendor_color_map\
                                    .get(vr["Vendor"],
                                         "#8C8C8C")
                                p   = vr["Price"]
                                others = [
                                    x for x in all_p
                                    if x != p]
                                ps  = None
                                if p>0 and others:
                                    ps,_,_,_,_ = \
                                        price_score(
                                            p,others)
                                sc_col = score_color(ps)
                                vt,_,vt_col = \
                                    get_verdict(ps)
                                pct = (round(
                                    (p-avg_p)/avg_p*100,1)
                                    if avg_p>0 else 0)
                                pct_col = (
                                    "#22992E"
                                    if pct < 0
                                    else "#E0301E"
                                    if pct > 5
                                    else "#856404")
                                pct_txt = (
                                    "{}% below avg"
                                    .format(abs(pct))
                                    if pct < 0
                                    else "{}% above avg"
                                    .format(abs(pct))
                                    if pct > 0
                                    else "At average")
                                src_badge = (
                                    "<span style='"
                                    "background:#22992E;"
                                    "color:white;"
                                    "padding:2px 6px;"
                                    "border-radius:2px;"
                                    "font-size:0.72em;"
                                    "font-weight:700'>"
                                    "EXTRACTED</span>"
                                    if vr["Source"]
                                    =="extracted"
                                    else
                                    "<span style='"
                                    "background:#FFB600;"
                                    "color:white;"
                                    "padding:2px 6px;"
                                    "border-radius:2px;"
                                    "font-size:0.72em;"
                                    "font-weight:700'>"
                                    "CATALOG</span>")
                                tbl.append(
                                    "<tr style='"
                                    "background:{}'>"
                                    "<td>{}</td>"
                                    "<td style='"
                                    "font-family:monospace;"
                                    "font-weight:700'>"
                                    "{}</td>"
                                    "<td style='color:{}'>"
                                    "{}</td>"
                                    "<td style='"
                                    "text-align:center'>"
                                    "<span style='"
                                    "font-weight:800;"
                                    "font-size:1.1em;"
                                    "color:{}'>"
                                    "{}/100</span></td>"
                                    "<td><span style='"
                                    "color:{};"
                                    "font-weight:700;"
                                    "font-size:0.85em'>"
                                    "{}</span></td>"
                                    "<td>{}</td>"
                                    "</tr>".format(
                                        bg,
                                        vendor_pill(
                                            vr["Vendor"],
                                            vc),
                                        _fmt(p),
                                        pct_col,pct_txt,
                                        sc_col,
                                        ps if ps
                                        is not None
                                        else "—",
                                        vt_col,vt,
                                        src_badge))

                            tbl.append(
                                "</tbody></table>")
                            st.markdown(
                                "".join(tbl),
                                unsafe_allow_html=True)

                            # Bar chart
                            st.markdown(
                                "<br>",
                                unsafe_allow_html=True)
                            fig_svc = go.Figure(
                                go.Bar(
                                    x=d_svc["Vendor"],
                                    y=d_svc["Price"],
                                    marker_color=[
                                        vendor_color_map
                                        .get(v,"#8C8C8C")
                                        for v in
                                        d_svc["Vendor"]],
                                    marker_line_width=0,
                                    text=d_svc["Price"]
                                    .apply(_fmt),
                                    textposition=
                                    "outside"))
                            fig_svc.add_hline(
                                y=avg_p,
                                line_dash="dash",
                                line_color="#FFB600",
                                line_width=2,
                                annotation_text=
                                "Avg: {}".format(
                                    _fmt(avg_p)),
                                annotation_position=
                                "top right")
                            fig_svc.update_layout(
                                height=280,
                                plot_bgcolor=CBG,
                                paper_bgcolor=CBG,
                                margin=dict(
                                    l=5,r=10,t=20,b=10),
                                font=CFONT,
                                yaxis=dict(
                                    title="Price (USD)",
                                    showgrid=True,
                                    gridcolor="#E0E0E0",
                                    zeroline=False),
                                xaxis=dict(
                                    tickangle=-15),
                                bargap=0.4)
                            st.plotly_chart(
                                fig_svc,
                                use_container_width=True)

                # ════════════════════════════════════
                # SECTION B — Vendor Overview
                # ════════════════════════════════════
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "VENDOR OVERVIEW — TOTAL SPEND",
                    "Total quoted value per vendor "
                    "across all services.")

                vendor_totals = (
                    df_analysis.groupby("Vendor")[
                        "Best Price"]
                    .agg(["sum","mean","count"])
                    .reset_index())
                vendor_totals.columns = [
                    "Vendor","Total","Average","Quotes"]
                vendor_totals = vendor_totals\
                    .sort_values("Total",
                                 ascending=False)

                vt1,vt2 = st.columns(2)

                with vt1:
                    fig_tot = go.Figure(go.Bar(
                        x=vendor_totals["Vendor"],
                        y=vendor_totals["Total"],
                        marker_color=[
                            vendor_color_map.get(
                                v,"#8C8C8C")
                            for v in
                            vendor_totals["Vendor"]],
                        marker_line_width=0,
                        text=vendor_totals[
                            "Total"].apply(_fmt),
                        textposition="outside"))
                    fig_tot.update_layout(
                        title="Total Quoted Value "
                              "per Vendor",
                        height=350,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(
                            l=5,r=10,t=40,b=10),
                        font=CFONT,
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-30),
                        bargap=0.35)
                    st.plotly_chart(
                        fig_tot,
                        use_container_width=True)

                with vt2:
                    fig_avg = go.Figure(go.Bar(
                        x=vendor_totals["Vendor"],
                        y=vendor_totals["Average"],
                        marker_color=[
                            vendor_color_map.get(
                                v,"#8C8C8C")
                            for v in
                            vendor_totals["Vendor"]],
                        marker_line_width=0,
                        text=vendor_totals[
                            "Average"].apply(_fmt),
                        textposition="outside"))
                    grand_avg = (
                        df_analysis["Best Price"]
                        .mean())
                    fig_avg.add_hline(
                        y=grand_avg,
                        line_dash="dash",
                        line_color="#FFB600",
                        line_width=2,
                        annotation_text=
                        "Overall Avg: {}".format(
                            _fmt(grand_avg)),
                        annotation_position=
                        "top right")
                    fig_avg.update_layout(
                        title="Average Quote Value "
                              "per Vendor",
                        height=350,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(
                            l=5,r=10,t=40,b=10),
                        font=CFONT,
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-30),
                        bargap=0.35)
                    st.plotly_chart(
                        fig_avg,
                        use_container_width=True)

                # ── Vendor summary table ─────────────
                section_title("VENDOR SUMMARY TABLE")
                vtbl = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>Vendor</th>"
                    "<th>Quotes</th>"
                    "<th>Total Value</th>"
                    "<th>Average Quote</th>"
                    "<th>Min Quote</th>"
                    "<th>Max Quote</th>"
                    "<th>Overall Verdict</th>"
                    "</tr></thead><tbody>"]

                overall_avg = (
                    df_analysis["Best Price"].mean())

                for i,vr in vendor_totals.iterrows():
                    bg  = ("white" if i%2==0
                           else "#F3F3F3")
                    vc  = vendor_color_map.get(
                        vr["Vendor"],"#8C8C8C")
                    v_df = df_analysis[
                        df_analysis["Vendor"]
                        ==vr["Vendor"]]
                    v_min = v_df["Best Price"].min()
                    v_max = v_df["Best Price"].max()
                    v_avg = vr["Average"]
                    pct   = round(
                        (v_avg-overall_avg)
                        /overall_avg*100,1
                    ) if overall_avg > 0 else 0
                    pct_col = (
                        "#22992E" if pct < -5
                        else "#E0301E" if pct > 5
                        else "#856404")
                    pct_txt = (
                        "{}% below avg".format(
                            abs(pct))
                        if pct < 0
                        else "{}% above avg".format(
                            abs(pct))
                        if pct > 0
                        else "At average")
                    # Overall verdict for vendor
                    if pct < -10:
                        ov = ("✅ COMPETITIVE",
                              "#22992E")
                    elif pct > 10:
                        ov = ("🔴 EXPENSIVE",
                              "#E0301E")
                    else:
                        ov = ("🟡 AVERAGE",
                              "#856404")
                    vtbl.append(
                        "<tr style='background:{}'>"
                        "<td>{}</td>"
                        "<td style='text-align:center;"
                        "font-weight:700'>{}</td>"
                        "<td style='font-family:"
                        "monospace;font-weight:700;"
                        "color:#D04A02'>{}</td>"
                        "<td style='font-family:"
                        "monospace'>{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#22992E'>"
                        "{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#E0301E'>"
                        "{}</td>"
                        "<td style='color:{};"
                        "font-weight:700;"
                        "font-size:0.85em'>{}</td>"
                        "</tr>".format(
                            bg,
                            vendor_pill(
                                vr["Vendor"],vc),
                            int(vr["Quotes"]),
                            _fmt(vr["Total"]),
                            _fmt(v_avg),
                            _fmt(v_min),
                            _fmt(v_max),
                            ov[1],ov[0]))

                vtbl.append("</tbody></table>")
                st.markdown(
                    "".join(vtbl),
                    unsafe_allow_html=True)

                # ════════════════════════════════════
                # SECTION C — Category Analysis
                # ════════════════════════════════════
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "CATEGORY PRICE ANALYSIS",
                    "Average spend and vendor "
                    "distribution per category.")

                cat_totals = (
                    df_analysis.groupby("Category")[
                        "Best Price"]
                    .agg(["mean","sum","count"])
                    .reset_index())
                cat_totals.columns = [
                    "Category","Average",
                    "Total","Quotes"]
                cat_totals = cat_totals.sort_values(
                    "Total",ascending=False)

                ct1,ct2 = st.columns(2)
                with ct1:
                    fig_cat_tot = px.bar(
                        cat_totals,
                        x="Category",
                        y="Total",
                        color="Category",
                        color_discrete_sequence=COLORS,
                        text="Total",
                        title="Total Spend "
                              "per Category")
                    fig_cat_tot.update_traces(
                        texttemplate="%{text:$,.0f}",
                        textposition="outside")
                    fig_cat_tot.update_layout(
                        height=350,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(
                            l=5,r=10,t=40,b=10),
                        font=CFONT,
                        showlegend=False,
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-30))
                    st.plotly_chart(
                        fig_cat_tot,
                        use_container_width=True)

                with ct2:
                    # Vendor distribution per category
                    cat_vendor = (
                        df_analysis.groupby(
                            ["Category","Vendor"])
                        ["Best Price"].mean()
                        .reset_index())
                    cat_vendor.columns = [
                        "Category","Vendor","Price"]
                    fig_cat_v = px.bar(
                        cat_vendor,
                        x="Category",
                        y="Price",
                        color="Vendor",
                        barmode="group",
                        color_discrete_sequence=COLORS,
                        title="Vendor Prices "
                              "by Category",
                        text="Price")
                    fig_cat_v.update_traces(
                        texttemplate="%{text:$,.0f}",
                        textposition="outside")
                    fig_cat_v.update_layout(
                        height=350,
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(
                            l=5,r=10,t=40,b=10),
                        font=CFONT,
                        yaxis=dict(
                            showgrid=True,
                            gridcolor="#E0E0E0",
                            zeroline=False),
                        xaxis=dict(tickangle=-30),
                        legend=dict(
                            orientation="h",
                            x=0,y=-0.3,
                            font=dict(size=9)))
                    st.plotly_chart(
                        fig_cat_v,
                        use_container_width=True)

                # ════════════════════════════════════
                # SECTION D — Single vendor services
                # ════════════════════════════════════
                if single_vendor_svcs:
                    st.markdown(
                        "<br>",
                        unsafe_allow_html=True)
                    section_title(
                        "SINGLE-VENDOR SERVICES",
                        "These services only have "
                        "one vendor — no competition.")
                    st.markdown(
                        "<div style='background:"
                        "#FFF8E1;"
                        "border-left:4px solid "
                        "#FFB600;"
                        "padding:8px 14px;"
                        "border-radius:2px;"
                        "margin-bottom:12px;"
                        "font-size:0.87em'>"
                        "⚠️ <b>{} service(s)</b> have "
                        "only one vendor — "
                        "consider sourcing additional "
                        "quotes for better negotiation."
                        "</div>".format(
                            len(single_vendor_svcs)),
                        unsafe_allow_html=True)

                    sv_rows = [
                        "<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>Service</th>"
                        "<th>Vendor</th>"
                        "<th>Category</th>"
                        "<th>Price</th>"
                        "<th>Recommendation</th>"
                        "</tr></thead><tbody>"]

                    for j,svc in enumerate(
                            sorted(single_vendor_svcs)):
                        d_sv  = df_svc_df[
                            df_svc_df["Service"]==svc]
                        if d_sv.empty: continue
                        row_s = d_sv.iloc[0]
                        bg    = ("white" if j%2==0
                                 else "#F3F3F3")
                        vc    = vendor_color_map.get(
                            row_s["Vendor"],"#8C8C8C")
                        sv_rows.append(
                            "<tr style='background:{}'>"
                            "<td style='font-weight:"
                            "600'>{}</td>"
                            "<td>{}</td>"
                            "<td style='color:#555'>"
                            "{}</td>"
                            "<td style='font-family:"
                            "monospace;font-weight:700;"
                            "color:#295477'>{}</td>"
                            "<td style='color:#856404;"
                            "font-size:0.82em'>"
                            "⚠️ Seek additional "
                            "quotes</td>"
                            "</tr>".format(
                                bg,
                                svc,
                                vendor_pill(
                                    row_s["Vendor"],
                                    vc),
                                row_s["Category"],
                                _fmt(row_s["Price"])))

                    sv_rows.append(
                        "</tbody></table>")
                    st.markdown(
                        "".join(sv_rows),
                        unsafe_allow_html=True)

                # ════════════════════════════════════
                # SECTION E — AI Recommendations
                # ════════════════════════════════════
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "AI PROCUREMENT RECOMMENDATIONS")

                recs = []

                # Cheapest vendor overall
                if not vendor_totals.empty:
                    cheapest_v = vendor_totals.iloc[
                        vendor_totals["Average"]
                        .values.argmin()]["Vendor"]
                    cheapest_avg = vendor_totals[
                        vendor_totals["Vendor"]
                        ==cheapest_v][
                        "Average"].values[0]
                    recs.append({
                        "icon" : "✅",
                        "color": "#22992E",
                        "text" : (
                            "<b>{}</b> offers the "
                            "lowest average quote at "
                            "<b>{}</b> — preferred "
                            "vendor for cost "
                            "optimisation.".format(
                                cheapest_v,
                                _fmt(cheapest_avg)))
                    })

                # Most expensive vendor
                if not vendor_totals.empty:
                    exp_v = vendor_totals.iloc[
                        vendor_totals["Average"]
                        .values.argmax()]["Vendor"]
                    exp_avg = vendor_totals[
                        vendor_totals["Vendor"]
                        ==exp_v][
                        "Average"].values[0]
                    pct_diff = round(
                        (exp_avg-cheapest_avg)
                        /cheapest_avg*100,1
                    ) if cheapest_avg > 0 else 0
                    if pct_diff > 10:
                        recs.append({
                            "icon" : "🔴",
                            "color": "#E0301E",
                            "text" : (
                                "<b>{}</b> is the most "
                                "expensive vendor — "
                                "<b>{}% higher</b> "
                                "than the cheapest. "
                                "Negotiate or "
                                "shortlist alternatives."
                                .format(
                                    exp_v,pct_diff))
                        })

                # High competition services
                if len(multi_vendor_svcs) > 0:
                    recs.append({
                        "icon" : "💡",
                        "color": "#295477",
                        "text" : (
                            "<b>{} service(s)</b> have "
                            "multiple vendors quoting — "
                            "use competitive pressure "
                            "to negotiate better "
                            "pricing.".format(
                                len(multi_vendor_svcs)))
                    })

                # Single vendor risk
                if len(single_vendor_svcs) > 0:
                    recs.append({
                        "icon" : "⚠️",
                        "color": "#FFB600",
                        "text" : (
                            "<b>{} service(s)</b> have "
                            "only one vendor — "
                            "procurement risk. "
                            "Seek additional quotes "
                            "before awarding.".format(
                                len(single_vendor_svcs)))
                    })

                # Price spread
                if multi_vendor_svcs:
                    spreads = []
                    for svc in multi_vendor_svcs:
                        d_s = df_svc_df[
                            df_svc_df["Service"]==svc]
                        mn  = d_s["Price"].min()
                        mx  = d_s["Price"].max()
                        if mn > 0:
                            spreads.append(
                                (mx-mn)/mn*100)
                    if spreads:
                        avg_spread = round(
                            sum(spreads)/len(spreads),1)
                        if avg_spread > 15:
                            recs.append({
                                "icon" : "📊",
                                "color": "#D04A02",
                                "text" : (
                                    "Average price "
                                    "spread across "
                                    "competitive "
                                    "services is "
                                    "<b>{}%</b> — "
                                    "significant "
                                    "savings available "
                                    "through vendor "
                                    "selection.".format(
                                        avg_spread))
                            })

                for rec in recs:
                    st.markdown(
                        "<div style='background:"
                        "white;"
                        "border-left:4px solid {};"
                        "padding:10px 16px;"
                        "border-radius:2px;"
                        "margin-bottom:8px;"
                        "font-size:0.87em'>"
                        "{} {}</div>".format(
                            rec["color"],
                            rec["icon"],
                            rec["text"]),
                        unsafe_allow_html=True)

                # ════════════════════════════════════
                # SECTION F — Full data download
                # ════════════════════════════════════
                st.markdown(
                    "<br>",
                    unsafe_allow_html=True)
                section_title(
                    "FULL ANALYSIS DATA")

                # Show full table
                display_cols = [
                    "Vendor","Category",
                    "File Name","Best Price","Source"]
                st.dataframe(
                    df_analysis[display_cols]
                    .sort_values(
                        ["Category","Best Price"]),
                    use_container_width=True,
                    height=320)

                # Download button
                csv_out = df_analysis[
                    display_cols].to_csv(index=False)
                st.download_button(
                    label="📥 Download Full Analysis CSV",
                    data=csv_out,
                    file_name="vendor_analysis.csv",
                    mime="text/csv",
                    type="primary")

        else:
            # Not yet loaded — show instructions
            st.markdown(
                "<br>",
                unsafe_allow_html=True)
            st.markdown(
                "<div style='background:white;"
                "border:1px solid #e0e0e0;"
                "border-radius:4px;"
                "padding:24px 28px;"
                "text-align:center;"
                "margin-top:20px'>"
                "<div style='font-size:2em;"
                "margin-bottom:12px'>🔍</div>"
                "<div style='font-size:1.1em;"
                "font-weight:700;color:#2D2D2D;"
                "margin-bottom:8px'>"
                "Click the button above to start "
                "analysis</div>"
                "<div style='font-size:0.85em;"
                "color:#7D7D7D;max-width:500px;"
                "margin:0 auto'>"
                "The dashboard will fetch all "
                "<b>{}</b> quote files directly from "
                "your GitHub repo and extract prices "
                "automatically. "
                "Catalog prices are used as fallback."
                "</div></div>".format(len(all_fnames)),
                unsafe_allow_html=True)

            # Show what will be loaded
            st.markdown(
                "<br>",
                unsafe_allow_html=True)
            with st.expander(
                    "📁 Files that will be loaded "
                    "({})".format(len(all_fnames)),
                    expanded=False):
                preview_rows = [
                    "<table class='comp-table'>"
                    "<thead><tr>"
                    "<th>File Name</th>"
                    "<th>Vendor</th>"
                    "<th>Category</th>"
                    "<th>Catalog Price</th>"
                    "</tr></thead><tbody>"]
                for i,(_,r) in enumerate(
                        df_master.iterrows()):
                    bg  = ("white" if i%2==0
                           else "#F3F3F3")
                    vc  = vendor_color_map.get(
                        r["Vendor"],"#8C8C8C")
                    qp  = _parse_num(str(
                        r.get("Quoted Price",
                              "")).strip())
                    preview_rows.append(
                        "<tr style='background:{}'>"
                        "<td style='font-family:"
                        "monospace;font-size:0.78em'>"
                        "{}</td>"
                        "<td>{}</td>"
                        "<td style='color:#555'>"
                        "{}</td>"
                        "<td style='font-family:"
                        "monospace;color:#295477'>"
                        "{}</td>"
                        "</tr>".format(
                            bg,
                            r["File Name"],
                            vendor_pill(
                                r["Vendor"],vc),
                            r["Category"],
                            _fmt(qp)
                            if qp>0 else "—"))
                preview_rows.append(
                    "</tbody></table>")
                st.markdown(
                    "".join(preview_rows),
                    unsafe_allow_html=True)




# ════════════════════════════════════════════════════════════
# TAB 7 — REAL CATALOG ANALYSIS
# ════════════════════════════════════════════════════════════
with tab7:

    st.markdown(
        "<div style='background:#2D2D2D;"
        "color:white;padding:20px 28px;"
        "border-radius:4px;"
        "border-left:6px solid #D04A02;"
        "margin-bottom:20px'>"
        "<div style='font-size:0.72em;"
        "font-weight:700;letter-spacing:2px;"
        "text-transform:uppercase;"
        "color:#D04A02;margin-bottom:5px'>"
        "Real Data Intelligence</div>"
        "<h1 style='margin:0;font-size:1.3em;"
        "font-weight:700;color:white'>"
        "Real Quotation Analysis</h1>"
        "<p style='margin:6px 0 0;opacity:0.6;"
        "font-size:0.85em'>"
        "Upload your Master Catalog Excel → "
        "dashboard follows the File Links → "
        "downloads each quote file → "
        "extracts prices → full vendor analysis"
        "</p></div>",
        unsafe_allow_html=True)

    # ── Upload real Master Catalog ───────────────
    st.markdown(
        "<div style='font-size:0.78em;"
        "font-weight:700;letter-spacing:1px;"
        "text-transform:uppercase;"
        "color:#D04A02;margin-bottom:8px'>"
        "STEP 1 — UPLOAD YOUR MASTER CATALOG EXCEL"
        "</div>",
        unsafe_allow_html=True)

    st.markdown(
        "<div style='background:#EEF4FF;"
        "border-left:4px solid #295477;"
        "padding:10px 14px;border-radius:2px;"
        "margin-bottom:12px;font-size:0.84em'>"
        "📋 Your Excel must have these columns: "
        "<b>Category</b> · <b>Vendor</b> · "
        "<b>File Name</b> · <b>Quoted Price</b> · "
        "<b>Comments</b> (services) · "
        "<b>File Link</b> (URL to each quote file)"
        "<br><br>"
        "🔗 File Link can be: "
        "SharePoint URL · OneDrive link · "
        "Google Drive link · "
        "Any direct download URL · "
        "GitHub raw URL"
        "</div>",
        unsafe_allow_html=True)

    uploaded_real = st.file_uploader(
        "Upload Master Catalog",
        type=["xlsx","xls","csv"],
        key="tab7_real_upload",
        label_visibility="collapsed")

    if uploaded_real is None:
        # Try using current loaded catalog
        if not NO_DATA and not df_master.empty:
            use_real_cat = st.checkbox(
                "Use currently loaded catalog "
                "({} files)".format(
                    len(df_master)),
                key="use_current_for_real")
            if use_real_cat:
                df_real_cat = df_master.copy()
                st.success(
                    "Using current catalog — "
                    "**{}** vendors · "
                    "**{}** files".format(
                        df_real_cat[
                            "Vendor"].nunique(),
                        len(df_real_cat)))
            else:
                st.info(
                    "👆 Upload your Master Catalog "
                    "Excel to start analysis.")
                st.stop()
        else:
            st.info(
                "👆 Upload your Master Catalog "
                "Excel to start analysis.")
            st.stop()
    else:
        # Process uploaded catalog
        real_bytes = uploaded_real.read()
        df_real_cat, _, real_err = \
            process_uploaded_catalog(
                real_bytes,
                uploaded_real.name)
        if real_err or df_real_cat is None:
            st.error(
                "❌ Could not read: {}".format(
                    real_err))
            st.stop()

        # Check for file link column
        link_col_found = any(
            c.lower().strip() in [
                "file link","file url","url",
                "hyperlink","link"]
            for c in df_real_cat.columns)

        if not link_col_found:
            st.warning(
                "⚠️ No 'File Link' column found. "
                "Quoted Price from catalog will "
                "be used instead of extracting "
                "from files.")

        st.success(
            "✅ Catalog loaded: "
            "**{}** vendors · "
            "**{}** files · "
            "**{}** categories".format(
                df_real_cat["Vendor"].nunique(),
                len(df_real_cat),
                df_real_cat["Category"].nunique()))

        # Show preview
        with st.expander(
                "Preview catalog data",
                expanded=False):
            st.dataframe(
                df_real_cat.drop(
                    columns=[
                        "Services List",
                        "Hyperlink"],
                    errors="ignore"),
                use_container_width=True,
                height=200)

    # ── Run Analysis ─────────────────────────────
    st.markdown("<br>",unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:0.78em;"
        "font-weight:700;letter-spacing:1px;"
        "text-transform:uppercase;"
        "color:#D04A02;margin-bottom:8px'>"
        "STEP 2 — RUN ANALYSIS</div>",
        unsafe_allow_html=True)

    run_real = st.button(
        "🚀 Run Full Analysis",
        type="primary",
        key="run_real_analysis",
        use_container_width=False)

    if run_real:
        st.session_state[
            "real_analysis_done"] = False
        st.session_state[
            "real_analysis_df"]   = None

    if run_real or st.session_state.get(
            "real_analysis_done", False):

        if run_real:
            with st.spinner(
                    "Analyzing {} quote files…"
                    .format(len(df_real_cat))):
                df_result = analyze_real_catalog(
                    df_real_cat)
            st.session_state[
                "real_analysis_done"] = True
            st.session_state[
                "real_analysis_df"]   = \
                df_result.to_dict("records")
        else:
            df_result = pd.DataFrame(
                st.session_state[
                    "real_analysis_df"])

        df_priced = df_result[
            df_result["Best Price"] > 0].copy()
        df_no_price = df_result[
            df_result["Best Price"] <= 0].copy()

        # ── KPIs ─────────────────────────────────
        st.markdown("<br>",unsafe_allow_html=True)
        k1,k2,k3,k4,k5 = st.columns(5)
        kpi(k1,
            len(df_result),
            "Total Files","#D04A02")
        kpi(k2,
            len(df_priced),
            "Prices Found","#22992E")
        kpi(k3,
            len(df_no_price),
            "No Price","#8C8C8C")
        kpi(k4,
            df_result["Vendor"].nunique(),
            "Vendors","#295477")
        kpi(k5,
            _fmt(df_priced["Best Price"].mean())
            if not df_priced.empty else "—",
            "Avg Quote","#299D8F")

        st.markdown("<br>",unsafe_allow_html=True)

        # ── Extraction Status ─────────────────────
        with st.expander(
                "📋 File Extraction Status",
                expanded=False):
            s_rows = [
                "<table class='comp-table'>"
                "<thead><tr>"
                "<th>File Name</th>"
                "<th>Vendor</th>"
                "<th>Quoted Price</th>"
                "<th>Extracted Price</th>"
                "<th>Used Price</th>"
                "<th>Source</th>"
                "<th>Status</th>"
                "</tr></thead><tbody>"]
            for i,r in df_result.iterrows():
                bg  = ("white" if i%2==0
                       else "#F3F3F3")
                vc  = vendor_color_map.get(
                    r["Vendor"],"#8C8C8C")
                sc  = (
                    "#22992E"
                    if r["Source"]=="extracted"
                    else "#FFB600"
                    if r["Source"]=="catalog"
                    else "#bbb")
                s_rows.append(
                    "<tr style='background:{}'>"
                    "<td style='font-family:"
                    "monospace;font-size:0.78em'>"
                    "{}</td>"
                    "<td>{}</td>"
                    "<td style='font-family:"
                    "monospace'>{}</td>"
                    "<td style='font-family:"
                    "monospace;color:#295477'>"
                    "{}</td>"
                    "<td style='font-family:"
                    "monospace;font-weight:700;"
                    "color:#D04A02'>{}</td>"
                    "<td><span style='background:{};"
                    "color:white;padding:2px 6px;"
                    "border-radius:2px;"
                    "font-size:0.72em;"
                    "font-weight:700'>"
                    "{}</span></td>"
                    "<td style='font-size:0.80em'>"
                    "{}</td>"
                    "</tr>".format(
                        bg,
                        r["File Name"],
                        vendor_pill(
                            r["Vendor"],vc),
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
            s_rows.append("</tbody></table>")
            st.markdown(
                "".join(s_rows),
                unsafe_allow_html=True)

        if df_priced.empty:
            st.warning(
                "No prices found. "
                "Check that File Link URLs are "
                "publicly accessible and files "
                "contain price data.")
            st.stop()

        # ════════════════════════════════════════
        # VENDOR PRICE ANALYSIS
        # ════════════════════════════════════════
        st.markdown("<br>",unsafe_allow_html=True)
        section_title(
            "VENDOR PRICE ANALYSIS",
            "Cheapest to most expensive — "
            "ranked by average quote.")

        v_sum = (
            df_priced.groupby("Vendor")[
                "Best Price"]
            .agg(["mean","sum","min","max","count"])
            .reset_index())
        v_sum.columns = [
            "Vendor","Average","Total",
            "Min","Max","Quotes"]
        v_sum = v_sum.sort_values("Average")
        ov_avg = df_priced["Best Price"].mean()

        # Vendor ranking table
        vt = [
            "<table class='comp-table'>"
            "<thead><tr>"
            "<th>Rank</th><th>Vendor</th>"
            "<th>Quotes</th><th>Avg Quote</th>"
            "<th>Min</th><th>Max</th>"
            "<th>vs Average</th><th>Verdict</th>"
            "</tr></thead><tbody>"]
        for rank,(i,vr) in enumerate(
                v_sum.iterrows(),start=1):
            bg  = ("white" if rank%2==0
                   else "#F3F3F3")
            vc  = vendor_color_map.get(
                vr["Vendor"],"#8C8C8C")
            pct = (round(
                (vr["Average"]-ov_avg)
                /ov_avg*100,1)
                if ov_avg>0 else 0)
            pc  = ("#22992E" if pct<-5
                   else "#E0301E" if pct>5
                   else "#856404")
            pt  = (
                "{}% below".format(abs(pct))
                if pct<0
                else "{}% above".format(abs(pct))
                if pct>0 else "At avg")
            ov  = (
                ("✅ COMPETITIVE","#22992E")
                if pct<-10
                else ("🔴 EXPENSIVE","#E0301E")
                if pct>10
                else ("🟡 AVERAGE","#856404"))
            medal = (
                "🥇" if rank==1
                else "🥈" if rank==2
                else "🥉" if rank==3
                else "{}".format(rank))
            vt.append(
                "<tr style='background:{}'>"
                "<td style='text-align:center;"
                "font-size:1.1em'>{}</td>"
                "<td>{}</td>"
                "<td style='text-align:center;"
                "font-weight:700'>{}</td>"
                "<td style='font-family:monospace;"
                "font-weight:700'>{}</td>"
                "<td style='font-family:monospace;"
                "color:#22992E'>{}</td>"
                "<td style='font-family:monospace;"
                "color:#E0301E'>{}</td>"
                "<td style='color:{}'>{}</td>"
                "<td style='color:{};"
                "font-weight:700;"
                "font-size:0.85em'>{}</td>"
                "</tr>".format(
                    bg,medal,
                    vendor_pill(vr["Vendor"],vc),
                    int(vr["Quotes"]),
                    _fmt(vr["Average"]),
                    _fmt(vr["Min"]),
                    _fmt(vr["Max"]),
                    pc,pt,ov[1],ov[0]))
        vt.append("</tbody></table>")
        st.markdown(
            "".join(vt),
            unsafe_allow_html=True)

        # Bar charts
        st.markdown("<br>",unsafe_allow_html=True)
        ch1,ch2 = st.columns(2)
        with ch1:
            fig_avg = go.Figure(go.Bar(
                x=v_sum["Vendor"],
                y=v_sum["Average"],
                marker_color=[
                    vendor_color_map.get(
                        v,"#8C8C8C")
                    for v in v_sum["Vendor"]],
                marker_line_width=0,
                text=v_sum["Average"].apply(_fmt),
                textposition="outside"))
            fig_avg.add_hline(
                y=ov_avg,
                line_dash="dash",
                line_color="#FFB600",
                line_width=2,
                annotation_text="Avg:{}".format(
                    _fmt(ov_avg)),
                annotation_position="top right")
            fig_avg.update_layout(
                title="Average Quote per Vendor",
                height=340,
                plot_bgcolor=CBG,
                paper_bgcolor=CBG,
                margin=dict(
                    l=5,r=10,t=40,b=10),
                font=CFONT,
                yaxis=dict(
                    showgrid=True,
                    gridcolor="#E0E0E0",
                    zeroline=False),
                xaxis=dict(tickangle=-25),
                bargap=0.35)
            st.plotly_chart(
                fig_avg,
                use_container_width=True)

        with ch2:
            fig_tot = go.Figure(go.Bar(
                x=v_sum["Vendor"],
                y=v_sum["Total"],
                marker_color=[
                    vendor_color_map.get(
                        v,"#8C8C8C")
                    for v in v_sum["Vendor"]],
                marker_line_width=0,
                text=v_sum["Total"].apply(_fmt),
                textposition="outside"))
            fig_tot.update_layout(
                title="Total Quoted Value "
                      "per Vendor",
                height=340,
                plot_bgcolor=CBG,
                paper_bgcolor=CBG,
                margin=dict(
                    l=5,r=10,t=40,b=10),
                font=CFONT,
                yaxis=dict(
                    showgrid=True,
                    gridcolor="#E0E0E0",
                    zeroline=False),
                xaxis=dict(tickangle=-25),
                bargap=0.35)
            st.plotly_chart(
                fig_tot,
                use_container_width=True)

        # ════════════════════════════════════════
        # PER SERVICE ANALYSIS
        # ════════════════════════════════════════
        st.markdown("<br>",unsafe_allow_html=True)
        section_title(
            "PER-SERVICE PRICE COMPARISON",
            "Each service — which vendor "
            "was cheapest vs most expensive.")

        # Explode services
        svc_rows = []
        for _,r in df_priced.iterrows():
            svcs = r.get("Services",[])
            if isinstance(svcs,str):
                svcs = parse_services_from_cell(
                    svcs)
            if not svcs:
                svcs = [r["Category"]
                        if r["Category"]
                        else "General"]
            for svc in svcs:
                svc_rows.append({
                    "Service" : svc,
                    "Vendor"  : r["Vendor"],
                    "Category": r["Category"],
                    "Price"   : r["Best Price"],
                    "File"    : r["File Name"],
                    "Source"  : r["Source"],
                })

        df_svc = pd.DataFrame(svc_rows)

        if df_svc.empty:
            st.info(
                "Add a Comments/Services column "
                "to your catalog for service-level "
                "analysis.")
        else:
            svc_vc = (
                df_svc.groupby("Service")[
                    "Vendor"].nunique())
            multi  = svc_vc[
                svc_vc>1].index.tolist()
            single = svc_vc[
                svc_vc==1].index.tolist()

            # Category filter
            cats7 = ["All"] + sorted(
                df_svc["Category"].unique().tolist())
            cat_sel = st.selectbox(
                "Filter by Category",
                cats7,
                key="tab7_cat_sel")
            if cat_sel != "All":
                df_svc_show = df_svc[
                    df_svc["Category"]==cat_sel]
                multi  = [s for s in multi
                          if s in df_svc_show[
                              "Service"].values]
                single = [s for s in single
                          if s in df_svc_show[
                              "Service"].values]
            else:
                df_svc_show = df_svc

            st.markdown(
                "<div style='display:flex;gap:10px;"
                "margin-bottom:14px'>"
                "<div style='background:#F0FFF4;"
                "border:1px solid #22992E;"
                "border-radius:4px;"
                "padding:8px 14px;"
                "font-size:0.85em'>"
                "✅ <b>{}</b> services with "
                "multiple vendor quotes"
                "</div>"
                "<div style='background:#FFF8E1;"
                "border:1px solid #FFB600;"
                "border-radius:4px;"
                "padding:8px 14px;"
                "font-size:0.85em'>"
                "⚠️ <b>{}</b> services with "
                "single vendor only"
                "</div></div>".format(
                    len(multi),len(single)),
                unsafe_allow_html=True)

            for svc in sorted(
                    df_svc_show[
                        "Service"].unique()):
                d_s = df_svc_show[
                    df_svc_show["Service"]==svc
                ].sort_values("Price")
                n_v   = d_s["Vendor"].nunique()
                mn_p  = d_s["Price"].min()
                mx_p  = d_s["Price"].max()
                avg_p = d_s["Price"].mean()
                sprd  = (round(
                    (mx_p-mn_p)/mn_p*100,1)
                    if mn_p>0 else 0)
                bv    = d_s.loc[
                    d_s["Price"].idxmin(),
                    "Vendor"]

                with st.expander(
                    "{}  ·  {} vendor(s){}  ·  "
                    "best: {} @ {}".format(
                        svc, n_v,
                        "  ·  spread {}%".format(
                            sprd) if n_v>1 else "",
                        bv,_fmt(mn_p)),
                    expanded=False):

                    if n_v > 1:
                        s1,s2,s3 = st.columns(3)
                        wv = d_s.loc[
                            d_s["Price"].idxmax(),
                            "Vendor"]
                        s1.markdown(
                            "<div class='score-card "
                            "green'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#22992E'>"
                            "Cheapest</div>"
                            "<div style='font-size:"
                            "1.6em;font-weight:800;"
                            "color:#22992E'>{}</div>"
                            "<div style='font-size:"
                            "0.78em;color:#555'>"
                            "{}</div></div>".format(
                                _fmt(mn_p),bv),
                            unsafe_allow_html=True)
                        s2.markdown(
                            "<div class='score-card "
                            "yellow'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#856404'>"
                            "Average</div>"
                            "<div style='font-size:"
                            "1.6em;font-weight:800;"
                            "color:#856404'>{}</div>"
                            "<div style='font-size:"
                            "0.78em;color:#555'>"
                            "{} vendors</div>"
                            "</div>".format(
                                _fmt(avg_p),n_v),
                            unsafe_allow_html=True)
                        s3.markdown(
                            "<div class='score-card "
                            "red'>"
                            "<div style='font-size:"
                            "0.68em;font-weight:700;"
                            "text-transform:uppercase;"
                            "color:#E0301E'>"
                            "Most Expensive</div>"
                            "<div style='font-size:"
                            "1.6em;font-weight:800;"
                            "color:#E0301E'>{}</div>"
                            "<div style='font-size:"
                            "0.78em;color:#555'>"
                            "{}</div></div>".format(
                                _fmt(mx_p),wv),
                            unsafe_allow_html=True)
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)

                    # Vendor table
                    all_p = d_s["Price"].tolist()
                    tbl   = [
                        "<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>Vendor</th>"
                        "<th>Price</th>"
                        "<th>vs Avg</th>"
                        "<th>Score</th>"
                        "<th>Verdict</th>"
                        "<th>File</th>"
                        "</tr></thead><tbody>"]
                    for j,(_,vr) in enumerate(
                            d_s.iterrows()):
                        bg  = ("white" if j%2==0
                               else "#F3F3F3")
                        vc  = vendor_color_map.get(
                            vr["Vendor"],"#8C8C8C")
                        p   = vr["Price"]
                        oth = [x for x in all_p
                               if x!=p]
                        ps  = None
                        if p>0 and oth:
                            ps,_,_,_,_ = \
                                price_score(p,oth)
                        sc_c= score_color(ps)
                        vt2,_,vc2=get_verdict(ps)
                        pct = (round(
                            (p-avg_p)/avg_p*100,1)
                            if avg_p>0 else 0)
                        pc  = (
                            "#22992E" if pct<0
                            else "#E0301E" if pct>5
                            else "#856404")
                        pt  = (
                            "{}% below".format(
                                abs(pct))
                            if pct<0
                            else "{}% above".format(
                                abs(pct))
                            if pct>0 else "At avg")
                        tbl.append(
                            "<tr style='"
                            "background:{}'>"
                            "<td>{}</td>"
                            "<td style='font-family:"
                            "monospace;"
                            "font-weight:700'>"
                            "{}</td>"
                            "<td style='color:{}'>"
                            "{}</td>"
                            "<td style='"
                            "text-align:center'>"
                            "<span style='"
                            "font-weight:800;"
                            "color:{}'>{}</span>"
                            "</td>"
                            "<td><span style='"
                            "color:{};"
                            "font-weight:700;"
                            "font-size:0.85em'>"
                            "{}</span></td>"
                            "<td style='font-family:"
                            "monospace;"
                            "font-size:0.78em;'>"
                            "{}</td>"
                            "</tr>".format(
                                bg,
                                vendor_pill(
                                    vr["Vendor"],
                                    vc),
                                _fmt(p),
                                pc,pt,
                                sc_c,
                                "{}/100".format(ps)
                                if ps is not None
                                else "—",
                                vc2,vt2,
                                vr["File"]))
                    tbl.append("</tbody></table>")
                    st.markdown(
                        "".join(tbl),
                        unsafe_allow_html=True)

                    if n_v > 1:
                        st.markdown(
                            "<br>",
                            unsafe_allow_html=True)
                        fig_s = go.Figure(go.Bar(
                            x=d_s["Vendor"],
                            y=d_s["Price"],
                            marker_color=[
                                vendor_color_map.get(
                                    v,"#8C8C8C")
                                for v in
                                d_s["Vendor"]],
                            marker_line_width=0,
                            text=d_s["Price"]
                            .apply(_fmt),
                            textposition="outside"))
                        fig_s.add_hline(
                            y=avg_p,
                            line_dash="dash",
                            line_color="#FFB600",
                            line_width=2,
                            annotation_text=
                            "Avg:{}".format(
                                _fmt(avg_p)),
                            annotation_position=
                            "top right")
                        fig_s.update_layout(
                            height=260,
                            plot_bgcolor=CBG,
                            paper_bgcolor=CBG,
                            margin=dict(
                                l=5,r=10,t=10,b=10),
                            font=CFONT,
                            yaxis=dict(
                                showgrid=True,
                                gridcolor="#E0E0E0",
                                zeroline=False),
                            bargap=0.4)
                        st.plotly_chart(
                            fig_s,
                            use_container_width=True)

        # ── Download ─────────────────────────────
        st.markdown("<br>",unsafe_allow_html=True)
        section_title("DOWNLOAD ANALYSIS")
        dl = df_priced[[
            "Vendor","Category","File Name",
            "Quoted Price","Extracted Price",
            "Best Price","Source","Status"
        ]].sort_values(["Category","Best Price"])
        st.download_button(
            label="📥 Download Analysis CSV",
            data=dl.to_csv(index=False),
            file_name="real_quotes_analysis.csv",
            mime="text/csv",
            type="primary")
