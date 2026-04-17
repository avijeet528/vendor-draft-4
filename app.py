# ============================================================
#  app.py — IT Procurement Intelligence Dashboard
#  Two data buckets: Master Catalog (Excel) + Dummy Data (CSV)
#  PwC Brand | No sidebar | Full-width
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
# COLOURS
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

CHART_SEQ = [C_ORANGE, C_DARK, C_ORANGE_MID,
             C_GREY_DARK, C_ORANGE_DARK, C_GREY,
             C_MID, C_BLACK]

CFONT    = dict(family="Georgia,'Source Sans Pro',Arial",
                size=11, color=C_DARK)
CBG      = C_GREY_BG
DEMO_DIR = "demo_quotes"

# ════════════════════════════════════════════════════════════
# CSS
# ════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+Pro:wght@300;400;600;700&display=swap');
html,body,[class*="css"],div,p,span,td,th,label,button,.stMarkdown{
    font-family:'Source Sans Pro','Helvetica Neue',Arial,sans-serif !important;}
h1,h2,h3{font-family:Georgia,'ITC Charter',serif !important;font-weight:700 !important;}
section[data-testid="stSidebar"]{display:none !important;}
[data-testid="collapsedControl"]{display:none !important;}
#MainMenu{visibility:hidden;}footer{visibility:hidden;}header{visibility:hidden;}
.main .block-container{
    background:#F3F3F3 !important;max-width:100% !important;
    padding:1.2rem 2.5rem !important;}
button[data-baseweb="tab"]{
    font-weight:600 !important;font-size:0.90em !important;color:#7D7D7D !important;}
button[data-baseweb="tab"][aria-selected="true"]{
    color:#D04A02 !important;border-bottom:3px solid #D04A02 !important;
    background:transparent !important;}
.kpi-box{border-radius:4px;padding:20px 12px;text-align:center;
    color:white;border-left:5px solid rgba(255,255,255,0.2);}
.kpi-value{font-size:2.4em;font-weight:700;margin:0;line-height:1.1;
    font-family:Georgia,serif !important;}
.kpi-label{font-size:0.72em;font-weight:700;opacity:0.9;margin-top:6px;
    letter-spacing:1.2px;text-transform:uppercase;}
.sec-head{font-size:0.73em;font-weight:700;letter-spacing:1.4px;
    text-transform:uppercase;color:#D04A02;margin:22px 0 10px;
    border-bottom:2px solid #D04A02;padding-bottom:5px;display:block;}
.comp-table{width:100%;border-collapse:collapse;font-size:0.81em;
    border:1px solid #E0E0E0;}
.comp-table thead tr{background:#2D2D2D;}
.comp-table thead th{padding:10px 12px;text-align:left;font-weight:700;
    font-size:0.78em;letter-spacing:0.5px;text-transform:uppercase;
    color:white !important;border:none;}
.comp-table tbody tr:nth-child(even){background:#F8F8F8;}
.comp-table tbody tr:hover{background:#FAD4C0;}
.comp-table tbody td{padding:9px 12px;border-bottom:1px solid #EBEBEB;
    vertical-align:middle;word-break:break-word;color:#2D2D2D;}
.vbadge{display:inline-block;padding:3px 9px;border-radius:2px;
    color:white;font-size:0.76em;font-weight:700;white-space:nowrap;}
.scard{border-radius:4px;padding:16px;margin-bottom:10px;
    border-left:5px solid #D04A02;background:white;}
.scard-orange{border-color:#D04A02;background:#FFF5F0;}
.scard-dark{border-color:#2D2D2D;background:#F5F5F5;}
.scard-grey{border-color:#7D7D7D;background:#FAFAFA;}
.verdict-good{background:#FFF5F0;border:2px solid #D04A02;
    border-radius:4px;padding:14px 18px;color:#D04A02;font-weight:700;}
.verdict-mid{background:#F5F5F5;border:2px solid #4A4A4A;
    border-radius:4px;padding:14px 18px;color:#4A4A4A;font-weight:700;}
.verdict-bad{background:#F0F0F0;border:2px solid #7D7D7D;
    border-radius:4px;padding:14px 18px;color:#2D2D2D;font-weight:700;}
.chat-header{background:#2D2D2D;color:white;padding:12px 18px;
    border-radius:6px 6px 0 0;}
.chat-outer{background:#F8F8F8;border:1px solid #E0E0E0;border-top:none;
    border-radius:0;padding:14px 14px 6px;
    min-height:320px;max-height:420px;overflow-y:auto;}
.msg-user{background:#D04A02;color:white;border-radius:14px 14px 3px 14px;
    padding:9px 14px;margin:5px 0 5px auto;max-width:74%;
    font-size:0.86em;display:inline-block;float:right;clear:both;}
.msg-bot{background:white;color:#2D2D2D;border:1px solid #E0E0E0;
    border-left:4px solid #D04A02;border-radius:14px 14px 14px 3px;
    padding:9px 14px;margin:5px 0;max-width:84%;font-size:0.86em;
    display:inline-block;float:left;clear:both;}
.chat-wrap{overflow:hidden;margin-bottom:3px;}
.chip{display:inline-block;background:white;border:1.5px solid #D04A02;
    color:#D04A02 !important;border-radius:20px;padding:4px 12px;
    font-size:0.78em;font-weight:600;margin:3px 4px;cursor:pointer;
    white-space:nowrap;}
.filter-bar{background:#2D2D2D;padding:14px 20px;border-radius:4px;
    margin-bottom:18px;border-left:4px solid #D04A02;}
.insight-box{background:#FFF5F0;border-left:4px solid #D04A02;
    border-radius:0 4px 4px 0;padding:11px 16px;margin:10px 0;
    font-size:0.86em;color:#2D2D2D;}
.bucket-header{background:#1A1A1A;color:white;padding:10px 20px;
    border-radius:4px;margin-bottom:6px;
    border-left:6px solid #D04A02;font-size:0.85em;
    font-weight:700;letter-spacing:1px;text-transform:uppercase;}
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
        return "—" if v<=0 else "${:,.2f}".format(v)
    except: return str(val)

def _best_price(text):
    tl = text.lower()
    for kw in TOTAL_KW:
        idx = tl.find(kw)
        if idx==-1: continue
        snip  = text[max(0,idx-20):idx+300]
        hits  = PRICE_RE.findall(snip)
        valid = [h.strip() for h in hits if _parse_num(h)>=50]
        if valid: return max(valid, key=_parse_num)
    all_h = PRICE_RE.findall(text)
    valid = [h.strip() for h in all_h if _parse_num(h)>=100]
    return max(valid, key=_parse_num) if valid else ""

def _text_from_bytes(content, ext):
    text=""; ext=ext.lower().strip(".")
    try:
        if ext=="pdf":
            if not PDF_OK: return ""
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for p in pdf.pages:
                    t=p.extract_text()
                    if t: text+=t+"\n"
        elif ext in ("xlsx","xls"):
            wb=openpyxl.load_workbook(
                io.BytesIO(content),data_only=True,read_only=True)
            rows_t=[]
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    rs="  ".join(str(c) for c in row if c is not None)
                    if rs.strip(): rows_t.append(rs)
            text="\n".join(rows_t); wb.close()
        elif ext=="docx":
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml=z.read("word/document.xml").decode("utf-8",errors="ignore")
                    text=re.sub(r"<[^>]+>"," ",xml)
                    text=re.sub(r"\s{2,}","\n",text)
    except: pass
    return text

def extract_price_from_bytes(content, ext):
    text  = _text_from_bytes(content, ext)
    price = _best_price(text)
    if not price or _parse_num(price)<=0:
        all_n=[h.strip() for h in PRICE_RE.findall(text) if _parse_num(h)>=1000]
        if all_n: price=max(all_n, key=_parse_num)
    return {"price":price,
            "price_num":_parse_num(price) if price else 0.0,
            "text":text[:5000]}

def extract_price_from_file(fp):
    try:
        with open(fp,"rb") as f: c=f.read()
        return extract_price_from_bytes(c, fp.rsplit(".",1)[-1])
    except: return {"price":"","price_num":0.0,"text":""}

# ════════════════════════════════════════════════════════════
# SCORING
# ════════════════════════════════════════════════════════════
def price_score(new_price, hist_prices):
    valid=[p for p in hist_prices if p>0]
    if not valid or new_price<=0:
        return None,"No comparison data",0,0,0
    mn=min(valid); mx=max(valid); avg=sum(valid)/len(valid)
    if mx==mn: return 50,"Same as historical average",avg,mn,mx
    score=round((1-(new_price-mn)/(mx-mn))*100,1)
    score=max(0,min(100,score))
    pct=round((new_price-avg)/avg*100,1)
    lbl=("{}% BELOW average — COMPETITIVE".format(abs(pct))
         if new_price<avg
         else "{}% ABOVE average — REVIEW NEEDED".format(abs(pct)))
    return score,lbl,avg,mn,mx

def score_color(s):
    if s is None: return C_GREY_DARK
    if s>=70:     return C_ORANGE
    if s>=40:     return C_GREY_DARK
    return C_MID

def get_verdict(ps):
    if ps is None:
        return "⚪ No Data","No comparison data.",C_GREY_DARK
    if ps>=70:
        return "✅ COMPETITIVE","Priced competitively.",C_ORANGE
    if ps>=40:
        return "🟡 AVERAGE","Within range. Negotiate.",C_GREY_DARK
    return "🔴 HIGH","Above average. Recommend negotiating.",C_MID

# ════════════════════════════════════════════════════════════
# DATA HELPERS
# ════════════════════════════════════════════════════════════
def _clean_df(df):
    safe=[c for c in df.columns if c!="Services List"]
    df=df[safe].copy()
    for col in df.columns:
        try:    df[col]=df[col].fillna("").apply(lambda x:str(x).strip())
        except: df[col]=""
    mask=(df["Category"].apply(lambda x:x in ["","nan"]) &
          df["Vendor"].apply(lambda x:x in ["","nan"]))
    df=df[~mask].copy()
    df.reset_index(drop=True,inplace=True)
    return df

def _parse_services(v):
    if not v or str(v).strip() in ["","nan","None"]:
        return ["(unspecified)"]
    s=str(v).replace("\\n","\n").replace("\r\n","\n").replace("\r","\n")
    parts=[p.strip() for p in s.split("\n")
           if p.strip() and p.strip()!="nan"]
    if not parts: parts=[p.strip() for p in s.split(";") if p.strip()]
    if not parts and len(s)<300:
        parts=[p.strip() for p in s.split(",") if p.strip()]
    return parts if parts else ["(unspecified)"]

def _explode(df):
    df2=df.copy()
    df2["Services List"]=df2["Comments"].apply(_parse_services)
    dfe=df2.explode("Services List").copy()
    dfe.rename(columns={"Services List":"Service"},inplace=True)
    dfe["Service"]=dfe["Service"].apply(lambda x:str(x).strip())
    dfe=dfe[~dfe["Service"].isin(
        ["","(unspecified)","nan","None"])].reset_index(drop=True)
    return df2,dfe

def _norm_cols(df):
    col_map={}
    for c in df.columns:
        cl=str(c).lower().strip()
        if cl=="category" and "Category" not in col_map:
            col_map["Category"]=c
        elif any(k in cl for k in ["vendor","supplier"]) and "Vendor" not in col_map:
            col_map["Vendor"]=c
        elif ("file name" in cl or cl=="filename") and "File Name" not in col_map:
            col_map["File Name"]=c
        elif any(k in cl for k in ["file link","file url"]) and "File Link" not in col_map:
            col_map["File Link"]=c
        elif any(k in cl for k in ["comment","service","description","scope"]) and "Comments" not in col_map:
            col_map["Comments"]=c
        elif any(k in cl for k in ["price","cost","amount","quoted"]) and "Quoted Price" not in col_map:
            col_map["Quoted Price"]=c
    df.rename(columns={v:k for k,v in col_map.items()},inplace=True)
    for req in ["Category","Vendor","File Name","Comments"]:
        if req not in df.columns: df[req]=""
    keep=["Category","Vendor","File Name","Comments"]
    for e in ["File Link","Quoted Price"]:
        if e in df.columns: keep.append(e)
    return df[[c for c in keep if c in df.columns]].copy()

# ════════════════════════════════════════════════════════════
# LOAD MASTER CATALOG (Excel — no prices extractable)
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_master_catalog():
    XLS="Master Catalog.xlsx"
    if not os.path.exists(XLS): return None,None
    try:
        raw=pd.read_excel(XLS,engine="openpyxl",header=None)
        hr=0
        for i,row in raw.iterrows():
            vals=[str(v).strip().lower()
                  for v in row.values if pd.notna(v)]
            if (any("category" in v for v in vals)
                    and any("vendor" in v for v in vals)):
                hr=i; break
        df=pd.read_excel(XLS,engine="openpyxl",header=hr)
        df.columns=[str(c).strip() for c in df.columns]
        # Extract embedded hyperlinks
        hmap={}
        try:
            wb=openpyxl.load_workbook(XLS)
            ws=wb.active; fc=None; hr2=None
            for row in ws.iter_rows():
                for cell in row:
                    if (cell.value and
                            str(cell.value).strip().lower()
                            =="file name"):
                        fc=cell.column; hr2=cell.row; break
                if fc: break
            if fc and hr2:
                for row in ws.iter_rows(min_row=hr2+1):
                    for cell in row:
                        if (cell.column==fc
                                and cell.value
                                and cell.hyperlink):
                            hmap[str(cell.value).strip()]=str(
                                cell.hyperlink.target).strip()
            wb.close()
        except: pass
        df=_norm_cols(df); df=_clean_df(df)
        df["Hyperlink"]=df["File Name"].map(hmap).fillna("")
        df,dfe=_explode(df)
        return df,dfe
    except Exception as e:
        st.warning("Master catalog error: {}".format(e))
        return None,None

# ════════════════════════════════════════════════════════════
# LOAD DUMMY DATA (CSV — prices available, fully analysable)
# ════════════════════════════════════════════════════════════
@st.cache_data
def load_dummy_data():
    CSV="dummy_catalog.csv"
    if not os.path.exists(CSV): return None,None
    try:
        df=pd.read_csv(CSV)
        df.columns=[str(c).strip() for c in df.columns]
        df=_norm_cols(df); df=_clean_df(df)
        df["Hyperlink"]=""
        df,dfe=_explode(df)
        return df,dfe
    except Exception as e:
        st.warning("Dummy data error: {}".format(e))
        return None,None

# ════════════════════════════════════════════════════════════
# DUMMY DATA GENERATOR
# Creates dummy_catalog.csv if it doesn't exist
# ════════════════════════════════════════════════════════════
def ensure_dummy_data():
    if os.path.exists("dummy_catalog.csv"): return
    import random
    random.seed(42)
    vendors=["NTT Data","Dimension Data","Telstra",
             "Optus","Vocus","Datacom"]
    cats={
        "Cybersecurity":[
            "Endpoint Protection","SIEM Monitoring",
            "Privileged Access Mgmt","Security Awareness",
            "Network Access Control"],
        "Network & Telecom":[
            "Cisco Catalyst 9200-L","Cisco Catalyst 9400",
            "Palo Alto NGFW","SD-WAN Solution",
            "Cisco Meraki MX"],
        "Hosting":[
            "VMware vSphere","NetApp Storage",
            "Oracle DB License","Colocation Build",
            "Backup & Recovery"],
        "M365 & Power Platform":[
            "M365 E3 License","M365 E5 License",
            "Power BI Premium","Teams Rooms"],
    }
    rows=[]
    for cat,svcs in cats.items():
        for svc in svcs:
            n_vendors=random.randint(2,4)
            chosen=random.sample(vendors,n_vendors)
            for v in chosen:
                base=random.uniform(20000,200000)
                price=round(base*random.uniform(0.85,1.15),2)
                fname="{}_{}_{}.pdf".format(
                    v.replace(" ","_"),
                    svc.replace(" ","_")[:15],
                    random.randint(1000,9999))
                rows.append({
                    "Category":cat,
                    "Vendor":v,
                    "File Name":fname,
                    "Comments":svc,
                    "Quoted Price":price})
    pd.DataFrame(rows).to_csv(
        "dummy_catalog.csv",index=False)

ensure_dummy_data()

# ════════════════════════════════════════════════════════════
# SUBCATEGORY
# ════════════════════════════════════════════════════════════
def infer_sub(cat,comments,fname):
    txt=(str(comments)+" "+str(fname)).lower()
    c=str(cat).lower().strip()
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
        if any(k in txt for k in ["vmware","vcf"]):   return "VMware"
        if "oracle" in txt:                             return "Oracle DB"
        if "netapp" in txt:                             return "NetApp"
        if any(k in txt for k in ["colo","colocation"]):return "Colocation"
        return "General Hosting"
    if "m365" in c:    return "M365 Licensing"
    if "idam" in c:    return "Identity Migration"
    if "snow" in c:    return "ServiceNow ITSM"
    if "summary" in c: return "Reporting"
    return str(cat).strip().title()

# ════════════════════════════════════════════════════════════
# CHATBOT
# ════════════════════════════════════════════════════════════
def chatbot_response(user_msg, df_master, df_exploded,
                     uploaded_file_bytes=None,
                     uploaded_file_name=None):
    msg=user_msg.lower().strip()
    if df_master is None or df_exploded is None:
        return {"type":"text","text":"No catalog loaded."}
    all_services=sorted(df_exploded["Service"].unique().tolist())
    all_vendors =sorted(df_master["Vendor"].unique().tolist())
    all_cats    =sorted(df_master["Category"].unique().tolist())

    # File upload
    if uploaded_file_bytes is not None and uploaded_file_name:
        ext=uploaded_file_name.rsplit(".",1)[-1].lower()
        result=extract_price_from_bytes(uploaded_file_bytes,ext)
        price=result["price_num"]
        st.session_state["tab2_upload_price"] =price
        st.session_state["tab2_upload_fname"] =uploaded_file_name
        st.session_state["tab2_file_bytes"]   =uploaded_file_bytes
        st.session_state["tab2_file_ext"]     =ext
        st.session_state["chat_redirect_upload"]=True
        hist=[]
        text_lower=result["text"].lower()
        matched_svcs=[s for s in all_services
                      if any(w in text_lower
                             for w in s.lower().split()
                             if len(w)>3)][:5]
        if matched_svcs:
            for s in matched_svcs:
                for _,r in df_exploded[
                        df_exploded["Service"]==s].iterrows():
                    qp=_parse_num(str(r.get("Quoted Price","")).strip())
                    if qp>0: hist.append(qp)
        verdict_txt=""
        if hist and price>0:
            ps,lbl,avg,mn,mx=price_score(price,hist)
            verdict_txt=("\n\n📊 vs {} similar: "
                         "Avg {} | **{}**".format(
                             len(hist),_fmt(avg),lbl))
        return {
            "type":"redirect_upload",
            "text":("📄 **{}** — price **{}**{}\n\n"
                    "🔄 Redirecting to Upload & Score…".format(
                        uploaded_file_name,
                        _fmt(price) if price>0
                        else "not found",
                        verdict_txt)),
            "price":price,
        }

    # Greeting
    if any(w in msg for w in ["hello","hi","hey"]):
        return {"type":"text","text":(
            "👋 Hello! I know the full catalog.\n\n"
            "**{} quotes · {} vendors · {} services**\n\n"
            "Ask:\n• *Who quoted Cisco Catalyst?*\n"
            "• *Compare Palo Alto prices*\n"
            "• *Cheapest Cybersecurity vendor?*\n"
            "• Upload a quote file for instant scoring".format(
                len(df_master),
                df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique()))}

    # Summary
    if any(w in msg for w in [
            "summary","overview","how many","total","catalog"]):
        lines=[]
        for cat in all_cats:
            dc=df_master[df_master["Category"]==cat]
            ns=df_exploded[df_exploded["Category"]==cat
                           ]["Service"].nunique()
            prices=[]
            if "Quoted Price" in dc.columns:
                for p in dc["Quoted Price"]:
                    n=_parse_num(str(p))
                    if n>0: prices.append(n)
            p_txt=("avg {}".format(
                _fmt(sum(prices)/len(prices)))
                    if prices else "no prices")
            lines.append("• **{}**: {} quotes · "
                         "{} vendors · {} svcs · {}".format(
                             cat,len(dc),
                             dc["Vendor"].nunique(),ns,p_txt))
        return {"type":"text","text":(
            "📊 **Catalog:**\n{} quotes · {} vendors · "
            "{} services · {} cats\n\n{}".format(
                len(df_master),df_master["Vendor"].nunique(),
                df_exploded["Service"].nunique(),
                df_master["Category"].nunique(),
                "\n".join(lines)))}

    # Price analysis
    if any(w in msg for w in [
            "cheapest","expensive","best price",
            "competitive","price analysis","which vendor"]):
        matched_cat=next(
            (c for c in all_cats
             if c.lower() in msg
             or any(w in msg for w in c.lower().split()
                    if len(w)>3)), None)
        scope=(df_master[df_master["Category"]==matched_cat]
               if matched_cat else df_master)
        va={}
        for v in scope["Vendor"].unique():
            dv=scope[scope["Vendor"]==v]
            ps=[]
            if "Quoted Price" in dv.columns:
                for p in dv["Quoted Price"]:
                    n=_parse_num(str(p))
                    if n>0: ps.append(n)
            if ps: va[v]=sum(ps)/len(ps)
        if va:
            bv=min(va,key=va.get)
            wv=max(va,key=va.get)
            oa=sum(va.values())/len(va)
            lines=[]
            for v,avg in sorted(va.items(),key=lambda x:x[1]):
                pct=round((avg-oa)/oa*100,1) if oa>0 else 0
                tag=("🟠 Cheapest" if v==bv
                     else "⚫ Most exp." if v==wv
                     else "⚪ Mid")
                lines.append("**{}**: {} ({:+.1f}%) {}".format(
                    v,_fmt(avg),pct,tag))
            return {
                "type":"price_analysis",
                "text":("💰 **Price Analysis{}**\n\nAvg: {}\n\n"
                        "{}\n\n✅ **{}** cheapest at **{}**".format(
                            " — "+matched_cat
                            if matched_cat else "",
                            _fmt(oa),"\n".join(lines),
                            bv,_fmt(va[bv]))),
                "vendor_avgs":va,"avg":oa,"best_vendor":bv,
            }
        return {"type":"text",
                "text":"No price data available."}

    matched_vendor=next(
        (v for v in all_vendors if v.lower() in msg),None)
    matched_services=[
        svc for svc in all_services
        if any(w in msg for w in svc.lower().split()
               if len(w)>3)
        or svc.lower() in msg][:8]
    matched_cat=next(
        (c for c in all_cats
         if c.lower() in msg
         or any(w in msg for w in c.lower().split()
                if len(w)>4)),None)

    # Vendor + service
    if matched_vendor and matched_services:
        svc=matched_services[0]
        d=df_exploded[(df_exploded["Vendor"]==matched_vendor)
                      &(df_exploded["Service"]==svc)]
        if not d.empty:
            files=d["File Name"].unique().tolist()
            prices=[]
            for f in files:
                rows=df_master[df_master["File Name"]==f]
                if len(rows)>0 and "Quoted Price" in rows.columns:
                    qp=_parse_num(str(rows["Quoted Price"].values[0]))
                    if qp>0: prices.append(qp)
            p_txt=("Price: **{}**".format(_fmt(prices[0]))
                    if prices else "No price on record.")
            all_p=[_parse_num(str(r.get("Quoted Price","")).strip())
                   for _,r in df_exploded[
                       df_exploded["Service"]==svc].iterrows()
                   if _parse_num(str(r.get("Quoted Price","")).strip())>0]
            v_txt=""
            if prices and all_p:
                ps2,lbl,avg,mn,mx=price_score(prices[0],all_p)
                v_txt="\n📊 vs market: **{}**".format(lbl)
            return {"type":"vendor_service","text":(
                "✅ **{}** quoted **{}**.\n\n{}{}\n📄 {}".format(
                    matched_vendor,svc,p_txt,v_txt,
                    ", ".join(files[:3])))}
        return {"type":"text","text":(
            "❌ **{}** has not quoted **{}**.".format(
                matched_vendor,svc))}

    # Compare prices
    if matched_services and any(w in msg for w in [
            "compare","price","cost","expensive",
            "cheap","competitive","vs"]):
        svc=matched_services[0]
        d=df_exploded[df_exploded["Service"]==svc
                      ].drop_duplicates(
            subset=["Vendor","File Name"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No quotes for **{}**.".format(svc)}
        vp={}
        for _,r in d.iterrows():
            v=r["Vendor"]
            qp=_parse_num(str(r.get("Quoted Price","")).strip())
            ck="px_{}".format(str(r.get("File Name","")).strip())
            ca=st.session_state.get(ck)
            ep=ca["price_num"] if ca else 0.0
            p=ep if ep>0 else qp
            if p>0: vp[v]=min(vp.get(v,p),p)
        if not vp:
            return {"type":"text","text":(
                "📋 **{}** quoted by: {}\nNo price data.".format(
                    svc,", ".join(d["Vendor"].unique().tolist())))}
        avg_p=sum(vp.values())/len(vp)
        bv=min(vp,key=vp.get)
        spread=round((max(vp.values())-min(vp.values()))
                     /min(vp.values())*100,1) \
            if min(vp.values())>0 else 0
        lines=[]
        for v,p in sorted(vp.items(),key=lambda x:x[1]):
            pct=round((p-avg_p)/avg_p*100,1) if avg_p>0 else 0
            tag=("🟠 Best" if v==bv
                 else "⚫ Most exp."
                 if p==max(vp.values()) else "⚪ Mid")
            lines.append("**{}**: {} ({:+.1f}%) {}".format(
                v,_fmt(p),pct,tag))
        return {
            "type":"comparison",
            "text":("📊 **{}**\n\n{}\n\nSpread: **{}%** · "
                    "Best: **{}** at **{}**".format(
                        svc,"\n".join(lines),
                        spread,bv,_fmt(min(vp.values())))),
            "service":svc,"vendor_prices":vp,
            "avg":avg_p,"best_vendor":bv,
        }

    # Who quoted
    if matched_services and any(w in msg for w in [
            "who","vendor","quoted","available"]):
        svc=matched_services[0]
        d=df_exploded[df_exploded["Service"]==svc
                      ].drop_duplicates(subset=["Vendor"])
        if d.empty:
            return {"type":"text",
                    "text":"❌ No vendor quoted **{}**.".format(svc)}
        vendors=d["Vendor"].unique().tolist()
        return {"type":"who_quoted","text":(
            "✅ **{}** vendor(s) for **{}**:\n\n{}\n\n"
            "📄 {} files".format(
                len(vendors),svc,
                "\n".join(["• **{}**".format(v) for v in vendors]),
                df_exploded[df_exploded["Service"]==svc
                            ]["File Name"].nunique()))}

    # Vendor profile
    if matched_vendor:
        d=df_exploded[df_exploded["Vendor"]==matched_vendor]
        svcs=sorted(d["Service"].unique().tolist())
        cats=sorted(d["Category"].unique().tolist())
        nq=len(df_master[df_master["Vendor"]==matched_vendor])
        prices=[]
        dv=df_master[df_master["Vendor"]==matched_vendor]
        if "Quoted Price" in dv.columns:
            for p in dv["Quoted Price"]:
                n=_parse_num(str(p))
                if n>0: prices.append(n)
        p_sum=("\n\n💰 avg {} · min {} · max {}".format(
            _fmt(sum(prices)/len(prices)),
            _fmt(min(prices)),_fmt(max(prices)))
                if prices else "")
        return {"type":"vendor_profile","text":(
            "🏢 **{}**\n\n📂 {}\n📄 {} quotes{}\n"
            "🛠 {} services:\n{}".format(
                matched_vendor,", ".join(cats),nq,p_sum,
                len(svcs),
                "\n".join(["• {}".format(s)
                            for s in svcs[:12]])
                +("\n…+{} more".format(len(svcs)-12)
                  if len(svcs)>12 else "")))}

    # Category
    if matched_cat:
        dc=df_master[df_master["Category"]==matched_cat]
        prices=[]
        if "Quoted Price" in dc.columns:
            for p in dc["Quoted Price"]:
                n=_parse_num(str(p))
                if n>0: prices.append(n)
        p_txt=("avg {} · min {} · max {}".format(
            _fmt(sum(prices)/len(prices)),
            _fmt(min(prices)),_fmt(max(prices)))
                if prices else "no price data")
        return {"type":"text","text":(
            "📂 **{}**\n\n📄 {} quotes · 🏢 {} vendors · "
            "🛠 {} services\n💰 {}\n\nVendors: {}".format(
                matched_cat,len(dc),dc["Vendor"].nunique(),
                df_exploded[df_exploded["Category"]==matched_cat
                            ]["Service"].nunique(),
                p_txt,
                ", ".join(sorted(dc["Vendor"].unique().tolist()))))}

    # Service info
    if matched_services:
        svc=matched_services[0]
        d=df_exploded[df_exploded["Service"]==svc]
        if not d.empty:
            vendors=d["Vendor"].unique().tolist()
            prices=[]
            for _,r in d.drop_duplicates(
                    subset=["File Name"]).iterrows():
                qp=_parse_num(str(r.get("Quoted Price","")).strip())
                if qp>0: prices.append(qp)
            p_txt=("\n\n💰 {} – {} (avg {})".format(
                _fmt(min(prices)),_fmt(max(prices)),
                _fmt(sum(prices)/len(prices)))
                    if prices else "")
            return {"type":"service_info","text":(
                "📋 **{}**\n\n🏢 {} vendor(s): {}\n"
                "📄 {} files{}\n\n"
                "💡 Ask *'compare {} prices'*".format(
                    svc,len(vendors),
                    ", ".join(vendors[:5]),
                    d["File Name"].nunique(),
                    p_txt,svc))}

    sugg=[]
    if matched_services:
        sugg.append("*'Compare {}?'*".format(matched_services[0]))
    if matched_vendor:
        sugg.append("*'Profile of {}?'*".format(matched_vendor))
    fb=("Couldn't find specific data. "
        "Try a vendor, service, or category.")
    if sugg: fb+="\n\n💡 Try: "+" or ".join(sugg)
    return {"type":"text","text":fb}

def render_chat_chart(resp):
    if resp is None: return
    vp=(resp.get("vendor_prices")
        or resp.get("vendor_avgs"))
    avg=resp.get("avg",0); bv=resp.get("best_vendor","")
    if not vp: return
    sv=sorted(vp.items(),key=lambda x:x[1])
    bc=[C_ORANGE if v==bv
        else C_DARK if p==max(vp.values())
        else C_GREY_DARK for v,p in sv]
    fig=go.Figure(go.Bar(
        x=[v for v,_ in sv],y=[p for _,p in sv],
        marker_color=bc,marker_line_width=0,
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
def vpill(v,color=None):
    return ("<span class='vbadge' style='background:{}'>"
            "{}</span>".format(color or C_DARK,v))

def sec(txt,caption=""):
    st.markdown(
        "<span class='sec-head'>{}</span>".format(txt),
        unsafe_allow_html=True)
    if caption:
        st.markdown(
            "<div style='font-size:0.81em;color:#7D7D7D;"
            "margin:-6px 0 10px'>{}</div>".format(caption),
            unsafe_allow_html=True)

def kpi_box(col,val,lbl,bg):
    col.markdown(
        "<div class='kpi-box' style='background:{}'>"
        "<div class='kpi-value'>{}</div>"
        "<div class='kpi-label'>{}</div></div>".format(bg,val,lbl),
        unsafe_allow_html=True)

def mini_kpi(col,val,lbl,bg,icon=""):
    col.markdown(
        "<div class='kpi-box' style='background:{};"
        "min-height:82px'>"
        "<div style='font-size:1.3em;margin-bottom:2px'>{}</div>"
        "<div class='kpi-value' style='font-size:1.3em'>{}</div>"
        "<div class='kpi-label'>{}</div></div>".format(bg,icon,val,lbl),
        unsafe_allow_html=True)

def insight(text):
    st.markdown(
        "<div class='insight-box'>💡 {}</div>".format(text),
        unsafe_allow_html=True)

def resolve_url(row):
    for col in ["Hyperlink","File Link"]:
        val=str(row.get(col,"")).strip()
        if val and val not in ["","nan"] and val.startswith("http"):
            return val
    fname=str(row.get("File Name","")).strip()
    if fname:
        local=os.path.join(DEMO_DIR,fname)
        if os.path.exists(local): return local
    return ""

def pwc_bar(fig_obj,title="",height=320):
    fig_obj.update_layout(
        height=height,plot_bgcolor=CBG,paper_bgcolor=CBG,
        margin=dict(l=5,r=10,t=36 if title else 16,b=10),
        font=CFONT,
        yaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,
                   zeroline=False),
        bargap=0.35,showlegend=False)
    if title:
        fig_obj.update_layout(title=dict(
            text=title,
            font=dict(size=12,color=C_DARK,
                      family="Georgia,serif"),
            x=0,xanchor="left"))

# ════════════════════════════════════════════════════════════
# CATALOG OVERVIEW TAB  (reusable for both data sources)
# ════════════════════════════════════════════════════════════
def render_catalog_overview(df_master, df_exploded,
                             label="Master Catalog"):
    df_ov=df_master.copy()
    df_ov["Subcategory"]=df_ov.apply(
        lambda r:infer_sub(r.get("Category",""),
                           r.get("Comments",""),
                           r.get("File Name","")),axis=1)
    all_cats_ov=sorted([
        c for c in df_ov["Category"].unique()
        if str(c).strip() not in ["","nan"]])

    k0a,k0b,k0c,k0d,k0e=st.columns(5)
    mini_kpi(k0a,len(df_ov),"Total Quotations",C_ORANGE,"📄")
    mini_kpi(k0b,df_ov["Vendor"].nunique(),"Unique Vendors",C_DARK,"🏢")
    mini_kpi(k0c,df_ov["Category"].nunique(),"Categories",C_MID,"📂")
    mini_kpi(k0d,df_ov["Subcategory"].nunique(),"Subcategories",C_GREY_DARK,"🏷️")
    mini_kpi(k0e,df_exploded["Service"].nunique() if df_exploded is not None else "—",
             "Unique Services",C_BLACK,"🛠")
    st.markdown("<br>",unsafe_allow_html=True)

    # Category cards
    sec("CATEGORIES AT A GLANCE")
    cat_stats=[]
    for cat in all_cats_ov:
        dc=df_ov[df_ov["Category"]==cat]
        ns=(df_exploded[df_exploded["Category"]==cat
                        ]["Service"].nunique()
             if df_exploded is not None else 0)
        cat_stats.append({
            "Category":cat,"Quotations":len(dc),
            "Vendors":dc["Vendor"].nunique(),
            "Subcategories":dc["Subcategory"].nunique(),
            "Services":ns})
    cs_df=pd.DataFrame(cat_stats).sort_values(
        "Quotations",ascending=False)
    CAT_ICONS={
        "Cybersecurity":"🛡️","Network & Telecom":"🌐",
        "Hosting":"🖥️","M365 & Power Platform":"☁️",
        "IdAM":"🔑","Service Management (SNow)":"⚙️",
        "Summary & Reporting":"📊"}

    rows3=[cs_df.iloc[i:i+3] for i in range(0,len(cs_df),3)]
    for chunk in rows3:
        cols=st.columns(len(chunk),gap="medium")
        for ci,(_,rs) in enumerate(chunk.iterrows()):
            cn=rs["Category"]
            icon=CAT_ICONS.get(cn,"📁")
            cidx=all_cats_ov.index(cn) if cn in all_cats_ov else 0
            color=CHART_SEQ[cidx%len(CHART_SEQ)]
            cols[ci].markdown(
                "<div style='background:white;"
                "border:1px solid #E0E0E0;border-radius:6px;"
                "padding:18px 16px;border-top:4px solid {}'>"
                "<div style='font-size:1.5em;margin-bottom:6px'>{}</div>"
                "<div style='font-size:0.92em;font-weight:700;"
                "color:#2D2D2D;margin-bottom:12px;"
                "font-family:Georgia,serif'>{}</div>"
                "<div style='display:flex;gap:6px;flex-wrap:wrap'>"
                "<span style='background:#F3F3F3;border-radius:3px;"
                "padding:3px 8px;font-size:0.73em;font-weight:700;"
                "color:#2D2D2D'>📄 {} quotes</span>"
                "<span style='background:#F3F3F3;border-radius:3px;"
                "padding:3px 8px;font-size:0.73em;font-weight:700;"
                "color:#4A4A4A'>🏢 {} vendors</span>"
                "<span style='background:#F3F3F3;border-radius:3px;"
                "padding:3px 8px;font-size:0.73em;font-weight:700;"
                "color:#7D7D7D'>🛠 {} services</span>"
                "</div></div>".format(
                    color,icon,cn,
                    rs["Quotations"],rs["Vendors"],rs["Services"]),
                unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)

    # Charts
    sec("DISTRIBUTION")
    ch1,ch2=st.columns(2,gap="large")
    with ch1:
        fig1=go.Figure(go.Bar(
            x=cs_df["Category"],y=cs_df["Quotations"],
            marker_color=C_ORANGE,marker_line_width=0,
            text=cs_df["Quotations"],textposition="outside"))
        pwc_bar(fig1,"Quotations per Category")
        fig1.update_xaxes(tickangle=-30,tickfont=dict(size=9.5))
        st.plotly_chart(fig1,use_container_width=True)
    with ch2:
        fig2=go.Figure(go.Bar(
            x=cs_df["Category"],y=cs_df["Vendors"],
            marker_color=C_DARK,marker_line_width=0,
            text=cs_df["Vendors"],textposition="outside"))
        pwc_bar(fig2,"Vendors per Category")
        fig2.update_xaxes(tickangle=-30,tickfont=dict(size=9.5))
        st.plotly_chart(fig2,use_container_width=True)

    st.markdown("<br>",unsafe_allow_html=True)
    sec("CATALOG COMPOSITION")
    dl,dr=st.columns([1,2],gap="large")
    with dl:
        fig_d=px.pie(cs_df,values="Quotations",names="Category",
                      hole=0.55,color_discrete_sequence=CHART_SEQ)
        fig_d.update_traces(textposition="outside",
                             textinfo="percent+label",
                             textfont_size=10,
                             pull=[0.03]*len(cs_df))
        fig_d.update_layout(height=360,
                             margin=dict(l=10,r=10,t=10,b=10),
                             paper_bgcolor=CBG,font=CFONT,
                             showlegend=False)
        st.plotly_chart(fig_d,use_container_width=True)
    with dr:
        tbl=["<table class='comp-table'><thead><tr>"
             "<th>Category</th>"
             "<th style='text-align:center'>Quotes</th>"
             "<th style='text-align:center'>Vendors</th>"
             "<th style='text-align:center'>Services</th>"
             "<th style='text-align:center'>Subcats</th>"
             "</tr></thead><tbody>"]
        for _,rs in cs_df.iterrows():
            cn=rs["Category"]
            cidx=all_cats_ov.index(cn) if cn in all_cats_ov else 0
            color=CHART_SEQ[cidx%len(CHART_SEQ)]
            icon=CAT_ICONS.get(cn,"📁")
            tbl.append(
                "<tr><td><span style='border-left:4px solid {};"
                "padding-left:8px;font-weight:600'>"
                "{} {}</span></td>"
                "<td style='text-align:center;font-weight:700;"
                "color:#D04A02'>{}</td>"
                "<td style='text-align:center;font-weight:700;"
                "color:#4A4A4A'>{}</td>"
                "<td style='text-align:center;font-weight:700;"
                "color:#7D7D7D'>{}</td>"
                "<td style='text-align:center;font-weight:700;"
                "color:#2D2D2D'>{}</td></tr>".format(
                    color,icon,cn,
                    rs["Quotations"],rs["Vendors"],
                    rs["Services"],rs["Subcategories"]))
        tbl.append("</tbody></table>")
        st.markdown("".join(tbl),unsafe_allow_html=True)

    # Drill-down
    st.markdown("<br>",unsafe_allow_html=True)
    sec("DRILL-DOWN BY CATEGORY")
    ordered=(["Cybersecurity"]
             +[c for c in all_cats_ov if c!="Cybersecurity"])
    cat_tabs=st.tabs([
        "🛡️ {}".format(c) if c=="Cybersecurity" else c
        for c in ordered])
    for tidx,cn in enumerate(ordered):
        with cat_tabs[tidx]:
            dc=df_ov[df_ov["Category"]==cn].copy()
            if dc.empty: st.info("No data."); continue
            ck1,ck2,ck3,ck4=st.columns(4)
            nsv=(df_exploded[df_exploded["Category"]==cn
                             ]["Service"].nunique()
                  if df_exploded is not None else 0)
            mini_kpi(ck1,len(dc),"Quotations",C_ORANGE,"📄")
            mini_kpi(ck2,dc["Vendor"].nunique(),"Vendors",C_DARK,"🏢")
            mini_kpi(ck3,nsv,"Services",C_MID,"🛠")
            mini_kpi(ck4,dc["Subcategory"].nunique(),"Subcats",C_GREY_DARK,"🏷️")
            st.markdown("<br>",unsafe_allow_html=True)

            ss=(dc.groupby("Subcategory")
                .agg(Quotations=("File Name","count"),
                     Vendors=("Vendor","nunique"))
                .reset_index()
                .sort_values("Quotations",ascending=False))
            sc1,sc2=st.columns([1,1],gap="medium")
            with sc1:
                st.markdown(
                    "<div style='font-size:0.77em;font-weight:700;"
                    "text-transform:uppercase;color:#2D2D2D;"
                    "margin-bottom:8px'>Subcategories</div>",
                    unsafe_allow_html=True)
                fs=go.Figure(go.Bar(
                    x=ss["Quotations"],y=ss["Subcategory"],
                    orientation="h",marker_color=C_ORANGE,
                    marker_line_width=0,
                    text=ss["Quotations"],textposition="outside"))
                fs.update_layout(
                    height=max(200,len(ss)*36),
                    plot_bgcolor=CBG,paper_bgcolor=CBG,
                    margin=dict(l=5,r=40,t=8,b=8),font=CFONT,
                    xaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,
                               zeroline=False),
                    yaxis=dict(autorange="reversed",
                               tickfont=dict(size=9.5)),
                    bargap=0.3)
                st.plotly_chart(fs,use_container_width=True)
            with sc2:
                st.markdown(
                    "<div style='font-size:0.77em;font-weight:700;"
                    "text-transform:uppercase;color:#2D2D2D;"
                    "margin-bottom:8px'>Vendors</div>",
                    unsafe_allow_html=True)
                vs2=(dc.groupby("Vendor")
                     .agg(Quotations=("File Name","count"))
                     .reset_index()
                     .sort_values("Quotations",ascending=False))
                fv2=go.Figure(go.Bar(
                    x=vs2["Quotations"],y=vs2["Vendor"],
                    orientation="h",marker_color=C_DARK,
                    marker_line_width=0,
                    text=vs2["Quotations"],textposition="outside"))
                fv2.update_layout(
                    height=max(200,len(vs2)*36),
                    plot_bgcolor=CBG,paper_bgcolor=CBG,
                    margin=dict(l=5,r=40,t=8,b=8),font=CFONT,
                    xaxis=dict(showgrid=True,gridcolor=C_GREY_LITE,
                               zeroline=False),
                    yaxis=dict(autorange="reversed",
                               tickfont=dict(size=9.5)),
                    bargap=0.3)
                st.plotly_chart(fv2,use_container_width=True)

            ft=["<table class='comp-table'><thead><tr>"
                "<th>File Name</th><th>Vendor</th>"
                "<th>Subcategory</th><th>Services</th>"
                "</tr></thead><tbody>"]
            for fi,(_,fr) in enumerate(
                    dc.sort_values("Subcategory").iterrows()):
                bg="white" if fi%2==0 else "#F8F8F8"
                vc=vendor_color_map.get(fr["Vendor"],C_DARK)
                cmt=str(fr.get("Comments","")
                         ).replace("\n"," · ")[:80]
                ft.append(
                    "<tr style='background:{}'>"
                    "<td style='font-family:monospace;"
                    "font-size:0.77em;word-break:break-all'>"
                    "{}</td><td>{}</td>"
                    "<td style='color:#7D7D7D;font-size:0.81em'>"
                    "{}</td>"
                    "<td style='font-size:0.79em'>{}</td>"
                    "</tr>".format(
                        bg,fr.get("File Name",""),
                        vpill(fr["Vendor"],vc),
                        fr["Subcategory"],cmt))
            ft.append("</tbody></table>")
            st.markdown("".join(ft),unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════
# BROWSE & VERDICT TAB  (reusable)
# has_prices=True means dummy data with real prices
# ════════════════════════════════════════════════════════════
def render_browse_verdict(df_master, df_exploded,
                           vcmap, label="",
                           has_prices=False,
                           chat_key_suffix=""):
    # Top filter bar
    st.markdown("<div class='filter-bar'>",
                 unsafe_allow_html=True)
    fb1,fb2,fb3=st.columns(3)
    all_cats_f=(["All"]+sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["","nan"]]))
    with fb1:
        st.markdown(
            "<p style='color:#D04A02;font-size:0.77em;"
            "font-weight:700;margin-bottom:4px;"
            "letter-spacing:0.6px;text-transform:uppercase'>"
            "📂 CATEGORY</p>",unsafe_allow_html=True)
        sel_cat=st.selectbox(
            "cat",all_cats_f,
            label_visibility="collapsed",
            key="bv_cat_{}".format(chat_key_suffix))
    vpool=(df_master if sel_cat=="All"
            else df_master[df_master["Category"]==sel_cat])
    all_vend_f=(["All"]+sorted([
        v for v in vpool["Vendor"].unique()
        if str(v).strip() not in ["","nan"]]))
    with fb2:
        st.markdown(
            "<p style='color:#D04A02;font-size:0.77em;"
            "font-weight:700;margin-bottom:4px;"
            "letter-spacing:0.6px;text-transform:uppercase'>"
            "🏢 VENDOR</p>",unsafe_allow_html=True)
        sel_vendor=st.selectbox(
            "ven",all_vend_f,
            label_visibility="collapsed",
            key="bv_ven_{}".format(chat_key_suffix))
    with fb3:
        st.markdown(
            "<p style='color:#D04A02;font-size:0.77em;"
            "font-weight:700;margin-bottom:4px;"
            "letter-spacing:0.6px;text-transform:uppercase'>"
            "🔍 SEARCH SERVICE</p>",unsafe_allow_html=True)
        svc_filter=st.text_input(
            "svc",placeholder="Type to filter…",
            label_visibility="collapsed",
            key="bv_svc_{}".format(chat_key_suffix))
    st.markdown("</div>",unsafe_allow_html=True)

    d_filt=df_exploded.copy()
    if sel_cat   !="All": d_filt=d_filt[d_filt["Category"]==sel_cat]
    if sel_vendor!="All": d_filt=d_filt[d_filt["Vendor"]==sel_vendor]
    avail_svcs=sorted([s for s in d_filt["Service"].unique()
                       if str(s).strip() not in ["","nan"]])
    if svc_filter:
        avail_svcs=[s for s in avail_svcs
                    if svc_filter.lower() in s.lower()]

    chat_key=chat_key_suffix
    left_col,right_col=st.columns([1,1],gap="large")

    # ── LEFT: CHATBOT ──
    with left_col:
        st.markdown(
            "<div class='chat-header'>"
            "<span style='font-weight:700;font-size:0.96em;"
            "font-family:Georgia,serif'>"
            "💬 Procurement Assistant</span>"
            "<span style='font-size:0.74em;opacity:0.6;"
            "margin-left:10px'>"
            "Full catalog knowledge · Upload a quote</span>"
            "</div>",unsafe_allow_html=True)

        chat_history_key="chat_history_{}".format(chat_key)
        if chat_history_key not in st.session_state:
            st.session_state[chat_history_key]=[]

        import re as _re
        chat_html="<div class='chat-outer'>"
        if not st.session_state[chat_history_key]:
            chat_html+=(
                "<div class='chat-wrap'><div class='msg-bot'>"
                "👋 Hello! I have full knowledge of this "
                "catalog.<br><br>"
                "<b>Try asking:</b><br>"
                "• <i>Who quoted Cisco Catalyst?</i><br>"
                "• <i>Compare Palo Alto prices</i><br>"
                "• <i>Cheapest Cybersecurity vendor?</i><br>"
                "• <i>What does TrendMicro offer?</i><br><br>"
                "Or upload a quote file below."
                "</div></div>")
        else:
            for turn in st.session_state[chat_history_key]:
                chat_html+=(
                    "<div class='chat-wrap'>"
                    "<div class='msg-user'>{}</div>"
                    "</div>".format(turn["user"]))
                bot=_re.sub(r'\*\*(.+?)\*\*',
                             r'<b>\1</b>',
                             turn["bot_text"].replace(
                                 "\n","<br>"))
                chat_html+=(
                    "<div class='chat-wrap'>"
                    "<div class='msg-bot'>{}</div>"
                    "</div>".format(bot))
        chat_html+="</div>"
        st.markdown(chat_html,unsafe_allow_html=True)

        # Show last chart
        if st.session_state[chat_history_key]:
            last_resp=st.session_state[
                chat_history_key][-1].get("bot_resp")
            if last_resp:
                render_chat_chart(last_resp)

        # File upload
        chat_file=st.file_uploader(
            "📎 Upload quote file for instant scoring",
            type=["pdf","xlsx","xls","docx"],
            key="chat_file_{}".format(chat_key))

        # Text input
        with st.form("chat_form_{}".format(chat_key),
                      clear_on_submit=True):
            ci1,ci2=st.columns([5,1])
            with ci1:
                user_input=st.text_input(
                    "msg",
                    placeholder="Ask a question…",
                    label_visibility="collapsed")
            with ci2:
                sent=st.form_submit_button(
                    "Send",type="primary",
                    use_container_width=True)

        # Handle send
        if sent and user_input.strip():
            resp=chatbot_response(
                user_input.strip(),df_master,df_exploded)
            st.session_state[chat_history_key].append({
                "user":user_input.strip(),
                "bot_text":resp["text"],
                "bot_resp":resp,
            })
            st.rerun()

        # Handle file upload
        last_file_key="last_chat_file_{}".format(chat_key)
        if (chat_file is not None
                and st.session_state.get(last_file_key)
                !=chat_file.name):
            st.session_state[last_file_key]=chat_file.name
            fb=chat_file.read()
            resp=chatbot_response(
                "uploaded file",df_master,df_exploded,
                uploaded_file_bytes=fb,
                uploaded_file_name=chat_file.name)
            st.session_state[chat_history_key].append({
                "user":"📎 {}".format(chat_file.name),
                "bot_text":resp["text"],
                "bot_resp":resp,
            })
            st.rerun()

        if st.button("🗑 Clear chat",
                      key="clr_{}".format(chat_key)):
            st.session_state[chat_history_key]=[]
            st.rerun()

    # ── RIGHT: MANUAL BROWSER ──
    with right_col:
        sec("SERVICE BROWSER")
        selected_svcs=st.multiselect(
            "Select services to analyse",
            options=avail_svcs,default=[],
            label_visibility="visible",
            key="sel_svcs_{}".format(chat_key))

        if not selected_svcs:
            sec("SERVICE COMPETITIVENESS MAP",
                "Orange = multiple vendors · "
                "Grey = single vendor only")
            svc_sum=(df_exploded.groupby("Service")[
                "Vendor"].nunique()
                      .reset_index()
                      .sort_values("Vendor",ascending=False))
            svc_sum.columns=["Service","Vendor Count"]
            top20=svc_sum.head(20)
            fig_sv=go.Figure(go.Bar(
                x=top20["Vendor Count"],
                y=top20["Service"].apply(lambda x:x[:48]),
                orientation="h",
                marker_color=[C_ORANGE if v>1 else C_GREY
                               for v in top20["Vendor Count"]],
                marker_line_width=0,
                text=top20["Vendor Count"],
                textposition="outside",
                textfont=dict(size=10)))
            fig_sv.update_layout(
                height=520,plot_bgcolor=CBG,paper_bgcolor=CBG,
                margin=dict(l=5,r=40,t=10,b=8),font=CFONT,
                xaxis=dict(title="Number of Vendors",
                           showgrid=True,gridcolor=C_GREY_LITE,
                           zeroline=False),
                yaxis=dict(autorange="reversed",
                           tickfont=dict(size=9.2)),
                bargap=0.28,showlegend=False)
            st.plotly_chart(fig_sv,use_container_width=True)
            n_multi=svc_sum[svc_sum["Vendor Count"]>1].shape[0]
            n_single=svc_sum[svc_sum["Vendor Count"]==1].shape[0]
            sm1,sm2,sm3=st.columns(3)
            kpi_box(sm1,len(svc_sum),"Total Services",C_DARK)
            kpi_box(sm2,n_multi,"Competitive (2+ vendors)",C_ORANGE)
            kpi_box(sm3,n_single,"Single Vendor Only",C_GREY_DARK)

        else:
            d_sel=d_filt[d_filt["Service"].isin(
                selected_svcs)].copy()
            if d_sel.empty:
                st.warning("No quotations found.")
            else:
                # Collect prices
                vpm={}
                for _,r in d_sel.drop_duplicates(
                        subset=["Vendor","File Name"]).iterrows():
                    v=r["Vendor"]
                    qp=_parse_num(str(r.get(
                        "Quoted Price","")).strip())
                    ck="px_{}_{}".format(
                        chat_key,
                        str(r.get("File Name","")).strip())
                    ca=st.session_state.get(ck)
                    ep=ca["price_num"] if ca else 0.0
                    ref=ep if ep>0 else qp
                    if ref>0:
                        vpm[v]=min(vpm.get(v,ref),ref)
                if not vpm:
                    for _,r in d_sel.drop_duplicates(
                            subset=["Vendor"]).iterrows():
                        v=r["Vendor"]
                        qp=_parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp>0 and v not in vpm:
                            vpm[v]=qp

                # Coverage verdict
                vsmap=defaultdict(set)
                for _,r in d_sel.iterrows():
                    vsmap[r["Vendor"]].add(r["Service"])
                full_cov=[v for v,s in vsmap.items()
                           if set(selected_svcs).issubset(s)]
                partial=[v for v,s in vsmap.items()
                          if set(selected_svcs)&s
                          and v not in full_cov]

                if len(selected_svcs)==1:
                    nv=d_sel["Vendor"].nunique()
                    if nv>1:
                        st.markdown(
                            "<div class='verdict-good'>"
                            "✅ <b>{}</b> — quoted by "
                            "<b>{} vendors</b>.</div>".format(
                                selected_svcs[0],nv),
                            unsafe_allow_html=True)
                    else:
                        st.markdown(
                            "<div class='verdict-mid'>"
                            "⚠️ <b>{}</b> — only "
                            "<b>1 vendor</b>.</div>".format(
                                selected_svcs[0]),
                            unsafe_allow_html=True)
                else:
                    if full_cov:
                        st.markdown(
                            "<div class='verdict-good'>"
                            "✅ <b>{}</b> vendor(s) cover "
                            "ALL {}: <b>{}</b></div>".format(
                                len(full_cov),
                                len(selected_svcs),
                                ", ".join(full_cov)),
                            unsafe_allow_html=True)
                    else:
                        st.markdown(
                            "<div class='verdict-mid'>"
                            "⚠️ No single vendor covers all "
                            "{}. Partial: <b>{}</b></div>".format(
                                len(selected_svcs),
                                ", ".join(partial)
                                if partial else "None"),
                            unsafe_allow_html=True)
                st.markdown("<br>",unsafe_allow_html=True)

                # ── CHART 1: Vendor × Service Heatmap ──
                sec("VENDOR × SERVICE COVERAGE MAP",
                    "✅ = vendor has quoted this service")
                heat=[]
                for svc in selected_svcs:
                    for v in sorted(d_sel["Vendor"].unique()):
                        has=int(len(d_sel[
                            (d_sel["Service"]==svc)
                            &(d_sel["Vendor"]==v)])>0)
                        heat.append({
                            "Service":svc[:40],
                            "Vendor":v,"Covered":has})
                heat_df=pd.DataFrame(heat)
                if not heat_df.empty:
                    pivot=heat_df.pivot_table(
                        index="Service",columns="Vendor",
                        values="Covered",fill_value=0)
                    fig_h=go.Figure(go.Heatmap(
                        z=pivot.values,
                        x=pivot.columns.tolist(),
                        y=pivot.index.tolist(),
                        colorscale=[[0,C_GREY_LITE],
                                     [1,C_ORANGE]],
                        showscale=False,
                        text=[["✅" if v==1 else "—"
                               for v in row]
                              for row in pivot.values],
                        texttemplate="%{text}",
                        textfont=dict(size=16)))
                    fig_h.update_layout(
                        height=max(180,
                                   len(selected_svcs)*65),
                        plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=10,b=10),
                        font=CFONT,
                        xaxis=dict(tickangle=-20,
                                   tickfont=dict(size=10)),
                        yaxis=dict(tickfont=dict(size=9.5),
                                   autorange="reversed"))
                    st.plotly_chart(fig_h,
                                     use_container_width=True)

                # ── CHART 2: Quotes per vendor ──
                sec("QUOTES PER VENDOR",
                    "Number of quote files per vendor")
                vpq=(d_sel.drop_duplicates(
                    subset=["Vendor","File Name"])
                      .groupby("Vendor").size()
                      .reset_index())
                vpq.columns=["Vendor","Quotes"]
                vpq=vpq.sort_values("Quotes",ascending=False)
                fig_vpq=go.Figure(go.Bar(
                    x=vpq["Vendor"],y=vpq["Quotes"],
                    marker_color=[vcmap.get(v,C_DARK)
                                   for v in vpq["Vendor"]],
                    marker_line_width=0,
                    text=vpq["Quotes"],
                    textposition="outside"))
                fig_vpq.update_layout(
                    height=260,plot_bgcolor=CBG,
                    paper_bgcolor=CBG,
                    margin=dict(l=5,r=10,t=10,b=8),
                    font=CFONT,
                    yaxis=dict(title="Quote Files",
                               showgrid=True,
                               gridcolor=C_GREY_LITE,
                               zeroline=False),
                    xaxis=dict(tickangle=-15,
                               tickfont=dict(size=10.5)),
                    bargap=0.4,showlegend=False)
                st.plotly_chart(fig_vpq,
                                 use_container_width=True)

                # ── CHART 3: Vendors per service ──
                if len(selected_svcs)>1:
                    sec("VENDORS PER SERVICE")
                    svc_vc=[]
                    for svc in selected_svcs:
                        nv2=d_sel[d_sel["Service"]==svc
                                  ]["Vendor"].nunique()
                        svc_vc.append({
                            "Service":svc[:45],
                            "Vendors":nv2})
                    svc_vc_df=pd.DataFrame(svc_vc)
                    fig_svc=go.Figure(go.Bar(
                        x=svc_vc_df["Service"],
                        y=svc_vc_df["Vendors"],
                        marker_color=[
                            C_ORANGE if v>1
                            else C_GREY_DARK
                            for v in svc_vc_df["Vendors"]],
                        marker_line_width=0,
                        text=svc_vc_df["Vendors"],
                        textposition="outside"))
                    fig_svc.update_layout(
                        height=260,plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=10,b=8),
                        font=CFONT,
                        yaxis=dict(title="Vendors",
                                   showgrid=True,
                                   gridcolor=C_GREY_LITE,
                                   zeroline=False),
                        xaxis=dict(tickangle=-15,
                                   tickfont=dict(size=10)),
                        bargap=0.4,showlegend=False)
                    st.plotly_chart(fig_svc,
                                     use_container_width=True)

                # ── Price section (only if has_prices) ──
                if vpm:
                    avg_p=sum(vpm.values())/len(vpm)
                    best_v=min(vpm,key=vpm.get)
                    worst_v=max(vpm,key=vpm.get)
                    spread=round(
                        (max(vpm.values())
                         -min(vpm.values()))
                        /min(vpm.values())*100,1
                    ) if min(vpm.values())>0 else 0

                    sec("PRICE VERDICT")
                    pv1,pv2,pv3=st.columns(3)
                    pv1.markdown(
                        "<div class='scard scard-orange'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;text-transform:"
                        "uppercase;color:#D04A02'>"
                        "Best Price</div>"
                        "<div style='font-size:1.5em;"
                        "font-weight:800;color:#D04A02;"
                        "font-family:Georgia,serif'>"
                        "{}</div>"
                        "<div style='font-size:0.73em;"
                        "color:#7D7D7D;margin-top:2px'>"
                        "{}</div></div>".format(
                            _fmt(min(vpm.values())),best_v),
                        unsafe_allow_html=True)
                    pv2.markdown(
                        "<div class='scard scard-dark'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;text-transform:"
                        "uppercase;color:#2D2D2D'>"
                        "Market Average</div>"
                        "<div style='font-size:1.5em;"
                        "font-weight:800;color:#2D2D2D;"
                        "font-family:Georgia,serif'>"
                        "{}</div>"
                        "<div style='font-size:0.73em;"
                        "color:#7D7D7D;margin-top:2px'>"
                        "{} vendors</div></div>".format(
                            _fmt(avg_p),len(vpm)),
                        unsafe_allow_html=True)
                    pv3.markdown(
                        "<div class='scard scard-grey'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;text-transform:"
                        "uppercase;color:#7D7D7D'>"
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
                    sec("PRICE COMPARISON CHART",
                        "Orange = best · Dark = highest "
                        "· Dashed = market average")
                    sv=sorted(vpm.items(),key=lambda x:x[1])
                    bc=[C_ORANGE if v==best_v
                         else C_DARK if v==worst_v
                         else C_GREY_DARK for v,_ in sv]
                    fig_c=go.Figure(go.Bar(
                        x=[v for v,_ in sv],
                        y=[p for _,p in sv],
                        marker_color=bc,
                        marker_line_width=0,
                        text=[_fmt(p) for _,p in sv],
                        textposition="outside",
                        textfont=dict(size=11,color=C_DARK)))
                    fig_c.add_hline(
                        y=avg_p,line_dash="dash",
                        line_color=C_MID,line_width=2,
                        annotation_text="Avg: {}".format(
                            _fmt(avg_p)),
                        annotation_position="top right")
                    fig_c.update_layout(
                        height=290,plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=24,b=8),
                        font=CFONT,
                        yaxis=dict(title="Price (USD)",
                                   showgrid=True,
                                   gridcolor=C_GREY_LITE,
                                   zeroline=False),
                        xaxis=dict(tickangle=-10,
                                   tickfont=dict(size=10.5)),
                        bargap=0.4,showlegend=False)
                    st.plotly_chart(fig_c,
                                     use_container_width=True)

                    # Score table
                    sec("VENDOR SCORE CARD")
                    all_pv=list(vpm.values())
                    tbl=["<table class='comp-table'>"
                         "<thead><tr><th>Rank</th>"
                         "<th>Vendor</th><th>Price</th>"
                         "<th>vs Average</th>"
                         "<th>Score</th><th>Verdict</th>"
                         "</tr></thead><tbody>"]
                    for rank,(v,p) in enumerate(
                            sorted(vpm.items(),
                                   key=lambda x:x[1]),1):
                        bg="white" if rank%2==0 else "#F8F8F8"
                        vc=vcmap.get(v,C_DARK)
                        oth=[x for x in all_pv if x!=p]
                        ps=None
                        if p>0 and oth:
                            ps,_,_,_,_=price_score(p,oth)
                        sc=score_color(ps)
                        pct=round((p-avg_p)/avg_p*100,1) \
                            if avg_p>0 else 0
                        vs=("{}% below ✅".format(abs(pct))
                            if pct<0
                             else "{}% above ⚠️".format(abs(pct))
                             if pct>0 else "At average")
                        vc2=(C_ORANGE if pct<0
                              else C_DARK if pct>10
                              else C_GREY_DARK)
                        vt,_,_=get_verdict(ps)
                        medal=("🥇" if rank==1
                                else "🥈" if rank==2
                                else "🥉" if rank==3
                                else str(rank))
                        tbl.append(
                            "<tr style='background:{}'>"
                            "<td style='text-align:center;"
                            "font-size:1.05em'>{}</td>"
                            "<td>{}</td>"
                            "<td style='font-family:monospace;"
                            "font-weight:700'>{}</td>"
                            "<td style='color:{}'>{}</td>"
                            "<td style='text-align:center'>"
                            "<span style='font-weight:800;"
                            "font-size:1.1em;color:{}'>"
                            "{}</span></td>"
                            "<td style='font-weight:700;"
                            "color:{}'>{}</td>"
                            "</tr>".format(
                                bg,medal,vpill(v,vc),
                                _fmt(p),vc2,vs,sc,
                                ps if ps is not None else "—",
                                sc,vt))
                    tbl.append("</tbody></table>")
                    st.markdown("".join(tbl),
                                 unsafe_allow_html=True)

                    pct_b=round(
                        (avg_p-min(vpm.values()))
                        /avg_p*100,1) if avg_p>0 else 0
                    insight(
                        "<b>{}</b> is most competitive "
                        "at <b>{}</b> — "
                        "<b>{}% below</b> market avg. "
                        "Spread <b>{}%</b> → <b>{}</b>.".format(
                            best_v,
                            _fmt(min(vpm.values())),
                            pct_b,spread,
                            "strong negotiation potential"
                            if spread>20
                            else "moderate room"
                            if spread>10
                            else "competitive market"))

                elif not has_prices:
                    # Master catalog — no prices
                    st.markdown(
                        "<div class='insight-box'>"
                        "ℹ️ Price data is not available "
                        "for the Master Catalog — "
                        "quotation files are stored on "
                        "SharePoint. Switch to the "
                        "<b>Dummy Data</b> tabs to see "
                        "full price analysis."
                        "</div>",
                        unsafe_allow_html=True)

                # File detail table
                st.markdown("<br>",unsafe_allow_html=True)
                sec("QUOTATION FILE DETAILS")
                has_price_col=(
                    "Quoted Price" in d_sel.columns)
                for svc in selected_svcs:
                    d_svc=(d_sel[d_sel["Service"]==svc]
                           .drop_duplicates(
                               subset=["Vendor","File Name"])
                           .sort_values("Vendor"))
                    nv2=d_svc["Vendor"].nunique()
                    st.markdown(
                        "<div style='background:white;"
                        "border-left:4px solid {};"
                        "padding:10px 14px;"
                        "border-radius:2px;margin:8px 0;"
                        "font-weight:700;"
                        "font-size:0.88em'>"
                        "{}  ·  {} vendor(s)  ·  {}"
                        "</div>".format(
                            C_ORANGE if nv2>1
                            else C_GREY_DARK,
                            svc,nv2,
                            "✅ COMPETITIVE"
                            if nv2>1
                            else "⚠️ SINGLE VENDOR"),
                        unsafe_allow_html=True)

                    all_p2=[]
                    for _,r in d_svc.iterrows():
                        qp=_parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp>0: all_p2.append(qp)

                    rt=["<table class='comp-table'>"
                        "<thead><tr>"
                        "<th>Vendor</th><th>File</th>"]
                    if has_price_col:
                        rt.append("<th>Quoted Price</th>")
                    rt.append(
                        "<th>Score</th><th>Verdict</th>"
                        "<th>Open</th>"
                        "</tr></thead><tbody>")

                    for i,(_,row) in enumerate(
                            d_svc.iterrows()):
                        bg2=("white" if i%2==0
                              else "#F8F8F8")
                        vc3=vcmap.get(row["Vendor"],C_DARK)
                        fname=str(row.get(
                            "File Name","")).strip()
                        url=resolve_url(row)
                        qpn=_parse_num(str(row.get(
                            "Quoted Price","")).strip())
                        ck2="px_{}_{}".format(
                            chat_key,fname)
                        ca=st.session_state.get(ck2)
                        ref=(ca["price_num"]
                              if ca and ca.get(
                                  "price_num",0)>0
                              else qpn
                              if qpn>0 else 0)
                        oth=[p for p in all_p2
                             if p!=ref]
                        ps2=None; vt2="—"
                        vc4=C_GREY_DARK
                        if ref>0 and oth:
                            ps2,_,_,_,_=price_score(
                                ref,oth)
                            vt2,_,vc4=get_verdict(ps2)
                        sc2=score_color(ps2)
                        lnk=("<a href='{}' "
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
                        if has_price_col:
                            rt.append(
                                "<td style='font-family:"
                                "monospace;font-weight:700;"
                                "color:#D04A02'>{}</td>"
                                .format(_fmt(qpn)
                                        if qpn>0
                                        else "—"))
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

                    # Extract prices button
                    # (only shown for dummy data or
                    #  if local files exist)
                    if has_prices or os.path.exists(
                            DEMO_DIR):
                        st.markdown("<br>",
                                     unsafe_allow_html=True)
                        if st.button(
                                "🔍 Extract Prices — "
                                "{}".format(svc[:38]),
                                key="ep_{}_{}".format(
                                    chat_key,svc[:28]),
                                type="primary"):
                            prog=st.progress(0)
                            nn=len(d_svc)
                            for ki,(_,r2) in enumerate(
                                    d_svc.iterrows()):
                                f2=str(r2.get(
                                    "File Name","")).strip()
                                ck3="px_{}_{}".format(
                                    chat_key,f2)
                                if not st.session_state.get(ck3):
                                    loc=os.path.join(DEMO_DIR,f2)
                                    if os.path.exists(loc):
                                        st.session_state[ck3]=(
                                            extract_price_from_file(loc))
                                    else:
                                        u2=resolve_url(r2)
                                        if (u2
                                                and u2.startswith("http")
                                                and REQUESTS_OK):
                                            try:
                                                rr=requests.get(
                                                    u2,timeout=20)
                                                ee=u2.split("?"
                                                            )[0].rsplit(
                                                    ".",1)[-1].lower()
                                                st.session_state[ck3]=(
                                                    extract_price_from_bytes(
                                                        rr.content,ee))
                                            except: pass
                                prog.progress((ki+1)/nn)
                            prog.empty()
                            st.rerun()


# ════════════════════════════════════════════════════════════
# SESSION STATE
# ════════════════════════════════════════════════════════════
for k,v in [
    ("tab2_upload_price",0.0),
    ("tab2_upload_fname",""),
    ("tab2_file_bytes",None),
    ("tab2_file_ext",""),
    ("chat_redirect_upload",False),
    ("gh_prices_loaded",False),
    ("real_analysis_done",False),
    ("real_analysis_df",None),
]:
    if k not in st.session_state:
        st.session_state[k]=v

# ════════════════════════════════════════════════════════════
# LOAD BOTH DATA SOURCES
# ════════════════════════════════════════════════════════════
df_master, df_exp_master = load_master_catalog()
df_dummy,  df_exp_dummy  = load_dummy_data()

NO_MASTER = (df_master is None or df_master.empty)
NO_DUMMY  = (df_dummy  is None or df_dummy.empty)

# Build vendor colour maps
vcmap_master={}
vcmap_dummy={}
if not NO_MASTER:
    for i,v in enumerate(
            sorted(df_master["Vendor"].unique())):
        vcmap_master[v]=CHART_SEQ[i%len(CHART_SEQ)]
if not NO_DUMMY:
    for i,v in enumerate(
            sorted(df_dummy["Vendor"].unique())):
        vcmap_dummy[v]=CHART_SEQ[i%len(CHART_SEQ)]

# For backward compat
vendor_color_map=vcmap_master if not NO_MASTER else vcmap_dummy

# ════════════════════════════════════════════════════════════
# MAIN HEADER
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
    "Master Catalog (Excel) · Dummy Data (CSV) · "
    "Browse &amp; Verdict · Upload &amp; Score · "
    "Vendor Analysis</p>"
    "</div>",
    unsafe_allow_html=True)

# Global KPIs — show both sources
col_m, col_d = st.columns(2, gap="large")
with col_m:
    st.markdown(
        "<div class='bucket-header'>"
        "📁 MASTER CATALOG (Excel)"
        "</div>",
        unsafe_allow_html=True)
    if not NO_MASTER:
        km1,km2,km3,km4=st.columns(4)
        kpi_box(km1,df_master["File Name"].nunique(),
                "Quotes",C_ORANGE)
        kpi_box(km2,df_exp_master["Service"].nunique(),
                "Services",C_DARK)
        kpi_box(km3,df_master["Vendor"].nunique(),
                "Vendors",C_MID)
        kpi_box(km4,df_master["Category"].nunique(),
                "Categories",C_GREY_DARK)
    else:
        st.warning("Master Catalog.xlsx not found.")

with col_d:
    st.markdown(
        "<div class='bucket-header'>"
        "📊 DUMMY DATA (CSV — with prices)"
        "</div>",
        unsafe_allow_html=True)
    if not NO_DUMMY:
        kd1,kd2,kd3,kd4=st.columns(4)
        kpi_box(kd1,df_dummy["File Name"].nunique(),
                "Quotes",C_ORANGE)
        kpi_box(kd2,df_exp_dummy["Service"].nunique(),
                "Services",C_DARK)
        kpi_box(kd3,df_dummy["Vendor"].nunique(),
                "Vendors",C_MID)
        kpi_box(kd4,df_dummy["Category"].nunique(),
                "Categories",C_GREY_DARK)
    else:
        st.warning("dummy_catalog.csv not found.")

st.markdown("<br>",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# GLOBAL NAV STRIP (category cards — master catalog)
# ════════════════════════════════════════════════════════════
if not NO_MASTER:
    cats_list=sorted([
        c for c in df_master["Category"].unique()
        if str(c).strip() not in ["","nan"]])
    nav_cols=st.columns(len(cats_list)+1)
    nav_cols[0].markdown(
        "<div style='font-size:0.72em;font-weight:700;"
        "color:#D04A02;letter-spacing:1px;"
        "text-transform:uppercase;padding-top:4px'>"
        "CATEGORIES</div>",
        unsafe_allow_html=True)
    for i,cat in enumerate(cats_list):
        cnt=len(df_master[df_master["Category"]==cat])
        nav_cols[i+1].markdown(
            "<div style='background:white;"
            "border:1px solid #E0E0E0;"
            "border-top:3px solid #D04A02;"
            "border-radius:4px;"
            "padding:8px 10px;text-align:center'>"
            "<div style='font-size:0.72em;font-weight:700;"
            "color:#2D2D2D;white-space:nowrap;"
            "overflow:hidden;text-overflow:ellipsis'>"
            "{}</div>"
            "<div style='font-size:1.1em;font-weight:800;"
            "color:#D04A02;font-family:Georgia,serif'>"
            "{}</div>"
            "<div style='font-size:0.63em;color:#7D7D7D;"
            "text-transform:uppercase;letter-spacing:0.5px'>"
            "quotes</div></div>".format(cat,cnt),
            unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TABS — Two buckets clearly labelled
# ════════════════════════════════════════════════════════════
(tab_mc_ov,      # Master Catalog Overview
 tab_mc_bv,      # Master Catalog Browse & Verdict
 tab_dd_ov,      # Dummy Data Overview
 tab_dd_bv,      # Dummy Data Browse & Verdict
 tab_upload,     # Upload & Score
 tab_data,       # Data Table
 tab_upload_cat, # Upload Catalog
 tab_vendor,     # Vendor Analysis
 ) = st.tabs([
    "📁 MC — Catalog Overview",
    "📁 MC — Browse & Verdict",
    "📊 DD — Catalog Overview",
    "📊 DD — Browse & Verdict",
    "📤 Upload & Score",
    "📄 Data Table",
    "🗂 Upload Catalog",
    "🔍 Vendor Analysis",
])

# ════════════════════════════════════════════════════════════
# TAB: MC — CATALOG OVERVIEW
# ════════════════════════════════════════════════════════════
with tab_mc_ov:
    st.markdown(
        "<div class='bucket-header'>"
        "📁 MASTER CATALOG — CATALOG OVERVIEW"
        "</div>",unsafe_allow_html=True)
    if NO_MASTER:
        st.info("Master Catalog.xlsx not found.")
    else:
        render_catalog_overview(
            df_master,df_exp_master,
            label="Master Catalog")

# ════════════════════════════════════════════════════════════
# TAB: MC — BROWSE & VERDICT
# ════════════════════════════════════════════════════════════
with tab_mc_bv:
    st.markdown(
        "<div class='bucket-header'>"
        "📁 MASTER CATALOG — BROWSE &amp; VERDICT"
        "<span style='font-size:0.8em;opacity:0.7;"
        "margin-left:12px'>"
        "No prices available (SharePoint files)</span>"
        "</div>",unsafe_allow_html=True)
    if NO_MASTER:
        st.info("Master Catalog.xlsx not found.")
    else:
        render_browse_verdict(
            df_master,df_exp_master,
            vcmap=vcmap_master,
            label="Master Catalog",
            has_prices=False,
            chat_key_suffix="mc")

# ════════════════════════════════════════════════════════════
# TAB: DD — CATALOG OVERVIEW
# ════════════════════════════════════════════════════════════
with tab_dd_ov:
    st.markdown(
        "<div class='bucket-header'>"
        "📊 DUMMY DATA — CATALOG OVERVIEW"
        "<span style='font-size:0.8em;opacity:0.7;"
        "margin-left:12px'>"
        "Full price analysis available</span>"
        "</div>",unsafe_allow_html=True)
    if NO_DUMMY:
        st.info("dummy_catalog.csv not found.")
    else:
        render_catalog_overview(
            df_dummy,df_exp_dummy,
            label="Dummy Data")

# ════════════════════════════════════════════════════════════
# TAB: DD — BROWSE & VERDICT
# ════════════════════════════════════════════════════════════
with tab_dd_bv:
    st.markdown(
        "<div class='bucket-header'>"
        "📊 DUMMY DATA — BROWSE &amp; VERDICT"
        "<span style='font-size:0.8em;opacity:0.7;"
        "margin-left:12px'>"
        "✅ Full price analysis · Scores · Verdicts"
        "</span></div>",unsafe_allow_html=True)
    if NO_DUMMY:
        st.info("dummy_catalog.csv not found.")
    else:
        render_browse_verdict(
            df_dummy,df_exp_dummy,
            vcmap=vcmap_dummy,
            label="Dummy Data",
            has_prices=True,
            chat_key_suffix="dd")

# ════════════════════════════════════════════════════════════
# TAB: UPLOAD & SCORE
# ════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown(
        "<div class='bucket-header'>"
        "📤 UPLOAD &amp; SCORE</div>",
        unsafe_allow_html=True)

    # Use dummy data as benchmark (has prices)
    df_bench   = df_dummy   if not NO_DUMMY else df_master
    df_exp_bench = df_exp_dummy if not NO_DUMMY else df_exp_master
    bench_label= "Dummy Data" if not NO_DUMMY else "Master Catalog"

    if df_bench is None:
        st.info("No catalog loaded for benchmarking.")
    else:
        insight(
            "Benchmarking against: <b>{}</b>".format(
                bench_label))
        prefill_price=st.session_state.get(
            "tab2_upload_price",0.0)
        prefill_fname=st.session_state.get(
            "tab2_upload_fname","")
        if prefill_fname:
            st.markdown(
                "<div class='insight-box'>"
                "📎 From chat: <b>{}</b> — "
                "price: <b>{}</b></div>".format(
                    prefill_fname,
                    _fmt(prefill_price)
                    if prefill_price>0
                    else "not found"),
                unsafe_allow_html=True)

        sec("STEP 1 — UPLOAD QUOTE FILE")
        uploaded=st.file_uploader(
            "Upload",
            type=["pdf","xlsx","xls","docx"],
            label_visibility="collapsed",
            key="up_score_file")
        new_price=0.0; fname_up=""
        if uploaded is not None:
            content=uploaded.read()
            ext_up=uploaded.name.rsplit(".",1)[-1]
            fname_up=uploaded.name
            st.success("Uploaded: **{}** ({} KB)".format(
                fname_up,round(len(content)/1024,1)))
            sec("STEP 2 — EXTRACTED PRICE")
            with st.spinner("Extracting…"):
                res=extract_price_from_bytes(
                    content,ext_up)
                new_price=res["price_num"]
            if new_price>0:
                st.markdown(
                    "<div class='scard scard-orange'>"
                    "<div style='font-size:0.70em;"
                    "font-weight:700;text-transform:"
                    "uppercase;color:#D04A02'>"
                    "Extracted Price</div>"
                    "<div style='font-size:2.1em;"
                    "font-weight:800;color:#D04A02;"
                    "font-family:Georgia,serif'>"
                    "{}</div></div>".format(_fmt(new_price)),
                    unsafe_allow_html=True)
            else:
                st.warning("Price not found automatically.")
                manual=st.number_input(
                    "Enter price manually (USD)",
                    min_value=0.0,step=100.0,value=0.0,
                    key="manual_price_up")
                if manual>0: new_price=manual
        elif prefill_price>0:
            new_price=prefill_price
            fname_up=prefill_fname

        sec("STEP 3 — SELECT SERVICES & FILTERS")
        up1,up2,up3=st.columns(3)
        with up1:
            cat_up=st.selectbox(
                "📂 Filter by Category",
                ["All"]+sorted([
                    c for c in df_bench[
                        "Category"].unique()
                    if str(c).strip()
                    not in ["","nan"]]),
                key="cat_up_score")
        with up2:
            svcs_up=sorted([
                s for s in df_exp_bench[
                    "Service"].unique()
                if str(s).strip()
                not in ["","nan"]])
            svc_srch=st.text_input(
                "🔍 Filter services",
                placeholder="Search…",
                key="svc_srch_score")
            if svc_srch:
                svcs_up=[s for s in svcs_up
                          if svc_srch.lower() in s.lower()]
        with up3:
            new_svcs=st.multiselect(
                "🛠 Select Services",
                options=svcs_up,
                key="new_svcs_score")

        sec("STEP 4 — COMPARISON & VERDICT")
        if new_price<=0 and not new_svcs:
            st.info(
                "Upload a file and select services "
                "to compare.")
        else:
            cands=(
                df_exp_bench[
                    df_exp_bench["Service"].isin(
                        new_svcs)].copy()
                if new_svcs
                else df_exp_bench.copy())
            if cat_up!="All":
                cands=cands[cands["Category"]==cat_up]
            cf=(cands.drop_duplicates(
                subset=["File Name","Vendor"])
                [["File Name","Vendor","Category",
                  "Quoted Price"]].copy()
                if "Quoted Price" in cands.columns
                else cands.drop_duplicates(
                    subset=["File Name","Vendor"])
                [["File Name","Vendor","Category"]].copy())

            if cf.empty:
                st.warning("No historical quotes found.")
            else:
                hist=[]
                if "Quoted Price" in cf.columns:
                    for _,r in cf.iterrows():
                        qp=_parse_num(str(r.get(
                            "Quoted Price","")).strip())
                        if qp>0: hist.append(qp)

                if new_price>0 and hist:
                    ps,lbl,avh,mnh,mxh=price_score(
                        new_price,hist)
                    vt,vd,vc5=get_verdict(ps)
                    css=("orange" if (ps or 0)>=70
                          else "mid" if (ps or 0)>=40
                          else "dark")
                    bgs={"orange":"#FFF5F0",
                          "mid":"#F5F5F5",
                          "dark":"#F0F0F0"}
                    bds={"orange":C_ORANGE,
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

                    sv1,sv2,sv3,sv4=st.columns(4)
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
                        "color:#2D2D2D'>Your Price</div>"
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
                            _fmt(avh),_fmt(mnh),_fmt(mxh)),
                        unsafe_allow_html=True)
                    sv4.markdown(
                        "<div class='scard scard-grey'>"
                        "<div style='font-size:0.67em;"
                        "font-weight:700;"
                        "text-transform:uppercase;"
                        "color:#7D7D7D'>vs Average</div>"
                        "<div style='font-size:0.95em;"
                        "font-weight:800;color:#4A4A4A;"
                        "margin-top:8px'>"
                        "{}</div></div>".format(lbl),
                        unsafe_allow_html=True)

                    st.markdown("<br>",
                                 unsafe_allow_html=True)
                    sec("PRICE POSITIONING CHART")
                    cd2=[]
                    if "Quoted Price" in cf.columns:
                        for _,r in cf.iterrows():
                            pv=_parse_num(str(r.get(
                                "Quoted Price","")).strip())
                            if pv>0:
                                cd2.append({
                                    "Label":"{}/{}".format(
                                        r["Vendor"],
                                        str(r["File Name"])[:10]),
                                    "Price":pv,
                                    "Type":"Historical"})
                    cd2.append({
                        "Label":"★ YOUR QUOTE",
                        "Price":new_price,
                        "Type":"New"})
                    cdf2=pd.DataFrame(cd2
                        ).sort_values("Price")
                    bc2=[C_ORANGE if t=="New"
                          else C_DARK
                          for t in cdf2["Type"]]
                    fig_up=go.Figure(go.Bar(
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
                        annotation_text="Avg: {}".format(
                            _fmt(avh)),
                        annotation_position="top right")
                    fig_up.update_layout(
                        height=380,plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=20,b=10),
                        font=CFONT,
                        yaxis=dict(
                            title="Price (USD)",
                            showgrid=True,
                            gridcolor=C_GREY_LITE,
                            zeroline=False),
                        xaxis=dict(tickangle=-25),
                        bargap=0.3,showlegend=False)
                    st.plotly_chart(fig_up,
                                     use_container_width=True)
                    pct_vs=round(
                        (new_price-avh)/avh*100,1
                    ) if avh>0 else 0
                    insight(
                        "Your quote <b>{}</b> is "
                        "<b>{}% {}</b> market avg "
                        "<b>{}</b>. Range: "
                        "<b>{}</b>–<b>{}</b>.".format(
                            _fmt(new_price),abs(pct_vs),
                            "below" if pct_vs<0
                            else "above",
                            _fmt(avh),_fmt(mnh),_fmt(mxh)))
                else:
                    st.info(
                        "No historical price data. "
                        "Select more services or use "
                        "Dummy Data tab.")

# ════════════════════════════════════════════════════════════
# TAB: DATA TABLE
# ════════════════════════════════════════════════════════════
with tab_data:
    st.markdown(
        "<div class='bucket-header'>"
        "📄 DATA TABLE</div>",
        unsafe_allow_html=True)
    src_choice=st.radio(
        "Data source",
        ["Master Catalog","Dummy Data"],
        horizontal=True,key="dt_src")
    df_dt=(df_master if src_choice=="Master Catalog"
            else df_dummy)
    df_dt_use=(df_dt if df_dt is not None
                else pd.DataFrame())

    tf1,tf2,tf3=st.columns(3)
    with tf1:
        dt_cat=st.selectbox(
            "📂 Category",
            ["All"]+sorted([
                c for c in df_dt_use.get(
                    "Category",pd.Series()).unique()
                if str(c).strip()
                not in ["","nan"]])
            if not df_dt_use.empty else ["All"],
            key="dt_cat")
    with tf2:
        vp_dt=(df_dt_use if dt_cat=="All"
                else df_dt_use[
                    df_dt_use["Category"]==dt_cat])
        dt_ven=st.selectbox(
            "🏢 Vendor",
            ["All"]+sorted([
                v for v in vp_dt.get(
                    "Vendor",pd.Series()).unique()
                if str(v).strip()
                not in ["","nan"]])
            if not vp_dt.empty else ["All"],
            key="dt_ven")
    with tf3:
        dt_srch=st.text_input(
            "🔍 Search",
            placeholder="File name or comments…",
            key="dt_srch")

    dm=df_dt_use.copy()
    if not dm.empty:
        if dt_cat!="All":
            dm=dm[dm["Category"]==dt_cat]
        if dt_ven!="All":
            dm=dm[dm["Vendor"]==dt_ven]
        if dt_srch:
            mask=(dm["File Name"].str.contains(
                dt_srch,case=False,na=False)
                  |dm["Comments"].str.contains(
                dt_srch,case=False,na=False))
            dm=dm[mask]
        st.markdown(
            "<div style='font-size:0.82em;"
            "color:#7D7D7D;margin:8px 0'>"
            "Showing <b>{}</b> of <b>{}</b> "
            "records</div>".format(
                len(dm),len(df_dt_use)),
            unsafe_allow_html=True)
        st.dataframe(
            dm.drop(
                columns=["Services List","Hyperlink"],
                errors="ignore"),
            use_container_width=True,height=520)
    else:
        st.info("No data available.")

# ════════════════════════════════════════════════════════════
# TAB: UPLOAD CATALOG
# ════════════════════════════════════════════════════════════
with tab_upload_cat:
    st.markdown(
        "<div class='bucket-header'>"
        "🗂 UPLOAD CATALOG</div>",
        unsafe_allow_html=True)
    insight(
        "Upload a new catalog to replace the "
        "current one. Supports Excel (.xlsx) "
        "and CSV formats.")
    cat_file=st.file_uploader(
        "Upload",type=["xlsx","xls","csv"],
        label_visibility="collapsed",
        key="catalog_upload")
    if cat_file is not None:
        fb=cat_file.read(); fn=cat_file.name
        with st.spinner("Analysing…"):
            from io import BytesIO as _BIO
            ext_c=fn.rsplit(".",1)[-1].lower()
            try:
                if ext_c in ("xlsx","xls"):
                    df_n=pd.read_excel(
                        _BIO(fb),engine="openpyxl")
                else:
                    df_n=pd.read_csv(_BIO(fb))
                df_n.columns=[str(c).strip()
                               for c in df_n.columns]
                df_n=_norm_cols(df_n)
                df_n=_clean_df(df_n)
                df_n["Hyperlink"]=""
                df_n,dfe_n=_explode(df_n)
                st.success(
                    "✅ **{}** rows · **{}** vendors · "
                    "**{}** categories".format(
                        len(df_n),
                        df_n["Vendor"].nunique(),
                        df_n["Category"].nunique()))
                st.dataframe(
                    df_n.drop(
                        columns=["Services List",
                                  "Hyperlink"],
                        errors="ignore").head(20),
                    use_container_width=True,
                    height=280)
                if st.button(
                        "✅ Apply as Master Catalog",
                        type="primary"):
                    st.session_state[
                        "uploaded_catalog_df"]=df_n
                    st.session_state[
                        "uploaded_catalog_exp"]=dfe_n
                    st.success("✅ Applied!")
                    st.rerun()
            except Exception as e:
                st.error("❌ {}".format(e))

# ════════════════════════════════════════════════════════════
# TAB: VENDOR ANALYSIS
# ════════════════════════════════════════════════════════════
with tab_vendor:
    st.markdown(
        "<div class='bucket-header'>"
        "🔍 VENDOR ANALYSIS — DUMMY DATA "
        "(Full Price Analysis)</div>",
        unsafe_allow_html=True)
    if NO_DUMMY:
        st.info("Dummy data not available.")
    else:
        sec("VENDOR PRICE SUMMARY",
            "Based on dummy data with actual prices")
        vt_dd=(df_dummy.groupby("Vendor")[
            "Quoted Price"]
                .agg(["mean","min","max","count"])
                .reset_index()
                if "Quoted Price" in df_dummy.columns
                else pd.DataFrame())

        if not vt_dd.empty:
            vt_dd.columns=[
                "Vendor","Average","Min","Max","Quotes"]
            vt_dd=vt_dd.sort_values("Average")
            oa=vt_dd["Average"].mean()

            k1,k2,k3,k4=st.columns(4)
            kpi_box(k1,len(df_dummy),
                     "Total Quotes",C_ORANGE)
            kpi_box(k2,df_dummy["Vendor"].nunique(),
                     "Vendors",C_DARK)
            kpi_box(k3,_fmt(oa),
                     "Overall Avg Quote",C_MID)
            kpi_box(k4,
                     df_exp_dummy["Service"].nunique(),
                     "Services",C_GREY_DARK)
            st.markdown("<br>",unsafe_allow_html=True)

            bc_va=[C_ORANGE if i==0
                    else C_DARK if i==len(vt_dd)-1
                    else C_GREY_DARK
                    for i in range(len(vt_dd))]
            fig_va=go.Figure(go.Bar(
                x=vt_dd["Vendor"],
                y=vt_dd["Average"],
                marker_color=bc_va,
                marker_line_width=0,
                text=vt_dd["Average"].apply(_fmt),
                textposition="outside"))
            fig_va.add_hline(
                y=oa,line_dash="dash",
                line_color=C_MID,line_width=2,
                annotation_text="Avg: {}".format(
                    _fmt(oa)),
                annotation_position="top right")
            pwc_bar(fig_va,
                     "Average Quote per Vendor",
                     height=360)
            fig_va.update_xaxes(tickangle=-20)
            st.plotly_chart(fig_va,
                             use_container_width=True)

            sec("VENDOR RANKING TABLE")
            vtbl=["<table class='comp-table'>"
                  "<thead><tr>"
                  "<th>Rank</th><th>Vendor</th>"
                  "<th>Quotes</th><th>Avg</th>"
                  "<th>Min</th><th>Max</th>"
                  "<th>vs Avg</th><th>Verdict</th>"
                  "</tr></thead><tbody>"]
            for rank,(_,vr) in enumerate(
                    vt_dd.iterrows(),start=1):
                bg=("white" if rank%2==0
                     else "#F8F8F8")
                vc=vcmap_dummy.get(vr["Vendor"],C_DARK)
                pct=round(
                    (vr["Average"]-oa)/oa*100,1
                ) if oa>0 else 0
                pc=(C_ORANGE if pct<-5
                     else C_DARK if pct>5
                     else C_GREY_DARK)
                pt=("{}% below".format(abs(pct))
                     if pct<0
                     else "{}% above".format(abs(pct))
                     if pct>0 else "At avg")
                ov=(("✅ COMPETITIVE",C_ORANGE)
                     if pct<-10
                     else ("🔴 EXPENSIVE",C_DARK)
                     if pct>10
                     else ("🟡 AVERAGE",C_GREY_DARK))
                medal=("🥇" if rank==1
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
                    "<td style='font-family:monospace;"
                    "font-weight:700;color:#D04A02'>"
                    "{}</td>"
                    "<td style='font-family:monospace;"
                    "color:#D04A02'>{}</td>"
                    "<td style='font-family:monospace;"
                    "color:#2D2D2D'>{}</td>"
                    "<td style='color:{}'>{}</td>"
                    "<td style='color:{};"
                    "font-weight:700'>{}</td>"
                    "</tr>".format(
                        bg,medal,vpill(vr["Vendor"],vc),
                        int(vr["Quotes"]),
                        _fmt(vr["Average"]),
                        _fmt(vr["Min"]),
                        _fmt(vr["Max"]),
                        pc,pt,ov[1],ov[0]))
            vtbl.append("</tbody></table>")
            st.markdown("".join(vtbl),
                         unsafe_allow_html=True)

            # Per-service
            st.markdown("<br>",unsafe_allow_html=True)
            sec("PER-SERVICE PRICE BENCHMARKING",
                "Services with multiple vendor quotes")
            svc_rows=[]
            for _,r in df_dummy.iterrows():
                svcs_raw=str(r.get("Comments","")
                              ).replace("\\n","\n"
                                        ).replace("\r\n","\n"
                                                   ).replace("\r","\n")
                ss=[s.strip() for s in svcs_raw.split("\n")
                    if s.strip()
                    and s.strip() not in ["nan","None",""]]
                if not ss: ss=[svcs_raw.strip()]
                for s in ss:
                    svc_rows.append({
                        "Service":s,
                        "Vendor":r["Vendor"],
                        "Price":_parse_num(str(r.get(
                            "Quoted Price","")).strip())})
            df_sv=pd.DataFrame(svc_rows)
            df_sv=df_sv[df_sv["Price"]>0]
            svc_vc=(df_sv.groupby("Service")[
                "Vendor"].nunique())
            multi=(svc_vc[svc_vc>1].index.tolist())

            if not multi:
                st.info("No multi-vendor services.")
            else:
                insight(
                    "<b>{}</b> services with "
                    "multi-vendor quotes.".format(
                        len(multi)))
                for svc in sorted(multi)[:10]:
                    ds=(df_sv[df_sv["Service"]==svc]
                        .sort_values("Price"))
                    mn=ds["Price"].min()
                    mx=ds["Price"].max()
                    av=ds["Price"].mean()
                    bv=ds.loc[ds["Price"].idxmin(),"Vendor"]
                    wv=ds.loc[ds["Price"].idxmax(),"Vendor"]
                    sp=round((mx-mn)/mn*100,1) if mn>0 else 0
                    st.markdown(
                        "<div style='background:white;"
                        "border-left:4px solid {};"
                        "padding:10px 14px;"
                        "border-radius:2px;"
                        "margin:10px 0;"
                        "font-weight:700;"
                        "font-size:0.88em'>"
                        "{} · {} vendors · spread {}% · "
                        "best: {} @ {}</div>".format(
                            C_ORANGE,svc,
                            ds["Vendor"].nunique(),
                            sp,bv,_fmt(mn)),
                        unsafe_allow_html=True)
                    bc_s=[C_ORANGE if v==bv
                           else C_DARK if v==wv
                           else C_GREY_DARK
                           for v in ds["Vendor"]]
                    fig_s=go.Figure(go.Bar(
                        x=ds["Vendor"],y=ds["Price"],
                        marker_color=bc_s,
                        marker_line_width=0,
                        text=ds["Price"].apply(_fmt),
                        textposition="outside"))
                    fig_s.add_hline(
                        y=av,line_dash="dash",
                        line_color=C_MID,line_width=1.5,
                        annotation_text="Avg: {}".format(
                            _fmt(av)),
                        annotation_position="top right")
                    fig_s.update_layout(
                        height=240,plot_bgcolor=CBG,
                        paper_bgcolor=CBG,
                        margin=dict(l=5,r=10,t=12,b=8),
                        font=CFONT,
                        yaxis=dict(showgrid=True,
                                   gridcolor=C_GREY_LITE,
                                   zeroline=False),
                        bargap=0.4,showlegend=False)
                    st.plotly_chart(fig_s,
                                     use_container_width=True)

            st.markdown("<br>",unsafe_allow_html=True)
            st.download_button(
                "📥 Download Dummy Data CSV",
                data=df_dummy.to_csv(index=False),
                file_name="dummy_analysis.csv",
                mime="text/csv",
                type="primary")
                                       
