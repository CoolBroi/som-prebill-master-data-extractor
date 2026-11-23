from __future__ import annotations
from copy import copy
from io import BytesIO
from hashlib import sha256
from datetime import datetime, date
import calendar
import math
import numbers
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from rapidfuzz import fuzz, utils

st.set_page_config(page_title="SOM Product Master Data Extractor", page_icon="📦", layout="wide")
st.markdown("""
<style>.block-container{max-width:1450px;padding-top:1.5rem}.hero{padding:1.4rem 1.7rem;border-radius:16px;background:linear-gradient(115deg,#14213d,#244b78);color:white;margin-bottom:1rem}.hero h1{color:white;margin:0 0 .35rem;font-size:2rem}.hero p{color:#dbeafe;margin:0}</style>
<div class="hero"><h1>SOM Product Master Data Extractor</h1><p>Match product/resource names from any Pre-Bill workbook to VBRP Item Descriptions and populate the Master Data Template.</p></div>
""", unsafe_allow_html=True)

MONTHS={m.lower():i for i,m in enumerate(calendar.month_name) if m}
MONTHS.update({m.lower():i for i,m in enumerate(calendar.month_abbr) if m})
MONTH_WORDS=set(MONTHS)
GENERIC_STOP={"compute","support","service","services","management","license","licenses","charge","charges","the","and","for","of","on","to","month","monthly","one","time","billing","os","instance","years","year","committed","resource","unit","ru"}
LOC={"adc","hdc","ldc","rdc"}
MONTH_RE=re.compile(r"\b(january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sept|sep|october|oct|november|nov|december|dec)[\s\-_/.,]*(19\d{2}|20\d{2}|\d{2})\b",re.I)

def text(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except (TypeError,ValueError): pass
    return str(v).strip()

def norm(v):
    return re.sub(r"\s+"," ",re.sub(r"[^a-z0-9]+"," ",text(v).lower().replace("&"," and "))).strip()

def tokens(v):
    # Numeric tokens are deliberately retained (e.g. Linux 8 vs Linux 9).
    return {x for x in norm(v).split() if x not in GENERIC_STOP and x not in LOC and x not in MONTH_WORDS and not re.fullmatch(r"(?:19|20)\\d{2}",x)}

def find_col(cols, aliases):
    normalized=[norm(c) for c in cols]
    for a in aliases:
        target=norm(a)
        for i,n in enumerate(normalized):
            if n==target:return cols[i]
    for a in aliases:
        target=norm(a)
        for i,n in enumerate(normalized):
            if target and target in n:return cols[i]
    return None

def all_sheets(upload):
    return pd.read_excel(BytesIO(upload.getvalue()),sheet_name=None,header=None,dtype=object)

def detect_sheet(sheets, groups, preferred=None):
    order=list(sheets)
    if preferred in order:order=[preferred]+[s for s in order if s!=preferred]
    for sn,raw in ((s,sheets[s]) for s in order):
        for hi in range(min(80,len(raw))):
            vals=[norm(x) for x in raw.iloc[hi].tolist()]
            if all(any(norm(alias)==v or norm(alias) in v for v in vals for alias in group) for group in groups):
                return sn,hi
    raise ValueError("Could not detect a worksheet/header containing the required columns: "+", ".join("/".join(g) for g in groups))

def month_key(value):
    if isinstance(value,(pd.Timestamp,datetime,date)):
        try:return (int(value.year),int(value.month))
        except Exception: return None
    s=text(value)
    if not s:return None
    m=MONTH_RE.search(s)
    if m:
        month=MONTHS[m.group(1).lower()]
        year=int(m.group(2))
        if year<100:year=2000+year if year<70 else 1900+year
        return (year,month)
    # ISO-like dates and Excel serials already decoded by pandas.
    try:
        dt=pd.to_datetime(s,errors="coerce",dayfirst=False)
        if not pd.isna(dt) and re.search(r"\b(19|20)\d{2}\b",s):
            return (int(dt.year),int(dt.month))
    except Exception: pass
    return None

def month_label(key):
    return datetime(key[0],key[1],1).strftime("%b %Y")

def months_in_sheets(sheets):
    found=set()
    for sn,raw in sheets.items():
        k=month_key(sn)
        if k:found.add(k)
        for row in raw.itertuples(index=False,name=None):
            for v in row:
                k=month_key(v)
                if k:found.add(k)
    return sorted(found)

def month_col_indices(header):
    return [(i,month_key(v)) for i,v in enumerate(header) if month_key(v)]

def table_end(raw, header_idx, product_col):
    blanks=0; end=len(raw)
    for ri in range(header_idx+1,len(raw)):
        vals=[text(x) for x in raw.iloc[ri].tolist()]
        product=text(raw.iat[ri,product_col]) if product_col<len(raw.columns) else ""
        rowtxt=" ".join(vals).lower()
        if not any(vals):
            blanks+=1
            if blanks>=2:
                end=ri-1;break
            continue
        blanks=0
        if re.search(r"\b(grand\s+total|sub\s*total|subtotal|total\s+amount|total\s+qty|notes?\s*:|footnote)\b",rowtxt):
            end=ri;break
        if not product:
            blanks+=1
            if blanks>=2:
                end=ri-1;break
        else:
            blanks=0
        if product and re.match(r"^(notes?|remarks?|source:|prepared by|continued)\b",product,re.I):
            end=ri;break
    return max(header_idx+1,end)

def header_bonus(header):
    h=norm(header)
    terms={"resource":1.8,"product":1.8,"service":1.5,"item":1.2,"description":1.2,"particular":1.0,"offering":.8,"deliverable":.8,"ru":.6}
    return min(2.5,sum(v for k,v in terms.items() if k in h))

def header_row_bonus(headers):
    recognized=("resource","product","service","item","description","particular","offering","deliverable","ru","period","month","quantity","qty","volume","baseline","uom","unit")
    count=sum(any(k in norm(h) for k in recognized) for h in headers if text(h))
    return 2.0 if count>=2 else 0.0

def content_column_score(values,header,headers=None):

    vals=[text(x) for x in values if text(x)]
    if len(vals)<3:return -100
    cleaned=[v for v in vals if not re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?",v)]
    text_ratio=len(cleaned)/len(vals)
    uniq=len({norm(v) for v in cleaned})/max(1,len(cleaned))
    avg=sum(len(v) for v in cleaned)/max(1,len(cleaned))
    moderate=max(0,1-abs(avg-28)/45)
    score=6*text_ratio+3*uniq+2*moderate+header_bonus(header)+(header_row_bonus(headers or []) if headers is not None else 0)
    h=norm(header)
    if any(k in h for k in ["id","code","number","serial","date","qty","quantity","amount","rate","price","currency","email","phone"]):score-=5
    if any(k in h for k in ["summary","total","inventory","category","department","tower","track","status","owner"]):score-=3
    if len(cleaned)<3:score-=4
    return score

def quantity_column_guess(headers):
    keys=["baseline volume","baseline qty","baseline quantity","quantity","qty","volume","units","number of"]
    for i,h in enumerate(headers):
        if any(k in norm(h) for k in keys):return i
    return None

def detect_candidates(sheets):
    candidates=[]
    for sn,raw in sheets.items():
        max_header=min(80,len(raw))
        for hi in range(max_header):
            headers=[text(x) for x in raw.iloc[hi].tolist()]
            if sum(bool(x) for x in headers)<2:continue
            month_cols=month_col_indices(headers)
            for ci,h in enumerate(headers):
                end=table_end(raw,hi,ci)
                vals=raw.iloc[hi+1:end,ci].tolist()
                score=content_column_score(vals,h,headers)
                if score<6.0:continue
                # A header row should have a reasonable table beneath it.
                nonblank=sum(bool(text(x)) for x in vals)
                if nonblank<3:continue
                qty=quantity_column_guess(headers)
                candidates.append({"sheet":sn,"header":hi,"col":ci,"qty":qty,"score":round(score,2),"end":end,"months":month_cols})
    if not candidates:
        raise ValueError("Could not detect a product/resource table. Please use the Detection controls to inspect the workbook structure.")
    # Collapse candidates with same sheet/header and retain column-level alternatives.
    candidates.sort(key=lambda x:x["score"],reverse=True)
    return candidates

def choose_candidate(candidates,sheet,header,product_col):
    for c in candidates:
        if c["sheet"]==sheet and c["header"]==header and c["col"]==product_col:return c
    raise ValueError("The selected sheet/header/product column does not form a detected table.")

def infer_month_column(raw,header_idx,headers):
    for i,h in enumerate(headers):
        if norm(h) in {"month","period","billing month","billing period","pricing month"} or "period" in norm(h):
            return i
    return None

def rows_for_candidate(raw,candidate,month_choice):
    hi=candidate["header"];ci=candidate["col"];end=table_end(raw,hi,ci)
    headers=[text(x) for x in raw.iloc[hi].tolist()]
    wide=month_col_indices(headers)
    period_col=infer_month_column(raw,hi,headers)
    qty_col=candidate.get("qty")
    selected_month=None if month_choice=="All / single period" else next((k for k in months_in_sheets({candidate["sheet"]:raw}) if month_label(k)==month_choice),None)
    if wide:
        if selected_month is None:selected_month=max(k for _,k in wide)
        chosen=next(((i,k) for i,k in wide if k==selected_month),None)
        if chosen:
            qty_col=chosen[0]
    output=[];duplicates=[]
    for ri in range(hi+1,end):
        name=text(raw.iat[ri,ci])
        if not name:continue
        rowtxt=" ".join(text(x).lower() for x in raw.iloc[ri].tolist())
        if re.search(r"\b(grand\s+total|sub\s*total|subtotal|total\s+amount|continued|notes?:|footnote)",rowtxt):continue
        if period_col is not None and selected_month is not None:
            rk=month_key(raw.iat[ri,period_col])
            if rk!=selected_month:continue
        qty=None
        if qty_col is not None and qty_col<len(raw.columns):
            val=raw.iat[ri,qty_col]
            try:
                num=pd.to_numeric(val,errors="coerce")
                qty=None if pd.isna(num) else float(num) if isinstance(num,numbers.Number) else None
                if qty is not None and qty.is_integer():qty=int(qty)
            except Exception:qty=None
        output.append({"ru":name,"qty":qty,"source_sheet":candidate["sheet"],"source_row":ri+1,"source_column":headers[ci] or f"Column {ci+1}"})
    return output

def deduplicate_products(rows):
    kept=[];dupes=[];seen={}
    for p in rows:
        key=norm(p["ru"])
        if key not in seen:
            seen[key]=p;kept.append(p)
        else:
            original=seen[key]
            conflict=original.get("qty")!=p.get("qty")
            dupes.append({**p,"duplicate_of_row":original.get("source_row"),"quantity_conflict":conflict,"original_qty":original.get("qty")})
    return kept,dupes

def tokens_match(v):
    return {x for x in norm(v).split() if x not in GENERIC_STOP and x not in LOC}

def alias_for_ru(ru,template):
    col=find_col(template.columns,["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"])
    if not col:return ru
    rt=tokens_match(ru);loc=set(norm(ru).split())&LOC;choices=[]
    for val in template[col].dropna():
        name=text(val);cloc=set(norm(name).split())&LOC
        if loc and cloc and loc!=cloc:continue
        ct=tokens_match(name)
        if rt and rt.issubset(ct):choices.append((len(ct-rt),name))
    return sorted(choices)[0][1] if choices else ru

def score_product(product,desc):
    pt=tokens_match(product);dt=tokens_match(desc)
    if not pt or not dt:return 0
    # Do not let distinct numeric specifications collapse into one match.
    pnums={x for x in norm(product).split() if x.isdigit() and not re.fullmatch(r"(?:19|20)[0-9]{2}",x)}
    dnums={x for x in norm(desc).split() if x.isdigit() and not re.fullmatch(r"(?:19|20)[0-9]{2}",x)}
    if pnums and not pnums.issubset(dnums):return 0
    pl=set(norm(product).split())&LOC;dl=set(norm(desc).split())&LOC
    if pl and dl and pl.isdisjoint(dl):return 0
    overlap=len(pt&dt)
    coverage=overlap/len(pt)
    precision=overlap/len(dt)
    if coverage<0.65:return 0
    # Penalize extra description tokens; long generic descriptions are not free matches.
    score=100*(0.72*coverage+0.28*precision)
    if len(dt)>len(pt)+3:score-=min(25,(len(dt)-len(pt)-3)*3)
    if len(pt)<=2 and coverage<1:return 0
    return round(max(0,score),1)

def parse_pricing_date(value):
    if isinstance(value,(pd.Timestamp,datetime,date)):
        return pd.Timestamp(value)
    s=text(value)
    if not s:return pd.NaT
    for dayfirst in (True,False):
        try:
            d=pd.to_datetime(s,errors="coerce",dayfirst=dayfirst)
            if not pd.isna(d):return d
        except Exception:pass
    return pd.NaT

def match_products(products,vbrp,template):
    desc_col=find_col(vbrp.columns,["Item Description"])
    date_col=find_col(vbrp.columns,["Pricing Date","Pricing date","Pricing Dt","Price Date"])
    if not desc_col:raise ValueError("VBRP Item Description column was not found.")
    v=vbrp[vbrp[desc_col].notna()].copy();v[desc_col]=v[desc_col].map(text)
    v["_pricing_date"]=v[date_col].map(parse_pricing_date) if date_col else pd.NaT
    descs=list(dict.fromkeys(x for x in v[desc_col] if x))
    out=[];unmatched=[]
    key_aliases=[["Sold-to","Customer ID"],["StatKF","Device ID"],["CoCd","Company Code"],["Project definition","Project Code"]]
    for p in products:
        alias=alias_for_ru(p["ru"],template)
        scored=[(d,score_product(alias,d),fuzz.ratio(alias,d,processor=utils.default_process)) for d in descs]
        scored=[x for x in scored if x[1]>0]
        if not scored:
            unmatched.append(p);continue
        scored.sort(key=lambda x:(x[1],x[2]),reverse=True)
        best=scored[0]
        near=[x for x in scored if best[1]-x[1]<=4 and x[0]!=best[0]]
        candidates=v[v[desc_col].isin([x[0] for x in [best]+near])].copy()
        latest=candidates["_pricing_date"].max()
        review_reasons=[]
        if len(near)>0:review_reasons.append("competing descriptions have similar scores")
        if date_col is None or candidates["_pricing_date"].isna().any():review_reasons.append("Pricing Date missing or unparsable")
        if pd.isna(latest):chosen=candidates.iloc[0]
        else:
            latest_rows=candidates[candidates["_pricing_date"]==latest]
            if len(latest_rows)>1:
                review_reasons.append("latest Pricing Date is tied")
            for aliases in key_aliases:
                col=find_col(vbrp.columns,aliases)
                if col and latest_rows[col].map(text).nunique(dropna=False)>1:
                    review_reasons.append(f"latest-date records differ in {col}")
            chosen=latest_rows.iloc[0]
        row=chosen.drop(labels=["_pricing_date"]).to_dict()
        status="Review required — "+"; ".join(dict.fromkeys(review_reasons)) if review_reasons else ("Exact description" if fuzz.ratio(alias,best[0],processor=utils.default_process)>=98 else ("Strong product-token match" if best[1]>=85 else "Review required — low-confidence match"))
        if best[1]<85 and not status.startswith("Review required"):status="Review required — low-confidence match"
        out.append({"ru":p["ru"],"alias":alias,"qty":p.get("qty"),"desc":text(chosen.get(desc_col)),"score":best[1],"status":status,"pricing_date":chosen.get("_pricing_date"),"vbrp":row})
    return out,unmatched

def sanitize(value):
    if value is None:return None
    try:
        if pd.isna(value):return None
    except (TypeError,ValueError):pass
    if hasattr(value,"item"):
        try:value=value.item()
        except Exception:pass
    if isinstance(value,(pd.Timestamp,datetime,date)):return value.to_pydatetime() if isinstance(value,pd.Timestamp) else value
    if isinstance(value,float) and (math.isnan(value) or math.isinf(value)):return None
    return value

def copy_style(ws,src,dst,n):
    for c in range(1,n+1):
        a,b=ws.cell(src,c),ws.cell(dst,c)
        if a.has_style:b._style=copy(a._style)
        if a.number_format:b.number_format=a.number_format
        if a.alignment:b.alignment=copy(a.alignment)
        if a.protection:b.protection=copy(a.protection)

def header_map(ws,header_row):
    return [(c,ws.cell(header_row,c).value,norm(ws.cell(header_row,c).value)) for c in range(1,ws.max_column+1)]

def find_header_index(headers,aliases):
    for alias in aliases:
        target=norm(alias)
        for i,h in enumerate(headers):
            if h==target:return i
    for alias in aliases:
        target=norm(alias)
        for i,h in enumerate(headers):
            if target and target in h:return i
    return None

def template_append_and_write(wb,template_sheet,header_row,template_df,matches,vbrp_columns):
    ws=wb[template_sheet];original_sheets=list(wb.sheetnames);original_max_col=ws.max_column
    header_cells=header_map(ws,header_row)
    normalized=[x[2] for x in header_cells]
    required={
      "SOW Product Name":["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"],
      "Customer ID":["Customer ID"],"Device ID":["Device ID"],"Company Code":["Company Code","CoCd"],
      "Project Code":["Project Code","Project definition"],"Output 2 Name":["Output 2 Name"],"Output 2 Value":["Output 2 Value"]}
    target={}
    missing=[]
    for field,aliases in required.items():
        idx=find_header_index(normalized,[norm(a) for a in aliases])
        if idx is None:missing.append(field)
        else:target[field]=header_cells[idx][0]
    key_col=target.get("SOW Product Name")
    if key_col is None:raise ValueError("Cannot determine template append key: SOW Product Name header is missing.")
    last_key=header_row
    for r in range(header_row+1,ws.max_row+1):
        if text(ws.cell(r,key_col).value):last_key=r
    append_row=last_key+1
    # Existing rows are mapped by the actual worksheet row number, not pandas index arithmetic.
    existing={}
    for r in range(header_row+1,last_key+1):
        val=text(ws.cell(r,key_col).value)
        if val:existing.setdefault(norm(val),r)
    sold=find_col(vbrp_columns,["Sold-to","Customer ID"]);stat=find_col(vbrp_columns,["StatKF","Device ID"])
    cocd=find_col(vbrp_columns,["CoCd","Company Code"]);proj=find_col(vbrp_columns,["Project definition","Project Code"])
    used=set();written=0;formula_notes=[]
    for m in matches:
        source=existing.get(norm(m["alias"])) or existing.get(norm(m["ru"])) or existing.get(norm(m["desc"]))
        if source and source not in used:outrow=source;used.add(source)
        else:
            outrow=append_row;append_row+=1
            ref_row=header_row+1
            if ref_row<=ws.max_row:
                copy_style(ws,ref_row,outrow,ws.max_column)
                for cc in range(1,ws.max_column+1):
                    formula=ws.cell(ref_row,cc).value
                    if isinstance(formula,str) and formula.startswith("="):
                        try:ws.cell(outrow,cc).value=Translator(formula,origin=ws.cell(ref_row,cc).coordinate).translate_formula(ws.cell(outrow,cc).coordinate)
                        except Exception:
                            ws.cell(outrow,cc).value=None;formula_notes.append(f"{ws.cell(ref_row,cc).coordinate}: formula left blank")
            else:
                formula_notes.append("No reference data row existed; new rows received no formulas.")
        values={
          "SOW Product Name":m["desc"],"Customer ID":m["vbrp"].get(sold) if sold else None,
          "Device ID":m["vbrp"].get(stat) if stat else None,"Company Code":m["vbrp"].get(cocd) if cocd else None,
          "Project Code":m["vbrp"].get(proj) if proj else None,"Output 2 Name":"QTY","Output 2 Value":m.get("qty")}
        for field,col in target.items():ws.cell(outrow,col).value=sanitize(values.get(field))
        written+=1
    return {"written":written,"missing":missing,"formula_notes":formula_notes,"sheetnames":original_sheets,"max_column":original_max_col,"last_key_row":last_key}

st.markdown("### Step 1 — Upload and review Pre-Bill products")
arc=st.file_uploader("Upload Pre-Bill workbook (any Excel layout)",type=["xlsx","xlsm"],key="prebill_upload")
if not arc:
    st.info("Upload the Pre-Bill first. VBRP is optional and is only needed in Step 2.")
    st.stop()

signature=sha256(arc.getvalue()).hexdigest()
if st.session_state.get("prebill_signature")!=signature:
    st.session_state["prebill_signature"]=signature
    for key in ["confirmed_products","prebill_editor","step2_ready","detection_state","duplicate_rows","unmatched_ack"]:
        st.session_state.pop(key,None)

try:
    asheets=all_sheets(arc)
    candidates=detect_candidates(asheets)
except Exception as e:
    st.error(f"Could not profile the Pre-Bill workbook: {e}")
    st.stop()

best=candidates[0]
if "detection_state" not in st.session_state:
    st.session_state["detection_state"]={"sheet":best["sheet"],"header":best["header"],"col":best["col"],"qty":best.get("qty"),"month":None}
state=st.session_state["detection_state"]
sheet_options=list(asheets.keys())
sel_sheet=st.selectbox("Sheet",sheet_options,index=sheet_options.index(state["sheet"]) if state["sheet"] in sheet_options else 0,key="detect_sheet")
headers_raw=asheets[sel_sheet]
header_options=list(range(min(80,len(headers_raw))))
sel_header=st.selectbox("Header row",header_options,index=min(state["header"],len(header_options)-1),format_func=lambda x:f"Excel row {x+1}",key="detect_header")
headers=[text(x) or f"Column {i+1}" for i,x in enumerate(headers_raw.iloc[sel_header].tolist())]
col_options=list(range(len(headers)))
default_col=state["col"] if sel_sheet==state["sheet"] and sel_header==state["header"] else max(range(len(headers)),key=lambda i:content_column_score(headers_raw.iloc[sel_header+1:min(len(headers_raw),sel_header+40),i].tolist(),headers[i]))
sel_col=st.selectbox("Product / RU column",col_options,index=min(default_col,len(col_options)-1),format_func=lambda i:f"{headers[i]} (Column {i+1})",key="detect_product_col")
qty_options=[None]+col_options
guessed_qty=quantity_column_guess(headers)
default_qty=state["qty"] if sel_sheet==state["sheet"] and sel_header==state["header"] else guessed_qty
qty_index=qty_options.index(default_qty) if default_qty in qty_options else 0
sel_qty=st.selectbox("Quantity column",qty_options,index=qty_index,format_func=lambda i:"Not detected" if i is None else f"{headers[i]} (Column {i+1})",key="detect_qty_col")
month_keys=months_in_sheets(asheets)
month_labels=[month_label(k) for k in month_keys]
if not month_labels:
    month_labels=["All / single period"]
    st.info("No month marker was found in sheet names, headers or cell values. The workbook will be treated as a single period.")
else:
    st.caption("Detected periods: "+", ".join(month_labels))
    month_labels=["Latest detected"]+month_labels
    default_month="Latest detected"
sel_month=st.selectbox("Month / Period",month_labels,index=0,key="detect_month")
if sel_month=="Latest detected" and month_keys:
    chosen_month=month_label(max(month_keys))
elif sel_month=="All / single period":
    chosen_month=sel_month
else:
    chosen_month=sel_month
selected_candidate={"sheet":sel_sheet,"header":sel_header,"col":sel_col,"qty":sel_qty,"score":content_column_score(headers_raw.iloc[sel_header+1:min(len(headers_raw),sel_header+40),sel_col].tolist(),headers[sel_col])}
end_row=table_end(headers_raw,sel_header,sel_col)
preview=headers_raw.iloc[sel_header:min(end_row,sel_header+16)].copy()
preview.columns=[text(x) or f"Column {i+1}" for i,x in enumerate(headers_raw.iloc[sel_header].tolist())]
st.markdown("#### Detection preview")
st.caption(f"Selected range: {sel_sheet}, Excel rows {sel_header+1}–{end_row}. Product-column confidence: {'High' if selected_candidate['score']>=9 else 'Medium' if selected_candidate['score']>=7 else 'Low'} (score {selected_candidate['score']:.1f}).")
preview_style=preview.style.apply(lambda col: ["background-color: rgba(40, 160, 100, 0.22)" if col.name==headers[sel_col] else "" for _ in col],axis=0)
st.dataframe(preview_style, use_container_width=True, height=300)
if len(candidates)>1 and candidates[0]["score"]-candidates[1]["score"]<=1.0:
    st.warning(f"Two candidate columns have similar scores ({candidates[0]['score']:.1f} and {candidates[1]['score']:.1f}). Please verify the highlighted product column before confirming.")
try:
    extracted=rows_for_candidate(headers_raw,selected_candidate,chosen_month)
    products,duplicates=deduplicate_products(extracted)
except Exception as e:
    st.error(f"Extraction failed: {e}")
    st.stop()
st.success(f"Extracted {len(products)} unique RU names from {sel_sheet}; {len(duplicates)} duplicate rows were detected.")
if duplicates:
    st.warning(f"{len(duplicates)} duplicate rows were excluded from the working list but are shown below. Conflicting quantities: {sum(bool(x['quantity_conflict']) for x in duplicates)}.")
    with st.expander("Review removed duplicate rows",expanded=False):
        st.dataframe(pd.DataFrame(duplicates),use_container_width=True)
        conflicts=[x for x in duplicates if x["quantity_conflict"]]
        if conflicts:st.error(f"{len(conflicts)} duplicates have conflicting quantities. Check which quantity should be retained.")
edit_df=pd.DataFrame([{"Product / Service Name":p["ru"],"Quantity / Volume":p["qty"],"Source Sheet":p["source_sheet"],"Source Row":p["source_row"],"Source Column":p["source_column"]} for p in products])
edited=st.data_editor(edit_df,use_container_width=True,height=420,num_rows="dynamic",key="prebill_editor_"+sha256(str((sel_sheet,sel_header,sel_col,sel_qty,chosen_month)).encode()).hexdigest()[:10],column_config={"Product / Service Name":st.column_config.TextColumn("Product / Service Name",required=True),"Quantity / Volume":st.column_config.NumberColumn("Quantity / Volume"),"Source Sheet":st.column_config.TextColumn("Source Sheet",disabled=True),"Source Row":st.column_config.NumberColumn("Source Row",disabled=True),"Source Column":st.column_config.TextColumn("Source Column",disabled=True)})
if st.button("Confirm product list",type="primary",use_container_width=True):
    clean=[];seen={}
    for _,row in edited.iterrows():
        name=text(row.get("Product / Service Name"));key=norm(name)
        if not name:continue
        if key in seen:continue
        seen[key]=True
        qty=sanitize(row.get("Quantity / Volume"))
        clean.append({"ru":name,"qty":qty,"source_sheet":text(row.get("Source Sheet")),"source_row":sanitize(row.get("Source Row")),"source_column":text(row.get("Source Column"))})
    if not clean:st.error("Please keep at least one RU name.")
    else:
        st.session_state["confirmed_products"]=clean
        st.session_state["step2_ready"]=False
        st.session_state.pop("unmatched_ack",None)
        st.rerun()
if "confirmed_products" not in st.session_state:st.stop()

products=st.session_state["confirmed_products"]
st.success(f"Product list confirmed: {len(products)} unique RU names.")
with st.expander("View confirmed RU list",expanded=False):
    st.dataframe(pd.DataFrame(products),use_container_width=True)
if st.button("Reset and review Pre-Bill products"):
    for key in ["confirmed_products","step2_ready","unmatched_ack"]:st.session_state.pop(key,None)
    st.rerun()
if not st.session_state.get("step2_ready",False):
    st.info("Your RU list is saved. Continue when you are ready to upload VBRP and the Master Data Template.")
    if st.button("Continue to Step 2 — Upload VBRP and Master Data Template",type="primary",use_container_width=True):
        st.session_state["step2_ready"]=True
        st.rerun()
    st.stop()

st.divider()
st.markdown("### Step 2 — Upload VBRP and Master Data Template")
a,b=st.columns(2)
with a:vbrp=st.file_uploader("Upload VBRP dump",type=["xlsx","xlsm"],key="vbrp_upload")
with b:tmpl=st.file_uploader("Upload Master Data Template",type=["xlsx","xlsm"],key="template_upload")
if not vbrp or not tmpl:
    st.info("Your confirmed RU list is saved. Upload both files to continue.")
    st.stop()
try:
    vsheets=all_sheets(vbrp);tsheets=all_sheets(tmpl)
    vn,vh=detect_sheet(vsheets,[["Item Description"],["Project definition"],["CoCd","Company Code"]])
    tn,th=detect_sheet(tsheets,[["SOW Product Name (Display text : this is Item Description on Invoice)","SOW Product Name"],["Customer ID"],["Company Code"]])
    vd=pd.read_excel(BytesIO(vbrp.getvalue()),sheet_name=vn,header=vh,dtype=object)
    td=pd.read_excel(BytesIO(tmpl.getvalue()),sheet_name=tn,header=th,dtype=object)
    matches,unmatched=match_products(products,vd,td)
except Exception as e:
    st.error(f"VBRP or template detection failed: {e}")
    st.stop()

st.markdown("### VBRP matching results")
x,y,z,w=st.columns(4)
review=sum(m["status"].startswith("Review required") for m in matches)
x.metric("Confirmed RU names",len(products));y.metric("VBRP matches",len(matches));z.metric("Unmatched",len(unmatched));w.metric("Needs review",review)
st.caption(f"VBRP sheet: {vn}; template sheet: {tn}. Latest Pricing Date is used only after candidate selection.")
if matches:
    st.dataframe(pd.DataFrame([{"Pre-Bill RU":m["ru"],"VBRP Item Description":m["desc"],"Pricing Date":m["pricing_date"],"Score":m["score"],"Status":m["status"]} for m in matches]),use_container_width=True,height=400)
else:st.warning("No RU names matched VBRP.")
if unmatched:
    st.warning(f"{len(unmatched)} RU names were not matched. Review them before proceeding.")
    st.dataframe(pd.DataFrame([{"Unmatched RU":p["ru"],"Quantity / Volume":p.get("qty"),"Source Sheet":p.get("source_sheet"),"Source Row":p.get("source_row")} for p in unmatched]),use_container_width=True)
    ack=st.checkbox("I acknowledge the unmatched RU names and want to generate a template using matched records only.",key="unmatched_ack")
else:
    ack=True
if review:
    st.warning(f"{review} matched records are marked Review required. Verify these before using the output.")
if not matches:
    st.stop()
if unmatched and not ack:
    st.info("Acknowledge the unmatched RU list to enable template generation.")
    st.stop()

if st.button("Generate Master Data Template",type="primary",use_container_width=True):
    keep_vba=tmpl.name.lower().endswith(".xlsm")
    wb=load_workbook(BytesIO(tmpl.getvalue()),keep_vba=keep_vba)
    original_sheets=list(wb.sheetnames)
    ws=wb[tn];original_header=[ws.cell(th+1,c).value for c in range(1,ws.max_column+1)]
    result=template_append_and_write(wb,tn,th+1,td,matches,list(vd.columns))
    # Structural verification after writing.
    unchanged=(list(wb.sheetnames)==original_sheets and [ws.cell(th+1,c).value for c in range(1,ws.max_column+1)]==original_header and ws.max_column==len(original_header))
    if not unchanged:st.error("Template structure verification failed. Do not use this output.");st.stop()
    buf=BytesIO();wb.save(buf)
    st.success(f"Generated {result['written']} matched records. Template sheet names, header row and column count were preserved.")
    if result["missing"]:st.error("Template columns not found: "+", ".join(result["missing"]))
    if result["formula_notes"]:st.warning("Formula handling: "+"; ".join(result["formula_notes"]))
    st.caption("New rows copy formatting only. Formulas from the reference row are translated where possible; failed translations are left blank.")
    st.warning("openpyxl may not preserve all embedded images, charts, slicers or advanced Excel objects. Review the output workbook.")
    if review:st.warning(f"{review} records require manual review before SAP upload.")
    mime="application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    st.download_button("Download completed Master Data Template",data=buf.getvalue(),file_name="Master_Data_Template_Product_Matched"+(".xlsm" if keep_vba else ".xlsx"),mime=mime,use_container_width=True)
    st.info("This tool prepares the workbook only; it does not create or update SAP records.")
