from __future__ import annotations
from copy import copy
from io import BytesIO
import re
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from rapidfuzz import fuzz, utils

st.set_page_config(page_title="SOM Product Master Data Extractor", page_icon="📦", layout="wide")
st.markdown("""
<style>.block-container{max-width:1450px;padding-top:1.5rem}.hero{padding:1.4rem 1.7rem;border-radius:16px;background:linear-gradient(115deg,#14213d,#244b78);color:white;margin-bottom:1rem}.hero h1{color:white;margin:0 0 .35rem;font-size:2rem}.hero p{color:#dbeafe;margin:0}</style>
<div class="hero"><h1>SOM Product Master Data Extractor</h1><p>Match product/resource names from any Pre-Bill workbook to VBRP Item Descriptions and populate the Master Data Template.</p></div>
""", unsafe_allow_html=True)

STOP={"compute","support","service","services","management","license","licenses","charge","charges","the","and","for","of","on","to","month","monthly","one","time","billing","os","instance","years","year","committed","resource","unit","ru","february","march","april","2026"}
LOC={"adc","hdc","ldc","rdc"}

def text(v):
    return "" if v is None or pd.isna(v) else str(v).strip()
def norm(v):
    return re.sub(r"\s+"," ",re.sub(r"[^a-z0-9]+"," ",text(v).lower().replace("&"," and "))).strip()
def tokens(v):
    return {x for x in norm(v).split() if not x.isdigit()}-STOP-LOC
def find_col(cols, aliases):
    ns={norm(c):c for c in cols}
    for a in aliases:
        if norm(a) in ns:return ns[norm(a)]
    for c in cols:
        if any(norm(a) in norm(c) for a in aliases):return c
    return None
def header_row(raw, groups, limit=15):
    best=(-1,None)
    for i in range(min(limit,len(raw))):
        vals=[norm(v) for v in raw.iloc[i].tolist() if text(v)]
        score=sum(any(norm(a)==v or norm(a) in v for v in vals for a in group) for group in groups)
        if score>best[0]:best=(score,i)
    return best[1] if best[0]==len(groups) else None
def all_sheets(upload):
    return pd.read_excel(BytesIO(upload.getvalue()),sheet_name=None,header=None,dtype=object)
def latest_month_sheet(sheets):
    months={"january":1,"jan":1,"february":2,"feb":2,"march":3,"mar":3,"april":4,"apr":4,"may":5,"june":6,"jun":6,"july":7,"jul":7,"august":8,"aug":8,"september":9,"sep":9,"sept":9,"october":10,"oct":10,"november":11,"nov":11,"december":12,"dec":12}
    candidates=[]
    for index,name in enumerate(sheets):
        words=re.findall(r"[a-z]+",text(name).lower())
        found=[months[w] for w in words if w in months]
        if found:
            years=re.findall(r"\\b(?:19|20)\\d{2}\\b",text(name))
            year=int(years[-1]) if years else 0
            candidates.append(((year,max(found),index),name))
    if not candidates:
        raise ValueError("Could not identify a month in any worksheet name. Please name the monthly tabs with the month (for example, CPO Tab July) so the latest month can be selected.")
    return max(candidates,key=lambda x:x[0])[1]
def detect_sheet(sheets, groups, preferred=None):
    order=list(sheets)
    if preferred:
        order=([s for s in order if norm(s)==norm(preferred)]+[s for s in order if norm(s)!=norm(preferred)])
    for s in order:
        h=header_row(sheets[s],groups)
        if h is not None:return s,h
    raise ValueError("Could not automatically detect a required worksheet/header. Check that the expected columns exist.")
def table(upload,sheet,h):
    d=pd.read_excel(BytesIO(upload.getvalue()),sheet_name=sheet,header=h,dtype=object).dropna(how="all").copy()
    d.columns=[str(c).strip() for c in d.columns]
    return d
def product_header_score(header):
    h=norm(header)
    if not h:return -10
    if any(x in h for x in ["serial","s no","sequence","line number","row number","quantity","qty","amount","rate","price","date","total","currency","uom","unit of measure"]):return -5
    if any(x in h for x in ["infrastructure services","business unit","service category","service line","tower","department","division"]):return -3
    if any(x in h for x in ["service area/domain","service area","service domain","product/service","service description"]):return 16
    terms={"resource unit":12,"resource name":11,"resource":7,"product name":12,"product":8,"service name":12,"service":7,"item description":11,"item name":11,"item":5,"description":7,"particulars":8,"scope of work":9,"work package":8,"offering":8,"component":6,"activity":6,"deliverable":7,"sku":7,"material description":8,"solution":5,"requirement":4}
    return max([v for k,v in terms.items() if k in h]+[-2])
def table_from_raw(raw,h):
    d=raw.iloc[h+1:].copy()
    d.columns=[str(x).strip() if text(x) else f"Unnamed_{i}" for i,x in enumerate(raw.iloc[h].tolist())]
    return d.dropna(how="all").copy()
def infer_product_column(raw,hi):
    sample=raw.iloc[hi+1:hi+min(35,len(raw))].copy()
    best=(-1,None)
    for j,h in enumerate(raw.iloc[hi].tolist()):
        vals=[text(x) for x in sample.iloc[:,j].tolist() if text(x)]
        if len(vals)<2:continue
        nonnumeric=sum(not re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?",v) for v in vals)/len(vals)
        unique=len({norm(v) for v in vals})/len(vals)
        avglen=sum(len(v) for v in vals)/len(vals)
        score=5*nonnumeric+3*unique+min(avglen,50)/50
        if any(k in norm(h) for k in ["id","code","number","date","qty","quantity","amount","rate","price"]):score-=4
        if score>best[0]:best=(score,j)
    return best[1] if best[0]>=6 else None
def detect_prebill(sheets):
    # Scan every worksheet and every plausible header row. Keep all credible
    # product/service columns instead of selecting only the single highest-scoring column.
    found=[]
    for sn,raw in sheets.items():
        sheet_found=[]
        fallback=[]
        for hi in range(min(60,len(raw))):
            headers=[text(x) for x in raw.iloc[hi].tolist()]
            if not any(headers): continue
            scores=[product_header_score(x) for x in headers]
            candidates=[j for j,s in enumerate(scores) if s>=5]
            if not candidates:
                inferred=infer_product_column(raw,hi)
                if inferred is not None: fallback.append((hi,inferred))
                continue
            qty_candidates=[j for j,h in enumerate(headers) if any(k in norm(h) for k in ["baseline volume","baseline qty","baseline quantity","quantity","qty","volume","units","no of","number of"])]
            for ci in candidates:
                vals=[text(x) for x in raw.iloc[hi+1:min(len(raw),hi+101),ci].tolist() if text(x)]
                if len(vals)<2: continue
                sheet_found.append({"sheet":sn,"header":hi,"col":ci,"qty":qty_candidates[0] if qty_candidates else None,"score":scores[ci]})
        # Use content-only inference once per sheet, and only when no recognizable
        # product header was found. This avoids treating ordinary data rows as headers.
        if not sheet_found and fallback:
            hi,ci=max(fallback,key=lambda pair: len([x for x in raw.iloc[pair[0]+1:min(len(raw),pair[0]+101),pair[1]].tolist() if text(x)]))
            sheet_found.append({"sheet":sn,"header":hi,"col":ci,"qty":None,"score":4})
        found.extend(sheet_found)
    if not found:
        raise ValueError("Could not identify product/resource columns in any worksheet. Check that the workbook contains product/service names.")
    # Remove duplicate detections of the same table/header/column.
    dedup={}
    for x in found: dedup[(x["sheet"],x["header"],x["col"])]=x
    return list(dedup.values())
def prebill_products(sheets, detections):
    out=[]
    seen_names=set()
    for d in detections:
        raw=sheets[d["sheet"]]
        headers=[text(x) for x in raw.iloc[d["header"]].tolist()]
        name_col=headers[d["col"]] or f"Column {d['col']+1}"
        qty_col=headers[d["qty"]] if d["qty"] is not None and d["qty"]<len(headers) else None
        for ri in range(d["header"]+1,len(raw)):
            name=text(raw.iat[ri,d["col"]])
            key=norm(name)
            if not name or len(key)<2: continue
            if any(x in key for x in ["grand total","sub total","subtotal","total amount","continued"]): continue
            # Keep only the first occurrence of each normalized product/service name.
            if key in seen_names: continue
            seen_names.add(key)
            qty=raw.iat[ri,d["qty"]] if d["qty"] is not None and d["qty"]<raw.shape[1] else None
            out.append({"ru":name,"qty":qty,"source_sheet":d["sheet"],"source_row":ri+1,"source_column":name_col})
    if not out: raise ValueError("No product/resource names were found in the detected columns.")
    return out
def alias_for_ru(ru, template):
    col=find_col(template.columns,["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"])
    if not col:return ru
    rt=tokens(ru); loc=set(norm(ru).split())&LOC; choices=[]
    for val in template[col].dropna():
        name=text(val); cloc=set(norm(name).split())&LOC
        if loc and cloc and loc!=cloc:continue
        ct=tokens(name)
        if rt and rt.issubset(ct):choices.append((len(ct-rt),name))
    return sorted(choices)[0][1] if choices else ru
def score_product(product, desc):
    pt=tokens(product); dt=tokens(desc)
    if not pt or not dt:return 0
    pl=set(norm(product).split())&LOC; dl=set(norm(desc).split())&LOC
    if pl and dl and pl.isdisjoint(dl):return 0
    coverage=len(pt&dt)/len(pt)
    if len(pt)<=2 and coverage<1:return 0
    if len(pt)==3 and coverage<2/3:return 0
    if len(pt)>=3:
        server=bool(pt&{"wintel","linux","aix","vxrail"})
        if coverage<(0.2 if server else 0.5):return 0
    return round(coverage*100,1)
def match_products(products,vbrp,template):
    desc_col=find_col(vbrp.columns,["Item Description"])
    date_col=find_col(vbrp.columns,["Pricing Date","Pricing date","Pricing Dt","Price Date"])
    if not desc_col:raise ValueError("VBRP Item Description column was not found.")
    v=vbrp[vbrp[desc_col].notna()].copy();v[desc_col]=v[desc_col].map(text)
    v["_pricing_date"]=pd.to_datetime(v[date_col],errors="coerce",dayfirst=True) if date_col else pd.NaT
    descs=list(dict.fromkeys(x for x in v[desc_col] if x));out=[]
    for p in products:
        alias=alias_for_ru(p["ru"],template)
        scored=[(d,score_product(alias,d),fuzz.ratio(alias,d,processor=utils.default_process)) for d in descs]
        scored=[x for x in scored if x[1]>0]
        if not scored:continue
        exact=[x for x in scored if x[2]>=98]
        if exact:scored=exact
        best_score=max(x[1] for x in scored)
        best_exact=max(x[2] for x in scored if x[1]==best_score)
        top=[x for x in scored if x[1]==best_score and x[2]==best_exact]
        candidate_descs={x[0] for x in top}
        candidates=v[v[desc_col].isin(candidate_descs)].copy()
        candidates=candidates.sort_values("_pricing_date",ascending=False,na_position="last",kind="stable")
        chosen=candidates.iloc[0]
        desc=text(chosen.get(desc_col));score=best_score;exactness=best_exact
        if len(candidate_descs)>1:
            status="Review required — competing descriptions; latest Pricing Date selected"
        else:
            status="Exact description" if exactness>=98 else ("Review required" if score<90 else "Strong product-token match")
        row=chosen.drop(labels=["_pricing_date"]).to_dict()
        out.append({"ru":p["ru"],"alias":alias,"qty":p["qty"],"desc":desc,"score":score,"status":status,"pricing_date":chosen.get("_pricing_date"),"vbrp":row})
    return out
def copy_style(ws,src,dst,n):
    for c in range(1,n+1):
        a,b=ws.cell(src,c),ws.cell(dst,c)
        if a.has_style:b._style=copy(a._style)
        if a.number_format:b.number_format=a.number_format
        if a.alignment:b.alignment=copy(a.alignment)
        if a.protection:b.protection=copy(a.protection)

st.markdown("### Step 1 — Upload and review Pre-Bill products")
arc=st.file_uploader("Upload Pre-Bill workbook (any Excel layout)",type=["xlsx","xlsm"],key="prebill_upload")
if not arc:
    st.info("Start by uploading only the Pre-Bill workbook. VBRP is not needed at this stage.")
    st.stop()

signature=(arc.name,len(arc.getvalue()))
if st.session_state.get("prebill_signature")!=signature:
    st.session_state["prebill_signature"]=signature
    st.session_state.pop("confirmed_products",None)
    st.session_state.pop("prebill_editor",None)

if "confirmed_products" not in st.session_state:
    try:
        asheets=all_sheets(arc)
        latest_sheet=latest_month_sheet(asheets)
        latest_sheets={latest_sheet:asheets[latest_sheet]}
        detections=detect_prebill(latest_sheets)
        extracted=prebill_products(latest_sheets,detections)
    except Exception as e:
        st.error(f"Could not extract products from this Pre-Bill: {e}")
        st.stop()
    st.success(f"Detected {len(extracted)} unique product/service names from the latest month worksheet: {latest_sheet}. Duplicate names are removed.")
    st.caption("Only the latest month worksheet is scanned. Duplicate names are removed case-insensitively after normalizing spaces and punctuation. Source row and column are shown for traceability.")
    edit_df=pd.DataFrame([{"Product / Service Name":p["ru"],"Quantity / Volume":p["qty"],"Source Sheet":p["source_sheet"],"Source Row":p["source_row"],"Source Column":p["source_column"]} for p in extracted])
    edited=st.data_editor(edit_df,use_container_width=True,height=420,num_rows="dynamic",key="prebill_editor",column_config={"Product / Service Name":st.column_config.TextColumn("Product / Service Name",required=True),"Quantity / Volume":st.column_config.TextColumn("Quantity / Volume"),"Source Sheet":st.column_config.TextColumn("Source Sheet",disabled=True),"Source Row":st.column_config.NumberColumn("Source Row",disabled=True),"Source Column":st.column_config.TextColumn("Source Column",disabled=True)})
    if st.button("Confirm product list",type="primary",use_container_width=True):
        clean=[]
        confirmed_seen=set()
        for _,row in edited.iterrows():
            name=text(row.get("Product / Service Name"))
            key=norm(name)
            if not name or key in confirmed_seen: continue
            confirmed_seen.add(key)
            clean.append({"ru":name,"qty":row.get("Quantity / Volume"),"source_sheet":text(row.get("Source Sheet")),"source_row":row.get("Source Row"),"source_column":text(row.get("Source Column"))})
        if not clean:
            st.error("Please keep at least one product/service name.")
        else:
            st.session_state["confirmed_products"]=clean
            st.rerun()
    st.stop()

products=st.session_state["confirmed_products"]
st.success(f"Product list confirmed: {len(products)} products/services.")
with st.expander("View confirmed product list",expanded=False):
    st.dataframe(pd.DataFrame([{"Product / Service Name":p["ru"],"Quantity / Volume":p.get("qty"),"Source Sheet":p.get("source_sheet"),"Source Row":p.get("source_row"),"Source Column":p.get("source_column")} for p in products]),use_container_width=True)
if st.button("Reset and review Pre-Bill products"):
    st.session_state.pop("confirmed_products",None)
    st.session_state.pop("prebill_editor",None)
    st.rerun()

st.divider()
st.markdown("### Step 2 — Upload VBRP and Master Data Template")
a,b=st.columns(2)
with a:vbrp=st.file_uploader("Upload VBRP dump",type=["xlsx","xlsm"],key="vbrp_upload")
with b:tmpl=st.file_uploader("Upload Master Data Template",type=["xlsx","xlsm"],key="template_upload")
if not vbrp or not tmpl:
    st.info("Your product list is saved. Upload VBRP and the Master Data Template to find the corresponding entries.")
    st.stop()

try:
    vsheets=all_sheets(vbrp);tsheets=all_sheets(tmpl)
    vn,vh=detect_sheet(vsheets,[["Item Description"],["Project definition"],["CoCd","Company Code"]])
    tn,th=detect_sheet(tsheets,[["SOW Product Name (Display text : this is Item Description on Invoice)","SOW Product Name"],["Customer ID"],["Company Code"]])
    vd=table(vbrp,vn,vh);td=table(tmpl,tn,th)
    matches=match_products(products,vd,td)
except Exception as e:
    st.error(f"VBRP or template detection failed: {e}")
    st.stop()

matched={norm(m["ru"]) for m in matches}
unmatched=[p["ru"] for p in products if norm(p["ru"]) not in matched]
cols=list(vd.columns)
key_cols=[["Sold-to","Customer ID"],["StatKF","Device ID"],["CoCd","Company Code"],["Project definition","Project Code"],["Material"]]
def match_key(m):
    r=m["vbrp"]
    return tuple(norm(x) for x in [m["ru"],m["desc"]]+[text(r.get(find_col(cols,a))) if find_col(cols,a) else "" for a in key_cols])
# Preserve every confirmed Pre-Bill source row, including repeated product names.
unique=matches
review=sum(m["status"].startswith("Review required") for m in unique)
high=len(unique)-review
st.markdown("### VBRP matching results")
x,y,z,w=st.columns(4)
x.metric("Confirmed products",len(products));y.metric("VBRP matches",len(unique));z.metric("Exact / strong matches",high);w.metric("Needs review",review)
st.caption(f"VBRP sheet: {vn}; template sheet: {tn}. Every confirmed Pre-Bill row is retained. For each row, one VBRP entry is selected using the latest Pricing Date.")
st.dataframe(pd.DataFrame([{"Pre-Bill Product / Service Name":m["ru"],"VBRP Item Description":m["desc"],"Pricing Date":m.get("pricing_date"),"Match Score":m["score"],"Match Status":m["status"]} for m in unique]),use_container_width=True,height=400)
if unmatched:st.warning(f"{len(unmatched)} confirmed product(s) have no VBRP match: "+", ".join(unmatched))
if not unique:
    st.error("No confirmed products matched VBRP. Your full confirmed product list is retained above; no template records can be generated until a match is found.")
    st.stop()

if st.button("Generate Master Data Template",type="primary",use_container_width=True):
    keep_vba=tmpl.name.lower().endswith(".xlsm")
    wb=load_workbook(BytesIO(tmpl.getvalue()),keep_vba=keep_vba)
    ws=wb[tn]
    hr=th+1
    headers=[text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)]
    hmap={norm(h):i+1 for i,h in enumerate(headers) if h}
    first=hr+1
    sow=find_col(td.columns,["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"])
    rowmap={}
    if sow:
        for ix,r in td.iterrows():
            n=text(r.get(sow))
            if n:rowmap.setdefault(norm(n),hr+1+int(ix))
    used=set();nextrow=ws.max_row+1;written=0
    sold=find_col(cols,["Sold-to","Customer ID"]);stat=find_col(cols,["StatKF","Device ID"])
    cocd=find_col(cols,["CoCd","Company Code"]);proj=find_col(cols,["Project definition","Project Code"])
    for m in unique:
        r=m["vbrp"]
        source=rowmap.get(norm(m["alias"])) or rowmap.get(norm(m["ru"])) or rowmap.get(norm(m["desc"]))
        if source and source not in used:
            outrow=source;used.add(source)
        else:
            outrow=nextrow;nextrow+=1;style=first
            copy_style(ws,style,outrow,ws.max_column)
            for cc in range(1,ws.max_column+1):ws.cell(outrow,cc).value=ws.cell(style,cc).value
        def put(names,value):
            h=find_col(headers,names)
            if h:ws.cell(outrow,hmap[norm(h)]).value=value
        put(["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"],m["desc"])
        put(["Customer ID"],r.get(sold) if sold else None)
        put(["Device ID"],r.get(stat) if stat else None)
        put(["Company Code","CoCd"],r.get(cocd) if cocd else None)
        put(["Project Code","Project definition"],r.get(proj) if proj else None)
        put(["Output 2 Name"],"QTY");put(["Output 2 Value"],m["qty"])
        written+=1
    buf=BytesIO();wb.save(buf)
    st.success(f"Generated {written} matched product records in the original Master Data Template format.")
    if review:st.warning(f"{review} selected VBRP matches need review before SAP upload.")
    st.download_button("Download completed Master Data Template",data=buf.getvalue(),file_name="Master_Data_Template_Product_Matched"+(".xlsm" if keep_vba else ".xlsx"),mime="application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.info("This tool prepares the workbook only; it does not create or update SAP records.")
