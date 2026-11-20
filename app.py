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
<div class="hero"><h1>SOM Product Master Data Extractor</h1><p>Match ARC RRC Resource Unit names to VBRP Item Descriptions and populate the Master Data Template.</p></div>
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
def arc_products(df):
    ru=find_col(df.columns,["Resource Unit (RU)","Resource Unit"])
    sn=find_col(df.columns,["S.No.","Serial No","S No"])
    qty=find_col(df.columns,["Baseline Volume","Baseline Qty","Baseline Quantity"])
    if not ru:raise ValueError("ARC RRC Resource Unit (RU) column was not found.")
    if sn:df=df[pd.to_numeric(df[sn],errors="coerce").notna()]
    out=[];seen=set()
    for _,r in df.iterrows():
        name=text(r.get(ru)); key=norm(name)
        if name and key not in seen:
            out.append({"ru":name,"qty":r.get(qty) if qty else None});seen.add(key)
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
    date_col=find_col(vbrp.columns,["Pricing Date","Pricing date","Price Date"])
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

st.markdown("### Upload source workbooks")
a,b,c=st.columns(3)
with a: arc=st.file_uploader("1. ARC RRC workbook",type=["xlsx","xlsm"],key="arc")
with b: vbrp=st.file_uploader("2. VBRP dump",type=["xlsx","xlsm"],key="vbrp")
with c: tmpl=st.file_uploader("3. Master Data Template",type=["xlsx","xlsm"],key="template")
if not all([arc,vbrp,tmpl]):
    st.info("Upload all three files. Worksheets, headers and product matching are detected automatically.");st.stop()
try:
    asheets=all_sheets(arc);vsheets=all_sheets(vbrp);tsheets=all_sheets(tmpl)
    an,ah=detect_sheet(asheets,[["S.No.","Serial No"],["Resource Unit (RU)","Resource Unit"]],"ARC RRC")
    vn,vh=detect_sheet(vsheets,[["Item Description"],["Project definition"],["CoCd","Company Code"]])
    tn,th=detect_sheet(tsheets,[["SOW Product Name (Display text : this is Item Description on Invoice)","SOW Product Name"],["Customer ID"],["Company Code"]])
    ad=table(arc,an,ah);vd=table(vbrp,vn,vh);td=table(tmpl,tn,th)
    products=arc_products(ad);matches=match_products(products,vd,td)
except Exception as e:
    st.error(f"Automatic detection failed: {e}");st.stop()

matched={norm(m["ru"]) for m in matches};unmatched=[p["ru"] for p in products if norm(p["ru"]) not in matched]
st.markdown("### Automatic detection")
x,y,z,w=st.columns(4);x.metric("ARC RRC sheet",an);y.metric("VBRP sheet",vn);z.metric("Template sheet",tn);w.metric("RU product names",len(products))
st.caption(f"Detected header rows: ARC RRC {ah+1}; VBRP {vh+1}; Template {th+1}.")
if not matches:st.error("No plausible VBRP matches were found.");st.stop()
cols=list(vd.columns)
key_cols=[["Sold-to","Customer ID"],["StatKF","Device ID"],["CoCd","Company Code"],["Project definition","Project Code"],["Material"]]
def match_key(m):
    r=m["vbrp"];return tuple(norm(x) for x in [m["ru"],m["desc"]]+[text(r.get(find_col(cols,a))) if find_col(cols,a) else "" for a in key_cols])
unique=[];seen=set()
for m in matches:
    k=match_key(m)
    if k not in seen:seen.add(k);unique.append(m)
review=sum(m["status"].startswith("Review required") for m in unique);high=len(unique)-review
x,y,z=st.columns(3);x.metric("Unique VBRP product records",len(unique));y.metric("Exact / strong matches",high);z.metric("Needs review",review)
st.dataframe(pd.DataFrame([{"ARC RRC RU Name":m["ru"],"VBRP Item Description":m["desc"],"Pricing Date":m.get("pricing_date"),"Match Score":m["score"],"Match Status":m["status"]} for m in unique]),use_container_width=True,height=400)
if unmatched:st.warning("No VBRP match detected for: "+", ".join(unmatched))
st.caption("The existing template's SOW Product Name is used as an approved alias when its product terms clearly match the RU. Otherwise, the RU name is matched directly. Ambiguous matches are flagged; unmatched products are listed in Match_Audit.")

if st.button("Generate Master Data Template",type="primary",use_container_width=True):
    keep_vba=tmpl.name.lower().endswith(".xlsm");wb=load_workbook(BytesIO(tmpl.getvalue()),keep_vba=keep_vba);ws=wb[tn]
    hr=th+1;headers=[text(ws.cell(hr,c).value) for c in range(1,ws.max_column+1)];hmap={norm(h):i+1 for i,h in enumerate(headers) if h}
    first=hr+1;sow=find_col(td.columns,["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"])
    rowmap={}
    if sow:
        for ix,r in td.iterrows():
            n=text(r.get(sow))
            if n:rowmap.setdefault(norm(n),hr+1+int(ix))
    used=set();nextrow=ws.max_row+1;written=0;audit=[]
    tower=hmap.get(norm(find_col(headers,["Tower"]) or ""));track=hmap.get(norm(find_col(headers,["Track"]) or ""))
    for rr in range(first,ws.max_row+1):
        if tower:ws.cell(rr,tower).value=None
        if track:ws.cell(rr,track).value=None
    for m in unique:
        r=m["vbrp"];source=rowmap.get(norm(m["alias"])) or rowmap.get(norm(m["ru"])) or rowmap.get(norm(m["desc"]))
        if source and source not in used:outrow=source;used.add(source)
        else:
            outrow=nextrow;nextrow+=1;style=source or first
            copy_style(ws,style,outrow,ws.max_column)
            for cc in range(1,ws.max_column+1):ws.cell(outrow,cc).value=ws.cell(style,cc).value
        def put(names,value):
            h=find_col(headers,names)
            if h:ws.cell(outrow,hmap[norm(h)]).value=value
        sold=find_col(cols,["Sold-to","Customer ID"]);stat=find_col(cols,["StatKF","Device ID"]);cocd=find_col(cols,["CoCd","Company Code"]);proj=find_col(cols,["Project definition","Project Code"]);mat=find_col(cols,["Material"]);bill=find_col(cols,["Bill. Doc.","Billing Document"])
        put(["SOW Product Name","SOW Product Name (Display text : this is Item Description on Invoice)"],m["desc"])
        put(["Customer ID"],r.get(sold) if sold else None);put(["Device ID"],r.get(stat) if stat else None)
        put(["Company Code","CoCd"],r.get(cocd) if cocd else None);put(["Project Code","Project definition"],r.get(proj) if proj else None)
        put(["Tower"],None);put(["Track"],None);put(["Output 2 Name"],"QTY");put(["Output 2 Value"],m["qty"])
        audit.append({"ARC RRC RU Name":m["ru"],"Lookup Product Label":m["alias"],"Template Output Row":outrow,"Baseline Volume":m["qty"],"VBRP Item Description":m["desc"],"Match Score":m["score"],"Match Status":m["status"],"Customer ID (Sold-to)":r.get(sold) if sold else None,"Device ID (StatKF)":r.get(stat) if stat else None,"Company Code (CoCd)":r.get(cocd) if cocd else None,"Project Code":r.get(proj) if proj else None,"Material":r.get(mat) if mat else None,"Billing Document":r.get(bill) if bill else None})
        written+=1
    for ru in unmatched:audit.append({"ARC RRC RU Name":ru,"Match Status":"No VBRP match"})
    if "Match_Audit" in wb.sheetnames:del wb["Match_Audit"]
    aw=wb.create_sheet("Match_Audit");adf=pd.DataFrame(audit)
    for j,h in enumerate(adf.columns,1):aw.cell(1,j).value=h
    for i,row in enumerate(adf.itertuples(index=False,name=None),2):
        for j,val in enumerate(row,1):aw.cell(i,j).value=None if pd.isna(val) else val
    aw.freeze_panes="A2";aw.auto_filter.ref=aw.dimensions
    for col in aw.columns:
        letter=col[0].column_letter;length=max(len(str(c.value)) if c.value is not None else 0 for c in col);aw.column_dimensions[letter].width=min(max(length+2,14),48)
    buf=BytesIO();wb.save(buf)
    st.success(f"Updated or generated {written} product records. Existing template rows were reused where possible.")
    if review:st.warning(f"{review} records need product-match review. See the Match_Audit worksheet before SAP migration.")
    st.download_button("Download completed Master Data Template",data=buf.getvalue(),file_name="Master_Data_Template_Product_Matched"+(".xlsm" if keep_vba else ".xlsx"),mime="application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.info("This tool prepares the workbook only; it does not create or update SAP records.")
