from __future__ import annotations

from io import BytesIO
import re
from copy import copy
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SOM Pre-Bill Master Data Extractor", page_icon="📊", layout="wide")

st.markdown("""
<style>
.block-container{padding-top:2rem;padding-bottom:3rem;max-width:1400px}
.hero{padding:1.5rem 1.7rem;border-radius:16px;background:linear-gradient(115deg,#14213d,#244b78);color:white;margin-bottom:1.2rem}
.hero h1{color:white;margin:0 0 .35rem 0;font-size:2rem}.hero p{color:#dbeafe;margin:0}
.step{background:#f1f5f9;border-radius:10px;padding:12px 14px;height:100%;color:#1e293b}
.small{color:#64748b;font-size:.88rem}
</style>
<div class="hero"><h1>Pre-Bill → Master Data Template</h1>
<p>Extract RU worksheet data, determine Tower / Track using mapping evidence, review the assignments, and export the completed Excel template.</p></div>
""", unsafe_allow_html=True)

st.markdown("**Workflow**")
c1,c2,c3=st.columns(3)
for col,num,title,desc in [(c1,"01","Upload","Pre-Bill workbook and Master Data Template"),(c2,"02","Map & review","Use a mapping sheet or assign Tower / Track per RU sheet"),(c3,"03","Export","Update matching template rows or append new rows")]:
    col.markdown(f'<div class="step"><b>{num} · {title}</b><br><span class="small">{desc}</span></div>',unsafe_allow_html=True)

def norm(x: Any) -> str:
    return re.sub(r"[^a-z0-9]+","",str(x or "").strip().lower())

def read_frames(file_bytes: bytes) -> dict[str,pd.DataFrame]:
    return pd.read_excel(BytesIO(file_bytes),sheet_name=None,header=None,dtype=object)

def guess_header(df: pd.DataFrame) -> int:
    best_i,best_score=0,-1
    for i in range(min(12,len(df))):
        vals=[str(v).strip() for v in df.iloc[i].tolist() if pd.notna(v) and str(v).strip()]
        score=len(vals)+sum(2 for v in vals if any(k in norm(v) for k in ["tower","track","project","service","description","resource","item","product","wbs","ru"]))
        if score>best_score: best_i,best_score=i,score
    return best_i

def as_table(df: pd.DataFrame, header_row: int) -> tuple[list[str],pd.DataFrame]:
    header=[str(v).strip() if pd.notna(v) and str(v).strip() else f"Column {i+1}" for i,v in enumerate(df.iloc[header_row].tolist())]
    data=df.iloc[header_row+1:].copy()
    data.columns=header
    data=data.dropna(how="all").reset_index(drop=True)
    return header,data

def find_col(cols, terms):
    for i,c in enumerate(cols):
        if norm(c) in terms: return i
    return 0

def mapping_from_rules(row_values, rules):
    text=" ".join(str(v) for v in row_values if pd.notna(v)).lower()
    matches=[]
    for keyword,tower,track in rules:
        kw=str(keyword).strip().lower()
        if kw and kw in text:
            matches.append((len(kw),tower,track,kw))
    if not matches: return None
    matches.sort(reverse=True)
    return matches[0][1],matches[0][2],matches[0][3]

def copy_row_style(ws, source_row, target_row, max_col):
    if source_row < 1: return
    for col in range(1,max_col+1):
        src=ws.cell(source_row,col); dst=ws.cell(target_row,col)
        if src.has_style: dst._style=copy(src._style)
        if src.number_format: dst.number_format=src.number_format
        if src.alignment: dst.alignment=copy(src.alignment)
        if src.protection: dst.protection=copy(src.protection)

st.header("1. Upload files")
left,right=st.columns(2)
with left:
    prebill=st.file_uploader("Pre-Bill workbook",type=["xlsx","xlsm"],key="prebill")
with right:
    template=st.file_uploader("Master Data Template",type=["xlsx","xlsm"],key="template")
if not prebill or not template:
    st.info("Upload both workbooks to continue. Files are processed in this Streamlit session and are not saved to this GitHub repository.")
    st.stop()

try:
    pre_bytes=prebill.getvalue(); template_bytes=template.getvalue()
    source_sheets=read_frames(pre_bytes)
    keep_vba=template.name.lower().endswith(".xlsm")
    target_wb=load_workbook(BytesIO(template_bytes),keep_vba=keep_vba)
except Exception as exc:
    st.error(f"Could not read one of the workbooks: {exc}")
    st.stop()

st.success(f"Loaded {len(source_sheets)} Pre-Bill worksheets and {len(target_wb.sheetnames)} template worksheets.")
st.header("2. Select the target template")
t1,t2,t3=st.columns(3)
with t1:
    target_sheet=st.selectbox("Template worksheet",target_wb.sheetnames)
target_ws=target_wb[target_sheet]
with t2:
    header_row=st.number_input("Template header row",min_value=1,max_value=100,value=1,step=1)
with t3:
    start_col_default="H"
    tower_col_letter=st.text_input("Tower column (Excel letter)",value=start_col_default).strip().upper()
    track_col_letter=st.text_input("Track column (Excel letter)",value="I").strip().upper()

def col_num(letter):
    from openpyxl.utils.cell import column_index_from_string
    try: return column_index_from_string(letter)
    except Exception: return None

tower_col=col_num(tower_col_letter); track_col=col_num(track_col_letter)
if not tower_col or not track_col:
    st.error("Enter valid Excel column letters, for example H and I.")
    st.stop()

target_headers={str(target_ws.cell(int(header_row),c).value or f"Column {c}").strip():c for c in range(1,max(target_ws.max_column,tower_col,track_col)+1)}
st.caption("Detected template headers: "+", ".join(list(target_headers.keys())[:30]))

st.header("3. Configure RU sheets and Tower / Track rules")
all_sheet_names=list(source_sheets)
default_selected=[s for s in all_sheet_names if s.strip().lower() not in ["summary","cover","contents","index"]]
selected_sheets=st.multiselect("Pre-Bill worksheets to process",all_sheet_names,default=default_selected or all_sheet_names)
if not selected_sheets:
    st.warning("Select at least one worksheet.")
    st.stop()

with st.expander("Optional: use a mapping / master sheet from the Pre-Bill workbook",expanded=False):
    st.write("If a supporting sheet contains keywords or service names with their Tower and Track, configure it here. The app uses the longest matching keyword found in each source row.")
    map_sheet=st.selectbox("Mapping worksheet",["(Do not use a mapping sheet)"]+all_sheet_names)
    rules=[]
    if map_sheet!="(Do not use a mapping sheet)":
        map_raw=source_sheets[map_sheet]
        map_hdr_guess=guess_header(map_raw)
        map_hdr=st.number_input("Mapping sheet header row",min_value=1,max_value=100,value=map_hdr_guess+1,key="map_header")-1
        map_cols,map_df=as_table(map_raw,int(map_hdr))
        st.dataframe(map_df.head(5),use_container_width=True)
        kw_col=st.selectbox("Keyword / service / RU description column",map_cols,key="kw")
        map_tower=st.selectbox("Mapping Tower column",map_cols,index=find_col(map_cols,{"tower"}),key="mt")
        map_track=st.selectbox("Mapping Track column",map_cols,index=find_col(map_cols,{"track"}),key="mr")
        for _,r in map_df.iterrows():
            kw=r.get(kw_col); tw=r.get(map_tower); tr=r.get(map_track)
            if pd.notna(kw) and pd.notna(tw) and pd.notna(tr):
                rules.append((str(kw),str(tw),str(tr)))

st.caption("For each RU sheet, confirm its header row and default Tower / Track. When mapping rules are available, matched rows use the rule; unmatched rows use the sheet defaults and are flagged for review.")
sheet_configs={}
with st.expander("Review source worksheet structure",expanded=True):
    for s in selected_sheets:
        raw=source_sheets[s]
        hdr_guess=guess_header(raw)
        with st.container(border=True):
            st.subheader(s)
            default_hdr=st.number_input(f"Header row · {s}",min_value=1,max_value=100,value=hdr_guess+1,key=f"hdr_{s}")-1
            cols,data=as_table(raw,int(default_hdr))
            st.dataframe(data.head(4),use_container_width=True)
            a,b=st.columns(2)
            with a:
                tower=st.text_input(f"Default Tower · {s}",value="",key=f"tower_{s}")
            with b:
                track=st.text_input(f"Default Track · {s}",value="",key=f"track_{s}")
            sheet_configs[s]={"header":int(default_hdr),"tower":tower.strip(),"track":track.strip(),"columns":cols,"data":data}

st.header("4. Choose how to populate the template")
mode=st.radio("Output method",["Match existing template rows using a common key","Append extracted source rows to the template"],horizontal=False)
match_config=None
if mode.startswith("Match"):
    m1,m2=st.columns(2)
    with m1:
        source_key_sheet=st.selectbox("Source sheet used for key matching",selected_sheets)
        source_cols=sheet_configs[source_key_sheet]["columns"]
        source_key=st.selectbox("Pre-Bill key column",source_cols,key="source_key")
    with m2:
        target_header_names=list(target_headers.keys())
        target_key_name=st.selectbox("Template key column",target_header_names,key="target_key")
    match_config=(source_key_sheet,source_key,target_headers[target_key_name])
else:
    st.warning("Append mode adds source data rows below the existing template data. Use match mode when the template already contains the target records.")

if rules:
    st.success(f"{len(rules)} mapping rules loaded from '{map_sheet}'.")
else:
    st.info("No mapping rules configured. Default Tower / Track values must be entered for each selected worksheet.")

if st.button("Build preview",type="primary",use_container_width=True):
    extracted=[]; unresolved=0
    for s,cfg in sheet_configs.items():
        data=cfg["data"]
        for _,row in data.iterrows():
            values=row.tolist()
            if not any(pd.notna(v) and str(v).strip() for v in values): continue
            found=mapping_from_rules(values,rules) if rules else None
            if found:
                tower,track,reason=found; status=f"Matched keyword: {reason}"
            else:
                tower,track=cfg["tower"],cfg["track"];status="Default / needs review"
                unresolved+=1
            rec={col:row.get(col) for col in cfg["columns"]}
            rec["Tower"]=tower;rec["Track"]=track;rec["Source Worksheet"]=s;rec["Mapping Status"]=status
            extracted.append(rec)
    if not extracted:
        st.error("No data rows were extracted. Check the selected sheets and header rows.")
    else:
        st.session_state["extracted_rows"]=extracted
        st.session_state["mapping_unresolved"]=unresolved
        st.success(f"Preview ready: {len(extracted)} data rows. {unresolved} rows use default values and need review.")
        preview=pd.DataFrame(extracted)
        st.dataframe(preview,use_container_width=True,height=360)
        st.download_button("Download extracted-data review CSV",preview.to_csv(index=False).encode("utf-8-sig"),"prebill_extraction_review.csv","text/csv")

if "extracted_rows" in st.session_state:
    st.header("5. Generate completed Master Data Template")
    output_rows=st.session_state["extracted_rows"]
    st.caption(f"{len(output_rows)} rows ready · {st.session_state.get('mapping_unresolved',0)} rows flagged for review.")
    if st.button("Generate Excel output",type="primary",use_container_width=True):
        ws=target_wb[target_sheet]
        updated=0; appended=0; skipped=0
        if mode.startswith("Match"):
            source_key_sheet,source_key,target_key_col=match_config
            source_data=sheet_configs[source_key_sheet]["data"]
            lookup={}
            for rec in output_rows:
                if rec["Source Worksheet"]!=source_key_sheet: continue
                key=str(rec.get(source_key,"")).strip()
                if key and key not in lookup: lookup[key]=rec
            target_lookup={}
            for r in range(int(header_row)+1,ws.max_row+1):
                key=str(ws.cell(r,target_key_col).value or "").strip()
                if key: target_lookup[key]=r
            for key,rec in lookup.items():
                target_r=target_lookup.get(key)
                if not target_r: skipped+=1;continue
                ws.cell(target_r,tower_col).value=rec.get("Tower")
                ws.cell(target_r,track_col).value=rec.get("Track")
                updated+=1
        else:
            start_row=ws.max_row+1
            for rec in output_rows:
                out_row=start_row+appended
                copy_row_style(ws,max(int(header_row)+1,ws.max_row),out_row,max(ws.max_column,tower_col,track_col))
                # Copy source fields into template columns where normalized headers match.
                for col_idx in range(1,ws.max_column+1):
                    h=str(ws.cell(int(header_row),col_idx).value or "").strip()
                    if h in rec and h not in ["Tower","Track"]:
                        ws.cell(out_row,col_idx).value=rec[h]
                ws.cell(out_row,tower_col).value=rec.get("Tower")
                ws.cell(out_row,track_col).value=rec.get("Track")
                appended+=1
        out=BytesIO()
        target_wb.save(out);out.seek(0)
        ext=".xlsm" if keep_vba else ".xlsx"
        st.download_button("Download completed Master Data Template",data=out.getvalue(),file_name="Master_Data_Template_Completed"+ext,mime="application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if mode.startswith("Match"):
            st.success(f"Output generated. Updated {updated} rows; {skipped} extracted keys were not found in the template.")
        else:
            st.success(f"Output generated. Appended {appended} rows.")
        if st.session_state.get("mapping_unresolved",0):
            st.warning("Some rows use default Tower / Track values. Review the extraction CSV and output before using the file for SAP migration.")
        st.info("The tool only prepares an Excel file. It does not connect to SAP or create master data.")
