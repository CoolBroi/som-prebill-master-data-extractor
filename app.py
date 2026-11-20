from __future__ import annotations

from io import BytesIO
from copy import copy
import re
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="SOM Pre-Bill Auto Extractor", page_icon="📊", layout="wide")

st.markdown("""
<style>
.block-container{max-width:1320px;padding-top:2rem}
.hero{padding:1.6rem 1.8rem;border-radius:16px;background:linear-gradient(115deg,#14213d,#244b78);color:white;margin-bottom:1.2rem}
.hero h1{color:white;margin:0 0 .35rem;font-size:2rem}.hero p{color:#dbeafe;margin:0}
.metric{padding:1rem;border:1px solid rgba(128,128,128,.25);border-radius:12px}
</style>
<div class="hero"><h1>Pre-Bill Auto Extractor</h1>
<p>Upload your Pre-Bill and Master Data Template. The app detects worksheets, headers, mapping tables and output columns automatically.</p></div>
""", unsafe_allow_html=True)

def norm(v: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(v or "").strip().lower())

def clean(v: Any) -> str:
    if pd.isna(v): return ""
    return str(v).strip()

def unique_headers(values):
    seen, result = {}, []
    for i, value in enumerate(values):
        base = clean(value) or f"Column {i+1}"
        seen[base] = seen.get(base, 0) + 1
        label = base if seen[base] == 1 else f"{base} ({seen[base]})"
        while label in result:
            seen[base] += 1
            label = f"{base} ({seen[base]})"
        result.append(label)
    return result

def detect_header(raw: pd.DataFrame, template=False):
    best = None
    terms = ["tower","track","customerid","deviceid","sowproductname","projectcode","ratingmodel","description","service","product","resource","item","wbs","key"]
    for i in range(min(20, len(raw))):
        vals = [norm(x) for x in raw.iloc[i].tolist() if clean(x)]
        if not vals: continue
        score = len(vals) + sum(3 for x in vals if any(t in x for t in terms))
        if template and any(x in {"tower","track"} for x in vals): score += 20
        if best is None or score > best[0]: best = (score, i)
    return best[1] if best else 0

def make_table(raw, header_idx):
    headers = unique_headers(raw.iloc[header_idx].tolist())
    data = raw.iloc[header_idx+1:].copy()
    data.columns = headers
    data = data.dropna(how="all").reset_index(drop=True)
    return headers, data

def find_header_col(headers, aliases):
    aliases = {norm(x) for x in aliases}
    for h in headers:
        if norm(h) in aliases: return h
    for h in headers:
        if any(a in norm(h) for a in aliases): return h
    return None

def workbook_sheets(upload):
    return pd.read_excel(BytesIO(upload.getvalue()), sheet_name=None, header=None, dtype=object)

def analyse_template(raw_sheets):
    candidates = []
    for name, raw in raw_sheets.items():
        hr = detect_header(raw, template=True)
        headers, df = make_table(raw, hr)
        tower = find_header_col(headers, ["Tower"])
        track = find_header_col(headers, ["Track"])
        score = (20 if tower else 0) + (20 if track else 0)
        score += sum(3 for x in ["Customer ID","Device ID","SOW Product Name","Rating Model","#Key"] if find_header_col(headers,[x]))
        candidates.append((score, name, hr, headers, df, tower, track))
    candidates.sort(key=lambda x:x[0], reverse=True)
    return candidates[0]

def extract_mapping_rules(sheet_tables):
    rules, mapping_sheets = [], []
    for sheet, (headers, df, _) in sheet_tables.items():
        tower_col = find_header_col(headers, ["Tower"])
        track_col = find_header_col(headers, ["Track"])
        if not tower_col or not track_col: continue
        mapping_sheets.append(sheet)
        keyword_cols = [h for h in headers if h not in {tower_col, track_col} and any(k in norm(h) for k in ["service","description","product","item","resource","ru","keyword","name","towertrack"])]
        if not keyword_cols:
            keyword_cols = [h for h in headers if h not in {tower_col, track_col}]
        for _, row in df.iterrows():
            tw, tr = clean(row.get(tower_col)), clean(row.get(track_col))
            if not tw or not tr: continue
            for col in keyword_cols:
                key = clean(row.get(col))
                if len(key) >= 4:
                    rules.append((key.lower(), tw, tr, sheet))
    # Keep all rules; collisions are detected when matching a row.
    return rules, mapping_sheets

def match_tower_track(row, rules):
    haystack = " | ".join(clean(v).lower() for v in row.tolist() if clean(v))
    found = []
    for keyword, tower, track, source in rules:
        if keyword in haystack:
            found.append((len(keyword), tower, track, keyword, source))
    if not found: return "", "", "No matching Tower/Track rule found"
    found.sort(key=lambda x:x[0], reverse=True)
    best_len = found[0][0]
    best = {(x[1], x[2]) for x in found if x[0] == best_len}
    if len(best) > 1: return "", "", "Conflicting mapping rules — review required"
    tower, track = next(iter(best))
    return tower, track, f"Matched '{found[0][3]}' from {found[0][4]}"

def copy_style(ws, source_row, target_row, max_col):
    if source_row < 1: return
    for c in range(1, max_col+1):
        src, dst = ws.cell(source_row,c), ws.cell(target_row,c)
        if src.has_style: dst._style = copy(src._style)
        if src.number_format: dst.number_format = src.number_format

st.markdown("### Upload files")
a,b = st.columns(2)
with a: prebill = st.file_uploader("Pre-Bill workbook", type=["xlsx","xlsm"], key="pb")
with b: template = st.file_uploader("Master Data Template", type=["xlsx","xlsm"], key="mt")
if not prebill or not template:
    st.info("Upload both Excel files. Worksheet and column configuration is automatic.")
    st.stop()

try:
    src_raw = workbook_sheets(prebill)
    tmpl_bytes = template.getvalue()
    keep_vba = template.name.lower().endswith(".xlsm")
    out_wb = load_workbook(BytesIO(tmpl_bytes), keep_vba=keep_vba)
    tmpl_raw = workbook_sheets(template)
except Exception as e:
    st.error(f"Unable to read the uploaded workbooks: {e}")
    st.stop()

# Detect all source worksheet headers and classify sheets.
source_tables = {}
for name, raw in src_raw.items():
    hr = detect_header(raw)
    headers, df = make_table(raw, hr)
    source_tables[name] = (headers, df, hr)

rules, mapping_sheets = extract_mapping_rules(source_tables)
template_pick = analyse_template(tmpl_raw)
_, target_name, target_header_idx, target_headers, target_df, target_tower_header, target_track_header = template_pick
target_ws = out_wb[target_name]

# A source sheet is a data/RU sheet if it has actual records and is not an obvious cover/summary/mapping sheet.
ru_sheets = []
for name, (headers, df, hr) in source_tables.items():
    n = norm(name)
    if len(df) == 0: continue
    if any(x in n for x in ["summary","cover","contents","index","mapping","masterdata","towertrack"]): continue
    if name in mapping_sheets: continue
    # Ignore sheets that are only a small lookup table; keep actual CPO/RU tabs.
    ru_sheets.append(name)

st.markdown("### Automatic detection results")
m1,m2,m3,m4 = st.columns(4)
for col, value, label in [
    (m1, len(ru_sheets), "RU / data worksheets"),
    (m2, len(mapping_sheets), "Mapping worksheets"),
    (m3, target_name, "Detected template sheet"),
    (m4, f"Row {target_header_idx+1}", "Template header row")
]:
    col.markdown(f'<div class="metric"><b>{value}</b><br><small>{label}</small></div>', unsafe_allow_html=True)

st.caption("Detected template fields: " + ", ".join(target_headers[:35]))
if not ru_sheets:
    st.error("No data worksheets were detected in the Pre-Bill workbook. Check that the workbook contains populated RU sheets.")
    st.stop()
if not target_tower_header or not target_track_header:
    st.error("The Master Data Template was detected, but its Tower and/or Track columns could not be identified by header. Please check that the template contains columns named Tower and Track.")
    st.stop()

if mapping_sheets:
    st.success("Detected mapping source(s): " + ", ".join(mapping_sheets))
else:
    st.warning("No supporting worksheet with explicit Tower and Track columns was detected. The app will extract the data, but it will leave Tower/Track blank rather than invent business mappings.")

# Extract rows and automatically map same-named columns to template.
template_norm_to_col = {}
for c in range(1, target_ws.max_column+1):
    h = clean(target_ws.cell(target_header_idx+1,c).value)
    if h: template_norm_to_col[norm(h)] = c

output_records, source_counts = [], {}
for sheet in ru_sheets:
    headers, df, hr = source_tables[sheet]
    source_counts[sheet] = len(df)
    for _, row in df.iterrows():
        if not any(clean(v) for v in row.tolist()): continue
        tower, track, status = match_tower_track(row, rules)
        record = {h: row.get(h) for h in headers}
        record["Tower"] = tower
        record["Track"] = track
        record["_Source Sheet"] = sheet
        record["_Mapping Status"] = status
        output_records.append(record)

if not output_records:
    st.error("No data rows found in the detected RU worksheets.")
    st.stop()

preview = pd.DataFrame(output_records)
# Show source sheet detection and mapping quality.
st.markdown("### Extraction preview")
x1,x2,x3 = st.columns(3)
x1.metric("Extracted rows", len(preview))
x2.metric("Tower/Track matched", int(((preview["Tower"] != "") & (preview["Track"] != "")).sum()))
x3.metric("Needs mapping review", int(((preview["Tower"] == "") | (preview["Track"] == "")).sum()))
with st.expander("Detected worksheets and row counts", expanded=False):
    st.dataframe(pd.DataFrame([{"Worksheet":k,"Detected rows":v,"Header row":source_tables[k][2]+1} for k,v in source_counts.items()]), use_container_width=True)
with st.expander("Preview extracted records", expanded=True):
    st.dataframe(preview.head(200), use_container_width=True, height=420)

st.markdown("### Generate Master Data Template")
st.write("Source columns are copied automatically when their names match template column names. Tower and Track are written to the detected template columns.")
if st.button("Generate completed Master Data Template", type="primary", use_container_width=True):
    ws = out_wb[target_name]
    start_row = max(ws.max_row + 1, target_header_idx + 2)
    written = 0
    for idx, rec in enumerate(output_records):
        row_num = start_row + idx
        copy_style(ws, max(target_header_idx+2, ws.max_row), row_num, max(ws.max_column, len(target_headers)))
        for key, value in rec.items():
            if key.startswith("_"): continue
            target_col = template_norm_to_col.get(norm(key))
            if target_col:
                ws.cell(row_num, target_col).value = value
        ws.cell(row_num, template_norm_to_col[norm(target_tower_header)]).value = rec["Tower"]
        ws.cell(row_num, template_norm_to_col[norm(target_track_header)]).value = rec["Track"]
        written += 1
    result = BytesIO()
    out_wb.save(result)
    st.download_button("Download completed Master Data Template", data=result.getvalue(),
        file_name="Master_Data_Template_Completed" + (".xlsm" if keep_vba else ".xlsx"),
        mime="application/vnd.ms-excel.sheet.macroEnabled.12" if keep_vba else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.success(f"Prepared {written} rows in '{target_name}'.")
    if any(not r["Tower"] or not r["Track"] for r in output_records):
        st.warning("Some Tower/Track values are blank because no authoritative mapping rule matched. Review those rows before SAP migration.")
    st.info("The workbook is prepared only; this app does not create or update SAP records.")
