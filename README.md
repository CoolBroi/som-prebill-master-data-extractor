# SOM Pre-Bill Auto Extractor

Streamlit app that automatically detects Pre-Bill data worksheets, header rows, a Master Data Template worksheet, Tower / Track columns, and supporting mapping tables.

## Deploy on Streamlit Community Cloud

1. Sign in at https://share.streamlit.io/ with the GitHub account that can access this repository.
2. Click **Create app**.
3. Repository: `CoolBroi/som-prebill-master-data-extractor`
4. Branch: `main`
5. Main file path: `app.py`
6. Click **Deploy**.

## Run locally

```bash
python -m venv .venv
# Windows:
.venv\\Scripts\\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Automatic workflow

1. Upload the Pre-Bill workbook and Master Data Template.
2. The app scans worksheets and detects likely RU/data sheets and header rows.
3. It detects the template worksheet and the Tower / Track columns from their headings.
4. It scans supporting sheets for explicit Tower and Track columns and builds keyword mapping rules from their other columns.
5. It extracts rows and matches Tower / Track where a rule matches.
6. It appends fields to the output template when source and template column headings match, then offers the completed workbook for download.

## Mapping and safety notes

- No worksheet, header row, output column letter, or mapping-column selection is required in the normal workflow.
- Tower / Track is only filled when an explicit supporting mapping table provides a match. If no mapping table exists, or a match is ambiguous, the value is left blank and flagged rather than guessed.
- The app currently uses keyword matching. It must be tested against representative Pre-Bill workbooks and the authoritative business mapping before production migration use.
- Source columns are copied only when their normalized header names match a template header. Different names require an explicit business-approved mapping enhancement.
- This tool prepares Excel files only; it does not connect to SAP or create SAP master data.
- Uploaded workbook contents are processed by the Streamlit app and are not committed to this public GitHub repository. Do not commit confidential customer files.
- Validate the downloaded workbook, including formatting and formulas, before SAP use.
