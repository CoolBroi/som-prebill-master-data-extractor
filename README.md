# SOM Pre-Bill Master Data Extractor

Streamlit application for extracting Pre-Bill worksheet data, assigning Tower / Track using configurable mapping evidence, reviewing the result, and exporting a completed Master Data Template.

## Deploy on Streamlit Community Cloud

1. Sign in at https://share.streamlit.io/ using the GitHub account that has access to this repository.
2. Click **Create app**.
3. Choose repository: `CoolBroi/som-prebill-master-data-extractor`.
4. Branch: `main`.
5. Main file path: `app.py`.
6. Click **Deploy**. Streamlit installs packages from `requirements.txt`.

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

## Workflow

1. Upload the Pre-Bill workbook and Master Data Template.
2. Select the template worksheet and Tower / Track output columns (defaults H and I).
3. Select the Pre-Bill worksheets to process and confirm each sheet's header row.
4. Optionally select a supporting mapping sheet and identify its keyword, Tower and Track columns. The app applies the longest keyword match found in each source row.
5. Enter default Tower / Track values for each RU worksheet. Rows that do not match a mapping rule use these defaults and are flagged for review.
6. Preview extracted rows, then either match to existing template rows using a common key or append extracted rows.
7. Download the completed workbook and validate it before using it for SAP migration.

## Notes and limitations

- This app prepares Excel files only. It does not connect to SAP or create SAP master data.
- The app does not guess Tower / Track when it has no mapping evidence. Unmatched rows should be reviewed.
- The current mapping engine uses keyword matching. Real Pre-Bill RU layouts and the team's authoritative mapping/master data should be tested and refined before production use.
- Match mode updates Tower / Track on existing template rows. Append mode copies fields when source and target headers match and appends rows.
- The app uses pandas and openpyxl. It aims to retain workbook styles in append mode, but complex workbook features, macros, formulas, pivot tables, and external links should be validated in the exported file.
- Uploaded workbook contents are processed in the Streamlit app session and are not committed to this public GitHub repository. Do not commit customer data or confidential files to GitHub.
- `index.html` is an earlier browser prototype; deploy `app.py` on Streamlit Community Cloud.
