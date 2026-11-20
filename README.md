# SOM Product Master Data Extractor

Streamlit application that matches product names from the **ARC RRC** worksheet to **VBRP Item Description** values and prepares the **Master Data Template**.

## Inputs

1. **ARC RRC workbook** — reads the numbered rows under `Resource Unit (RU)` and `Baseline Volume`.
2. **VBRP dump** — identifies `Item Description`, `Sold-to`, `StatKF`, `CoCd`, `Project definition`, and `Material`.
3. **Master Data Template** — detects the template worksheet and its existing `SOW Product Name` values.

## Automatic workflow

1. Upload the three Excel workbooks. The app detects the relevant worksheets and header rows automatically.
2. Read each numbered ARC RRC RU as a product name.
3. Use an existing SOW Product Name in the template as an approved alias only when its product terms clearly correspond to the RU; otherwise match the RU directly.
4. Match the product label to VBRP `Item Description`. Exact descriptions are preferred. Ambiguous matches and products with no match are identified.
5. Populate the template with:
   - SOW Product Name ← VBRP Item Description
   - Customer ID ← VBRP Sold-to
   - Device ID ← VBRP StatKF
   - Company Code ← VBRP CoCd
   - Project Code ← VBRP Project definition
   - Output 2 Name / Value ← QTY / ARC RRC Baseline Volume
6. Reuse existing product rows where possible; append additional VBRP product variants.
7. Add a `Match_Audit` worksheet with source RU, matched description, score/status, and VBRP identifiers.

Tower and Track are not used by this workflow and their values are cleared in the generated product rows. Other template configuration fields are retained from the corresponding existing template row where available.

## Matching and validation

- Exact VBRP descriptions are preferred when an existing template product label matches exactly.
- Product-token matching is used for product variants such as server, RHEL, and backup descriptions.
- Multiple candidate descriptions, shared matches across RU names, and no-match RU names are flagged in `Match_Audit`.
- A similarity/token score is a matching aid, not proof of a correct business decision. Review flagged rows before SAP migration.
- The app does not invent a VBRP match when no meaningful product terms are found.

## Deploy on Streamlit Community Cloud

1. Sign in at https://share.streamlit.io/ using the GitHub account with access to this repository.
2. Select **Create app**.
3. Repository: `CoolBroi/som-prebill-master-data-extractor`
4. Branch: `main`
5. Main file path: `app.py`
6. Click **Deploy**. Streamlit will install packages from `requirements.txt`.

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

The app prepares an Excel workbook only. It does not connect to SAP or create SAP master data. Uploaded customer workbooks are processed by the app and should never be committed to this public repository.
