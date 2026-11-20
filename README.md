# SOM Pre-Bill Master Data Extractor

A browser-based Excel utility for transferring Pre-Bill worksheet data into a Master Data Template and assigning Tower / Track values.

## Run locally
Open `index.html` in a modern browser while connected to the internet (SheetJS is loaded from jsDelivr). Excel files are processed in the browser.

## GitHub Pages deployment
1. Open **Settings → Pages** in this repository.
2. Under **Build and deployment**, choose **Deploy from a branch**.
3. Select branch **main** and folder **/(root)**, then Save.
4. Wait for the Pages deployment to finish. The site URL will be shown in Settings → Pages.

## Current MVP workflow
1. Upload the Pre-Bill workbook and Master Data Template.
2. Select the target template worksheet and Tower / Track columns.
3. Select Pre-Bill worksheets to include.
4. Review and set Tower / Track for each selected worksheet.
5. Preview rows and export the completed workbook.

## Important limitations
- Tower / Track assignments are user-reviewed per worksheet. The tool does not yet infer mappings from an RU page or mapping/master sheets.
- Append mode maps source fields to template columns when their header names match. Review the output carefully.
- Match mode updates Tower / Track for rows whose selected key values match.
- This browser MVP uses SheetJS Community Edition. Rebuilding a worksheet for export may not preserve all advanced Excel formatting, macros, formulas, or workbook features. Keep an untouched copy of the original template and validate the exported file before SAP upload.
- Do not upload confidential customer data to a public repository. The workbook contents are processed in the browser and are not committed to GitHub by this app.
