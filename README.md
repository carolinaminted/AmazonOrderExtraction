# Amazon Order Extraction Scripts

This repository contains two Google Apps Script utilities that help automate Amazon order management from Gmail and Google Drive:

- `ingestAmazonPurchases.js` — parses order confirmation emails and appends structured purchase data to a Google Sheet.
- `exportAmazonPurchasesToPDFs.js` — converts qualifying order emails into PDF files saved to Google Drive.

Both scripts rely on Gmail labels and require access to a Google Sheet (for ingestion) and Google Drive (for PDF export). Deploy them inside the same Apps Script project that has access to Gmail, Drive, Sheets, and the Spreadsheet that will display toast notifications.

## Function Reference

### ingestAmazonPurchases.js

| Function | Purpose | Key Output / Side Effect |
| --- | --- | --- |
| `ingestAmazonPurchases()` | Entry point that scans labeled Gmail threads, filters out non-order messages, parses relevant ones, and appends rows to the configured sheet. | Appends order data rows to the active sheet and shows a toast summary. |
| `parseAmazonEmail_(msg)` | Extracts order metadata (date, number, item title, total) from an Apps Script `GmailMessage`. | Returns an object: `{ orderDate, orderNumber, itemTitle, orderTotal }` or `null` on failure. |
| `loadExistingAmazonMessageIds_(sheet)` | Reads the fifth column of the sheet to gather message IDs already processed. | Returns a `Set` of message IDs for deduplication. |
| `round2_(n)` | Rounds a number to two decimal places. | Returns the rounded number. |

### exportAmazonPurchasesToPDFs.js

| Function | Purpose | Key Output / Side Effect |
| --- | --- | --- |
| `exportAmazonPurchasesToPDFs()` | Entry point that converts filtered Gmail messages into PDFs stored in Drive. | Creates PDF files in the configured Drive folder and logs processed IDs in a sheet. |
| `buildAmazonFileName_(msg)` | Generates a descriptive PDF filename using the message subject and any detected order number. | Returns a filename string ending in `.pdf`. |
| `renderMessageToPDF_(msg)` | Wraps email metadata and HTML body into a printable template and converts it to a PDF blob. | Returns a PDF blob suitable for saving to Drive. |
| `getOrCreateFolderByPath_(path)` | Ensures a nested Drive folder path exists. | Returns a `DriveApp.Folder` instance. |
| `inlineCidImages_(html, msg)` | Replaces inline CID image references with data URIs so they render in the PDF. | Returns modified HTML with inline images embedded. |
| `inlineExternalImages_(html)` | Resolves external image URLs to inline data URIs when feasible. | Returns HTML with remote images embedded when successful. |
| `normalizeGoogleProxyUrl_(u)` | Cleans Google proxy URLs to improve fetch reliability. | Returns the normalized URL string. |
| `loadProcessedIds_(logSheetName)` | Reads a helper sheet that stores processed message IDs. | Returns a `Set` of message IDs and creates the sheet if needed. |
| `saveProcessedIds_(set, logSheetName)` | Writes the processed message ID set back to the helper sheet. | Updates the helper sheet with the latest IDs. |
| `escape_(s)` | Escapes HTML-sensitive characters. | Returns a safe string for HTML output. |

## Setup & Usage

1. **Create/identify your Spreadsheet**
   - Add tabs that match `SHEET_NAME_AMAZON` (default `"Amazon Orders"`) and the optional PDF log sheet (`"Amazon PDFs"`).
   - Ensure the Apps Script project is bound to this spreadsheet so calls to `SpreadsheetApp.getActive()` succeed.

2. **Label your Gmail messages**
   - Create the Gmail label named `Amazon Orders` (or adjust the constants). Apply it to order confirmation emails you want processed.

3. **Configure constants**
   - In each script, review the `LABEL_NAME_*`, `SHEET_NAME_AMAZON`, `DRIVE_FOLDER_PATH_AMAZON`, `LOG_SHEET_AMAZON`, and per-run caps. Modify them if your environment uses different names or limits.

4. **Authorize services**
   - On first execution, grant the script permission to access Gmail, Drive, and Sheets.

5. **Run the desired function**
   - Use Apps Script's *Run* menu to execute either `ingestAmazonPurchases` or `exportAmazonPurchasesToPDFs`.
   - Consider setting time-driven triggers (e.g., hourly) if you want automated processing.

## Troubleshooting

| Symptom | Likely Cause | Resolution |
| --- | --- | --- |
| Execution stops with "Sheet named ... could not be found" | Spreadsheet tab is missing or misspelled. | Create/rename the tab to match `SHEET_NAME_AMAZON`. |
| Execution stops with "Label not found" | Gmail label is missing or named differently. | Create the label or update `LABEL_NAME_AMAZON` / `LABEL_NAME_PDF_AMAZON`. |
| Emails skipped despite being Amazon orders | Filters require `from` to include `auto-confirm@amazon.com` and subject to contain `ordered`. | Verify sender and subject format or adjust the filter logic. |
| PDFs not appearing in Drive | Drive path constant is wrong or permissions denied. | Confirm `DRIVE_FOLDER_PATH_AMAZON` is valid and the script has Drive access. |
| Duplicate rows or PDFs | Helper sheets missing or cleared. | Ensure the log sheet exists and is not manually cleared; rerun to rebuild state. |
| Order total parsed as `null` | Email layout differs or lacks a "Total" line. | Adjust the regex or parsing logic within `parseAmazonEmail_` as needed. |

## Logging & Monitoring

- Both scripts use `console.log` and `console.error` extensively. Open **View → Logs** in Apps Script after a run to review messages.
- `SpreadsheetApp.getActive().toast(...)` provides quick success summaries.
- For PDF export, the helper sheet logs processed IDs to avoid duplication between runs.

## Extending the Scripts

- **Different retailers**: Duplicate the scripts and update filter logic to match other sender domains/subjects.
- **Custom file naming**: Modify `buildAmazonFileName_` to include additional metadata.
- **Additional data fields**: Extend `parseAmazonEmail_` and update `appendRow` calls to include more purchase details.

## Support Checklist

Before contacting support or making changes, confirm the following:

- [ ] Gmail label exists and is applied to target emails.
- [ ] Spreadsheet tabs referenced by the constants exist.
- [ ] Apps Script project has the necessary service scopes.
- [ ] Helper sheets (`Amazon Orders` data tab and `Amazon PDFs`) contain headers in row 1.
- [ ] Script executed recently and logs show expected activity.

