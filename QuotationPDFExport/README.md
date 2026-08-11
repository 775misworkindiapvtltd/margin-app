# Quotation PDF Export Script

Apps Script that generates a paginated PDF from the "Copy of Sales Quotation Format" sheet.

## Setup (do this before running)

1. Open the Google Sheet → **Extensions > Apps Script**.
2. Paste the full contents of `PDFExport.gs` into the editor (a new script file, or replace an existing one).
3. At the top of the file, edit the config block:
   - `PDF_FOLDER_ID` — **you must set this**. Open your Drive folder in the browser and copy the ID from the URL:
     `https://drive.google.com/drive/folders/<PASTE_THIS_PART>`
   - `SHEET_NAME` — defaults to `'PDF FINAL'` (the tab shown in your screenshots). Change if your tab name differs.
   - `HEADER_LAST_ROW` (25), `ITEM_FIRST_ROW` (26), `ITEM_LAST_ROW` (67) — already set to match your layout.
   - `TNC_HEADING_TEXT` — the exact heading text used to auto-locate the "Technical Terms & Conditions" block (`'TECHNICAL TERMS & CONDITIONS'`). If it can't be found, the script falls back to `FOOTER_FIRST_ROW_FALLBACK` (77).
   - `QUOTATION_REF_CELL` — cell used to name the PDF file (defaults to `N16`, the "Quotaion Ref No." value cell in your screenshot). Adjust if it's a different cell in your actual sheet.
4. Save the script, then run `generateQuotationPDF` once from the Apps Script editor to grant permissions (Sheets + Drive).
5. After that, run it via the **Quotation Tools > Generate Quotation PDF** menu that appears in the sheet (added automatically on open), or call `generateQuotationPDF()` from a button/menu of your choice.

## What it does

1. **Rows 26–67 (item rows):** blank rows in this block are hidden. If data exists in only the first 13 rows, rows 26–38 stay visible and 39–67 are hidden. A minimum of 5 rows always stays visible even if fewer (or none) of the rows have data.
2. **Rows 1–25 (header/customer/quotation info) — conditional repeat:** the header only repeats on a later page if the item table itself overflows past page 1. If all the items fit on page 1 and it's only the Terms & Conditions block that gets pushed to page 2, the header will **not** reappear on that page. If the item rows themselves spill onto page 2 (or further), the header repeats on every page that has item rows, as before.
3. **"TECHNICAL TERMS & CONDITIONS" heading through the end of the sheet** (Commercial Terms, Declaration, "Thanks for your business" banner): this whole block is treated as one unit that can never be cut mid-way. The script auto-locates the row containing this heading text (so it keeps working even if rows shift after future edits) and measures the block's height. If it wouldn't fully fit in the remaining space on the current page, it's pushed entirely onto a fresh page by adjusting the spacer row height for the export only — the sheet's own layout is restored immediately after the PDF is generated.
4. The PDF is saved to your configured Drive folder, and the **PDF URL is printed to the console/log** (`console.log` and `Logger.log`) after each run. Open **Execution log** (or **View > Logs**) in the Apps Script editor after running to see it.

## Calibration note (for points 2 & 3)

Both the header-repeat decision and the terms-block-no-split logic estimate how tall one printed page is using `PAGE_WIDTH_BASE_PX` / `PAGE_HEIGHT_BASE_PX` constants (defaults tuned for A4 portrait). Real Google Sheets page breaks can vary slightly by content/margins. After your first PDF:

- If pagination still looks off (terms block cut, or header repeating/not repeating incorrectly), run the `debugPageBreakEstimate` function (from the Apps Script editor, select it in the function dropdown and click Run). It logs the detected terms-block row, estimated page height, header height, items+charges content height, terms block height, and the header-repeat decision. Compare against where the actual page break falls in your PDF, then adjust `PAGE_HEIGHT_BASE_PX` up or down slightly and re-run.

## Testing

`test_logic.js` contains the pure decision logic (row-visibility counting, header-repeat page-count math, and terms-block-fit math) copied out of the `.gs` file, with no Apps Script API calls, so it can run standalone in plain Node.js. This was used to verify correctness before delivery — it is not required inside Apps Script.

```
node test_logic.js
```

Covers: exact 13-filled-row example from the request, fewer-than-minimum cases, zero-filled case, all-filled case, gaps between filled rows, whitespace-only cells, the three terms-block-fit scenarios (fits already / needs push to next page / exact fit), and the header-repeat page-count scenarios (fits on one page / overflows to two / exact fit / the specific "items fit, only terms block pushed" case / zero-height guard).
