# Receivables Dashboard (Finance 360° — standalone module)

This is a **separate, standalone Google Apps Script web app** for the Receivables
Dashboard. It is not yet merged into the main Margin Entry System — it will be
combined into the Finance 360° Command Center later, as requested.

## Files (copy all 4 into a new Apps Script project)

| File | Purpose |
|---|---|
| `ReceivablesCode.gs` | Server-side: `doGet`, auto-detects the Receivables / Receipt / Balance (ledger) sheets by matching column headers (no hardcoded tab names or column letters), exposes `getReceivablesBootstrap()` and `getLedgerForCustomer()`. |
| `ReceivablesDashboard.html` | Page shell: fonts, CSS (matches the approved Finance 360° look — dark navy sidebar, white cards, blue/green/orange/red status colors), includes the two files below. |
| `ReceivablesEngine.html` | Pure calculation engine (header detection, currency/date parsing, dedup, KPI/ageing/customer/sales-person/forecast/insight calculations). This logic was unit-tested in Node before being embedded here — do not hand-edit without re-testing, since date/currency parsing edge cases are easy to break silently. |
| `ReceivablesUI.html` | State management, rendering, charts (Chart.js), filters, customer drill-down modal, event handling. |

## How to deploy

1. Create a **new** Google Apps Script project (Extensions → Apps Script, or script.google.com) bound to your Google Sheet, or create a fresh standalone project and set the Sheet ID if needed.
2. Create 4 files matching the names above (`.gs` file type for `ReceivablesCode`, `.html` for the other three) and paste in the content from GitHub (use the **Raw** view).
3. Required sheet: a tab with header columns (any tab name works, headers are auto-detected):
   `TIMESTAMP, Bill_Date, Bill_Ref_No, Party_Name, Party_Group, Sub_Group, Main_Group, Sales Person, Phone No., Payment Terms, Pending Amount, Due_Date, Overdue_Days`
4. Optional sheets (auto-detected the same way):
   - **Receipt** sheet — columns like `Receipt_Date, Receipt_No, Party_Name, Amount, Mode, Reference, Sales Person` — powers Collected Today, Collection Trend, Recent Collections, Collection Efficiency.
   - **Balance / Ledger** sheet — columns like `Voucher_Date, Voucher_Particular, Voucher_Type, Voucher_No, Voucher_Debit, Voucher_Credit, Opening_Balance, Closing_Balance` — powers the Ledger tab in the customer drill-down.
5. Deploy → New deployment → type "Web app" → Execute as "Me" → Who has access "Anyone" → Deploy.

## Calculation rules implemented (per spec)

- Base data rule: `Party_Group = Debtors`, `Pending Amount > 0`, `Party_Name` not blank, `Bill_Date <= As On Date`.
- `Null` / blank / `-` / `NA` / `N/A` treated as empty everywhere.
- Currency strings like `₹1,06,200` parsed to `106200`.
- Dedup key: normalized Party Name + Bill_Ref_No + Bill_Date, keeping the row with the latest `TIMESTAMP`.
- Overdue Days recalculated from `As On Date - Due_Date` (not blindly trusted from the sheet's `Overdue_Days` column).
- Ageing buckets: Not Due, 0–30, 31–60, 61–90, 90+ — bucket totals always sum to Total Receivable.
- Indian Financial Year (1 Apr – 31 Mar) used for QTD/YTD period calculations.
- All KPI cards, charts, and tables operate on the same filtered/deduped dataset.

## Testing performed before delivery

A Node-based test harness (not shipped with the app) was used to verify:
- Calculation engine unit tests (currency/date parsing, dedup, ageing, KPIs, insights) against sample data matching the provided screenshot.
- A DOM-harness functional test that runs the actual `RD_loadAll() → RD_render()` pipeline with mocked `google.script.run`, verifying no runtime errors, correct totals, and no `undefined`/`NaN` leaking into rendered HTML — across all views, the customer drill-down modal (all 6 tabs), and loading/empty/error states.

Two real bugs were found and fixed during this testing:
1. `TIMESTAMP` values in `DD/MM/YYYY` format were being misparsed by `Date.parse`, which broke the "keep latest row" dedup logic.
2. Literal `"Null"` strings in text fields (Sales Person, Phone, etc.) were not being cleaned to empty, which broke "Unassigned" grouping.
