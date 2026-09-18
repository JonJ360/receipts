# 360 Receipt Tracker

Supabase-backed receipt capture for the 360 Apps hub. Data and private receipt sources remain in Supabase under existing RLS. Version **1.2**.

## Export reports

On **Receipts**, choose month/company/status/search filters, then **Export report** beside Reload. The modal refreshes the ledger and freezes those filters for the report. It shows counts and per-currency totals before download.

- **PDF with receipt pictures** follows the verified historical Smart Receipts layout: A4, blue rules, a six-column striped summary (Date, Name, Price, Currency, Category Name, Reimbursable), then four numbered receipt panels per page. All source PDF pages occupy panels in order; images are fitted whole without cropping. Original bytes are also attached inside the PDF for full-resolution review. No per-receipt cover sheets or giant screenshot pagination. Browser-font summary/captions preserve Unicode but are not selectable; embedded source PDF text stays selectable. The footer accurately identifies Receipt Tracker, not Smart Receipts.
- **Excel verification** downloads a real `.xlsx`: Summary plus Receipts sheets; the first eight receipt columns match the historical CSV exactly (Date, Name, Price, Currency, Category Name, Category Code, Comment, Reimbursable), followed by verification/provenance fields. Native dates and numeric signed amounts, separate currencies, coding, IDs, provenance and source hashes, frozen headings and autofilter. No pictures or expiring signed URLs. Text beginning with `=` stays text, not a formula.
- Excluded receipts are omitted by default. Explicit opt-in includes them as evidence/verification rows with **separate excluded totals**, never in report or reimbursable totals. This also applies when the status filter is Excluded. Positive/negative wash entries retain their signs.
- Export does not change status, write records, email, share or send anything. Downloads contain private information; handle locally according to your normal accounting process.
- Empty selections disable export. Source download/authentication errors, encrypted PDFs and unreadable images block the **entire PDF**, never silently drop evidence. Excel can still verify rows with missing evidence. HEIC/HEIF decoding depends on browser support; if unsupported, convert the source to JPEG/PNG first. There is no automatic replacement of originals.
- Fresh ledger reads paginate in stable ID order rather than silently truncating at 1,000. Reports sort by date/vendor/ID. These are point-in-time browser snapshots, not transactionally locked month closes; reload after concurrent edits. Large image batches use browser memory; narrow the month/company filter on constrained devices.

## Dependencies

Export libraries are vendored and loaded only when needed; they do not receive receipt data. No schema changes or new backend service.

| File | Version | License | Upstream download |
|---|---|---|---|
| `vendor/pdf-lib.min.js` | pdf-lib 1.17.1 | MIT (`pdf-lib.LICENSE`) | https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js |
| `vendor/exceljs.min.js` | ExcelJS 4.4.0 | MIT (`exceljs.LICENSE`) | https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js |

The pre-existing Supabase client CDN dependency is unchanged. All production source fetches use signed private URLs obtained through that authenticated client. No service key is in this repository.

## Tests

```
node --test tests/reports.test.cjs
python -m pip install playwright pymupdf openpyxl
python tests/browser_reports.py
python tests/ui_reports.py
```

Browser tests use installed headless Microsoft Edge. Fixtures are synthetic only, not inserted into Supabase. Artifacts go to the system temp directory `receipts-export-tests`, never the public repository. Tests cover signed wash amounts, filters, currencies, excluded opt-in, stable sorting, 1,001-row pagination, real XLSX round-trip, embedded tall-image and multipage PDF evidence, empty/missing/unreadable sources, expired authentication, failed ledger reads, desktop/mobile layout and browser download.

Production verification artifacts, when authorized, must be kept outside this repository. Never commit source receipts, production snapshots, credentials or signed URLs.
