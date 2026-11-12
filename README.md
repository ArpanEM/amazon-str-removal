# Amazon STR & Removal Control Tower — Emami

A self-managed static web app for GitHub Pages + Supabase. It is designed for multiple employees and does **not** connect to SAP. The Amazon STR report is the source file for PO grouping and SAP document references.

## Included modules

- Dashboard with month and financial-year filters.
- Amazon STR import: only `Transaction Type = FC_REMOVAL` is accepted into the operational STR table. `FC_TRANSFER` and `FC_REMOVAL-Cancel` are excluded.
- Removal Order Detail register.
- Removal Shipment Detail tracker, one row per package/tracking line. A generated internal database ID is distinct from Amazon's `tracking-number`.
- Warehouse receipt allocation: Good, Damaged, Near Expiry, Expired, Short, Not Received.
- Amazon Removal Fees: actual `removal-fee` and `currency` values are captured directly from Amazon Removal Order Detail, by order/SKU/FNSKU. No separate fee report is required.
- PO grouping and SAP document tracker based on the Amazon STR source file. PO Number, Delivery No, Billing Document and MIGO No are read from that same STR file; no separate SAP upload is required.
- Claims, import history, employee roles and audit logs.

## Important operating definitions

1. **Physical received quantity** = Good + Damaged + Near Expiry + Expired.
2. **Final allocation** = physical received + Short + Not Received.
3. **Pending receipt** = shipped quantity − final allocation. This prevents Short and Not Received from being double-counted as physical receipts.
4. Amazon shipment status such as `Delivered` is not warehouse receipt confirmation.
5. Shipment month uses shipment date; receipt month uses receipt date; order month uses request date.
6. Removal fee totals are based on Amazon-reported fee rows and only accepted when the removal order ID exists in the Removal Orders register. Fee currency is retained. Do not aggregate mixed currencies as one INR total without a currency conversion policy.
7. Average fee per unit is meaningful only when the report provides comparable fee and unit quantities. This starter displays a simple total-fees / quantity measure only when shipment quantities exist; validate currency and period before using it for finance.
8. STR values/invoice value are not calculated. Quantity and process reconciliation are the core controls.

## 1. Create Supabase project

1. Create a project at https://supabase.com.
2. Open **SQL Editor** and run all of `supabase_schema.sql` only for a brand-new database. For your existing database, do NOT rerun the original schema or the already-applied migrations. Run the new `excel_auto_mapping_migration.sql` exactly once.
3. Go to **Project Settings → API** and copy Project URL and `anon` public key.
4. Copy `config.example.js` to `config.js` and replace the placeholders.
5. Never put the Supabase `service_role` key in browser code or GitHub.

## 2. Bootstrap the first admin

1. In Supabase **Authentication → Users**, create the first user (your own email).
2. Sign in once through the app so the `profiles` row is created.
3. In SQL Editor, run the following with your auth user UUID:

```sql
update public.profiles
set role = 'admin'
where email = 'YOUR_EMAIL_ADDRESS';
```

Create other employees under **Authentication → Users**. Their profiles are created on first sign-in. Admin can then assign `employee`, `manager`, or `admin`.

## 3. Publish on GitHub Pages

1. Create a private or organization-approved GitHub repository.
2. Upload `index.html`, `styles.css`, `app.js`, `config.js`, and optionally `supabase_schema.sql`/README.
3. Repository → **Settings → Pages** → Deploy from branch → `main` → `/ (root)`.
4. Open the published URL. Use HTTPS.
5. Supabase Authentication → URL Configuration: set Site URL to the GitHub Pages URL and add it to Redirect URLs.

**Security note:** GitHub Pages is static hosting. The Supabase anon key is visible in browser source by design; database security depends on correctly configured Row Level Security. This starter's operational policies allow authenticated employees broad read/insert/update access. Before production, tighten write policies by role and test with separate employee accounts. Never use the service-role key in the frontend.

## 4. Import order

Recommended initial upload order:
1. Product Mapping Master Excel or CSV (upload the original `.xlsx`, `.xls` or `.csv` file).
2. Removal Order Detail Excel or CSV.
3. Removal Shipment Detail Excel or CSV.
4. Amazon STR report Excel or CSV (only FC_REMOVAL rows will be stored). This same file is used for PO grouping and carries the SAP document reference columns.

Removal fees are imported automatically during step 1 from the `removal-fee` and `currency` columns in the Removal Order Detail CSV. No separate fee report is available or needed.

### Excel / CSV upload and automatic mapping

The import screen accepts `.xlsx`, `.xls` and `.csv` directly. Excel files are read in the browser using SheetJS; the first visible worksheet is imported. Do not convert Excel files manually. Upload the Product Mapping Master first. STR rows with blank SKU are retained: the app attempts unique mapping by source SKU, FNSKU, then ASIN. Shipment rows use available FNSKU, SKU and ASIN identifiers. If a key resolves to more than one distinct SKU/material combination, the row is marked `conflict` and no mapping is guessed. If no unique mapping exists, the source row remains stored with `unmapped` status and its original row data.

The Recent Imports table now stores rejected-row counts and the first row-level error messages in Supabase, so the reason for a failed row remains visible after refresh.

### Amazon source column mapping

The importer recognizes common Amazon headers such as:
- Order: `order-id` / `order id` / `removal-order-id`
- Shipment tracking: `tracking-number`
- FNSKU: `fnsku`
- Quantity: `shipped-quantity` / `quantity-shipped` / `quantity`
- STR transaction type: `transaction-type`
- STR transaction ID: `transaction-id`

If your export uses different column names, update the `pick()` aliases in `app.js` or share a sample export to map the exact columns.

## 5. Amazon removal fee logic

The `Removal Order Detail` CSV is the source of truth. On import, the app reads `removal-fee` and `currency` from each row and stores nonblank fee lines linked to the same removal order, SKU and FNSKU. The report does not provide a separate fee-posted date, so the order request date is used for month filtering and is labelled Order date in the fee register. Blank fee/currency rows are ignored, not treated as zero. Re-importing the same report filename updates its existing order/SKU/FNSKU fee line rather than adding a duplicate. No fee is estimated from quantity.

## 6. PO grouping and SAP document tracking from Amazon STR

The Amazon STR report is the source file used to prepare PO groups. The project does not connect to SAP and does not create SAP purchase orders automatically. It prepares the grouping view and tracks the SAP document references that are present in the same Amazon STR file.

Exact source columns used from the supplied header:
- `Transaction Type` (only `FC_REMOVAL` accepted)
- `Transaction Id`, `Order Id`, `Ship From Fc`, `Invoice Number`, `Invoice Date`, `Asin`, `Sku`, `Quantity`, `ACTUAL QTY`, `Invoice Quantity`, `Hsn Code`, `Irn Filing Status`
- `Po Number`, `Delivery No`, `Billing Document`, `MIGO No`, `Status`

PO grouping rules confirmed by the business user:
1. `GOVT_ACCEPTED`: keep each Transaction ID as an individual PO group. All SKU lines belonging to that Transaction ID stay together in that group.
2. `GOVT_REJECTED` and blank IRN status: consolidate together into one PO group per `Ship From Fc`, including all SKUs from that same location.
3. Different Ship From FC values always produce different consolidated groups.
4. A rejected/blank row with no Ship From FC is not silently consolidated; it appears in the review queue until the location is supplied.
5. Original Transaction ID, Order ID, SKU, quantity, invoice and document references remain on the underlying STR lines for traceability.

The PO grouping view is a preparation/tracking aid. Actual PO creation is still performed in SAP by the business team. After PO, Delivery, Billing and MIGO references are available, upload the refreshed Amazon STR report to update those columns in the tracker. The app does not infer missing document numbers or treat blank document references as completed.

The CSV parser preserves duplicate source headers by assigning positional suffixes internally. Excel imports use the first visible worksheet and preserve the source row data. The first storage location and plant columns are treated as Ship From values; the second duplicate columns remain in the original raw source row.

## 7. Current scope and known limitations

This is a working starter project, not a certified production finance system. Before operational go-live:
- Validate exact Amazon report column headers against current exports.
- Add duplicate-file hash detection and safer idempotent shipment imports. Shipment rows are currently appended; repeated imports can duplicate lines.
- Configure date parsing for the actual date format in your exports.
- Confirm removal-fee amounts and currencies against the latest Amazon Removal Order Detail export before financial sign-off.
- Validate PO group totals and the SAP document references with a small test upload before operational use.
- Tighten RLS policies (especially employee update/delete rights), add manager approval rules, and test access with multiple users.
- Add source file storage upload if retention of original reports is required.
- Back up Supabase regularly and restrict GitHub repository access.

## Existing database upgrade

For an existing Control Tower installation, deploy the updated frontend files and run `excel_auto_mapping_migration.sql` once in Supabase SQL Editor. This additive migration allows blank STR SKU values, adds a stable source-line key and mapping status, and persists import rejection details. It preserves existing operational rows. Do not rerun `supabase_schema.sql` or the previously applied `str_po_logic_migration.sql`.

## Folder contents

- `index.html` — user interface
- `styles.css` — responsive layout
- `app.js` — Supabase client, imports, dashboard and workflows
- `config.example.js` — configuration template
- `supabase_schema.sql` — database and RLS schema
- `templates/amazon_str_template.csv` — exact Amazon STR header template, including the duplicate Ship From / Ship To SAP location headers.
- `str_po_logic_migration.sql` — earlier STR/PO migration (already applied in this installation).
- `excel_auto_mapping_migration.sql` — new additive migration for direct Excel imports, blank-SKU STR rows, and persisted import errors.


## Revised SKU-wise Warehouse Receipt Tracker (v2)
1. In Supabase SQL Editor, run `receipt_sku_mapping_migration.sql` before replacing the website files.
2. Upload the original product mapping master workbook directly from Reports & Imports using report type **Product Mapping Master**; no CSV conversion or manual SKU mapping is required.
3. Import Removal Order Detail first, then Removal Shipment Detail. The receipt selector now shows Order ID, SKU, FNSKU, tracking number and shipped quantity.
4. Warehouse Receipt reconciliation displays SKU, FNSKU, Material Code, tracking number, shipped quantity, each receipt category, physical received, short, not received, pending allocation and mapping status.
5. FNSKU mapping (or SKU fallback when FNSKU is absent) is applied only when it resolves to one unique SKU/material combination. Conflicting or missing mappings remain flagged and are not silently assigned.

Important: This is an incremental MVP correction, not a production certification. Existing shipment records will show mapped SKU/material dynamically from the mapping master where a unique match exists. The new database columns are used for future shipment imports. Validate Amazon source column names and mapping results with a small test import before bulk use. Duplicate shipment import prevention and a full exception-resolution workflow remain future hardening items.


### Amazon STR date handling update
The importer normalizes Amazon India date values in `DD/MM/YY HH:mm` and `DD/MM/YYYY` formats into ISO `YYYY-MM-DD` values before writing to the Supabase DATE columns. The time portion is intentionally omitted because `str_records.transaction_date` and `str_records.invoice_date` are DATE fields. No additional database migration is required for this date-format correction.
