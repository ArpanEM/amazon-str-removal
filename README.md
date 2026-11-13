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
4. Edit `config.js` and paste your Project URL and `anon` public key. Never paste the `service_role` key.
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

**Security note:** GitHub Pages is static hosting. The Supabase anon key is visible in browser source by design; database security depends on Row Level Security. Permissions are set by `supabase_fix.sql` (see the Security update section).

## 4. Import order

Recommended initial upload order:
1. Product Mapping Master Excel or CSV (upload the original `.xlsx`, `.xls` or `.csv` file).
2. Removal Order Detail Excel or CSV.
3. Removal Shipment Detail Excel or CSV.
4. Amazon STR report Excel or CSV (only FC_REMOVAL rows will be stored). This same file is used for PO grouping and carries the SAP document reference columns. The SAP Material Code and Material Description shown in the trackers come from the Product Mapping Master (its `Material Code` and `Title` columns), matched by FNSKU, then SKU, then ASIN. If no mapping is found, the tracker shows "Not mapped" rather than the SKU.

Removal fees are imported automatically during step 1 from the `removal-fee` and `currency` columns in the Removal Order Detail CSV. No separate fee report is available or needed.

### Excel / CSV upload and automatic mapping

The import screen accepts `.xlsx`, `.xls` and `.csv` directly. Excel files are read in the browser using SheetJS; the first visible worksheet is imported. Do not convert Excel files manually. The STR/SAP workbook is the reference source for PO Number, Delivery No, Billing Document and MIGO No. The SKU column is the SAP Material Code. Material Description is not available in the supplied source header, so it remains blank. Existing product mapping imports remain optional for enriching ASIN/FNSKU identifiers; they are not required to derive the SAP Material Code when SKU is present.

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
1. Keep each Transaction ID / Invoice reference as its own PO group, retaining all SKU lines associated with that transaction and invoice.
2. Do not consolidate or split PO groups based on `GOVT_ACCEPTED`, `GOVT_REJECTED` or blank IRN status. IRN status remains a displayed source attribute only.
3. PO Number, Delivery No, Billing Document and MIGO No are read from the STR/SAP Excel or CSV source columns. Missing values remain pending and are visible in the review queue.
4. SAP Material Code and Material Description come from the Product Mapping Master (`Material Code` and `Title`), matched by FNSKU, then SKU, then ASIN. Unmatched rows show "Not mapped".

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

## Security update: permissions and duplicate protection (read first)

This update changes how employees, managers and admins can work with data. Follow these steps in order.

### Step 1: Back up your data (Free plan)
The Supabase Free plan does not include automatic backups. Before running the SQL, export your data:
- In **Table Editor**, open each table and use **Export** (CSV), or run `select * from public.<table>;` in the SQL Editor and save the result.
- Tables to export: `removal_orders`, `shipment_lines`, `warehouse_receipts`, `amazon_removal_fees`, `amazon_claims`, `str_records`, `import_batches`, `product_mappings`, `sap_postings`.

### Step 2: Turn off public sign-ups
In Supabase, go to **Authentication → Sign In / Providers** and turn off **Allow new users to sign up**. Add employees yourself through **Authentication → Users → Invite user**. Anyone who can sign up gets access to the data, so this step matters.

### Step 3: Run the SQL once
Open **SQL Editor**, paste the full contents of `supabase_fix.sql`, and run it. This is the only SQL file to run.
- The file runs as one transaction. If any step fails, nothing is changed. Copy the error message and send it to your developer.
- Part 1 removes 3 identical duplicate shipment rows from the first removal shipment import (`473573020769.csv`). Shipment lines that have warehouse receipts are never removed.
- Part 2 applies the permissions, audit log and duplicate protection.
- At the end, the CHECK query lists the policies now in force. Each operational table should show `update managers` and `delete managers`, and no `update authenticated`.
- Do not re-run the original schema or the earlier four migrations, which are already applied.

### Step 3a: Confirm the shipment count
In SQL Editor, run `select count(*) from public.shipment_lines;`. If only the 473573020769.csv file has been imported for shipments, the count should be 48.

### Step 4: Upload the updated files
Upload `app.js`, `index.html`, `supabase_fix.sql` and `README.md` to the GitHub repository, replacing the old copies. Wait for GitHub Pages to rebuild, then open the app and press Ctrl+F5 (or Cmd+Shift+R on Mac) to clear the old script from the browser cache.

### Step 5: Test with a non-admin account
Sign in as an employee account and check:
1. You can view all pages, and you can save a warehouse receipt.
2. The Employees page shows roles as "View only" (no dropdown).
3. Re-importing a file that contains existing rows reports them as "Already present and skipped", not as errors.
4. Editing or deleting an existing record in Supabase fails with a permission message.
5. Signing in as a manager or admin, the same re-import updates existing rows, and the Employees page shows the role dropdown (admin only).

### Permissions after this update

| Action | Employee | Manager / Admin |
|---|---|---|
| View all records | Yes | Yes |
| Add records (imports, receipts, claims) | Yes | Yes |
| Re-import an existing row (updates it) | No: skipped and reported | Yes |
| Edit or delete records | No | Yes |
| Change roles | No | Admin only |
| View audit log, employees, uploaded source reports | No | Yes |

### What this update does not change yet
- **STR upload key.** STR rows are still identified by the older key, which includes the row position in the file. Changing it now would insert your existing STR rows again. This will be handled in a separate step after a duplicate review. Until then, after each STR upload run the read-only check at the end of `supabase_fix.sql`.
- **Audit log size.** Every insert, update and delete is logged. Large imports create many audit entries.
- **Duplicate protection on imports.** Shipments are identified by order, tracking number, SKU, FNSKU and shipped quantity, so the same line cannot be added twice, while two real lines with different quantities (for example 9 and 3 units on one tracking number) are both kept. Fees are identified by order, SKU and FNSKU. Fee rows are no longer separated by filename.

## PO Tracker (current)

The **PO & SAP Tracker** page shows one row per PO and SKU:
- PO, Delivery, Billing Document and MIGO from the STR / SAP master.
- SAP Material Code and product title from the Product Mapping Master.
- PO date (the invoice date in the master).
- Quantity sent (Invoice Quantity; falls back to Quantity when blank).
- Quantity received at the warehouse (Good + Damaged + Near Expiry + Expired).
- Short / Not received quantity, and Pending quantity (sent minus all of these).
- Pending days (today minus the PO date, for open lines).
- Status: Awaiting PO, Pending, Partial or Complete.

**Recording receipts:** use the "Record warehouse receipt against a PO" panel. Type the PO number, pick it from the suggestions, enter the quantities for each SKU line, and save. The database refuses any entry that would exceed the quantity sent.

**Re-uploading the STR master:** you can upload the same master again, or a newer one. Existing lines are matched by Transaction ID, Invoice, SKU, quantity and invoice quantity, so they are not added twice. PO, Delivery, Billing, MIGO and status values are filled in or updated for existing lines (managers and admins update existing rows; employees see those lines counted as already present).

**One-time database step:** run `po_tracker_setup.sql` once in SQL Editor. It creates the receipt table, its permissions and the audit log. It deletes no data.

## Invoices and reconciliation (current)

**Invoice upload:** Reports & Imports → *Amazon removal tax invoices (bulk ZIP)*. Upload one or more ZIP files (or single PDFs). Each PDF is read in the browser. Invoice number, invoice date, shipment ID, RA number, ASIN and order number are extracted. Invoices already in the system and ZIP files already uploaded are skipped. The result shows read, accepted, updated, duplicates, failed and unmatched counts.

**RA matching key:** the text between the first and second hyphen of the RA number. Example: `ZAXQT-Xqa4dfs578-BOM5-1` gives `Xqa4dfs578`. The prefix `ZAXQT-` and the location and final suffix are dropped. If the RA number does not follow this structure, the invoice is flagged for review instead of being guessed.

**Matching:** invoice RA key = Removal Order ID (Removal Order Detail). Invoice Shipment ID = STR Transaction ID (FC_REMOVAL only). Invoice ASIN is linked to the FNSKU through the Product Mapping Master.

**Reconciliation page:** one row per removal order and product (FNSKU). It shows the invoice, STR units (STR Quantity × Pack Of), SAP references, requested and shipped units, shipment total, returned to Amazon, expected at warehouse (shipped minus returned to Amazon), physical received, short and not received, pending, tracking numbers, removal fee and currency. Status is Matched, Not shipped, No invoice, STR not found, or a variance.

**Rules:** Amazon-returned units are not expected at the warehouse. Removal fees come only from the removal fee data and are never shown as zero when blank. Product title comes from the Product Mapping Master.

**One-time database step:** run `reconciliation_setup.sql` once in SQL Editor. It creates the invoice and ZIP import tables, their permissions and audit log, and deletes no data.

**Note:** invoice PDFs must contain selectable text. Scanned PDFs will fail and are counted as failed.

## Existing database upgrade

For an existing Control Tower installation, deploy the updated frontend files in this package. This frontend revision uses the existing database columns and does not require a new SQL migration. Do not rerun `supabase_schema.sql` or previously applied migrations.

## Folder contents

- `index.html` — user interface
- `styles.css` — responsive layout
- `app.js` — Supabase client, imports, dashboard and workflows
- `config.js` — your Supabase URL and anon key (edit directly)
- `supabase_fix.sql` — permissions, audit log and duplicate protection (run once)
- `supabase_schema.sql` — database and RLS schema
- `templates/amazon_str_template.csv` — exact Amazon STR header template, including the duplicate Ship From / Ship To SAP location headers.
- `str_po_logic_migration.sql` — earlier STR/PO migration (already applied in this installation).
- `excel_auto_mapping_migration.sql` — new additive migration for direct Excel imports, blank-SKU STR rows, and persisted import errors.


## Revised SKU-wise Warehouse Receipt Tracker (v2)
1. In Supabase SQL Editor, run `receipt_sku_mapping_migration.sql` before replacing the website files.
2. Upload the original product mapping master workbook directly from Reports & Imports using report type **Product Mapping Master**; no CSV conversion or manual SKU mapping is required.
3. Import Removal Order Detail first, then Removal Shipment Detail. The receipt selector now shows Order ID, SKU, FNSKU, tracking number and shipped quantity.
4. Warehouse Receipt reconciliation displays SKU/FNSKU, SAP Material Code (from the Product Mapping Master), tracking number, shipped quantity, each receipt category, physical received, short, not received and pending allocation. Order and shipment trackers show SAP document references where a matching STR row exists. The report tables can be exported to `.xlsx` and respect the currently displayed month/year filters.
5. FNSKU mapping (or SKU fallback when FNSKU is absent) is applied only when it resolves to one unique SKU/material combination. Conflicting or missing mappings remain flagged and are not silently assigned.

Important: This is an incremental MVP correction, not a production certification. Existing shipment records will show mapped SKU/material dynamically from the mapping master where a unique match exists. The new database columns are used for future shipment imports. Validate Amazon source column names and mapping results with a small test import before bulk use. Duplicate shipment import prevention and a full exception-resolution workflow remain future hardening items.


### Amazon STR date handling update
The importer normalizes Amazon India date values in `DD/MM/YY HH:mm` and `DD/MM/YYYY` formats into ISO `YYYY-MM-DD` values before writing to the Supabase DATE columns. The time portion is intentionally omitted because `str_records.transaction_date` and `str_records.invoice_date` are DATE fields. No additional database migration is required for this date-format correction.


## Tracker controls and warehouse receipt workflow (12 Nov 2026 update)

- Removal Orders, Shipment Tracker, and Warehouse Receipt Reconciliation now provide per-column filters under each table header. Click a column heading to sort ascending/descending; filters can be combined. Existing page-level search and financial-year/month filters remain available.
- Export buttons export the currently visible rows after the active table filters/sort.
- Warehouse Receipts now starts with a lookup mode: choose Amazon Order ID or Tracking Number, then choose the identifier. The page loads every shipment line associated with that identifier, including SKU/SAP Material Code, FNSKU, tracking number, Amazon status and Amazon shipped quantity.
- Enter Good, Damaged, Near Expiry, Expired, Short and Not Received quantities for each line, then save all lines together. The app validates each line against its shipped quantity plus previous allocations. Amazon delivery status remains separate from physical warehouse receipt.
- SAP Material Code and Material Description come from the Product Mapping Master. Unmatched rows show "Not mapped".
- The Product Mapping Master's `Product Title` column is read as the product description. If product descriptions were blank after an earlier import, run this once in SQL Editor (no re-import needed):

```sql
update public.product_mappings
set title = raw_data->>'Product Title'
where (title is null or trim(title) = '')
  and coalesce(trim(raw_data->>'Product Title'), '') <> '';
```


### Horizontal report navigation
Wide report tables now display **Left** and **Right** scroll buttons whenever columns extend beyond the visible area. The buttons move the report horizontally in smooth steps; the left/right controls disable automatically at the respective ends. This is available across report and tracker tables, including Removal Orders, Shipment Tracker and Warehouse Receipt reconciliation.
