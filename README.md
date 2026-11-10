# Amazon STR & Removal Control Tower — Emami

A self-managed static web app for GitHub Pages + Supabase. It is designed for multiple employees and does **not** connect to SAP. Employees manually export SAP reports and upload them.

## Included modules

- Dashboard with month and financial-year filters.
- Amazon STR import: only `Transaction Type = FC_REMOVAL` is accepted into the operational STR table. `FC_TRANSFER` and `FC_REMOVAL-Cancel` are excluded.
- Removal Order Detail register.
- Removal Shipment Detail tracker, one row per package/tracking line. A generated internal database ID is distinct from Amazon's `tracking-number`.
- Warehouse receipt allocation: Good, Damaged, Near Expiry, Expired, Short, Not Received.
- Amazon Removal Fees: actual reported fees only, linked to known Amazon removal order IDs. No estimated rates and no product/inventory value KPI.
- SAP PO, Delivery, Invoice and MIGO manual CSV uploads with reference/material matching. Ambiguous/unmatched lines require review.
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
2. Open **SQL Editor** and run all of `supabase_schema.sql`.
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
1. Removal Order Detail CSV.
2. Removal Shipment Detail CSV.
3. STR report CSV (only FC_REMOVAL rows will be stored).
4. Amazon removal fee report CSV (after order IDs exist).
5. SAP PO / Delivery / Invoice / MIGO exports as CSV.

### Amazon CSV mapping

The importer recognizes common Amazon headers such as:
- Order: `order-id` / `order id` / `removal-order-id`
- Shipment tracking: `tracking-number`
- FNSKU: `fnsku`
- Quantity: `shipped-quantity` / `quantity-shipped` / `quantity`
- STR transaction type: `transaction-type`
- STR transaction ID: `transaction-id`

If your export uses different column names, update the `pick()` aliases in `app.js` or share a sample export to map the exact columns.

## 5. Amazon removal fee logic

Use the Amazon report's actual fee amount, currency, date and removal order ID. The import will skip fee lines that cannot be linked to an existing removal order. Do not load FBA storage, referral, selling, disposal or other non-removal charges into this table. If Amazon's report contains mixed fee types, filter to removal-related fee types before upload.

No fee is inferred from quantity. A blank fee remains absent; it is not converted to zero.

## 6. SAP manual upload and matching

- No SAP integration/API exists in this app.
- Download SAP exports manually and save as CSV.
- Select PO, Delivery, Invoice or MIGO stage.
- The starter parser looks for common reference and material headers (`PO`, `Delivery`, `Invoice`, `MIGO`, `Material`).
- To make matching reliable, configure aliases to the actual SAP export column names and define the authoritative matching keys with your team. Recommended hierarchy:
  1. PO + Material
  2. Delivery + Material
  3. Invoice + Material
  4. MIGO document + Material
- Do not auto-resolve ambiguous matches. Review exceptions manually.

**Important limitation:** The current starter persists parsed SAP reference rows and provides the tracker structure, but it does not yet implement a fully validated cross-document exception workflow or exact SAP-specific matching for your company exports. Sample PO, Delivery, Invoice and MIGO files are required to finish that mapping accurately.

## 7. Current scope and known limitations

This is a working starter project, not a certified production finance system. Before operational go-live:
- Validate exact Amazon report column headers against current exports.
- Add duplicate-file hash detection and safer idempotent shipment imports. Shipment rows are currently appended; repeated imports can duplicate lines.
- Configure date parsing for the actual date format in your exports.
- Configure fee type filters and currency handling for the exact Amazon fee report.
- Build a reviewed exception queue for SAP matches and fee lines.
- Tighten RLS policies (especially employee update/delete rights), add manager approval rules, and test access with multiple users.
- Add source file storage upload if retention of original reports is required.
- Back up Supabase regularly and restrict GitHub repository access.

## Folder contents

- `index.html` — user interface
- `styles.css` — responsive layout
- `app.js` — Supabase client, imports, dashboard and workflows
- `config.example.js` — configuration template
- `supabase_schema.sql` — database and RLS schema
- `templates/` — sample CSV headers for each import
