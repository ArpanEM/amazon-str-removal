# Amazon STR & Removal Control Tower (Emami Limited)

A static web app (GitHub Pages) with a Supabase database. It reconciles Amazon STR, SAP posting, Amazon removal invoices, shipments and warehouse receipts, and prepares the SAP SCAT upload files.

Stack: HTML, CSS and JavaScript (no build step), Supabase (database, login, file storage), Google Apps Script (approval mail, free).

## Files

| File | Purpose |
|---|---|
| `index.html`, `app.js`, `styles.css` | The app |
| `config.js` | Your Supabase URL and anon key. Keep your existing file. Never put the `service_role` key here |
| `assets/` | Emami logo files |
| `templates/` | Import templates |
| `apps_script/mailer.gs` | Approval mail and 2-day follow-up (runs in your Google account) |
| `sql/already_applied/` | Original schema and earlier scripts. **Do not run again** |
| `sql/` (numbered) | Scripts to run in order, once each |

## Database scripts: run order

Run in **Supabase > SQL Editor**, one file at a time, in this order. Skip any you have already run. Each one is safe to run twice, except where marked. If a paste is cut off, paste the file in two halves at a statement boundary.

| # | File | Adds |
|---|---|---|
| 01 | `po_tracker_setup.sql` | PO receipt table and guard |
| 02 | `po_date_setup.sql` | SAP PO created date |
| 03 | `po_units_setup.sql` | Sent units (STR quantity x Pack Of) |
| 04 | `shipment_line_seq_setup.sql` | Keeps genuine repeated shipment lines |
| 05 | `reconciliation_setup.sql` | Invoice tables |
| 06 | `mapping_update_setup.sql` | Product status (run once) |
| 07 | `warehouse_role_setup.sql` | Warehouse role |
| 08 | `warehouse_access_setup.sql` | Per-user page access |
| 09 | `batch_setup.sql` | Batch on receipts |
| 10 | `warehouse_import_setup.sql` | Rework, Mfg and Exp dates, import key |
| 11 | `warehouse_reimport_setup.sql` | Marks hand-edited receipts |
| 12 | `approval_setup.sql` | Approval mail tables and screenshot bucket |
| 13 | `followup_fc_setup.sql` | Follow-up tracking, FC mapping table |
| 14 | `fc_mapping_seed.sql` | Loads your 23 FC mappings |
| 15 | `scat_setup.sql` | SAP stock report and SCAT export log |
| 16 | `evidence_setup.sql` | Photo and video evidence storage |

After the last script, run `notify pgrst, 'reload schema';` once.

**Never run again:** `part1_cleanup.sql`, `cleanup_shipment_duplicates.sql`, `supabase_fix.sql`. They delete rows that look identical and would remove genuine repeated shipment lines. They are not in this package.

## Security setup

1. Authentication > Sign In / Providers: turn **off** "Allow new users to sign up". Add people with Authentication > Users > Invite user.
2. The free plan has **no automatic backups**. Every week export each table to CSV (Table Editor > Export) and keep the files in a separate folder.
3. Row Level Security is on for every table. Changes to operational tables are written to the audit log.

## Roles

| Role | Can do |
|---|---|
| admin | Everything, including roles and page access |
| manager | Add, edit, delete data. Approvals, SCAT, product master, FC mapping, stock report |
| employee | View and add records |
| warehouse | Warehouse Receipts only (by default). Edits only receipts they created |

Employees > Page access (admin only) sets which pages each user can open. A user with no custom list gets the role default.

## Daily use

**Imports (Reports & Imports).** STR report (Amazon dump), STR master (SAP-processed), Removal Shipment Detail, Product Mapping Master, SAP stock report, invoice ZIP files. Uploading the same file again does not duplicate rows. The Removal Order report is no longer shown.

**Product master.** Edit, add or remove products in the app. Mark products you no longer sell as Inactive instead of deleting them.

**Amazon FC to SAP location.** Reports & Imports lists each FC with its SAP virtual plant and storage location. Ship-to is fixed: plant DP50, storage location ED01.

**SAP SCAT** (managers and admins).
1. Upload the SAP stock report (Reports & Imports > SAP stock report). Upload a fresh one before each SCAT run.
2. Open SAP SCAT. Each STR removal line is checked and shown as Ready, Review or Blocked.
3. Template: IRN status `GOVT_ACCEPTED` uses the Govt Approved template. `GOVT_REJECTED` and blank use the Govt Rejected / Blank template.
4. Fields: ship-from plant and storage location come from the FC mapping. Material comes from the product master. Quantity is STR quantity x Pack Of. Value is taxable value divided by that quantity, rounded to 2 decimals.
5. Valuation type comes from the stock report by material and ship-from storage location. If a material has several valuation types there, the one with the most stock is used. If two are tied, you must choose one.
6. If the material has no stock at that location, a type is suggested from the same plant, or from other plants. These lines are marked **Review** and are not selected until you tick them.
7. In the Rejected / Blank file, lines with the same ship-from storage location form one PO number, with items 10, 20, 30. The Approved file uses the same grouping and item numbers (it has no PO number column).
8. Blocked reasons: FC not mapped, ASIN not in the product master, more than one material code in the master, missing taxable value, approved status without an IRN number, material not in the stock report with no type chosen.
9. Generate SCAT files downloads the files and records them. A line already exported is marked, and is shown only when you choose "All lines".

**Warehouse Receipts.** Pick an Order ID or Tracking Number. Enter quantities per shipment line. Good, Damaged and Rework open the batch popup, where each batch has its own quantity, Mfg and Exp date. Good stock expiring within 90 days is saved as Near Expiry, and past expiry as Expired. Damaged, Short and Not Received **require a photo or video**. Add or view evidence later with the Evidence button. A Missing tag marks lines without evidence.
Warehouse Excel upload reads the PHY DATA sheet. Re-uploading updates rows from that file and keeps rows you edited by hand.
Limits: photos 10 MB, videos 30 MB. Only managers and warehouse users can open evidence files.

**Approvals** (managers and admins). Damaged, Expired and Near Expiry lines with quantity and value (quantity x MRP) are mailed to the Ecom team. They reply by mail. Record the decision and upload the approval screenshot. Approved needs a screenshot. If there is no decision after 2 days, a follow-up mail is sent automatically, up to 3 times.
Sender name, To, CC and mailer details are remembered from the last mail.

**Apps Script mailer setup.**
1. script.google.com > New project, paste `apps_script/mailer.gs`.
2. Project Settings > Script properties: add `SUPABASE_URL`, `SUPABASE_SERVICE_KEY` (service role key, kept only here) and `MAILER_TOKEN` (any long private text).
3. Deploy > New deployment > Web app: Execute as Me, access Anyone. Copy the URL.
4. Triggers > Add trigger > `sendFollowUps` > Time-driven > Day timer.
5. In the app (Approvals) enter the URL and the same token.
Mail goes from the Google account that owns the script. Free accounts have a daily sending limit.

**Reconciliation.** One row per removal order and product: invoice, STR units, SAP references, shipped units, returned to Amazon, expected at warehouse, received, short, pending, tracking, fee and batches. Amazon "Delivered" does not mean received at the warehouse.

**Tracking.** The Track button in Shipment Tracker opens the Delhivery page directly. For Blue Dart it copies the number and opens the Blue Dart page, where you paste it.

**Filters and export.** Every table has column filters and click-to-sort headers. The first three columns stay fixed while scrolling sideways. Excel export follows the active filters.

## Troubleshooting

| Message | Cause and fix |
|---|---|
| Could not find the table or column ... in the schema cache | A script from the run order was not run. Run it, then `notify pgrst, 'reload schema';` |
| Allocation exceeds shipped quantity | Received plus short plus not received is above the shipped quantity |
| You do not have permission | The action needs a manager or admin |
| Page looks old after an update | Press Ctrl+F5, or open in a private window |
| Syntax error at end of input in SQL Editor | The paste was cut off. Paste the file in smaller parts |
| SCAT line is Blocked | Read the Note column. Fix the product master, FC mapping or stock report and reload |

## Testing

Not yet tested end to end on your live Supabase project. Test with one admin, one manager, one employee and one warehouse account before live use. Suggested checks:
- STR import keeps only FC_REMOVAL rows, and a second upload adds no duplicates.
- Invoice ZIP: RA key is the text between the first and second hyphen.
- Shipment re-upload gives no duplicates.
- A receipt over the shipped quantity is refused. Damaged without a photo is refused.
- SAP stock upload replaces the old one. SCAT files open with the same columns as your templates.
- Approval mail arrives with your sender name, CC works, the follow-up arrives after 2 days.
- A warehouse user sees only Warehouse Receipts. An employee cannot edit or delete.
- Audit log shows inserts, updates and deletes.
