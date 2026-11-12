-- Amazon STR + PO grouping logic migration
-- Run once in Supabase SQL Editor after the original schema and prior migrations.
alter table public.str_records
  add column if not exists ship_from_fc text,
  add column if not exists ship_from_storage_location text,
  add column if not exists ship_from_plant_code text,
  add column if not exists sku text,
  add column if not exists asin text,
  add column if not exists order_id text,
  add column if not exists receiver_gstin text,
  add column if not exists actual_qty numeric(14,3),
  add column if not exists invoice_quantity numeric(14,3),
  add column if not exists irn_status text,
  add column if not exists po_number text,
  add column if not exists delivery_no text,
  add column if not exists billing_document text,
  add column if not exists migo_no text,
  add column if not exists sap_status text,
  add column if not exists invoice_value numeric(14,2),
  add column if not exists taxable_value numeric(14,2),
  add column if not exists invoice_date date;

-- The Amazon STR source has SKU but may not contain SAP Material Code.
-- Use transaction + SKU as the stable source-line key.
alter table public.str_records drop constraint if exists str_records_transaction_id_material_code_key;
do $$ begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'str_records_transaction_id_sku_key'
      and conrelid = 'public.str_records'::regclass
  ) then
    alter table public.str_records
      add constraint str_records_transaction_id_sku_key unique (transaction_id, sku);
  end if;
end $$;

create index if not exists str_records_ship_from_idx on public.str_records(ship_from_fc);
create index if not exists str_records_irn_idx on public.str_records(irn_status);
create index if not exists str_records_po_idx on public.str_records(po_number);
