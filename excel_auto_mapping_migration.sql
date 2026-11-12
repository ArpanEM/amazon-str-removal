-- Excel upload + automatic mapping support migration
-- Run ONCE in Supabase SQL Editor after the already-applied project migrations.
-- Additive/safe: preserves existing rows and existing application tables.

-- STR rows can legitimately have no SKU in Amazon's source export.
alter table public.str_records alter column sku drop not null;
alter table public.str_records add column if not exists source_line_key text;
alter table public.str_records add column if not exists mapping_status text not null default 'unmapped';

-- Use a stable source-line key for idempotent imports, including rows with blank SKU.
alter table public.str_records drop constraint if exists str_records_transaction_id_sku_key;
create unique index if not exists str_records_source_line_key_uidx
  on public.str_records(source_line_key);

create index if not exists str_records_mapping_status_idx on public.str_records(mapping_status);
create index if not exists product_mappings_asin_idx on public.product_mappings(asin);
create index if not exists product_mappings_sku_idx on public.product_mappings(sku);
create index if not exists product_mappings_fnsku_idx on public.product_mappings(fnsku);

-- Persist row-level import failures so users can review them after refresh.
alter table public.import_batches add column if not exists rejected_count integer not null default 0;
alter table public.import_batches add column if not exists error_details jsonb not null default '[]'::jsonb;
