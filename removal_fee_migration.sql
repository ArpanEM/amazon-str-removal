-- Run once in Supabase SQL Editor before deploying this fee-capture update.
-- Stores Amazon Removal Order Detail fee lines at order + SKU + FNSKU + source-file level.
alter table public.amazon_removal_fees
  add column if not exists sku text not null default '',
  add column if not exists fnsku text not null default '';

create unique index if not exists amazon_removal_fees_import_key
  on public.amazon_removal_fees(order_id, sku, fnsku, source_file);
