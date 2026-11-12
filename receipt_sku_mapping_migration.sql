-- Run this once in Supabase SQL Editor BEFORE deploying the revised receipt tracker.
-- Adds SKU mapping fields to shipment lines. Safe to run more than once.
alter table public.removal_orders add column if not exists material_code text;
alter table public.shipment_lines
  add column if not exists material_code text,
  add column if not exists mapping_status text not null default 'unmapped';

-- Product mapping table is expected to exist from the original project schema.
create table if not exists public.product_mappings(
 id uuid primary key default gen_random_uuid(),
 asin text, sku text, fnsku text, material_code text, title text,
 case_size numeric(14,3), channel text, combo_check text, new_mrp numeric(14,2),
 mapping_version text, is_conflict boolean not null default false,
 raw_data jsonb, created_at timestamptz not null default now()
);
alter table public.product_mappings enable row level security;
drop policy if exists "read authenticated" on public.product_mappings;
create policy "read authenticated" on public.product_mappings
 for select to authenticated using (true);
drop policy if exists "insert authenticated" on public.product_mappings;
create policy "insert authenticated" on public.product_mappings
 for insert to authenticated with check (true);
