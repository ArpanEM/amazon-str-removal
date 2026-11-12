-- =====================================================================
-- PO Tracker: warehouse receipts against PO number + SKU.
-- Run ONCE in Supabase SQL Editor (after the earlier parts).
-- Runs as one transaction: if any step fails, nothing is changed.
-- Deletes no data.
-- =====================================================================

begin;

create table if not exists public.po_receipts (
  id uuid primary key default gen_random_uuid(),
  po_number text not null,
  sku text not null,
  receipt_date date not null,
  good_qty numeric(14,3) not null default 0 check (good_qty >= 0),
  damaged_qty numeric(14,3) not null default 0 check (damaged_qty >= 0),
  near_expiry_qty numeric(14,3) not null default 0 check (near_expiry_qty >= 0),
  expired_qty numeric(14,3) not null default 0 check (expired_qty >= 0),
  short_qty numeric(14,3) not null default 0 check (short_qty >= 0),
  not_received_qty numeric(14,3) not null default 0 check (not_received_qty >= 0),
  remarks text,
  created_by uuid default auth.uid() references public.profiles(id),
  created_at timestamptz not null default now()
);

create index if not exists po_receipts_po_sku_idx on public.po_receipts (po_number, sku);

alter table public.po_receipts enable row level security;

-- Option A: everyone signed in can add and view; managers and admins can edit and delete.
drop policy if exists "read authenticated" on public.po_receipts;
create policy "read authenticated" on public.po_receipts
  for select to authenticated using (true);

drop policy if exists "insert authenticated" on public.po_receipts;
create policy "insert authenticated" on public.po_receipts
  for insert to authenticated with check (true);

drop policy if exists "update managers" on public.po_receipts;
create policy "update managers" on public.po_receipts
  for update to authenticated using (public.is_manager()) with check (public.is_manager());

drop policy if exists "delete managers" on public.po_receipts;
create policy "delete managers" on public.po_receipts
  for delete to authenticated using (public.is_manager());

-- Never allow more units to be received/short/not received than were sent for that PO + SKU.
create or replace function public.guard_po_receipt()
returns trigger
language plpgsql
set search_path = public
as $$
declare
  sent numeric;
  allocated numeric;
begin
  select coalesce(sum(coalesce(nullif(invoice_quantity, 0), quantity)), 0)
    into sent
  from public.str_records
  where po_number = new.po_number and sku = new.sku;

  select coalesce(sum(good_qty + damaged_qty + near_expiry_qty + expired_qty + short_qty + not_received_qty), 0)
    into allocated
  from public.po_receipts
  where po_number = new.po_number and sku = new.sku
    and id <> coalesce(new.id, '00000000-0000-0000-0000-000000000000'::uuid);

  allocated := allocated + new.good_qty + new.damaged_qty + new.near_expiry_qty
                         + new.expired_qty + new.short_qty + new.not_received_qty;

  if allocated > sent then
    raise exception 'Allocation (%) exceeds sent quantity (%) for PO % / SKU %', allocated, sent, new.po_number, new.sku;
  end if;
  return new;
end $$;

drop trigger if exists guard_po_receipt on public.po_receipts;
create trigger guard_po_receipt
  before insert or update on public.po_receipts
  for each row execute function public.guard_po_receipt();

-- Audit log for every change (uses the function created by the earlier update).
drop trigger if exists audit_po_receipts on public.po_receipts;
create trigger audit_po_receipts
  after insert or update or delete on public.po_receipts
  for each row execute function public.log_row_change();

commit;

-- ---------------------------------------------------------------------
-- CHECK (read-only). Duplicate STR lines would double-count PO quantities.
-- Expected: no rows.
-- ---------------------------------------------------------------------
select transaction_id, invoice_no, sku, quantity, invoice_quantity, count(*) as copies
from public.str_records
group by transaction_id, invoice_no, sku, quantity, invoice_quantity
having count(*) > 1;
