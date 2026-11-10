-- Amazon STR & Removal Control Tower
-- Run in Supabase SQL Editor. Review RLS and bootstrap admin steps in README.
create extension if not exists pgcrypto;

create table if not exists public.profiles(
 id uuid primary key references auth.users(id) on delete cascade,
 email text not null,
 full_name text,
 role text not null default 'employee' check(role in ('admin','manager','employee')),
 created_at timestamptz not null default now()
);
create or replace function public.handle_new_user() returns trigger language plpgsql security definer set search_path=public as $$
begin
 insert into public.profiles(id,email,full_name,role)
 values(new.id,new.email,coalesce(new.raw_user_meta_data->>'full_name',split_part(new.email,'@',1)),'employee')
 on conflict(id) do nothing;
 return new;
end; $$;
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created after insert on auth.users for each row execute procedure public.handle_new_user();

create or replace function public.current_role() returns text language sql stable security definer set search_path=public as $$
 select role from public.profiles where id=auth.uid()
$$;
create or replace function public.is_manager() returns boolean language sql stable security definer set search_path=public as $$
 select coalesce(public.current_role() in ('admin','manager'),false)
$$;
create or replace function public.is_admin() returns boolean language sql stable security definer set search_path=public as $$
 select coalesce(public.current_role()='admin',false)
$$;

create table if not exists public.import_batches(
 id uuid primary key default gen_random_uuid(),file_name text not null,report_type text not null,
 rows_read integer not null default 0,rows_accepted integer not null default 0,
 imported_by uuid references public.profiles(id),imported_by_email text,created_at timestamptz not null default now()
);
create table if not exists public.str_records(
 id uuid primary key default gen_random_uuid(),transaction_id text not null,invoice_no text,material_code text,
 transaction_date date,quantity numeric(14,3) not null default 0,hsn text,gst_rate numeric(8,3),
 transaction_type text not null check(transaction_type='FC_REMOVAL'),raw_data jsonb,created_at timestamptz not null default now(),
 unique(transaction_id,material_code)
);
create table if not exists public.removal_orders(
 id uuid primary key default gen_random_uuid(),order_id text not null,sku text not null default '',fnsku text not null default '',
 requested_qty numeric(14,3) not null default 0,cancelled_qty numeric(14,3) not null default 0,
 shipped_qty numeric(14,3) not null default 0,in_process_qty numeric(14,3) not null default 0,
 status text,request_date date,raw_data jsonb,created_at timestamptz not null default now(),
 unique(order_id,sku,fnsku)
);
create table if not exists public.shipment_lines(
 id uuid primary key default gen_random_uuid(),order_id text not null,tracking_number text not null,
 sku text,fnsku text,shipped_qty numeric(14,3) not null default 0,amazon_status text,
 shipment_date date,disposition text,raw_data jsonb,created_at timestamptz not null default now()
);
create index if not exists shipment_order_idx on public.shipment_lines(order_id);
create index if not exists shipment_tracking_idx on public.shipment_lines(tracking_number);
create table if not exists public.warehouse_receipts(
 id uuid primary key default gen_random_uuid(),shipment_line_id uuid not null references public.shipment_lines(id) on delete cascade,
 good_qty numeric(14,3) not null default 0,damaged_qty numeric(14,3) not null default 0,
 near_expiry_qty numeric(14,3) not null default 0,expired_qty numeric(14,3) not null default 0,
 short_qty numeric(14,3) not null default 0,not_received_qty numeric(14,3) not null default 0,
 receipt_date date not null,remarks text,created_by uuid references public.profiles(id),created_at timestamptz not null default now(),
 check(good_qty>=0 and damaged_qty>=0 and near_expiry_qty>=0 and expired_qty>=0 and short_qty>=0 and not_received_qty>=0)
);
create or replace function public.guard_receipt_allocation() returns trigger language plpgsql as $$
declare shipped numeric; allocated numeric;
begin
 select shipped_qty into shipped from public.shipment_lines where id=new.shipment_line_id for update;
 select coalesce(sum(good_qty+damaged_qty+near_expiry_qty+expired_qty+short_qty+not_received_qty),0)
 into allocated from public.warehouse_receipts where shipment_line_id=new.shipment_line_id and id<>coalesce(new.id,'00000000-0000-0000-0000-000000000000'::uuid);
 allocated:=allocated+new.good_qty+new.damaged_qty+new.near_expiry_qty+new.expired_qty+new.short_qty+new.not_received_qty;
 if allocated>shipped then raise exception 'Final receipt allocation (%) exceeds shipped quantity (%)',allocated,shipped; end if;
 return new;
end; $$;
drop trigger if exists receipt_allocation_guard on public.warehouse_receipts;
create trigger receipt_allocation_guard before insert or update on public.warehouse_receipts for each row execute procedure public.guard_receipt_allocation();

-- Fee table is deliberately restricted to Amazon removal order references.
create table if not exists public.amazon_removal_fees(
 id uuid primary key default gen_random_uuid(),order_id text not null,
 fee_amount numeric(14,2) not null check(fee_amount>=0),currency text not null default 'INR',
 fee_date date,source_file text,raw_data jsonb,created_at timestamptz not null default now()
);
create index if not exists removal_fee_order_idx on public.amazon_removal_fees(order_id);
create table if not exists public.amazon_claims(
 id uuid primary key default gen_random_uuid(),order_id text not null,claim_ref text,claim_type text not null,
 quantity numeric(14,3) not null default 0,status text not null default 'Open',amount numeric(14,2),remarks text,
 created_at timestamptz not null default now()
);
create table if not exists public.sap_postings(
 id uuid primary key default gen_random_uuid(),reference_no text not null,material_code text not null default '',
 po_no text,delivery_no text,invoice_no text,migo_doc text,source_stage text,source_file text,raw_data jsonb,
 updated_at timestamptz not null default now(),unique(reference_no,material_code)
);
create table if not exists public.product_mappings(
 id uuid primary key default gen_random_uuid(),asin text,sku text,fnsku text,material_code text,title text,
 case_size numeric(14,3),channel text,combo_check text,new_mrp numeric(14,2),mapping_version text,
 is_conflict boolean not null default false,raw_data jsonb,created_at timestamptz not null default now()
);
create table if not exists public.audit_logs(
 id uuid primary key default gen_random_uuid(),user_id uuid references public.profiles(id),action text not null,
 entity text not null,details jsonb,created_at timestamptz not null default now()
);

-- Enable RLS on all application tables.
do $$ declare t text; begin
 foreach t in array array['profiles','import_batches','str_records','removal_orders','shipment_lines','warehouse_receipts','amazon_removal_fees','amazon_claims','sap_postings','product_mappings','audit_logs'] loop
  execute format('alter table public.%I enable row level security',t);
 end loop;
end $$;

-- Signed-in employees can read operational data and create/update records.
drop policy if exists "profiles read" on public.profiles;
create policy "profiles read" on public.profiles for select to authenticated using (true);
drop policy if exists "profile update admin" on public.profiles;
create policy "profile update admin" on public.profiles for update to authenticated using(public.is_admin()) with check(public.is_admin());

do $$ declare t text; begin
 foreach t in array array['import_batches','str_records','removal_orders','shipment_lines','warehouse_receipts','amazon_removal_fees','amazon_claims','sap_postings','product_mappings'] loop
  execute format('drop policy if exists %I on public.%I','read authenticated',t);
  execute format('create policy %I on public.%I for select to authenticated using(true)','read authenticated',t);
  execute format('drop policy if exists %I on public.%I','insert authenticated',t);
  execute format('create policy %I on public.%I for insert to authenticated with check(true)','insert authenticated',t);
  execute format('drop policy if exists %I on public.%I','update authenticated',t);
  execute format('create policy %I on public.%I for update to authenticated using(true) with check(true)','update authenticated',t);
 end loop;
end $$;
create policy "audit read managers" on public.audit_logs for select to authenticated using(public.is_manager());
create policy "audit insert own" on public.audit_logs for insert to authenticated with check(user_id=auth.uid());

-- Private storage bucket for uploaded source exports. Optional: this build records source filename
-- and raw rows in DB; add object storage upload if you want original files retained.
insert into storage.buckets(id,name,public) values('source-reports','source-reports',false) on conflict(id) do nothing;
create policy "authenticated upload source reports" on storage.objects for insert to authenticated with check(bucket_id='source-reports');
create policy "authenticated read source reports" on storage.objects for select to authenticated using(bucket_id='source-reports');
