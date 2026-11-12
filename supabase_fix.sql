-- =====================================================================
-- Amazon STR & Removal Control Tower: permissions, audit log and
-- duplicate protection (Option A).
--
-- RUN ONCE in Supabase SQL Editor, AFTER the four earlier migrations.
-- Option A: employees can add and view records;
--           managers and admins can edit and delete records.
--
-- This file deletes NO data. If duplicate rows exist, Step 1 stops the
-- script and nothing is changed.
-- =====================================================================


-- ---------------------------------------------------------------------
-- STEP 1: Safety check. Stops the whole script if duplicates exist.
-- ---------------------------------------------------------------------
do $$
begin
  if exists (
    select 1 from public.shipment_lines
    group by order_id, tracking_number, sku, fnsku
    having count(*) > 1
  ) then
    raise exception 'STOPPED. Duplicate shipment_lines exist. Nothing was changed. Send the duplicate query output before continuing.';
  end if;

  if exists (
    select 1 from public.amazon_removal_fees
    group by order_id, sku, fnsku
    having count(*) > 1
  ) then
    raise exception 'STOPPED. Duplicate amazon_removal_fees exist. Nothing was changed. Send the duplicate query output before continuing.';
  end if;
end $$;


-- ---------------------------------------------------------------------
-- STEP 2: Duplicate protection (unique keys).
-- ---------------------------------------------------------------------
create unique index if not exists shipment_lines_identity_uidx
  on public.shipment_lines (order_id, tracking_number, sku, fnsku);

drop index if exists public.amazon_removal_fees_import_key;
create unique index if not exists amazon_removal_fees_identity_uidx
  on public.amazon_removal_fees (order_id, sku, fnsku);

-- Needed for the STR upload's upsert (already created by the Excel migration).
create unique index if not exists str_records_source_line_key_uidx
  on public.str_records (source_line_key);


-- ---------------------------------------------------------------------
-- STEP 3: Permissions for operational tables (Option A).
--   select / insert : any signed-in user (kept from the original schema)
--   update / delete : managers and admins only
-- ---------------------------------------------------------------------
do $$
declare t text;
begin
  foreach t in array array[
    'removal_orders','shipment_lines','warehouse_receipts','amazon_removal_fees',
    'amazon_claims','sap_postings','import_batches','str_records','product_mappings'
  ] loop
    execute format('drop policy if exists %I on public.%I', 'update authenticated', t);
    execute format('drop policy if exists %I on public.%I', 'update managers', t);
    execute format(
      'create policy %I on public.%I for update to authenticated using (public.is_manager()) with check (public.is_manager())',
      'update managers', t);

    execute format('drop policy if exists %I on public.%I', 'delete managers', t);
    execute format(
      'create policy %I on public.%I for delete to authenticated using (public.is_manager())',
      'delete managers', t);
  end loop;
end $$;


-- ---------------------------------------------------------------------
-- STEP 4: Audit log for every insert, update and delete on operational
-- tables. Writes to the existing audit_logs table.
-- ---------------------------------------------------------------------
create or replace function public.log_row_change()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if tg_op = 'INSERT' then
    insert into public.audit_logs (user_id, action, entity, details)
    values (auth.uid(), 'insert', tg_table_name,
            jsonb_build_object('record_id', new.id));
  elsif tg_op = 'UPDATE' then
    insert into public.audit_logs (user_id, action, entity, details)
    values (auth.uid(), 'update', tg_table_name,
            jsonb_build_object('record_id', new.id, 'old', to_jsonb(old), 'new', to_jsonb(new)));
  else
    insert into public.audit_logs (user_id, action, entity, details)
    values (auth.uid(), 'delete', tg_table_name,
            jsonb_build_object('record_id', old.id, 'old', to_jsonb(old)));
  end if;
  return coalesce(new, old);
end $$;

do $$
declare t text;
begin
  foreach t in array array[
    'removal_orders','shipment_lines','warehouse_receipts','amazon_removal_fees',
    'amazon_claims','sap_postings','import_batches','str_records','product_mappings'
  ] loop
    execute format('drop trigger if exists audit_%I on public.%I', t, t);
    execute format(
      'create trigger audit_%I after insert or update or delete on public.%I for each row execute function public.log_row_change()',
      t, t);
  end loop;
end $$;


-- ---------------------------------------------------------------------
-- STEP 5: Profiles. Each user sees their own profile; managers and
-- admins see all profiles (needed for the Employees and Audit screens).
-- ---------------------------------------------------------------------
drop policy if exists "profiles read" on public.profiles;
create policy "profiles read own or manager" on public.profiles
  for select to authenticated
  using (id = auth.uid() or public.is_manager());


-- ---------------------------------------------------------------------
-- STEP 6: Uploaded source reports. Only managers and admins can read them.
-- ---------------------------------------------------------------------
drop policy if exists "authenticated read source reports" on storage.objects;
create policy "managers read source reports" on storage.objects
  for select to authenticated
  using (bucket_id = 'source-reports' and public.is_manager());


-- ---------------------------------------------------------------------
-- CHECK (read-only): shows the policies now in force on operational tables.
-- Expected: "update managers" and "delete managers" on each table, and
-- no "update authenticated" rows.
-- ---------------------------------------------------------------------
select tablename, policyname, cmd, roles
from pg_policies
where schemaname = 'public'
  and tablename in ('removal_orders','shipment_lines','warehouse_receipts','amazon_removal_fees',
                    'amazon_claims','sap_postings','import_batches','str_records','product_mappings','profiles')
order by tablename, cmd, policyname;

-- Separate check for the STR table (run any time; read-only).
-- Duplicates here would double-count PO and invoice quantities.
-- select transaction_id, sku, count(*) from public.str_records
--   where sku is not null group by 1, 2 having count(*) > 1;
