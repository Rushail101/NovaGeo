-- ============================================================================
-- Vendor Projects  (run once in the Supabase SQL editor)
--
--  vendor_projects          : a job / work / thing a vendor is doing for you
--  vendor_project_payments  : links a payment (bank debit, expense, partner-paid
--                             capex) -- or a project-only manual payment -- to a
--                             project. A payment can be split across projects
--                             by linking part of its amount to each.
--
-- The link columns copy the data type of your existing id columns
-- (bank_txns.id, expenses.id, capital_items.id), so this works whether those
-- are bigint, integer or uuid. Safe to re-run.
-- ============================================================================

create table if not exists public.vendor_projects (
  id             bigint generated always as identity primary key,
  project_no     text unique,
  vendor_name    text not null,
  title          text not null,
  description    text,
  category       text not null default 'misc',
  status         text not null default 'active'
                 check (status in ('planned','active','on_hold','completed','cancelled')),
  contract_value numeric(14,2) not null default 0,   -- 0 = open-ended / no fixed quote
  start_date     date,
  end_date       date,
  notes          text,
  created_at     timestamptz not null default now()
);

create index if not exists vendor_projects_vendor_idx on public.vendor_projects (lower(vendor_name));

-- Auto project number: VP-0001, VP-0002 ...
create or replace function public.vendor_projects_set_no() returns trigger
language plpgsql as $$
begin
  if new.project_no is null or new.project_no = '' then
    new.project_no := 'VP-' || lpad(new.id::text, 4, '0');
  end if;
  return new;
end $$;

drop trigger if exists vendor_projects_set_no_trg on public.vendor_projects;
create trigger vendor_projects_set_no_trg
  before insert on public.vendor_projects
  for each row execute function public.vendor_projects_set_no();

do $$
declare
  bt text; ex text; ci text;
begin
  select data_type into bt from information_schema.columns
   where table_schema='public' and table_name='bank_txns'     and column_name='id';
  select data_type into ex from information_schema.columns
   where table_schema='public' and table_name='expenses'      and column_name='id';
  select data_type into ci from information_schema.columns
   where table_schema='public' and table_name='capital_items' and column_name='id';

  execute format($f$
    create table if not exists public.vendor_project_payments (
      id              bigint generated always as identity primary key,
      project_id      bigint not null references public.vendor_projects(id) on delete cascade,
      bank_txn_id     %s references public.bank_txns(id)     on delete cascade,
      expense_id      %s references public.expenses(id)      on delete cascade,
      capital_item_id %s references public.capital_items(id) on delete cascade,
      amount          numeric(14,2) not null check (amount > 0),
      pay_date        date,
      payment_mode    text,
      reference       text,
      remarks         text,
      created_at      timestamptz not null default now(),
      -- at most one source; none = a manual, project-only payment
      check (
        (case when bank_txn_id     is not null then 1 else 0 end) +
        (case when expense_id      is not null then 1 else 0 end) +
        (case when capital_item_id is not null then 1 else 0 end) <= 1
      )
    )$f$,
    coalesce(bt,'bigint'), coalesce(ex,'bigint'), coalesce(ci,'bigint'));
end $$;

create index if not exists vpp_project_idx on public.vendor_project_payments (project_id);
create index if not exists vpp_bank_idx    on public.vendor_project_payments (bank_txn_id);
create index if not exists vpp_expense_idx on public.vendor_project_payments (expense_id);
create index if not exists vpp_capital_idx on public.vendor_project_payments (capital_item_id);

-- The app talks to Supabase with the anon key. Tables created in the SQL editor
-- have RLS off by default, matching a typical setup for this app. If your other
-- tables have RLS ON with open policies, mirror that here, e.g.:
--   alter table public.vendor_projects         enable row level security;
--   alter table public.vendor_project_payments enable row level security;
--   create policy "open" on public.vendor_projects         for all using (true) with check (true);
--   create policy "open" on public.vendor_project_payments for all using (true) with check (true);
