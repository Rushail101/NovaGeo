-- ============================================================================
-- OPTIONAL — permanently deletes the plant-operations tables and ALL their data
-- (weighments, production, lots, shift logs, boulder receipts, maintenance,
-- QC, samples, cost config, vehicles, customers, grades).
--
-- Only run this when you're sure you no longer need that history, e.g. because
-- the operations ERP now lives elsewhere. EXPORT / BACKUP FIRST (Supabase
-- dashboard > Database > Backups, or export each table to CSV).
--
-- Kept: suppliers, purchases, expenses, capital_items, bank_*, partner_cashbook
-- and the new vendor_projects tables.
-- Order is children-first. Wrapped in a transaction: if anything fails, nothing
-- is dropped. Change ROLLBACK to COMMIT only after you've reviewed it.
-- ============================================================================
begin;

drop table if exists public.shift_log_outputs;
drop table if exists public.shift_logs;
drop table if exists public.production_entry_outputs;
drop table if exists public.production_entries;
drop table if exists public.lot_grade_outputs;
drop table if exists public.lots;
drop table if exists public.weighments;
drop table if exists public.boulder_receipts;
drop table if exists public.maint_logs;
drop table if exists public.maint_tasks;
drop table if exists public.equipment;
drop table if exists public.qc_tests;
drop table if exists public.samples;
drop table if exists public.cost_config;
drop table if exists public.vehicles;
drop table if exists public.customers;
drop table if exists public.grades;

rollback;  -- <- switch to COMMIT; to actually apply
