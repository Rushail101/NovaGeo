// Data layer for GST, TDS and the compliance calendar (tables from supabase/002_compliance.sql).
import { supabase } from "./supabase.js";

async function run(p) {
  const { data, error } = await p;
  if (error) throw error;
  return data;
}
export const isMissingTable = e => /relation .* does not exist|Could not find the table|schema cache/i.test(e?.message || "");

// ── Company settings ──────────────────────────────────────────────────────────
export const DEFAULT_SETTINGS = {
  legal_name: "Nova Geo Resources LLP", gstin: "36AAYFN2455Q1Z8", pan: "AAYFN2455Q", tan: "", llpin: "",
  state: "Telangana", state_code: "36", address: "", office_address: "", phone: "", email: "",
  bank_name: "", bank_account: "", bank_ifsc: "", bank_branch: "",
  invoice_prefix: "NGR", default_hsn: "25062020", default_gst_rate: 5, default_unit: "MTS",
  gst_frequency: "monthly", einvoice: false, tax_audit: true, pf_esi: false, professional_tax: true, buyer_194q: false,
  signatory: "Authorised Signatory", invoice_terms: "", track_from: "",
};
export async function loadSettings() {
  const row = await run(supabase.from("company_settings").select("data").eq("id", 1).maybeSingle());
  return { ...DEFAULT_SETTINGS, ...(row?.data || {}) };
}
export async function saveSettings(data) {
  await run(supabase.from("company_settings").upsert({ id: 1, data, updated_at: new Date().toISOString() }));
}

// ── Customers ─────────────────────────────────────────────────────────────────
export const loadCustomers = () => run(supabase.from("customers").select("*").order("name"));
export async function saveCustomer(c) {
  const row = {
    name: c.name.trim(), gstin: (c.gstin || "").trim().toUpperCase() || null, pan: (c.pan || "").trim().toUpperCase() || null,
    contact: c.contact || null, phone: c.phone || null, email: c.email || null, address: c.address || null,
    ship_address: c.ship_address || null, state: c.state || null, state_code: c.state_code || null,
    credit_days: c.credit_days === "" || c.credit_days == null ? 30 : +c.credit_days, deducts_tds: !!c.deducts_tds, notes: c.notes || null,
  };
  return c.id
    ? run(supabase.from("customers").update(row).eq("id", c.id).select().single())
    : run(supabase.from("customers").insert(row).select().single());
}
export const deleteCustomer = id => run(supabase.from("customers").delete().eq("id", id));

// ── Suppliers (full tax details) ──────────────────────────────────────────────
export const loadSuppliersFull = () => run(supabase.from("suppliers").select("*").order("name"));
export async function saveSupplierFull(s) {
  const row = {
    name: s.name.trim(), contact: s.contact || null, phone: s.phone || null, address: s.address || null, email: s.email || null,
    gstin: (s.gstin || "").trim().toUpperCase() || null, pan: (s.pan || "").trim().toUpperCase() || null,
    state_code: s.state_code || null, entity_type: s.entity_type || "company", tds_section: s.tds_section || null,
    lower_tds_cert: s.lower_tds_cert || null, lower_tds_rate: s.lower_tds_rate === "" || s.lower_tds_rate == null ? null : +s.lower_tds_rate,
    msme_udyam: s.msme_udyam || null,
  };
  return s.id
    ? run(supabase.from("suppliers").update(row).eq("id", s.id).select().single())
    : run(supabase.from("suppliers").insert(row).select().single());
}

// ── Sales invoices ────────────────────────────────────────────────────────────
export const loadInvoices = () => run(supabase.from("sales_invoices").select("*").order("invoice_date", { ascending: false }).order("id", { ascending: false }));
export const loadReceipts = () => run(supabase.from("invoice_receipts").select("*").order("receipt_date", { ascending: false }));
export async function nextDocNo(docType, date) {
  return run(supabase.rpc("next_sales_doc_no", { p_doc_type: docType, p_date: date }));
}
export async function saveInvoice(inv) {
  const { id, ...row } = inv;
  return id
    ? run(supabase.from("sales_invoices").update(row).eq("id", id).select().single())
    : run(supabase.from("sales_invoices").insert(row).select().single());
}
export const cancelInvoice = (id, reason) =>
  run(supabase.from("sales_invoices").update({ status: "cancelled", cancel_reason: reason || null }).eq("id", id).select().single());
export const saveReceipt = r => run(supabase.from("invoice_receipts").insert(r).select().single());
export const deleteReceipt = id => run(supabase.from("invoice_receipts").delete().eq("id", id));

// ── ITC register ──────────────────────────────────────────────────────────────
export const loadInward = () => run(supabase.from("gst_inward").select("*").order("bill_date", { ascending: false }));
export async function saveInward(r) {
  const { id, ...row } = r;
  return id
    ? run(supabase.from("gst_inward").update(row).eq("id", id).select().single())
    : run(supabase.from("gst_inward").insert(row).select().single());
}
export const insertInwardMany = rows => run(supabase.from("gst_inward").upsert(rows, { onConflict: "source,source_id", ignoreDuplicates: true }).select());
export const deleteInward = id => run(supabase.from("gst_inward").delete().eq("id", id));

// ── TDS ───────────────────────────────────────────────────────────────────────
export const loadTdsSections = () => run(supabase.from("tds_sections").select("*").order("sort"));
export const saveTdsSection = s => run(supabase.from("tds_sections").upsert(s).select().single());
export const loadTds = () => run(supabase.from("tds_entries").select("*").order("deducted_on", { ascending: false }).order("id", { ascending: false }));
export async function saveTds(t) {
  const { id, ...row } = t;
  return id
    ? run(supabase.from("tds_entries").update(row).eq("id", id).select().single())
    : run(supabase.from("tds_entries").insert(row).select().single());
}
export const markTdsDeposited = (ids, ch) => run(supabase.from("tds_entries").update(ch).in("id", ids).select());
export const deleteTds = id => run(supabase.from("tds_entries").delete().eq("id", id));

// ── Compliance calendar ───────────────────────────────────────────────────────
export const loadFilings = () => run(supabase.from("compliance_filings").select("*"));
export const saveFiling = f => run(supabase.from("compliance_filings").upsert({ ...f, updated_at: new Date().toISOString() }).select().single());
export const deleteFiling = key => run(supabase.from("compliance_filings").delete().eq("key", key));
export const loadTasks = () => run(supabase.from("compliance_tasks").select("*").order("title"));
export async function saveTask(t) {
  const { id, ...row } = t;
  return id
    ? run(supabase.from("compliance_tasks").update(row).eq("id", id).select().single())
    : run(supabase.from("compliance_tasks").insert(row).select().single());
}
export const deleteTask = id => run(supabase.from("compliance_tasks").delete().eq("id", id));
