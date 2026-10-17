// src/vendorProjectsDb.js — Supabase layer for Vendor Projects.
// Kept separate from db.js so it can be dropped into any version of the app.
import { supabase } from './supabase.js';

const mapProject = r => ({
  id: r.id, projectNo: r.project_no || '', vendorName: r.vendor_name || '',
  title: r.title || '', description: r.description || '',
  category: r.category || 'misc', status: r.status || 'active',
  contractValue: +r.contract_value || 0,
  startDate: r.start_date || '', endDate: r.end_date || '', notes: r.notes || '',
});

const mapPayment = r => ({
  id: r.id, projectId: r.project_id,
  bankTxnId: r.bank_txn_id ?? null,
  expenseId: r.expense_id ?? null,
  capitalItemId: r.capital_item_id ?? null,
  amount: +r.amount || 0, date: r.pay_date || '',
  paymentMode: r.payment_mode || '', reference: r.reference || '', remarks: r.remarks || '',
});

async function one(query) {
  const { data, error } = await query;
  if (error) throw error;
  return data;
}

// Never throws: if the tables haven't been created yet, returns ready:false so
// the rest of the app keeps working and the tab can show a setup hint.
export async function loadVendorProjects() {
  const [p, pay] = await Promise.all([
    supabase.from('vendor_projects').select('*').order('id', { ascending: false }),
    supabase.from('vendor_project_payments').select('*').order('pay_date', { ascending: false }).order('id', { ascending: false }),
  ]);
  const error = p.error || pay.error;
  if (error) return { projects: [], payments: [], ready: false, error: error.message || String(error) };
  return { projects: (p.data || []).map(mapProject), payments: (pay.data || []).map(mapPayment), ready: true, error: '' };
}

export async function saveVendorProject(data, id) {
  const row = {
    vendor_name: (data.vendorName || '').trim(),
    title: (data.title || '').trim(),
    description: data.description || null,
    category: data.category || 'misc',
    status: data.status || 'active',
    contract_value: +data.contractValue || 0,
    start_date: data.startDate || null,
    end_date: data.endDate || null,
    notes: data.notes || null,
  };
  const res = id
    ? await one(supabase.from('vendor_projects').update(row).eq('id', id).select().single())
    : await one(supabase.from('vendor_projects').insert(row).select().single());
  return mapProject(res);
}

export async function deleteVendorProject(id) {
  const { error } = await supabase.from('vendor_projects').delete().eq('id', id);
  if (error) throw error;
}

// Link (part of) a payment to a project. Exactly one of bankTxnId / expenseId /
// capitalItemId for a linked payment; none for a manual, project-only payment.
export async function addProjectPayment(data) {
  const row = {
    project_id: data.projectId,
    bank_txn_id: data.bankTxnId ?? null,
    expense_id: data.expenseId ?? null,
    capital_item_id: data.capitalItemId ?? null,
    amount: +data.amount,
    pay_date: data.date || null,
    payment_mode: data.paymentMode || null,
    reference: data.reference || null,
    remarks: data.remarks || null,
  };
  return mapPayment(await one(supabase.from('vendor_project_payments').insert(row).select().single()));
}

export async function deleteProjectPayment(id) {
  const { error } = await supabase.from('vendor_project_payments').delete().eq('id', id);
  if (error) throw error;
}
