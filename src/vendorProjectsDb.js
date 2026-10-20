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

const mapEstimate = r => ({
  id: r.id, projectId: r.project_id, description: r.description || '',
  category: r.category || 'other', amount: +r.amount || 0,
  status: r.status || 'considering', treatment: r.treatment || 'additional',
  expectedDate: r.expected_date || '', notes: r.notes || '',
  createdAt: r.created_at || '', updatedAt: r.updated_at || '',
});

async function one(query) {
  const { data, error } = await query;
  if (error) throw error;
  return data;
}

// Never throws: if the tables haven't been created yet, returns ready:false so
// the rest of the app keeps working and the tab can show a setup hint.
export async function loadVendorProjects() {
  const [p, pay, est] = await Promise.all([
    supabase.from('vendor_projects').select('*').order('id', { ascending: false }),
    supabase.from('vendor_project_payments').select('*').order('pay_date', { ascending: false }).order('id', { ascending: false }),
    supabase.from('vendor_project_estimates').select('*').order('id', { ascending: false }),
  ]);
  const error = p.error || pay.error;
  if (error) return { projects: [], payments: [], estimates: [], ready: false, estimatesReady: !est.error, error: error.message || String(error), estimatesError: est.error?.message || '' };
  return { projects: (p.data || []).map(mapProject), payments: (pay.data || []).map(mapPayment), estimates: (est.data || []).map(mapEstimate), ready: true, estimatesReady: !est.error, error: '', estimatesError: est.error?.message || '' };
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

// Edit a manual / loan payment in place (amount, date, reference, remarks).
export async function updateProjectPayment(id, data) {
  const row = {
    amount: +data.amount, pay_date: data.date || null,
    reference: data.reference || null, remarks: data.remarks || null,
  };
  return mapPayment(await one(supabase.from('vendor_project_payments').update(row).eq('id', id).select().single()));
}


// Estimates are planning records only; they never create accounting ledger entries.
export async function saveProjectEstimate(data, id) {
  const row = {
    project_id: data.projectId,
    description: (data.description || '').trim(),
    category: data.category || 'other',
    amount: +data.amount || 0,
    status: data.status || 'considering',
    treatment: data.treatment || 'additional',
    expected_date: data.expectedDate || null,
    notes: (data.notes || '').trim() || null,
    updated_at: new Date().toISOString(),
  };
  const res = id
    ? await one(supabase.from('vendor_project_estimates').update(row).eq('id', id).select().single())
    : await one(supabase.from('vendor_project_estimates').insert(row).select().single());
  return mapEstimate(res);
}

export async function deleteProjectEstimate(id) {
  const { error } = await supabase.from('vendor_project_estimates').delete().eq('id', id);
  if (error) throw error;
}
