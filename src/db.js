// src/db.js — Complete Supabase Query & Automation Layer
import { supabase } from './supabase.js';

// ── Generic helpers ────────────────────────────────────────────────────────────
async function q(fn) {
  const { data, error } = await fn;
  if (error) throw error;
  return data || [];
}

async function q1(fn) {
  const { data, error } = await fn;
  if (error) throw error;
  return data;
}

// ── Load ALL data on app mount ────────────────────────────────────────────────
export async function loadAll() {
  const [
    suppliers, purchases, expenses, capitalItems,
    bankBatches, bankTxns, bankLabels, partnerCashbook,
  ] = await Promise.all([
    q(supabase.from('suppliers').select('*').order('name')),
    q(supabase.from('purchases').select('*, suppliers(name)').order('p_date', { ascending: false })),
    q(supabase.from('expenses').select('*').order('e_date', { ascending: false })),
    q(supabase.from('capital_items').select('*').order('c_date', { ascending: false })),
    q(supabase.from('bank_import_batches').select('*').order('imported_at', { ascending: false })),
    q(supabase.from('bank_txns').select('*').order('txn_date', { ascending: false })),
    q(supabase.from('bank_labels').select('*')),
    q(supabase.from('partner_cashbook').select('*').order('entry_date', { ascending: false })),
  ]);

  const appSuppliers = suppliers.map(s => ({
    id: s.id, name: s.name, contact: s.contact, phone: s.phone, address: s.address,
  }));

  const appPurchases = purchases.map(p => ({
    id: p.id, poNo: p.po_no, date: p.p_date,
    supplierId: p.supplier_id, supplierName: p.suppliers?.name || '',
    description: p.description || '', quantityMT: +p.quantity_mt,
    ratePerMT: +p.rate_per_mt, gstPct: +p.gst_pct,
    taxableAmount: +p.taxable_amount, gstAmount: +p.gst_amount,
    totalAmount: +p.total_amount,
    invoiceNo: p.invoice_no || '', remarks: p.remarks || '',
    bankTxnId: p.bank_txn_id,
  }));

  const appExpenses = expenses.map(e => ({
    id: e.id, expNo: e.exp_no, date: e.e_date,
    category: e.category, amount: +e.amount,
    description: e.description || '', paidTo: e.paid_to || '',
    reference: e.reference || '', opsAuto: e.ops_auto,
    paymentMode: e.payment_mode || 'bank',
    paidByPartner: e.paid_by_partner || '',
    bankTxnId: e.bank_txn_id,
  }));

  const appCapitalItems = capitalItems.map(c => ({
    id: c.id, capNo: c.cap_no, date: c.c_date,
    category: c.category, description: c.description,
    amount: +c.amount, paidTo: c.paid_to || '',
    reference: c.reference || '', fundedBy: c.funded_by,
    paymentMode: c.payment_mode || 'bank',
    paidByPartner: c.paid_by_partner || '',
    bankTxnId: c.bank_txn_id,
  }));

  const appBankBatches = bankBatches.map(b => ({
    id: b.id, source: b.source_filename, importedAt: b.imported_at, count: 0,
  }));
  bankTxns.forEach(t => {
    const batch = appBankBatches.find(b => b.id === t.batch_id);
    if (batch) batch.count++;
  });

  const appBankTxns = bankTxns.map(t => ({
    id: t.id, batchId: t.batch_id,
    date: t.txn_date, description: t.description,
    debit: +t.debit, credit: +t.credit,
    key: t.group_key,
    source: bankBatches.find(b => b.id === t.batch_id)?.source_filename || '',
    importedAt: bankBatches.find(b => b.id === t.batch_id)?.imported_at || '',
  }));

  const appBankLabels = {};
  bankLabels.forEach(l => { 
    appBankLabels[l.group_key] = { 
      label: l.label || '', 
      type: l.type || 'unlabeled',
      targetCategory: l.target_category || '',
      policy: l.policy || 'suggest'
    }; 
  });

  const appPartnerCashbook = partnerCashbook.map(p => ({
    id: p.id,
    partnerName: p.partner_name,
    entryDate: p.entry_date,
    accountType: p.account_type || 'current',
    type: p.type,
    amount: +p.amount,
    paymentMode: p.payment_mode || 'cash',
    bankTxnId: p.bank_txn_id,
    remarks: p.remarks || '',
  }));

  return {
    suppliers: appSuppliers,
    purchases: appPurchases, expenses: appExpenses,
    capitalItems: appCapitalItems,
    bankBatches: appBankBatches, bankTxns: appBankTxns,
    bankLabels: appBankLabels,
    partnerCashbook: appPartnerCashbook,
  };
}

// ── MASTERS ────────────────────────────────────────────────────────────────────
export async function saveSupplier(data, id) {
  const row = { name: data.name, contact: data.contact, phone: data.phone, address: data.address };
  const res = id ? await q1(supabase.from('suppliers').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('suppliers').insert(row).select().single());
  return { id: res.id, name: res.name, contact: res.contact, phone: res.phone, address: res.address };
}
export async function deleteSupplier(id) { await supabase.from('suppliers').delete().eq('id', id); }

// ── COMMERCIAL SUB-LEDGERS ───────────────────────────────────────────────────
export async function savePurchase(data) {
  const taxable = (+data.quantityMT) * (+data.ratePerMT);
  const gstAmt  = taxable * (+data.gstPct) / 100;
  const row = {
    p_date: data.date, supplier_id: +data.supplierId,
    description: data.description || null,
    quantity_mt: +data.quantityMT, rate_per_mt: +data.ratePerMT,
    gst_pct: +data.gstPct, taxable_amount: taxable,
    gst_amount: gstAmt, total_amount: taxable + gstAmt,
    invoice_no: data.invoiceNo || null, remarks: data.remarks || null,
    bank_txn_id: data.bankTxnId || null,
  };
  const p = await q1(supabase.from('purchases').insert(row).select('*, suppliers(name)').single());
  return {
    id: p.id, poNo: p.po_no, date: p.p_date,
    supplierId: p.supplier_id, supplierName: p.suppliers?.name || '',
    description: p.description || '', quantityMT: +p.quantity_mt,
    ratePerMT: +p.rate_per_mt, gstPct: +p.gst_pct,
    taxableAmount: +p.taxable_amount, gstAmount: +p.gst_amount,
    totalAmount: +p.total_amount, invoiceNo: p.invoice_no || '', remarks: p.remarks || '',
    bankTxnId: p.bank_txn_id,
  };
}
export async function deletePurchase(id) { await supabase.from('purchases').delete().eq('id', id); }

export async function saveExpense(data) {
  const row = {
    e_date: data.date, category: data.category, amount: +data.amount,
    description: data.description, paid_to: data.paidTo || null,
    reference: data.reference || null, ops_auto: !!data.opsAuto,
    exp_no: data.expNo || null,
    payment_mode: data.paymentMode || 'bank',
    paid_by_partner: data.paidByPartner || null,
    bank_txn_id: data.bankTxnId || null,
  };
  const e = await q1(supabase.from('expenses').insert(row).select().single());
  return {
    id: e.id, expNo: e.exp_no, date: e.e_date,
    category: e.category, amount: +e.amount,
    description: e.description || '', paidTo: e.paid_to || '',
    reference: e.reference || '', opsAuto: e.ops_auto,
    paymentMode: e.payment_mode, paidByPartner: e.paid_by_partner || '',
    bankTxnId: e.bank_txn_id,
  };
}
export async function deleteExpense(id) { await supabase.from('expenses').delete().eq('id', id); }

export async function saveMonthlyOpsCosts(ym, catAmounts) {
  await supabase.from('expenses')
    .delete()
    .eq('ops_auto', true)
    .gte('e_date', `${ym}-01`)
    .lte('e_date', `${ym}-31`);

  const rows = Object.entries(catAmounts)
    .filter(([, amt]) => +amt > 0)
    .map(([cat, amt]) => ({
      e_date: `${ym}-01`, category: cat, amount: +amt,
      description: `Monthly ${cat} total`,
      ops_auto: true,
      exp_no: `EXP-OPS-${ym}-${cat}`,
      payment_mode: 'bank',
    }));

  if (!rows.length) return [];
  const res = await q(supabase.from('expenses').insert(rows).select());
  return res.map(e => ({
    id: e.id, expNo: e.exp_no, date: e.e_date,
    category: e.category, amount: +e.amount,
    description: e.description || '', paidTo: e.paid_to || '',
    reference: e.reference || '', opsAuto: e.ops_auto,
    paymentMode: e.payment_mode, paidByPartner: '',
  }));
}

export async function saveCapitalItem(data) {
  const row = {
    c_date: data.date, category: data.category,
    description: data.description, amount: +data.amount,
    paid_to: data.paidTo || null, reference: data.reference || null,
    payment_mode: data.paymentMode || 'bank',
    paid_by_partner: data.paidByPartner || null,
    funded_by: data.fundedBy || 'own',
    bank_txn_id: data.bankTxnId || null,
  };
  const c = await q1(supabase.from('capital_items').insert(row).select().single());
  return {
    id: c.id, capNo: c.cap_no, date: c.c_date,
    category: c.category, description: c.description,
    amount: +c.amount, paidTo: c.paid_to || '',
    reference: c.reference || '', fundedBy: c.funded_by,
    paymentMode: c.payment_mode, paidByPartner: c.paid_by_partner || '',
    bankTxnId: c.bank_txn_id,
  };
}
export async function deleteCapitalItem(id) { await supabase.from('capital_items').delete().eq('id', id); }

// ── PARTNER CASHBOOK & CURRENT ACCOUNT ─────────────────────────────────────────
export async function savePartnerCashbookEntry(data) {
  const row = {
    partner_name: data.partnerName,
    entry_date: data.entryDate,
    account_type: data.accountType || 'current',
    type: data.type,
    amount: +data.amount,
    payment_mode: data.paymentMode || 'cash',
    bank_txn_id: data.bankTxnId || null,
    remarks: data.remarks || null,
  };
  const p = await q1(supabase.from('partner_cashbook').insert(row).select().single());
  return {
    id: p.id,
    partnerName: p.partner_name,
    entryDate: p.entry_date,
    accountType: p.account_type,
    type: p.type,
    amount: +p.amount,
    paymentMode: p.payment_mode,
    bankTxnId: p.bank_txn_id,
    remarks: p.remarks || '',
  };
}

// Atomically cross-posts an off-statement partner direct cash expenditure
export async function recordPartnerDirectCashExpense(partnerName, date, category, amount, description, ref, isCapex = false, vendorName = '') {
  // The vendor is who actually got the cash. Falling back to ref/partner only
  // when no vendor was given keeps old call sites working.
  const paidTo = (vendorName || '').trim() || ref || partnerName;
  let createdItem = null;
  if (isCapex) {
    createdItem = await saveCapitalItem({
      date, category, description, amount: +amount, paidTo,
      reference: ref, paymentMode: 'partner_personal', paidByPartner: partnerName, fundedBy: 'own'
    });
  } else {
    createdItem = await saveExpense({
      date, category, description, amount: +amount, paidTo,
      reference: ref, opsAuto: false, paymentMode: 'partner_personal', paidByPartner: partnerName
    });
  }

  // Credit the partner's current account for paying on the company's behalf
  const cb = await savePartnerCashbookEntry({
    partnerName, entryDate: date, accountType: 'current',
    type: 'Direct Cash Expense', amount: +amount, paymentMode: 'cash',
    remarks: `${isCapex ? '[Capex]' : '[Opex]'} ${description} (${category})`
  });

  return { item: createdItem, cashbookEntry: cb, isCapex };
}

// Re-point an existing cash expense / capex entry at the right vendor
// (older partner-cash entries were saved under the partner's name).
export async function setPaidTo(table, id, vendorName) {
  if (table !== 'expenses' && table !== 'capital_items') throw new Error('Unsupported table');
  const { error } = await supabase.from(table).update({ paid_to: (vendorName || '').trim() || null }).eq('id', id);
  if (error) throw error;
}

// ── BANK STATEMENT & PROMOTIONS ───────────────────────────────────────────────
export async function importBankBatch(filename, txns) {
  const batch = await q1(supabase
    .from('bank_import_batches').insert({ source_filename: filename }).select().single());

  const uniqueKeys = [...new Set(txns.map(t => t.key))];
  if (uniqueKeys.length) {
    await supabase.from('bank_labels')
      .upsert(uniqueKeys.map(k => ({ group_key: k })), { onConflict: 'group_key', ignoreDuplicates: true });
  }

  const rows = txns.map(t => ({
    batch_id: batch.id, txn_date: t.date,
    description: t.description, ref_no: t.refNo || null,
    debit: +t.debit, credit: +t.credit, group_key: t.key,
  }));

  const inserted = [];
  for (let i = 0; i < rows.length; i += 200) {
    const chunk = await q(supabase.from('bank_txns')
      .upsert(rows.slice(i, i + 200), { onConflict: 'txn_date,description,debit,credit', ignoreDuplicates: true }).select());
    inserted.push(...chunk);
  }

  return {
    batch: { id: batch.id, source: batch.source_filename, importedAt: batch.imported_at, count: inserted.length },
    txns: inserted.map(t => ({
      id: t.id, batchId: t.batch_id, date: t.txn_date,
      description: t.description, debit: +t.debit, credit: +t.credit,
      key: t.group_key, source: filename, importedAt: batch.imported_at,
    }))
  };
}

export async function deleteBankBatch(batchId) {
  await supabase.from('bank_txns').delete().eq('batch_id', batchId);
  await supabase.from('bank_import_batches').delete().eq('id', batchId);
}

export async function updateBankLabel(groupKey, patch) {
  const row = {
    group_key: groupKey,
    ...(patch.label !== undefined ? { label: patch.label } : {}),
    ...(patch.type !== undefined ? { type: patch.type } : {}),
    ...(patch.targetCategory !== undefined ? { target_category: patch.targetCategory } : {}),
    ...(patch.policy !== undefined ? { policy: patch.policy } : {}),
    updated_at: new Date().toISOString()
  };
  const { error } = await supabase
    .from('bank_labels')
    .upsert(row, { onConflict: 'group_key' });

  if (error) throw error;
}

export async function updateBankLabels(groupKeys, patch) {
  await Promise.all(groupKeys.map(k => updateBankLabel(k, patch)));
}

export async function promoteTxnToCapital(txn, category = "machinery", fundedBy = "own") {
  return saveCapitalItem({
    date: txn.date || txn.txn_date,
    category,
    description: txn.description,
    amount: +txn.debit,
    paidTo: txn.key || null,
    reference: txn.refNo || txn.reference || null,
    fundedBy,
    paymentMode: 'bank',
    bankTxnId: txn.id,
  });
}

export async function promoteTxnToExpense(txn, category = "misc") {
  return saveExpense({
    date: txn.date || txn.txn_date,
    category,
    amount: +txn.debit,
    description: txn.description,
    paidTo: txn.key || null,
    reference: txn.refNo || txn.reference || null,
    opsAuto: false,
    paymentMode: 'bank',
    bankTxnId: txn.id,
  });
}

export async function logPartnerCashbookEntry(txn, partnerName, type, accountType = "capital") {
  const amount = type === "Capital Infusion" ? +txn.credit : +txn.debit;
  return savePartnerCashbookEntry({
    partnerName,
    entryDate: txn.date || txn.txn_date,
    accountType,
    type,
    amount,
    paymentMode: 'bank',
    bankTxnId: txn.id,
    remarks: txn.description,
  });
}
