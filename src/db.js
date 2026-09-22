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
    vehicles, customers, grades, suppliers, equipment,
    maintTasks, weighments, boulderReceipts,
    lots, lotGradeOutputs,
    productionEntries, productionEntryOutputs,
    shiftLogs, shiftLogOutputs,
    purchases, expenses, capitalItems,
    bankBatches, bankTxns, bankLabels,
    costConfigRows,
    qcTests, samples,
    maintLogs,
    partnerCashbook,
  ] = await Promise.all([
    q(supabase.from('vehicles').select('*').order('id')),
    q(supabase.from('customers').select('*').order('name')),
    q(supabase.from('grades').select('*').order('id')),
    q(supabase.from('suppliers').select('*').order('name')),
    q(supabase.from('equipment').select('*').order('code')),
    q(supabase.from('maint_tasks').select('*, equipment(code,name)').order('id')),
    q(supabase.from('weighments').select('*, vehicles(vehicle_no,type), customers(name), grades(code,name)').order('w_date', { ascending: false }).order('id', { ascending: false })),
    q(supabase.from('boulder_receipts').select('*, suppliers(name), vehicles(vehicle_no)').order('r_date', { ascending: false })),
    q(supabase.from('lots').select('*').order('lot_date', { ascending: false })),
    q(supabase.from('lot_grade_outputs').select('*, grades(code,name)')),
    q(supabase.from('production_entries').select('*').order('p_date', { ascending: false })),
    q(supabase.from('production_entry_outputs').select('*, grades(code,name)')),
    q(supabase.from('shift_logs').select('*').order('s_date', { ascending: false })),
    q(supabase.from('shift_log_outputs').select('*, grades(code,name)')),
    q(supabase.from('purchases').select('*, suppliers(name)').order('p_date', { ascending: false })),
    q(supabase.from('expenses').select('*').order('e_date', { ascending: false })),
    q(supabase.from('capital_items').select('*').order('c_date', { ascending: false })),
    q(supabase.from('bank_import_batches').select('*').order('imported_at', { ascending: false })),
    q(supabase.from('bank_txns').select('*').order('txn_date', { ascending: false })),
    q(supabase.from('bank_labels').select('*')),
    q(supabase.from('cost_config').select('*').eq('id', 1)),
    q(supabase.from('qc_tests').select('*').order('t_date', { ascending: false })),
    q(supabase.from('samples').select('*').order('s_date', { ascending: false })),
    q(supabase.from('maint_logs').select('*, equipment(code,name)').order('log_date', { ascending: false })),
    q(supabase.from('partner_cashbook').select('*').order('entry_date', { ascending: false })),
  ]);

  const appVehicles = vehicles.map(v => ({
    id: v.id, vehicleNo: v.vehicle_no, type: v.type,
    tareWeight: +v.tare_weight_kg, owner: v.owner, phone: v.phone,
  }));

  const appCustomers = customers.map(c => ({
    id: c.id, name: c.name, gstin: c.gstin, contact: c.contact,
    phone: c.phone, address: c.address, state: c.state,
  }));

  const appGrades = grades.map(g => ({
    id: g.id, code: g.code, name: g.name, unit: g.unit,
    description: g.description, isBoulder: g.is_boulder,
  }));

  const appSuppliers = suppliers.map(s => ({
    id: s.id, name: s.name, contact: s.contact, phone: s.phone, address: s.address,
  }));

  const appEquipment = equipment.map(e => ({
    id: e.id, code: e.code, name: e.name, type: e.type,
    capacity: e.capacity, make: e.make,
    installDate: e.install_date, status: e.status, notes: e.notes,
  }));

  const appMaintTasks = maintTasks.map(t => ({
    id: t.id, equipmentId: t.equipment_id,
    code: t.equipment?.code || '', equipmentName: t.equipment?.name || '',
    taskName: t.task_name, category: t.category,
    intervalType: t.interval_type, intervalValue: t.interval_value,
    estimatedMins: t.estimated_mins,
    lastDone: t.last_done || '', nextDue: t.next_due || '',
    active: t.active, notes: t.notes || '',
  }));

  const appMaintLogs = maintLogs.map(l => ({
    id: l.id, logNo: l.log_no, equipmentId: l.equipment_id,
    equipmentCode: l.equipment?.code || '', equipmentName: l.equipment?.name || '',
    taskId: l.task_id, taskName: l.task_name,
    date: l.log_date, type: l.type, performedBy: l.performed_by,
    durationHours: +l.duration_hours, cost: +l.cost,
    partsUsed: l.parts_used || '', findings: l.findings || '',
    nextServiceDate: l.next_service_date || '',
    ballsAddedKg: l.balls_added_kg, magnetMaterialGrams: l.magnet_material_grams,
    liningWearPct: l.lining_wear_pct, meshCondition: l.mesh_condition,
    remarks: l.remarks || '',
  }));

  const appWeighments = weighments.map(w => ({
    id: w.id, slipNo: w.slip_no, date: w.w_date, time: w.w_time,
    vehicleId: w.vehicle_id, customerId: w.customer_id, gradeId: w.grade_id,
    lotNo: w.lot_no || '',
    grossWeight: +w.gross_weight_kg, tareWeight: +w.tare_weight_kg,
    netWeight: +w.net_weight_kg,
    purpose: w.purpose, remarks: w.remarks || '',
    vehicleNo: w.vehicles?.vehicle_no || '',
    customerName: w.customers?.name || '—',
    gradeCode: w.grades?.code || '', gradeName: w.grades?.name || '',
  }));

  const appBoulderReceipts = boulderReceipts.map(r => ({
    id: r.id, receiptNo: r.receipt_no, date: r.r_date,
    vehicleId: r.vehicle_id, supplierId: r.supplier_id,
    quantityMT: +r.quantity_mt,
    royaltyNo: r.royalty_no || '', challanNo: r.challan_no || '',
    remarks: r.remarks || '',
    supplierName: r.suppliers?.name || '',
    vehicleNo: r.vehicles?.vehicle_no || '—',
  }));

  const lgByLotId = {};
  lotGradeOutputs.forEach(o => {
    if (!lgByLotId[o.lot_id]) lgByLotId[o.lot_id] = [];
    lgByLotId[o.lot_id].push({
      gradeId: o.grade_id,
      gradeCode: o.grades?.code || '',
      gradeName: o.grades?.name || '',
      quantityMT: +o.quantity_mt,
    });
  });

  const appLots = lots.map(l => ({
    id: l.id, lotNo: l.lot_no, date: l.lot_date,
    shift: l.shift || '', operator: l.operator || '',
    totalOutputMT: +l.total_output_mt,
    boulderConsumedMT: +l.boulder_consumed_mt,
    yieldPct: l.yield_pct ? +l.yield_pct : null,
    status: l.status, remarks: l.remarks || '',
    gradeOutputs: lgByLotId[l.id] || [],
  }));

  const peoByEntryId = {};
  productionEntryOutputs.forEach(o => {
    if (!peoByEntryId[o.production_entry_id]) peoByEntryId[o.production_entry_id] = [];
    peoByEntryId[o.production_entry_id].push({
      gradeId: o.grade_id,
      gradeCode: o.grades?.code || '',
      gradeName: o.grades?.name || '',
      quantityMT: +o.quantity_mt,
    });
  });

  const appProductionEntries = productionEntries.map(pe => ({
    id: pe.id, entryNo: pe.entry_no, date: pe.p_date,
    shift: pe.shift, operator: pe.operator, lotNo: pe.lot_no,
    boulderConsumedMT: +pe.boulder_consumed_mt,
    machineHours: +pe.machine_hours, idleHours: +pe.idle_hours,
    totalOutputMT: +pe.total_output_mt,
    yieldPct: pe.yield_pct ? +pe.yield_pct : null,
    remarks: pe.remarks || '',
    outputs: peoByEntryId[pe.id] || [],
  }));

  const sloByLogId = {};
  shiftLogOutputs.forEach(o => {
    if (!sloByLogId[o.shift_log_id]) sloByLogId[o.shift_log_id] = [];
    sloByLogId[o.shift_log_id].push({
      gradeId: o.grade_id,
      gradeCode: o.grades?.code || '',
      gradeName: o.grades?.name || '',
      quantity: +o.quantity,
    });
  });

  const appShiftLogs = shiftLogs.map(l => ({
    id: l.id, logNo: l.log_no, date: l.s_date,
    shift: l.shift, operator: l.operator,
    machineStatus: l.machine_status, boulderFed: +l.boulder_fed_mt,
    breakdowns: l.breakdowns || '', remarks: l.remarks || '',
    production: sloByLogId[l.id] || [],
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

  const cfg = costConfigRows[0] || {};
  const appCostConfig = {
    bagCostPerTonne:  cfg.bag_cost_per_tonne  ? String(cfg.bag_cost_per_tonne)  : '',
    freightPerTonne:  cfg.freight_per_tonne   ? String(cfg.freight_per_tonne)   : '',
    loadingPerTonne:  cfg.loading_per_tonne   ? String(cfg.loading_per_tonne)   : '',
    monthlyLabCost:   cfg.monthly_lab_cost    ? String(cfg.monthly_lab_cost)    : '',
    monthlyRent:      cfg.monthly_rent        ? String(cfg.monthly_rent)        : '',
    miscFixed:        cfg.misc_fixed          ? String(cfg.misc_fixed)          : '',
    stdSellingRate:   cfg.std_selling_rate    ? String(cfg.std_selling_rate)    : '',
  };

  const appQcTests = qcTests.map(t => ({
    id: t.id, testNo: t.test_no, date: t.t_date,
    lotNo: t.lot_no, testedBy: t.tested_by || '',
    labType: t.lab_type, certNo: t.cert_no || '',
    result: t.result, remarks: t.remarks || '',
    params: {
      whiteness:    t.whiteness    != null ? String(t.whiteness)    : '',
      sio2:         t.sio2         != null ? String(t.sio2)         : '',
      fe2o3:        t.fe2o3        != null ? String(t.fe2o3)        : '',
      al2o3:        t.al2o3        != null ? String(t.al2o3)        : '',
      moisture:     t.moisture     != null ? String(t.moisture)     : '',
      loi:          t.loi          != null ? String(t.loi)          : '',
      d50:          t.d50          != null ? String(t.d50)          : '',
      bulkDensity: t.bulk_density!= null ? String(t.bulk_density): '',
    },
  }));

  const appSamples = samples.map(s => ({
    id: s.id, sampleNo: s.sample_no, date: s.s_date,
    lotNo: s.lot_no || '', location: s.location || '',
    collectedBy: s.collected_by || '', quantity: s.quantity || '',
    disposalDate: s.disposal_date, status: s.status,
    disposedDate: s.disposed_date || '',
    remarks: s.remarks || '',
  }));

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
    vehicles: appVehicles, customers: appCustomers,
    grades: appGrades, suppliers: appSuppliers,
    equipment: appEquipment, maintTasks: appMaintTasks,
    maintLogs: appMaintLogs, weighments: appWeighments,
    boulderReceipts: appBoulderReceipts, lots: appLots,
    productionEntries: appProductionEntries, shiftLogs: appShiftLogs,
    purchases: appPurchases, expenses: appExpenses,
    capitalItems: appCapitalItems,
    bankBatches: appBankBatches, bankTxns: appBankTxns,
    bankLabels: appBankLabels, costConfig: appCostConfig,
    qcTests: appQcTests, samples: appSamples,
    partnerCashbook: appPartnerCashbook,
  };
}

// ── MASTERS ────────────────────────────────────────────────────────────────────
export async function saveVehicle(data, id) {
  const row = { vehicle_no: data.vehicleNo, type: data.type, tare_weight_kg: +data.tareWeight, owner: data.owner, phone: data.phone };
  const res = id ? await q1(supabase.from('vehicles').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('vehicles').insert(row).select().single());
  return { id: res.id, vehicleNo: res.vehicle_no, type: res.type, tareWeight: +res.tare_weight_kg, owner: res.owner, phone: res.phone };
}
export async function deleteVehicle(id) { await supabase.from('vehicles').delete().eq('id', id); }

export async function saveCustomer(data, id) {
  const row = { name: data.name, gstin: data.gstin, contact: data.contact, phone: data.phone, address: data.address, state: data.state };
  const res = id ? await q1(supabase.from('customers').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('customers').insert(row).select().single());
  return { id: res.id, name: res.name, gstin: res.gstin, contact: res.contact, phone: res.phone, address: res.address, state: res.state };
}
export async function deleteCustomer(id) { await supabase.from('customers').delete().eq('id', id); }

export async function saveGrade(data, id) {
  const row = { code: data.code, name: data.name, unit: data.unit, description: data.description, is_boulder: !!data.isBoulder };
  const res = id ? await q1(supabase.from('grades').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('grades').insert(row).select().single());
  return { id: res.id, code: res.code, name: res.name, unit: res.unit, description: res.description, isBoulder: res.is_boulder };
}
export async function deleteGrade(id) { await supabase.from('grades').delete().eq('id', id); }

export async function saveSupplier(data, id) {
  const row = { name: data.name, contact: data.contact, phone: data.phone, address: data.address };
  const res = id ? await q1(supabase.from('suppliers').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('suppliers').insert(row).select().single());
  return { id: res.id, name: res.name, contact: res.contact, phone: res.phone, address: res.address };
}
export async function deleteSupplier(id) { await supabase.from('suppliers').delete().eq('id', id); }

export async function saveEquipment(data, id) {
  const row = { code: data.code, name: data.name, type: data.type, capacity: data.capacity, make: data.make, install_date: data.installDate || null, status: data.status, notes: data.notes };
  const res = id ? await q1(supabase.from('equipment').update(row).eq('id', id).select().single())
                 : await q1(supabase.from('equipment').insert(row).select().single());
  return { id: res.id, code: res.code, name: res.name, type: res.type, capacity: res.capacity, make: res.make, installDate: res.install_date, status: res.status, notes: res.notes };
}
export async function deleteEquipment(id) { await supabase.from('equipment').delete().eq('id', id); }

// ── MAINTENANCE ────────────────────────────────────────────────────────────────
export async function saveMaintTask(data, id) {
  const row = {
    equipment_id: data.equipmentId, task_name: data.taskName,
    category: data.category, interval_type: data.intervalType,
    interval_value: +data.intervalValue, estimated_mins: +data.estimatedMins,
    last_done: data.lastDone || null, next_due: data.nextDue || null,
    active: data.active, notes: data.notes,
  };
  const res = id ? await q1(supabase.from('maint_tasks').update(row).eq('id', id).select('*, equipment(code,name)').single())
                 : await q1(supabase.from('maint_tasks').insert(row).select('*, equipment(code,name)').single());
  return {
    id: res.id, equipmentId: res.equipment_id,
    code: res.equipment?.code || '', equipmentName: res.equipment?.name || '',
    taskName: res.task_name, category: res.category,
    intervalType: res.interval_type, intervalValue: res.interval_value,
    estimatedMins: res.estimated_mins, lastDone: res.last_done || '', nextDue: res.next_due || '',
    active: res.active, notes: res.notes || '',
  };
}
export async function deleteMaintTask(id) { await supabase.from('maint_tasks').delete().eq('id', id); }
export async function updateMaintTaskDue(id, lastDone, nextDue) {
  await supabase.from('maint_tasks').update({ last_done: lastDone, next_due: nextDue }).eq('id', id);
}

export async function saveMaintLog(data) {
  const row = {
    equipment_id: data.equipmentId, task_id: data.taskId || null,
    task_name: data.taskName, log_date: data.date,
    type: data.type, performed_by: data.performedBy,
    duration_hours: +data.durationHours || 0, cost: +data.cost || 0,
    parts_used: data.partsUsed || null, findings: data.findings || null,
    next_service_date: data.nextServiceDate || null,
    balls_added_kg: data.ballsAddedKg ? +data.ballsAddedKg : null,
    magnet_material_grams: data.magnetMaterialGrams ? +data.magnetMaterialGrams : null,
    lining_wear_pct: data.liningWearPct ? +data.liningWearPct : null,
    mesh_condition: data.meshCondition || null,
    remarks: data.remarks || null,
  };
  const res = await q1(supabase.from('maint_logs').insert(row).select('*, equipment(code,name)').single());
  return {
    id: res.id, logNo: res.log_no, equipmentId: res.equipment_id,
    equipmentCode: res.equipment?.code || '', equipmentName: res.equipment?.name || '',
    taskId: res.task_id, taskName: res.task_name,
    date: res.log_date, type: res.type, performedBy: res.performed_by,
    durationHours: +res.duration_hours, cost: +res.cost,
    partsUsed: res.parts_used || '', findings: res.findings || '',
    nextServiceDate: res.next_service_date || '',
    ballsAddedKg: res.balls_added_kg, magnetMaterialGrams: res.magnet_material_grams,
    liningWearPct: res.lining_wear_pct, meshCondition: res.mesh_condition,
    remarks: res.remarks || '',
  };
}
export async function deleteMaintLog(id) { await supabase.from('maint_logs').delete().eq('id', id); }

// ── OPERATIONS ─────────────────────────────────────────────────────────────────
export async function saveWeighment(data) {
  const row = {
    w_date: data.date, w_time: data.time || null,
    vehicle_id: data.vehicleId || null, customer_id: data.customerId || null,
    grade_id: data.gradeId || null, lot_no: data.lotNo || null,
    gross_weight_kg: +data.grossWeight, tare_weight_kg: +data.tareWeight,
    purpose: data.purpose, remarks: data.remarks || null,
  };
  const saved = await q1(supabase.from('weighments').insert(row).select(`
    *, vehicles(vehicle_no), customers(name), grades(code,name)
  `).single());
  return {
    id: saved.id, slipNo: saved.slip_no, date: saved.w_date, time: saved.w_time,
    vehicleId: saved.vehicle_id, customerId: saved.customer_id, gradeId: saved.grade_id,
    lotNo: saved.lot_no || '', grossWeight: +saved.gross_weight_kg,
    tareWeight: +saved.tare_weight_kg, netWeight: +saved.net_weight_kg,
    purpose: saved.purpose, remarks: saved.remarks || '',
    vehicleNo: saved.vehicles?.vehicle_no || '',
    customerName: saved.customers?.name || '—',
    gradeCode: saved.grades?.code || '', gradeName: saved.grades?.name || '',
  };
}

export async function saveBoulderReceipt(data) {
  const row = {
    r_date: data.date, vehicle_id: data.vehicleId || null,
    supplier_id: data.supplierId, quantity_mt: +data.quantityMT,
    royalty_no: data.royaltyNo || null, challan_no: data.challanNo || null,
    remarks: data.remarks || null,
  };
  const saved = await q1(supabase.from('boulder_receipts').insert(row)
    .select('*, suppliers(name), vehicles(vehicle_no)').single());
  return {
    id: saved.id, receiptNo: saved.receipt_no, date: saved.r_date,
    vehicleId: saved.vehicle_id, supplierId: saved.supplier_id,
    quantityMT: +saved.quantity_mt, royaltyNo: saved.royalty_no || '',
    challanNo: saved.challan_no || '', remarks: saved.remarks || '',
    supplierName: saved.suppliers?.name || '', vehicleNo: saved.vehicles?.vehicle_no || '—',
  };
}

export async function saveProductionEntry(data, outputs) {
  const lotRow = await q1(supabase.from('lots').insert({
    lot_date: data.date, shift: data.shift, operator: data.operator,
    total_output_mt: data.totalOutputMT, boulder_consumed_mt: +data.boulderConsumedMT,
    yield_pct: data.yieldPct || null, status: 'in-stock', remarks: data.remarks || null,
  }).select().single());

  if (outputs.length) {
    const lgRows = outputs.map(o => ({ lot_id: lotRow.id, grade_id: +o.gradeId, quantity_mt: +o.quantityMT }));
    await q(supabase.from('lot_grade_outputs').insert(lgRows));
  }

  const peRow = await q1(supabase.from('production_entries').insert({
    p_date: data.date, shift: data.shift, operator: data.operator,
    lot_no: lotRow.lot_no, boulder_consumed_mt: +data.boulderConsumedMT,
    machine_hours: +data.machineHours || 0, idle_hours: +data.idleHours || 0,
    total_output_mt: data.totalOutputMT, yield_pct: data.yieldPct || null,
    remarks: data.remarks || null,
  }).select().single());

  if (outputs.length) {
    const peoRows = outputs.map(o => ({ production_entry_id: peRow.id, grade_id: +o.gradeId, quantity_mt: +o.quantityMT }));
    await q(supabase.from('production_entry_outputs').insert(peoRows));
  }

  return {
    lot: {
      id: lotRow.id, lotNo: lotRow.lot_no, date: lotRow.lot_date,
      shift: lotRow.shift, operator: lotRow.operator,
      totalOutputMT: +lotRow.total_output_mt, boulderConsumedMT: +lotRow.boulder_consumed_mt,
      yieldPct: lotRow.yield_pct ? +lotRow.yield_pct : null, status: lotRow.status,
      remarks: lotRow.remarks || '', gradeOutputs: outputs,
    },
    entry: {
      id: peRow.id, entryNo: peRow.entry_no, date: peRow.p_date,
      shift: peRow.shift, operator: peRow.operator, lotNo: peRow.lot_no,
      boulderConsumedMT: +peRow.boulder_consumed_mt,
      machineHours: +peRow.machine_hours, idleHours: +peRow.idle_hours,
      totalOutputMT: +peRow.total_output_mt, yieldPct: peRow.yield_pct ? +peRow.yield_pct : null,
      remarks: peRow.remarks || '', outputs,
    }
  };
}

export async function updateLotStatus(id, status) {
  await supabase.from('lots').update({ status }).eq('id', id);
}

export async function saveShiftLog(data, outputs) {
  const slRow = await q1(supabase.from('shift_logs').insert({
    s_date: data.date, shift: data.shift, operator: data.operator,
    machine_status: data.machineStatus, boulder_fed_mt: +data.boulderFed,
    breakdowns: data.breakdowns || null, remarks: data.remarks || null,
  }).select().single());

  if (outputs.length) {
    const rows = outputs.map(o => ({ shift_log_id: slRow.id, grade_id: +o.gradeId, quantity: +o.quantity }));
    await q(supabase.from('shift_log_outputs').insert(rows));
  }
  return {
    id: slRow.id, logNo: slRow.log_no, date: slRow.s_date,
    shift: slRow.shift, operator: slRow.operator,
    machineStatus: slRow.machine_status, boulderFed: +slRow.boulder_fed_mt,
    breakdowns: slRow.breakdowns || '', remarks: slRow.remarks || '',
    production: outputs,
  };
}

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
export async function recordPartnerDirectCashExpense(partnerName, date, category, amount, description, ref, isCapex = false) {
  let createdItem = null;
  if (isCapex) {
    createdItem = await saveCapitalItem({
      date, category, description, amount: +amount, paidTo: ref || partnerName,
      reference: ref, paymentMode: 'partner_personal', paidByPartner: partnerName, fundedBy: 'own'
    });
  } else {
    createdItem = await saveExpense({
      date, category, description, amount: +amount, paidTo: ref || partnerName,
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

// ── COST CONFIG ───────────────────────────────────────────────────────────────
export async function saveCostConfig(data) {
  const row = {
    id: 1,
    bag_cost_per_tonne:  +data.bagCostPerTonne  || 0,
    freight_per_tonne:   +data.freightPerTonne   || 0,
    loading_per_tonne:   +data.loadingPerTonne   || 0,
    monthly_lab_cost:    +data.monthlyLabCost    || 0,
    monthly_rent:        +data.monthlyRent       || 0,
    misc_fixed:          +data.miscFixed         || 0,
    std_selling_rate:    +data.stdSellingRate    || 0,
  };
  await supabase.from('cost_config').upsert(row, { onConflict: 'id' });
}

// ── QC / LAB ──────────────────────────────────────────────────────────────────
export async function saveQcTest(data) {
  const p = data.params || {};
  const row = {
    t_date: data.date, lot_no: data.lotNo,
    tested_by: data.testedBy || null, lab_type: data.labType,
    cert_no: data.certNo || null, result: data.result,
    remarks: data.remarks || null,
    whiteness:    p.whiteness    ? +p.whiteness    : null,
    sio2:         p.sio2         ? +p.sio2         : null,
    fe2o3:        p.fe2o3        ? +p.fe2o3        : null,
    al2o3:        p.al2o3        ? +p.al2o3        : null,
    moisture:     p.moisture     ? +p.moisture     : null,
    loi:          p.loi          ? +p.loi          : null,
    d50:          p.d50          ? +p.d50          : null,
    bulk_density: p.bulkDensity  ? +p.bulkDensity  : null,
  };
  const t = await q1(supabase.from('qc_tests').insert(row).select().single());
  return {
    id: t.id, testNo: t.test_no, date: t.t_date,
    lotNo: t.lot_no, testedBy: t.tested_by || '',
    labType: t.lab_type, certNo: t.cert_no || '',
    result: t.result, remarks: t.remarks || '',
    params: data.params || {},
  };
}

export async function saveSample(data) {
  const row = {
    s_date: data.date, lot_no: data.lotNo || null,
    location: data.location, collected_by: data.collectedBy || null,
    quantity: data.quantity || null, disposal_date: data.disposalDate,
    status: 'retained', remarks: data.remarks || null,
  };
  const s = await q1(supabase.from('samples').insert(row).select().single());
  return {
    id: s.id, sampleNo: s.sample_no, date: s.s_date,
    lotNo: s.lot_no || '', location: s.location || '',
    collectedBy: s.collected_by || '', quantity: s.quantity || '',
    disposalDate: s.disposal_date, status: s.status,
    disposedDate: s.disposed_date || '', remarks: s.remarks || '',
  };
}

export async function markSampleDisposed(id) {
  const today = new Date().toISOString().split('T')[0];
  await supabase.from('samples').update({ status: 'disposed', disposed_date: today }).eq('id', id);
}
