// Run: node tools/test-tax.mjs
import assert from 'node:assert/strict'
import * as T from '../src/taxLogic.js'

let n = 0
const t = (name, fn) => { fn(); n++; console.log('ok -', name) }

t('own GSTIN passes', () => assert.equal(T.checkGSTIN('36AAYFN2455Q1Z8'), ''))
t('known sample GSTIN passes', () => assert.equal(T.checkGSTIN('27AAPFU0939F1ZV'), ''))
t('typo is caught', () => assert.match(T.checkGSTIN('36AAYFN2455Q1Z9'), /check digit/))
t('PAN from GSTIN', () => assert.equal(T.panFromGSTIN('36AAYFN2455Q1Z8'), 'AAYFN2455Q'))

t('intra-state invoice splits CGST/SGST', () => {
  const r = T.computeInvoice([{ qty: 30.5, rate: 4200, gst_rate: 5, hsn: '25062020', unit: 'MTS' }], '36', '36')
  assert.equal(r.taxable_total, 128100); assert.equal(r.cgst, 3202.5); assert.equal(r.sgst, 3202.5); assert.equal(r.igst, 0)
  assert.equal(r.grand_total, 134505); assert.equal(r.round_off, 0)
})
t('inter-state invoice uses IGST and rounds', () => {
  const r = T.computeInvoice([{ qty: 10.123, rate: 3999, gst_rate: 5 }], '36', '37')
  assert.equal(r.inter, true); assert.equal(r.igst, T.computeInvoice([{ qty: 10.123, rate: 3999, gst_rate: 5 }], '36', '37').igst)
  assert.ok(Math.abs(r.round_off) <= 0.5); assert.equal(r.grand_total % 1, 0)
})

t('set-off follows Rule 88A order', () => {
  const r = T.setOff({ igst: 1000, cgst: 500, sgst: 500 }, { igst: 1200, cgst: 100, sgst: 600 })
  assert.deepEqual(r.cash, { igst: 0, cgst: 200, sgst: 0 })
  assert.equal(r.used.igst.cgst, 200); assert.equal(r.carryForward.sgst, 100)
})
t('CGST credit never pays SGST', () => {
  const r = T.setOff({ igst: 0, cgst: 0, sgst: 500 }, { igst: 0, cgst: 900, sgst: 0 })
  assert.equal(r.cash.sgst, 500); assert.equal(r.carryForward.cgst, 900)
})

t('TDS due 7th of next month, March → 30 April', () => {
  assert.equal(T.tdsDepositDue('2026-10-25'), '2026-11-07')
  assert.equal(T.tdsDepositDue('2026-12-31'), '2027-01-07')
  assert.equal(T.tdsDepositDue('2027-03-15'), '2027-04-30')
})
t('TDS late interest counts part months', () => {
  const r = T.tdsLateInterest({ deducted_on: '2026-10-25', tds_amount: 10000 }, '2026-11-08')
  assert.equal(r.months, 2); assert.equal(r.interest, 300)
  assert.equal(T.tdsLateInterest({ deducted_on: '2026-10-25', tds_amount: 10000 }, '2026-11-07').interest, 0)
})
t('TDS rate: no PAN → 20%, individual 1%', () => {
  const s = { rate_ind: 1, rate_other: 2 }
  assert.equal(T.tdsRateFor(s, { entity_type: 'individual', pan: 'ABCDE1234F' }), 1)
  assert.equal(T.tdsRateFor(s, { entity_type: 'company', pan: 'ABCDE1234F' }), 2)
  assert.equal(T.tdsRateFor(s, { entity_type: 'company', pan: '' }), 20)
})

t('calendar has GST, TDS and LLP items', () => {
  const items = T.complianceItems('2026-27', { gst_frequency: 'monthly', tax_audit: true, professional_tax: true }, [
    { id: 1, title: 'Weighbridge stamping', recurrence: 'yearly', due_day: 15, due_month: 2, active: true },
  ])
  const by = k => items.find(i => i.key === k)
  assert.equal(by('GSTR3B:2026-10').due, '2026-11-20')
  assert.equal(by('GSTR1:2027-03').due, '2027-04-11')
  assert.equal(by('TDSPAY:2027-03').due, '2027-04-30')
  assert.equal(by('TDSRET:2026-27:Q4').due, '2027-05-31')
  assert.equal(by('FORM8:2026-27').due, '2027-10-30')
  assert.equal(by('FORM11:2026-27').due, '2027-05-30')
  assert.equal(by('ITR:2026-27').due, '2027-10-31')
  assert.equal(by('CUSTOM:1:2026-27').due, '2027-02-15')
  assert.ok(items.every((x, i) => !i || items[i - 1].due <= x.due))
})
t('QRMP uses 22nd for Telangana', () => {
  const items = T.complianceItems('2026-27', { gst_frequency: 'qrmp' })
  assert.equal(items.find(i => i.key === 'GSTR3B:2026-06').due, '2026-07-22')
  assert.ok(items.find(i => i.key === 'PMT06:2026-04'))
})

t('GSTR-1 buckets', () => {
  const inv = (o) => ({ status: 'issued', doc_type: 'invoice', invoice_date: '2026-10-05', items: [{ hsn: '2506', unit: 'MTS', qty: 1, taxable: 100, igst: 5, cgst: 0, sgst: 0, gst_rate: 5 }], grand_total: 105, ...o })
  const r = T.gstr1([
    inv({ invoice_no: 'A/1', bill_gstin: '27AAPFU0939F1ZV', place_of_supply: '27' }),
    inv({ invoice_no: 'A/2', bill_gstin: '', place_of_supply: '27', grand_total: 150000 }),
    inv({ invoice_no: 'A/3', bill_gstin: '', place_of_supply: '36' }),
    inv({ invoice_no: 'A/4', status: 'cancelled' }),
  ], '2026-10', '36')
  assert.equal(r.b2b.length, 1); assert.equal(r.b2cl.length, 1); assert.equal(r.b2cs.length, 1)
  assert.equal(r.docs[0].total, 4); assert.equal(r.docs[0].cancelled, 1)
})

t('amount months label', () => assert.equal(T.fyMonths('2026-27')[0], '2026-04'))
console.log(`\n${n} tests passed`)
