// Pure tax logic (no React, no database) so it can be tested on its own: node tools/test-tax.mjs
// GST invoice maths, GSTIN check, GSTR-1 / GSTR-3B working, TDS interest and the compliance calendar.

export const STATES = {
  "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana",
  "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
  "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
  "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
  "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep",
  "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman & Nicobar Islands", "36": "Telangana",
  "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
};
export const stateLabel = code => (code && STATES[code] ? `${STATES[code]} (${code})` : code || "");

const r2 = n => Math.round((+n || 0) * 100) / 100;

// ── GSTIN ─────────────────────────────────────────────────────────────────────
const GSTIN_RE = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;
const CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
export function gstinChecksumOk(g) {
  let sum = 0;
  for (let i = 0; i < 14; i++) {
    const v = CHARS.indexOf(g[i]) * (i % 2 ? 2 : 1);
    sum += Math.floor(v / 36) + (v % 36);
  }
  return CHARS[(36 - (sum % 36)) % 36] === g[14];
}
// Returns "" when fine, otherwise a short reason.
export function checkGSTIN(raw) {
  const g = String(raw || "").trim().toUpperCase();
  if (!g) return "";
  if (g.length !== 15) return "GSTIN must be 15 characters";
  if (!GSTIN_RE.test(g)) return "GSTIN format looks wrong";
  if (!STATES[g.slice(0, 2)]) return "Unknown state code in GSTIN";
  if (!gstinChecksumOk(g)) return "GSTIN check digit does not match (typo?)";
  return "";
}
export const panFromGSTIN = g => (g && g.length >= 12 ? g.slice(2, 12).toUpperCase() : "");
export const checkPAN = p => (!p ? "" : /^[A-Z]{5}[0-9]{4}[A-Z]$/.test(String(p).toUpperCase()) ? "" : "PAN format looks wrong");

// ── Invoice maths ─────────────────────────────────────────────────────────────
// items: [{ description, hsn, qty, unit, rate, discount, gst_rate }]
export function computeInvoice(items, sellerState, placeOfSupply) {
  const inter = String(sellerState) !== String(placeOfSupply);
  const lines = (items || []).map(it => {
    const taxable = r2((+it.qty || 0) * (+it.rate || 0) - (+it.discount || 0));
    const rate = +it.gst_rate || 0;
    const igst = inter ? r2(taxable * rate / 100) : 0;
    const cgst = inter ? 0 : r2(taxable * rate / 200);
    const sgst = cgst;
    return { ...it, taxable, igst, cgst, sgst, total: r2(taxable + igst + cgst + sgst) };
  });
  const sum = k => r2(lines.reduce((a, l) => a + l[k], 0));
  const taxable_total = sum("taxable"), igst = sum("igst"), cgst = sum("cgst"), sgst = sum("sgst");
  const exact = r2(taxable_total + igst + cgst + sgst);
  const grand_total = Math.round(exact);
  return { inter, lines, taxable_total, igst, cgst, sgst, round_off: r2(grand_total - exact), grand_total };
}

// Rate-wise tax breakup (for the invoice footer and HSN summary)
export function taxBreakup(lines) {
  const m = {};
  lines.forEach(l => {
    const k = `${l.hsn || ""}|${+l.gst_rate || 0}`;
    if (!m[k]) m[k] = { hsn: l.hsn || "", rate: +l.gst_rate || 0, unit: l.unit || "", qty: 0, taxable: 0, igst: 0, cgst: 0, sgst: 0 };
    m[k].qty += +l.qty || 0; m[k].taxable += l.taxable; m[k].igst += l.igst; m[k].cgst += l.cgst; m[k].sgst += l.sgst;
  });
  return Object.values(m).map(x => ({ ...x, qty: r2(x.qty), taxable: r2(x.taxable), igst: r2(x.igst), cgst: r2(x.cgst), sgst: r2(x.sgst) }));
}

// Unit Quantity Codes used on the GST portal
export const UQC = { MTS: "MTS-METRIC TON", KGS: "KGS-KILOGRAMS", BAG: "BAG-BAGS", NOS: "NOS-NUMBERS", TON: "TON-TONNES", QTL: "QTL-QUINTAL", OTH: "OTH-OTHERS" };

// ── Periods ───────────────────────────────────────────────────────────────────
export const monthKey = d => String(d || "").slice(0, 7);             // 'YYYY-MM'
export function fyOfDate(d) {
  const [y, m] = String(d).split("-").map(Number);
  const s = m >= 4 ? y : y - 1;
  return `${s}-${String((s + 1) % 100).padStart(2, "0")}`;
}
export function fyMonths(fy) {
  const y = +fy.slice(0, 4);
  return Array.from({ length: 12 }, (_, i) => {
    const mm = ((i + 3) % 12) + 1, yy = i < 9 ? y : y + 1;
    return `${yy}-${String(mm).padStart(2, "0")}`;
  });
}
export const monthLabel = mk => {
  const [y, m] = mk.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("en-IN", { month: "short", year: "numeric" });
};
const iso = (y, m, d) => {
  const last = new Date(y, m, 0).getDate();            // m is 1-12
  return `${y}-${String(m).padStart(2, "0")}-${String(Math.min(d, last)).padStart(2, "0")}`;
};
const nextMonth = mk => { const [y, m] = mk.split("-").map(Number); return m === 12 ? [y + 1, 1] : [y, m + 1]; };

// ── GSTR-1 working ────────────────────────────────────────────────────────────
// B2C (large) = inter-state invoice to an unregistered buyer above ₹1,00,000.
export const B2CL_LIMIT = 100000;
export function gstr1(invoices, month, sellerState) {
  const live = invoices.filter(i => monthKey(i.invoice_date) === month && i.status === "issued");
  const inv = live.filter(i => i.doc_type === "invoice");
  const notes = live.filter(i => i.doc_type !== "invoice");
  const b2b = inv.filter(i => i.bill_gstin);
  const b2cl = inv.filter(i => !i.bill_gstin && String(i.place_of_supply) !== String(sellerState) && +i.grand_total > B2CL_LIMIT);
  const b2csRows = inv.filter(i => !i.bill_gstin && !b2cl.includes(i));
  const b2cs = {};
  b2csRows.forEach(i => (i.items || []).forEach(l => {
    const k = `${i.place_of_supply}|${+l.gst_rate || 0}`;
    if (!b2cs[k]) b2cs[k] = { pos: i.place_of_supply, rate: +l.gst_rate || 0, taxable: 0, igst: 0, cgst: 0, sgst: 0 };
    b2cs[k].taxable += +l.taxable || 0; b2cs[k].igst += +l.igst || 0; b2cs[k].cgst += +l.cgst || 0; b2cs[k].sgst += +l.sgst || 0;
  }));
  // HSN summary, split B2B / B2C as the portal now asks
  const hsnOf = list => {
    const m = {};
    list.forEach(i => {
      const sign = i.doc_type === "credit_note" ? -1 : 1;
      (i.items || []).forEach(l => {
        const k = `${l.hsn}|${l.unit}|${+l.gst_rate || 0}`;
        if (!m[k]) m[k] = { hsn: l.hsn, unit: l.unit, rate: +l.gst_rate || 0, qty: 0, value: 0, taxable: 0, igst: 0, cgst: 0, sgst: 0 };
        m[k].qty += sign * (+l.qty || 0); m[k].taxable += sign * (+l.taxable || 0);
        m[k].igst += sign * (+l.igst || 0); m[k].cgst += sign * (+l.cgst || 0); m[k].sgst += sign * (+l.sgst || 0);
        m[k].value += sign * ((+l.taxable || 0) + (+l.igst || 0) + (+l.cgst || 0) + (+l.sgst || 0));
      });
    });
    return Object.values(m).map(x => Object.fromEntries(Object.entries(x).map(([k, v]) => [k, typeof v === "number" && !["rate"].includes(k) ? r2(v) : v])));
  };
  // Documents issued (table 13): includes cancelled numbers
  const allOfMonth = invoices.filter(i => monthKey(i.invoice_date) === month && i.status !== "draft");
  const docs = ["invoice", "credit_note", "debit_note"].map(t => {
    const xs = allOfMonth.filter(i => i.doc_type === t).sort((a, b) => a.invoice_no.localeCompare(b.invoice_no, undefined, { numeric: true }));
    return xs.length ? { type: t, from: xs[0].invoice_no, to: xs[xs.length - 1].invoice_no, total: xs.length, cancelled: xs.filter(i => i.status === "cancelled").length } : null;
  }).filter(Boolean);
  return {
    b2b, b2cl, b2cs: Object.values(b2cs).map(x => ({ ...x, taxable: r2(x.taxable), igst: r2(x.igst), cgst: r2(x.cgst), sgst: r2(x.sgst) })),
    cdnr: notes.filter(i => i.bill_gstin), cdnur: notes.filter(i => !i.bill_gstin),
    hsnB2B: hsnOf(live.filter(i => i.bill_gstin)), hsnB2C: hsnOf(live.filter(i => !i.bill_gstin)), docs,
  };
}

// ── GSTR-3B working ───────────────────────────────────────────────────────────
export function outwardTotals(invoices, month) {
  const t = { taxable: 0, igst: 0, cgst: 0, sgst: 0, count: 0 };
  invoices.filter(i => monthKey(i.invoice_date) === month && i.status === "issued").forEach(i => {
    const s = i.doc_type === "credit_note" ? -1 : 1;
    t.taxable += s * +i.taxable_total; t.igst += s * +i.igst; t.cgst += s * +i.cgst; t.sgst += s * +i.sgst; t.count++;
  });
  return { taxable: r2(t.taxable), igst: r2(t.igst), cgst: r2(t.cgst), sgst: r2(t.sgst), count: t.count };
}
export function itcTotals(inward, month) {
  const t = { taxable: 0, igst: 0, cgst: 0, sgst: 0, blocked: 0, count: 0 };
  inward.filter(r => (r.claim_month || monthKey(r.bill_date)) === month).forEach(r => {
    if (!r.itc_eligible || r.itc_category === "blocked") { t.blocked += (+r.igst || 0) + (+r.cgst || 0) + (+r.sgst || 0); return; }
    t.taxable += +r.taxable || 0; t.igst += +r.igst || 0; t.cgst += +r.cgst || 0; t.sgst += +r.sgst || 0; t.count++;
  });
  return Object.fromEntries(Object.entries(t).map(([k, v]) => [k, k === "count" ? v : r2(v)]));
}
// Set-off order (Sec 49 / Rule 88A): IGST credit → IGST, then CGST, then SGST.
// CGST credit → CGST, then IGST. SGST credit → SGST, then IGST. CGST and SGST never cross.
export function setOff(liab, credit) {
  const L = { igst: +liab.igst || 0, cgst: +liab.cgst || 0, sgst: +liab.sgst || 0 };
  const C = { igst: +credit.igst || 0, cgst: +credit.cgst || 0, sgst: +credit.sgst || 0 };
  const used = { igst: { igst: 0, cgst: 0, sgst: 0 }, cgst: { igst: 0, cgst: 0 }, sgst: { igst: 0, sgst: 0 } };
  const take = (from, to) => { const x = Math.min(C[from], L[to]); C[from] -= x; L[to] -= x; used[from][to] += x; };
  take("igst", "igst"); take("igst", "cgst"); take("igst", "sgst");
  take("cgst", "cgst"); take("cgst", "igst");
  take("sgst", "sgst"); take("sgst", "igst");
  const round = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [k, typeof v === "number" ? r2(v) : round(v)]));
  return { cash: round(L), carryForward: round(C), used: round(used) };
}

// ── TDS ───────────────────────────────────────────────────────────────────────
export function tdsRateFor(section, supplier) {
  if (!section) return 0;
  if (supplier?.lower_tds_rate != null && supplier.lower_tds_rate !== "") return +supplier.lower_tds_rate;
  if (supplier && !supplier.pan) return Math.max(20, ["individual", "huf"].includes(supplier.entity_type) ? +section.rate_ind : +section.rate_other);
  return ["individual", "huf"].includes(supplier?.entity_type) ? +section.rate_ind : +section.rate_other;
}
// Deposit due: 7th of next month; tax deducted in March → 30 April.
export function tdsDepositDue(deductedOn) {
  const [y, m] = String(deductedOn).split("-").map(Number);
  if (m === 3) return iso(y, 4, 30);
  const [ny, nm] = nextMonth(`${y}-${String(m).padStart(2, "0")}`);
  return iso(ny, nm, 7);
}
// Interest for late deposit: 1.5% per month or part of a month, from the date of deduction to the date of deposit.
export function tdsLateInterest(entry, depositDate) {
  const paid = depositDate || entry.challan_date;
  const due = tdsDepositDue(entry.deducted_on);
  const until = paid || isoToday();
  if (until <= due) return { months: 0, interest: 0, due };
  const [y1, m1] = entry.deducted_on.split("-").map(Number);
  const [y2, m2] = until.split("-").map(Number);
  const months = (y2 - y1) * 12 + (m2 - m1) + 1;
  return { months, interest: Math.ceil((+entry.tds_amount || 0) * 0.015 * months), due };
}
export const tdsQuarter = d => {
  const m = +String(d).slice(5, 7);
  return m >= 4 && m <= 6 ? "Q1" : m >= 7 && m <= 9 ? "Q2" : m >= 10 && m <= 12 ? "Q3" : "Q4";
};
export function isoToday() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// ── Compliance calendar ───────────────────────────────────────────────────────
// Returns every due item for the financial year (April–March), plus the run-in of next April/May
// for year-end items. Each item: { key, title, authority, due, period, category, help }
export function complianceItems(fy, s = {}, tasks = []) {
  const y = +fy.slice(0, 4);
  const out = [];
  const add = (key, title, authority, due, period, category, help) => out.push({ key, title, authority, due, period, category, help });
  const months = fyMonths(fy);
  const qrmp = s.gst_frequency === "qrmp";
  const ptDay = 10;

  months.forEach((mk, i) => {
    const [ny, nm] = nextMonth(mk);
    const lbl = monthLabel(mk);
    const quarterEnd = [2, 5, 8, 11].includes(i);
    // GST
    if (!qrmp) {
      add(`GSTR1:${mk}`, `GSTR-1 for ${lbl}`, "GST", iso(ny, nm, 11), mk, "GST", "Outward supplies. Use GST › GSTR-1 for the working and Excel.");
      add(`GSTR3B:${mk}`, `GSTR-3B for ${lbl} + pay tax`, "GST", iso(ny, nm, 20), mk, "GST", "Summary return and payment. Use GST › GSTR-3B for the working.");
    } else if (quarterEnd) {
      add(`GSTR1:${mk}`, `GSTR-1 (quarterly) up to ${lbl}`, "GST", iso(ny, nm, 13), mk, "GST", "QRMP quarterly GSTR-1.");
      add(`GSTR3B:${mk}`, `GSTR-3B (quarterly) up to ${lbl}`, "GST", iso(ny, nm, 22), mk, "GST", "QRMP: Telangana / Andhra Pradesh file by the 22nd.");
    } else {
      add(`IFF:${mk}`, `IFF (optional B2B upload) for ${lbl}`, "GST", iso(ny, nm, 13), mk, "GST", "Optional: lets B2B buyers see your invoices early.");
      add(`PMT06:${mk}`, `PMT-06 GST payment for ${lbl}`, "GST", iso(ny, nm, 25), mk, "GST", "QRMP monthly tax payment.");
    }
    // TDS deposit
    add(`TDSPAY:${mk}`, `Deposit TDS deducted in ${lbl}`, "Income Tax", +mk.slice(5) === 3 ? iso(ny, 4, 30) : iso(ny, nm, 7), mk, "TDS", "Challan for all TDS deducted in the month. Late deposit costs 1.5% per month.");
    if (s.pf_esi) {
      add(`PF:${mk}`, `PF contribution + ECR for ${lbl}`, "EPFO", iso(ny, nm, 15), mk, "Payroll", "");
      add(`ESI:${mk}`, `ESI contribution for ${lbl}`, "ESIC", iso(ny, nm, 15), mk, "Payroll", "");
    }
    if (s.professional_tax) add(`PT:${mk}`, `Professional tax for ${lbl}`, "Commercial Taxes (TS)", iso(ny, nm, ptDay), mk, "Payroll", "Telangana PT on salaries paid, by the 10th.");
  });

  // TDS quarterly returns (Form 140, old 26Q) and certificates (Form 131, old 16A)
  [["Q1", y, 7, 31, 8, 15], ["Q2", y, 10, 31, 11, 15], ["Q3", y + 1, 1, 31, 2, 15], ["Q4", y + 1, 5, 31, 6, 15]].forEach(([q, yy, m, d, cm, cd]) => {
    add(`TDSRET:${fy}:${q}`, `TDS return ${q} (Form 140, old 26Q)`, "Income Tax", iso(yy, m, d), `${fy} ${q}`, "TDS", "Quarterly statement of non-salary TDS. Export from TDS › Quarterly return.");
    add(`TDSCERT:${fy}:${q}`, `TDS certificates ${q} (Form 131, old 16A)`, "Income Tax", iso(yy, cm, cd), `${fy} ${q}`, "TDS", "Download from TRACES and send to each deductee.");
  });
  // Advance tax
  [[6, 15, "15%"], [9, 45, "45%"], [12, 75, "75%"], [3, 100, "100%"]].forEach(([m, pct, lbl]) => {
    const yy = m === 3 ? y + 1 : y;
    add(`ADVTAX:${fy}:${pct}`, `Advance tax: ${lbl} of the year's tax`, "Income Tax", iso(yy, m, 15), fy, "Income tax", "Cumulative instalment for the LLP.");
  });
  // Annual GST
  add(`GSTR9:${fy}`, `GSTR-9 annual return for ${fy}`, "GST", iso(y + 1, 12, 31), fy, "GST", "Annual return (optional below ₹2 Cr turnover).");
  add(`GSTR9C:${fy}`, `GSTR-9C reconciliation for ${fy} (if turnover > ₹5 Cr)`, "GST", iso(y + 1, 12, 31), fy, "GST", "Self-certified reconciliation statement.");
  // LLP / MCA and income tax for the year
  add(`FORM11:${fy}`, `LLP Form 11 annual return for ${fy}`, "MCA", iso(y + 1, 5, 30), fy, "LLP / MCA", "Partners and contributions as on 31 March. ₹100 per day late fee.");
  add(`FORM8:${fy}`, `LLP Form 8 accounts & solvency for ${fy}`, "MCA", iso(y + 1, 10, 30), fy, "LLP / MCA", "Statement of account and solvency, signed by designated partners. ₹100 per day late fee.");
  add(`DIRKYC:${fy}`, "DIR-3 KYC of designated partners (check if due this year)", "MCA", iso(y + 1, 9, 30), fy, "LLP / MCA", "");
  if (s.tax_audit) {
    add(`TAXAUDIT:${fy}`, `Tax audit report for ${fy}`, "Income Tax", iso(y + 1, 9, 30), fy, "Income tax", "Required when turnover crosses the audit limit.");
    add(`ITR:${fy}`, `Income tax return (ITR-5) for ${fy}, audit case`, "Income Tax", iso(y + 1, 10, 31), fy, "Income tax", "");
  } else {
    add(`ITR:${fy}`, `Income tax return (ITR-5) for ${fy}`, "Income Tax", iso(y + 1, 7, 31), fy, "Income tax", "");
  }
  // Your own items
  tasks.filter(t => t.active).forEach(t => {
    const base = `CUSTOM:${t.id}`;
    if (t.recurrence === "once" && t.once_date && fyOfDate(t.once_date) === fy) add(base, t.title, t.authority || "", t.once_date, "", "Custom", t.notes || "");
    if (t.recurrence === "monthly") months.forEach(mk => { const [yy, mm] = mk.split("-").map(Number); add(`${base}:${mk}`, `${t.title} (${monthLabel(mk)})`, t.authority || "", iso(yy, mm, t.due_day), mk, "Custom", t.notes || ""); });
    if (t.recurrence === "quarterly") [6, 9, 12, 3].forEach(m => { const yy = m === 3 ? y + 1 : y; add(`${base}:${yy}-${m}`, `${t.title} (quarter to ${monthLabel(`${yy}-${String(m).padStart(2, "0")}`)})`, t.authority || "", iso(yy, m, t.due_day), "", "Custom", t.notes || ""); });
    if (t.recurrence === "yearly" && t.due_month) { const yy = t.due_month >= 4 ? y : y + 1; add(`${base}:${fy}`, t.title, t.authority || "", iso(yy, t.due_month, t.due_day), fy, "Custom", t.notes || ""); }
  });
  return out.sort((a, b) => a.due.localeCompare(b.due));
}

// Status of a calendar item given the saved filings
export function itemStatus(item, filing, today = isoToday()) {
  if (filing?.done_on) return { code: "done", label: filing.done_on > item.due ? "Filed late" : "Filed", late: filing.done_on > item.due };
  const days = Math.round((new Date(item.due) - new Date(today)) / 864e5);
  if (days < 0) return { code: "overdue", label: `Overdue ${-days}d`, days };
  if (days <= 7) return { code: "soon", label: days === 0 ? "Due today" : `Due in ${days}d`, days };
  return { code: "upcoming", label: `In ${days}d`, days };
}
