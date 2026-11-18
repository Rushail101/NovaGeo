// GST returns working: GSTR-1, GSTR-3B (with set-off), and the input tax credit (ITC) register with GSTR-2B matching.
import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import * as cdb from "./complianceDb.js";
import { useStore, upsertIn, removeIn, reload } from "./finStore.js";
import { useAuth, canWrite } from "./auth.jsx";
import { STATES, stateLabel, gstr1, outwardTotals, itcTotals, setOff, monthKey, monthLabel, fyOfDate, fyMonths, checkGSTIN, UQC } from "./taxLogic.js";
import { fmt, fmtDate, today, notify, Badge, StatCard, FG, Modal, EmptyState, SearchBar, SetupNotice, exportSheetsToExcel, exportRowsToExcel } from "./ui.jsx";

const SETUP = "Run supabase/002_compliance.sql in the Supabase SQL editor (after 001_security.sql), then reload.";
const prevMonth = () => { const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; };
const normNo = s => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "").replace(/^0+/, "");

function MonthPicker({ value, onChange }) {
  const fy = fyOfDate(today());
  const prevFy = `${+fy.slice(0, 4) - 1}-${fy.slice(2, 4)}`;
  const ms = [...fyMonths(prevFy), ...fyMonths(fy)].filter(m => m <= monthKey(today())).reverse();
  return <select className="inline-select" value={value} onChange={e => onChange(e.target.value)}>{ms.map(m => <option key={m} value={m}>{monthLabel(m)}</option>)}</select>;
}

function FiledButton({ fkey, title, due, filings, role }) {
  const f = filings.find(x => x.key === fkey);
  const [open, setOpen] = useState(false);
  if (f?.done_on) return <span className="chip" style={{ color: "var(--green)" }}>✓ Filed {fmtDate(f.done_on)}{f.ack_no ? ` · ARN ${f.ack_no}` : ""}</span>;
  if (!canWrite(role)) return <span className="chip">Not marked filed</span>;
  return <>
    <button className="btn btn-ghost btn-sm" onClick={() => setOpen(true)}>Mark as filed</button>
    {open && <FilingModal fkey={fkey} title={title} due={due} onClose={() => setOpen(false)} />}
  </>;
}

export function FilingModal({ fkey, title, due, existing, onClose }) {
  const { name } = useAuth();
  const [f, setF] = useState({ done_on: existing?.done_on || today(), ack_no: existing?.ack_no || "", amount_paid: existing?.amount_paid ?? "", notes: existing?.notes || "" });
  async function save() {
    try {
      const row = await cdb.saveFiling({ key: fkey, title, due_date: due || null, done_on: f.done_on, ack_no: f.ack_no || null, amount_paid: f.amount_paid === "" ? null : +f.amount_paid, notes: f.notes || null, filed_by: name });
      upsertIn("filings", row, "key"); notify("Marked as filed", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); }
  }
  async function undo() {
    if (!confirm("Remove the filed mark?")) return;
    try { await cdb.deleteFiling(fkey); removeIn("filings", fkey, "key"); onClose(); } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={title} onClose={onClose} foot={<>{existing && <button className="btn btn-danger" style={{ marginRight: "auto" }} onClick={undo}>Undo</button>}<button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
      {due && <div className="mono" style={{ color: "var(--text3)", marginBottom: 12 }}>Due {fmtDate(due)}</div>}
      <div className="form-row cols-2">
        <FG label="Filed / paid on"><input type="date" value={f.done_on} onChange={e => setF(x => ({ ...x, done_on: e.target.value }))} /></FG>
        <FG label="ARN / acknowledgement / challan no."><input value={f.ack_no} onChange={e => setF(x => ({ ...x, ack_no: e.target.value }))} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Amount paid (optional)"><input type="number" step="0.01" value={f.amount_paid} onChange={e => setF(x => ({ ...x, amount_paid: e.target.value }))} /></FG>
        <FG label="Notes"><input value={f.notes} onChange={e => setF(x => ({ ...x, notes: e.target.value }))} /></FG>
      </div>
      {f.done_on > (due || "9999") && <div className="callout callout-amber">Filed after the due date. Late fee / interest may apply.</div>}
    </Modal>
  );
}

export function GstView({ purchases = [] }) {
  const st = useStore(["settings", "invoices", "inward", "filings", "suppliers"]);
  const [tab, setTab] = useState("3b");
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;
  return (
    <>
      <div className="pill-tabs">
        {[["3b", "GSTR-3B working"], ["1", "GSTR-1 working"], ["itc", "ITC register (purchases)"]].map(([k, l]) =>
          <button key={k} className={`pill-tab${tab === k ? " active" : ""}`} onClick={() => setTab(k)}>{l}</button>)}
      </div>
      {tab === "1" && <Gstr1 st={st} />}
      {tab === "3b" && <Gstr3b st={st} />}
      {tab === "itc" && <ItcRegister st={st} purchases={purchases} />}
    </>
  );
}

// ── GSTR-1 ───────────────────────────────────────────────────────────────────
function Gstr1({ st }) {
  const { role } = useAuth();
  const [month, setMonth] = useState(prevMonth());
  const r = useMemo(() => gstr1(st.invoices, month, st.settings.state_code), [st.invoices, month, st.settings.state_code]);
  const [ny, nm] = month.split("-").map(Number);
  const due = `${nm === 12 ? ny + 1 : ny}-${String(nm === 12 ? 1 : nm + 1).padStart(2, "0")}-${st.settings.gst_frequency === "qrmp" ? "13" : "11"}`;
  const invRows = list => list.map(i => ({
    "GSTIN/UIN of Recipient": i.bill_gstin || "", "Receiver Name": i.bill_name, "Invoice Number": i.invoice_no, "Invoice date": fmtDate(i.invoice_date),
    "Invoice Value": +i.grand_total, "Place Of Supply": `${i.place_of_supply}-${STATES[i.place_of_supply] || ""}`, "Reverse Charge": i.reverse_charge ? "Y" : "N",
    "Invoice Type": "Regular B2B", "Taxable Value": +i.taxable_total, "Integrated Tax": +i.igst, "Central Tax": +i.cgst, "State/UT Tax": +i.sgst,
  }));
  const hsnRows = list => list.map(h => ({ HSN: h.hsn, UQC: UQC[h.unit] || h.unit, "Total Quantity": h.qty, "Total Value": h.value, Rate: h.rate, "Taxable Value": h.taxable, "Integrated Tax": h.igst, "Central Tax": h.cgst, "State/UT Tax": h.sgst }));
  function exportAll() {
    exportSheetsToExcel(`GSTR1-${month}`, [
      { name: "b2b", rows: invRows(r.b2b) },
      { name: "b2cl", rows: invRows(r.b2cl) },
      { name: "b2cs", rows: r.b2cs.map(x => ({ Type: "OE", "Place Of Supply": `${x.pos}-${STATES[x.pos] || ""}`, Rate: x.rate, "Taxable Value": x.taxable, "Integrated Tax": x.igst, "Central Tax": x.cgst, "State/UT Tax": x.sgst })) },
      { name: "cdnr", rows: r.cdnr.map(i => ({ "GSTIN/UIN of Recipient": i.bill_gstin, "Receiver Name": i.bill_name, "Note Number": i.invoice_no, "Note Date": fmtDate(i.invoice_date), "Note Type": i.doc_type === "credit_note" ? "C" : "D", "Place Of Supply": `${i.place_of_supply}-${STATES[i.place_of_supply] || ""}`, "Note Value": +i.grand_total, "Taxable Value": +i.taxable_total, "Integrated Tax": +i.igst, "Central Tax": +i.cgst, "State/UT Tax": +i.sgst, "Original invoice": i.against_invoice })) },
      { name: "cdnur", rows: r.cdnur.map(i => ({ "Note Number": i.invoice_no, "Note Date": fmtDate(i.invoice_date), "Note Type": i.doc_type === "credit_note" ? "C" : "D", "Place Of Supply": `${i.place_of_supply}-${STATES[i.place_of_supply] || ""}`, "Note Value": +i.grand_total, "Taxable Value": +i.taxable_total, "Integrated Tax": +i.igst })) },
      { name: "hsn(b2b)", rows: hsnRows(r.hsnB2B) },
      { name: "hsn(b2c)", rows: hsnRows(r.hsnB2C) },
      { name: "docs", rows: r.docs.map(d => ({ "Nature of Document": { invoice: "Invoices for outward supply", credit_note: "Credit Note", debit_note: "Debit Note" }[d.type], "Sr. No. From": d.from, "Sr. No. To": d.to, "Total Number": d.total, Cancelled: d.cancelled })) },
    ]);
  }
  const sum = (list, k) => list.reduce((a, x) => a + (+x[k] || 0), 0);
  const Section = ({ title, children, count }) => <div className="table-wrap"><div className="table-toolbar"><h3>{title} <span className="mono" style={{ color: "var(--text3)" }}>· {count}</span></h3></div>{children}</div>;
  return (
    <>
      <div className="filter-bar">
        <MonthPicker value={month} onChange={setMonth} />
        <button className="btn btn-primary btn-sm" onClick={exportAll}>Download GSTR-1 Excel</button>
        <FiledButton fkey={`GSTR1:${month}`} title={`GSTR-1 for ${monthLabel(month)}`} due={due} filings={st.filings} role={role} />
        <span className="mono" style={{ color: "var(--text3)" }}>Due {fmtDate(due)}</span>
      </div>
      <div className="callout callout-blue">Built from the Sales Invoices register. The Excel has one sheet per GSTR-1 table (b2b, b2cl, b2cs, cdnr, hsn split B2B / B2C, docs) for copying into the offline tool or checking against what your CA uploads.</div>
      <div className="grid-auto">
        <StatCard label="B2B invoices" value={r.b2b.length} sub={fmt(sum(r.b2b, "taxable_total"))} color="blue" />
        <StatCard label="B2C large / small" value={`${r.b2cl.length} / ${r.b2cs.length} rows`} sub={fmt(sum(r.b2cl, "taxable_total") + sum(r.b2cs, "taxable"))} color="purple" />
        <StatCard label="Credit / debit notes" value={r.cdnr.length + r.cdnur.length} color="amber" />
        <StatCard label="Documents issued" value={r.docs.reduce((a, d) => a + d.total, 0)} sub={`${r.docs.reduce((a, d) => a + d.cancelled, 0)} cancelled`} color="teal" />
      </div>
      <Section title="B2B (4A)" count={r.b2b.length}>
        {!r.b2b.length ? <EmptyState icon="—" message="No B2B invoices this month" /> : <div className="table-scroll"><table>
          <thead><tr><th>Invoice</th><th>Date</th><th>Recipient</th><th>GSTIN</th><th>POS</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th className="num">Value</th></tr></thead>
          <tbody>{r.b2b.map(i => <tr key={i.id}><td className="mono">{i.invoice_no}</td><td className="mono">{fmtDate(i.invoice_date)}</td><td>{i.bill_name}</td><td className="mono">{i.bill_gstin}</td><td>{i.place_of_supply}</td><td className="num">{fmt(i.taxable_total)}</td><td className="num">{fmt(i.igst)}</td><td className="num">{fmt(i.cgst)}</td><td className="num">{fmt(i.sgst)}</td><td className="num">{fmt(i.grand_total)}</td></tr>)}</tbody>
        </table></div>}
      </Section>
      {(r.b2cl.length > 0 || r.b2cs.length > 0) && <Section title="B2C (5 large, 7 small)" count={r.b2cl.length + r.b2cs.length}>
        <div className="table-scroll"><table><thead><tr><th>Type</th><th>POS</th><th className="num">Rate</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th></tr></thead>
          <tbody>
            {r.b2cl.map(i => <tr key={i.id}><td>B2CL {i.invoice_no}</td><td>{stateLabel(i.place_of_supply)}</td><td className="num">—</td><td className="num">{fmt(i.taxable_total)}</td><td className="num">{fmt(i.igst)}</td><td className="num">—</td><td className="num">—</td></tr>)}
            {r.b2cs.map((x, k) => <tr key={k}><td>B2CS</td><td>{stateLabel(x.pos)}</td><td className="num">{x.rate}%</td><td className="num">{fmt(x.taxable)}</td><td className="num">{fmt(x.igst)}</td><td className="num">{fmt(x.cgst)}</td><td className="num">{fmt(x.sgst)}</td></tr>)}
          </tbody></table></div>
      </Section>}
      <Section title="HSN summary (12)" count={r.hsnB2B.length + r.hsnB2C.length}>
        {!(r.hsnB2B.length + r.hsnB2C.length) ? <EmptyState icon="—" message="No supplies" /> : <div className="table-scroll"><table>
          <thead><tr><th>Tab</th><th>HSN</th><th>UQC</th><th className="num">Qty</th><th className="num">Rate</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th></tr></thead>
          <tbody>{[...r.hsnB2B.map(h => ({ ...h, tab: "B2B" })), ...r.hsnB2C.map(h => ({ ...h, tab: "B2C" }))].map((h, k) => <tr key={k}><td>{h.tab}</td><td className="mono">{h.hsn}</td><td className="mono">{UQC[h.unit] || h.unit}</td><td className="num">{h.qty}</td><td className="num">{h.rate}%</td><td className="num">{fmt(h.taxable)}</td><td className="num">{fmt(h.igst)}</td><td className="num">{fmt(h.cgst)}</td><td className="num">{fmt(h.sgst)}</td></tr>)}</tbody>
        </table></div>}
      </Section>
      <Section title="Documents issued (13)" count={r.docs.length}>
        {!r.docs.length ? <EmptyState icon="—" message="No documents" /> : <table><thead><tr><th>Nature</th><th>From</th><th>To</th><th className="num">Total</th><th className="num">Cancelled</th></tr></thead>
          <tbody>{r.docs.map(d => <tr key={d.type}><td>{d.type.replace("_", " ")}</td><td className="mono">{d.from}</td><td className="mono">{d.to}</td><td className="num">{d.total}</td><td className="num">{d.cancelled}</td></tr>)}</tbody></table>}
      </Section>
    </>
  );
}

// ── GSTR-3B ──────────────────────────────────────────────────────────────────
function Gstr3b({ st }) {
  const { role } = useAuth();
  const [month, setMonth] = useState(prevMonth());
  const [carry, setCarry] = useState({ igst: "", cgst: "", sgst: "" });
  const out = outwardTotals(st.invoices, month);
  const itc = itcTotals(st.inward, month);
  const credit = { igst: itc.igst + (+carry.igst || 0), cgst: itc.cgst + (+carry.cgst || 0), sgst: itc.sgst + (+carry.sgst || 0) };
  const so = setOff(out, credit);
  const [ny, nm] = month.split("-").map(Number);
  const due = `${nm === 12 ? ny + 1 : ny}-${String(nm === 12 ? 1 : nm + 1).padStart(2, "0")}-${st.settings.gst_frequency === "qrmp" ? "22" : "20"}`;
  const cash = so.cash.igst + so.cash.cgst + so.cash.sgst;
  const notIn2b = st.inward.filter(r => (r.claim_month || monthKey(r.bill_date)) === month && r.itc_eligible && r.in_2b === false);
  const unchecked = st.inward.filter(r => (r.claim_month || monthKey(r.bill_date)) === month && r.itc_eligible && r.in_2b == null);

  function exportIt() {
    exportRowsToExcel(`GSTR3B-working-${month}`, [
      { Table: "3.1(a) Outward taxable supplies", Taxable: out.taxable, IGST: out.igst, CGST: out.cgst, SGST: out.sgst },
      { Table: "4(A)(5) All other ITC", Taxable: itc.taxable, IGST: itc.igst, CGST: itc.cgst, SGST: itc.sgst },
      { Table: "ITC carried forward (entered)", IGST: +carry.igst || 0, CGST: +carry.cgst || 0, SGST: +carry.sgst || 0 },
      { Table: "Paid through ITC: IGST credit", IGST: so.used.igst.igst, CGST: so.used.igst.cgst, SGST: so.used.igst.sgst },
      { Table: "Paid through ITC: CGST credit", IGST: so.used.cgst.igst, CGST: so.used.cgst.cgst },
      { Table: "Paid through ITC: SGST credit", IGST: so.used.sgst.igst, SGST: so.used.sgst.sgst },
      { Table: "Tax payable in cash", IGST: so.cash.igst, CGST: so.cash.cgst, SGST: so.cash.sgst },
      { Table: "ITC left after set-off", IGST: so.carryForward.igst, CGST: so.carryForward.cgst, SGST: so.carryForward.sgst },
    ], "3B");
  }

  const Row = ({ label, v, cls }) => <tr className={cls}><td className="lbl">{label}</td><td className="num">{v.taxable != null ? fmt(v.taxable) : ""}</td><td className="num">{fmt(v.igst)}</td><td className="num">{fmt(v.cgst)}</td><td className="num">{fmt(v.sgst)}</td></tr>;
  return (
    <>
      <div className="filter-bar">
        <MonthPicker value={month} onChange={setMonth} />
        <button className="btn btn-ghost btn-sm" onClick={exportIt}>Export working</button>
        <FiledButton fkey={`GSTR3B:${month}`} title={`GSTR-3B for ${monthLabel(month)}`} due={due} filings={st.filings} role={role} />
        <span className="mono" style={{ color: "var(--text3)" }}>Due {fmtDate(due)}</span>
      </div>
      <div className="stats-grid">
        <StatCard label="Output tax" value={fmt(out.igst + out.cgst + out.sgst)} sub={`${out.count} documents`} color="blue" />
        <StatCard label="ITC available" value={fmt(itc.igst + itc.cgst + itc.sgst)} sub={`${itc.count} bills${itc.blocked ? ` · ${fmt(itc.blocked)} blocked` : ""}`} color="green" />
        <StatCard label="Cash payable" value={fmt(cash)} color={cash > 0 ? "amber" : "green"} />
        <StatCard label="ITC carried forward" value={fmt(so.carryForward.igst + so.carryForward.cgst + so.carryForward.sgst)} color="teal" />
      </div>
      {(notIn2b.length > 0 || unchecked.length > 0) && <div className="callout callout-amber">
        {notIn2b.length > 0 && <div><b>{notIn2b.length} bill(s) not found in GSTR-2B</b> but counted as ITC this month ({fmt(notIn2b.reduce((a, r) => a + +r.igst + +r.cgst + +r.sgst, 0))}). Credit can only be taken when the supplier has reported the bill; move them to a later month or follow up with the supplier.</div>}
        {unchecked.length > 0 && <div>{unchecked.length} bill(s) not yet checked against GSTR-2B. Import the 2B Excel under <b>ITC register</b>.</div>}
      </div>}
      <div className="table-wrap"><div className="table-toolbar"><h3>Working for {monthLabel(month)}</h3></div>
        <div className="table-scroll"><table className="ret-table">
          <thead><tr><th>Item</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th></tr></thead>
          <tbody>
            <Row label="3.1(a) Outward taxable supplies (net of credit notes)" v={out} />
            <Row label="4(A)(5) ITC: all other (from ITC register)" v={itc} />
            <tr><td className="lbl">ITC brought forward from last month (enter from the portal's credit ledger)</td><td></td>
              {["igst", "cgst", "sgst"].map(k => <td key={k} className="num"><input className="inline-select" style={{ width: 110, textAlign: "right" }} type="number" value={carry[k]} onChange={e => setCarry(c => ({ ...c, [k]: e.target.value }))} placeholder="0" /></td>)}</tr>
            <Row cls="sub" label="Credit used: IGST credit" v={{ igst: so.used.igst.igst, cgst: so.used.igst.cgst, sgst: so.used.igst.sgst }} />
            <Row label="Credit used: CGST credit" v={{ igst: so.used.cgst.igst, cgst: so.used.cgst.cgst, sgst: 0 }} />
            <Row label="Credit used: SGST credit" v={{ igst: so.used.sgst.igst, cgst: 0, sgst: so.used.sgst.sgst }} />
            <Row cls="grand" label="6.1 Tax payable in cash" v={so.cash} />
            <Row label="Credit left to carry forward" v={so.carryForward} />
          </tbody>
        </table></div>
      </div>
      <div className="callout">Set-off follows the GST rules: IGST credit is used for IGST first, then CGST and SGST; CGST credit for CGST then IGST; SGST credit for SGST then IGST. CGST and SGST credit never cross. Always check the system-generated 3B on the portal before you file.</div>
    </>
  );
}

// ── ITC register ─────────────────────────────────────────────────────────────
function ItcRegister({ st, purchases }) {
  const { role } = useAuth();
  const [month, setMonth] = useState(prevMonth());
  const [q, setQ] = useState("");
  const [edit, setEdit] = useState(null);
  const [match, setMatch] = useState(null);
  const rows = st.inward.filter(r => (r.claim_month || monthKey(r.bill_date)) === month && (!q || `${r.supplier_name} ${r.bill_no} ${r.supplier_gstin}`.toLowerCase().includes(q.toLowerCase())));
  const tot = itcTotals(st.inward, month);

  async function importPurchases() {
    const have = new Set(st.inward.filter(r => r.source === "purchase").map(r => String(r.source_id)));
    const todo = purchases.filter(p => +p.gstAmount > 0 && !have.has(String(p.id)));
    if (!todo.length) return notify("All purchases with GST are already in the register", "info");
    if (!confirm(`Add ${todo.length} purchase bill(s) with GST to the ITC register?`)) return;
    const rows = todo.map(p => {
      const sup = st.suppliers.find(s => s.id === p.supplierId);
      const inter = sup?.state_code && sup.state_code !== st.settings.state_code;
      const g = +p.gstAmount || 0;
      return { bill_date: p.date, bill_no: p.invoiceNo || p.poNo || `PUR-${p.id}`, supplier_id: p.supplierId || null, supplier_name: p.supplierName || sup?.name || "Supplier",
        supplier_gstin: sup?.gstin || null, description: p.description || null, taxable: +p.taxableAmount || 0,
        igst: inter ? g : 0, cgst: inter ? 0 : Math.round(g * 50) / 100, sgst: inter ? 0 : Math.round(g * 50) / 100,
        itc_eligible: true, itc_category: "inputs", source: "purchase", source_id: String(p.id) };
    });
    try { await cdb.insertInwardMany(rows); await reload("inward"); notify(`${rows.length} bill(s) added`, "ok"); } catch (e) { notify(e.message, "err"); }
  }

  async function del(r) {
    if (!confirm(`Delete bill ${r.bill_no}?`)) return;
    try { await cdb.deleteInward(r.id); removeIn("inward", r.id); } catch (e) { notify(e.message, "err"); }
  }

  return (
    <>
      <div className="filter-bar">
        <MonthPicker value={month} onChange={setMonth} />
        <SearchBar value={q} onChange={setQ} placeholder="Supplier, bill no…" style={{ minWidth: 200 }} />
        {canWrite(role) && <>
          <button className="btn btn-primary btn-sm" onClick={() => setEdit({ bill_date: `${month}-01`, itc_eligible: true, itc_category: "inputs", igst: 0, cgst: 0, sgst: 0 })}>+ Bill</button>
          <button className="btn btn-ghost btn-sm" onClick={importPurchases}>Pull from Purchases</button>
          <label className="btn btn-ghost btn-sm" style={{ cursor: "pointer" }}>Match GSTR-2B Excel<input type="file" accept=".xlsx,.xls" hidden onChange={e => { const file = e.target.files[0]; e.target.value = ""; if (file) parse2B(file).then(setMatch).catch(err => notify(err.message, "err")); }} /></label>
        </>}
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel(`ITC-${month}`, rows.map(r => ({ Date: fmtDate(r.bill_date), "Bill no": r.bill_no, Supplier: r.supplier_name, GSTIN: r.supplier_gstin, Taxable: +r.taxable, IGST: +r.igst, CGST: +r.cgst, SGST: +r.sgst, Eligible: r.itc_eligible ? "Yes" : "No", Category: r.itc_category, "In 2B": r.in_2b == null ? "" : r.in_2b ? "Yes" : "No" })))}>Export</button>
      </div>
      <div className="grid-auto">
        <StatCard label="Eligible ITC" value={fmt(tot.igst + tot.cgst + tot.sgst)} sub={`IGST ${fmt(tot.igst)} · CGST ${fmt(tot.cgst)} · SGST ${fmt(tot.sgst)}`} color="green" />
        <StatCard label="Blocked / ineligible" value={fmt(tot.blocked)} color="red" />
        <StatCard label="Matched with 2B" value={`${rows.filter(r => r.in_2b).length} / ${rows.length}`} color="blue" />
      </div>
      <div className="table-wrap">
        {!rows.length ? <EmptyState icon="📥" message="No purchase bills for this month" sub="Add bills with GST, pull them from the Purchases register, or import the GSTR-2B Excel from the GST portal." /> : (
          <div className="table-scroll"><table>
            <thead><tr><th>Date</th><th>Bill no.</th><th>Supplier</th><th className="num">Taxable</th><th className="num">IGST</th><th className="num">CGST</th><th className="num">SGST</th><th>ITC</th><th>2B</th><th></th></tr></thead>
            <tbody>{rows.map(r => (
              <tr key={r.id}>
                <td className="mono">{fmtDate(r.bill_date)}</td><td className="mono">{r.bill_no}</td>
                <td>{r.supplier_name}<div className="mono" style={{ fontSize: 10.5, color: "var(--text3)" }}>{r.supplier_gstin || "no GSTIN"}</div></td>
                <td className="num">{fmt(r.taxable)}</td><td className="num">{fmt(r.igst)}</td><td className="num">{fmt(r.cgst)}</td><td className="num">{fmt(r.sgst)}</td>
                <td>{r.itc_eligible && r.itc_category !== "blocked" ? <Badge type="green">{r.itc_category.replace("_", " ")}</Badge> : <Badge type="red">blocked</Badge>}</td>
                <td>{r.in_2b == null ? <span className="chip">?</span> : r.in_2b ? <Badge type="green">✓</Badge> : <Badge type="red">missing</Badge>}</td>
                <td className="r">{canWrite(role) && <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}><button className="btn btn-ghost btn-sm" onClick={() => setEdit(r)}>Edit</button><button className="btn btn-danger btn-sm" onClick={() => del(r)}>✕</button></div>}</td>
              </tr>
            ))}</tbody>
          </table></div>
        )}
      </div>
      {edit && <InwardModal value={edit} st={st} onClose={() => setEdit(null)} />}
      {match && <Match2B data={match} inward={st.inward} onClose={() => setMatch(null)} />}
    </>
  );
}

function InwardModal({ value, st, onClose }) {
  const [f, setF] = useState({ ...value });
  const set = k => e => setF(x => ({ ...x, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  const inter = f.supplier_gstin && f.supplier_gstin.slice(0, 2) !== st.settings.state_code;
  function quickTax(rate) {
    const t = +f.taxable || 0, g = Math.round(t * rate) / 100;
    setF(x => ({ ...x, igst: inter ? g : 0, cgst: inter ? 0 : Math.round(g * 50) / 100, sgst: inter ? 0 : Math.round(g * 50) / 100 }));
  }
  async function save() {
    if (!f.bill_no || !f.supplier_name || !f.bill_date) return notify("Bill date, number and supplier are required", "err");
    const gErr = checkGSTIN(f.supplier_gstin); if (gErr) return notify(gErr, "err");
    try {
      const row = await cdb.saveInward({ ...f, taxable: +f.taxable || 0, igst: +f.igst || 0, cgst: +f.cgst || 0, sgst: +f.sgst || 0, supplier_gstin: (f.supplier_gstin || "").toUpperCase() || null, claim_month: f.claim_month || null, supplier_id: f.supplier_id ? +f.supplier_id : null });
      upsertIn("inward", row); notify("Saved", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={f.id ? "Edit purchase bill" : "Add purchase bill"} onClose={onClose} size="modal-lg" foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
      <div className="form-row cols-3">
        <FG label="Supplier"><select value={f.supplier_id || ""} onChange={e => { const s = st.suppliers.find(x => x.id === +e.target.value); setF(x => ({ ...x, supplier_id: e.target.value, supplier_name: s?.name || x.supplier_name, supplier_gstin: s?.gstin || x.supplier_gstin })); }}><option value="">Select or type →</option>{st.suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG>
        <FG label="Supplier name *"><input value={f.supplier_name || ""} onChange={set("supplier_name")} /></FG>
        <FG label="Supplier GSTIN" note={checkGSTIN(f.supplier_gstin) || (f.supplier_gstin ? (inter ? "Other state → IGST" : "Same state → CGST + SGST") : "")}><input maxLength={15} value={f.supplier_gstin || ""} onChange={e => setF(x => ({ ...x, supplier_gstin: e.target.value.toUpperCase() }))} /></FG>
      </div>
      <div className="form-row cols-3">
        <FG label="Bill date *"><input type="date" value={f.bill_date || ""} onChange={set("bill_date")} /></FG>
        <FG label="Bill no. *"><input value={f.bill_no || ""} onChange={set("bill_no")} /></FG>
        <FG label="Claim in month (if later)"><input type="month" value={f.claim_month || ""} onChange={set("claim_month")} /></FG>
      </div>
      <div className="form-row cols-4">
        <FG label="Taxable value" note={<span>{[5, 12, 18, 28].map(r => <button key={r} type="button" className="link-btn" style={{ marginRight: 8 }} onClick={() => quickTax(r)}>{r}%</button>)}</span>}><input type="number" step="0.01" value={f.taxable ?? ""} onChange={set("taxable")} /></FG>
        <FG label="IGST"><input type="number" step="0.01" value={f.igst ?? ""} onChange={set("igst")} /></FG>
        <FG label="CGST"><input type="number" step="0.01" value={f.cgst ?? ""} onChange={set("cgst")} /></FG>
        <FG label="SGST"><input type="number" step="0.01" value={f.sgst ?? ""} onChange={set("sgst")} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="ITC type"><select value={f.itc_category} onChange={set("itc_category")}><option value="inputs">Inputs (raw quartz, consumables)</option><option value="capital_goods">Capital goods (machinery)</option><option value="input_services">Input services (transport, repairs)</option><option value="blocked">Blocked (Sec 17(5): food, personal vehicles…)</option></select></FG>
        <FG label="Description"><input value={f.description || ""} onChange={set("description")} /></FG>
      </div>
      <label className="check"><input type="checkbox" checked={!!f.itc_eligible} onChange={set("itc_eligible")} /> ITC can be claimed on this bill</label>
    </Modal>
  );
}

async function parse2B(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheetName = wb.SheetNames.find(n => /^b2b$/i.test(n.trim())) || wb.SheetNames.find(n => /b2b/i.test(n));
  if (!sheetName) throw new Error("No B2B sheet found. Download GSTR-2B as Excel from the GST portal and pick that file.");
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: "" });
  const hi = grid.findIndex(r => r.some(c => /gstin of supplier/i.test(String(c))));
  if (hi < 0) throw new Error("Could not find the 'GSTIN of supplier' header in the B2B sheet.");
  const h1 = grid[hi].map(String), h2 = (grid[hi + 1] || []).map(String);
  const head = h1.map((h, i) => `${h} ${h2[i] || ""}`.toLowerCase());
  const col = re => head.findIndex(h => re.test(h));
  const c = { gstin: col(/gstin of supplier/), name: col(/trade|legal name/), no: col(/invoice number|invoice no/), date: col(/invoice date/), taxable: col(/taxable value/), igst: col(/integrated tax/), cgst: col(/central tax/), sgst: col(/state\/ut tax|state tax/), itc: col(/itc availability/) };
  if (c.gstin < 0 || c.no < 0) throw new Error("This B2B sheet is not in the expected GSTR-2B layout.");
  const num = v => +String(v).replace(/[,₹\s]/g, "") || 0;
  const asDate = v => v instanceof Date ? `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}` : (() => { const m = String(v).match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : ""; })();
  const rows = grid.slice(hi + 2).filter(r => /^[0-9]{2}[A-Z0-9]{13}$/.test(String(r[c.gstin]).trim())).map(r => ({
    gstin: String(r[c.gstin]).trim(), name: String(r[c.name] || ""), no: String(r[c.no]), date: asDate(r[c.date]),
    taxable: num(r[c.taxable]), igst: num(r[c.igst]), cgst: num(r[c.cgst]), sgst: num(r[c.sgst]), itc: c.itc >= 0 ? String(r[c.itc]) : "",
  }));
  if (!rows.length) throw new Error("No invoices found in the B2B sheet.");
  return { file: file.name, rows };
}

function Match2B({ data, inward, onClose }) {
  const key = (g, no) => `${g}|${normNo(no)}`;
  const mine = new Map(inward.map(r => [key(r.supplier_gstin, r.bill_no), r]));
  const theirs = new Set(data.rows.map(r => key(r.gstin, r.no)));
  const matched = data.rows.filter(r => mine.has(key(r.gstin, r.no)));
  const onlyIn2b = data.rows.filter(r => !mine.has(key(r.gstin, r.no)));
  const dates = data.rows.map(r => r.date).filter(Boolean).sort();
  const from = dates[0] || "", to = dates[dates.length - 1] || "";
  const onlyMine = inward.filter(r => r.supplier_gstin && r.bill_date >= from.slice(0, 7) + "-01" && r.bill_date <= to && !theirs.has(key(r.supplier_gstin, r.bill_no)));
  const diff = matched.filter(r => { const m = mine.get(key(r.gstin, r.no)); return Math.abs((+m.igst + +m.cgst + +m.sgst) - (r.igst + r.cgst + r.sgst)) > 1; });
  const [busy, setBusy] = useState(false);
  async function apply(addMissing) {
    setBusy(true);
    try {
      for (const r of matched) { const m = mine.get(key(r.gstin, r.no)); if (m.in_2b !== true) await cdb.saveInward({ id: m.id, in_2b: true }); }
      for (const m of onlyMine) if (m.in_2b !== false) await cdb.saveInward({ id: m.id, in_2b: false });
      if (addMissing && onlyIn2b.length) await cdb.insertInwardMany(onlyIn2b.map(r => ({ bill_date: r.date || today(), bill_no: r.no, supplier_name: r.name || r.gstin, supplier_gstin: r.gstin, taxable: r.taxable, igst: r.igst, cgst: r.cgst, sgst: r.sgst, itc_eligible: !/^no$/i.test(r.itc), itc_category: "inputs", in_2b: true, source: "gstr2b", source_id: `${r.gstin}|${normNo(r.no)}` })));
      await reload("inward"); notify("GSTR-2B match saved", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); } finally { setBusy(false); }
  }
  const List = ({ title, list, render, tone }) => list.length ? <div style={{ marginBottom: 14 }}><div className="section-title" style={{ marginTop: 0, color: tone }}>{title} · {list.length}</div><div style={{ maxHeight: 160, overflow: "auto", fontSize: 12 }}>{list.map(render)}</div></div> : null;
  return (
    <Modal title={`GSTR-2B match · ${data.file}`} onClose={onClose} size="modal-lg"
      foot={<><button className="btn btn-ghost" onClick={onClose}>Close</button><button className="btn btn-ghost" disabled={busy} onClick={() => apply(false)}>Save match only</button><button className="btn btn-primary" disabled={busy} onClick={() => apply(true)}>Save match + add {onlyIn2b.length} missing bills</button></>}>
      <div className="grid-auto">
        <StatCard label="In 2B" value={data.rows.length} color="blue" />
        <StatCard label="Matched" value={matched.length} color="green" />
        <StatCard label="Amount differs" value={diff.length} color={diff.length ? "amber" : "green"} />
        <StatCard label="In our books, not in 2B" value={onlyMine.length} color={onlyMine.length ? "red" : "green"} />
      </div>
      <List title="Not in 2B: supplier hasn't reported these yet (hold the ITC)" tone="var(--red)" list={onlyMine} render={r => <div key={r.id} className="mono">{fmtDate(r.bill_date)} · {r.bill_no} · {r.supplier_name} · {fmt(+r.igst + +r.cgst + +r.sgst)}</div>} />
      <List title="Tax amount differs from 2B" tone="var(--amber)" list={diff} render={r => { const m = mine.get(key(r.gstin, r.no)); return <div key={r.gstin + r.no} className="mono">{r.no} · {r.name}: ours {fmt(+m.igst + +m.cgst + +m.sgst)} vs 2B {fmt(r.igst + r.cgst + r.sgst)}</div>; }} />
      <List title="In 2B but not in our register" tone="var(--blue)" list={onlyIn2b} render={r => <div key={r.gstin + r.no} className="mono">{fmtDate(r.date)} · {r.no} · {r.name} · {fmt(r.igst + r.cgst + r.sgst)}</div>} />
    </Modal>
  );
}
