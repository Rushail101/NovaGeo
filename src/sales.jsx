// Sales: customers, GST tax invoices / credit & debit notes, receipts and receivables ageing.
import React, { useMemo, useState } from "react";
import * as cdb from "./complianceDb.js";
import { useStore, reload, upsertIn, removeIn } from "./finStore.js";
import { useAuth, canWrite } from "./auth.jsx";
import { STATES, stateLabel, checkGSTIN, checkPAN, panFromGSTIN, computeInvoice, fyOfDate, monthKey, UQC } from "./taxLogic.js";
import { invoicePDF } from "./invoicePdf.js";
import { fmt, fmtDate, today, isoDate, notify, Badge, StatCard, FG, Modal, EmptyState, SearchBar, SetupNotice, exportRowsToExcel } from "./ui.jsx";

const SETUP = "Run supabase/002_compliance.sql in the Supabase SQL editor (after 001_security.sql), then reload.";
const DOC_LABEL = { invoice: "Invoice", credit_note: "Credit note", debit_note: "Debit note" };
const addDays = (d, n) => { const x = new Date(`${d}T12:00:00`); x.setDate(x.getDate() + (+n || 0)); return isoDate(x); };
const stateOptions = Object.entries(STATES).map(([c, n]) => <option key={c} value={c}>{n} ({c})</option>);

// ── Customers ────────────────────────────────────────────────────────────────
export function CustomersView() {
  const { role } = useAuth();
  const st = useStore(["customers", "invoices", "receipts"]);
  const [q, setQ] = useState("");
  const [edit, setEdit] = useState(null);
  const outstanding = useMemo(() => balances(st.invoices, st.receipts), [st.invoices, st.receipts]);
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;
  const rows = st.customers.filter(c => !q || `${c.name} ${c.gstin} ${c.state}`.toLowerCase().includes(q.toLowerCase()));

  async function del(c) {
    if (st.invoices.some(i => i.customer_id === c.id)) return notify("This customer has invoices, so it cannot be deleted.", "err");
    if (!confirm(`Delete ${c.name}?`)) return;
    try { await cdb.deleteCustomer(c.id); removeIn("customers", c.id); notify("Deleted", "ok"); } catch (e) { notify(e.message, "err"); }
  }

  return (
    <>
      <div className="table-wrap">
        <div className="table-toolbar">
          <h3>Customers <span className="mono" style={{ color: "var(--text3)" }}>· {st.customers.length}</span></h3>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <SearchBar value={q} onChange={setQ} placeholder="Search name, GSTIN, state…" />
            <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("Customers", st.customers.map(c => ({ Name: c.name, GSTIN: c.gstin, PAN: c.pan, State: stateLabel(c.state_code), Contact: c.contact, Phone: c.phone, Email: c.email, Address: c.address, "Credit days": c.credit_days })))}>Export</button>
            {canWrite(role) && <button className="btn btn-primary btn-sm" onClick={() => setEdit({ credit_days: 30, state_code: "36" })}>+ Customer</button>}
          </div>
        </div>
        {!rows.length ? <EmptyState icon="🏢" message={st.loading ? "Loading…" : "No customers yet"} sub="Add the buyers you invoice. GSTIN fills in the state and PAN." /> : (
          <div className="table-scroll"><table>
            <thead><tr><th>Name</th><th>GSTIN</th><th>State</th><th>Contact</th><th className="num">Credit days</th><th className="num">Outstanding</th><th></th></tr></thead>
            <tbody>{rows.map(c => {
              const o = outstanding.byCustomer[c.id] || 0;
              return (
                <tr key={c.id}>
                  <td><b>{c.name}</b>{c.deducts_tds && <span className="chip" style={{ marginLeft: 6 }}>deducts TDS</span>}</td>
                  <td className="mono">{c.gstin || <span style={{ color: "var(--text3)" }}>Unregistered</span>}</td>
                  <td>{stateLabel(c.state_code)}</td>
                  <td>{[c.contact, c.phone].filter(Boolean).join(" · ")}</td>
                  <td className="num">{c.credit_days}</td>
                  <td className="num" style={{ color: o > 0.5 ? "var(--amber)" : "var(--text3)" }}>{fmt(o)}</td>
                  <td className="r">{canWrite(role) && <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <button className="btn btn-ghost btn-sm" onClick={() => setEdit(c)}>Edit</button>
                    <button className="btn btn-danger btn-sm" onClick={() => del(c)}>Delete</button>
                  </div>}</td>
                </tr>
              );
            })}</tbody>
          </table></div>
        )}
      </div>
      {edit && <CustomerModal value={edit} onClose={() => setEdit(null)} />}
    </>
  );
}

function CustomerModal({ value, onClose, onSaved }) {
  const [f, setF] = useState({ ...value });
  const [busy, setBusy] = useState(false);
  const set = k => e => setF(x => ({ ...x, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  const gErr = checkGSTIN(f.gstin);
  function onGstin(e) {
    const g = e.target.value.toUpperCase();
    setF(x => ({ ...x, gstin: g, ...(g.length >= 12 && !checkGSTIN(g) ? { state_code: g.slice(0, 2), pan: panFromGSTIN(g) } : {}) }));
  }
  async function save() {
    if (!f.name?.trim()) return notify("Name is required", "err");
    if (gErr) return notify(gErr, "err");
    setBusy(true);
    try {
      const row = await cdb.saveCustomer({ ...f, state: STATES[f.state_code] || f.state });
      upsertIn("customers", row); notify("Customer saved", "ok"); onSaved?.(row); onClose();
    } catch (e) { notify(e.message, "err"); } finally { setBusy(false); }
  }
  return (
    <Modal title={f.id ? "Edit customer" : "New customer"} onClose={onClose} size="modal-lg"
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" disabled={busy} onClick={save}>Save</button></>}>
      <div className="form-row cols-2">
        <FG label="Name *"><input value={f.name || ""} onChange={set("name")} autoFocus /></FG>
        <FG label="GSTIN" note={gErr || (f.gstin ? "✓ Valid format and check digit" : "Leave blank for an unregistered buyer")}><input value={f.gstin || ""} onChange={onGstin} maxLength={15} style={gErr ? { borderColor: "var(--red)" } : {}} /></FG>
      </div>
      <div className="form-row cols-3">
        <FG label="State (place of supply)"><select value={f.state_code || ""} onChange={set("state_code")}><option value="">Select…</option>{stateOptions}</select></FG>
        <FG label="PAN" note={checkPAN(f.pan)}><input value={f.pan || ""} onChange={e => setF(x => ({ ...x, pan: e.target.value.toUpperCase() }))} maxLength={10} /></FG>
        <FG label="Credit days"><input type="number" min="0" value={f.credit_days ?? 30} onChange={set("credit_days")} /></FG>
      </div>
      <div className="form-row cols-3">
        <FG label="Contact person"><input value={f.contact || ""} onChange={set("contact")} /></FG>
        <FG label="Phone"><input value={f.phone || ""} onChange={set("phone")} /></FG>
        <FG label="Email"><input type="email" value={f.email || ""} onChange={set("email")} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Billing address"><textarea rows={3} value={f.address || ""} onChange={set("address")} /></FG>
        <FG label="Shipping address (if different)"><textarea rows={3} value={f.ship_address || ""} onChange={set("ship_address")} /></FG>
      </div>
      <label className="check"><input type="checkbox" checked={!!f.deducts_tds} onChange={set("deducts_tds")} /> This buyer deducts TDS on purchase of goods (0.1%) from our payments</label>
    </Modal>
  );
}

// ── Balances / ageing ────────────────────────────────────────────────────────
function balances(invoices, receipts) {
  const paid = {};
  receipts.forEach(r => { paid[r.invoice_id] = (paid[r.invoice_id] || 0) + (+r.amount || 0) + (+r.tds_by_buyer || 0); });
  const byInvoice = {}, byCustomer = {};
  invoices.filter(i => i.status === "issued").forEach(i => {
    const sign = i.doc_type === "credit_note" ? -1 : 1;
    const bal = sign * (+i.grand_total) - (paid[i.id] || 0);
    byInvoice[i.id] = { paid: paid[i.id] || 0, balance: bal };
    if (i.customer_id) byCustomer[i.customer_id] = (byCustomer[i.customer_id] || 0) + bal;
  });
  return { byInvoice, byCustomer };
}

// ── Sales invoices ───────────────────────────────────────────────────────────
export function SalesInvoicesView({ bankTxns = [] }) {
  const { role } = useAuth();
  const st = useStore(["settings", "customers", "invoices", "receipts", "filings"]);
  const [tab, setTab] = useState("list");
  const [q, setQ] = useState("");
  const [month, setMonth] = useState("");
  const [type, setType] = useState("");
  const [edit, setEdit] = useState(null);
  const [recFor, setRecFor] = useState(null);
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;

  const bal = balances(st.invoices, st.receipts);
  const months = [...new Set(st.invoices.map(i => monthKey(i.invoice_date)))].sort().reverse();
  const rows = st.invoices.filter(i =>
    (!month || monthKey(i.invoice_date) === month) && (!type || i.doc_type === type) &&
    (!q || `${i.invoice_no} ${i.bill_name} ${i.bill_gstin} ${i.vehicle_no} ${i.eway_bill_no}`.toLowerCase().includes(q.toLowerCase())));

  const live = st.invoices.filter(i => i.status === "issued");
  const thisMonth = monthKey(today());
  const monthSales = live.filter(i => monthKey(i.invoice_date) === thisMonth).reduce((a, i) => a + (i.doc_type === "credit_note" ? -1 : 1) * +i.taxable_total, 0);
  const totalDue = Object.values(bal.byInvoice).reduce((a, b) => a + Math.max(0, b.balance), 0);
  const overdue = live.filter(i => i.doc_type === "invoice" && i.due_date && i.due_date < today() && bal.byInvoice[i.id]?.balance > 0.5);
  const noEway = live.filter(i => i.doc_type === "invoice" && +i.grand_total > 50000 && !i.eway_bill_no);
  const noIrn = st.settings.einvoice ? live.filter(i => i.bill_gstin && !i.irn) : [];

  function newDoc(docType = "invoice", base = null) {
    const s = st.settings;
    setEdit({
      doc_type: docType, invoice_date: today(), place_of_supply: s.state_code, reverse_charge: false,
      items: [{ description: "Quartz powder", hsn: s.default_hsn, qty: "", unit: s.default_unit || "MTS", rate: "", discount: "", gst_rate: s.default_gst_rate }],
      ...(base ? {
        customer_id: base.customer_id, bill_name: base.bill_name, bill_gstin: base.bill_gstin, bill_address: base.bill_address,
        bill_state_code: base.bill_state_code, ship_name: base.ship_name, ship_address: base.ship_address,
        place_of_supply: base.place_of_supply, against_invoice: base.invoice_no, items: (base.items || []).map(l => ({ ...l })),
      } : {}),
    });
  }

  async function cancel(inv) {
    const reason = prompt(`Cancel ${inv.invoice_no}? The number stays used and shows as cancelled in GSTR-1.\n\nReason:`);
    if (reason == null) return;
    try { const row = await cdb.cancelInvoice(inv.id, reason); upsertIn("invoices", row); notify("Cancelled", "ok"); } catch (e) { notify(e.message, "err"); }
  }

  function exportList() {
    exportRowsToExcel("Sales-register", rows.map(i => ({
      Type: DOC_LABEL[i.doc_type], Number: i.invoice_no, Date: fmtDate(i.invoice_date), Status: i.status, Customer: i.bill_name, GSTIN: i.bill_gstin,
      "Place of supply": stateLabel(i.place_of_supply), Taxable: +i.taxable_total, CGST: +i.cgst, SGST: +i.sgst, IGST: +i.igst, "Round off": +i.round_off,
      Total: +i.grand_total, Received: bal.byInvoice[i.id]?.paid || 0, Balance: bal.byInvoice[i.id]?.balance ?? 0,
      "E-way bill": i.eway_bill_no, Vehicle: i.vehicle_no, IRN: i.irn,
    })), "Sales");
  }

  return (
    <>
      <div className="stats-grid">
        <StatCard label="Sales this month (taxable)" value={fmt(monthSales)} color="accent" />
        <StatCard label="Receivable" value={fmt(totalDue)} color="amber" onClick={() => setTab("ageing")} sub="Click for ageing" />
        <StatCard label="Overdue invoices" value={overdue.length} color={overdue.length ? "red" : "green"} onClick={() => setTab("ageing")} />
        <StatCard label="Missing e-way / IRN" value={noEway.length + noIrn.length} color={noEway.length + noIrn.length ? "red" : "green"} sub={`${noEway.length} e-way · ${noIrn.length} IRN`} />
      </div>

      {(noEway.length > 0 || noIrn.length > 0) && (
        <div className="callout callout-amber">
          {noEway.length > 0 && <div><b>E-way bill missing</b> on {noEway.map(i => i.invoice_no).slice(0, 6).join(", ")}{noEway.length > 6 ? "…" : ""}. Goods worth more than ₹50,000 need an e-way bill before they move.</div>}
          {noIrn.length > 0 && <div><b>IRN missing</b> on {noIrn.length} B2B invoice(s). E-invoicing is switched on in Tax Settings: generate the IRN on the IRP and enter it on the invoice.</div>}
        </div>
      )}

      <div className="pill-tabs">
        <button className={`pill-tab${tab === "list" ? " active" : ""}`} onClick={() => setTab("list")}>Register</button>
        <button className={`pill-tab${tab === "ageing" ? " active" : ""}`} onClick={() => setTab("ageing")}>Receivables ageing</button>
      </div>

      {tab === "ageing" ? <Ageing invoices={st.invoices} customers={st.customers} bal={bal} onReceipt={setRecFor} /> : (
        <div className="table-wrap">
          <div className="table-toolbar">
            <h3>Sales register</h3>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <SearchBar value={q} onChange={setQ} placeholder="Number, customer, vehicle…" style={{ minWidth: 200 }} />
              <select className="inline-select" value={month} onChange={e => setMonth(e.target.value)}><option value="">All months</option>{months.map(m => <option key={m} value={m}>{m}</option>)}</select>
              <select className="inline-select" value={type} onChange={e => setType(e.target.value)}><option value="">All types</option>{Object.entries(DOC_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select>
              <button className="btn btn-ghost btn-sm" onClick={exportList}>Export</button>
              {canWrite(role) && <button className="btn btn-primary btn-sm" onClick={() => newDoc()}>+ New invoice</button>}
            </div>
          </div>
          {!rows.length ? <EmptyState icon="🧾" message={st.loading ? "Loading…" : "No invoices yet"} sub="Create GST tax invoices with consecutive numbering, PDF in three copies, e-way and IRN fields." /> : (
            <div className="table-scroll"><table>
              <thead><tr><th>Number</th><th>Date</th><th>Customer</th><th className="num">Taxable</th><th className="num">GST</th><th className="num">Total</th><th className="num">Balance</th><th>Docs</th><th></th></tr></thead>
              <tbody>{rows.map(i => {
                const b = bal.byInvoice[i.id];
                const due = i.status === "issued" && i.due_date && i.due_date < today() && b?.balance > 0.5;
                return (
                  <tr key={i.id} style={i.status === "cancelled" ? { opacity: .55 } : {}}>
                    <td className="mono"><b>{i.invoice_no}</b>{i.doc_type !== "invoice" && <div><Badge type="purple">{DOC_LABEL[i.doc_type]}</Badge></div>}{i.status === "cancelled" && <div><Badge type="red">Cancelled</Badge></div>}</td>
                    <td className="mono">{fmtDate(i.invoice_date)}</td>
                    <td>{i.bill_name}<div className="mono" style={{ color: "var(--text3)", fontSize: 10.5 }}>{i.bill_gstin || "B2C"} · {STATES[i.place_of_supply]}</div></td>
                    <td className="num">{fmt(i.taxable_total)}</td>
                    <td className="num">{fmt(+i.cgst + +i.sgst + +i.igst)}</td>
                    <td className="num"><b>{fmt(i.grand_total)}</b></td>
                    <td className="num" style={{ color: !b ? "var(--text3)" : b.balance > 0.5 ? (due ? "var(--red)" : "var(--amber)") : "var(--green)" }}>{b ? (b.balance > 0.5 ? fmt(b.balance) : "Paid") : "—"}</td>
                    <td>
                      {+i.grand_total > 50000 && i.doc_type === "invoice" && <span className={`chip${i.eway_bill_no ? "" : " chip-warn"}`} title={i.eway_bill_no || "No e-way bill"}>EWB {i.eway_bill_no ? "✓" : "✗"}</span>}
                      {st.settings.einvoice && i.bill_gstin && <span className={`chip${i.irn ? "" : " chip-warn"}`} style={{ marginLeft: 4 }}>IRN {i.irn ? "✓" : "✗"}</span>}
                    </td>
                    <td className="r"><div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button className="btn btn-info btn-sm" onClick={() => invoicePDF(i, st.settings)}>PDF</button>
                      {i.status === "issued" && i.doc_type === "invoice" && <button className="btn btn-ghost btn-sm" onClick={() => setRecFor(i)}>Receipts</button>}
                      {canWrite(role) && i.status === "issued" && <button className="btn btn-ghost btn-sm" onClick={() => setEdit({ ...i, items: i.items.map(l => ({ ...l })) })}>Edit</button>}
                      {canWrite(role) && i.status === "issued" && i.doc_type === "invoice" && <button className="btn btn-ghost btn-sm" title="Credit note against this invoice" onClick={() => newDoc("credit_note", i)}>CN</button>}
                      {canWrite(role) && i.status === "issued" && <button className="btn btn-danger btn-sm" onClick={() => cancel(i)}>Cancel</button>}
                    </div></td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </div>
      )}
      {edit && <InvoiceModal value={edit} settings={st.settings} customers={st.customers} filings={st.filings} onClose={() => setEdit(null)} />}
      {recFor && <ReceiptsModal invoice={recFor} receipts={st.receipts.filter(r => r.invoice_id === recFor.id)} bankTxns={bankTxns} onClose={() => setRecFor(null)} canEdit={canWrite(role)} />}
    </>
  );
}

function Ageing({ invoices, customers, bal, onReceipt }) {
  const t = today();
  const buckets = ["Not due", "0–30", "31–60", "61–90", "90+"];
  const rows = {};
  invoices.filter(i => i.status === "issued" && i.doc_type === "invoice" && bal.byInvoice[i.id]?.balance > 0.5).forEach(i => {
    const days = i.due_date ? Math.floor((new Date(t) - new Date(i.due_date)) / 864e5) : 0;
    const b = days <= 0 ? 0 : days <= 30 ? 1 : days <= 60 ? 2 : days <= 90 ? 3 : 4;
    const k = i.customer_id || i.bill_name;
    if (!rows[k]) rows[k] = { name: customers.find(c => c.id === i.customer_id)?.name || i.bill_name, b: [0, 0, 0, 0, 0], list: [] };
    rows[k].b[b] += bal.byInvoice[i.id].balance;
    rows[k].list.push({ ...i, days, balance: bal.byInvoice[i.id].balance });
  });
  const list = Object.values(rows).sort((a, b) => b.b.reduce((x, y) => x + y) - a.b.reduce((x, y) => x + y));
  const tot = buckets.map((_, j) => list.reduce((a, r) => a + r.b[j], 0));
  const [open, setOpen] = useState(null);
  return (
    <div className="table-wrap">
      <div className="table-toolbar"><h3>Receivables ageing <span className="mono" style={{ color: "var(--text3)" }}>· by days past due date</span></h3>
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("Receivables-ageing", list.flatMap(r => r.list.map(i => ({ Customer: r.name, Invoice: i.invoice_no, Date: fmtDate(i.invoice_date), Due: fmtDate(i.due_date), "Days overdue": Math.max(0, i.days), Balance: i.balance }))))}>Export</button>
      </div>
      {!list.length ? <EmptyState icon="✅" message="Nothing outstanding" /> : (
        <div className="table-scroll"><table>
          <thead><tr><th>Customer</th>{buckets.map(b => <th key={b} className="num">{b}</th>)}<th className="num">Total</th></tr></thead>
          <tbody>
            {list.map(r => (
              <React.Fragment key={r.name}>
                <tr style={{ cursor: "pointer" }} onClick={() => setOpen(open === r.name ? null : r.name)}>
                  <td><b>{r.name}</b> <span className="mono" style={{ color: "var(--text3)" }}>· {r.list.length} inv</span></td>
                  {r.b.map((v, j) => <td key={j} className="num" style={{ color: v > 0.5 ? (j >= 3 ? "var(--red)" : j >= 1 ? "var(--amber)" : "var(--text)") : "var(--text4)" }}>{v > 0.5 ? fmt(v) : "—"}</td>)}
                  <td className="num"><b>{fmt(r.b.reduce((x, y) => x + y))}</b></td>
                </tr>
                {open === r.name && r.list.map(i => (
                  <tr key={i.id} style={{ background: "var(--bg3)" }}>
                    <td className="mono" style={{ paddingLeft: 28 }}>{i.invoice_no} · {fmtDate(i.invoice_date)} · due {fmtDate(i.due_date)}</td>
                    <td colSpan={5} className="mono" style={{ color: i.days > 0 ? "var(--red)" : "var(--text3)" }}>{i.days > 0 ? `${i.days} days overdue` : "not yet due"}</td>
                    <td className="num">{fmt(i.balance)} <button className="btn btn-ghost btn-sm" onClick={e => { e.stopPropagation(); onReceipt(i); }}>Receipt</button></td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
            <tr style={{ borderTop: "2px solid var(--border2)" }}><td><b>Total</b></td>{tot.map((v, j) => <td key={j} className="num"><b>{fmt(v)}</b></td>)}<td className="num"><b>{fmt(tot.reduce((a, b) => a + b))}</b></td></tr>
          </tbody>
        </table></div>
      )}
    </div>
  );
}

function InvoiceModal({ value, settings, customers, filings, onClose }) {
  const [f, setF] = useState(value);
  const [busy, setBusy] = useState(false);
  const [newCust, setNewCust] = useState(false);
  const isEdit = !!f.id;
  const lockedMonth = isEdit && filings.some(x => x.key === `GSTR1:${monthKey(value.invoice_date)}` && x.done_on);
  const calc = useMemo(() => computeInvoice(f.items, settings.state_code, f.place_of_supply), [f.items, f.place_of_supply, settings.state_code]);
  const set = k => e => setF(x => ({ ...x, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  const setItem = (i, k, v) => setF(x => ({ ...x, items: x.items.map((l, j) => (j === i ? { ...l, [k]: v } : l)) }));
  const cust = customers.find(c => c.id === +f.customer_id);

  function pickCustomer(id) {
    const c = customers.find(x => x.id === +id);
    if (!c) return setF(x => ({ ...x, customer_id: "" }));
    setF(x => ({
      ...x, customer_id: c.id, bill_name: c.name, bill_gstin: c.gstin || "", bill_address: c.address || "", bill_state_code: c.state_code || "",
      ship_name: c.name, ship_address: c.ship_address || c.address || "", place_of_supply: c.state_code || x.place_of_supply,
      due_date: addDays(x.invoice_date, c.credit_days ?? 30),
    }));
  }

  const warnings = [];
  const gErr = checkGSTIN(f.bill_gstin);
  if (gErr) warnings.push(gErr);
  if (f.doc_type === "invoice" && calc.grand_total > 50000 && !f.eway_bill_no) warnings.push("Value over ₹50,000: an e-way bill is needed before the goods move. Add its number once generated.");
  if (settings.einvoice && f.bill_gstin && !f.irn) warnings.push("E-invoicing is on: report this invoice on the IRP and enter the IRN / Ack No.");
  if (f.doc_type !== "invoice" && !f.against_invoice) warnings.push("Credit / debit notes should name the original invoice.");
  if (cust?.deducts_tds) warnings.push(`${cust.name} deducts TDS on goods (0.1%); record it as "TDS by buyer" when the payment comes in.`);

  async function save() {
    if (!f.bill_name?.trim()) return notify("Choose or type the customer", "err");
    if (!f.place_of_supply) return notify("Place of supply is required", "err");
    if (gErr) return notify(gErr, "err");
    if (!calc.lines.length || calc.lines.some(l => !l.description || !(+l.qty > 0) || !(+l.rate > 0))) return notify("Every line needs a description, quantity and rate", "err");
    if (calc.lines.some(l => !l.hsn)) return notify("Every line needs an HSN code", "err");
    setBusy(true);
    try {
      const fy = fyOfDate(f.invoice_date);
      const invoice_no = isEdit ? f.invoice_no : await cdb.nextDocNo(f.doc_type, f.invoice_date);
      const row = {
        ...(isEdit ? { id: f.id } : {}),
        doc_type: f.doc_type, invoice_no, invoice_date: f.invoice_date, fy, status: "issued",
        customer_id: f.customer_id ? +f.customer_id : null, bill_name: f.bill_name.trim(), bill_gstin: (f.bill_gstin || "").toUpperCase() || null,
        bill_address: f.bill_address || null, bill_state_code: f.bill_state_code || null, ship_name: f.ship_name || null, ship_address: f.ship_address || null,
        place_of_supply: f.place_of_supply, reverse_charge: !!f.reverse_charge, against_invoice: f.against_invoice || null,
        items: calc.lines.map(l => ({ description: l.description, hsn: l.hsn, qty: +l.qty, unit: l.unit, rate: +l.rate, discount: +l.discount || 0, gst_rate: +l.gst_rate, taxable: l.taxable, cgst: l.cgst, sgst: l.sgst, igst: l.igst, total: l.total })),
        taxable_total: calc.taxable_total, cgst: calc.cgst, sgst: calc.sgst, igst: calc.igst, round_off: calc.round_off, grand_total: calc.grand_total,
        due_date: f.due_date || null, vehicle_no: (f.vehicle_no || "").toUpperCase() || null, transporter: f.transporter || null, lr_no: f.lr_no || null,
        eway_bill_no: f.eway_bill_no || null, eway_bill_date: f.eway_bill_date || null, irn: f.irn || null, ack_no: f.ack_no || null, ack_date: f.ack_date || null,
        ops_ref: f.ops_ref || null, notes: f.notes || null,
      };
      const saved = await cdb.saveInvoice(row);
      upsertIn("invoices", saved);
      notify(`${DOC_LABEL[saved.doc_type]} ${saved.invoice_no} saved`, "ok");
      onClose();
      if (!isEdit && confirm(`Download the PDF of ${saved.invoice_no} now?`)) invoicePDF(saved, settings);
    } catch (e) { notify(e.message || String(e), "err"); } finally { setBusy(false); }
  }

  const amountsLocked = lockedMonth;
  return (
    <Modal title={isEdit ? `Edit ${f.invoice_no}` : `New ${DOC_LABEL[f.doc_type].toLowerCase()}`} onClose={onClose} size="modal-xl"
      foot={<><span className="mono" style={{ marginRight: "auto", color: "var(--text3)" }}>{isEdit ? "Number stays the same" : "Number is given on save (next in series)"}</span>
        <button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" disabled={busy} onClick={save}>{busy ? "Saving…" : isEdit ? "Save changes" : "Issue"}</button></>}>
      {amountsLocked && <div className="callout callout-red">GSTR-1 for this month is marked filed, so customer and amounts are locked. Use a credit or debit note to correct them. Transport, e-way and IRN details can still be added.</div>}
      <div className="form-row cols-4">
        <FG label="Type"><select value={f.doc_type} disabled={isEdit} onChange={set("doc_type")}>{Object.entries(DOC_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}</select></FG>
        <FG label="Date"><input type="date" value={f.invoice_date} disabled={isEdit} onChange={e => setF(x => ({ ...x, invoice_date: e.target.value, due_date: cust ? addDays(e.target.value, cust.credit_days ?? 30) : x.due_date }))} /></FG>
        <FG label="Due date"><input type="date" value={f.due_date || ""} onChange={set("due_date")} /></FG>
        {f.doc_type !== "invoice" ? <FG label="Against invoice no."><input value={f.against_invoice || ""} onChange={set("against_invoice")} /></FG>
          : <FG label="Ops / dispatch ref."><input value={f.ops_ref || ""} onChange={set("ops_ref")} placeholder="Order / dispatch no." /></FG>}
      </div>
      <fieldset disabled={amountsLocked} style={{ border: 0, padding: 0, margin: 0, minWidth: 0 }}>
        <div className="form-row cols-3">
          <FG label="Customer" note={<button type="button" className="link-btn" onClick={() => setNewCust(true)}>+ add a new customer</button>}>
            <select value={f.customer_id || ""} onChange={e => pickCustomer(e.target.value)}><option value="">Select or type below…</option>{customers.map(c => <option key={c.id} value={c.id}>{c.name}{c.gstin ? ` · ${c.gstin}` : ""}</option>)}</select>
          </FG>
          <FG label="Bill-to name *"><input value={f.bill_name || ""} onChange={set("bill_name")} /></FG>
          <FG label="Buyer GSTIN" note={gErr}><input value={f.bill_gstin || ""} maxLength={15} onChange={e => { const g = e.target.value.toUpperCase(); setF(x => ({ ...x, bill_gstin: g, ...(g.length === 15 && !checkGSTIN(g) ? { bill_state_code: g.slice(0, 2), place_of_supply: g.slice(0, 2) } : {}) })); }} /></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Billing address"><textarea rows={2} value={f.bill_address || ""} onChange={set("bill_address")} /></FG>
          <FG label="Ship to (name & address)"><textarea rows={2} value={[f.ship_name, f.ship_address].filter(Boolean).join("\n")} onChange={e => { const [n, ...rest] = e.target.value.split("\n"); setF(x => ({ ...x, ship_name: n, ship_address: rest.join("\n") })); }} /></FG>
          <div>
            <FG label="Place of supply *" note={calc.inter ? "Inter-state → IGST" : "Intra-state → CGST + SGST"}><select value={f.place_of_supply || ""} onChange={set("place_of_supply")}>{stateOptions}</select></FG>
            <label className="check" style={{ marginTop: 8 }}><input type="checkbox" checked={!!f.reverse_charge} onChange={set("reverse_charge")} /> Reverse charge</label>
          </div>
        </div>

        <div className="section-title">Items</div>
        <div className="table-scroll" style={{ border: "1px solid var(--border)", borderRadius: "var(--r)" }}>
          <table className="edit-table">
            <thead><tr><th>Description</th><th>HSN</th><th className="num">Qty</th><th>Unit</th><th className="num">Rate ₹</th><th className="num">Discount ₹</th><th className="num">GST %</th><th className="num">Taxable</th><th className="num">Tax</th><th></th></tr></thead>
            <tbody>{f.items.map((l, i) => {
              const c = calc.lines[i] || {};
              return (
                <tr key={i}>
                  <td><input value={l.description || ""} onChange={e => setItem(i, "description", e.target.value)} placeholder="e.g. Quartz powder 200 mesh" /></td>
                  <td style={{ width: 96 }}><input value={l.hsn || ""} onChange={e => setItem(i, "hsn", e.target.value)} /></td>
                  <td style={{ width: 86 }}><input type="number" step="0.001" value={l.qty} onChange={e => setItem(i, "qty", e.target.value)} /></td>
                  <td style={{ width: 80 }}><select value={l.unit} onChange={e => setItem(i, "unit", e.target.value)}>{Object.keys(UQC).map(u => <option key={u}>{u}</option>)}</select></td>
                  <td style={{ width: 100 }}><input type="number" step="0.01" value={l.rate} onChange={e => setItem(i, "rate", e.target.value)} /></td>
                  <td style={{ width: 90 }}><input type="number" step="0.01" value={l.discount || ""} onChange={e => setItem(i, "discount", e.target.value)} /></td>
                  <td style={{ width: 72 }}><select value={l.gst_rate} onChange={e => setItem(i, "gst_rate", e.target.value)}>{[0, 5, 12, 18, 28, 40].map(r => <option key={r} value={r}>{r}</option>)}</select></td>
                  <td className="num">{fmt(c.taxable)}</td>
                  <td className="num">{fmt((c.igst || 0) + (c.cgst || 0) + (c.sgst || 0))}</td>
                  <td>{f.items.length > 1 && <button type="button" className="btn btn-danger btn-sm" onClick={() => setF(x => ({ ...x, items: x.items.filter((_, j) => j !== i) }))}>✕</button>}</td>
                </tr>
              );
            })}</tbody>
          </table>
        </div>
        <button type="button" className="btn btn-ghost btn-sm" style={{ marginTop: 8 }} onClick={() => setF(x => ({ ...x, items: [...x.items, { ...x.items[x.items.length - 1], description: "", qty: "", rate: "", discount: "" }] }))}>+ Line</button>
      </fieldset>

      <div className="invoice-totals">
        <div><span>Taxable</span><b>{fmt(calc.taxable_total)}</b></div>
        {calc.inter ? <div><span>IGST</span><b>{fmt(calc.igst)}</b></div> : <><div><span>CGST</span><b>{fmt(calc.cgst)}</b></div><div><span>SGST</span><b>{fmt(calc.sgst)}</b></div></>}
        {calc.round_off !== 0 && <div><span>Round off</span><b>{fmt(calc.round_off)}</b></div>}
        <div className="grand"><span>Total</span><b>{fmt(calc.grand_total)}</b></div>
      </div>

      <div className="section-title">Transport & e-documents</div>
      <div className="form-row cols-4">
        <FG label="Vehicle no."><input value={f.vehicle_no || ""} onChange={set("vehicle_no")} /></FG>
        <FG label="Transporter"><input value={f.transporter || ""} onChange={set("transporter")} /></FG>
        <FG label="LR / GR no."><input value={f.lr_no || ""} onChange={set("lr_no")} /></FG>
        <FG label="E-way bill no." note={calc.grand_total > 50000 ? "Required (> ₹50,000)" : ""}><input value={f.eway_bill_no || ""} onChange={set("eway_bill_no")} /></FG>
      </div>
      <div className="form-row cols-4">
        <FG label="E-way bill date"><input type="date" value={f.eway_bill_date || ""} onChange={set("eway_bill_date")} /></FG>
        <FG label="IRN (e-invoice)"><input value={f.irn || ""} onChange={set("irn")} /></FG>
        <FG label="Ack no."><input value={f.ack_no || ""} onChange={set("ack_no")} /></FG>
        <FG label="Ack date"><input type="date" value={f.ack_date || ""} onChange={set("ack_date")} /></FG>
      </div>
      <FG label="Note on invoice"><input value={f.notes || ""} onChange={set("notes")} /></FG>
      {warnings.length > 0 && <div className="callout callout-amber" style={{ marginTop: 14 }}>{warnings.map((w, i) => <div key={i}>• {w}</div>)}</div>}
      {newCust && <CustomerModal value={{ credit_days: 30, state_code: settings.state_code }} onClose={() => setNewCust(false)} onSaved={c => setTimeout(() => pickCustomer(c.id), 0)} />}
    </Modal>
  );
}

function ReceiptsModal({ invoice, receipts, bankTxns, onClose, canEdit }) {
  const [f, setF] = useState({ receipt_date: today(), amount: "", tds_by_buyer: "", mode: "bank", reference: "", bank_txn_id: "" });
  const paid = receipts.reduce((a, r) => a + +r.amount + +r.tds_by_buyer, 0);
  const balance = +invoice.grand_total - paid;
  const credits = useMemo(() => bankTxns.filter(t => +t.credit > 0 && Math.abs(new Date(t.date) - new Date(invoice.invoice_date)) < 120 * 864e5)
    .sort((a, b) => Math.abs(+a.credit - balance) - Math.abs(+b.credit - balance)).slice(0, 30), [bankTxns, invoice, balance]);
  async function add() {
    const amt = +f.amount || 0, tds = +f.tds_by_buyer || 0;
    if (amt + tds <= 0) return notify("Enter the amount received", "err");
    if (amt + tds > balance + 1 && !confirm("This is more than the balance. Save anyway?")) return;
    try {
      const row = await cdb.saveReceipt({ invoice_id: invoice.id, receipt_date: f.receipt_date, amount: amt, tds_by_buyer: tds, mode: f.mode, reference: f.reference || null, bank_txn_id: f.bank_txn_id || null });
      upsertIn("receipts", row); notify("Receipt saved", "ok");
      setF(x => ({ ...x, amount: "", tds_by_buyer: "", reference: "", bank_txn_id: "" }));
    } catch (e) { notify(e.message, "err"); }
  }
  async function del(r) {
    if (!confirm("Delete this receipt?")) return;
    try { await cdb.deleteReceipt(r.id); removeIn("receipts", r.id); } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={`Receipts · ${invoice.invoice_no}`} onClose={onClose} size="modal-lg">
      <div className="grid-auto">
        <StatCard label="Invoice total" value={fmt(invoice.grand_total)} color="blue" />
        <StatCard label="Received + TDS" value={fmt(paid)} color="green" />
        <StatCard label="Balance" value={fmt(balance)} color={balance > 0.5 ? "amber" : "green"} />
      </div>
      {receipts.length > 0 && <table style={{ marginBottom: 16 }}><thead><tr><th>Date</th><th>Mode</th><th>Reference</th><th className="num">Amount</th><th className="num">TDS by buyer</th><th></th></tr></thead>
        <tbody>{receipts.map(r => <tr key={r.id}><td className="mono">{fmtDate(r.receipt_date)}</td><td>{r.mode}</td><td className="mono">{r.reference}{r.bank_txn_id && <span className="chip" style={{ marginLeft: 6 }}>bank-linked</span>}</td><td className="num">{fmt(r.amount)}</td><td className="num">{fmt(r.tds_by_buyer)}</td><td className="r">{canEdit && <button className="btn btn-danger btn-sm" onClick={() => del(r)}>✕</button>}</td></tr>)}</tbody></table>}
      {canEdit && balance > 0.5 && <>
        <div className="section-title" style={{ marginTop: 0 }}>Add receipt</div>
        <div className="form-row cols-4">
          <FG label="Date"><input type="date" value={f.receipt_date} onChange={e => setF(x => ({ ...x, receipt_date: e.target.value }))} /></FG>
          <FG label="Amount received"><input type="number" step="0.01" value={f.amount} onChange={e => setF(x => ({ ...x, amount: e.target.value }))} placeholder={balance.toFixed(2)} /></FG>
          <FG label="TDS deducted by buyer"><input type="number" step="0.01" value={f.tds_by_buyer} onChange={e => setF(x => ({ ...x, tds_by_buyer: e.target.value }))} /></FG>
          <FG label="Mode"><select value={f.mode} onChange={e => setF(x => ({ ...x, mode: e.target.value }))}><option value="bank">Bank</option><option value="cash">Cash</option><option value="cheque">Cheque</option><option value="adjustment">Adjustment</option></select></FG>
        </div>
        <div className="form-row cols-2">
          <FG label="Match a bank credit (optional)"><select value={f.bank_txn_id} onChange={e => { const t = credits.find(x => String(x.id) === e.target.value); setF(x => ({ ...x, bank_txn_id: e.target.value, ...(t ? { amount: String(t.credit), receipt_date: t.date, reference: (t.description || "").slice(0, 60) } : {}) })); }}>
            <option value="">—</option>{credits.map(t => <option key={t.id} value={t.id}>{fmtDate(t.date)} · {fmt(t.credit)} · {(t.description || "").slice(0, 40)}</option>)}</select></FG>
          <FG label="Reference / UTR"><input value={f.reference} onChange={e => setF(x => ({ ...x, reference: e.target.value }))} /></FG>
        </div>
        <button className="btn btn-primary" onClick={add}>Save receipt</button>
      </>}
    </Modal>
  );
}

export { CustomerModal, balances };
export const reloadSales = () => { reload("invoices"); reload("receipts"); };
