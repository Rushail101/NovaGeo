// TDS: deductions register, deposit (challan) tracking with late interest, quarterly return export, threshold check.
import React, { useMemo, useState } from "react";
import * as cdb from "./complianceDb.js";
import { useStore, upsertIn, removeIn, reload } from "./finStore.js";
import { useAuth, canWrite } from "./auth.jsx";
import { tdsRateFor, tdsDepositDue, tdsLateInterest, tdsQuarter, fyOfDate, monthKey, monthLabel, checkPAN } from "./taxLogic.js";
import { fmt, fmtDate, today, notify, Badge, StatCard, FG, Modal, EmptyState, SearchBar, SetupNotice, exportRowsToExcel } from "./ui.jsx";

const SETUP = "Run supabase/002_compliance.sql in the Supabase SQL editor (after 001_security.sql), then reload.";

export function TdsView({ bankTxns = [], bankLabels = {} }) {
  const { role } = useAuth();
  const st = useStore(["tds", "tdsSections", "suppliers", "settings"]);
  const [tab, setTab] = useState("register");
  const [fy, setFy] = useState(fyOfDate(today()));
  const [q, setQ] = useState("");
  const [edit, setEdit] = useState(null);
  const [sel, setSel] = useState(new Set());
  const [deposit, setDeposit] = useState(false);
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;

  const secOf = k => st.tdsSections.find(s => s.key === k);
  const inFy = st.tds.filter(t => fyOfDate(t.deducted_on) === fy);
  const pending = inFy.filter(t => !t.challan_date);
  const pendingAmt = pending.reduce((a, t) => a + +t.tds_amount, 0);
  const overdue = pending.filter(t => tdsDepositDue(t.deducted_on) < today());
  const interestDue = overdue.reduce((a, t) => a + tdsLateInterest(t).interest, 0);
  const rows = inFy.filter(t => !q || `${t.party_name} ${t.party_pan} ${t.challan_no} ${secOf(t.section_key)?.old_section}`.toLowerCase().includes(q.toLowerCase()));
  const fys = [...new Set([fyOfDate(today()), ...st.tds.map(t => fyOfDate(t.deducted_on))])].sort().reverse();

  async function del(t) {
    if (!confirm(`Delete TDS entry for ${t.party_name}?`)) return;
    try { await cdb.deleteTds(t.id); removeIn("tds", t.id); } catch (e) { notify(e.message, "err"); }
  }

  return (
    <>
      {!st.settings.tan && <div className="callout callout-amber"><b>TAN is not set.</b> Add it under Tax Settings; it is needed on every TDS challan and return.</div>}
      <div className="stats-grid">
        <StatCard label={`TDS deducted ${fy}`} value={fmt(inFy.reduce((a, t) => a + +t.tds_amount, 0))} sub={`${inFy.length} entries`} color="blue" />
        <StatCard label="Not yet deposited" value={fmt(pendingAmt)} sub={`${pending.length} entries`} color={pendingAmt ? "amber" : "green"} />
        <StatCard label="Deposit overdue" value={overdue.length} sub={overdue.length ? `Interest so far ≈ ${fmt(interestDue)}` : "All on time"} color={overdue.length ? "red" : "green"} />
        <StatCard label="Certificates pending" value={inFy.filter(t => t.challan_date && !t.certificate_issued).length} color="teal" />
      </div>
      <div className="filter-bar">
        <div className="pill-tabs" style={{ marginBottom: 0 }}>
          {[["register", "Register"], ["monthly", "Monthly deposits"], ["quarter", "Quarterly return"], ["check", "Threshold check"], ["rates", "Rates"]].map(([k, l]) =>
            <button key={k} className={`pill-tab${tab === k ? " active" : ""}`} onClick={() => setTab(k)}>{l}</button>)}
        </div>
        <select className="inline-select" value={fy} onChange={e => setFy(e.target.value)}>{fys.map(f => <option key={f}>{f}</option>)}</select>
      </div>

      {tab === "register" && (
        <div className="table-wrap">
          <div className="table-toolbar">
            <h3>TDS register</h3>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <SearchBar value={q} onChange={setQ} placeholder="Party, PAN, challan…" style={{ minWidth: 200 }} />
              {canWrite(role) && sel.size > 0 && <button className="btn btn-primary btn-sm" onClick={() => setDeposit(true)}>Record challan for {sel.size}</button>}
              {canWrite(role) && <button className="btn btn-primary btn-sm" onClick={() => setEdit({ deducted_on: today(), section_key: "contract" })}>+ Deduction</button>}
            </div>
          </div>
          {!rows.length ? <EmptyState icon="🧮" message="No TDS entries" sub="Record TDS whenever you pay or credit a contractor, professional, landlord or partner above the limit." /> : (
            <div className="table-scroll"><table>
              <thead><tr><th></th><th>Date</th><th>Party</th><th>Section</th><th className="num">Base</th><th className="num">Rate</th><th className="num">TDS</th><th>Deposit</th><th></th></tr></thead>
              <tbody>{rows.map(t => {
                const s = secOf(t.section_key);
                const due = tdsDepositDue(t.deducted_on);
                const late = !t.challan_date && due < today();
                const li = tdsLateInterest(t);
                return (
                  <tr key={t.id} className={late ? "due-row-overdue" : ""}>
                    <td>{!t.challan_date && canWrite(role) && <input type="checkbox" checked={sel.has(t.id)} onChange={e => setSel(x => { const n = new Set(x); e.target.checked ? n.add(t.id) : n.delete(t.id); return n; })} />}</td>
                    <td className="mono">{fmtDate(t.deducted_on)}</td>
                    <td>{t.party_name}<div className="mono" style={{ fontSize: 10.5, color: t.party_pan ? "var(--text3)" : "var(--red)" }}>{t.party_pan || "No PAN → 20%"}</div></td>
                    <td><span className="chip">{s?.old_section || t.section_key}</span><div className="mono" style={{ fontSize: 10, color: "var(--text3)" }}>{s?.new_ref}</div></td>
                    <td className="num">{fmt(t.base_amount)}</td><td className="num">{t.rate}%</td><td className="num"><b>{fmt(t.tds_amount)}</b></td>
                    <td>{t.challan_date
                      ? <><Badge type={li.interest ? "amber" : "green"}>Paid {fmtDate(t.challan_date)}</Badge><div className="mono" style={{ fontSize: 10.5, color: "var(--text3)" }}>{t.challan_no}{li.interest ? ` · int ${fmt(li.interest)}` : ""}</div></>
                      : <><Badge type={late ? "red" : "amber"}>{late ? "Overdue" : "Due"} {fmtDate(due)}</Badge>{late && <div className="mono" style={{ fontSize: 10.5, color: "var(--red)" }}>interest {fmt(li.interest)} so far</div>}</>}</td>
                    <td className="r">{canWrite(role) && <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      {t.challan_date && !t.certificate_issued && <button className="btn btn-ghost btn-sm" title="Mark Form 131 (old 16A) sent" onClick={async () => { try { upsertIn("tds", await cdb.saveTds({ id: t.id, certificate_issued: true })); } catch (e) { notify(e.message, "err"); } }}>Cert sent</button>}
                      <button className="btn btn-ghost btn-sm" onClick={() => setEdit(t)}>Edit</button>
                      <button className="btn btn-danger btn-sm" onClick={() => del(t)}>✕</button></div>}</td>
                  </tr>
                );
              })}</tbody>
            </table></div>
          )}
        </div>
      )}
      {tab === "monthly" && <Monthly entries={inFy} />}
      {tab === "quarter" && <Quarterly entries={inFy} fy={fy} secOf={secOf} settings={st.settings} />}
      {tab === "check" && <ThresholdCheck fy={fy} bankTxns={bankTxns} bankLabels={bankLabels} entries={inFy} suppliers={st.suppliers} />}
      {tab === "rates" && <Rates sections={st.tdsSections} role={role} />}
      {edit && <TdsModal value={edit} st={st} onClose={() => setEdit(null)} />}
      {deposit && <DepositModal ids={[...sel]} entries={st.tds} onClose={() => { setDeposit(false); setSel(new Set()); }} />}
    </>
  );
}

function TdsModal({ value, st, onClose }) {
  const [f, setF] = useState({ ...value });
  const sup = st.suppliers.find(s => s.id === +f.supplier_id);
  const sec = st.tdsSections.find(s => s.key === f.section_key);
  const suggested = tdsRateFor(sec, sup || { entity_type: f.entity_type || "company", pan: f.party_pan });
  const rate = f.rate === undefined || f.rate === "" ? suggested : +f.rate;
  const tds = Math.round((+f.base_amount || 0) * rate / 100);
  // How much this party has had in the year under this section (for the threshold)
  const ytd = st.tds.filter(t => t.id !== f.id && t.section_key === f.section_key && fyOfDate(t.deducted_on) === fyOfDate(f.deducted_on) && (sup ? t.supplier_id === sup.id : t.party_name === f.party_name)).reduce((a, t) => a + +t.base_amount, 0);
  const below = sec && (+f.base_amount || 0) < (sec.single_limit || Infinity) && ytd + (+f.base_amount || 0) <= (sec.annual_limit || 0);
  function pickSup(id) {
    const s = st.suppliers.find(x => x.id === +id);
    setF(x => ({ ...x, supplier_id: id, party_name: s?.name || x.party_name, party_pan: s?.pan || x.party_pan, section_key: s?.tds_section || x.section_key, rate: undefined }));
  }
  async function save() {
    if (!f.party_name || !(+f.base_amount > 0)) return notify("Party and amount are required", "err");
    if (checkPAN(f.party_pan)) return notify(checkPAN(f.party_pan), "err");
    try {
      const row = await cdb.saveTds({ id: f.id, deducted_on: f.deducted_on, section_key: f.section_key, supplier_id: f.supplier_id ? +f.supplier_id : null, party_name: f.party_name, party_pan: (f.party_pan || "").toUpperCase() || null, base_amount: +f.base_amount, rate, tds_amount: tds, source_type: f.source_type || "manual", notes: f.notes || null, ...(f.id ? {} : {}) });
      upsertIn("tds", row); notify("TDS entry saved", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={f.id ? "Edit TDS entry" : "Record TDS deduction"} onClose={onClose} size="modal-lg" foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
      <div className="form-row cols-3">
        <FG label="Supplier"><select value={f.supplier_id || ""} onChange={e => pickSup(e.target.value)}><option value="">Other / type name →</option>{st.suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG>
        <FG label="Party name *"><input value={f.party_name || ""} onChange={e => setF(x => ({ ...x, party_name: e.target.value }))} /></FG>
        <FG label="PAN" note={checkPAN(f.party_pan) || (!f.party_pan ? "Without PAN, deduct at 20%" : "")}><input maxLength={10} value={f.party_pan || ""} onChange={e => setF(x => ({ ...x, party_pan: e.target.value.toUpperCase(), rate: undefined }))} /></FG>
      </div>
      <div className="form-row cols-3">
        <FG label="Date paid / credited"><input type="date" value={f.deducted_on} onChange={e => setF(x => ({ ...x, deducted_on: e.target.value }))} /></FG>
        <FG label="Nature of payment"><select value={f.section_key} onChange={e => setF(x => ({ ...x, section_key: e.target.value, rate: undefined }))}>{st.tdsSections.filter(s => s.active).map(s => <option key={s.key} value={s.key}>{s.label} ({s.old_section})</option>)}</select></FG>
        {!sup && <FG label="Payee type"><select value={f.entity_type || "company"} onChange={e => setF(x => ({ ...x, entity_type: e.target.value, rate: undefined }))}><option value="individual">Individual</option><option value="huf">HUF</option><option value="company">Company</option><option value="firm">Firm / LLP</option><option value="other">Other</option></select></FG>}
      </div>
      <div className="form-row cols-3">
        <FG label="Amount paid / credited (excl. GST) *"><input type="number" step="0.01" value={f.base_amount || ""} onChange={e => setF(x => ({ ...x, base_amount: e.target.value }))} /></FG>
        <FG label="Rate %" note={`Suggested ${suggested}%${sup?.lower_tds_cert ? " (lower deduction certificate)" : ""}`}><input type="number" step="0.01" value={f.rate ?? suggested} onChange={e => setF(x => ({ ...x, rate: e.target.value }))} /></FG>
        <FG label="TDS to deduct"><input value={fmt(tds)} readOnly /></FG>
      </div>
      {sec && <div className={`callout ${below ? "callout-blue" : ""}`}>
        {sec.old_section} → Income-tax Act 2025 {sec.new_ref}{sec.pay_code ? `, payment code ${sec.pay_code}` : ""}. Limits: {sec.single_limit ? `${fmt(sec.single_limit)} per payment or ` : ""}{fmt(sec.annual_limit)} in the year.
        Already this year for this party: {fmt(ytd)}.{below && <b> This payment is still under the limit; TDS may not be needed yet.</b>}
      </div>}
      <FG label="Notes"><input value={f.notes || ""} onChange={e => setF(x => ({ ...x, notes: e.target.value }))} placeholder="Bill no. / project" /></FG>
    </Modal>
  );
}

function DepositModal({ ids, entries, onClose }) {
  const [f, setF] = useState({ challan_date: today(), challan_no: "", bsr_code: "" });
  const list = entries.filter(e => ids.includes(e.id));
  const total = list.reduce((a, e) => a + +e.tds_amount, 0);
  const interest = list.reduce((a, e) => a + tdsLateInterest(e, f.challan_date).interest, 0);
  async function save() {
    if (!f.challan_no || !f.challan_date) return notify("Challan number and date are required", "err");
    try {
      for (const e of list) {
        const row = await cdb.saveTds({ id: e.id, challan_no: f.challan_no, challan_date: f.challan_date, bsr_code: f.bsr_code || null, interest_paid: tdsLateInterest(e, f.challan_date).interest });
        upsertIn("tds", row);
      }
      notify("Deposit recorded", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={`Record challan · ${list.length} entries`} onClose={onClose} foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
      <div className="form-row cols-3">
        <FG label="Deposit date"><input type="date" value={f.challan_date} onChange={e => setF(x => ({ ...x, challan_date: e.target.value }))} /></FG>
        <FG label="Challan serial no."><input value={f.challan_no} onChange={e => setF(x => ({ ...x, challan_no: e.target.value }))} /></FG>
        <FG label="BSR code"><input value={f.bsr_code} onChange={e => setF(x => ({ ...x, bsr_code: e.target.value }))} maxLength={7} /></FG>
      </div>
      <div className="invoice-totals"><div><span>TDS</span><b>{fmt(total)}</b></div><div><span>Late interest (1.5%/month)</span><b style={{ color: interest ? "var(--red)" : undefined }}>{fmt(interest)}</b></div><div className="grand"><span>Challan total</span><b>{fmt(total + interest)}</b></div></div>
    </Modal>
  );
}

function Monthly({ entries }) {
  const by = {};
  entries.forEach(e => { const m = monthKey(e.deducted_on); (by[m] = by[m] || []).push(e); });
  const months = Object.keys(by).sort();
  return (
    <div className="table-wrap"><div className="table-scroll"><table>
      <thead><tr><th>Month deducted</th><th>Deposit due</th><th className="num">Deducted</th><th className="num">Deposited</th><th className="num">Pending</th><th className="num">Interest</th><th>Status</th></tr></thead>
      <tbody>{months.map(m => {
        const xs = by[m], due = tdsDepositDue(`${m}-01`);
        const ded = xs.reduce((a, e) => a + +e.tds_amount, 0), dep = xs.filter(e => e.challan_date).reduce((a, e) => a + +e.tds_amount, 0);
        const int = xs.reduce((a, e) => a + tdsLateInterest(e).interest, 0);
        const pend = ded - dep;
        return <tr key={m} className={pend > 0 && due < today() ? "due-row-overdue" : ""}><td>{monthLabel(m)}</td><td className="mono">{fmtDate(due)}</td><td className="num">{fmt(ded)}</td><td className="num">{fmt(dep)}</td><td className="num">{fmt(pend)}</td><td className="num">{int ? fmt(int) : "—"}</td>
          <td>{pend <= 0 ? <Badge type={int ? "amber" : "green"}>{int ? "Paid late" : "Paid"}</Badge> : due < today() ? <Badge type="red">Overdue</Badge> : <Badge type="amber">Due</Badge>}</td></tr>;
      })}</tbody>
    </table></div>{!months.length && <EmptyState icon="—" message="Nothing deducted in this year" />}</div>
  );
}

function Quarterly({ entries, fy, secOf, settings }) {
  const [q, setQ] = useState("Q1");
  const list = entries.filter(e => tdsQuarter(e.deducted_on) === q);
  const undeposited = list.filter(e => !e.challan_date);
  const noPan = list.filter(e => !e.party_pan);
  function exportIt() {
    exportRowsToExcel(`TDS-Form140-${fy}-${q}`, list.map((e, i) => {
      const s = secOf(e.section_key);
      return { "Sr": i + 1, "Deductor TAN": settings.tan, "Deductee PAN": e.party_pan || "PANNOTAVBL", "Deductee name": e.party_name, "Old section": s?.old_section, "Section (2025 Act)": s?.new_ref, "Payment code": s?.pay_code,
        "Date of payment/credit": fmtDate(e.deducted_on), "Amount paid/credited": +e.base_amount, "Rate %": +e.rate, "TDS deducted": +e.tds_amount, "Date of deduction": fmtDate(e.deducted_on),
        "Challan serial no.": e.challan_no, "BSR code": e.bsr_code, "Date of deposit": fmtDate(e.challan_date), "Interest paid": +e.interest_paid || 0 };
    }), q);
  }
  const due = { Q1: `31-07-${fy.slice(0, 4)}`, Q2: `31-10-${fy.slice(0, 4)}`, Q3: `31-01-${+fy.slice(0, 4) + 1}`, Q4: `31-05-${+fy.slice(0, 4) + 1}` }[q];
  return (
    <>
      <div className="filter-bar">
        <div className="pill-tabs" style={{ marginBottom: 0 }}>{["Q1", "Q2", "Q3", "Q4"].map(x => <button key={x} className={`pill-tab${q === x ? " active" : ""}`} onClick={() => setQ(x)}>{x}</button>)}</div>
        <button className="btn btn-primary btn-sm" onClick={exportIt} disabled={!list.length}>Download deductee details</button>
        <span className="mono" style={{ color: "var(--text3)" }}>Form 140 (old 26Q) due {due} · ₹200/day late fee</span>
      </div>
      {(undeposited.length > 0 || noPan.length > 0) && <div className="callout callout-red">
        {undeposited.length > 0 && <div>• {undeposited.length} entr{undeposited.length > 1 ? "ies" : "y"} not deposited yet. The return cannot be filed until all TDS of the quarter is paid.</div>}
        {noPan.length > 0 && <div>• {noPan.length} deductee(s) without PAN. Collect PAN or make sure 20% was deducted.</div>}
      </div>}
      <div className="table-wrap"><div className="table-scroll"><table>
        <thead><tr><th>Deductee</th><th>PAN</th><th>Section</th><th className="num">Amount</th><th className="num">TDS</th><th>Challan</th></tr></thead>
        <tbody>{list.map(e => <tr key={e.id}><td>{e.party_name}</td><td className="mono">{e.party_pan || "—"}</td><td>{secOf(e.section_key)?.old_section}</td><td className="num">{fmt(e.base_amount)}</td><td className="num">{fmt(e.tds_amount)}</td><td className="mono">{e.challan_no ? `${e.challan_no} · ${fmtDate(e.challan_date)}` : <Badge type="red">pending</Badge>}</td></tr>)}</tbody>
      </table></div>{!list.length && <EmptyState icon="—" message={`No deductions in ${q}`} />}</div>
    </>
  );
}

// Who got paid more than ₹1,00,000 this year (bank debits by counterparty) with no TDS recorded?
function ThresholdCheck({ fy, bankTxns, bankLabels, entries, suppliers }) {
  const rows = useMemo(() => {
    const m = {};
    bankTxns.filter(t => +t.debit > 0 && t.date && fyOfDate(t.date) === fy).forEach(t => {
      const meta = bankLabels[t.key] || {};
      if (["owner", "loan", "tax", "customer"].includes(meta.type)) return;
      const name = (meta.label || t.key || "").trim();
      if (!name) return;
      if (!m[name]) m[name] = { name, type: meta.type || "unlabeled", paid: 0, n: 0, biggest: 0 };
      m[name].paid += +t.debit; m[name].n++; m[name].biggest = Math.max(m[name].biggest, +t.debit);
    });
    const norm = s => s.toLowerCase().replace(/[^a-z0-9]/g, "");
    return Object.values(m).map(r => {
      const tds = entries.filter(e => norm(e.party_name) === norm(r.name)).reduce((a, e) => a + +e.tds_amount, 0);
      const sup = suppliers.find(s => norm(s.name) === norm(r.name));
      return { ...r, tds, flagged: (r.paid > 100000 || r.biggest > 30000) && !tds, sup };
    }).filter(r => r.paid > 30000).sort((a, b) => (b.flagged - a.flagged) || b.paid - a.paid);
  }, [fy, bankTxns, bankLabels, entries, suppliers]);
  return (
    <div className="table-wrap">
      <div className="table-toolbar"><h3>Payments that may need TDS</h3><span className="mono" style={{ color: "var(--text3)" }}>From bank debits this year, by counterparty label</span></div>
      <div className="callout" style={{ margin: 12 }}>A payee is flagged when they got more than ₹1,00,000 in the year, or more than ₹30,000 at once, and no TDS is recorded against the same name. Material purchases (quartz, diesel) usually don't need TDS; contractors, transporters, rent and professionals do. Check each flag and record TDS if needed.</div>
      {!rows.length ? <EmptyState icon="✅" message="No payees above ₹30,000 found" sub="Label counterparties on the Bank Statement screen so names group properly." /> : <div className="table-scroll"><table>
        <thead><tr><th>Payee</th><th>Label type</th><th className="num">Payments</th><th className="num">Largest</th><th className="num">Total paid</th><th className="num">TDS recorded</th><th>Check</th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.name} className={r.flagged ? "due-row-soon" : ""}><td>{r.name}{r.sup?.pan && <div className="mono" style={{ fontSize: 10.5, color: "var(--text3)" }}>{r.sup.pan}</div>}</td><td>{r.type}</td><td className="num">{r.n}</td><td className="num">{fmt(r.biggest)}</td><td className="num">{fmt(r.paid)}</td><td className="num">{r.tds ? fmt(r.tds) : "—"}</td><td>{r.flagged ? <Badge type="amber">Check TDS</Badge> : r.tds ? <Badge type="green">TDS on file</Badge> : <Badge type="muted">Under limit</Badge>}</td></tr>)}</tbody>
      </table></div>}
    </div>
  );
}

function Rates({ sections, role }) {
  const [edit, setEdit] = useState(null);
  async function save() {
    try { const row = await cdb.saveTdsSection({ ...edit, rate_ind: +edit.rate_ind, rate_other: +edit.rate_other, single_limit: edit.single_limit === "" || edit.single_limit == null ? null : +edit.single_limit, annual_limit: edit.annual_limit === "" || edit.annual_limit == null ? null : +edit.annual_limit }); upsertIn("tdsSections", row, "key"); setEdit(null); notify("Saved", "ok"); } catch (e) { notify(e.message, "err"); }
  }
  return (
    <div className="table-wrap">
      <div className="callout callout-blue" style={{ margin: 12 }}>From 1 April 2026 TDS falls under the Income-tax Act, 2025 (Section 393). Rates and limits mostly carried over from the old sections. These values are a starting point: confirm them with your CA. Only an admin can change them.</div>
      <div className="table-scroll"><table>
        <thead><tr><th>Nature</th><th>Old</th><th>2025 Act</th><th>Code</th><th className="num">Indiv/HUF</th><th className="num">Others</th><th className="num">Per payment</th><th className="num">Per year</th><th></th></tr></thead>
        <tbody>{sections.map(s => <tr key={s.key}><td>{s.label}</td><td className="mono">{s.old_section}</td><td className="mono">{s.new_ref}</td><td className="mono">{s.pay_code}</td><td className="num">{s.rate_ind}%</td><td className="num">{s.rate_other}%</td><td className="num">{s.single_limit ? fmt(s.single_limit) : "—"}</td><td className="num">{s.annual_limit ? fmt(s.annual_limit) : "—"}</td><td>{role === "admin" && <button className="btn btn-ghost btn-sm" onClick={() => setEdit({ ...s })}>Edit</button>}</td></tr>)}</tbody>
      </table></div>
      {edit && <Modal title={`Edit: ${edit.label}`} onClose={() => setEdit(null)} foot={<><button className="btn btn-ghost" onClick={() => setEdit(null)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
        <div className="form-row cols-2">{[["new_ref", "2025 Act reference"], ["pay_code", "Payment code"], ["rate_ind", "Rate: individual / HUF %"], ["rate_other", "Rate: others %"], ["single_limit", "Per payment limit ₹"], ["annual_limit", "Per year limit ₹"]].map(([k, l]) =>
          <FG key={k} label={l}><input value={edit[k] ?? ""} onChange={e => setEdit(x => ({ ...x, [k]: e.target.value }))} /></FG>)}</div>
        <label className="check"><input type="checkbox" checked={edit.active} onChange={e => setEdit(x => ({ ...x, active: e.target.checked }))} /> In use</label>
      </Modal>}
    </div>
  );
}

export const reloadTds = () => reload("tds");
