// Compliance calendar (GST, TDS, LLP/MCA, income tax, payroll, own items), tax settings and supplier tax details.
import React, { useMemo, useState } from "react";
import * as cdb from "./complianceDb.js";
import { useStore, upsertIn, removeIn, reload } from "./finStore.js";
import { useAuth, canWrite } from "./auth.jsx";
import { complianceItems, itemStatus, fyOfDate, monthLabel, STATES, checkGSTIN, checkPAN, panFromGSTIN } from "./taxLogic.js";
import { FilingModal } from "./gst.jsx";
import { fmt, fmtDate, today, notify, Badge, StatCard, FG, Modal, EmptyState, SearchBar, SetupNotice, exportRowsToExcel } from "./ui.jsx";

const SETUP = "Run supabase/002_compliance.sql in the Supabase SQL editor (after 001_security.sql), then reload.";
const CATS = ["GST", "TDS", "Income tax", "LLP / MCA", "Payroll", "Custom"];

// Items due soon / overdue, for the dashboard
export function useDueSoon(days = 15) {
  const st = useStore(["settings", "filings", "tasks"]);
  return useMemo(() => {
    if (st.missing) return { missing: true, items: [] };
    const fy = fyOfDate(today());
    const prev = `${+fy.slice(0, 4) - 1}-${fy.slice(2, 4)}`;
    const all = [...complianceItems(prev, st.settings, st.tasks), ...complianceItems(fy, st.settings, st.tasks)];
    const from = st.settings.track_from || "";
    const items = all.filter(i => !from || i.due >= from).map(i => ({ ...i, status: itemStatus(i, st.filings.find(f => f.key === i.key)) }))
      .filter(i => i.status.code === "overdue" || (i.status.code !== "done" && i.status.days <= days));
    return { missing: false, items };
  }, [st.settings, st.filings, st.tasks, st.missing, days]);
}

export function DueSoonPanel({ onOpen }) {
  const { missing, items } = useDueSoon(15);
  if (missing) return null;
  const overdue = items.filter(i => i.status.code === "overdue");
  return (
    <div className="table-wrap" style={{ marginBottom: 20 }}>
      <div className="table-toolbar">
        <h3>Compliance due <span className="mono" style={{ color: "var(--text3)" }}>· next 15 days</span></h3>
        <button className="btn btn-ghost btn-sm" onClick={onOpen}>Open calendar →</button>
      </div>
      {!items.length ? <EmptyState icon="✅" message="Nothing due in the next 15 days" /> : (
        <div style={{ padding: 12 }}>
          {overdue.length > 0 && <div className="callout callout-red" style={{ marginBottom: 10 }}><b>{overdue.length} overdue.</b> Late fees and interest keep adding up until filed.</div>}
          {items.slice(0, 8).map(i => (
            <div key={i.key} className={`cal-item ${i.status.code}`}>
              <div className="cal-date">{fmtDate(i.due)}</div>
              <div><div className="cal-title">{i.title}</div><div className="cal-sub">{i.authority}</div></div>
              <Badge type={i.status.code === "overdue" ? "red" : "amber"}>{i.status.label}</Badge>
              <span />
            </div>
          ))}
          {items.length > 8 && <div className="mono" style={{ color: "var(--text3)", marginTop: 6 }}>+ {items.length - 8} more</div>}
        </div>
      )}
    </div>
  );
}

export function ComplianceCalendarView() {
  const { role } = useAuth();
  const st = useStore(["settings", "filings", "tasks"]);
  const [fy, setFy] = useState(fyOfDate(today()));
  const [cat, setCat] = useState("");
  const [show, setShow] = useState("open"); // open | all | done
  const [q, setQ] = useState("");
  const [filing, setFiling] = useState(null);
  const [task, setTask] = useState(null);
  const from = st.settings.track_from || "";
  const items = useMemo(() => complianceItems(fy, st.settings, st.tasks).map(i => {
    const f = st.filings.find(x => x.key === i.key);
    const status = !f && from && i.due < from ? { code: "untracked", label: "Before tracking start" } : itemStatus(i, f);
    return { ...i, filing: f, status };
  }), [fy, st.settings, st.tasks, st.filings, from]);
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;

  const shown = items.filter(i => (!cat || i.category === cat) && (!q || i.title.toLowerCase().includes(q.toLowerCase())) &&
    (show === "all" || (show === "done" ? i.status.code === "done" : !["done", "untracked"].includes(i.status.code))));
  async function startTracking() {
    if (!confirm("Treat everything due before today as already handled (by you or your CA)? Only items due from today on will be tracked. You can still mark older items filed by choosing 'Everything'.")) return;
    try { await cdb.saveSettings({ ...st.settings, track_from: today() }); await reload("settings"); notify("Tracking from today", "ok"); } catch (e) { notify(e.message, "err"); }
  }
  const byMonth = {};
  shown.forEach(i => { const m = i.due.slice(0, 7); (byMonth[m] = byMonth[m] || []).push(i); });
  const overdue = items.filter(i => i.status.code === "overdue");
  const soon = items.filter(i => i.status.code === "soon");
  const done = items.filter(i => i.status.code === "done");
  const late = done.filter(i => i.status.late);
  const thisFy = fyOfDate(today());
  const fys = [`${+thisFy.slice(0, 4) - 1}-${thisFy.slice(2, 4)}`, thisFy, `${+thisFy.slice(0, 4) + 1}-${String((+thisFy.slice(0, 4) + 2) % 100).padStart(2, "0")}`];

  return (
    <>
      <div className="stats-grid">
        <StatCard label="Overdue" value={overdue.length} color={overdue.length ? "red" : "green"} onClick={() => setShow("open")} />
        <StatCard label="Due within 7 days" value={soon.length} color={soon.length ? "amber" : "green"} />
        <StatCard label={`Filed in ${fy}`} value={`${done.length} / ${items.length}`} color="blue" onClick={() => setShow("done")} />
        <StatCard label="Filed late" value={late.length} color={late.length ? "amber" : "green"} />
      </div>
      {!from && overdue.length > 3 && <div className="callout callout-blue">
        <b>Just starting with this calendar?</b> It lists every due date of the year, so returns already filed outside the app show as overdue.
        {role === "admin" ? <> <button className="btn btn-primary btn-sm" style={{ marginLeft: 8 }} onClick={startTracking}>Start tracking from today</button></> : " Ask an admin to set a tracking start date."}
      </div>}
      <div className="filter-bar">
        <select className="inline-select" value={fy} onChange={e => setFy(e.target.value)}>{fys.map(f => <option key={f} value={f}>FY {f}</option>)}</select>
        <select className="inline-select" value={cat} onChange={e => setCat(e.target.value)}><option value="">All areas</option>{CATS.map(c => <option key={c}>{c}</option>)}</select>
        <select className="inline-select" value={show} onChange={e => setShow(e.target.value)}><option value="open">Pending</option><option value="done">Filed</option><option value="all">Everything</option></select>
        <SearchBar value={q} onChange={setQ} placeholder="Search…" style={{ minWidth: 180 }} />
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel(`Compliance-${fy}`, items.map(i => ({ Due: fmtDate(i.due), Item: i.title, Area: i.category, Authority: i.authority, Status: i.status.label, "Filed on": fmtDate(i.filing?.done_on), "ARN / Ack": i.filing?.ack_no, "Amount paid": i.filing?.amount_paid, "Filed by": i.filing?.filed_by, Notes: i.filing?.notes })))}>Export</button>
        {canWrite(role) && <button className="btn btn-primary btn-sm" onClick={() => setTask({ recurrence: "yearly", due_day: 1, due_month: 4, remind_days: 15, active: true })}>+ Own item</button>}
      </div>
      {!shown.length ? <EmptyState icon="✅" message={show === "done" ? "Nothing marked filed yet" : "Nothing pending"} /> :
        Object.keys(byMonth).sort().map(m => (
          <div key={m} className="cal-month">
            <h4>{monthLabel(m)}</h4>
            {byMonth[m].map(i => (
              <div key={i.key} className={`cal-item ${i.status.code}`}>
                <div className="cal-date">{fmtDate(i.due)}</div>
                <div>
                  <div className="cal-title">{i.title}</div>
                  <div className="cal-sub">{[i.authority, i.help].filter(Boolean).join(" · ")}{i.filing?.ack_no ? ` · Ack ${i.filing.ack_no}` : ""}{i.filing?.filed_by ? ` · by ${i.filing.filed_by}` : ""}</div>
                </div>
                <Badge type={{ done: i.status.late ? "amber" : "green", overdue: "red", soon: "amber", upcoming: "muted", untracked: "muted" }[i.status.code]}>{i.status.code === "done" ? `${i.status.label} ${fmtDate(i.filing.done_on)}` : i.status.label}</Badge>
                <div style={{ display: "flex", gap: 6 }}>
                  {canWrite(role) && <button className="btn btn-ghost btn-sm" onClick={() => setFiling(i)}>{i.filing ? "Edit" : "Mark done"}</button>}
                  {canWrite(role) && i.key.startsWith("CUSTOM:") && <button className="btn btn-ghost btn-sm" title="Edit this recurring item" onClick={() => setTask(st.tasks.find(t => t.id === +i.key.split(":")[1]))}>⚙</button>}
                </div>
              </div>
            ))}
          </div>
        ))}
      <div className="callout" style={{ marginTop: 10 }}>Due dates are the normal statutory dates. The government often extends them by notification; when that happens, file by the extended date and note it here. Turn items on or off (QRMP, tax audit, PF/ESI, professional tax) in Tax Settings.</div>
      {filing && <FilingModal fkey={filing.key} title={filing.title} due={filing.due} existing={filing.filing} onClose={() => setFiling(null)} />}
      {task && <TaskModal value={task} onClose={() => setTask(null)} />}
    </>
  );
}

function TaskModal({ value, onClose }) {
  const [f, setF] = useState({ ...value });
  const set = k => e => setF(x => ({ ...x, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  async function save() {
    if (!f.title?.trim()) return notify("Title is required", "err");
    try {
      const row = await cdb.saveTask({ ...f, due_day: +f.due_day || 1, due_month: f.recurrence === "yearly" ? +f.due_month || 4 : null, once_date: f.recurrence === "once" ? f.once_date || null : null, remind_days: +f.remind_days || 7 });
      upsertIn("tasks", row); notify("Saved", "ok"); onClose();
    } catch (e) { notify(e.message, "err"); }
  }
  async function del() {
    if (!confirm("Delete this recurring item? Past filed marks stay.")) return;
    try { await cdb.deleteTask(f.id); removeIn("tasks", f.id); onClose(); } catch (e) { notify(e.message, "err"); }
  }
  return (
    <Modal title={f.id ? "Edit own item" : "Add own compliance item"} onClose={onClose}
      foot={<>{f.id && <button className="btn btn-danger" style={{ marginRight: "auto" }} onClick={del}>Delete</button>}<button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
      <FG label="What"><input value={f.title || ""} onChange={set("title")} placeholder="e.g. Factory licence renewal, Fire NOC, Insurance renewal" autoFocus /></FG>
      <div className="form-row cols-2" style={{ marginTop: 12 }}>
        <FG label="Authority"><input value={f.authority || ""} onChange={set("authority")} placeholder="e.g. TGPCB, Factories Dept" /></FG>
        <FG label="Repeats"><select value={f.recurrence} onChange={set("recurrence")}><option value="once">Once</option><option value="monthly">Every month</option><option value="quarterly">Every quarter</option><option value="yearly">Every year</option></select></FG>
      </div>
      <div className="form-row cols-2">
        {f.recurrence === "once" ? <FG label="Date"><input type="date" value={f.once_date || ""} onChange={set("once_date")} /></FG> : <>
          <FG label="Day of month"><input type="number" min="1" max="31" value={f.due_day} onChange={set("due_day")} /></FG>
          {f.recurrence === "yearly" && <FG label="Month"><select value={f.due_month} onChange={set("due_month")}>{Array.from({ length: 12 }, (_, i) => <option key={i + 1} value={i + 1}>{new Date(2000, i, 1).toLocaleString("en-IN", { month: "long" })}</option>)}</select></FG>}
        </>}
      </div>
      <FG label="Notes"><input value={f.notes || ""} onChange={set("notes")} /></FG>
      <label className="check" style={{ marginTop: 10 }}><input type="checkbox" checked={!!f.active} onChange={set("active")} /> Active</label>
    </Modal>
  );
}

// ── Tax settings (admin) ─────────────────────────────────────────────────────
export function TaxSettingsView() {
  const { role } = useAuth();
  const st = useStore(["settings"]);
  const [f, setF] = useState(null);
  const v = f || st.settings;
  if (st.missing) return <SetupNotice title="GST & compliance tables not set up">{SETUP}</SetupNotice>;
  const set = k => e => setF({ ...v, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value });
  const isAdmin = role === "admin";
  async function save() {
    const g = checkGSTIN(v.gstin); if (g) return notify(g, "err");
    try { await cdb.saveSettings({ ...v, default_gst_rate: +v.default_gst_rate }); await reload("settings"); setF(null); notify("Settings saved", "ok"); } catch (e) { notify(e.message, "err"); }
  }
  const Toggle = ({ k, label, note }) => <label className="check" style={{ marginBottom: 10, alignItems: "flex-start" }}><input type="checkbox" disabled={!isAdmin} checked={!!v[k]} onChange={set(k)} /><span><b style={{ color: "var(--text)" }}>{label}</b>{note && <div style={{ fontSize: 11, color: "var(--text3)" }}>{note}</div>}</span></label>;
  return (
    <fieldset disabled={!isAdmin} style={{ border: 0, padding: 0, margin: 0, minWidth: 0 }}>
      {!isAdmin && <div className="callout">Only an admin can change these settings.</div>}
      <div className="config-card"><h3>Business details (printed on invoices)</h3>
        <div className="form-row cols-3">
          <FG label="Legal name"><input value={v.legal_name} onChange={set("legal_name")} /></FG>
          <FG label="GSTIN" note={checkGSTIN(v.gstin)}><input value={v.gstin} onChange={e => setF({ ...v, gstin: e.target.value.toUpperCase(), pan: panFromGSTIN(e.target.value) || v.pan, state_code: e.target.value.slice(0, 2) })} /></FG>
          <FG label="State"><select value={v.state_code} onChange={e => setF({ ...v, state_code: e.target.value, state: STATES[e.target.value] })}>{Object.entries(STATES).map(([c, n]) => <option key={c} value={c}>{n} ({c})</option>)}</select></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="PAN" note={checkPAN(v.pan)}><input value={v.pan} onChange={set("pan")} /></FG>
          <FG label="TAN (for TDS)" note={!v.tan ? "Needed for TDS challans and returns" : ""}><input value={v.tan} maxLength={10} onChange={e => setF({ ...v, tan: e.target.value.toUpperCase() })} /></FG>
          <FG label="LLPIN"><input value={v.llpin} onChange={set("llpin")} /></FG>
        </div>
        <div className="form-row cols-2">
          <FG label="Plant / registered address (on invoices)"><textarea rows={2} value={v.address} onChange={set("address")} /></FG>
          <FG label="Office address"><textarea rows={2} value={v.office_address} onChange={set("office_address")} /></FG>
        </div>
        <div className="form-row cols-4">
          <FG label="Phone"><input value={v.phone} onChange={set("phone")} /></FG>
          <FG label="Email"><input value={v.email} onChange={set("email")} /></FG>
          <FG label="Signatory line"><input value={v.signatory} onChange={set("signatory")} /></FG>
          <FG label="Invoice prefix" note="NGR → NGR/26-27/0001"><input value={v.invoice_prefix} maxLength={5} onChange={e => setF({ ...v, invoice_prefix: e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, "") })} /></FG>
        </div>
        <div className="form-row cols-4">
          <FG label="Bank"><input value={v.bank_name} onChange={set("bank_name")} /></FG>
          <FG label="Account no."><input value={v.bank_account} onChange={set("bank_account")} /></FG>
          <FG label="IFSC"><input value={v.bank_ifsc} onChange={set("bank_ifsc")} /></FG>
          <FG label="Branch"><input value={v.bank_branch} onChange={set("bank_branch")} /></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Default HSN" note="Quartz powder: 25062020"><input value={v.default_hsn} onChange={set("default_hsn")} /></FG>
          <FG label="Default GST rate %"><select value={v.default_gst_rate} onChange={set("default_gst_rate")}>{[0, 5, 12, 18, 28].map(r => <option key={r}>{r}</option>)}</select></FG>
          <FG label="Default unit"><select value={v.default_unit} onChange={set("default_unit")}>{["MTS", "KGS", "BAG", "NOS"].map(u => <option key={u}>{u}</option>)}</select></FG>
        </div>
        <FG label="Terms printed on invoices"><textarea rows={3} value={v.invoice_terms} onChange={set("invoice_terms")} /></FG>
      </div>
      <div className="config-card"><h3>What applies to us (drives the compliance calendar)</h3>
        <FG label="GST return filing"><select value={v.gst_frequency} onChange={set("gst_frequency")} style={{ maxWidth: 360 }}><option value="monthly">Monthly (GSTR-1 by 11th, GSTR-3B by 20th)</option><option value="qrmp">Quarterly QRMP (turnover up to ₹5 Cr)</option></select></FG>
        <div style={{ height: 12 }} />
        <FG label="Track compliance from" note="Due dates before this are treated as handled outside the app"><input type="date" value={v.track_from || ""} onChange={set("track_from")} style={{ maxWidth: 200 }} /></FG>
        <div style={{ height: 12 }} />
        <Toggle k="einvoice" label="E-invoicing applies" note="Mandatory once turnover in any year since 2017-18 crossed ₹5 Cr. Invoices to GST-registered buyers then need an IRN." />
        <Toggle k="tax_audit" label="Tax audit applies" note="Moves the ITR due date to 31 October and adds the audit report (30 September)." />
        <Toggle k="buyer_194q" label="Our turnover last year was over ₹10 Cr" note="We must deduct 0.1% TDS on purchases of goods above ₹50 lakh from one seller in the year." />
        <Toggle k="pf_esi" label="PF / ESI registered" note="Adds monthly PF ECR and ESI payments (15th)." />
        <Toggle k="professional_tax" label="Professional tax (Telangana)" note="Adds the monthly PT payment (10th)." />
      </div>
      {isAdmin && <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>{f && <button className="btn btn-ghost" onClick={() => setF(null)}>Discard</button>}<button className="btn btn-primary" disabled={!f} onClick={save}>Save settings</button></div>}
    </fieldset>
  );
}

// ── Suppliers with tax details (replaces the simple master) ─────────────────
export function SuppliersFullView({ onChanged }) {
  const { role } = useAuth();
  const st = useStore(["suppliers", "tdsSections"]);
  const [q, setQ] = useState("");
  const [edit, setEdit] = useState(null);
  const rows = st.suppliers.filter(s => !q || `${s.name} ${s.gstin} ${s.pan}`.toLowerCase().includes(q.toLowerCase()));
  const missing = st.suppliers.filter(s => !s.pan).length;
  async function save() {
    if (!edit.name?.trim()) return notify("Name is required", "err");
    const g = checkGSTIN(edit.gstin); if (g) return notify(g, "err");
    const p = checkPAN(edit.pan); if (p) return notify(p, "err");
    try { const row = await cdb.saveSupplierFull(edit); upsertIn("suppliers", row); onChanged?.(edit.id ? st.suppliers.map(x => (x.id === row.id ? row : x)) : [...st.suppliers, row]); setEdit(null); notify("Supplier saved", "ok"); } catch (e) { notify(e.message, "err"); }
  }
  const set = k => e => setEdit(x => ({ ...x, [k]: e.target.value }));
  return (
    <>
      {missing > 0 && !st.missing && <div className="callout callout-amber">{missing} supplier(s) have no PAN. TDS on them must be deducted at 20%, and the TDS return needs PANs.</div>}
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Suppliers & vendors <span className="mono" style={{ color: "var(--text3)" }}>· {st.suppliers.length}</span></h3>
          <div style={{ display: "flex", gap: 8 }}><SearchBar value={q} onChange={setQ} placeholder="Name, GSTIN, PAN…" />
            <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("Suppliers", st.suppliers.map(s => ({ Name: s.name, GSTIN: s.gstin, PAN: s.pan, Type: s.entity_type, "TDS nature": s.tds_section, "MSME Udyam": s.msme_udyam, Contact: s.contact, Phone: s.phone, Email: s.email, Address: s.address })))}>Export</button>
            {canWrite(role) && <button className="btn btn-primary btn-sm" onClick={() => setEdit({ entity_type: "company" })}>+ Supplier</button>}</div></div>
        {!rows.length ? <EmptyState icon="⛏" message={st.loading ? "Loading…" : "No suppliers"} /> : <div className="table-scroll"><table>
          <thead><tr><th>Name</th><th>GSTIN</th><th>PAN</th><th>Type</th><th>TDS</th><th>Contact</th><th></th></tr></thead>
          <tbody>{rows.map(s => <tr key={s.id}>
            <td><b>{s.name}</b>{s.msme_udyam && <span className="chip" style={{ marginLeft: 6 }}>MSME</span>}</td>
            <td className="mono">{s.gstin || "—"}</td><td className="mono" style={{ color: s.pan ? undefined : "var(--red)" }}>{s.pan || "missing"}</td>
            <td>{s.entity_type || "—"}</td><td>{st.tdsSections.find(x => x.key === s.tds_section)?.old_section || "—"}</td>
            <td>{[s.contact, s.phone].filter(Boolean).join(" · ")}</td>
            <td className="r">{canWrite(role) && <button className="btn btn-ghost btn-sm" onClick={() => setEdit({ ...s })}>Edit</button>}</td>
          </tr>)}</tbody></table></div>}
      </div>
      {edit && <Modal title={edit.id ? `Edit ${edit.name}` : "New supplier"} onClose={() => setEdit(null)} size="modal-lg" foot={<><button className="btn btn-ghost" onClick={() => setEdit(null)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
        <div className="form-row cols-3">
          <FG label="Name *"><input value={edit.name || ""} onChange={set("name")} autoFocus /></FG>
          <FG label="GSTIN" note={checkGSTIN(edit.gstin)}><input maxLength={15} value={edit.gstin || ""} onChange={e => { const g = e.target.value.toUpperCase(); setEdit(x => ({ ...x, gstin: g, ...(g.length === 15 && !checkGSTIN(g) ? { pan: panFromGSTIN(g), state_code: g.slice(0, 2) } : {}) })); }} /></FG>
          <FG label="PAN" note={checkPAN(edit.pan)}><input maxLength={10} value={edit.pan || ""} onChange={e => setEdit(x => ({ ...x, pan: e.target.value.toUpperCase() }))} /></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Payee type (for TDS rate)"><select value={edit.entity_type || "company"} onChange={set("entity_type")}><option value="individual">Individual / proprietor</option><option value="huf">HUF</option><option value="company">Company</option><option value="firm">Firm / LLP</option><option value="other">Other</option></select></FG>
          <FG label="Usual TDS nature"><select value={edit.tds_section || ""} onChange={set("tds_section")}><option value="">None (e.g. material supplier)</option>{st.tdsSections.map(x => <option key={x.key} value={x.key}>{x.label} ({x.old_section})</option>)}</select></FG>
          <FG label="MSME Udyam no."><input value={edit.msme_udyam || ""} onChange={set("msme_udyam")} placeholder="UDYAM-TS-…" /></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Contact"><input value={edit.contact || ""} onChange={set("contact")} /></FG>
          <FG label="Phone"><input value={edit.phone || ""} onChange={set("phone")} /></FG>
          <FG label="Email"><input value={edit.email || ""} onChange={set("email")} /></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Address" span={2}><textarea rows={2} value={edit.address || ""} onChange={set("address")} /></FG>
          <div><FG label="Lower TDS certificate no."><input value={edit.lower_tds_cert || ""} onChange={set("lower_tds_cert")} /></FG>
            <FG label="Certificate rate %"><input type="number" step="0.01" value={edit.lower_tds_rate ?? ""} onChange={set("lower_tds_rate")} /></FG></div>
        </div>
      </Modal>}
    </>
  );
}
