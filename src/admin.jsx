// Admin screens: approve users / change roles, and the read-only audit trail.
import React, { useEffect, useMemo, useState } from "react";
import { supabase } from "./supabase.js";
import { useAuth, ROLE_LABELS } from "./auth.jsx";
import { Badge, EmptyState, SearchBar, notify, exportRowsToExcel, fmtDate } from "./ui.jsx";

const ROLE_BADGE = { admin: "accent", accountant: "blue", viewer: "muted", pending: "amber" };

export function UsersView() {
  const { session } = useAuth();
  const [rows, setRows] = useState(null);
  const load = () => supabase.from("fin_users").select("*").order("created_at").then(({ data, error }) => {
    if (error) notify(error.message, "err");
    setRows(data || []);
  });
  useEffect(() => { load(); }, []);

  async function patch(u, change) {
    const { error } = await supabase.from("fin_users").update(change).eq("user_id", u.user_id);
    if (error) return notify(error.message, "err");
    notify("Saved", "ok");
    load();
  }

  if (!rows) return <div className="mono" style={{ color: "var(--text3)" }}>Loading…</div>;
  const pending = rows.filter(r => r.role === "pending" && r.active);

  return (
    <>
      {pending.length > 0 && (
        <div className="callout callout-amber">
          <b>{pending.length} account{pending.length > 1 ? "s" : ""} waiting for approval.</b> Give each person the lowest role they need:
          <i> Viewer</i> sees everything but cannot change anything; <i>Accountant</i> can add and edit entries; <i>Admin</i> also manages users and sees the audit log.
        </div>
      )}
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Users</h3><span className="mono" style={{ color: "var(--text3)" }}>New people sign up from the login page, then appear here as Pending.</span></div>
        <div className="table-scroll">
          <table>
            <thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th>Joined</th><th className="r">Actions</th></tr></thead>
            <tbody>
              {rows.map(u => {
                const me = u.user_id === session.user.id;
                return (
                  <tr key={u.user_id}>
                    <td>{u.name}{me && <span className="mono" style={{ color: "var(--text3)" }}> (you)</span>}</td>
                    <td className="mono">{u.email}</td>
                    <td>
                      <select className="inline-select" value={u.role} disabled={me} onChange={e => patch(u, { role: e.target.value })}>
                        {Object.keys(ROLE_LABELS).map(r => <option key={r} value={r}>{ROLE_LABELS[r]}</option>)}
                      </select>
                    </td>
                    <td>{u.active ? <Badge type={ROLE_BADGE[u.role]}>{u.role === "pending" ? "Pending" : "Active"}</Badge> : <Badge type="red">Blocked</Badge>}</td>
                    <td className="mono">{fmtDate(u.created_at)}</td>
                    <td className="r">
                      {!me && (u.active
                        ? <button className="btn btn-danger btn-sm" onClick={() => confirm(`Block ${u.email}? They will be signed out of the data immediately.`) && patch(u, { active: false })}>Block</button>
                        : <button className="btn btn-ghost btn-sm" onClick={() => patch(u, { active: true })}>Unblock</button>)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );
}

const TABLE_NAMES = {
  bank_txns: "Bank entry", bank_labels: "Bank label", bank_import_batches: "Bank import", expenses: "Expense",
  purchases: "Purchase", capital_items: "Capital item", partner_cashbook: "Partner entry", loan_entries: "Loan",
  vendor_projects: "Vendor project", vendor_project_payments: "Project payment", vendor_project_estimates: "Project estimate",
  planned_expenses: "Planned expense", purchase_orders: "Purchase order", suppliers: "Supplier", attachments: "Attachment",
  sales_invoices: "Sales invoice", customers: "Customer", tds_entries: "TDS entry", compliance_tasks: "Compliance task",
  compliance_filings: "Filing", company_settings: "Settings",
};

function summarize(r) {
  const d = r.new_data || r.old_data || {};
  const pick = ["invoice_no", "po_number", "name", "description", "particulars", "label", "vendor_name", "lender", "partner_name", "title", "party_name"]
    .map(k => d[k]).find(Boolean);
  const amt = ["grand_total", "total_amount", "amount", "debit", "credit", "tds_amount"].map(k => d[k]).find(v => v != null && +v);
  return [pick, amt != null ? `₹${(+amt).toLocaleString("en-IN")}` : ""].filter(Boolean).join(" · ");
}

function changedFields(r) {
  if (r.action !== "UPDATE" || !r.old_data || !r.new_data) return [];
  return Object.keys(r.new_data).filter(k => !["updated_at"].includes(k) && JSON.stringify(r.old_data[k]) !== JSON.stringify(r.new_data[k]))
    .map(k => ({ k, from: r.old_data[k], to: r.new_data[k] }));
}

export function AuditView() {
  const [rows, setRows] = useState(null);
  const [q, setQ] = useState("");
  const [action, setAction] = useState("");
  const [days, setDays] = useState(30);
  const [open, setOpen] = useState(null);

  useEffect(() => {
    const since = new Date(Date.now() - days * 864e5).toISOString();
    supabase.from("fin_audit_log").select("*").gte("at", since).order("at", { ascending: false }).limit(2000)
      .then(({ data, error }) => { if (error) notify(error.message, "err"); setRows(data || []); });
  }, [days]);

  const shown = useMemo(() => (rows || []).filter(r =>
    (!action || r.action === action) &&
    (!q || `${r.user_email} ${r.table_name} ${summarize(r)}`.toLowerCase().includes(q.toLowerCase()))
  ), [rows, q, action]);

  return (
    <div className="table-wrap">
      <div className="table-toolbar">
        <h3>Audit trail <span className="mono" style={{ color: "var(--text3)", fontWeight: 400 }}>· every add, change and delete, cannot be edited</span></h3>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <SearchBar value={q} onChange={setQ} placeholder="Search user, record…" style={{ minWidth: 200 }} />
          <select className="inline-select" value={action} onChange={e => setAction(e.target.value)}>
            <option value="">All actions</option><option value="INSERT">Added</option><option value="UPDATE">Changed</option><option value="DELETE">Deleted</option>
          </select>
          <select className="inline-select" value={days} onChange={e => setDays(+e.target.value)}>
            {[7, 30, 90, 365, 3650].map(d => <option key={d} value={d}>{d >= 3650 ? "All time" : `Last ${d} days`}</option>)}
          </select>
          <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("Audit-trail", shown.map(r => ({
            When: new Date(r.at).toLocaleString("en-IN"), User: r.user_email || "system", Action: r.action,
            Record: TABLE_NAMES[r.table_name] || r.table_name, RowId: r.row_id, Summary: summarize(r),
            Changes: changedFields(r).map(c => `${c.k}: ${c.from} → ${c.to}`).join("; "),
          })))}>Export</button>
        </div>
      </div>
      {!rows ? <div className="empty mono">Loading…</div> : !shown.length ? <EmptyState icon="🛡" message="No activity in this period" /> : (
        <div className="table-scroll">
          <table>
            <thead><tr><th>When</th><th>User</th><th>Action</th><th>Record</th><th>Details</th></tr></thead>
            <tbody>
              {shown.map(r => {
                const ch = changedFields(r);
                return (
                  <React.Fragment key={r.id}>
                    <tr onClick={() => setOpen(open === r.id ? null : r.id)} style={{ cursor: "pointer" }}>
                      <td className="mono">{new Date(r.at).toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" })}</td>
                      <td className="mono">{r.user_email || "system"}</td>
                      <td><Badge type={r.action === "DELETE" ? "red" : r.action === "UPDATE" ? "amber" : "green"}>{r.action === "INSERT" ? "Added" : r.action === "UPDATE" ? "Changed" : "Deleted"}</Badge></td>
                      <td>{TABLE_NAMES[r.table_name] || r.table_name} <span className="mono" style={{ color: "var(--text3)" }}>#{r.row_id}</span></td>
                      <td style={{ color: "var(--text2)" }}>{summarize(r)}{ch.length > 0 && <span className="mono" style={{ color: "var(--text3)" }}> · {ch.length} field{ch.length > 1 ? "s" : ""}</span>}</td>
                    </tr>
                    {open === r.id && (
                      <tr><td colSpan={5} style={{ background: "var(--bg3)" }}>
                        {ch.length ? ch.map(c => (
                          <div key={c.k} className="mono" style={{ padding: "2px 0" }}>
                            <span style={{ color: "var(--text3)" }}>{c.k}:</span> <span style={{ color: "var(--red)" }}>{String(c.from ?? "—")}</span> → <span style={{ color: "var(--green)" }}>{String(c.to ?? "—")}</span>
                          </div>
                        )) : <pre className="sql-pre" style={{ maxHeight: 220 }}>{JSON.stringify(r.new_data || r.old_data, null, 2)}</pre>}
                      </td></tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
