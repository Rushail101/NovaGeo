// Shared building blocks used across the finance app (moved out of App.jsx so new modules can use them).
import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

// ── Formatting ────────────────────────────────────────────────────────────────
export const fmtMT = v => `${(+v || 0).toFixed(3)} MT`;
export const fmt = n => `₹${(+n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
export const fmt0 = n => `₹${Math.round(+n || 0).toLocaleString("en-IN")}`;
// Local (IST) date as YYYY-MM-DD. The old version used UTC, which gave yesterday's date before 5:30 am.
export const isoDate = d => {
  const x = d instanceof Date ? d : new Date(d);
  return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
};
export const today = () => isoDate(new Date());
export const fmtDate = s => {
  if (!s) return "";
  const [y, m, d] = String(s).slice(0, 10).split("-");
  return d && m && y ? `${d}-${m}-${y}` : s;
};
export const round2 = n => Math.round((+n || 0) * 100) / 100;

// Indian financial year helpers ("2026-27" runs 1 Apr 2026 – 31 Mar 2027)
export function fyOf(dateStr) {
  const d = new Date(dateStr || today());
  const y = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
  return `${y}-${String((y + 1) % 100).padStart(2, "0")}`;
}
export function fyRange(fy) {
  const y = +fy.slice(0, 4);
  return { from: `${y}-04-01`, to: `${y + 1}-03-31` };
}
export const quarterOf = dateStr => {
  const m = new Date(dateStr).getMonth(); // 0 = Jan
  return m >= 3 && m <= 5 ? "Q1" : m >= 6 && m <= 8 ? "Q2" : m >= 9 && m <= 11 ? "Q3" : "Q4";
};

// Amount in words (Indian system) for invoices
export function amountInWords(n) {
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
    "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  const two = x => (x < 20 ? ones[x] : `${tens[Math.floor(x / 10)]}${x % 10 ? " " + ones[x % 10] : ""}`);
  const three = x => `${x >= 100 ? ones[Math.floor(x / 100)] + " Hundred" + (x % 100 ? " " : "") : ""}${x % 100 ? two(x % 100) : ""}`;
  const words = x => {
    if (x === 0) return "Zero";
    const parts = [];
    const cr = Math.floor(x / 1e7); x %= 1e7;
    const lk = Math.floor(x / 1e5); x %= 1e5;
    const th = Math.floor(x / 1e3); x %= 1e3;
    if (cr) parts.push(`${words(cr)} Crore`);
    if (lk) parts.push(`${two(lk)} Lakh`);
    if (th) parts.push(`${two(th)} Thousand`);
    if (x) parts.push(three(x));
    return parts.join(" ");
  };
  const rupees = Math.floor(Math.abs(+n || 0));
  const paise = Math.round((Math.abs(+n || 0) - rupees) * 100);
  return `Rupees ${words(rupees)}${paise ? ` and ${two(paise)} Paise` : ""} Only`;
}

// ── Toasts (replace blocking alert() pop-ups) ────────────────────────────────
const toastSubs = new Set();
let toastId = 0;
export function notify(message, kind) {
  const text = String(message ?? "");
  const k = kind || (/fail|error|could not|cannot|can't|invalid|required|must|not /i.test(text) ? "err" : "ok");
  const t = { id: ++toastId, text, kind: k };
  toastSubs.forEach(f => f(t));
}
export function Toaster() {
  const [items, setItems] = useState([]);
  useEffect(() => {
    const on = t => {
      setItems(xs => [...xs, t]);
      setTimeout(() => setItems(xs => xs.filter(x => x.id !== t.id)), t.kind === "err" ? 7000 : 3500);
    };
    toastSubs.add(on);
    return () => toastSubs.delete(on);
  }, []);
  return (
    <div className="toasts">
      {items.map(t => (
        <div key={t.id} className={`toast toast-${t.kind}`} onClick={() => setItems(xs => xs.filter(x => x.id !== t.id))}>{t.text}</div>
      ))}
    </div>
  );
}

// ── Components ────────────────────────────────────────────────────────────────
export function BadgeComponent({ type, children }) { return <span className={`badge badge-${type || "muted"}`}>{children}</span>; }
export const Badge = BadgeComponent;

export function StatCard({ label, value, sub, color = "blue", onClick }) {
  return (
    <div className={`stat-card ${color}${onClick ? " clickable" : ""}`} onClick={onClick}>
      <div className="stat-label">{label}</div>
      <div className={`stat-val ${color}`}>{value}</div>
      {sub && <div className="stat-sub">{sub}</div>}
    </div>
  );
}

export function FG({ label, note, children, span }) {
  return <div className="form-group" style={span ? { gridColumn: `span ${span}` } : {}}>{label && <label>{label}</label>}{children}{note && <div className="form-note">{note}</div>}</div>;
}

export function Modal({ title, onClose, children, foot, size = "" }) {
  useEffect(() => {
    const k = e => { if (e.key === "Escape") onClose?.(); };
    window.addEventListener("keydown", k);
    return () => window.removeEventListener("keydown", k);
  }, [onClose]);
  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className={`modal ${size}`} role="dialog" aria-label={typeof title === "string" ? title : undefined}>
        <div className="modal-head"><h3>{title}</h3><button className="btn btn-ghost btn-sm" onClick={onClose} aria-label="Close">✕</button></div>
        <div className="modal-body">{children}</div>
        {foot && <div className="modal-foot">{foot}</div>}
      </div>
    </div>
  );
}

export function EmptyState({ icon, message, sub, action }) {
  return (
    <div className="empty">
      <div className="empty-icon">{icon || "📭"}</div>
      <p style={{ fontWeight: 600, marginBottom: 4 }}>{message}</p>
      {sub && <p style={{ fontSize: 11.5, color: "var(--text3)", marginTop: 4 }}>{sub}</p>}
      {action && <div style={{ marginTop: 14 }}>{action}</div>}
    </div>
  );
}

export function SearchBar({ value, onChange, placeholder, style }) {
  return (
    <input
      className="search-input"
      type="search"
      placeholder={placeholder}
      value={value}
      onChange={e => onChange(e.target.value)}
      style={style}
    />
  );
}

export function RowActions({ onEdit, onDelete }) {
  if (!onEdit && !onDelete) return null;
  return (
    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
      {onEdit && <button className="btn btn-ghost btn-sm" onClick={onEdit}>Edit</button>}
      {onDelete && <button className="btn btn-danger btn-sm" onClick={onDelete}>Delete</button>}
    </div>
  );
}

export function SetupNotice({ title, sql, children }) {
  const [copied, setCopied] = useState(false);
  return (
    <div className="config-card setup-card">
      <h3>⚙ {title}</h3>
      <div style={{ fontSize: 12.5, color: "var(--text2)", lineHeight: 1.6, marginBottom: 10 }}>{children}</div>
      {sql && (
        <>
          <button className="btn btn-ghost btn-sm" onClick={() => { navigator.clipboard?.writeText(sql); setCopied(true); }}>{copied ? "Copied ✓" : "Copy SQL"}</button>
          <pre className="sql-pre">{sql}</pre>
        </>
      )}
    </div>
  );
}

export function exportRowsToExcel(fileName, rows, sheetName = "Sheet1") {
  if (!rows || !rows.length) { notify("Nothing to export yet.", "info"); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`);
}

// Several sheets in one workbook: sheets = [{ name, rows }]
export function exportSheetsToExcel(fileName, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(s => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(s.rows.length ? s.rows : [{ Note: "No entries" }]), s.name.slice(0, 31)));
  XLSX.writeFile(wb, fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`);
}
