// Tiny shared cache for the compliance modules: each list loads once and every screen sees updates.
import { useEffect, useSyncExternalStore } from "react";
import * as cdb from "./complianceDb.js";

const loaders = {
  settings: cdb.loadSettings,
  customers: cdb.loadCustomers,
  suppliers: cdb.loadSuppliersFull,
  invoices: cdb.loadInvoices,
  receipts: cdb.loadReceipts,
  inward: cdb.loadInward,
  tdsSections: cdb.loadTdsSections,
  tds: cdb.loadTds,
  filings: cdb.loadFilings,
  tasks: cdb.loadTasks,
};

const state = {};      // key -> { data, loading, error, missing }
const subs = new Set();
let snapshot = {};
const emit = () => { snapshot = { ...state }; subs.forEach(f => f()); };

export async function reload(key) {
  state[key] = { ...(state[key] || {}), loading: true };
  emit();
  try {
    const data = await loaders[key]();
    state[key] = { data, loading: false, error: "", missing: false };
  } catch (e) {
    state[key] = { data: state[key]?.data ?? (key === "settings" ? { ...cdb.DEFAULT_SETTINGS } : []), loading: false, error: e.message || String(e), missing: cdb.isMissingTable(e) };
  }
  emit();
}

// Replace a list locally after a save, without another round trip
export function patch(key, fn) {
  state[key] = { ...(state[key] || {}), data: fn(state[key]?.data ?? []) };
  emit();
}
export const upsertIn = (key, row, idKey = "id") => patch(key, xs => {
  const i = xs.findIndex(x => x[idKey] === row[idKey]);
  return i < 0 ? [row, ...xs] : xs.map((x, j) => (j === i ? row : x));
});
export const removeIn = (key, id, idKey = "id") => patch(key, xs => xs.filter(x => x[idKey] !== id));

export function useStore(keys) {
  const snap = useSyncExternalStore(f => { subs.add(f); return () => subs.delete(f); }, () => snapshot);
  useEffect(() => { keys.forEach(k => { if (!state[k]) reload(k); }); }, [keys.join(",")]);
  const out = { loading: false, missing: false, errors: [] };
  keys.forEach(k => {
    const s = snap[k];
    out[k] = s?.data ?? (k === "settings" ? { ...cdb.DEFAULT_SETTINGS } : []);
    if (!s || s.loading && s.data === undefined) out.loading = true;
    if (s?.missing) out.missing = true;
    if (s?.error && !s.missing) out.errors.push(s.error);
  });
  return out;
}
