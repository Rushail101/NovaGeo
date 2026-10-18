import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import * as db from "./db.js";
import * as vpdb from "./vendorProjectsDb.js";
import * as loansdb from "./loansDb.js";
import { exportStatementPDF, exportVendorSummaryPDF } from "./pdfExport.js";

// ── Design System ─────────────────────────────────────────────────────────────
const FONTS = `@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');`;
const CSS = `
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0d0d0d;--bg2:#161616;--bg3:#1e1e1e;--bg4:#252525;
  --border:#2a2a2a;--border2:#333;--border3:#3d3d3d;
  --text:#e8e6df;--text2:#999;--text3:#555;--text4:#333;
  --accent:#c8f064;--accent2:#a8d844;--accent3:#e8ff8a;
  --red:#ff5c5c;--amber:#f5a623;--blue:#5ca0ff;
  --green:#4ecb71;--purple:#a78bfa;--teal:#2dd4bf;
  --font:'Syne',sans-serif;--mono:'DM Mono',monospace;
  --r:8px;--r2:12px;--r3:16px;
  --shadow:0 2px 12px rgba(0,0,0,.4);--shadow2:0 4px 24px rgba(0,0,0,.6);
}
body{background:var(--bg);color:var(--text);font-family:var(--font);font-size:14px;min-height:100vh;-webkit-font-smoothing:antialiased}
.app{display:flex;min-height:100vh}

/* ── Mobile Sidebar Drawer Fix ── */
.sidebar {
  width: 220px;
  background: var(--bg2);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  position: sticky;
  top: 0;
  height: 100vh;
  flex-shrink: 0;
  overflow: hidden;
  transition: transform 0.25s ease-in-out;
  z-index: 300;
}
.sidebar-logo{padding:20px 18px 14px;border-bottom:1px solid var(--border)}
.sidebar-logo h1{font-size:17px;font-weight:800;letter-spacing:-.4px;color:var(--accent)}
.sidebar-logo p{font-size:10px;color:var(--text3);margin-top:2px;letter-spacing:.1em;text-transform:uppercase;font-family:var(--mono)}
.nav{flex:1;padding:6px;overflow-y:auto}
.nav-label{font-size:9px;color:var(--text4);letter-spacing:.1em;text-transform:uppercase;font-family:var(--mono);padding:10px 10px 4px;margin-top:4px}
.nav-item{display:flex;align-items:center;gap:9px;padding:7px 10px;border-radius:var(--r);cursor:pointer;color:var(--text2);font-size:12.5px;font-weight:500;transition:all .12s;border:none;background:none;width:100%;text-align:left;font-family:var(--font)}
.nav-item:hover{background:var(--bg3);color:var(--text)}
.nav-item.active{background:var(--accent);color:#0d0d0d;font-weight:600}
.nav-item .icon{width:15px;text-align:center;font-size:12px;flex-shrink:0}
.sidebar-foot{padding:10px 12px;border-top:1px solid var(--border);font-size:10px;color:var(--text4);font-family:var(--mono)}
.main{flex:1;display:flex;flex-direction:column;min-width:0;overflow:hidden}
.topbar{display:flex;align-items:center;justify-content:space-between;padding:14px 26px;border-bottom:1px solid var(--border);background:var(--bg);flex-shrink:0}
.topbar h2{font-size:17px;font-weight:700;letter-spacing:-.3px}
.topbar-right{display:flex;gap:8px;align-items:center}
.content{padding:22px 26px;flex:1;overflow-y:auto}
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border-radius:var(--r);font-family:var(--font);font-size:12.5px;font-weight:500;cursor:pointer;border:none;transition:all .12s;white-space:nowrap}
.btn-primary{background:var(--accent);color:#0d0d0d}.btn-primary:hover{background:var(--accent2)}
.btn-ghost{background:transparent;color:var(--text2);border:1px solid var(--border2)}.btn-ghost:hover{color:var(--text);border-color:var(--border3)}
.btn-danger{background:transparent;color:var(--red);border:1px solid #3a1a1a}.btn-danger:hover{background:#1e0a0a}
.btn-info{background:transparent;color:var(--blue);border:1px solid #1a2e4a}.btn-info:hover{background:#0a1a2e}
.btn-sm{padding:4px 10px;font-size:11.5px}
.btn:disabled{opacity:.45;cursor:not-allowed}
.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:20px}
.stats-grid-5{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:20px}
.stat-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:14px 18px;position:relative;overflow:hidden}
.stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px}
.stat-card.green::before{background:var(--green)}.stat-card.amber::before{background:var(--amber)}
.stat-card.red::before{background:var(--red)}.stat-card.blue::before{background:var(--blue)}
.stat-card.purple::before{background:var(--purple)}.stat-card.teal::before{background:var(--teal)}
.stat-card.accent::before{background:var(--accent)}
.stat-label{font-size:10px;color:var(--text3);letter-spacing:.07em;text-transform:uppercase;font-family:var(--mono);margin-bottom:7px}
.stat-val{font-size:22px;font-weight:700;letter-spacing:-.4px}
.stat-val.green{color:var(--green)}.stat-val.amber{color:var(--amber)}.stat-val.red{color:var(--red)}
.stat-val.blue{color:var(--blue)}.stat-val.purple{color:var(--purple)}.stat-val.teal{color:var(--teal)}
.stat-val.accent{color:var(--accent)}.stat-val.white{color:var(--text)}
.stat-sub{font-size:10px;color:var(--text3);margin-top:3px;font-family:var(--mono)}
.table-wrap{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);overflow:hidden;margin-bottom:20px}
.table-toolbar{display:flex;align-items:center;justify-content:space-between;padding:12px 16px;border-bottom:1px solid var(--border);gap:12px;flex-wrap:wrap}
.table-toolbar h3{font-size:13px;font-weight:600}
table{width:100%;border-collapse:collapse}
th{text-align:left;padding:9px 16px;font-size:10px;color:var(--text3);letter-spacing:.08em;text-transform:uppercase;font-family:var(--mono);border-bottom:1px solid var(--border);font-weight:500;white-space:nowrap}
th.r,td.r{text-align:right}
td{padding:10px 16px;font-size:12.5px;border-bottom:1px solid var(--border);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:rgba(255,255,255,.02)}
.mono{font-family:var(--mono);font-size:11.5px}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:500;font-family:var(--mono);text-transform:uppercase;letter-spacing:.04em;white-space:nowrap}
.badge-green,.badge-pass,.badge-running,.badge-sale{background:#0a2016;color:var(--green);border:1px solid #1a4028}
.badge-amber,.badge-conditional,.badge-low,.badge-morning{background:#1a1500;color:var(--amber);border:1px solid #3a2e00}
.badge-red,.badge-fail,.badge-rejected,.badge-negative{background:#200808;color:var(--red);border:1px solid #401010}
.badge-blue,.badge-under-test,.badge-inward{background:#081828;color:var(--blue);border:1px solid #143050}
.badge-muted,.badge-disposed{background:var(--bg3);color:var(--text3);border:1px solid var(--border)}
.badge-teal,.badge-released,.badge-retained{background:#051e1a;color:var(--teal);border:1px solid #0d3830}
.badge-purple,.badge-dispatched{background:#100a28;color:var(--purple);border:1px solid #201848}
.badge-accent{background:#0d1a00;color:var(--accent);border:1px solid #1e3a00}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.8);display:flex;align-items:center;justify-content:center;z-index:200;padding:20px;backdrop-filter:blur(2px)}
.modal{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--r3);width:100%;max-width:600px;max-height:92vh;display:flex;flex-direction:column;box-shadow:var(--shadow2)}
.modal-lg{max-width:820px}
.modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 22px;border-bottom:1px solid var(--border);flex-shrink:0}
.modal-head h3{font-size:15px;font-weight:600}
.modal-body{padding:20px 22px;overflow-y:auto;flex:1}
.modal-foot{display:flex;justify-content:flex-end;gap:8px;padding:14px 22px;border-top:1px solid var(--border);flex-shrink:0}
.form-row{display:grid;gap:14px;margin-bottom:14px}
.form-row.cols-2{grid-template-columns:1fr 1fr}
.form-row.cols-3{grid-template-columns:1fr 1fr 1fr}
.form-group{display:flex;flex-direction:column;gap:5px}
.form-group label{font-size:10.5px;color:var(--text2);letter-spacing:.05em;text-transform:uppercase;font-family:var(--mono);font-weight:500}
.form-group input,.form-group select,.form-group textarea{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:8px 11px;color:var(--text);font-family:var(--font);font-size:13px;outline:none;width:100%}
.form-group input:focus,.form-group select:focus{border-color:var(--accent)}
.form-note{font-size:10.5px;color:var(--text3);font-family:var(--mono);margin-top:2px}
.empty{text-align:center;padding:48px 24px;color:var(--text3)}
.empty-icon{font-size:32px;margin-bottom:12px}
.filter-bar{display:flex;gap:10px;margin-bottom:16px;align-items:center;flex-wrap:wrap}
.config-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;margin-bottom:16px}
.config-card h3{font-size:13px;font-weight:600;margin-bottom:14px}
.config-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.cost-row{display:flex;justify-content:space-between;align-items:center;padding:9px 0;border-bottom:1px solid var(--border)}
.cost-row.grand{border-top:2px solid var(--border2);margin-top:4px;padding-top:12px}
.cost-row.grand .cost-val{font-size:16px;font-weight:700;color:var(--accent)}
.cost-label{font-size:12.5px;color:var(--text2)}
.cost-val{font-family:var(--mono);font-size:12.5px;color:var(--text);text-align:right}
.report-section{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:20px 22px;margin-bottom:20px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
.cpt-banner{background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r2);padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:20px}
.cpt-val{font-size:30px;font-weight:800;color:var(--accent);font-family:var(--mono)}
.cpt-sub{font-size:11px;color:var(--text3);margin-top:3px;font-family:var(--mono)}
.pill-tabs{display:flex;gap:4px;background:var(--bg3);border-radius:var(--r2);padding:4px;margin-bottom:16px;width:fit-content}
.pill-tab{padding:6px 16px;border-radius:var(--r);font-size:12px;font-weight:500;cursor:pointer;border:none;background:none;color:var(--text3);font-family:var(--font);transition:all .12s}
.pill-tab.active{background:var(--bg2);color:var(--text);box-shadow:0 1px 4px rgba(0,0,0,.4)}
.lot-chip{display:inline-flex;align-items:center;gap:5px;background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r);padding:3px 10px;font-family:var(--mono);font-size:11px;color:var(--accent)}
.output-rows{display:flex;flex-direction:column;gap:8px}
.output-row{display:grid;grid-template-columns:1fr 1fr auto;gap:8px;align-items:end}
.hamburger-btn{display:none;background:none;border:1px solid var(--border2);border-radius:var(--r);color:var(--text2);font-size:15px;padding:5px 9px;cursor:pointer;line-height:1;align-items:center;justify-content:center}
.hamburger-btn:hover{color:var(--text);border-color:var(--border3)}
.sidebar-close{display:none;background:none;border:none;color:var(--text2);font-size:15px;cursor:pointer;padding:4px 8px;line-height:1}
.sidebar-close:hover{color:var(--text)}
.due-badge{font-size:10px;font-family:var(--mono);padding:1px 7px;border-radius:20px;background:#1a1500;color:var(--amber);border:1px solid #3a2e00}

/* ── Mobile & Tablet Responsiveness (Off-Canvas Sidebar Drawer) ── */
@media (max-width: 860px) {
  .sidebar {
    position: fixed;
    left: 0;
    top: 0;
    transform: translateX(-100%);
    width: 240px !important;
    box-shadow: var(--shadow2);
  }
  .sidebar.open {
    transform: translateX(0);
  }
  .sidebar-backdrop {
    position: fixed;
    inset: 0;
    background: rgba(0, 0, 0, 0.7);
    z-index: 250;
    backdrop-filter: blur(2px);
  }
  .main {
    width: 100vw;
  }
  .form-group input, .form-group select, .form-group textarea {
    font-size: 16px !important;
  }
  .btn {
    padding: 10px 14px;
    font-size: 13px;
    min-height: 38px;
  }
  .btn-sm {
    padding: 6px 10px;
    font-size: 11.5px;
    min-height: 30px;
  }
  .topbar {
    padding: 12px 18px;
  }
  .content {
    padding: 16px 18px;
  }
  .table-wrap {
    -webkit-overflow-scrolling: touch;
    overflow-x: auto;
  }
  .stats-grid, .stats-grid-5 {
    grid-template-columns: 1fr 1fr;
    gap: 8px;
  }
  .stat-card {
    padding: 12px 14px;
  }
  .stat-val {
    font-size: 19px;
  }
  .pill-tabs {
    width: 100%;
    overflow-x: auto;
    flex-wrap: nowrap;
    padding-bottom: 4px;
  }
  .pill-tab {
    flex-shrink: 0;
  }
  .two-col {
    grid-template-columns: 1fr;
  }
  .form-row.cols-2, .form-row.cols-3 {
    grid-template-columns: 1fr;
  }
  .hamburger-btn {
    display: inline-flex;
  }
  .sidebar-close {
    display: inline-flex;
  }
}
@media (max-width: 480px) {
  .stats-grid, .stats-grid-5 {
    grid-template-columns: 1fr;
  }
}
`;

// ── Helpers ───────────────────────────────────────────────────────────────────
const fmtMT = v => `${(+v||0).toFixed(3)} MT`;
const fmt = n => `₹${(+n||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const today = () => new Date().toISOString().split("T")[0];

const BANK_STOPWORDS = new Set(["UPI","NEFT","IMPS","RTGS","ACH","ACH DR","ACH CR","POS","ATM","DR","CR","WDL","TRF","TRANSFER","BIL","MOB","INB","ECOM","ECS","ONL","P2A","P2M","P2PM","TXN","REF","VPA","PAYMENT","AUTOPAY","MANDATE","NACH"]);
function normalizeDesc(desc) {
  const raw = (desc || "").toUpperCase();
  const tokens = raw.split(/[-\/|,]+/).map(t => t.trim()).filter(Boolean);
  const kept = tokens.filter(t => {
    if (BANK_STOPWORDS.has(t)) return false;
    if (/^\d+$/.test(t) && t.length >= 4) return false;
    if (t.length >= 10 && ((t.match(/\d/g) || []).length / t.length) >= 0.55) return false;
    if (/^[A-Z]{4}0[A-Z0-9]{6}$/.test(t)) return false;
    if (/^\d{1,2}[A-Z]{3}\d{2,4}$/.test(t)) return false;
    if (/^\d+@[A-Z0-9.]+$/.test(t)) return false;
    return t.length > 1;
  });
  return (kept.length ? kept : tokens).join(" ").replace(/\s+/g, " ").trim() || "UNKNOWN COUNTERPARTY";
}
const parseAmt = v => { if (!v) return 0; const n = parseFloat(String(v).replace(/[₹,\s]/g, "")); return isNaN(n) ? 0 : Math.abs(n); };
function parseBankDate(v) {
  if (!v) return "";
  if (typeof v === "number") { const ms = Date.UTC(1899, 11, 30) + v * 86400000; return new Date(ms).toISOString().slice(0,10); }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) { let [, d, mo, y] = m; if (y.length === 2) y = "20" + y; return `${y}-${mo.padStart(2,"0")}-${d.padStart(2,"0")}`; }
  return s.slice(0, 10);
}
function guessCol(headers, patterns) {
  const low = headers.map(h => String(h).toLowerCase());
  for (const p of patterns) { const i = low.findIndex(h => h.includes(p)); if (i > -1) return headers[i]; }
  return "";
}

const EXP_CATS = [
  { id:"fuel", label:"Fuel & Diesel" },
  { id:"salary", label:"Staff Salaries" },
  { id:"labour", label:"Contract / Casual Labour" },
  { id:"maint", label:"Maintenance" },
  { id:"royalty", label:"Royalty / Levy" },
  { id:"electric", label:"Electricity" },
  { id:"water", label:"Water" },
  { id:"transport", label:"Transport" },
  { id:"misc", label:"Miscellaneous" },
];
const OPS_CATS = ["electric","water","salary","labour","fuel","maint","transport","misc"];

const CAP_CATS = [
  { id: "land", label: "Land Acquisition & Development", defaultRate: 0 },
  { id: "civil", label: "Buildings & Civil Structures", defaultRate: 10 },
  { id: "machinery", label: "Plant & Machinery (Ball Mill, etc.)", defaultRate: 15 },
  { id: "electrical", label: "Electrical Infrastructure (Transformers/Panels)", defaultRate: 15 },
  { id: "utilities", label: "Utilities & Water Infrastructure", defaultRate: 15 },
  { id: "vehicle", label: "Vehicles & Material Handling", defaultRate: 30 },
  { id: "deposits", label: "Utility & Security Deposits", defaultRate: 0 },
  { id: "office", label: "Office & IT Equipment", defaultRate: 40 },
  { id: "other", label: "Other Capital Assets", defaultRate: 15 },
];

const CP_TYPES = [
  { id:"unlabeled", label:"Unlabeled" }, { id:"vendor", label:"Vendor / Supplier" },
  { id:"customer", label:"Customer / Buyer" }, { id:"staff", label:"Staff / Labour" },
  { id:"owner", label:"Owner / Capital" }, { id:"loan", label:"Loan / Bank" },
  { id:"utility", label:"Utility / Rent" }, { id:"tax", label:"Tax / Government" },
  { id:"capital_machinery", label:"Capital Asset: Machinery" },
  { id:"capital_land", label:"Capital Asset: Land & Civil" },
  { id:"capital_electrical", label:"Capital Asset: Electrical" },
  { id:"capital_vehicles", label:"Capital Asset: Vehicles" },
  { id:"other", label:"Other" },
];

// ── PDF.js Loader ─────────────────────────────────────────────────────────────
let _pdfjsPromise = null;
function loadPdfjs() {
  if (!_pdfjsPromise) {
    _pdfjsPromise = import("pdfjs-dist/legacy/build/pdf.mjs").then((lib) => {
      lib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${lib.version || "4.10.38"}/pdf.worker.min.mjs`;
      return lib;
    }).catch(() => {
      return import("pdfjs-dist").then((lib) => {
        lib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${lib.version || "4.10.38"}/pdf.worker.min.mjs`;
        return lib;
      });
    });
  }
  return _pdfjsPromise;
}

async function parsePdfStatement(file) {
  const pdfjsLib = await loadPdfjs();
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const allWords = []; let pageWidth = 612;
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    if (p === 1) pageWidth = viewport.width;
    const content = await page.getTextContent();
    content.items.forEach(it => {
      const t = (it.str || "").trim();
      if (!t) return;
      allWords.push({ page: p, x0: it.transform[4], top: viewport.height - it.transform[5], text: t });
    });
  }
  const sorted = [...allWords].sort((a,b) => a.page - b.page || a.top - b.top || a.x0 - b.x0);
  const lines = []; let cur = [];
  sorted.forEach(w => {
    if (cur.length && (w.page !== cur[0].page || Math.abs(w.top - cur[cur.length-1].top) > 4.5)) { lines.push(cur); cur = []; }
    cur.push(w);
  });
  if (cur.length) lines.push(cur);
  const lineText = line => line.map(w => w.text).join("").replace(/\s+/g, "");

  const REF = [[0,65],[65,278],[278,358],[358,404],[404,489],[489,563],[563,638]];
  const order = ["date","narr","ref","vdt","wd","dep","bal"];
  const k = pageWidth / 638;
  const colBounds = {};
  order.forEach((c, i) => { colBounds[c] = [REF[i][0] * k, i === order.length - 1 ? pageWidth + 50 : REF[i][1] * k]; });

  const wordsInCol = (line, col) => { const [lo, hi] = colBounds[col]; return line.filter(w => w.x0 >= lo && w.x0 < hi).sort((a,b) => a.x0 - b.x0); };

  const filtered = []; let inTable = false;
  for (const line of lines) {
    const txt = lineText(line);
    if (txt.includes("Statementofaccount")) { inTable = true; continue; }
    if (txt.includes("HDFCBANKLIMITED") || txt.toUpperCase().startsWith("STATEMENTSUMMARY")) { inTable = false; continue; }
    if (txt.startsWith("DateNarrationChq")) continue;
    if (inTable) filtered.push(line);
  }

  const parseAmtLocal = s => { const n = parseFloat(String(s).replace(/,/g, "")); return isNaN(n) ? null : n; };
  const txns = []; let t = null;
  for (const line of filtered) {
    const dwords = wordsInCol(line, "date");
    const isNew = dwords.length > 0 && /^\d{2}\/\d{2}\/\d{2}$/.test(dwords[0].text);
    if (isNew) {
      if (t) txns.push(t);
      t = { date: dwords[0].text, narr: [], ref: [], vdt: "", wd: null, dep: null };
      wordsInCol(line, "narr").forEach(w => t.narr.push(w.text));
      wordsInCol(line, "ref").forEach(w => t.ref.push(w.text));
      const vd = wordsInCol(line, "vdt"); if (vd.length) t.vdt = vd[0].text;
      const wd = wordsInCol(line, "wd"); if (wd.length) t.wd = parseAmtLocal(wd[0].text);
      const dep = wordsInCol(line, "dep"); if (dep.length) t.dep = parseAmtLocal(dep[0].text);
    } else if (t) {
      wordsInCol(line, "narr").forEach(w => t.narr.push(w.text));
      wordsInCol(line, "ref").forEach(w => t.ref.push(w.text));
    }
  }
  if (t) txns.push(t);
  return {
    headers: ["Date","Narration","Chq/Ref No","Value Dt","Withdrawal Amt","Deposit Amt"],
    rows: txns.map(tx => ({
      "Date": tx.date, "Narration": tx.narr.join(" "), "Chq/Ref No": tx.ref.join(""),
      "Value Dt": tx.vdt, "Withdrawal Amt": tx.wd != null ? tx.wd : "", "Deposit Amt": tx.dep != null ? tx.dep : "",
    })),
    fileName: file.name,
  };
}

// ── Navigation ────────────────────────────────────────────────────────────────
const NAV = [
  { id:"dashboard", label:"Dashboard", icon:"◈", group:null },
  { id:"vendorprojects", label:"Vendor Projects", icon:"🛠", group:"Commercial" },
  { id:"purchases", label:"Purchases", icon:"📥", group:"Commercial" },
  { id:"opscosts", label:"Monthly Overheads", icon:"🧾", group:"Commercial" },
  { id:"capital", label:"Capital & Infra", icon:"🏗", group:"Commercial" },
  { id:"bankstatement", label:"Bank Statement", icon:"🏦", group:"Banking" },
  { id:"partners", label:"Partners & Capital", icon:"🤝", group:"Banking" },
  { id:"loans", label:"Loans", icon:"💳", group:"Banking" },
  { id:"balancesheet", label:"Balance Sheet", icon:"🏛", group:"Reports" },
  { id:"suppliers", label:"Suppliers", icon:"⛏", group:"Masters" },
];
const TITLES = Object.fromEntries(NAV.map(n => [n.id, n.label]));
const GROUPS = [...new Set(NAV.filter(n => n.group).map(n => n.group))];

function BadgeComponent({ type, children }) { return <span className={`badge badge-${type||"muted"}`}>{children}</span>; }
function StatCard({ label, value, sub, color="blue" }) {
  return <div className={`stat-card ${color}`}><div className="stat-label">{label}</div><div className={`stat-val ${color}`}>{value}</div>{sub&&<div className="stat-sub">{sub}</div>}</div>;
}
function FG({ label, note, children, span }) {
  return <div className="form-group" style={span?{gridColumn:`span ${span}`}:{}}>{label&&<label>{label}</label>}{children}{note&&<div className="form-note">{note}</div>}</div>;
}
function Modal({ title, onClose, children, foot, size="" }) {
  return (
    <div className="modal-overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className={`modal ${size}`}>
        <div className="modal-head"><h3>{title}</h3><button className="btn btn-ghost btn-sm" onClick={onClose}>✕</button></div>
        <div className="modal-body">{children}</div>
        {foot&&<div className="modal-foot">{foot}</div>}
      </div>
    </div>
  );
}

function EmptyState({ icon, message, sub }) {
  return (
    <div className="empty">
      <div className="empty-icon">{icon || "📭"}</div>
      <p style={{ fontWeight: 600, marginBottom: 4 }}>{message}</p>
      {sub && <p style={{ fontSize: 11.5, color: "var(--text3)", marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

// ── Shared: search box used across the ledger/register views ──────────────────
function SearchBar({ value, onChange, placeholder, style }) {
  return (
    <input
      placeholder={placeholder}
      value={value}
      onChange={e => onChange(e.target.value)}
      style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: "var(--r)", padding: "7px 11px", color: "var(--text)", fontSize: 12.5, minWidth: 260, ...style }}
    />
  );
}

// ── Shared: Edit/Delete action pair for a transaction row ─────────────────────
function RowActions({ onEdit, onDelete }) {
  if (!onEdit && !onDelete) return null;
  return (
    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
      {onEdit && <button className="btn btn-ghost btn-sm" onClick={onEdit}>Edit</button>}
      {onDelete && <button className="btn btn-danger btn-sm" onClick={onDelete}>Delete</button>}
    </div>
  );
}

// ── Shared: export any row-set to an .xlsx file (uses the xlsx lib already in use for bank imports) ──
function exportRowsToExcel(fileName, rows) {
  if (!rows || !rows.length) { alert("Nothing to export yet."); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, fileName.endsWith(".xlsx") ? fileName : `${fileName}.xlsx`);
}

// ── Shared: bank transactions grouped + merged by counterparty label ──────────
// Centralizes the grouping logic that Expenses, Capital, Partners, and the Bank
// Statement view each used to re-derive independently. Returns one row per
// distinct counterparty label (or per raw narration key if never labeled), with
// its `type` (from CP_TYPES) so callers can filter to what they need:
//   - ExpensesView    -> exclude type === "owner" and type.startsWith("capital_")
//   - CapitalRegisterView -> only type.startsWith("capital_")
//   - PartnersCapitalDashboard -> only type === "owner"
function useCounterpartyGroups(bankTxns, bankLabels) {
  return useMemo(() => {
    const raw = {};
    (bankTxns || []).forEach(t => {
      if (!t) return;
      const k = t.key || t.group_key || (t.description ? normalizeDesc(t.description) : "UNKNOWN");
      if (!raw[k]) raw[k] = { key: k, txns: [], totalDebit: 0, totalCredit: 0 };
      raw[k].txns.push(t);
      raw[k].totalDebit += (+t.debit || 0);
      raw[k].totalCredit += (+t.credit || 0);
    });
    const merged = {};
    Object.values(raw).forEach(g => {
      const meta = (bankLabels || {})[g.key] || {};
      const type = meta.type || "unlabeled";
      const label = (meta.label && meta.label.trim()) ? meta.label.trim() : g.key;
      const mergeKey = label.toLowerCase();
      if (!merged[mergeKey]) {
        merged[mergeKey] = { mergeKey, label, type, rawKeys: [], txns: [], totalDebit: 0, totalCredit: 0 };
      }
      merged[mergeKey].rawKeys.push(g.key);
      merged[mergeKey].txns.push(...g.txns);
      merged[mergeKey].totalDebit += g.totalDebit;
      merged[mergeKey].totalCredit += g.totalCredit;
    });
    return Object.values(merged);
  }, [bankTxns, bankLabels]);
}

// ── Shared: straight-line net book value for a capital asset ──────────────────
function computeNetBookValue(category, amount, acquisitionDate) {
  const catConfig = CAP_CATS.find(x => x.id === category) || { defaultRate: 15 };
  const rate = catConfig.defaultRate / 100;
  const acqDate = new Date(acquisitionDate || today());
  const ageYears = Math.max(0, (new Date() - acqDate) / (1000 * 60 * 60 * 24 * 365.25));
  return rate === 0 ? amount : Math.max(0, amount * Math.pow(1 - rate, ageYears));
}

// ── APP MAIN ──────────────────────────────────────────────────────────────────
export default function App() {
  const [view, setView] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [loading, setLoading] = useState(true);

  const [suppliers, setSuppliers] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [capitalItems, setCapitalItems] = useState([]);
  const [capitalProjects, setCapitalProjects] = useState([]);
  const [bankBatches, setBankBatches] = useState([]);
  const [bankTxns, setBankTxns] = useState([]);
  const [bankLabels, setBankLabels] = useState({});
  const [partnerCashbook, setPartnerCashbook] = useState([]);
  const [vendorProjects, setVendorProjects] = useState([]);
  const [loanEntries, setLoanEntries] = useState([]);
  const [loanStatus, setLoanStatus] = useState({ ready: true, error: "" });
  const [projectPayments, setProjectPayments] = useState([]);
  const [vpStatus, setVpStatus] = useState({ ready: true, error: "" });

  useEffect(() => {
    async function init() {
      try {
        const d = await db.loadAll();
        setSuppliers(d.suppliers || []);
        setPurchases(d.purchases || []); setExpenses(d.expenses || []); setCapitalItems(d.capitalItems || []);
        setCapitalProjects(d.capitalProjects || []);
        setBankBatches(d.bankBatches || []); setBankTxns(d.bankTxns || []); setBankLabels(d.bankLabels || {});
        setPartnerCashbook(d.partnerCashbook || []);
      } catch (err) {
        console.error("Database sync failed:", err);
      }
      // Loaded on its own so the rest of the app still works if the
      // vendor-projects tables haven't been created yet.
      try {
        const v = await vpdb.loadVendorProjects();
        setVendorProjects(v.projects); setProjectPayments(v.payments);
        setVpStatus({ ready: v.ready, error: v.error });
      } catch (err) {
        setVpStatus({ ready: false, error: err.message || String(err) });
      }
      try {
        const l = await loansdb.loadLoans();
        setLoanEntries(l.entries); setLoanStatus({ ready: l.ready, error: l.error });
      } catch (err) {
        setLoanStatus({ ready: false, error: err.message || String(err) });
      } finally {
        setLoading(false);
      }
    }
    init();
  }, []);

  const goTo = id => { setView(id); setNavOpen(false); };

  if (loading) {
    return <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", background:"var(--bg)", color:"var(--accent)", fontFamily:"monospace" }}>Syncing Cloud Database...</div>;
  }

  return (
    <>
      <style>{FONTS}{CSS}</style>
      <div className="app">
        {navOpen && <div className="sidebar-backdrop" onClick={() => setNavOpen(false)} />}
        <aside className={`sidebar${navOpen ? " open" : ""}`}>
          <div className="sidebar-logo">
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
              <div><h1>Finance ERP</h1><p>Nova Geo Resources</p></div>
              <button className="sidebar-close" onClick={() => setNavOpen(false)}>✕</button>
            </div>
          </div>
          <nav className="nav">
            {NAV.filter(n => !n.group).map(n => (
              <button key={n.id} className={`nav-item${view === n.id ? " active" : ""}`} onClick={() => goTo(n.id)}>
                <span className="icon">{n.icon}</span>{n.label}
              </button>
            ))}
            {GROUPS.map(grp => (
              <div key={grp}>
                <div className="nav-label">{grp}</div>
                {NAV.filter(n => n.group === grp).map(n => (
                  <button key={n.id} className={`nav-item${view === n.id ? " active" : ""}`} onClick={() => goTo(n.id)}>
                    <span className="icon">{n.icon}</span>{n.label}
                  </button>
                ))}
              </div>
            ))}
          </nav>
          <div className="sidebar-foot">Real-Time Sync Active</div>
        </aside>

        <div className="main">
          <div className="topbar">
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <button className="hamburger-btn" onClick={() => setNavOpen(true)}>☰</button>
              <h2>{TITLES[view] || view}</h2>
            </div>
            <div className="topbar-right">
              <span className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{today()}</span>
              <BadgeComponent type="accent">Cloud Active</BadgeComponent>
            </div>
          </div>

          <div className="content">
            {view === "dashboard" && <DashboardView {...{ expenses, purchases, capitalItems, bankTxns, partnerCashbook, vendorProjects, projectPayments, goTo }} />}
            {view === "vendorprojects" && (
              <VendorProjectsView
                bankTxns={bankTxns} bankLabels={bankLabels} expenses={expenses} capitalItems={capitalItems} suppliers={suppliers}
                projects={vendorProjects} setProjects={setVendorProjects}
                projectPayments={projectPayments} setProjectPayments={setProjectPayments}
                partnerCashbook={partnerCashbook} setPartnerCashbook={setPartnerCashbook}
                setExpenses={setExpenses} setCapitalItems={setCapitalItems}
                status={vpStatus}
              />
            )}
            {view === "purchases" && <PurchasesView {...{ suppliers, purchases, setPurchases }} />}
            {view === "opscosts" && <MonthlyOpsCostsView {...{ expenses, setExpenses }} />}
            {view === "capital" && <CapitalRegisterView {...{ bankTxns, bankLabels, capitalItems, capitalProjects, setCapitalProjects }} />}
            {view === "bankstatement" && <BankStatementGroupingView {...{ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, capitalItems, setCapitalItems, setExpenses, setPartnerCashbook }} />}
            {view === "loans" && <LoansView {...{ bankTxns, bankLabels, loanEntries, setLoanEntries, loanStatus, vendorProjects, setProjectPayments }} />}
            {view === "partners" && (
              <PartnersCapitalDashboard
                bankTxns={bankTxns}
                bankLabels={bankLabels}
                partnerCashbook={partnerCashbook}
                setPartnerCashbook={setPartnerCashbook}
                setExpenses={setExpenses}
                setCapitalItems={setCapitalItems}
              />
            )}
            {view === "balancesheet" && <BalanceSheetView {...{ capitalItems, bankTxns, bankLabels, partnerCashbook, purchases, expenses }} />}
            {view === "suppliers" && <MasterTableView title="Suppliers" noun="Supplier" items={suppliers} setItems={setSuppliers} saveFn={db.saveSupplier} delFn={db.deleteSupplier} fields={[{ key:"name", label:"Supplier Name", required:true },{ key:"contact", label:"Contact" }]} cols={[{ label:"Name", key:"name" },{ label:"Contact", key:"contact" }]} />}
          </div>
        </div>
      </div>
    </>
  );
}

// ── COMMERCIAL & REPORTING VIEWS ──────────────────────────────────────────────

function PurchasesView({ suppliers, purchases, setPurchases }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [expandedSupplier, setExpandedSupplier] = useState(null);
  const [search, setSearch] = useState("");

  const blank = { date: today(), supplierId: "", description: "", quantityMT: "", ratePerMT: "", gstPct: "5", invoiceNo: "", remarks: "" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  function openAdd() { setEditingId(null); setF(blank); setOpen(true); }
  function openEdit(p) {
    setEditingId(p.id);
    setF({
      date: p.date || today(), supplierId: p.supplierId ? String(p.supplierId) : "", description: p.description || "",
      quantityMT: String(p.quantityMT ?? ""), ratePerMT: String(p.ratePerMT ?? ""), gstPct: String(p.gstPct ?? "5"),
      invoiceNo: p.invoiceNo || "", remarks: p.remarks || "",
    });
    setOpen(true);
  }

  async function save() {
    if (!f.supplierId || !f.quantityMT || !f.ratePerMT) return alert("Supplier, Quantity, and Rate required.");
    try {
      const payload = editingId ? { ...f, id: editingId } : f;
      const row = await db.savePurchase(payload);
      if (editingId) setPurchases(ps => ps.map(p => p.id === editingId ? row : p));
      else setPurchases(ps => [row, ...ps]);
      setOpen(false); setF(blank); setEditingId(null);
    } catch (err) { alert(err.message); }
  }

  async function remove(id) {
    if (!confirm("Delete this purchase order? This cannot be undone.")) return;
    if (typeof db.deletePurchase !== "function") return alert("db.deletePurchase isn't implemented yet — add it to db.js to enable deleting POs.");
    try {
      await db.deletePurchase(id);
      setPurchases(ps => ps.filter(p => p.id !== id));
    } catch (err) { alert(err.message); }
  }

  const supplierGroups = useMemo(() => {
    const map = {};
    (purchases || []).forEach(p => {
      if (!p) return;
      const sName = p.supplierName || "Unknown Supplier";
      const key = sName.toLowerCase();
      if (!map[key]) {
        map[key] = { key, supplierName: sName, totalQtyMT: 0, totalSpend: 0, pos: [] };
      }
      map[key].totalQtyMT += (+p.quantityMT || 0);
      map[key].totalSpend += (+p.totalAmount || 0);
      map[key].pos.push(p);
    });
    return Object.values(map).sort((a, b) => b.totalSpend - a.totalSpend);
  }, [purchases]);

  const filteredGroups = supplierGroups.filter(g => {
    if (!search) return true;
    return g.supplierName.toLowerCase().includes(search.toLowerCase()) || g.pos.some(p => (p.invoiceNo || "").toLowerCase().includes(search.toLowerCase()));
  });

  const totalSpendAll = (purchases || []).reduce((s, p) => s + (+p.totalAmount || 0), 0);
  const totalQtyAll = (purchases || []).reduce((s, p) => s + (+p.quantityMT || 0), 0);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h3 style={{ fontSize: 15, fontWeight: 700, marginBottom: 4 }}>Raw Material Purchases & Supplier Ledger</h3>
          <p style={{ fontSize: 12, color: "var(--text3)" }}>Track raw boulder deliveries, supplier rates, and purchase order invoices.</p>
        </div>
        <button className="btn btn-primary" onClick={openAdd}>+ New RM Purchase Order</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Total Spend" value={fmt(totalSpendAll)} color="amber" sub="Cumulative PO value" />
        <StatCard label="Total Boulder Procured" value={fmtMT(totalQtyAll)} color="teal" sub="Raw material tonnage" />
        <StatCard label="Active Suppliers" value={supplierGroups.length} color="blue" sub="Supplier accounts" />
        <StatCard label="Total POs Logged" value={(purchases || []).length} color="purple" sub="Invoice entries" />
      </div>

      <div className="filter-bar">
        <SearchBar value={search} onChange={setSearch} placeholder="Search supplier or invoice #…" />
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("purchases", (purchases || []).map(p => ({ PO: p.poNo, Date: p.date, Supplier: p.supplierName, Invoice: p.invoiceNo || "", QtyMT: p.quantityMT, RatePerMT: p.ratePerMT, GST: p.gstPct, TotalAmount: p.totalAmount })))}>⬇ Export Excel</button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filteredGroups.length === 0 ? (
          <div className="config-card" style={{ textAlign: "center", padding: 32 }}>
            <EmptyState icon="📥" message="No purchases recorded yet" sub="Click above to log raw material purchase orders." />
          </div>
        ) : (
          filteredGroups.map(g => (
            <div key={g.key} className="config-card" style={{ padding: 14 }}>
              <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", justifyContent: "space-between" }}>
                <div>
                  <div style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>Supplier Account</div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: "var(--text)", marginTop: 2 }}>{g.supplierName}</div>
                </div>
                <div style={{ display: "flex", gap: 18, fontFamily: "var(--mono)", fontSize: 12.5 }}>
                  <div><div style={{ color: "var(--text3)", fontSize: 10 }}>TOTAL SPEND</div><div style={{ color: "var(--amber)", fontWeight: 700 }}>{fmt(g.totalSpend)}</div></div>
                  <div><div style={{ color: "var(--text3)", fontSize: 10 }}>QTY PROCURED</div><div style={{ color: "var(--teal)", fontWeight: 700 }}>{fmtMT(g.totalQtyMT)}</div></div>
                  <div><div style={{ color: "var(--text3)", fontSize: 10 }}>ORDERS</div><div style={{ fontWeight: 700 }}>{g.pos.length}</div></div>
                </div>
              </div>

              <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)" }}>
                  Avg Rate: ₹{g.totalQtyMT > 0 ? (g.totalSpend / g.totalQtyMT).toFixed(0) : 0} / MT
                </span>
                <button className="btn btn-ghost btn-sm" onClick={() => setExpandedSupplier(x => x === g.key ? null : g.key)}>
                  {expandedSupplier === g.key ? "Hide" : "View"} purchase orders ({g.pos.length})
                </button>
              </div>

              {expandedSupplier === g.key && (
                <div className="table-wrap" style={{ marginTop: 10 }}>
                  <table>
                    <thead>
                      <tr>
                        <th>PO #</th>
                        <th>Date</th>
                        <th>Invoice #</th>
                        <th className="r">Qty (MT)</th>
                        <th className="r">Rate / MT</th>
                        <th className="r">Total Amount</th>
                        <th className="r">Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {g.pos.sort((a, b) => (b.date || "").localeCompare(a.date || "")).map(p => (
                        <tr key={p.id}>
                          <td className="mono" style={{ color: "var(--accent)" }}>{p.poNo}</td>
                          <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{p.date}</td>
                          <td className="mono">{p.invoiceNo || "—"}</td>
                          <td className="r mono">{fmtMT(p.quantityMT)}</td>
                          <td className="r mono">₹{p.ratePerMT}</td>
                          <td className="r mono" style={{ fontWeight: 700, color: "var(--amber)" }}>{fmt(p.totalAmount)}</td>
                          <td className="r"><RowActions onEdit={() => openEdit(p)} onDelete={() => remove(p.id)} /></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          ))
        )}
      </div>

      {open && (
        <Modal title={editingId ? "Edit Raw Material Purchase" : "Record Raw Material Purchase"} onClose={() => { setOpen(false); setEditingId(null); }} foot={<><button className="btn btn-ghost" onClick={() => { setOpen(false); setEditingId(null); }}>Cancel</button><button className="btn btn-primary" onClick={save}>{editingId ? "Save Changes" : "Save Purchase"}</button></>}>
          <div className="form-row cols-2">
            <FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG>
            <FG label="Supplier *">
              <select value={f.supplierId} onChange={e => set("supplierId")(e.target.value)}>
                <option value="">Select supplier...</option>
                {suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-3">
            <FG label="Quantity (MT) *"><input type="number" step="0.001" value={f.quantityMT} onChange={e => set("quantityMT")(e.target.value)} /></FG>
            <FG label="Rate / MT (₹) *"><input type="number" value={f.ratePerMT} onChange={e => set("ratePerMT")(e.target.value)} /></FG>
            <FG label="GST %">
              <select value={f.gstPct} onChange={e => set("gstPct")(e.target.value)}>
                {[0, 5, 12, 18].map(r => <option key={r} value={r}>{r}%</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row">
            <FG label="Invoice / Challan #"><input value={f.invoiceNo} onChange={e => set("invoiceNo")(e.target.value)} /></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

function MonthlyOpsCostsView({ expenses, setExpenses }) {
  const [ym, setYm] = useState(today().slice(0, 7));
  const [vals, setVals] = useState({});
  useEffect(() => {
    const next = {};
    OPS_CATS.forEach(c => {
      const ex = expenses.filter(e => e.opsAuto && e.category === c && e.date?.slice(0, 7) === ym);
      next[c] = ex.length ? String(ex.reduce((s, e) => s + e.amount, 0)) : "";
    });
    setVals(next);
  }, [ym, expenses]);

  async function save() {
    try {
      const updated = await db.saveMonthlyOpsCosts(ym, vals);
      setExpenses(es => [...es.filter(e => !(e.opsAuto && e.date?.slice(0,7) === ym)), ...updated]);
      alert("Monthly overheads synchronized!");
    } catch (err) { alert(err.message); }
  }

  return (
    <div className="config-card">
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
        <h3>Monthly Overheads Quick Entry</h3>
        <FG label="Month Filter"><input type="month" value={ym} onChange={e => setYm(e.target.value)} style={{ padding:4 }} /></FG>
      </div>
      <div className="config-row">
        {OPS_CATS.map(c => (
          <FG key={c} label={`${c.toUpperCase()} (₹)`}>
            <input type="number" value={vals[c] || ""} onChange={e => setVals({ ...vals, [c]: e.target.value })} />
          </FG>
        ))}
      </div>
      <button className="btn btn-primary" style={{ marginTop: 14 }} onClick={save}>Sync Overheads to Ledger</button>
    </div>
  );
}

function CapitalRegisterView({ bankTxns, bankLabels, capitalItems, capitalProjects, setCapitalProjects }) {
  const [activeTab, setActiveTab] = useState("ledger");
  const [expandedVendor, setExpandedVendor] = useState(null);
  const [search, setSearch] = useState("");
  const [selectedVendorForStmt, setSelectedVendorForStmt] = useState(null);
  const [includeNBV, setIncludeNBV] = useState(true);

  const bankGroups = useCounterpartyGroups(bankTxns, bankLabels);

  // Capital & Infra is bank-statement-only: every row here is an actual bank
  // debit whose counterparty was tagged as a capital_* type in Bank Statement.
  // There's no manual "add a fixed asset" path — that used to double-count
  // against the same bank debit, so it's gone rather than fixed a third time.
  const allCapitalRows = useMemo(() => {
    const list = [];
    const bankCategoryByVendor = {}; // vendor name (lowercase) -> category, from bank-tagged rows

    bankGroups.forEach(g => {
      if (!g.type.startsWith("capital_")) return;
      let cat = "machinery";
      if (g.type === "capital_land") cat = "land";
      else if (g.type === "capital_electrical") cat = "electrical";
      else if (g.type === "capital_vehicles") cat = "vehicle";

      bankCategoryByVendor[g.label.trim().toLowerCase()] = cat;

      g.txns.filter(t => t.debit > 0).forEach(t => {
        const amount = +t.debit || 0;
        list.push({
          id: `bank-${t.id}`,
          category: cat,
          vendorName: g.label.toUpperCase(),
          date: t.date || t.txn_date,
          description: t.description,
          amount,
          netBookValue: computeNetBookValue(cat, amount, t.date || t.txn_date),
          source: "Bank Statement",
        });
      });
    });

    // Only cash capex payments an owner made directly (Partners tab ->
    // "Record Cash Spent by Partner" -> "Fixed Asset / Land"). These have no
    // bankTxnId and are tagged with paidByPartner — that combination is what
    // marks them as an owner cash spend rather than any other kind of entry.
    // Manual/plain capitalItems entries (no paidByPartner) are intentionally
    // excluded — this tab never shows a bare manual entry, only bank-tagged
    // debits and owner cash spends.
    (capitalItems || []).forEach(c => {
      if (!c || c.bankTxnId || !c.paidByPartner) return;
      const vendorName = (c.paidTo || c.description || "General Capital Expense").trim();
      const vendorKey = vendorName.toLowerCase();
      // If this vendor already has a block from a tagged bank counterparty,
      // join that same block instead of opening a new one under whatever
      // category was picked in the cash-spend modal.
      const cat = bankCategoryByVendor[vendorKey] || c.category || "other";

      list.push({
        id: `cap-${c.id}`,
        category: cat,
        vendorName: vendorName.toUpperCase(),
        date: c.date,
        description: c.description,
        amount: c.amount,
        netBookValue: computeNetBookValue(cat, c.amount, c.date),
        source: `Cash — Paid by ${c.paidByPartner}`,
      });
    });

    return list.sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  }, [bankGroups, capitalItems]);

  // Group vendors directly by category -> Vendor Label name
  const categoryGroups = useMemo(() => {
    const map = {};
    CAP_CATS.forEach(c => {
      map[c.id] = { categoryId: c.id, label: c.label, totalCost: 0, totalNBV: 0, vendorsMap: {} };
    });

    allCapitalRows.forEach(item => {
      const cat = item.category || "other";
      if (!map[cat]) {
        map[cat] = { categoryId: cat, label: cat.toUpperCase(), totalCost: 0, totalNBV: 0, vendorsMap: {} };
      }
      map[cat].totalCost += item.amount;
      map[cat].totalNBV += item.netBookValue;

      const vName = (item.vendorName || "General").trim();
      const vKey = vName.toLowerCase();
      if (!map[cat].vendorsMap[vKey]) {
        map[cat].vendorsMap[vKey] = { vendorKey: `${cat}-${vKey}`, vendorName: vName, category: cat, totalCost: 0, totalNBV: 0, txns: [] };
      }
      map[cat].vendorsMap[vKey].totalCost += item.amount;
      map[cat].vendorsMap[vKey].totalNBV += item.netBookValue;
      map[cat].vendorsMap[vKey].txns.push(item);
    });

    return Object.values(map)
      .filter(g => g.totalCost > 0)
      .map(g => ({
        ...g,
        vendors: Object.values(g.vendorsMap).sort((a, b) => b.totalCost - a.totalCost)
      }))
      .sort((a, b) => b.totalCost - a.totalCost);
  }, [allCapitalRows]);

  const filteredGroups = categoryGroups.filter(g => {
    if (!search) return true;
    return g.label.toLowerCase().includes(search.toLowerCase()) || g.vendors.some(v => v.vendorName.toLowerCase().includes(search.toLowerCase()) || v.txns.some(t => t.description.toLowerCase().includes(search.toLowerCase())));
  });

  const totalCapexAll = allCapitalRows.reduce((s, c) => s + c.amount, 0);
  const totalBookValAll = allCapitalRows.reduce((s, c) => s + c.netBookValue, 0);

  return (
    <div>
      <div className="pill-tabs">
        <button className={`pill-tab ${activeTab === "ledger" ? "active" : ""}`} onClick={() => setActiveTab("ledger")}>Asset Ledger</button>
        <button className={`pill-tab ${activeTab === "projects" ? "active" : ""}`} onClick={() => setActiveTab("projects")}>
          Projects & Dues {(capitalProjects || []).filter(p => p.due > 0).length > 0 && `(${(capitalProjects || []).filter(p => p.due > 0).length} open)`}
        </button>
      </div>

      {activeTab === "ledger" ? (
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
            <div>
              <h3 style={{ fontSize: 15, fontWeight: 700, marginBottom: 4 }}>Capital & Infrastructure Assets (Capex)</h3>
              <p style={{ fontSize: 12, color: "var(--text3)" }}>Gross Invested: <strong style={{ color: "var(--accent)" }}>{fmt(totalCapexAll)}</strong> | Net Book Value: <strong style={{ color: "var(--teal)" }}>{fmt(totalBookValAll)}</strong></p>
              <p style={{ fontSize: 11, color: "var(--text3)", marginTop: 2 }}>Sourced from Bank Statement (tag a counterparty as a Capital Asset type there) and from cash payments made directly by an owner/partner (Partners & Capital tab).</p>
            </div>
          </div>

          <div className="stats-grid">
            <StatCard label="Gross Capital Invested" value={fmt(totalCapexAll)} color="accent" sub="Historical cost" />
            <StatCard label="Net Book Value" value={fmt(totalBookValAll)} color="teal" sub="Post-depreciation value" />
            <StatCard label="Active Asset Blocks" value={categoryGroups.length} color="blue" sub="Categories utilized" />
            <StatCard label="Registered Items" value={allCapitalRows.length} color="purple" sub="Total asset records" />
          </div>

          <div className="filter-bar">
            <SearchBar value={search} onChange={setSearch} placeholder="Search asset block, vendor, or description…" style={{ minWidth: 280 }} />
            <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("capital_assets", allCapitalRows.map(r => ({ Category: r.category, Vendor: r.vendorName, Date: r.date, Description: r.description, Amount: r.amount, NetBookValue: Math.round(r.netBookValue), Source: r.source })))}>⬇ Export Excel</button>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
            {filteredGroups.length === 0 ? (
              <div className="config-card" style={{ textAlign: "center", padding: 32 }}>
                <EmptyState icon="🏗" message="No capital assets found" sub="Tag a counterparty as a Capital Asset type in Bank Statement to see it here." />
              </div>
            ) : (
              filteredGroups.map(g => (
                <div key={g.categoryId} className="config-card" style={{ padding: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid var(--border)", paddingBottom: 10, marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>Asset Block</div>
                      <div style={{ fontSize: 16, fontWeight: 700, color: "var(--text)", marginTop: 2 }}>{g.label}</div>
                    </div>
                    <div style={{ display: "flex", gap: 16, fontFamily: "var(--mono)", fontSize: 12 }}>
                      <div>Gross Cost: <span style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(g.totalCost)}</span></div>
                      <div>Net Book Value: <span style={{ color: "var(--teal)", fontWeight: 700 }}>{fmt(g.totalNBV)}</span></div>
                    </div>
                  </div>

                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {g.vendors.map(v => (
                      <div key={v.vendorKey} style={{ background: "var(--bg3)", border: "1px solid var(--border)", borderRadius: "var(--r)", padding: 14 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 10 }}>
                          <div>
                            <div style={{ fontSize: 10, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>Vendor / Counterparty Label</div>
                            <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)", marginTop: 2 }}>{v.vendorName}</div>
                          </div>
                          <div style={{ display: "flex", gap: 8, fontFamily: "var(--mono)", fontSize: 12, alignItems: "center" }}>
                            <div>PAID: <span style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(v.totalCost)}</span></div>
                            <div style={{ marginLeft: 8 }}>NET VALUE: <span style={{ color: "var(--teal)", fontWeight: 700 }}>{fmt(v.totalNBV)}</span></div>
                            <button className="btn btn-info btn-sm" onClick={() => setSelectedVendorForStmt(v)}>📄 Statement</button>
                            <button className="btn btn-ghost btn-sm" onClick={() => setExpandedVendor(x => x === v.vendorKey ? null : v.vendorKey)}>
                              {expandedVendor === v.vendorKey ? `Hide transactions (${v.txns.length})` : `View transactions (${v.txns.length})`}
                            </button>
                          </div>
                        </div>

                        {expandedVendor === v.vendorKey && (
                          <div className="table-wrap" style={{ marginTop: 12, marginBottom: 0 }}>
                            <table>
                              <thead>
                                <tr>
                                  <th>Date</th>
                                  <th>Narration / Description</th>
                                  <th>Source</th>
                                  <th className="r">Amount</th>
                                  <th className="r">Net Book Value</th>
                                </tr>
                              </thead>
                              <tbody>
                                {v.txns.sort((a, b) => (b.date || "").localeCompare(a.date || "")).map(t => (
                                  <tr key={t.id}>
                                    <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{t.date}</td>
                                    <td style={{ fontSize: 12 }}>{t.description}</td>
                                    <td><BadgeComponent type={t.source === "Bank Statement" ? "blue" : "accent"}>{t.source}</BadgeComponent></td>
                                    <td className="r mono" style={{ fontWeight: 700 }}>{fmt(t.amount)}</td>
                                    <td className="r mono" style={{ color: "var(--teal)", fontWeight: 700 }}>{fmt(t.netBookValue)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              ))
            )}
          </div>

          {selectedVendorForStmt && (
            <Modal
              title={`Capital Asset Statement: ${selectedVendorForStmt.vendorName}`}
              size="modal-lg"
              onClose={() => setSelectedVendorForStmt(null)}
              foot={
                <button className="btn btn-ghost" onClick={() => exportStatementPDF({
                  title: `Capital Asset Statement: ${selectedVendorForStmt.vendorName}`,
                  subtitle: `Category: ${CAP_CATS.find(c => c.id === selectedVendorForStmt.category)?.label || selectedVendorForStmt.category}`,
                  summary: [
                    { label: "Total Paid", value: fmt(selectedVendorForStmt.totalCost), color: [200, 160, 40] },
                    ...(includeNBV ? [{ label: "Net Book Value", value: fmt(selectedVendorForStmt.totalNBV), color: [40, 180, 160] }] : []),
                  ],
                  columns: [
                    { key: "date", label: "Date" },
                    { key: "description", label: "Description" },
                    { key: "source", label: "Source" },
                    { key: "amount", label: "Amount", align: "right" },
                    ...(includeNBV ? [{ key: "nbv", label: "Net Book Value", align: "right" }] : []),
                  ],
                  rows: selectedVendorForStmt.txns
                    .slice().sort((a, b) => (b.date || "").localeCompare(a.date || ""))
                    .map(t => ({
                      date: t.date, description: t.description, source: t.source, amount: fmt(t.amount),
                      ...(includeNBV ? { nbv: fmt(t.netBookValue) } : {}),
                    })),
                  fileName: `${selectedVendorForStmt.vendorName}_capital_statement.pdf`,
                })}>📄 Export PDF</button>
              }
            >
              <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, fontSize: 12.5, color: "var(--text2)" }}>
                <input type="checkbox" checked={includeNBV} onChange={e => setIncludeNBV(e.target.checked)} />
                Include Net Book Value in exported PDF
              </label>
              <div style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", gap: 20, fontFamily: "var(--mono)", fontSize: 13, background: "var(--bg3)", padding: 12, borderRadius: "var(--r)" }}>
                  <div>Total Paid: <span style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(selectedVendorForStmt.totalCost)}</span></div>
                  <div>Net Book Value: <span style={{ color: "var(--teal)", fontWeight: 700 }}>{fmt(selectedVendorForStmt.totalNBV)}</span></div>
                </div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr><th>Date</th><th>Description</th><th>Source</th><th className="r">Amount</th><th className="r">Net Book Value</th></tr>
                  </thead>
                  <tbody>
                    {selectedVendorForStmt.txns.map(t => (
                      <tr key={t.id}>
                        <td className="mono">{t.date}</td>
                        <td>{t.description}</td>
                        <td><BadgeComponent type="muted">{t.source}</BadgeComponent></td>
                        <td className="r mono" style={{ fontWeight: 700 }}>{fmt(t.amount)}</td>
                        <td className="r mono" style={{ color: "var(--teal)" }}>{fmt(t.netBookValue)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Modal>
          )}
        </div>
      ) : (
        <CapitalProjectsView bankTxns={bankTxns} capitalProjects={capitalProjects} setCapitalProjects={setCapitalProjects} />
      )}
    </div>
  );
}

function CapitalProjectsView({ bankTxns, capitalProjects, setCapitalProjects }) {
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [search, setSearch] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [selectedForStmt, setSelectedForStmt] = useState(null);
  const [paymentModalFor, setPaymentModalFor] = useState(null);
  const [linkTxnModalFor, setLinkTxnModalFor] = useState(null);
  const [linkTxnSearch, setLinkTxnSearch] = useState("");

  const blank = { date: today(), category: "machinery", vendorName: "", description: "", totalCost: "", reference: "" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  function openAdd() { setEditingId(null); setF(blank); setOpen(true); }
  function openEdit(p) {
    setEditingId(p.id);
    setF({ date: p.date, category: p.category, vendorName: p.vendorName, description: p.description, totalCost: String(p.totalCost), reference: p.reference });
    setOpen(true);
  }

  async function save() {
    if (!f.vendorName || !f.totalCost) return alert("Vendor and total cost required.");
    try {
      const payload = editingId ? { ...f, id: editingId } : f;
      const row = await db.saveCapitalProject(payload);
      if (editingId) {
        setCapitalProjects(ps => ps.map(p => p.id === editingId ? { ...p, ...row, payments: p.payments, paid: p.paid, due: Math.max(0, row.totalCost - p.paid) } : p));
      } else {
        setCapitalProjects(ps => [{ ...row, payments: [], paid: 0, due: row.totalCost }, ...ps]);
      }
      setOpen(false); setF(blank); setEditingId(null);
    } catch (err) { alert(err.message); }
  }

  async function remove(id) {
    if (!confirm("Delete this capital project and all its linked payments? This cannot be undone.")) return;
    try { await db.deleteCapitalProject(id); setCapitalProjects(ps => ps.filter(p => p.id !== id)); }
    catch (err) { alert(err.message); }
  }

  function recompute(p) {
    const paid = (p.payments || []).reduce((s, x) => s + x.amount, 0);
    return { ...p, paid, due: Math.max(0, p.totalCost - paid) };
  }

  async function addPayment(project, data) {
    if (!data.amount) return alert("Amount is required.");
    try {
      const p = await db.addCapitalProjectPayment(project.id, data);
      setCapitalProjects(ps => ps.map(x => x.id === project.id ? recompute({ ...x, payments: [p, ...(x.payments || [])] }) : x));
      setPaymentModalFor(null);
    } catch (err) { alert(err.message); }
  }

  async function removePayment(project, paymentId) {
    if (!confirm("Remove this payment?")) return;
    try {
      await db.deleteCapitalProjectPayment(paymentId);
      setCapitalProjects(ps => ps.map(x => x.id === project.id ? recompute({ ...x, payments: (x.payments || []).filter(p => p.id !== paymentId) }) : x));
    } catch (err) { alert(err.message); }
  }

  async function linkBankTxn(project, txn) {
    await addPayment(project, { date: txn.date, amount: txn.debit, paymentMode: "bank", bankTxnId: txn.id, remarks: txn.description });
    setLinkTxnModalFor(null); setLinkTxnSearch("");
  }

  const linkedBankTxnIds = useMemo(
    () => new Set((capitalProjects || []).flatMap(p => (p.payments || []).map(x => x.bankTxnId).filter(Boolean))),
    [capitalProjects]
  );
  const unlinkedBankDebits = useMemo(() => (bankTxns || []).filter(t => t.debit > 0 && !linkedBankTxnIds.has(t.id)), [bankTxns, linkedBankTxnIds]);
  const filteredUnlinkedTxns = unlinkedBankDebits.filter(t => !linkTxnSearch || (t.description || "").toLowerCase().includes(linkTxnSearch.toLowerCase()));

  const filtered = (capitalProjects || []).filter(p => !search || p.vendorName.toLowerCase().includes(search.toLowerCase()) || (p.description || "").toLowerCase().includes(search.toLowerCase())).sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  const totalCommitted = (capitalProjects || []).reduce((s, p) => s + p.totalCost, 0);
  const totalPaid = (capitalProjects || []).reduce((s, p) => s + p.paid, 0);
  const totalDue = (capitalProjects || []).reduce((s, p) => s + p.due, 0);
  const openCount = (capitalProjects || []).filter(p => p.due > 0).length;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div>
          <h3 style={{ fontSize: 15, fontWeight: 700, marginBottom: 4 }}>Capital Projects & Dues</h3>
          <p style={{ fontSize: 12, color: "var(--text3)" }}>Quote or budget a capex job, then link payments to it to track what's still owed.</p>
        </div>
        <button className="btn btn-primary" onClick={openAdd}>+ New Capital Project</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Total Committed" value={fmt(totalCommitted)} color="blue" sub="Quoted/budgeted" />
        <StatCard label="Total Paid" value={fmt(totalPaid)} color="green" sub="Linked payments" />
        <StatCard label="Total Due" value={fmt(totalDue)} color="amber" sub="Outstanding" />
        <StatCard label="Open Projects" value={openCount} color="purple" sub={`of ${(capitalProjects || []).length} total`} />
      </div>

      <div className="filter-bar">
        <SearchBar value={search} onChange={setSearch} placeholder="Search vendor or description…" />
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("capital_projects", filtered.map(p => ({ Project: p.projectNo, Date: p.date, Vendor: p.vendorName, Category: p.category, Description: p.description, TotalCost: p.totalCost, Paid: p.paid, Due: p.due })))}>⬇ Export Excel</button>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtered.length === 0 ? (
          <div className="config-card" style={{ textAlign: "center", padding: 32 }}>
            <EmptyState icon="🏗" message="No capital projects yet" sub="Create one to start tracking a job that'll be paid off over time." />
          </div>
        ) : filtered.map(p => (
          <div key={p.id} className="config-card" style={{ padding: 14 }}>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", justifyContent: "space-between" }}>
              <div>
                <div style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>{p.projectNo} · {p.date}</div>
                <div style={{ fontSize: 15, fontWeight: 700, marginTop: 2 }}>{p.vendorName}</div>
                <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{p.description}</div>
              </div>
              <div style={{ display: "flex", gap: 18, fontFamily: "var(--mono)", fontSize: 12.5, alignItems: "center" }}>
                <div><div style={{ color: "var(--text3)", fontSize: 10 }}>TOTAL</div><div style={{ fontWeight: 700 }}>{fmt(p.totalCost)}</div></div>
                <div><div style={{ color: "var(--text3)", fontSize: 10 }}>PAID</div><div style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(p.paid)}</div></div>
                {p.due > 0 ? <span className="due-badge">DUE {fmt(p.due)}</span> : <BadgeComponent type="green">Fully Paid</BadgeComponent>}
              </div>
            </div>
            <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
              <span style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)" }}>Category: {p.category}</span>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {p.due > 0 && <button className="btn btn-primary btn-sm" onClick={() => setPaymentModalFor(p)}>+ Add Payment</button>}
                {p.due > 0 && <button className="btn btn-ghost btn-sm" onClick={() => setLinkTxnModalFor(p)}>🔗 Link Bank Txn</button>}
                <button className="btn btn-info btn-sm" onClick={() => setSelectedForStmt(p)}>📄 Statement</button>
                <button className="btn btn-ghost btn-sm" onClick={() => openEdit(p)}>Edit</button>
                <button className="btn btn-danger btn-sm" onClick={() => remove(p.id)}>Delete</button>
                <button className="btn btn-ghost btn-sm" onClick={() => setExpanded(x => x === p.id ? null : p.id)}>{expanded === p.id ? "Hide" : "Show"} payments ({(p.payments || []).length})</button>
              </div>
            </div>
            {expanded === p.id && (
              <div className="table-wrap" style={{ marginTop: 10 }}>
                <table>
                  <thead><tr><th>Date</th><th>Mode</th><th>Reference / Remarks</th><th className="r">Amount</th><th className="r">Actions</th></tr></thead>
                  <tbody>
                    {(p.payments || []).length === 0 ? (
                      <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--text3)", padding: 16 }}>No payments linked yet</td></tr>
                    ) : p.payments.map(x => (
                      <tr key={x.id}>
                        <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{x.date}</td>
                        <td><BadgeComponent type={x.bankTxnId ? "blue" : "accent"}>{x.bankTxnId ? "Bank Txn" : x.paymentMode}</BadgeComponent></td>
                        <td style={{ fontSize: 12 }}>{x.reference || x.remarks || "—"}</td>
                        <td className="r mono" style={{ fontWeight: 700, color: "var(--green)" }}>{fmt(x.amount)}</td>
                        <td className="r"><RowActions onDelete={() => removePayment(p, x.id)} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        ))}
      </div>

      {open && (
        <Modal title={editingId ? "Edit Capital Project" : "New Capital Project"} onClose={() => { setOpen(false); setEditingId(null); }} foot={<><button className="btn btn-ghost" onClick={() => { setOpen(false); setEditingId(null); }}>Cancel</button><button className="btn btn-primary" onClick={save}>{editingId ? "Save Changes" : "Create Project"}</button></>}>
          <div className="form-row cols-2">
            <FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG>
            <FG label="Category"><select value={f.category} onChange={e => set("category")(e.target.value)}>{CAP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></FG>
          </div>
          <div className="form-row"><FG label="Vendor Name *" note="Match your vendor's exact name to keep statements consistent"><input value={f.vendorName} onChange={e => set("vendorName")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Description"><input value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
          <div className="form-row cols-2">
            <FG label="Total Quoted Cost (₹) *"><input type="number" value={f.totalCost} onChange={e => set("totalCost")(e.target.value)} /></FG>
            <FG label="Reference"><input value={f.reference} onChange={e => set("reference")(e.target.value)} /></FG>
          </div>
        </Modal>
      )}

      {paymentModalFor && <AddCapitalPaymentModal project={paymentModalFor} onClose={() => setPaymentModalFor(null)} onSave={data => addPayment(paymentModalFor, data)} />}

      {linkTxnModalFor && (
        <Modal title={`Link Bank Transaction: ${linkTxnModalFor.vendorName}`} size="modal-lg" onClose={() => { setLinkTxnModalFor(null); setLinkTxnSearch(""); }}>
          <div style={{ marginBottom: 12 }}><SearchBar value={linkTxnSearch} onChange={setLinkTxnSearch} placeholder="Search narration…" style={{ minWidth: "100%" }} /></div>
          <p style={{ fontSize: 11, color: "var(--text3)", marginBottom: 10 }}>{fmt(linkTxnModalFor.due)} still due. Showing bank debits not already linked to another project.</p>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Date</th><th>Narration</th><th className="r">Amount</th><th className="r">Action</th></tr></thead>
              <tbody>
                {filteredUnlinkedTxns.length === 0 ? (
                  <tr><td colSpan={4} style={{ textAlign: "center", color: "var(--text3)", padding: 16 }}>No unlinked bank debits found</td></tr>
                ) : filteredUnlinkedTxns.slice(0, 100).map(t => (
                  <tr key={t.id}>
                    <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{t.date}</td>
                    <td style={{ fontSize: 12 }}>{t.description}</td>
                    <td className="r mono" style={{ fontWeight: 700 }}>{fmt(t.debit)}</td>
                    <td className="r"><button className="btn btn-primary btn-sm" onClick={() => linkBankTxn(linkTxnModalFor, t)}>Link</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Modal>
      )}

      {selectedForStmt && (
        <Modal title={`Project Statement: ${selectedForStmt.projectNo} — ${selectedForStmt.vendorName}`} size="modal-lg" onClose={() => setSelectedForStmt(null)}
          foot={<button className="btn btn-ghost" onClick={() => exportStatementPDF({
            title: `Capital Project Statement: ${selectedForStmt.projectNo}`,
            subtitle: `${selectedForStmt.vendorName} — ${selectedForStmt.description || ""}`,
            summary: [
              { label: "Total Cost", value: fmt(selectedForStmt.totalCost), color: [90, 160, 220] },
              { label: "Paid", value: fmt(selectedForStmt.paid), color: [70, 180, 100] },
              { label: "Due", value: fmt(selectedForStmt.due), color: [220, 150, 30] },
            ],
            columns: [{ key: "date", label: "Date" }, { key: "mode", label: "Mode" }, { key: "reference", label: "Reference" }, { key: "amount", label: "Amount", align: "right" }],
            rows: (selectedForStmt.payments || []).map(p => ({ date: p.date, mode: p.bankTxnId ? "Bank Txn" : p.paymentMode, reference: p.reference || p.remarks || "—", amount: fmt(p.amount) })),
            fileName: `${selectedForStmt.projectNo}_statement.pdf`,
          })}>📄 Export PDF</button>}
        >
          <div style={{ marginBottom: 16, display: "flex", gap: 20, fontFamily: "var(--mono)", fontSize: 13, background: "var(--bg3)", padding: 12, borderRadius: "var(--r)", flexWrap: "wrap" }}>
            <div>Total: <span style={{ fontWeight: 700 }}>{fmt(selectedForStmt.totalCost)}</span></div>
            <div>Paid: <span style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(selectedForStmt.paid)}</span></div>
            <div>Due: <span style={{ color: "var(--amber)", fontWeight: 700 }}>{fmt(selectedForStmt.due)}</span></div>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Date</th><th>Mode</th><th>Reference</th><th className="r">Amount</th></tr></thead>
              <tbody>{(selectedForStmt.payments || []).map(p => (
                <tr key={p.id}><td className="mono">{p.date}</td><td><BadgeComponent type={p.bankTxnId ? "blue" : "accent"}>{p.bankTxnId ? "Bank Txn" : p.paymentMode}</BadgeComponent></td><td>{p.reference || p.remarks || "—"}</td><td className="r mono" style={{ fontWeight: 700 }}>{fmt(p.amount)}</td></tr>
              ))}</tbody>
            </table>
          </div>
        </Modal>
      )}
    </div>
  );
}

function AddCapitalPaymentModal({ project, onClose, onSave }) {
  const [pf, setPf] = useState({ date: today(), amount: "", paymentMode: "bank", reference: "", remarks: "" });
  const setP = k => v => setPf(x => ({ ...x, [k]: v }));
  return (
    <Modal title={`Add Payment: ${project.vendorName}`} onClose={onClose} foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(pf)}>Save Payment</button></>}>
      <p style={{ fontSize: 11, color: "var(--text3)", marginBottom: 10 }}>{fmt(project.due)} still due on {project.projectNo}.</p>
      <div className="form-row cols-2">
        <FG label="Date"><input type="date" value={pf.date} onChange={e => setP("date")(e.target.value)} /></FG>
        <FG label="Amount (₹) *"><input type="number" value={pf.amount} onChange={e => setP("amount")(e.target.value)} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Mode"><select value={pf.paymentMode} onChange={e => setP("paymentMode")(e.target.value)}><option value="bank">Bank / UPI</option><option value="cash">Cash</option></select></FG>
        <FG label="Reference"><input value={pf.reference} onChange={e => setP("reference")(e.target.value)} /></FG>
      </div>
      <div className="form-row"><FG label="Remarks"><input value={pf.remarks} onChange={e => setP("remarks")(e.target.value)} /></FG></div>
    </Modal>
  );
}

// ── BANK STATEMENT GROUPING VIEW ─────────────────────────────────────────────
function CounterpartyNameInput({ rawKeys, currentLabel, placeholder, onSave }) {
  const [val, setVal] = useState(currentLabel || "");

  useEffect(() => {
    setVal(currentLabel || "");
  }, [currentLabel]);

  function commit() {
    if (val !== currentLabel) {
      onSave(rawKeys, { label: val.trim() });
    }
  }

  function handleKeyDown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      e.target.blur();
    }
  }

  return (
    <input
      placeholder={placeholder}
      value={val}
      onChange={e => setVal(e.target.value)}
      onBlur={commit}
      onKeyDown={handleKeyDown}
      title="Press Enter or click outside to save"
    />
  );
}

function BankStatementGroupingView({ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, capitalItems, setCapitalItems, setExpenses, setPartnerCashbook }) {
  const [sheet, setSheet] = useState(null);
  const [map, setMap] = useState({ dateCol:"", descCol:"", mode:"separate", debitCol:"", creditCol:"", amountCol:"" });
  const [search, setSearch] = useState("");
  const [unlabeledOnly, setUnlabeledOnly] = useState(false);
  const [expanded, setExpanded] = useState(null);

  async function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = "";

    if (file.name.toLowerCase().endsWith(".pdf")) {
      try {
        const parsedSheet = await parsePdfStatement(file);
        setSheet(parsedSheet);
        setMap({ dateCol: "Date", descCol: "Narration", mode: "separate", debitCol: "Withdrawal Amt", creditCol: "Deposit Amt", amountCol: "" });
      } catch (err) { alert("PDF Error: " + err.message); }
      return;
    }

    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const headers = Object.keys(rows[0] || {});
        setSheet({ headers, rows, fileName: file.name });
        setMap({
          dateCol: guessCol(headers, ["date"]),
          descCol: guessCol(headers, ["narration", "description", "particular"]),
          mode: "separate",
          debitCol: guessCol(headers, ["debit", "withdrawal"]),
          creditCol: guessCol(headers, ["credit", "deposit"]),
          amountCol: guessCol(headers, ["amount"]),
        });
      } catch (err) { alert("File read error: " + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  async function importRows() {
    if (!sheet || !map.dateCol || !map.descCol) return alert("Select required mapping columns");
    const parsed = sheet.rows.map(r => {
      const date = parseBankDate(r[map.dateCol]);
      const description = String(r[map.descCol] || "").trim();
      let debit = 0, credit = 0;
      if (map.mode === "separate") {
        debit = parseAmt(r[map.debitCol]);
        credit = parseAmt(r[map.creditCol]);
      } else {
        const a = parseAmt(r[map.amountCol]);
        if (String(r[map.amountCol]).includes("-")) debit = a; else credit = a;
      }
      return { date, description, debit, credit, key: normalizeDesc(description) };
    }).filter(r => r.date && r.description && (r.debit > 0 || r.credit > 0));

    try {
      const res = await db.importBankBatch(sheet.fileName, parsed);
      setBankBatches(bs => [res.batch, ...bs]);
      setBankTxns(txs => [...txs, ...res.txns]);
      setSheet(null);
      alert(`Imported ${res.txns.length} transactions and synchronized with Supabase!`);
    } catch (err) { alert(err.message); }
  }

  async function updateLabel(rawKeys, patch) {
    try {
      await db.updateBankLabels(rawKeys, patch);
      setBankLabels(bl => {
        const next = { ...bl };
        rawKeys.forEach(k => { next[k] = { ...next[k], ...patch }; });
        return next;
      });
      // Note: tagging a counterparty here as capital_* (or "owner", or any vendor
      // type) does NOT write a row into capitalItems/expenses/partnerCashbook.
      // Capital & Infra, Expenses, and Partners & Capital each derive their
      // bank-tagged rows live from bankTxns + bankLabels on every render (see
      // useCounterpartyGroups) — so the label change shows up there immediately,
      // and a transaction is never double-counted once as a live bank row and
      // again as a persisted manual entry with the same amount.
    } catch (err) { alert(err.message); }
  }

  async function logCashbook(t, partnerName, type, accountType) {
    try {
      const entry = await db.logPartnerCashbookEntry(t, partnerName, type, accountType);
      setPartnerCashbook(cb => [entry, ...cb]);
      alert(`Logged ${type} of ${fmt(entry.amount)} for ${partnerName}`);
    } catch (err) { alert(err.message); }
  }

  const groups = useMemo(() => {
    const m = {};
    bankTxns.forEach(t => {
      if (!m[t.key]) m[t.key] = { key: t.key, txns: [], totalDebit: 0, totalCredit: 0 };
      m[t.key].txns.push(t); m[t.key].totalDebit += t.debit; m[t.key].totalCredit += t.credit;
    });
    return Object.values(m).map(g => ({
      ...g,
      label: bankLabels[g.key]?.label || "",
      type: bankLabels[g.key]?.type || "unlabeled",
    }));
  }, [bankTxns, bankLabels]);

  const mergedGroups = useMemo(() => {
    const m = {};
    groups.forEach(g => {
      const mergeKey = g.label.trim() ? g.label.trim().toLowerCase() : `__raw__${g.key}`;
      if (!m[mergeKey]) m[mergeKey] = { mergeKey, label: g.label, type: g.type, rawKeys: [], txns: [], totalDebit: 0, totalCredit: 0 };
      m[mergeKey].rawKeys.push(g.key);
      m[mergeKey].txns.push(...g.txns);
      m[mergeKey].totalDebit += g.totalDebit;
      m[mergeKey].totalCredit += g.totalCredit;
    });
    return Object.values(m).sort((a,b) => (b.totalDebit + b.totalCredit) - (a.totalDebit + a.totalCredit));
  }, [groups]);

  const filtered = mergedGroups.filter(g => {
    if (unlabeledOnly && g.label) return false;
    if (!search) return true;
    const s = search.toLowerCase();
    return g.rawKeys.some(k => k.toLowerCase().includes(s)) || g.label.toLowerCase().includes(s) || g.txns.some(t => t.description.toLowerCase().includes(s));
  });

  const totalPaid = bankTxns.reduce((s, t) => s + t.debit, 0);
  const totalRecv = bankTxns.reduce((s, t) => s + t.credit, 0);

  return (
    <div>
      <div className="config-card">
        <h3>Import Net-Banking Statement</h3>
        {!sheet ? (
          <div style={{ display:"flex", alignItems:"center", gap:14, flexWrap:"wrap" }}>
            <label className="btn btn-primary" style={{ cursor:"pointer" }}>
              Choose PDF / CSV / XLSX File
              <input type="file" accept=".pdf,.csv,.xlsx,.xls" onChange={onFile} style={{ display:"none" }} />
            </label>
            <span style={{ fontSize:11.5, color:"var(--text3)" }}>Bank statement PDF (HDFC tested) or CSV/XLSX export from net-banking.</span>
          </div>
        ) : (
          <div>
            <div style={{ fontSize:12, color:"var(--text2)", marginBottom:12 }}>{sheet.fileName} — {sheet.rows.length} rows found. Confirm mapping:</div>
            <div className="config-row">
              <FG label="Date Column"><select value={map.dateCol} onChange={e => setMap({ ...map, dateCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Description Column"><select value={map.descCol} onChange={e => setMap({ ...map, descCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Debit Column"><select value={map.debitCol} onChange={e => setMap({ ...map, debitCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Credit Column"><select value={map.creditCol} onChange={e => setMap({ ...map, creditCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
            </div>
            <div style={{ display:"flex", gap:10, marginTop:14 }}>
              <button className="btn btn-ghost" onClick={() => setSheet(null)}>Cancel</button>
              <button className="btn btn-primary" onClick={importRows}>Import Transactions</button>
            </div>
          </div>
        )}
      </div>

      <div className="stats-grid">
        <StatCard label="Total Transactions" value={bankTxns.length} sub={`${bankBatches.length} file(s)`} color="blue" />
        <StatCard label="Total Paid Out" value={fmt(totalPaid)} sub="debits" color="red" />
        <StatCard label="Total Received" value={fmt(totalRecv)} sub="credits" color="green" />
        <StatCard label="Counterparty Groups" value={mergedGroups.length} sub="distinct names" color="purple" />
      </div>

      {bankTxns.length > 0 && (
        <>
          <div className="filter-bar">
            <input
              placeholder="Search by name or narration…"
              value={search}
              onChange={e => setSearch(e.target.value)}
              style={{ background:"var(--bg3)", border:"1px solid var(--border2)", borderRadius:"var(--r)", padding:"7px 11px", color:"var(--text)", fontSize:12.5, minWidth:220 }}
            />
            <button className={`btn btn-sm ${unlabeledOnly ? "btn-primary" : "btn-ghost"}`} onClick={() => setUnlabeledOnly(u => !u)}>
              Unlabeled Only
            </button>
          </div>

          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
            {filtered.map(g => (
              <div key={g.rawKeys[0]} className="config-card" style={{ padding: 14 }}>
                <div style={{ display:"flex", gap:12, flexWrap:"wrap", alignItems:"flex-end", justifyContent:"space-between" }}>
                  <div style={{ display:"flex", gap:12, flexWrap:"wrap", flex:1, minWidth:280 }}>
                    <FG label="Name this group" note="Press Enter to save name">
                      <CounterpartyNameInput
                        rawKeys={g.rawKeys}
                        currentLabel={g.label}
                        placeholder={g.rawKeys[0]}
                        onSave={updateLabel}
                      />
                    </FG>
                    <FG label="Type / Category Routing">
                      <select value={g.type} onChange={e => updateLabel(g.rawKeys, { type: e.target.value })}>
                        {CP_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
                      </select>
                    </FG>
                  </div>
                  <div style={{ display:"flex", gap:18, fontFamily:"var(--mono)", fontSize:12.5 }}>
                    <div><div style={{ color:"var(--text3)", fontSize:10 }}>PAID</div><div style={{ color:"var(--red)", fontWeight:700 }}>{fmt(g.totalDebit)}</div></div>
                    <div><div style={{ color:"var(--text3)", fontSize:10 }}>RECEIVED</div><div style={{ color:"var(--green)", fontWeight:700 }}>{fmt(g.totalCredit)}</div></div>
                    <div><div style={{ color:"var(--text3)", fontSize:10 }}>TXNS</div><div style={{ fontWeight:700 }}>{g.txns.length}</div></div>
                  </div>
                </div>

                <div style={{ marginTop:8, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ fontSize:11, color:"var(--text3)", fontFamily:"var(--mono)" }}>
                    {g.rawKeys.length > 1 ? `Merged from ${g.rawKeys.length} narration variations` : g.rawKeys[0]}
                  </span>
                  <button className="btn btn-ghost btn-sm" onClick={() => setExpanded(x => x === g.mergeKey ? null : g.mergeKey)}>
                    {expanded === g.mergeKey ? "Hide" : "Show"} transactions
                  </button>
                </div>

                {expanded === g.mergeKey && (
                  <div className="table-wrap" style={{ marginTop:10 }}>
                    <table>
                      <thead><tr><th>Date</th><th>Narration</th><th className="r">Debit</th><th className="r">Credit</th><th className="r">Actions</th></tr></thead>
                      <tbody>
                        {g.txns.sort((a,b) => b.date.localeCompare(a.date)).map(t => (
                          <tr key={t.id}>
                            <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{t.date}</td>
                            <td style={{ fontSize:12 }}>{t.description}</td>
                            <td className="r mono" style={{ fontSize:11, color:"var(--red)" }}>{t.debit > 0 ? fmt(t.debit) : "—"}</td>
                            <td className="r mono" style={{ fontSize:11, color:"var(--green)" }}>{t.credit > 0 ? fmt(t.credit) : "—"}</td>
                            <td className="r">
                              <div style={{ display:"flex", gap:4, justifyContent:"flex-end" }}>
                                {t.credit > 0 && g.type === "owner" && (
                                  <button className="btn btn-primary btn-sm" onClick={() => logCashbook(t, g.label || g.key, "Capital Infusion", "capital")}>+ Equity</button>
                                )}
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ── Loans: money borrowed (from a bank or anyone else) and what's been repaid ──
//  received  = a bank-statement credit you mark as a loan
//  direct    = the lender paid a vendor directly (not through your bank); added by hand and
//              optionally booked against a Vendor Project so it counts as paid there
//  repayment = money paid back to the lender
// Bank-statement counterparties already tagged "Loan / Bank" are included automatically.
const LOAN_KINDS = {
  received: { label: "Received in bank", badge: "green" },
  direct: { label: "Paid to vendor", badge: "purple" },
  repayment: { label: "Repayment", badge: "amber" },
};

function LoansView({ bankTxns, bankLabels, loanEntries, setLoanEntries, loanStatus, vendorProjects, setProjectPayments }) {
  const [modal, setModal] = useState(null); // "mark" | "direct" | "repay"
  const groups = useCounterpartyGroups(bankTxns, bankLabels);
  const ready = loanStatus?.ready !== false;

  const legacy = useMemo(() => groups.filter(g => g.type === "loan"), [groups]);
  const legacyTxnIds = useMemo(() => new Set(legacy.flatMap(g => g.txns.map(t => String(t.id)))), [legacy]);
  const markedIds = useMemo(() => new Set((loanEntries || []).filter(e => e.bankTxnId != null).map(e => String(e.bankTxnId))), [loanEntries]);

  const lenders = useMemo(() => {
    const m = {};
    const get = name => {
      const n = (name || "").trim() || "Unnamed lender";
      const k = n.toLowerCase();
      if (!m[k]) m[k] = { name: n, received: 0, direct: 0, repaid: 0 };
      return m[k];
    };
    legacy.forEach(g => { const l = get(g.label); l.received += g.totalCredit; l.repaid += g.totalDebit; });
    (loanEntries || []).forEach(e => {
      const l = get(e.lender);
      if (e.kind === "received") l.received += e.amount;
      else if (e.kind === "direct") l.direct += e.amount;
      else l.repaid += e.amount;
    });
    return Object.values(m)
      .map(l => ({ ...l, outstanding: l.received + l.direct - l.repaid }))
      .sort((a, b) => b.outstanding - a.outstanding);
  }, [legacy, loanEntries]);

  const tot = lenders.reduce((a, l) => ({ received: a.received + l.received, direct: a.direct + l.direct, repaid: a.repaid + l.repaid, outstanding: a.outstanding + l.outstanding }), { received: 0, direct: 0, repaid: 0, outstanding: 0 });
  const lenderNames = lenders.map(l => l.name);
  const vendorNames = useMemo(() => [...new Set((vendorProjects || []).map(p => p.vendorName).filter(Boolean))].sort((a, b) => a.localeCompare(b)), [vendorProjects]);

  async function addEntry(data) {
    try {
      const row = await loansdb.saveLoanEntry(data);
      setLoanEntries(es => [row, ...es]);
      return true;
    } catch (err) { alert(err.message); return false; }
  }

  async function markReceipt(txn, lender) {
    if (!(lender || "").trim()) return alert("Enter the lender's name first.");
    return addEntry({ date: txn.date, kind: "received", lender, amount: txn.credit, bankTxnId: txn.id, note: txn.description });
  }

  async function saveDirect(d) {
    if (!d.lender.trim()) return alert("Enter the lender's name.");
    if (!d.vendorName.trim()) return alert("Enter the vendor who was paid.");
    if (!(+d.amount > 0)) return alert("Enter an amount greater than zero.");
    const project = d.projectId ? (vendorProjects || []).find(p => String(p.id) === d.projectId) : null;
    try {
      let pay = null;
      if (project) {
        pay = await vpdb.addProjectPayment({
          projectId: project.id, amount: +d.amount, date: d.date, paymentMode: "bank",
          reference: `Loan: ${d.lender.trim()}`, remarks: d.note.trim() || `Paid directly by ${d.lender.trim()}`,
        });
        setProjectPayments(pp => [pay, ...pp]);
      }
      const ok = await addEntry({ date: d.date, kind: "direct", lender: d.lender, amount: d.amount, vendorName: d.vendorName, projectId: project ? project.id : null, projectPaymentId: pay ? pay.id : null, note: d.note });
      if (ok) setModal(null);
    } catch (err) { alert(err.message); }
  }

  async function saveRepayment(d) {
    if (!d.lender.trim()) return alert("Enter the lender's name.");
    if (!(+d.amount > 0)) return alert("Enter an amount greater than zero.");
    if (await addEntry({ date: d.date, kind: "repayment", lender: d.lender, amount: d.amount, note: d.note })) setModal(null);
  }

  async function removeEntry(e) {
    const hasPay = e.projectPaymentId != null;
    if (!confirm(`Delete this ${LOAN_KINDS[e.kind]?.label.toLowerCase() || "loan"} entry of ${fmt(e.amount)}?${hasPay ? "\n\nThe matching payment on the vendor project is removed too." : ""}`)) return;
    try {
      if (hasPay) {
        try { await vpdb.deleteProjectPayment(e.projectPaymentId); } catch (_) { /* already removed from the project */ }
        setProjectPayments(pp => pp.filter(x => String(x.id) !== String(e.projectPaymentId)));
      }
      await loansdb.deleteLoanEntry(e.id);
      setLoanEntries(es => es.filter(x => x.id !== e.id));
    } catch (err) { alert(err.message); }
  }

  const ledger = [...(loanEntries || [])].sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  return (
    <div>
      {!ready && (
        <div style={{ background: "var(--bg3)", border: "1px solid var(--amber)", borderRadius: "var(--r)", padding: 14, marginBottom: 16, fontSize: 12.5 }}>
          <div style={{ fontWeight: 700, marginBottom: 4 }}>One-time setup: create the loans table</div>
          Run this in the Supabase SQL editor, then refresh the app. If your other tables use row-level-security policies, add the same policy to <span className="mono">loan_entries</span>.
          <pre className="mono" style={{ margin: "8px 0 0", whiteSpace: "pre-wrap", fontSize: 11 }}>{loansdb.LOANS_SETUP_SQL}</pre>
          {loanStatus?.error && <div className="mono" style={{ color: "var(--text3)", fontSize: 10.5, marginTop: 6 }}>{loanStatus.error}</div>}
        </div>
      )}

      <div className="stats-grid-5">
        <StatCard label="Loan Drawn" value={fmt(tot.received + tot.direct)} color="blue" sub="Received + paid to vendors" />
        <StatCard label="Received in Bank" value={fmt(tot.received)} color="green" sub="Marked bank receipts" />
        <StatCard label="Paid to Vendors" value={fmt(tot.direct)} color="purple" sub="Lender paid directly" />
        <StatCard label="Repaid" value={fmt(tot.repaid)} color="amber" sub="Principal + interest paid" />
        <StatCard label="Outstanding" value={fmt(tot.outstanding)} color={tot.outstanding > 0 ? "red" : "green"} sub="Drawn − repaid" />
      </div>

      <div className="filter-bar" style={{ marginBottom: 14 }}>
        <button className="btn btn-primary" disabled={!ready} onClick={() => setModal("mark")}>🏦 Mark bank receipt as loan</button>
        <button className="btn btn-primary" disabled={!ready} onClick={() => setModal("direct")}>+ Lender paid a vendor directly</button>
        <button className="btn btn-ghost" disabled={!ready} onClick={() => setModal("repay")}>+ Record repayment</button>
        <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("loans", ledger.map(e => ({ Date: e.date, Type: LOAN_KINDS[e.kind]?.label || e.kind, Lender: e.lender, Vendor: e.vendorName, Amount: e.amount, Note: e.note })))}>⬇ Export Excel</button>
      </div>

      <div className="table-wrap" style={{ marginBottom: 18 }}>
        <div className="table-toolbar"><h3>By lender</h3></div>
        <table>
          <thead><tr><th>Lender</th><th className="r">Received in bank</th><th className="r">Paid to vendors</th><th className="r">Repaid</th><th className="r">Outstanding</th></tr></thead>
          <tbody>
            {lenders.length === 0 ? (
              <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--text3)", padding: 24 }}>No loans recorded yet</td></tr>
            ) : lenders.map(l => (
              <tr key={l.name}>
                <td style={{ fontWeight: 600 }}>{l.name}</td>
                <td className="r mono">{fmt(l.received)}</td>
                <td className="r mono">{fmt(l.direct)}</td>
                <td className="r mono">{fmt(l.repaid)}</td>
                <td className="r mono" style={{ fontWeight: 700, color: l.outstanding > 0 ? "var(--red)" : "var(--green)" }}>{fmt(l.outstanding)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {legacy.length > 0 && <div className="mono" style={{ fontSize: 11, color: "var(--text3)", padding: "8px 16px 12px" }}>Includes bank-statement counterparties tagged "Loan / Bank": {legacy.map(g => g.label).join(", ")} (credits = received, debits = repaid).</div>}
      </div>

      <div className="table-wrap">
        <div className="table-toolbar"><h3>Loan entries</h3></div>
        <table>
          <thead><tr><th>Date</th><th>Type</th><th>Lender</th><th>Vendor</th><th>Note</th><th className="r">Amount</th><th className="r">Actions</th></tr></thead>
          <tbody>
            {ledger.length === 0 ? (
              <tr><td colSpan={7} style={{ textAlign: "center", color: "var(--text3)", padding: 24 }}>Nothing here yet — mark a bank receipt as a loan, or add a payment the lender made straight to a vendor.</td></tr>
            ) : ledger.map(e => (
              <tr key={e.id}>
                <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{e.date}</td>
                <td><BadgeComponent type={LOAN_KINDS[e.kind]?.badge}>{LOAN_KINDS[e.kind]?.label || e.kind}</BadgeComponent></td>
                <td style={{ fontWeight: 600 }}>{e.lender}</td>
                <td style={{ fontSize: 12 }}>{e.vendorName || "—"}{e.projectPaymentId != null && <div style={{ color: "var(--text3)", fontSize: 10.5 }}>linked to a vendor project</div>}</td>
                <td style={{ fontSize: 12, color: "var(--text2)" }}>{e.note || "—"}</td>
                <td className="r mono" style={{ fontWeight: 700 }}>{fmt(e.amount)}</td>
                <td className="r"><RowActions onDelete={() => removeEntry(e)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {modal === "mark" && <MarkLoanReceiptModal bankTxns={bankTxns} bankLabels={bankLabels} hidden={id => legacyTxnIds.has(id) || markedIds.has(id)} lenderNames={lenderNames} onMark={markReceipt} onClose={() => setModal(null)} />}
      {modal === "direct" && <LoanDirectModal lenderNames={lenderNames} vendorNames={vendorNames} projects={vendorProjects || []} onSave={saveDirect} onClose={() => setModal(null)} />}
      {modal === "repay" && <LoanRepaymentModal lenderNames={lenderNames} onSave={saveRepayment} onClose={() => setModal(null)} />}
    </div>
  );
}

function MarkLoanRow({ txn, defaultLender, lenderNames, onMark }) {
  const [lender, setLender] = useState(defaultLender);
  const [busy, setBusy] = useState(false);
  return (
    <tr>
      <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{txn.date}</td>
      <td style={{ fontSize: 12 }}>{txn.description}</td>
      <td className="r mono" style={{ fontWeight: 700, color: "var(--green)" }}>{fmt(txn.credit)}</td>
      <td className="r"><input list="loan-lenders" value={lender} onChange={e => setLender(e.target.value)} placeholder="Lender" style={{ width: 150, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: "var(--r)", padding: "5px 8px", color: "var(--text)" }} /></td>
      <td className="r"><button className="btn btn-primary btn-sm" disabled={busy} onClick={async () => { setBusy(true); await onMark(txn, lender); setBusy(false); }}>Mark as loan</button></td>
    </tr>
  );
}

function MarkLoanReceiptModal({ bankTxns, bankLabels, hidden, lenderNames, onMark, onClose }) {
  const [search, setSearch] = useState("");
  const rows = useMemo(() => {
    const s = search.toLowerCase();
    return (bankTxns || [])
      .filter(t => t && t.credit > 0 && !hidden(String(t.id)))
      .filter(t => !s || (t.description || "").toLowerCase().includes(s) || String(t.credit).includes(s))
      .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
      .slice(0, 200);
  }, [bankTxns, search, hidden]);
  return (
    <Modal title="Mark a bank receipt as a loan" size="modal-lg" onClose={onClose}>
      <p style={{ fontSize: 11.5, color: "var(--text3)", marginBottom: 10, lineHeight: 1.5 }}>
        Money credited to your account that is really a loan. Type the lender and click Mark as loan — it then counts as loan received. Showing the 200 most recent credits; use search to find older ones.
      </p>
      <div className="filter-bar" style={{ marginBottom: 10 }}><SearchBar value={search} onChange={setSearch} placeholder="Search narration or amount…" style={{ minWidth: 240 }} /></div>
      <datalist id="loan-lenders">{lenderNames.map(n => <option key={n} value={n} />)}</datalist>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Date</th><th>Narration</th><th className="r">Credit</th><th className="r">Lender</th><th className="r"></th></tr></thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--text3)", padding: 20 }}>No unmarked bank credits found</td></tr>
            ) : rows.map(t => <MarkLoanRow key={t.id} txn={t} defaultLender={(bankLabels?.[t.key]?.label || "").trim()} lenderNames={lenderNames} onMark={onMark} />)}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

function LoanDirectModal({ lenderNames, vendorNames, projects, onSave, onClose }) {
  const [f, setF] = useState({ date: today(), lender: "", vendorName: "", amount: "", projectId: "", note: "" });
  const set = k => v => setF(x => ({ ...x, [k]: v, ...(k === "vendorName" ? { projectId: "" } : {}) }));
  const vendorProjects = projects.filter(p => normVendor(p.vendorName) === normVendor(f.vendorName));
  return (
    <Modal title="Lender paid a vendor directly" onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(f)}>Save</button></>}>
      <p style={{ fontSize: 11.5, color: "var(--text3)", marginBottom: 12, lineHeight: 1.5 }}>
        For when the bank pays the vendor straight from the loan, so it never shows in your bank statement. It counts as loan drawn, and if you pick a project it is also booked as a payment on that project.
      </p>
      <datalist id="loan-lenders-d">{lenderNames.map(n => <option key={n} value={n} />)}</datalist>
      <datalist id="loan-vendors-d">{vendorNames.map(n => <option key={n} value={n} />)}</datalist>
      <div className="form-row cols-2">
        <FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG>
        <FG label="Amount (₹) *"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Lender *"><input list="loan-lenders-d" value={f.lender} onChange={e => set("lender")(e.target.value)} placeholder="e.g. HDFC Bank" /></FG>
        <FG label="Vendor paid *"><input list="loan-vendors-d" value={f.vendorName} onChange={e => set("vendorName")(e.target.value)} /></FG>
      </div>
      <div className="form-row">
        <FG label="Vendor project (optional)" note={vendorProjects.length === 0 && f.vendorName ? "No project found for this vendor name." : "Pick one to count this as paid on the project."}>
          <select value={f.projectId} onChange={e => set("projectId")(e.target.value)}>
            <option value="">Don't link to a project</option>
            {vendorProjects.map(p => <option key={p.id} value={String(p.id)}>{p.projectNo} — {p.title}</option>)}
          </select>
        </FG>
      </div>
      <div className="form-row"><FG label="Note"><input value={f.note} onChange={e => set("note")(e.target.value)} placeholder="e.g. Machinery stage 2, loan disbursement" /></FG></div>
    </Modal>
  );
}

function LoanRepaymentModal({ lenderNames, onSave, onClose }) {
  const [f, setF] = useState({ date: today(), lender: lenderNames[0] || "", amount: "", note: "" });
  const set = k => v => setF(x => ({ ...x, [k]: v }));
  return (
    <Modal title="Record a loan repayment" onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(f)}>Save</button></>}>
      <datalist id="loan-lenders-r">{lenderNames.map(n => <option key={n} value={n} />)}</datalist>
      <div className="form-row cols-2">
        <FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG>
        <FG label="Amount (₹) *"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG>
      </div>
      <div className="form-row"><FG label="Lender *"><input list="loan-lenders-r" value={f.lender} onChange={e => set("lender")(e.target.value)} /></FG></div>
      <div className="form-row"><FG label="Note"><input value={f.note} onChange={e => set("note")(e.target.value)} placeholder="e.g. EMI for October" /></FG></div>
      <p style={{ fontSize: 11, color: "var(--text3)" }}>Repayments that already appear on the bank statement under a "Loan / Bank" counterparty are counted automatically — don't add those again.</p>
    </Modal>
  );
}


function PartnersCapitalDashboard({ bankTxns, bankLabels, partnerCashbook, setPartnerCashbook, setExpenses, setCapitalItems }) {
  const [activeTab, setActiveTab] = useState("overview");
  const [selectedPartner, setSelectedPartner] = useState("All");
  const [openCashModal, setOpenCashModal] = useState(false);
  const [openActionModal, setOpenActionModal] = useState(null);
  const [selectedPartnerForStmt, setSelectedPartnerForStmt] = useState(null);

  const allBankGroups = useCounterpartyGroups(bankTxns, bankLabels);

  const bankPartnerGroups = useMemo(() => {
    const merged = {};
    allBankGroups.forEach(g => {
      if (g.type !== "owner") return;
      merged[g.label.toLowerCase()] = { displayName: g.label, totalDebit: g.totalDebit, totalCredit: g.totalCredit, txns: g.txns };
    });
    return merged;
  }, [allBankGroups]);

  const partners = useMemo(() => {
    const s = new Set();
    Object.values(bankPartnerGroups).forEach(g => s.add(g.displayName));
    (partnerCashbook || []).forEach(e => {
      if (e && e.partnerName) s.add(e.partnerName.trim());
    });
    return Array.from(s);
  }, [bankPartnerGroups, partnerCashbook]);

  const partnerSummaries = useMemo(() => {
    return partners.map(name => {
      const bankGroup = bankPartnerGroups[name.toLowerCase()] || { totalCredit: 0, totalDebit: 0, txns: [] };

      const cashbookRecords = (partnerCashbook || []).filter(
        e => e && e.partnerName && e.partnerName.trim().toLowerCase() === name.toLowerCase()
      );

      const manualInfusions = cashbookRecords
        .filter(e => !e.bankTxnId && (e.accountType === "capital" || e.type === "Capital Infusion"))
        .reduce((sum, e) => sum + (+e.amount || 0), 0);

      const manualDrawings = cashbookRecords
        .filter(e => !e.bankTxnId && e.type === "Drawings")
        .reduce((sum, e) => sum + (+e.amount || 0), 0);

      const cashSpentOnSite = cashbookRecords
        .filter(e => e.type === "Direct Cash Expense" || e.type === "Expense Reimbursement")
        .reduce((sum, e) => sum + (+e.amount || 0), 0);

      const capitalInjected = bankGroup.totalCredit + manualInfusions;
      const drawingsTaken = bankGroup.totalDebit + manualDrawings;
      const netStanding = (capitalInjected + cashSpentOnSite) - drawingsTaken;

      return { name, capitalInjected, cashSpentOnSite, drawingsTaken, netStanding };
    }).sort((a, b) => b.capitalInjected - a.capitalInjected);
  }, [partners, bankPartnerGroups, partnerCashbook]);

  const totalCapitalAll = partnerSummaries.reduce((s, p) => s + p.capitalInjected, 0);
  const totalCashSpentAll = partnerSummaries.reduce((s, p) => s + p.cashSpentOnSite, 0);
  const totalDrawingsAll = partnerSummaries.reduce((s, p) => s + p.drawingsTaken, 0);

  const [cf, setCf] = useState({ partnerName: "", date: today(), amount: "", isCapex: false, category: "machinery", vendorName: "", description: "", ref: "" });

  useEffect(() => {
    if (!cf.partnerName && partners.length > 0) {
      setCf(prev => ({ ...prev, partnerName: partners[0] }));
    }
  }, [partners, cf.partnerName]);

  async function handleDirectCashSubmit() {
    if (!cf.partnerName || !cf.amount || !cf.description) return alert("Partner name, amount, and description are required.");
    if (!cf.vendorName) return alert("Paid To / Vendor is required so this can be grouped correctly.");
    try {
      const res = await db.recordPartnerDirectCashExpense(cf.partnerName, cf.date, cf.category, cf.amount, cf.description, cf.ref, cf.isCapex, cf.vendorName);
      setPartnerCashbook(prev => [res.cashbookEntry, ...(prev || [])]);
      if (res.isCapex) setCapitalItems(prev => [res.item, ...(prev || [])]);
      else setExpenses(prev => [res.item, ...(prev || [])]);
      setOpenCashModal(false);
      setCf({ partnerName: partners[0] || "", date: today(), amount: "", isCapex: false, category: "machinery", vendorName: "", description: "", ref: "" });
      alert(`Logged ${fmt(cf.amount)} paid by ${cf.partnerName}!`);
    } catch (err) { alert(err.message); }
  }

  const [quickForm, setQuickForm] = useState({ date: today(), amount: "", remarks: "" });
  async function handleQuickActionSubmit() {
    if (!quickForm.amount) return alert("Amount is required.");
    try {
      const entry = await db.savePartnerCashbookEntry({
        partnerName: openActionModal.partner,
        entryDate: quickForm.date,
        accountType: openActionModal.type === "Capital Infusion" ? "capital" : "current",
        type: openActionModal.type,
        amount: +quickForm.amount,
        paymentMode: "cash",
        remarks: quickForm.remarks || `${openActionModal.type} recorded directly`,
      });
      setPartnerCashbook(prev => [entry, ...(prev || [])]);
      setOpenActionModal(null);
      setQuickForm({ date: today(), amount: "", remarks: "" });
    } catch (err) { alert(err.message); }
  }

  const ledgerRows = useMemo(() => {
    const list = [];
    (bankTxns || []).forEach(t => {
      if (!t) return;
      const k = t.key || t.group_key || (t.description ? normalizeDesc(t.description) : "");
      const meta = (bankLabels || {})[k];
      if (meta && meta.type === "owner" && meta.label) {
        const pName = meta.label.trim();
        if (selectedPartner === "All" || pName.toLowerCase() === selectedPartner.toLowerCase()) {
          list.push({
            id: `bank-${t.id}`,
            entryDate: t.date || t.txn_date,
            partnerName: pName,
            accountType: t.credit > 0 ? "capital" : "current",
            type: t.credit > 0 ? "Capital Infusion" : "Drawings",
            remarks: t.description,
            paymentMode: "bank",
            amount: t.credit > 0 ? +t.credit : +t.debit,
          });
        }
      }
    });

    (partnerCashbook || []).forEach(e => {
      if (!e || e.bankTxnId) return;
      if (selectedPartner === "All" || (e.partnerName && e.partnerName.toLowerCase() === selectedPartner.toLowerCase())) {
        list.push({
          id: `cb-${e.id}`,
          entryDate: e.entryDate,
          partnerName: e.partnerName,
          accountType: e.accountType || "current",
          type: e.type,
          remarks: e.remarks,
          paymentMode: e.paymentMode || "cash",
          amount: +e.amount,
        });
      }
    });
    return list.sort((a, b) => (b.entryDate || "").localeCompare(a.entryDate || ""));
  }, [bankTxns, bankLabels, partnerCashbook, selectedPartner]);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div className="pill-tabs">
          <button className={`pill-tab ${activeTab === "overview" ? "active" : ""}`} onClick={() => setActiveTab("overview")}>Partner Equity & Summaries</button>
          <button className={`pill-tab ${activeTab === "ledger" ? "active" : ""}`} onClick={() => setActiveTab("ledger")}>Transaction Ledger ({ledgerRows.length})</button>
        </div>
        <button className="btn btn-primary" onClick={() => setOpenCashModal(true)}>+ Record Cash Spent by Partner</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Total Capital Infused" value={fmt(totalCapitalAll)} color="accent" sub="Core Equity" />
        <StatCard label="Direct Cash Spent on Site" value={fmt(totalCashSpentAll)} color="green" sub="Paid On Behalf" />
        <StatCard label="Drawings / Withdrawals" value={fmt(totalDrawingsAll)} color="red" sub="Taken Out" />
        <StatCard label="Net Firm Obligation" value={fmt((totalCapitalAll + totalCashSpentAll) - totalDrawingsAll)} color="blue" sub="Total Net Standing" />
      </div>

      {activeTab === "overview" ? (
        <div className="table-wrap">
          <div className="table-toolbar">
            <h3>Partner Accounts Summary</h3>
            <span style={{ fontSize: 11, color: "var(--text3)" }}>{partners.length} registered partner(s)</span>
          </div>
          <table>
            <thead>
              <tr>
                <th>Partner Name</th>
                <th className="r">Capital Contributed</th>
                <th className="r">Cash Spent on Site</th>
                <th className="r">Drawings / Salary Taken</th>
                <th className="r">Net Equity Standing</th>
                <th className="r">Actions</th>
              </tr>
            </thead>
            <tbody>
              {partnerSummaries.length === 0 ? (
                <tr>
                  <td colSpan={6} style={{ textAlign: "center", color: "var(--text3)", padding: 32 }}>
                    <EmptyState icon="🤝" message="No partners detected yet" sub="Tag counterparties as 'Owner / Capital' in Bank Statement." />
                  </td>
                </tr>
              ) : (
                partnerSummaries.map(p => (
                  <tr key={p.name}>
                    <td style={{ fontWeight: 600 }}>{p.name}</td>
                    <td className="r mono" style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(p.capitalInjected)}</td>
                    <td className="r mono" style={{ color: "var(--green)" }}>{fmt(p.cashSpentOnSite)}</td>
                    <td className="r mono" style={{ color: "var(--red)" }}>{fmt(p.drawingsTaken)}</td>
                    <td className="r mono" style={{ fontWeight: 700, fontSize: 13, color: p.netStanding >= 0 ? "var(--teal)" : "var(--amber)" }}>{fmt(p.netStanding)}</td>
                    <td className="r">
                      <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                        <button className="btn btn-info btn-sm" onClick={() => setSelectedPartnerForStmt(p)}>📄 Statement</button>
                        <button className="btn btn-ghost btn-sm" onClick={() => setOpenActionModal({ partner: p.name, type: "Capital Infusion" })}>+ Capital</button>
                        <button className="btn btn-danger btn-sm" onClick={() => setOpenActionModal({ partner: p.name, type: "Drawings" })}>+ Draw</button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      ) : (
        <div>
          <div className="filter-bar">
            <FG label="Filter by Partner">
              <select value={selectedPartner} onChange={e => setSelectedPartner(e.target.value)}>
                <option value="All">All Partners</option>
                {partners.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            </FG>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Partner</th>
                  <th>Account</th>
                  <th>Type</th>
                  <th>Narration / Remarks</th>
                  <th>Channel</th>
                  <th className="r">Amount</th>
                </tr>
              </thead>
              <tbody>
                {ledgerRows.length === 0 ? (
                  <tr><td colSpan={7} style={{ textAlign: "center", color: "var(--text3)", padding: 32 }}><EmptyState icon="📖" message="No ledger transactions found" /></td></tr>
                ) : (
                  ledgerRows.map(e => (
                    <tr key={e.id}>
                      <td className="mono">{e.entryDate}</td>
                      <td style={{ fontWeight: 600 }}>{e.partnerName}</td>
                      <td><BadgeComponent type={e.accountType === "capital" ? "accent" : "muted"}>{e.accountType}</BadgeComponent></td>
                      <td><BadgeComponent type={e.type === "Drawings" ? "red" : "green"}>{e.type}</BadgeComponent></td>
                      <td style={{ fontSize: 12 }}>{e.remarks}</td>
                      <td><BadgeComponent type="muted">{e.paymentMode}</BadgeComponent></td>
                      <td className="r mono" style={{ fontWeight: 700, color: e.type === "Drawings" ? "var(--red)" : "var(--green)" }}>{e.type === "Drawings" ? "-" : "+"}{fmt(e.amount)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {selectedPartnerForStmt && (
        <Modal 
          title={`Partner Capital Account: ${selectedPartnerForStmt.name}`} 
          size="modal-lg" 
          onClose={() => setSelectedPartnerForStmt(null)} 
          foot={
            <div style={{ display: "flex", gap: 8 }}>
              <button className="btn btn-ghost" onClick={() => exportStatementPDF({
                title: `Partner Capital Account: ${selectedPartnerForStmt.name}`,
                summary: [
                  { label: "Capital Injected", value: fmt(selectedPartnerForStmt.capitalInjected), color: [190, 190, 90] },
                  { label: "Site Cash Spent", value: fmt(selectedPartnerForStmt.cashSpentOnSite), color: [70, 180, 100] },
                  { label: "Drawings", value: fmt(selectedPartnerForStmt.drawingsTaken), color: [220, 70, 70] },
                  { label: "Net Standing", value: fmt(selectedPartnerForStmt.netStanding), color: [40, 180, 160] },
                ],
                columns: [
                  { key: "date", label: "Date" },
                  { key: "particulars", label: "Activity / Particulars" },
                  { key: "amount", label: "Amount", align: "right" },
                ],
                rows: ledgerRows
                  .filter(r => r.partnerName.toLowerCase() === selectedPartnerForStmt.name.toLowerCase())
                  .map(r => ({
                    date: r.entryDate,
                    particulars: `${r.type}: ${r.remarks}`,
                    amount: `${r.type === "Drawings" ? "-" : "+"}${fmt(r.amount)}`,
                  })),
                fileName: `${selectedPartnerForStmt.name}_capital_account.pdf`,
              })}>📄 Export PDF</button>
              <button className="btn btn-primary" onClick={() => alert("Partner statement ready to share via WhatsApp!")}>💬 Share via WhatsApp</button>
            </div>
          }
        >
          <div style={{ marginBottom: 16 }}>
            <div style={{ display: "flex", gap: 16, fontFamily: "var(--mono)", fontSize: 12.5, background: "var(--bg3)", padding: 12, borderRadius: "var(--r)", flexWrap: "wrap" }}>
              <div>Capital Injected: <span style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(selectedPartnerForStmt.capitalInjected)}</span></div>
              <div>Site Cash Spent: <span style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(selectedPartnerForStmt.cashSpentOnSite)}</span></div>
              <div>Drawings: <span style={{ color: "var(--red)", fontWeight: 700 }}>{fmt(selectedPartnerForStmt.drawingsTaken)}</span></div>
              <div>Net Standing: <span style={{ color: "var(--teal)", fontWeight: 700 }}>{fmt(selectedPartnerForStmt.netStanding)}</span></div>
            </div>
          </div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Activity / Particulars</th>
                  <th className="r">Amount</th>
                </tr>
              </thead>
              <tbody>
                {ledgerRows.filter(r => r.partnerName.toLowerCase() === selectedPartnerForStmt.name.toLowerCase()).map(r => (
                  <tr key={r.id}>
                    <td className="mono">{r.entryDate}</td>
                    <td>{r.type}: {r.remarks}</td>
                    <td className="r mono" style={{ fontWeight: 700, color: r.type === "Drawings" ? "var(--red)" : "var(--green)" }}>{r.type === "Drawings" ? "-" : "+"}{fmt(r.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Modal>
      )}

      {openCashModal && (
        <Modal title="Record Site Cash Expense Paid by Partner" onClose={() => setOpenCashModal(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpenCashModal(false)}>Cancel</button><button className="btn btn-primary" onClick={handleDirectCashSubmit}>Save Entry</button></>}>
          <div className="form-row cols-2">
            <FG label="Partner Name *">
              <input list="partner-suggestions" value={cf.partnerName} placeholder="Select or type partner name" onChange={e => setCf({ ...cf, partnerName: e.target.value })} />
              <datalist id="partner-suggestions">{partners.map(p => <option key={p} value={p} />)}</datalist>
            </FG>
            <FG label="Date *"><input type="date" value={cf.date} onChange={e => setCf({ ...cf, date: e.target.value })} /></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Amount Paid (₹) *"><input type="number" value={cf.amount} onChange={e => setCf({ ...cf, amount: e.target.value })} /></FG>
            <FG label="Expense Nature">
              <select value={cf.isCapex ? "capex" : "opex"} onChange={e => setCf({ ...cf, isCapex: e.target.value === "capex" })}>
                <option value="opex">Operational Overhead</option>
                <option value="capex">Fixed Asset / Land</option>
              </select>
            </FG>
          </div>
          <div className="form-row">
            <FG label="Paid To / Vendor *" note="Match an existing vendor name exactly to group this into that vendor's card in Capital & Infra / Expenses">
              <input value={cf.vendorName} onChange={e => setCf({ ...cf, vendorName: e.target.value })} placeholder="e.g. Pole Venketeash" />
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Category">
              <select value={cf.category} onChange={e => setCf({ ...cf, category: e.target.value })}>
                {(cf.isCapex ? CAP_CATS : EXP_CATS).map(c => <option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </FG>
            <FG label="Voucher / Cash Memo Ref"><input value={cf.ref} onChange={e => setCf({ ...cf, ref: e.target.value })} /></FG>
          </div>
          <div className="form-row"><FG label="Description *"><input value={cf.description} onChange={e => setCf({ ...cf, description: e.target.value })} /></FG></div>
        </Modal>
      )}

      {openActionModal && (
        <Modal title={`${openActionModal.type}: ${openActionModal.partner}`} onClose={() => setOpenActionModal(null)} foot={<><button className="btn btn-ghost" onClick={() => setOpenActionModal(null)}>Cancel</button><button className="btn btn-primary" onClick={handleQuickActionSubmit}>Record</button></>}>
          <div className="form-row cols-2">
            <FG label="Date"><input type="date" value={quickForm.date} onChange={e => setQuickForm({ ...quickForm, date: e.target.value })} /></FG>
            <FG label="Amount (₹) *"><input type="number" value={quickForm.amount} onChange={e => setQuickForm({ ...quickForm, amount: e.target.value })} /></FG>
          </div>
          <div className="form-row"><FG label="Remarks / Purpose"><input value={quickForm.remarks} onChange={e => setQuickForm({ ...quickForm, remarks: e.target.value })} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

// ── BALANCE SHEET VIEW ─────────────────────────────────────────────────────
function BalanceSheetView({ capitalItems, bankTxns, bankLabels, partnerCashbook, purchases, expenses }) {
  const fixedAssets = useMemo(() => {
    return (capitalItems || []).reduce((acc, c) => {
      acc[c.category] = (acc[c.category] || 0) + (c.amount || 0);
      return acc;
    }, {});
  }, [capitalItems]);
  const totalFixedAssets = Object.values(fixedAssets).reduce((s, a) => s + a, 0);

  const bankBalance = useMemo(() => {
    return (bankTxns || []).reduce((s, t) => s + ((t.credit || 0) - (t.debit || 0)), 0);
  }, [bankTxns]);

  const totalCurrentAssets = bankBalance > 0 ? bankBalance : 0;
  const totalAssets = totalFixedAssets + totalCurrentAssets;

  const loanCapital = useMemo(() => {
    return (capitalItems || []).filter(c => c.fundedBy === "loan").reduce((s, c) => s + (c.amount || 0), 0);
  }, [capitalItems]);
  const bankOverdraft = bankBalance < 0 ? Math.abs(bankBalance) : 0;
  const totalLiabilities = loanCapital + bankOverdraft;

  const partnerEquity = useMemo(() => {
    const raw = {};
    (partnerCashbook || []).forEach(e => {
      const delta = (e.type === "Capital Infusion" || e.type === "Direct Cash Expense" || e.type === "Expense Reimbursement") ? (e.amount || 0) : -(e.amount || 0);
      raw[e.partnerName] = (raw[e.partnerName] || 0) + delta;
    });
    return raw;
  }, [partnerCashbook]);
  const totalPartnerEquity = Object.values(partnerEquity).reduce((s, a) => s + a, 0);

  // Plant operations are no longer tracked here, so revenue is what actually hit
  // the bank from counterparties tagged "Customer / Buyer" on the Bank Statement.
  const bankGroups = useCounterpartyGroups(bankTxns, bankLabels);
  const totalRevenue = bankGroups.filter(g => g.type === "customer").reduce((s, g) => s + g.totalCredit, 0);
  const totalSpend = (purchases || []).reduce((s, p) => s + (p.taxableAmount || 0), 0) + (expenses || []).reduce((s, e) => s + (e.amount || 0), 0);
  const retainedEarnings = totalRevenue - totalSpend;
  const totalFinanced = totalLiabilities + totalPartnerEquity + retainedEarnings;

  return (
    <div>
      <div className="stats-grid">
        <StatCard label="Gross Fixed Assets" value={fmt(totalFixedAssets)} color="teal" sub="Land, Civil, Plant" />
        <StatCard label="Current Assets" value={fmt(totalCurrentAssets)} color="blue" sub="Bank balance" />
        <StatCard label="Total Liabilities" value={fmt(totalLiabilities)} color="red" sub="Loans & Overdraft" />
        <StatCard label="Partner Net Worth" value={fmt(totalPartnerEquity + retainedEarnings)} color="accent" sub="Capital + Retained P&L" />
      </div>

      <div className="two-col">
        <div className="report-section">
          <h3>Assets (Gross Block & Cash)</h3>
          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"10px 0" }}>FIXED INFRASTRUCTURE ASSETS (CAPEX)</div>
          {Object.entries(fixedAssets).map(([cat, amt]) => (
            <div key={cat} className="cost-row">
              <span className="cost-label" style={{ textTransform:"capitalize" }}>{cat}</span>
              <span className="cost-val">{fmt(amt)}</span>
            </div>
          ))}

          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"16px 0 8px" }}>CURRENT ASSETS</div>
          <div className="cost-row">
            <span className="cost-label">Operating Bank Balance</span>
            <span className="cost-val" style={{ color: bankBalance >= 0 ? "var(--green)" : "var(--amber)" }}>{bankBalance >= 0 ? fmt(bankBalance) : "₹0.00"}</span>
          </div>
          <div className="cost-row grand">
            <span className="cost-label">Total Assets Base</span>
            <span className="cost-val accent">{fmt(totalAssets)}</span>
          </div>
        </div>

        <div className="report-section">
          <h3>Liabilities & Partner Net Worth</h3>
          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"10px 0" }}>DEBT & CURRENT LIABILITIES</div>
          <div className="cost-row">
            <span className="cost-label">Machinery & Project Debt</span>
            <span className="cost-val" style={{ color:"var(--red)" }}>{fmt(loanCapital)}</span>
          </div>
          {bankOverdraft > 0 && (
            <div className="cost-row">
              <span className="cost-label">Bank Overdraft Facility</span>
              <span className="cost-val" style={{ color:"var(--red)" }}>{fmt(bankOverdraft)}</span>
            </div>
          )}

          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"16px 0 8px" }}>PARTNER EQUITY & REVENUE RESERVES</div>
          {Object.entries(partnerEquity).map(([partner, amt]) => (
            <div key={partner} className="cost-row">
              <span className="cost-label">{partner} (Net Standing)</span>
              <span className="cost-val" style={{ color:"var(--accent)" }}>{fmt(amt)}</span>
            </div>
          ))}
          <div className="cost-row">
            <span className="cost-label">Receipts from Customers (bank)</span>
            <span className="cost-val" style={{ color:"var(--green)" }}>{fmt(totalRevenue)}</span>
          </div>
          <div className="cost-row">
            <span className="cost-label">Accumulated Retained Earnings</span>
            <span className="cost-val" style={{ color: retainedEarnings >= 0 ? "var(--green)" : "var(--red)" }}>{fmt(retainedEarnings)}</span>
          </div>
          <div className="cost-row grand">
            <span className="cost-label">Total Liabilities & Equity</span>
            <span className="cost-val accent">{fmt(totalFinanced)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── MASTER TABLE HELPER WITH EDIT & DELETE SUPPORT ───────────────────────────
function MasterTableView({ title, noun, items, setItems, saveFn, delFn, fields, cols }) {
  const blank = Object.fromEntries(fields.map(f => [f.key, ""]));
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(blank);
  const [editingId, setEditingId] = useState(null);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  function openAdd() {
    setEditingId(null);
    setF(blank);
    setOpen(true);
  }

  function openEdit(item) {
    setEditingId(item.id);
    setF(item);
    setOpen(true);
  }

  async function save() {
    try {
      const payload = editingId ? { ...f, id: editingId } : f;
      const res = await saveFn(payload);
      if (editingId) {
        setItems(xs => xs.map(x => x.id === editingId ? res : x));
      } else {
        setItems(xs => [...xs, res]);
      }
      setOpen(false); setF(blank); setEditingId(null);
    } catch (err) { alert(err.message); }
  }

  async function remove(id) {
    if (!confirm(`Are you sure you want to delete this ${noun}?`)) return;
    try {
      if (delFn) await delFn(id);
      setItems(xs => xs.filter(x => x.id !== id));
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={openAdd}>+ Add {noun}</button></div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              {cols.map(c => <th key={c.label}>{c.label}</th>)}
              <th className="r">Actions</th>
            </tr>
          </thead>
          <tbody>
            {(items || []).length === 0 ? (
              <tr><td colSpan={cols.length + 1} style={{ textAlign: "center", color: "var(--text3)", padding: 24 }}>No {noun.toLowerCase()} records found</td></tr>
            ) : (
              (items || []).map(row => (
                <tr key={row.id}>
                  {cols.map(c => <td key={c.label}>{row[c.key] || "—"}</td>)}
                  <td className="r">
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button className="btn btn-ghost btn-sm" onClick={() => openEdit(row)}>Edit</button>
                      {delFn && <button className="btn btn-danger btn-sm" onClick={() => remove(row.id)}>Delete</button>}
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title={`${editingId ? "Edit" : "New"} ${noun}`} onClose={() => { setOpen(false); setEditingId(null); }} foot={<><button className="btn btn-ghost" onClick={() => { setOpen(false); setEditingId(null); }}>Cancel</button><button className="btn btn-primary" onClick={save}>Save {noun}</button></>}>
          {fields.map(fl => (
            <FG key={fl.key} label={fl.label} span={fl.full ? 2 : undefined}>
              <input value={f[fl.key] !== undefined ? f[fl.key] : ""} onChange={e => set(fl.key)(e.target.value)} />
            </FG>
          ))}
        </Modal>
      )}
    </div>
  );
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function DashboardView({ expenses, purchases, capitalItems, bankTxns, partnerCashbook, vendorProjects, projectPayments, goTo }) {
  const totalCapex = (capitalItems || []).reduce((s, c) => s + (c.amount || 0), 0);
  const bankBalance = (bankTxns || []).reduce((s, t) => s + ((t.credit || 0) - (t.debit || 0)), 0);
  const totalPartnerCapital = (partnerCashbook || []).filter(e => e.type === "Capital Infusion").reduce((s, e) => s + (e.amount || 0), 0);
  const totalSpend = (purchases || []).reduce((s, p) => s + (p.totalAmount || 0), 0) + (expenses || []).reduce((s, e) => s + (e.amount || 0), 0);

  const projects = useMemo(() => summarizeProjects(vendorProjects, projectPayments), [vendorProjects, projectPayments]);
  const open = projects.filter(p => p.status !== "completed" && p.status !== "cancelled");
  const dueTotal = projects.reduce((s, p) => s + p.due, 0);
  const owing = projects.filter(p => p.due > EPS).sort((a, b) => b.due - a.due).slice(0, 8);

  return (
    <div>
      <div className="stats-grid-5">
        <StatCard label="Bank Balance" value={fmt(bankBalance)} color={bankBalance >= 0 ? "green" : "red"} sub="Liquid Cash" />
        <StatCard label="Fixed Assets Gross" value={fmt(totalCapex)} color="teal" sub="Land & Plant" />
        <StatCard label="Partner Capital" value={fmt(totalPartnerCapital)} color="accent" sub="Contributed" />
        <StatCard label="Spend to Date" value={fmt(totalSpend)} color="amber" sub="Purchases + expenses" />
        <StatCard label="Vendor Dues" value={fmt(dueTotal)} color="red" sub={`${open.length} open project(s)`} />
      </div>

      <div className="table-wrap">
        <div className="table-toolbar">
          <h3>Vendor projects with a balance due</h3>
          <button className="btn btn-ghost btn-sm" onClick={() => goTo("vendorprojects")}>Open Vendor Projects →</button>
        </div>
        <table>
          <thead><tr><th>Project</th><th>Vendor</th><th className="r">Contract</th><th className="r">Paid</th><th className="r">Due</th></tr></thead>
          <tbody>
            {owing.length === 0 ? (
              <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--text3)", padding: 24 }}>No outstanding vendor project balances</td></tr>
            ) : owing.map(p => (
              <tr key={p.id}>
                <td><span className="mono" style={{ color: "var(--accent)", marginRight: 8 }}>{p.projectNo}</span>{p.title}</td>
                <td>{p.vendorName}</td>
                <td className="r mono">{fmt(p.contractValue)}</td>
                <td className="r mono" style={{ color: "var(--green)" }}>{fmt(p.paid)}</td>
                <td className="r mono" style={{ color: "var(--amber)", fontWeight: 700 }}>{fmt(p.due)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── VENDOR PROJECTS ───────────────────────────────────────────────────────────
// A "project" is a job / work / thing a vendor is doing for you. Payments that
// already exist elsewhere (bank debits, expenses, partner-paid capex) are LINKED
// to a project — a payment can be split across projects by linking part of its
// amount to each. Linking never creates or changes money in the ledgers; it only
// tags where a payment went. Cash that isn't recorded anywhere else can be added
// as a manual, project-only payment.
const PROJECT_STATUSES = [
  { id: "planned", label: "Planned", badge: "blue" },
  { id: "active", label: "In Progress", badge: "accent" },
  { id: "on_hold", label: "On Hold", badge: "amber" },
  { id: "completed", label: "Completed", badge: "green" },
  { id: "cancelled", label: "Cancelled", badge: "muted" },
];
const PROJECT_CATS = [
  ...EXP_CATS,
  ...CAP_CATS.map(c => ({ id: c.id, label: `Capital — ${c.label}` })),
];
const catLabel = id => (PROJECT_CATS.find(c => c.id === id) || {}).label || id;
const normVendor = s => (s || "").trim().toLowerCase();
const EPS = 0.005;

// Projects with their payments, paid and due worked out.
function summarizeProjects(projects, payments) {
  const byProject = {};
  (payments || []).forEach(p => { (byProject[p.projectId] = byProject[p.projectId] || []).push(p); });
  return (projects || []).map(p => {
    const pays = (byProject[p.id] || []).slice().sort((a, b) => (b.date || "").localeCompare(a.date || ""));
    const paid = pays.reduce((s, x) => s + x.amount, 0);
    const hasValue = p.contractValue > 0;
    return {
      ...p, payments: pays, paid, hasValue,
      due: hasValue ? Math.max(0, p.contractValue - paid) : 0,
      overpaid: hasValue ? Math.max(0, paid - p.contractValue) : 0,
      pct: hasValue ? Math.min(100, (paid / p.contractValue) * 100) : 0,
    };
  });
}

// Every vendor payment that exists in the ledgers, plus how much of each is
// already linked to projects.
function usePaymentPool(bankTxns, bankLabels, expenses, capitalItems, projectPayments) {
  const bankGroups = useCounterpartyGroups(bankTxns, bankLabels);

  const pool = useMemo(() => {
    const list = [];
    bankGroups.forEach(g => {
      if (g.type === "owner") return; // partner drawings aren't vendor payments
      g.txns.forEach(t => {
        if (!(t.debit > 0)) return;
        list.push({
          key: `bank-${t.id}`, sourceType: "bank", sourceId: t.id,
          date: t.date || t.txn_date || "", vendor: g.label, description: t.description || "",
          amount: +t.debit, sourceLabel: "Bank",
        });
      });
    });
    (expenses || []).forEach(e => {
      // Monthly overhead totals aren't a payment to one vendor; promoted bank
      // debits are already represented by their bank row above.
      if (!e || e.opsAuto || e.bankTxnId) return;
      list.push({
        key: `expense-${e.id}`, sourceType: "expense", sourceId: e.id,
        date: e.date || "", vendor: (e.paidTo || "").trim(), description: e.description || "",
        amount: +e.amount || 0, sourceLabel: e.paidByPartner ? `Cash · ${e.paidByPartner}` : "Expense",
      });
    });
    (capitalItems || []).forEach(c => {
      if (!c || c.bankTxnId || !c.paidByPartner) return; // same rule Capital & Infra uses
      list.push({
        key: `capital-${c.id}`, sourceType: "capital", sourceId: c.id,
        date: c.date || "", vendor: (c.paidTo || "").trim(), description: c.description || "",
        amount: +c.amount || 0, sourceLabel: `Cash · ${c.paidByPartner} · Capex`,
      });
    });
    return list;
  }, [bankGroups, expenses, capitalItems]);

  const allocated = useMemo(() => {
    const m = {};
    (projectPayments || []).forEach(p => {
      let k = null;
      if (p.bankTxnId != null) k = `bank-${p.bankTxnId}`;
      else if (p.expenseId != null) k = `expense-${p.expenseId}`;
      else if (p.capitalItemId != null) k = `capital-${p.capitalItemId}`;
      if (k) m[k] = (m[k] || 0) + p.amount;
    });
    return m;
  }, [projectPayments]);

  const items = useMemo(
    () => pool.map(x => ({ ...x, allocated: allocated[x.key] || 0, remaining: Math.max(0, x.amount - (allocated[x.key] || 0)) })),
    [pool, allocated]
  );
  const byKey = useMemo(() => Object.fromEntries(items.map(x => [x.key, x])), [items]);
  return { items, byKey };
}

function paymentSourceKey(p) {
  if (p.bankTxnId != null) return `bank-${p.bankTxnId}`;
  if (p.expenseId != null) return `expense-${p.expenseId}`;
  if (p.capitalItemId != null) return `capital-${p.capitalItemId}`;
  return null;
}

function VendorProjectsView({ bankTxns, bankLabels, expenses, capitalItems, suppliers, projects, setProjects, projectPayments, setProjectPayments, partnerCashbook, setPartnerCashbook, setExpenses, setCapitalItems, status }) {
  const [tab, setTab] = useState("projects");
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("open"); // open | all | <status id>
  const [expanded, setExpanded] = useState(null);
  const [formOpen, setFormOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [linkFor, setLinkFor] = useState(null);
  const [manualFor, setManualFor] = useState(null);
  const [stmtFor, setStmtFor] = useState(null);
  const [assignItem, setAssignItem] = useState(null);
  const [unSearch, setUnSearch] = useState("");
  const [unShowAll, setUnShowAll] = useState(false);
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [summaryWithTx, setSummaryWithTx] = useState(false);
  const [cashFor, setCashFor] = useState(null);       // project to record a partner cash payment against
  const [vendorEditItem, setVendorEditItem] = useState(null); // cash payment whose vendor is being corrected

  const blank = { vendorName: "", title: "", category: "misc", status: "active", contractValue: "", startDate: "", endDate: "", description: "", notes: "" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  const ready = status?.ready !== false;
  const { items: poolItems, byKey: poolByKey } = usePaymentPool(bankTxns, bankLabels, expenses, capitalItems, projectPayments);
  const summarized = useMemo(() => summarizeProjects(projects, projectPayments), [projects, projectPayments]);

  const vendorNames = useMemo(() => {
    const m = {};
    const add = n => { const t = (n || "").trim(); if (t && !m[normVendor(t)]) m[normVendor(t)] = t; };
    (projects || []).forEach(p => add(p.vendorName));
    (suppliers || []).forEach(s => add(s.name));
    poolItems.forEach(x => add(x.vendor));
    return Object.values(m).sort((a, b) => a.localeCompare(b));
  }, [projects, suppliers, poolItems]);

  const projectVendorKeys = useMemo(() => new Set((projects || []).map(p => normVendor(p.vendorName))), [projects]);

  function openAdd() { setEditingId(null); setF(blank); setFormOpen(true); }
  function openEdit(p) {
    setEditingId(p.id);
    setF({
      vendorName: p.vendorName, title: p.title, category: p.category, status: p.status,
      contractValue: p.contractValue ? String(p.contractValue) : "", startDate: p.startDate || "", endDate: p.endDate || "",
      description: p.description || "", notes: p.notes || "",
    });
    setFormOpen(true);
  }
  async function saveProject() {
    if (!f.vendorName.trim() || !f.title.trim()) return alert("Vendor and project name are required.");
    try {
      const row = await vpdb.saveVendorProject(f, editingId);
      if (editingId) setProjects(ps => ps.map(p => p.id === editingId ? row : p));
      else setProjects(ps => [row, ...ps]);
      setFormOpen(false); setEditingId(null); setF(blank);
    } catch (err) { alert(err.message); }
  }
  async function removeProject(p) {
    if (!confirm(`Delete "${p.title}" (${p.projectNo})?\n\nIts ${p.payments.length} payment link(s) are removed. The underlying bank transactions / expenses stay exactly as they are.`)) return;
    try {
      await vpdb.deleteVendorProject(p.id);
      setProjects(ps => ps.filter(x => x.id !== p.id));
      setProjectPayments(pp => pp.filter(x => x.projectId !== p.id));
    } catch (err) { alert(err.message); }
  }

  async function linkPayment(project, item, amount) {
    const amt = +amount;
    if (!(amt > 0)) return alert("Enter an amount greater than zero.");
    if (amt > item.remaining + EPS) return alert(`Only ${fmt(item.remaining)} of this payment is still unassigned.`);
    try {
      const row = await vpdb.addProjectPayment({
        projectId: project.id,
        bankTxnId: item.sourceType === "bank" ? item.sourceId : null,
        expenseId: item.sourceType === "expense" ? item.sourceId : null,
        capitalItemId: item.sourceType === "capital" ? item.sourceId : null,
        amount: amt, date: item.date, paymentMode: item.sourceType === "bank" ? "bank" : "cash",
        remarks: item.description,
      });
      setProjectPayments(pp => [row, ...pp]);
      return true;
    } catch (err) { alert(err.message); return false; }
  }
  async function addManualPayment(project, data) {
    if (!(+data.amount > 0)) return alert("Enter an amount greater than zero.");
    try {
      const row = await vpdb.addProjectPayment({ projectId: project.id, ...data });
      setProjectPayments(pp => [row, ...pp]);
      setManualFor(null);
    } catch (err) { alert(err.message); }
  }
  async function unlinkPayment(p) {
    if (!confirm("Remove this payment from the project? The payment itself is not deleted.")) return;
    try {
      await vpdb.deleteProjectPayment(p.id);
      setProjectPayments(pp => pp.filter(x => x.id !== p.id));
    } catch (err) { alert(err.message); }
  }

  // Partners = anyone tagged as an owner on the bank statement + anyone with a cashbook entry.
  const bankGroupsForPartners = useCounterpartyGroups(bankTxns, bankLabels);
  const partners = useMemo(() => {
    const m = {};
    bankGroupsForPartners.forEach(g => { if (g.type === "owner") m[normVendor(g.label)] = g.label; });
    (partnerCashbook || []).forEach(e => { const n = (e?.partnerName || "").trim(); if (n && !m[normVendor(n)]) m[normVendor(n)] = n; });
    return Object.values(m).sort((a, b) => a.localeCompare(b));
  }, [bankGroupsForPartners, partnerCashbook]);

  // Partner paid a vendor in cash: books the expense (or capex), credits the partner's
  // current account, and links the whole amount to this project in one go.
  async function savePartnerCash(project, d) {
    if (!d.partnerName) return alert("Pick the partner who paid.");
    if (!(+d.amount > 0)) return alert("Enter an amount greater than zero.");
    if (!d.description.trim()) return alert("Add a short description.");
    const isCapex = CAP_CATS.some(c => c.id === project.category);
    try {
      const res = await db.recordPartnerDirectCashExpense(
        d.partnerName, d.date, project.category, d.amount, d.description.trim(), d.ref, isCapex, project.vendorName
      );
      setPartnerCashbook(prev => [res.cashbookEntry, ...(prev || [])]);
      if (isCapex) setCapitalItems(prev => [res.item, ...(prev || [])]);
      else setExpenses(prev => [res.item, ...(prev || [])]);
      const row = await vpdb.addProjectPayment({
        projectId: project.id,
        expenseId: isCapex ? null : res.item.id,
        capitalItemId: isCapex ? res.item.id : null,
        amount: +d.amount, date: d.date, paymentMode: "cash", remarks: d.description.trim(),
      });
      setProjectPayments(pp => [row, ...pp]);
      setCashFor(null);
    } catch (err) { alert(err.message); }
  }

  // Fix the vendor on an existing partner-cash entry so it lines up with the project's vendor.
  async function saveItemVendor(item, vendorName) {
    const name = (vendorName || "").trim();
    if (!name) return alert("Enter the vendor name.");
    try {
      if (item.sourceType === "expense") {
        await db.setPaidTo("expenses", item.sourceId, name);
        setExpenses(es => es.map(e => e.id === item.sourceId ? { ...e, paidTo: name } : e));
      } else if (item.sourceType === "capital") {
        await db.setPaidTo("capital_items", item.sourceId, name);
        setCapitalItems(cs => cs.map(c => c.id === item.sourceId ? { ...c, paidTo: name } : c));
      }
      setVendorEditItem(null);
    } catch (err) { alert(err.message); }
  }

  const filtered = summarized.filter(p => {
    if (statusFilter === "open" && (p.status === "completed" || p.status === "cancelled")) return false;
    if (statusFilter !== "open" && statusFilter !== "all" && p.status !== statusFilter) return false;
    if (!search) return true;
    const s = search.toLowerCase();
    return p.vendorName.toLowerCase().includes(s) || p.title.toLowerCase().includes(s) || (p.projectNo || "").toLowerCase().includes(s) || (p.description || "").toLowerCase().includes(s);
  });

  // Group projects under their vendor, with the vendor's overall picture.
  const vendorGroups = useMemo(() => {
    const m = {};
    filtered.forEach(p => {
      const k = normVendor(p.vendorName);
      if (!m[k]) m[k] = { key: k, vendor: p.vendorName, projects: [], contracted: 0, paid: 0, due: 0 };
      m[k].projects.push(p);
      m[k].contracted += p.contractValue; m[k].paid += p.paid; m[k].due += p.due;
    });
    return Object.values(m).map(g => {
      const mine = poolItems.filter(x => normVendor(x.vendor) === g.key);
      return {
        ...g,
        totalPaidToVendor: mine.reduce((s, x) => s + x.amount, 0),
        unassigned: mine.reduce((s, x) => s + x.remaining, 0),
        unassignedCount: mine.filter(x => x.remaining > EPS).length,
      };
    }).sort((a, b) => a.vendor.localeCompare(b.vendor));
  }, [filtered, poolItems]);

  const unassignedRows = useMemo(() => {
    const s = unSearch.toLowerCase();
    return poolItems
      .filter(x => x.remaining > EPS)
      .filter(x => unShowAll || projectVendorKeys.has(normVendor(x.vendor)))
      .filter(x => !s || x.vendor.toLowerCase().includes(s) || x.description.toLowerCase().includes(s))
      .sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  }, [poolItems, unShowAll, projectVendorKeys, unSearch]);
  const unassignedTotal = unassignedRows.reduce((s, x) => s + x.remaining, 0);

  const openProjects = summarized.filter(p => p.status !== "completed" && p.status !== "cancelled");
  const totalContract = summarized.reduce((s, p) => s + p.contractValue, 0);
  const totalPaid = summarized.reduce((s, p) => s + p.paid, 0);
  const totalDue = summarized.reduce((s, p) => s + p.due, 0);

  const statusMeta = id => PROJECT_STATUSES.find(s => s.id === id) || { label: id, badge: "muted" };
  const paymentRowInfo = p => {
    const src = poolByKey[paymentSourceKey(p)];
    if (paymentSourceKey(p) == null) return { kind: "Manual", badge: "purple", text: p.reference || p.remarks || "—" };
    return { kind: src ? src.sourceLabel : "Linked", badge: src?.sourceType === "bank" ? "blue" : "accent", text: src?.description || p.remarks || "—" };
  };

  function exportVendorSummary(withTransactions) {
    const vendors = vendorGroups.map(g => ({
      vendor: g.vendor, contracted: g.contracted, paid: g.paid, due: g.due,
      projects: g.projects.map(p => ({
        projectNo: p.projectNo, title: p.title, hasValue: p.hasValue, contractValue: p.contractValue, paid: p.paid, due: p.due,
        payments: p.payments.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(x => {
          const info = paymentRowInfo(x);
          return { date: x.date, source: info.kind, particulars: info.text, amount: x.amount };
        }),
      })),
    }));
    const filterLabel = statusFilter === "open" ? "Open projects" : statusFilter === "all" ? "All projects" : `${statusMeta(statusFilter).label} projects`;
    exportVendorSummaryPDF({
      vendors, withTransactions,
      subtitle: `${filterLabel}${search ? ` matching "${search}"` : ""} · ${vendors.length} vendor(s)`,
      fileName: `vendor_summary${withTransactions ? "_with_transactions" : ""}.pdf`,
    });
    setSummaryOpen(false);
  }

  function exportStatement(p) {
    exportStatementPDF({
      title: `Vendor Project Statement: ${p.projectNo}`,
      subtitle: `${p.vendorName} — ${p.title}${p.description ? ` (${p.description})` : ""}`,
      summary: [
        { label: p.hasValue ? "Contract Value" : "Contract Value", value: p.hasValue ? fmt(p.contractValue) : "Open-ended", color: [90, 160, 220] },
        { label: "Paid", value: fmt(p.paid), color: [70, 180, 100] },
        ...(p.hasValue ? [{ label: "Balance Due", value: fmt(p.due), color: [220, 150, 30] }] : []),
      ],
      columns: [
        { key: "date", label: "Date" }, { key: "source", label: "Source" },
        { key: "particulars", label: "Particulars" }, { key: "amount", label: "Amount", align: "right" },
      ],
      rows: p.payments.slice().sort((a, b) => (a.date || "").localeCompare(b.date || "")).map(x => {
        const info = paymentRowInfo(x);
        return { date: x.date, source: info.kind, particulars: info.text, amount: fmt(x.amount) };
      }),
      fileName: `${p.projectNo}_${p.vendorName.replace(/[^\w]+/g, "_")}_statement.pdf`,
    });
  }

  return (
    <div>
      {!ready && (
        <div className="config-card" style={{ borderColor: "#4a3500", background: "#201500" }}>
          <h3 style={{ color: "var(--amber)" }}>Vendor Projects database setup needed</h3>
          <p style={{ fontSize: 12.5, color: "var(--text2)", lineHeight: 1.6 }}>
            The <span className="mono">vendor_projects</span> tables don't exist yet. Open the Supabase SQL editor, run the
            contents of <span className="mono">supabase/vendor_projects.sql</span>, then reload this page.
          </p>
          {status?.error && <p className="mono" style={{ fontSize: 11, color: "var(--text3)", marginTop: 8 }}>{status.error}</p>}
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div className="pill-tabs" style={{ marginBottom: 0 }}>
          <button className={`pill-tab ${tab === "projects" ? "active" : ""}`} onClick={() => setTab("projects")}>Projects ({openProjects.length} open)</button>
          <button className={`pill-tab ${tab === "unassigned" ? "active" : ""}`} onClick={() => setTab("unassigned")}>Unassigned Payments</button>
        </div>
        <button className="btn btn-primary" disabled={!ready} onClick={openAdd}>+ New Vendor Project</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Open Projects" value={openProjects.length} color="purple" sub={`of ${summarized.length} total`} />
        <StatCard label="Contract Value" value={fmt(totalContract)} color="blue" sub="Quoted / agreed" />
        <StatCard label="Paid Against Projects" value={fmt(totalPaid)} color="green" sub="Linked payments" />
        <StatCard label="Balance Due" value={fmt(totalDue)} color="amber" sub="On projects with a contract value" />
      </div>

      {tab === "projects" ? (
        <div>
          <div className="filter-bar">
            <SearchBar value={search} onChange={setSearch} placeholder="Search vendor, project or PO number…" />
            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} style={{ background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: "var(--r)", padding: "7px 11px", color: "var(--text)", fontSize: 12.5 }}>
              <option value="open">Open projects</option>
              <option value="all">All projects</option>
              {PROJECT_STATUSES.map(s => <option key={s.id} value={s.id}>{s.label} only</option>)}
            </select>
            <button className="btn btn-ghost btn-sm" onClick={() => exportRowsToExcel("vendor_projects", filtered.flatMap(p => (p.payments.length ? p.payments : [null]).map(x => ({
              Project: p.projectNo, Vendor: p.vendorName, Title: p.title, Category: catLabel(p.category), Status: statusMeta(p.status).label,
              ContractValue: p.contractValue, ProjectPaid: p.paid, ProjectDue: p.due,
              PaymentDate: x?.date || "", PaymentSource: x ? paymentRowInfo(x).kind : "", PaymentDetail: x ? paymentRowInfo(x).text : "", PaymentAmount: x ? x.amount : "",
            }))))}>⬇ Export Excel</button>
            <button className="btn btn-ghost btn-sm" disabled={vendorGroups.length === 0} onClick={() => setSummaryOpen(true)}>📄 Vendor Summary PDF</button>
          </div>

          {vendorGroups.length === 0 ? (
            <div className="config-card" style={{ textAlign: "center", padding: 32 }}>
              <EmptyState icon="🛠" message="No vendor projects yet" sub="Create one for each job a vendor is doing, then link that vendor's payments to it." />
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
              {vendorGroups.map(g => (
                <div key={g.key}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 8, marginBottom: 8, padding: "0 2px" }}>
                    <div>
                      <div style={{ fontSize: 10, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>Vendor</div>
                      <div style={{ fontSize: 16, fontWeight: 700 }}>{g.vendor}</div>
                    </div>
                    <div style={{ display: "flex", gap: 16, fontFamily: "var(--mono)", fontSize: 11.5, flexWrap: "wrap" }}>
                      <div><span style={{ color: "var(--text3)" }}>PAID TO VENDOR </span><span style={{ fontWeight: 700 }}>{fmt(g.totalPaidToVendor)}</span></div>
                      <div><span style={{ color: "var(--text3)" }}>ON PROJECTS </span><span style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(g.paid)}</span></div>
                      <div><span style={{ color: "var(--text3)" }}>NOT ASSIGNED </span><span style={{ color: g.unassigned > EPS ? "var(--amber)" : "var(--text3)", fontWeight: 700 }}>{fmt(g.unassigned)}</span></div>
                    </div>
                  </div>

                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    {g.projects.map(p => {
                      const sm = statusMeta(p.status);
                      return (
                        <div key={p.id} className="config-card" style={{ padding: 14, marginBottom: 0 }}>
                          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end", justifyContent: "space-between" }}>
                            <div style={{ minWidth: 220, flex: 1 }}>
                              <div style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)", textTransform: "uppercase" }}>{p.projectNo} · {catLabel(p.category)}</div>
                              <div style={{ fontSize: 15, fontWeight: 700, marginTop: 2, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>{p.title} <BadgeComponent type={sm.badge}>{sm.label}</BadgeComponent></div>
                              {p.description && <div style={{ fontSize: 12, color: "var(--text2)", marginTop: 2 }}>{p.description}</div>}
                              {(p.startDate || p.endDate) && <div style={{ fontSize: 11, color: "var(--text3)", fontFamily: "var(--mono)", marginTop: 3 }}>{p.startDate || "…"} → {p.endDate || "…"}</div>}
                            </div>
                            <div style={{ display: "flex", gap: 18, fontFamily: "var(--mono)", fontSize: 12.5, alignItems: "center" }}>
                              <div><div style={{ color: "var(--text3)", fontSize: 10 }}>CONTRACT</div><div style={{ fontWeight: 700 }}>{p.hasValue ? fmt(p.contractValue) : "Open-ended"}</div></div>
                              <div><div style={{ color: "var(--text3)", fontSize: 10 }}>PAID</div><div style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(p.paid)}</div></div>
                              {p.hasValue && (p.overpaid > EPS
                                ? <span className="badge badge-red">OVERPAID {fmt(p.overpaid)}</span>
                                : p.due > EPS ? <span className="due-badge">DUE {fmt(p.due)}</span> : <BadgeComponent type="green">Fully Paid</BadgeComponent>)}
                            </div>
                          </div>

                          {p.hasValue && (
                            <div style={{ height: 5, background: "var(--bg4)", borderRadius: 4, marginTop: 12, overflow: "hidden" }}>
                              <div style={{ width: `${p.pct}%`, height: "100%", background: p.overpaid > EPS ? "var(--red)" : "var(--green)" }} />
                            </div>
                          )}

                          <div style={{ marginTop: 10, display: "flex", justifyContent: "flex-end", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                            <button className="btn btn-primary btn-sm" disabled={!ready} onClick={() => setLinkFor(p)}>🔗 Link Payment</button>
                            <button className="btn btn-ghost btn-sm" disabled={!ready} onClick={() => setCashFor(p)}>💵 Partner Cash</button>
                            <button className="btn btn-ghost btn-sm" disabled={!ready} onClick={() => setManualFor(p)}>+ Manual Payment</button>
                            <button className="btn btn-info btn-sm" onClick={() => setStmtFor(p)}>📄 Statement</button>
                            <button className="btn btn-ghost btn-sm" onClick={() => openEdit(p)}>Edit</button>
                            <button className="btn btn-danger btn-sm" onClick={() => removeProject(p)}>Delete</button>
                            <button className="btn btn-ghost btn-sm" onClick={() => setExpanded(x => x === p.id ? null : p.id)}>{expanded === p.id ? "Hide" : "Show"} payments ({p.payments.length})</button>
                          </div>

                          {expanded === p.id && (
                            <div className="table-wrap" style={{ marginTop: 10, marginBottom: 0 }}>
                              <table>
                                <thead><tr><th>Date</th><th>Source</th><th>Narration / Details</th><th className="r">Amount</th><th className="r">Actions</th></tr></thead>
                                <tbody>
                                  {p.payments.length === 0 ? (
                                    <tr><td colSpan={5} style={{ textAlign: "center", color: "var(--text3)", padding: 16 }}>No payments linked yet</td></tr>
                                  ) : p.payments.map(x => {
                                    const info = paymentRowInfo(x);
                                    return (
                                      <tr key={x.id}>
                                        <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{x.date}</td>
                                        <td><BadgeComponent type={info.badge}>{info.kind}</BadgeComponent></td>
                                        <td style={{ fontSize: 12 }}>{info.text}</td>
                                        <td className="r mono" style={{ fontWeight: 700, color: "var(--green)" }}>{fmt(x.amount)}</td>
                                        <td className="r"><RowActions onDelete={() => unlinkPayment(x)} /></td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                  {g.unassignedCount > 0 && (
                    <div style={{ fontSize: 11.5, color: "var(--amber)", marginTop: 8, fontFamily: "var(--mono)" }}>
                      {g.unassignedCount} payment(s) to {g.vendor} not fully assigned to a project — see the Unassigned Payments tab.
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      ) : (
        <div>
          <p style={{ fontSize: 12, color: "var(--text3)", marginBottom: 12 }}>
            Bank debits, expenses and partner-paid cash that haven't been fully assigned to a project. Assign all of a payment, or just part of it to split it across projects.
          </p>
          <div className="filter-bar">
            <SearchBar value={unSearch} onChange={setUnSearch} placeholder="Search vendor or narration…" />
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "var(--text2)" }}>
              <input type="checkbox" checked={unShowAll} onChange={e => setUnShowAll(e.target.checked)} />
              Include vendors with no project
            </label>
            <span className="mono" style={{ fontSize: 11.5, color: "var(--text3)" }}>{unassignedRows.length} payment(s) · {fmt(unassignedTotal)} unassigned</span>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Date</th><th>Vendor</th><th>Narration</th><th>Source</th><th className="r">Amount</th><th className="r">Unassigned</th><th className="r">Action</th></tr></thead>
              <tbody>
                {unassignedRows.length === 0 ? (
                  <tr><td colSpan={7} style={{ textAlign: "center", color: "var(--text3)", padding: 28 }}>{projectVendorKeys.size === 0 ? "Create a vendor project first — payments to that vendor will show up here." : "Everything is assigned. 🎉"}</td></tr>
                ) : unassignedRows.slice(0, 300).map(x => (
                  <tr key={x.key}>
                    <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{x.date}</td>
                    <td style={{ fontWeight: 600 }}>{x.vendor || "—"}</td>
                    <td style={{ fontSize: 12 }}>{x.description}</td>
                    <td><BadgeComponent type={x.sourceType === "bank" ? "blue" : "accent"}>{x.sourceLabel}</BadgeComponent></td>
                    <td className="r mono">{fmt(x.amount)}</td>
                    <td className="r mono" style={{ fontWeight: 700, color: "var(--amber)" }}>{fmt(x.remaining)}</td>
                    <td className="r" style={{ whiteSpace: "nowrap" }}>
                      {x.sourceType !== "bank" && <button className="btn btn-ghost btn-sm" style={{ marginRight: 6 }} onClick={() => setVendorEditItem(x)}>Set vendor</button>}
                      <button className="btn btn-primary btn-sm" disabled={!ready || !(projects || []).length} onClick={() => setAssignItem(x)}>Assign</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {unassignedRows.length > 300 && <p style={{ fontSize: 11, color: "var(--text3)" }}>Showing the latest 300 — use search to narrow.</p>}
        </div>
      )}

      {formOpen && (
        <Modal title={editingId ? "Edit Vendor Project" : "New Vendor Project"} onClose={() => { setFormOpen(false); setEditingId(null); }}
          foot={<><button className="btn btn-ghost" onClick={() => { setFormOpen(false); setEditingId(null); }}>Cancel</button><button className="btn btn-primary" onClick={saveProject}>{editingId ? "Save Changes" : "Create Project"}</button></>}>
          <div className="form-row">
            <FG label="Vendor *" note="Pick an existing vendor or type a new one. Use the same name as on the Bank Statement label so their payments line up.">
              <input list="vp-vendors" value={f.vendorName} onChange={e => set("vendorName")(e.target.value)} placeholder="e.g. Mateshwari Industries" />
              <datalist id="vp-vendors">{vendorNames.map(n => <option key={n} value={n} />)}</datalist>
            </FG>
          </div>
          <div className="form-row"><FG label="Project / Work *"><input value={f.title} onChange={e => set("title")(e.target.value)} placeholder="e.g. Boundary wall, Transformer installation" /></FG></div>
          <div className="form-row cols-2">
            <FG label="Category">
              <select value={f.category} onChange={e => set("category")(e.target.value)}>
                <optgroup label="Operating">{EXP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</optgroup>
                <optgroup label="Capital">{CAP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</optgroup>
              </select>
            </FG>
            <FG label="Status">
              <select value={f.status} onChange={e => set("status")(e.target.value)}>{PROJECT_STATUSES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}</select>
            </FG>
          </div>
          <div className="form-row"><FG label="Contract / Quoted Value (₹)" note="Optional. Leave blank for ongoing work with no fixed quote — payments are still tracked."><input type="number" value={f.contractValue} onChange={e => set("contractValue")(e.target.value)} /></FG></div>
          <div className="form-row cols-2">
            <FG label="Start Date"><input type="date" value={f.startDate} onChange={e => set("startDate")(e.target.value)} /></FG>
            <FG label="End Date"><input type="date" value={f.endDate} onChange={e => set("endDate")(e.target.value)} /></FG>
          </div>
          <div className="form-row"><FG label="Description"><input value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Notes"><input value={f.notes} onChange={e => set("notes")(e.target.value)} /></FG></div>
        </Modal>
      )}

      {linkFor && (
        <LinkPaymentModal
          project={summarized.find(p => p.id === linkFor.id) || linkFor}
          items={poolItems}
          onClose={() => setLinkFor(null)}
          onLink={linkPayment}
        />
      )}

      {assignItem && (
        <AssignPaymentModal
          item={poolItems.find(x => x.key === assignItem.key) || assignItem}
          projects={summarized}
          onClose={() => setAssignItem(null)}
          onAssign={async (project, amount) => { const ok = await linkPayment(project, poolItems.find(x => x.key === assignItem.key) || assignItem, amount); if (ok) setAssignItem(null); }}
        />
      )}

      {summaryOpen && (
        <Modal title="Vendor Summary PDF" onClose={() => setSummaryOpen(false)}
          foot={<><button className="btn btn-ghost" onClick={() => setSummaryOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={() => exportVendorSummary(summaryWithTx)}>📄 Download PDF</button></>}>
          <p style={{ fontSize: 12.5, color: "var(--text2)", lineHeight: 1.6, marginBottom: 12 }}>
            One line per vendor with <strong>Total</strong> (contract value), <strong>Paid</strong> and <strong>Left</strong>. Covers the {vendorGroups.length} vendor(s) matching your current search and status filter.
          </p>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: "var(--text)" }}>
            <input type="checkbox" checked={summaryWithTx} onChange={e => setSummaryWithTx(e.target.checked)} />
            Include transactions (every project and payment under each vendor)
          </label>
          <p className="mono" style={{ fontSize: 11, color: "var(--text3)", marginTop: 10 }}>
            {summaryWithTx ? "Page 1 is the summary; the transactions follow on the next pages." : "Summary only — a single page for most setups."}
          </p>
        </Modal>
      )}

      {cashFor && <PartnerCashModal project={cashFor} partners={partners} onClose={() => setCashFor(null)} onSave={d => savePartnerCash(cashFor, d)} />}
      {vendorEditItem && <SetVendorModal item={vendorEditItem} vendorNames={vendorNames} onClose={() => setVendorEditItem(null)} onSave={name => saveItemVendor(vendorEditItem, name)} />}

      {manualFor && <ManualPaymentModal project={manualFor} onClose={() => setManualFor(null)} onSave={data => addManualPayment(manualFor, data)} />}

      {stmtFor && (() => {
        const p = summarized.find(x => x.id === stmtFor.id) || stmtFor;
        return (
          <Modal title={`Project Statement: ${p.projectNo} — ${p.vendorName}`} size="modal-lg" onClose={() => setStmtFor(null)}
            foot={<button className="btn btn-ghost" onClick={() => exportStatement(p)}>📄 Export PDF</button>}>
            <div style={{ marginBottom: 6, fontSize: 13, fontWeight: 600 }}>{p.title}</div>
            <div style={{ marginBottom: 16, display: "flex", gap: 20, fontFamily: "var(--mono)", fontSize: 13, background: "var(--bg3)", padding: 12, borderRadius: "var(--r)", flexWrap: "wrap" }}>
              <div>Contract: <span style={{ fontWeight: 700 }}>{p.hasValue ? fmt(p.contractValue) : "Open-ended"}</span></div>
              <div>Paid: <span style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(p.paid)}</span></div>
              {p.hasValue && <div>Due: <span style={{ color: "var(--amber)", fontWeight: 700 }}>{fmt(p.due)}</span></div>}
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Date</th><th>Source</th><th>Details</th><th className="r">Amount</th></tr></thead>
                <tbody>
                  {p.payments.length === 0 ? <tr><td colSpan={4} style={{ textAlign: "center", color: "var(--text3)", padding: 16 }}>No payments linked yet</td></tr>
                    : p.payments.map(x => { const info = paymentRowInfo(x); return (
                      <tr key={x.id}><td className="mono">{x.date}</td><td><BadgeComponent type={info.badge}>{info.kind}</BadgeComponent></td><td>{info.text}</td><td className="r mono" style={{ fontWeight: 700 }}>{fmt(x.amount)}</td></tr>
                    ); })}
                </tbody>
              </table>
            </div>
          </Modal>
        );
      })()}
    </div>
  );
}

function PartnerCashModal({ project, partners, onClose, onSave }) {
  const [pf, setPf] = useState({ partnerName: partners[0] || "", date: today(), amount: "", description: "", ref: "" });
  const setP = k => v => setPf(x => ({ ...x, [k]: v }));
  const isCapex = CAP_CATS.some(c => c.id === project.category);
  return (
    <Modal title={`Partner Cash Payment: ${project.vendorName} — ${project.title}`} onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(pf)}>Save & Link</button></>}>
      <p style={{ fontSize: 11.5, color: "var(--text3)", marginBottom: 12, lineHeight: 1.5 }}>
        A partner paid {project.vendorName} in cash. This records it as {isCapex ? "a capital expense" : "an expense"} against the vendor, credits the partner's
        current account, and links the full amount to this project.
      </p>
      <div className="form-row cols-2">
        <FG label="Paid by (partner) *">
          <select value={pf.partnerName} onChange={e => setP("partnerName")(e.target.value)}>
            {partners.length === 0 && <option value="">No partners found</option>}
            {partners.map(n => <option key={n} value={n}>{n}</option>)}
          </select>
        </FG>
        <FG label="Date"><input type="date" value={pf.date} onChange={e => setP("date")(e.target.value)} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Amount (₹) *"><input type="number" value={pf.amount} onChange={e => setP("amount")(e.target.value)} /></FG>
        <FG label="Reference"><input value={pf.ref} onChange={e => setP("ref")(e.target.value)} /></FG>
      </div>
      <div className="form-row"><FG label="Description *"><input value={pf.description} onChange={e => setP("description")(e.target.value)} placeholder="e.g. Cash advance for labour" /></FG></div>
    </Modal>
  );
}

function SetVendorModal({ item, vendorNames, onClose, onSave }) {
  const [name, setName] = useState(item.vendor || "");
  return (
    <Modal title="Set Vendor for Cash Payment" onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(name)}>Save Vendor</button></>}>
      <div style={{ background: "var(--bg3)", padding: 12, borderRadius: "var(--r)", marginBottom: 14, fontSize: 12.5 }}>
        <div style={{ fontWeight: 700 }}>{fmt(item.amount)} · {item.sourceLabel}</div>
        <div className="mono" style={{ color: "var(--text3)", fontSize: 11, marginTop: 3 }}>{item.date} · {item.description}</div>
        <div className="mono" style={{ color: "var(--text3)", fontSize: 11, marginTop: 3 }}>Currently under: {item.vendor || "(no vendor)"}</div>
      </div>
      <div className="form-row">
        <FG label="Vendor *" note="Pick the vendor who received the cash so it lines up with their projects. Only the vendor name on this entry changes.">
          <input list="vp-setvendor" value={name} onChange={e => setName(e.target.value)} />
          <datalist id="vp-setvendor">{vendorNames.map(n => <option key={n} value={n} />)}</datalist>
        </FG>
      </div>
    </Modal>
  );
}

// One row in the "link a payment" list, with its own amount box so a payment can be split.
function LinkRow({ item, onLink }) {
  const [amt, setAmt] = useState(String(Math.round(item.remaining * 100) / 100));
  const [busy, setBusy] = useState(false);
  useEffect(() => { setAmt(String(Math.round(item.remaining * 100) / 100)); }, [item.remaining]);
  return (
    <tr>
      <td className="mono" style={{ fontSize: 11, color: "var(--text3)" }}>{item.date}</td>
      <td style={{ fontSize: 12 }}><div style={{ fontWeight: 600 }}>{item.vendor || "—"}</div><div style={{ color: "var(--text3)", fontSize: 11 }}>{item.description}</div></td>
      <td><BadgeComponent type={item.sourceType === "bank" ? "blue" : "accent"}>{item.sourceLabel}</BadgeComponent></td>
      <td className="r mono" style={{ fontSize: 11.5 }}>{fmt(item.amount)}{item.allocated > EPS && <div style={{ color: "var(--text3)", fontSize: 10 }}>{fmt(item.remaining)} left</div>}</td>
      <td className="r"><input type="number" value={amt} onChange={e => setAmt(e.target.value)} style={{ width: 110, background: "var(--bg3)", border: "1px solid var(--border2)", borderRadius: "var(--r)", padding: "5px 8px", color: "var(--text)", textAlign: "right" }} /></td>
      <td className="r"><button className="btn btn-primary btn-sm" disabled={busy} onClick={async () => { setBusy(true); await onLink(item, amt); setBusy(false); }}>Link</button></td>
    </tr>
  );
}

function LinkPaymentModal({ project, items, onClose, onLink }) {
  const [search, setSearch] = useState("");
  const [showOthers, setShowOthers] = useState(false);
  const vk = normVendor(project.vendorName);
  const matching = items.filter(x => normVendor(x.vendor) === vk && x.remaining > EPS);
  const rows = useMemo(() => {
    const s = search.toLowerCase();
    return items
      .filter(x => x.remaining > EPS)
      .filter(x => showOthers || normVendor(x.vendor) === vk)
      .filter(x => !s || x.vendor.toLowerCase().includes(s) || x.description.toLowerCase().includes(s))
      .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
      .slice(0, 200);
  }, [items, search, showOthers, vk]);

  return (
    <Modal title={`Link Payments: ${project.vendorName} — ${project.title}`} size="modal-lg" onClose={onClose}>
      <div style={{ display: "flex", gap: 20, fontFamily: "var(--mono)", fontSize: 12, background: "var(--bg3)", padding: 10, borderRadius: "var(--r)", flexWrap: "wrap", marginBottom: 12 }}>
        <div>Paid so far: <span style={{ color: "var(--green)", fontWeight: 700 }}>{fmt(project.paid)}</span></div>
        {project.hasValue && <div>Due: <span style={{ color: "var(--amber)", fontWeight: 700 }}>{fmt(project.due)}</span></div>}
      </div>
      <div className="filter-bar" style={{ marginBottom: 10 }}>
        <SearchBar value={search} onChange={setSearch} placeholder="Search narration or vendor…" style={{ minWidth: 220 }} />
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "var(--text2)" }}>
          <input type="checkbox" checked={showOthers} onChange={e => setShowOthers(e.target.checked)} />
          Show other vendors' payments
        </label>
      </div>
      <p style={{ fontSize: 11, color: "var(--text3)", marginBottom: 10 }}>
        {showOthers ? "Showing every unassigned payment." : `${matching.length} unassigned payment(s) to ${project.vendorName}.`} Change the amount to link only part of a payment — the rest stays available for another project.
        Payments made through a different name? Tick "Show other vendors' payments".
      </p>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Date</th><th>Vendor / Narration</th><th>Source</th><th className="r">Payment</th><th className="r">Link Amount</th><th className="r"></th></tr></thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={6} style={{ textAlign: "center", color: "var(--text3)", padding: 20 }}>No unassigned payments found{showOthers ? "" : " for this vendor"}</td></tr>
            ) : rows.map(x => <LinkRow key={x.key} item={x} onLink={(item, amt) => onLink(project, item, amt)} />)}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

function AssignPaymentModal({ item, projects, onClose, onAssign }) {
  const vk = normVendor(item.vendor);
  const mine = projects.filter(p => normVendor(p.vendorName) === vk);
  const others = projects.filter(p => normVendor(p.vendorName) !== vk);
  const [projectId, setProjectId] = useState(String((mine[0] || others[0] || {}).id ?? ""));
  const [amount, setAmount] = useState(String(Math.round(item.remaining * 100) / 100));
  const project = projects.find(p => String(p.id) === projectId);
  return (
    <Modal title="Assign Payment to a Project" onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" disabled={!project} onClick={() => onAssign(project, amount)}>Assign</button></>}>
      <div style={{ background: "var(--bg3)", padding: 12, borderRadius: "var(--r)", marginBottom: 14, fontSize: 12.5 }}>
        <div style={{ fontWeight: 700 }}>{item.vendor || "(no vendor name)"} · {fmt(item.amount)}</div>
        <div className="mono" style={{ color: "var(--text3)", fontSize: 11, marginTop: 3 }}>{item.date} · {item.description}</div>
        <div className="mono" style={{ color: "var(--amber)", fontSize: 11, marginTop: 3 }}>{fmt(item.remaining)} not yet assigned</div>
      </div>
      <div className="form-row">
        <FG label="Project">
          <select value={projectId} onChange={e => setProjectId(e.target.value)}>
            {mine.length > 0 && <optgroup label={`${item.vendor}'s projects`}>{mine.map(p => <option key={p.id} value={p.id}>{p.projectNo} — {p.title}</option>)}</optgroup>}
            {others.length > 0 && <optgroup label="Other vendors' projects">{others.map(p => <option key={p.id} value={p.id}>{p.projectNo} — {p.vendorName}: {p.title}</option>)}</optgroup>}
          </select>
        </FG>
      </div>
      <div className="form-row">
        <FG label="Amount to assign (₹)" note="Assign less than the full amount to split this payment across projects."><input type="number" value={amount} onChange={e => setAmount(e.target.value)} /></FG>
      </div>
    </Modal>
  );
}

function ManualPaymentModal({ project, onClose, onSave }) {
  const [pf, setPf] = useState({ date: today(), amount: "", paymentMode: "cash", reference: "", remarks: "" });
  const setP = k => v => setPf(x => ({ ...x, [k]: v }));
  return (
    <Modal title={`Manual Payment: ${project.vendorName} — ${project.title}`} onClose={onClose}
      foot={<><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={() => onSave(pf)}>Save Payment</button></>}>
      <p style={{ fontSize: 11.5, color: "var(--text3)", marginBottom: 12, lineHeight: 1.5 }}>
        Records a payment on this project only. It does <strong>not</strong> appear in the Bank Statement, Expenses or Balance Sheet — use it for money already
        accounted for elsewhere or tracked only at project level. For cash paid by a partner, record it under Partners & Capital and link it instead.
      </p>
      <div className="form-row cols-2">
        <FG label="Date"><input type="date" value={pf.date} onChange={e => setP("date")(e.target.value)} /></FG>
        <FG label="Amount (₹) *"><input type="number" value={pf.amount} onChange={e => setP("amount")(e.target.value)} /></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Mode">
          <select value={pf.paymentMode} onChange={e => setP("paymentMode")(e.target.value)}>
            <option value="cash">Cash</option><option value="bank">Bank / UPI</option><option value="cheque">Cheque</option><option value="other">Other</option>
          </select>
        </FG>
        <FG label="Reference"><input value={pf.reference} onChange={e => setP("reference")(e.target.value)} /></FG>
      </div>
      <div className="form-row"><FG label="Remarks"><input value={pf.remarks} onChange={e => setP("remarks")(e.target.value)} /></FG></div>
    </Modal>
  );
}
