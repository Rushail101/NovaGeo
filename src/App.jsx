import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import * as db from "./db.js";

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
}
@media (max-width: 480px) {
  .stats-grid, .stats-grid-5 {
    grid-template-columns: 1fr;
  }
}
`;



// ── Helpers ───────────────────────────────────────────────────────────────────
const SHIFTS = ["Morning (6AM–2PM)","Afternoon (2PM–10PM)","Night (10PM–6AM)"];
const OPERATORS = ["Suresh Kumar","Manoj Singh","Ravi Patil","Dinesh Yadav"];
const fmtMT = v => `${(+v||0).toFixed(3)} MT`;
const fmtWt = kg => kg>=1000?`${(kg/1000).toFixed(3)} MT`:`${kg} kg`;
const fmt = n => `₹${(+n||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const today = () => new Date().toISOString().split("T")[0];
const nowTime = () => new Date().toTimeString().slice(0,5);

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
  { id:"machinery", label:"Machinery & Equipment" },
  { id:"civil", label:"Civil / Foundation Works" },
  { id:"land", label:"Land Acquisition & Registry" },
  { id:"vehicle", label:"Vehicles & Loaders" },
  { id:"license", label:"Licenses & Approvals" },
  { id:"deposit", label:"Electricity/Utility Deposits" },
  { id:"office", label:"Furniture & Site Office" },
  { id:"other", label:"Other Capital Assets" },
];
const CP_TYPES = [
  { id:"unlabeled", label:"Unlabeled" }, { id:"vendor", label:"Vendor / Supplier" },
  { id:"customer", label:"Customer / Buyer" }, { id:"staff", label:"Staff / Labour" },
  { id:"owner", label:"Owner / Capital" }, { id:"loan", label:"Loan / Bank" },
  { id:"utility", label:"Utility / Rent" }, { id:"tax", label:"Tax / Government" },
  { id:"other", label:"Other" },
];

function computeStock(grades, receipts, entries, weighments) {
  const s = {};
  grades.forEach(g => { s[g.id] = { ...g, gradeId: g.id, in: 0, produced: 0, sold: 0, transferred: 0 }; });
  receipts.forEach(r => { const b = grades.find(g => g.isBoulder); if (b && s[b.id]) s[b.id].in += r.quantityMT; });
  entries.forEach(pe => {
    const b = grades.find(g => g.isBoulder);
    if (b && s[b.id]) s[b.id].produced -= pe.boulderConsumedMT;
    (pe.outputs || []).forEach(o => { if (s[o.gradeId]) s[o.gradeId].produced += o.quantityMT; });
  });
  weighments.forEach(w => {
    const mt = (w.netWeight || 0) / 1000;
    if (w.purpose === "Sale" && s[w.gradeId]) s[w.gradeId].sold += mt;
    if (w.purpose === "Inward (Raw)" && s[w.gradeId]) s[w.gradeId].in += mt;
    if (w.purpose === "Stock Transfer" && s[w.gradeId]) s[w.gradeId].transferred += mt;
  });
  Object.values(s).forEach(g => { g.closing = g.isBoulder ? g.in + g.produced : g.produced - g.sold - g.transferred; });
  return s;
}

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
  { id:"weighment", label:"Weighment", icon:"⚖", group:"Operations" },
  { id:"boulder", label:"Boulder In", icon:"⬡", group:"Operations" },
  { id:"production", label:"Production", icon:"⚙", group:"Operations" },
  { id:"stock", label:"Stock Ledger", icon:"▤", group:"Operations" },
  { id:"shift", label:"Shift Log", icon:"◷", group:"Operations" },
  { id:"purchases", label:"Purchases", icon:"📥", group:"Commercial" },
  { id:"opscosts", label:"Monthly Ops Costs", icon:"🧾", group:"Commercial" },
  { id:"expenses", label:"Expenses", icon:"💸", group:"Commercial" },
  { id:"capital", label:"Capital & Infra", icon:"🏗", group:"Commercial" },
  { id:"bankstatement", label:"Bank Statement", icon:"🏦", group:"Banking" },
  { id: "partners", label: "Partners & Capital", icon: "🤝", group: "Banking" },
  { id:"balancesheet", label:"Balance Sheet", icon:"🏛", group:"Reports" },
  { id:"costs", label:"Cost Sheet", icon:"💰", group:"Reports" },
  { id:"costconfig", label:"Cost Config", icon:"⚙", group:"Reports" },
  { id:"vehicles", label:"Vehicles", icon:"◉", group:"Masters" },
  { id:"customers", label:"Customers", icon:"◈", group:"Masters" },
  { id:"suppliers", label:"Suppliers", icon:"⛏", group:"Masters" },
  { id:"grades", label:"Grades", icon:"💎", group:"Masters" },
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

// ── APP MAIN ──────────────────────────────────────────────────────────────────
export default function App() {
  const [view, setView] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [loading, setLoading] = useState(true);

  const [vehicles, setVehicles] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [grades, setGrades] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [weighments, setWeighments] = useState([]);
  const [shiftLogs, setShiftLogs] = useState([]);
  const [boulderReceipts, setBoulderReceipts] = useState([]);
  const [productionEntries, setProductionEntries] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [capitalItems, setCapitalItems] = useState([]);
  const [bankBatches, setBankBatches] = useState([]);
  const [bankTxns, setBankTxns] = useState([]);
  const [bankLabels, setBankLabels] = useState({});
  const [equipment, setEquipment] = useState([]);
  const [maintTasks, setMaintTasks] = useState([]);
  const [maintLogs, setMaintLogs] = useState([]);
  const [lots, setLots] = useState([]);
  const [costConfig, setCostConfig] = useState({});
  const [qcTests, setQcTests] = useState([]);
  const [samples, setSamples] = useState([]);
  const [partnerCashbook, setPartnerCashbook] = useState([]);

  useEffect(() => {
    async function init() {
      try {
        const d = await db.loadAll();
        setVehicles(d.vehicles || []); setCustomers(d.customers || []); setGrades(d.grades || []);
        setSuppliers(d.suppliers || []); setEquipment(d.equipment || []); setMaintTasks(d.maintTasks || []);
        setMaintLogs(d.maintLogs || []); setWeighments(d.weighments || []); setBoulderReceipts(d.boulderReceipts || []);
        setLots(d.lots || []); setProductionEntries(d.productionEntries || []); setShiftLogs(d.shiftLogs || []);
        setPurchases(d.purchases || []); setExpenses(d.expenses || []); setCapitalItems(d.capitalItems || []);
        setBankBatches(d.bankBatches || []); setBankTxns(d.bankTxns || []); setBankLabels(d.bankLabels || {});
        setCostConfig(d.costConfig || {}); setQcTests(d.qcTests || []); setSamples(d.samples || []);
        setPartnerCashbook(d.partnerCashbook || []);
      } catch (err) {
        console.error("Database sync failed:", err);
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
              <div><h1>BallMill ERP</h1><p>Nova Geo Resources</p></div>
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
            {view === "dashboard" && <DashboardView {...{ weighments, boulderReceipts, productionEntries, expenses, capitalItems, bankTxns, partnerCashbook }} />}
            {view === "weighment" && <WeighmentsView {...{ vehicles, customers, grades, weighments, setWeighments }} />}
            {view === "boulder" && <BoulderInView {...{ suppliers, vehicles, boulderReceipts, setBoulderReceipts }} />}
            {view === "production" && <ProductionView {...{ grades, productionEntries, setProductionEntries, lots, setLots }} />}
            {view === "stock" && <StockLedgerView {...{ grades, boulderReceipts, productionEntries, weighments }} />}
            {view === "shift" && <ShiftLogView {...{ grades, shiftLogs, setShiftLogs }} />}
            {view === "purchases" && <PurchasesView {...{ suppliers, purchases, setPurchases }} />}
            {view === "opscosts" && <MonthlyOpsCostsView {...{ expenses, setExpenses }} />}
            {view === "expenses" && <ExpensesView {...{ expenses, setExpenses }} />}
            {view === "capital" && <CapitalRegisterView {...{ capitalItems, setCapitalItems }} />}
            {view === "bankstatement" && <BankStatementGroupingView {...{ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, setCapitalItems, setExpenses, setPartnerCashbook }} />}
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
            {view === "balancesheet" && <BalanceSheetView {...{ capitalItems, bankTxns, partnerCashbook, purchases, expenses, weighments, grades, boulderReceipts, productionEntries, costConfig }} />}
            {view === "costs" && <CostSheetView {...{ purchases, expenses, productionEntries, weighments, costConfig }} />}
            {view === "costconfig" && <CostConfigView {...{ costConfig, setCostConfig }} />}
            {view === "vehicles" && <MasterTableView title="Vehicles" noun="Vehicle" items={vehicles} setItems={setVehicles} saveFn={db.saveVehicle} delFn={db.deleteVehicle} fields={[{ key:"vehicleNo", label:"Vehicle Number", required:true },{ key:"type", label:"Type" },{ key:"tareWeight", label:"Tare Wt (kg)", required:true }]} cols={[{ label:"Vehicle", key:"vehicleNo" },{ label:"Type", key:"type" },{ label:"Tare Wt", key:"tareWeight" }]} />}
            {view === "customers" && <MasterTableView title="Customers" noun="Customer" items={customers} setItems={setCustomers} saveFn={db.saveCustomer} delFn={db.deleteCustomer} fields={[{ key:"name", label:"Company Name", required:true },{ key:"gstin", label:"GSTIN" },{ key:"contact", label:"Contact Person" },{ key:"phone", label:"Phone" }]} cols={[{ label:"Name", key:"name" },{ label:"GSTIN", key:"gstin" },{ label:"Contact", key:"contact" }]} />}
            {view === "suppliers" && <MasterTableView title="Suppliers" noun="Supplier" items={suppliers} setItems={setSuppliers} saveFn={db.saveSupplier} delFn={db.deleteSupplier} fields={[{ key:"name", label:"Supplier Name", required:true },{ key:"contact", label:"Contact" }]} cols={[{ label:"Name", key:"name" },{ label:"Contact", key:"contact" }]} />}
            {view === "grades" && <MasterTableView title="Grades" noun="Grade" items={grades} setItems={setGrades} saveFn={db.saveGrade} delFn={db.deleteGrade} fields={[{ key:"code", label:"Code", required:true },{ key:"name", label:"Grade Name", required:true }]} cols={[{ label:"Code", key:"code" },{ label:"Name", key:"name" }]} />}
          </div>
        </div>
      </div>
    </>
  );
}

// ── OPERATIONS VIEWS WITH RESTORED CREATION MODALS ────────────────────────────

function WeighmentsView({ vehicles, customers, grades, weighments, setWeighments }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), time:nowTime(), vehicleId:"", customerId:"", gradeId:"", lotNo:"", grossWeight:"", tareWeight:"", purpose:"Sale", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));
  const pickV = id => { const v = vehicles.find(x => String(x.id) === String(id)); setF(x => ({ ...x, vehicleId:id, tareWeight: v ? String(v.tareWeight) : "" })); };
  const net = () => { const g = parseFloat(f.grossWeight), t = parseFloat(f.tareWeight); return (!isNaN(g) && !isNaN(t)) ? g - t : null; };

  async function save() {
    if (!f.vehicleId || !f.gradeId || !f.grossWeight || !f.tareWeight) return alert("Vehicle, Grade, Gross, and Tare required.");
    if (net() <= 0) return alert("Net weight must be positive.");
    try {
      const row = await db.saveWeighment(f);
      setWeighments(ws => [row, ...ws]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Weighment Slip</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Slip #</th><th>Date/Time</th><th>Vehicle</th><th>Customer</th><th>Grade</th><th className="r">Gross</th><th className="r">Tare</th><th className="r">Net</th><th>Purpose</th></tr></thead>
          <tbody>
            {weighments.length === 0 ? (
              <tr><td colSpan={9} style={{ textAlign:"center", color:"var(--text3)", padding:24 }}>No weighment slips logged yet</td></tr>
            ) : (
              weighments.map(w => (
                <tr key={w.id}>
                  <td className="mono" style={{ color:"var(--accent)" }}>{w.slipNo}</td>
                  <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{w.date} {w.time}</td>
                  <td className="mono">{w.vehicleNo}</td><td>{w.customerName}</td>
                  <td><BadgeComponent type="muted">{w.gradeCode}</BadgeComponent></td>
                  <td className="r mono">{w.grossWeight} kg</td><td className="r mono">{w.tareWeight} kg</td>
                  <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtWt(w.netWeight)}</td>
                  <td><BadgeComponent type={w.purpose === "Sale" ? "green" : "blue"}>{w.purpose}</BadgeComponent></td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="New Weighment Slip" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Slip</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Time"><input type="time" value={f.time} onChange={e => set("time")(e.target.value)} /></FG></div>
          <div className="form-row cols-2">
            <FG label="Vehicle *"><select value={f.vehicleId} onChange={e => pickV(e.target.value)}><option value="">Select...</option>{vehicles.map(v => <option key={v.id} value={v.id}>{v.vehicleNo} ({v.type})</option>)}</select></FG>
            <FG label="Customer"><select value={f.customerId} onChange={e => set("customerId")(e.target.value)}><option value="">Select...</option>{customers.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Grade *"><select value={f.gradeId} onChange={e => set("gradeId")(e.target.value)}><option value="">Select...</option>{grades.map(g => <option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></FG>
            <FG label="Purpose"><select value={f.purpose} onChange={e => set("purpose")(e.target.value)}>{["Sale","Inward (Raw)","Stock Transfer"].map(p => <option key={p}>{p}</option>)}</select></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Gross Weight (kg) *"><input type="number" value={f.grossWeight} onChange={e => set("grossWeight")(e.target.value)} /></FG>
            <FG label="Tare Weight (kg) *"><input type="number" value={f.tareWeight} onChange={e => set("tareWeight")(e.target.value)} /></FG>
          </div>
          <div className="form-row"><FG label="Remarks"><input value={f.remarks} onChange={e => set("remarks")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function BoulderInView({ suppliers, vehicles, boulderReceipts, setBoulderReceipts }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), vehicleId:"", supplierId:"", quantityMT:"", royaltyNo:"", challanNo:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.supplierId || !f.quantityMT) return alert("Supplier and Quantity required.");
    try {
      const row = await db.saveBoulderReceipt(f);
      setBoulderReceipts(rs => [row, ...rs]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Boulder Receipt</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Receipt #</th><th>Date</th><th>Supplier</th><th>Vehicle</th><th>Royalty #</th><th>Challan #</th><th className="r">Qty (MT)</th></tr></thead>
          <tbody>
            {boulderReceipts.length === 0 ? (
              <tr><td colSpan={7} style={{ textAlign:"center", color:"var(--text3)", padding:24 }}>No boulder deliveries recorded yet</td></tr>
            ) : (
              boulderReceipts.map(r => (
                <tr key={r.id}>
                  <td className="mono" style={{ color:"var(--accent)" }}>{r.receiptNo}</td>
                  <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{r.date}</td>
                  <td>{r.supplierName}</td><td className="mono">{r.vehicleNo}</td><td>{r.royaltyNo || "—"}</td><td>{r.challanNo || "—"}</td>
                  <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtMT(r.quantityMT)}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="New Raw Boulder Receipt" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Receipt</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Supplier *"><select value={f.supplierId} onChange={e => set("supplierId")(e.target.value)}><option value="">Select...</option>{suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Vehicle"><select value={f.vehicleId} onChange={e => set("vehicleId")(e.target.value)}><option value="">Select...</option>{vehicles.map(v => <option key={v.id} value={v.id}>{v.vehicleNo}</option>)}</select></FG><FG label="Quantity (MT) *"><input type="number" step="0.001" value={f.quantityMT} onChange={e => set("quantityMT")(e.target.value)} /></FG></div>
          <div className="form-row cols-2"><FG label="Royalty Pass #"><input value={f.royaltyNo} onChange={e => set("royaltyNo")(e.target.value)} /></FG><FG label="Challan / Invoice #"><input value={f.challanNo} onChange={e => set("challanNo")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function ProductionView({ grades, productionEntries, setProductionEntries, lots, setLots }) {
  const finished = grades.filter(g => !g.isBoulder);
  const blankOut = () => ({ gradeId: finished[0]?.id || "", quantityMT: "" });
  const blank = { date:today(), shift:SHIFTS[0], operator:OPERATORS[0], boulderConsumedMT:"", machineHours:"", idleHours:"0", outputs:[blankOut()], remarks:"" };
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));
  const setOut = (i, k, v) => setF(x => { const o = [...x.outputs]; o[i] = { ...o[i], [k]: v }; return { ...x, outputs: o }; });
  const totalOut = f.outputs.reduce((s, o) => s + (parseFloat(o.quantityMT) || 0), 0);
  const yieldPct = f.boulderConsumedMT > 0 ? ((totalOut / +f.boulderConsumedMT) * 100).toFixed(1) : null;

  async function save() {
    if (!f.boulderConsumedMT) return alert("Boulder consumed required.");
    if (!f.outputs.every(o => o.gradeId && o.quantityMT)) return alert("Fill all output rows.");
    try {
      const payloadOutputs = f.outputs.map(o => {
        const g = grades.find(gx => String(gx.id) === String(o.gradeId));
        return { gradeId: +o.gradeId, quantityMT: +o.quantityMT, gradeName: g?.name, gradeCode: g?.code };
      });
      const res = await db.saveProductionEntry({ ...f, totalOutputMT: totalOut, yieldPct: yieldPct ? +yieldPct : null }, payloadOutputs);
      setLots(ls => [res.lot, ...ls]);
      setProductionEntries(es => [res.entry, ...es]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Production Entry</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Entry #</th><th>Lot #</th><th>Date</th><th>Shift</th><th>Operator</th><th className="r">Consumed</th><th className="r">Produced</th><th>Yield</th></tr></thead>
          <tbody>
            {productionEntries.length === 0 ? (
              <tr><td colSpan={8} style={{ textAlign:"center", color:"var(--text3)", padding:24 }}>No production recorded yet</td></tr>
            ) : (
              productionEntries.map(pe => (
                <tr key={pe.id}>
                  <td className="mono" style={{ color:"var(--accent)" }}>{pe.entryNo}</td>
                  <td><span className="lot-chip">🏷 {pe.lotNo}</span></td>
                  <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{pe.date}</td>
                  <td><BadgeComponent type="muted">{(pe.shift || "").split(" ")[0]}</BadgeComponent></td>
                  <td>{pe.operator}</td>
                  <td className="r mono">{fmtMT(pe.boulderConsumedMT)}</td>
                  <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtMT(pe.totalOutputMT)}</td>
                  <td>{pe.yieldPct ? <BadgeComponent type={pe.yieldPct >= 90 ? "green" : "amber"}>{pe.yieldPct}%</BadgeComponent> : "—"}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="New Milling Run / Production Entry" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Entry</button></>}>
          <div className="form-row cols-3"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Shift"><select value={f.shift} onChange={e => set("shift")(e.target.value)}>{SHIFTS.map(s => <option key={s}>{s}</option>)}</select></FG><FG label="Operator"><select value={f.operator} onChange={e => set("operator")(e.target.value)}>{OPERATORS.map(o => <option key={o}>{o}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Boulder Consumed (MT) *"><input type="number" step="0.001" value={f.boulderConsumedMT} onChange={e => set("boulderConsumedMT")(e.target.value)} /></FG><FG label="Machine Running Hours"><input type="number" value={f.machineHours} onChange={e => set("machineHours")(e.target.value)} /></FG></div>
          <div style={{ marginBottom: 14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom: 8 }}><label style={{ fontSize:11, color:"var(--text2)", textTransform:"uppercase", fontFamily:"var(--mono)" }}>Produced Outputs by Grade</label><button className="btn btn-ghost btn-sm" onClick={() => setF(x => ({ ...x, outputs:[...x.outputs, blankOut()] }))}>+ Add Grade Row</button></div>
            <div className="output-rows">
              {f.outputs.map((row, i) => (
                <div key={i} className="output-row">
                  <div className="form-group"><select value={row.gradeId} onChange={e => setOut(i, "gradeId", e.target.value)}>{finished.map(g => <option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></div>
                  <div className="form-group"><input type="number" step="0.001" placeholder="Quantity (MT)" value={row.quantityMT} onChange={e => setOut(i, "quantityMT", e.target.value)} /></div>
                  <button className="btn btn-danger btn-sm" onClick={() => setF(x => ({ ...x, outputs: x.outputs.filter((_, idx) => idx !== i) }))}>✕</button>
                </div>
              ))}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function ShiftLogView({ grades, shiftLogs, setShiftLogs }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), shift:SHIFTS[0], operator:OPERATORS[0], machineStatus:"Running", boulderFed:"", production:[{ gradeId:grades[0]?.id||"", quantity:"" }], breakdowns:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.boulderFed) return alert("Boulder fed is required.");
    try {
      const enriched = f.production.map(p => {
        const g = grades.find(gx => String(gx.id) === String(p.gradeId));
        return { gradeId: +p.gradeId, quantity: +p.quantity, gradeCode: g?.code, gradeName: g?.name };
      });
      const row = await db.saveShiftLog(f, enriched);
      setShiftLogs(ls => [row, ...ls]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Shift Log</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Log #</th><th>Date</th><th>Shift</th><th>Operator</th><th>Status</th><th className="r">Fed (MT)</th><th>Breakdowns</th></tr></thead>
          <tbody>
            {shiftLogs.length === 0 ? (
              <tr><td colSpan={7} style={{ textAlign:"center", color:"var(--text3)", padding:24 }}>No shift logs recorded yet</td></tr>
            ) : (
              shiftLogs.map(l => (
                <tr key={l.id}>
                  <td className="mono" style={{ color:"var(--accent)" }}>{l.logNo}</td>
                  <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{l.date}</td>
                  <td><BadgeComponent type="muted">{(l.shift || "").split(" ")[0]}</BadgeComponent></td>
                  <td>{l.operator}</td>
                  <td><BadgeComponent type={l.machineStatus === "Running" ? "green" : "red"}>{l.machineStatus}</BadgeComponent></td>
                  <td className="r mono">{l.boulderFed}</td>
                  <td style={{ fontSize:12 }}>{l.breakdowns || "None"}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="Record Shift Operating Log" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Log</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Shift"><select value={f.shift} onChange={e => set("shift")(e.target.value)}>{SHIFTS.map(s => <option key={s}>{s}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Operator"><select value={f.operator} onChange={e => set("operator")(e.target.value)}>{OPERATORS.map(o => <option key={o}>{o}</option>)}</select></FG><FG label="Boulder Fed (MT) *"><input type="number" value={f.boulderFed} onChange={e => set("boulderFed")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Breakdowns / Downtime"><input placeholder="e.g. 30 min classifier mesh choke" value={f.breakdowns} onChange={e => set("breakdowns")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function StockLedgerView({ grades, boulderReceipts, productionEntries, weighments }) {
  const stock = useMemo(() => computeStock(grades, boulderReceipts, productionEntries, weighments), [grades, boulderReceipts, productionEntries, weighments]);
  return (
    <div className="table-wrap">
      <div className="table-toolbar"><h3>Material Inventory Balance (Real-Time Calculation)</h3></div>
      <table>
        <thead><tr><th>Grade</th><th>Type</th><th className="r">Total Inflow / Produced</th><th className="r">Total Outflow / Sold</th><th className="r">Closing Stock Balance</th><th>Status</th></tr></thead>
        <tbody>
          {Object.values(stock).map(s => (
            <tr key={s.gradeId}>
              <td><BadgeComponent type={s.isBoulder ? "teal" : "muted"}>{s.code}</BadgeComponent> {s.name}</td>
              <td><BadgeComponent type={s.isBoulder ? "teal" : "blue"}>{s.isBoulder ? "Raw Material" : "Finished Product"}</BadgeComponent></td>
              <td className="r mono" style={{ color:"var(--teal)" }}>{fmtMT(s.isBoulder ? s.in : s.produced)}</td>
              <td className="r mono" style={{ color:"var(--amber)" }}>{fmtMT(s.isBoulder ? Math.abs(s.produced) : s.sold)}</td>
              <td className="r mono" style={{ fontWeight:700, color: s.closing < 0 ? "var(--red)" : "var(--accent)" }}>{fmtMT(s.closing)}</td>
              <td><BadgeComponent type={s.closing < 0 ? "red" : s.closing < 10 ? "amber" : "green"}>{s.closing < 0 ? "Negative" : s.closing < 10 ? "Low" : "Optimal"}</BadgeComponent></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ── COMMERCIAL & REPORTING VIEWS ──────────────────────────────────────────────

function PurchasesView({ suppliers, purchases, setPurchases }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), supplierId:"", description:"", quantityMT:"", ratePerMT:"", gstPct:"5", invoiceNo:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.supplierId || !f.quantityMT || !f.ratePerMT) return alert("Supplier, Quantity, and Rate required.");
    try {
      const row = await db.savePurchase(f);
      setPurchases(ps => [row, ...ps]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New RM Purchase Order</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>PO #</th><th>Date</th><th>Supplier</th><th className="r">Qty (MT)</th><th className="r">Rate</th><th className="r">Taxable</th><th className="r">Total</th></tr></thead>
          <tbody>
            {purchases.map(p => (
              <tr key={p.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{p.poNo}</td>
                <td className="mono">{p.date}</td><td>{p.supplierName}</td>
                <td className="r mono">{fmtMT(p.quantityMT)}</td><td className="r mono">₹{p.ratePerMT}</td>
                <td className="r mono">₹{p.taxableAmount.toLocaleString()}</td>
                <td className="r mono" style={{ fontWeight:700, color:"var(--amber)" }}>₹{p.totalAmount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="Record Raw Material Purchase" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Purchase</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Supplier"><select value={f.supplierId} onChange={e => set("supplierId")(e.target.value)}><option value="">Select...</option>{suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG></div>
          <div className="form-row cols-3"><FG label="Quantity (MT)"><input type="number" step="0.001" value={f.quantityMT} onChange={e => set("quantityMT")(e.target.value)} /></FG><FG label="Rate / MT (₹)"><input type="number" value={f.ratePerMT} onChange={e => set("ratePerMT")(e.target.value)} /></FG><FG label="GST %"><select value={f.gstPct} onChange={e => set("gstPct")(e.target.value)}>{[0,5,12,18].map(r => <option key={r} value={r}>{r}%</option>)}</select></FG></div>
          <div className="form-row"><FG label="Invoice #"><input value={f.invoiceNo} onChange={e => set("invoiceNo")(e.target.value)} /></FG></div>
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
        <h3>Monthly Operations Quick Entry</h3>
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

function ExpensesView({ expenses, setExpenses }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), category:"fuel", amount:"", description:"", paidTo:"", reference:"", paymentMode:"bank" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.amount || !f.description) return alert("Amount and description required.");
    try {
      const row = await db.saveExpense(f);
      setExpenses(es => [row, ...es]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Expense Entry</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Exp #</th><th>Date</th><th>Category</th><th>Description</th><th>Paid To</th><th>Channel</th><th className="r">Amount</th></tr></thead>
          <tbody>
            {expenses.map(e => (
              <tr key={e.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{e.expNo}</td>
                <td className="mono">{e.date}</td><td><BadgeComponent type="muted">{e.category}</BadgeComponent></td>
                <td>{e.description}</td><td>{e.paidTo || "—"}</td>
                <td><BadgeComponent type={e.paymentMode === "cash" ? "amber" : "muted"}>{e.paymentMode}</BadgeComponent></td>
                <td className="r mono" style={{ color:"var(--red)", fontWeight: 700 }}>₹{e.amount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="Record Direct Expense" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Expense</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Category"><select value={f.category} onChange={e => set("category")(e.target.value)}>{EXP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Amount (₹)"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG><FG label="Channel"><select value={f.paymentMode} onChange={e => set("paymentMode")(e.target.value)}><option value="bank">Bank / UPI</option><option value="cash">Site Petty Cash</option></select></FG></div>
          <div className="form-row"><FG label="Description"><input value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
          <div className="form-row cols-2"><FG label="Paid To"><input value={f.paidTo} onChange={e => set("paidTo")(e.target.value)} /></FG><FG label="Bill / Voucher Ref"><input value={f.reference} onChange={e => set("reference")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function CapitalRegisterView({ capitalItems, setCapitalItems }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), category:"land", description:"", amount:"", paidTo:"", reference:"", fundedBy:"own", paymentMode:"cash", paidByPartner:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.amount || !f.description) return alert("Amount and description required.");
    try {
      const row = await db.saveCapitalItem(f);
      setCapitalItems(cs => [row, ...cs]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Fixed Asset (Land, Cash Capex, Plant)</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Cap #</th><th>Date</th><th>Category</th><th>Description</th><th>Channel</th><th>Funded By</th><th className="r">Amount</th></tr></thead>
          <tbody>
            {capitalItems.map(c => (
              <tr key={c.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{c.capNo}</td>
                <td className="mono">{c.date}</td>
                <td><BadgeComponent type="muted">{c.category}</BadgeComponent></td>
                <td style={{ fontWeight: 500 }}>{c.description}</td>
                <td><BadgeComponent type={c.paymentMode === "cash" ? "amber" : "muted"}>{c.paymentMode}</BadgeComponent></td>
                <td><BadgeComponent type={c.fundedBy === "loan" ? "amber" : "green"}>{c.fundedBy}</BadgeComponent></td>
                <td className="r mono" style={{ color:"var(--accent)", fontWeight: 700 }}>₹{c.amount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {open && (
        <Modal title="Record Capital / Land / Plant Asset" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Asset</button></>}>
          <div className="form-row cols-2"><FG label="Acquisition Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Asset Category"><select value={f.category} onChange={e => set("category")(e.target.value)}>{CAP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Total Cost / Value (₹)"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG><FG label="Payment Method"><select value={f.paymentMode} onChange={e => set("paymentMode")(e.target.value)}><option value="cash">Direct Cash / Off-Book</option><option value="bank">Company Bank Account</option><option value="partner_personal">Partner Personal Account</option></select></FG></div>
          <div className="form-row cols-2"><FG label="Financing Source"><select value={f.fundedBy} onChange={e => set("fundedBy")(e.target.value)}><option value="own">Partner Equity</option><option value="loan">Bank / Equipment Loan</option></select></FG><FG label="Paid By Partner (Optional)"><input placeholder="Partner Name" value={f.paidByPartner} onChange={e => set("paidByPartner")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Asset Description"><input placeholder="e.g. Land Plot Registry (Peddapur) or Machinery Advance" value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
          <div className="form-row cols-2"><FG label="Paid To"><input value={f.paidTo} onChange={e => set("paidTo")(e.target.value)} /></FG><FG label="Registry / Deed / Invoice Ref"><input value={f.reference} onChange={e => set("reference")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

// ── 6. BANK STATEMENT GROUPING VIEW (RESTORED CLEAN CARDS) ───────────────────
// Keeps local input state while typing so the card doesn't re-key and drop focus
function CounterpartyNameInput({ rawKeys, currentLabel, placeholder, onSave }) {
  const [val, setVal] = useState(currentLabel || "");

  // Sync if parent updates from outside
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
      e.target.blur(); // Triggers commit via onBlur
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

function BankStatementGroupingView({ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, setCapitalItems, setExpenses, setPartnerCashbook }) {
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
    } catch (err) { alert(err.message); }
  }

  async function promoteToCapital(t) {
    try {
      const item = await db.promoteTxnToCapital(t, "machinery", "own");
      setCapitalItems(cs => [item, ...cs]);
      alert(`Promoted ${fmt(t.debit)} to Fixed Assets!`);
    } catch (err) { alert(err.message); }
  }

  async function promoteToExpense(t) {
    try {
      const exp = await db.promoteTxnToExpense(t, "misc");
      setExpenses(es => [exp, ...es]);
      alert(`Promoted ${fmt(t.debit)} to Expenses!`);
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
                    <FG label="Type">
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
                                {t.debit > 0 && (
                                  <>
                                    <button className="btn btn-ghost btn-sm" title="Promote to Fixed Assets" onClick={() => promoteToCapital(t)}>+ Cap</button>
                                    <button className="btn btn-ghost btn-sm" title="Promote to Expenses" onClick={() => promoteToExpense(t)}>+ Exp</button>
                                    {g.type === "owner" && (
                                      <button className="btn btn-danger btn-sm" title="Log Drawing" onClick={() => logCashbook(t, g.label || g.key, "Drawings", "current")}>+ Draw</button>
                                    )}
                                  </>
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

function PartnersCapitalDashboard({ bankTxns, bankLabels, partnerCashbook, setPartnerCashbook, setExpenses, setCapitalItems }) {
  const [activeTab, setActiveTab] = useState("overview"); // 'overview' | 'ledger'
  const [selectedPartner, setSelectedPartner] = useState("All");
  const [openCashModal, setOpenCashModal] = useState(false);
  const [openActionModal, setOpenActionModal] = useState(null); // { partner, type }

  // 1. Recreate the EXACT same merged groups calculation engine as Bank Statement cards
  const bankPartnerGroups = useMemo(() => {
    const rawMap = {};
    (bankTxns || []).forEach(t => {
      if (!t) return;
      // Key can be t.key, t.group_key, or normalized description fallback
      const k = t.key || t.group_key || (t.description ? normalizeDesc(t.description) : "UNKNOWN");
      if (!rawMap[k]) rawMap[k] = { key: k, txns: [], totalDebit: 0, totalCredit: 0 };
      rawMap[k].txns.push(t);
      rawMap[k].totalDebit += (+t.debit || 0);
      rawMap[k].totalCredit += (+t.credit || 0);
    });

    const merged = {};
    Object.values(rawMap).forEach(g => {
      const meta = (bankLabels || {})[g.key] || {};
      const type = meta.type || "unlabeled";
      const label = (meta.label || "").trim();

      // Filter groups classified as Owner / Capital
      if (type === "owner" && label) {
        const mergeKey = label.toLowerCase();
        if (!merged[mergeKey]) {
          merged[mergeKey] = {
            displayName: label,
            totalDebit: 0,
            totalCredit: 0,
            txns: []
          };
        }
        merged[mergeKey].totalDebit += g.totalDebit;
        merged[mergeKey].totalCredit += g.totalCredit;
        merged[mergeKey].txns.push(...g.txns);
      }
    });

    return merged;
  }, [bankTxns, bankLabels]);

  // 2. Discover all partner names (from Owner cards and manual cashbook entries)
  const partners = useMemo(() => {
    const s = new Set();
    Object.values(bankPartnerGroups).forEach(g => s.add(g.displayName));
    (partnerCashbook || []).forEach(e => {
      if (e && e.partnerName) s.add(e.partnerName.trim());
    });
    return Array.from(s);
  }, [bankPartnerGroups, partnerCashbook]);

  // 3. Roll up summary per partner combining bank statement movements & manual entries
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

      // Total Capital = Bank Inward Deposits + Manual Off-Book Infusions
      const capitalInjected = bankGroup.totalCredit + manualInfusions;

      // Total Drawings = Bank Withdrawals by Partner + Manual Drawings
      const drawingsTaken = bankGroup.totalDebit + manualDrawings;

      // Net Standing = (Capital + Site Expenses) - Drawings
      const netStanding = (capitalInjected + cashSpentOnSite) - drawingsTaken;

      return {
        name,
        capitalInjected,
        cashSpentOnSite,
        drawingsTaken,
        netStanding,
      };
    }).sort((a, b) => b.capitalInjected - a.capitalInjected);
  }, [partners, bankPartnerGroups, partnerCashbook]);

  const totalCapitalAll = partnerSummaries.reduce((s, p) => s + p.capitalInjected, 0);
  const totalCashSpentAll = partnerSummaries.reduce((s, p) => s + p.cashSpentOnSite, 0);
  const totalDrawingsAll = partnerSummaries.reduce((s, p) => s + p.drawingsTaken, 0);

  // Form state for direct site cash spend
  const [cf, setCf] = useState({
    partnerName: "",
    date: today(),
    amount: "",
    isCapex: false,
    category: "misc",
    description: "",
    ref: "",
  });

  useEffect(() => {
    if (!cf.partnerName && partners.length > 0) {
      setCf(prev => ({ ...prev, partnerName: partners[0] }));
    }
  }, [partners, cf.partnerName]);

  async function handleDirectCashSubmit() {
    if (!cf.partnerName || !cf.amount || !cf.description) {
      return alert("Partner name, amount, and description are required.");
    }
    try {
      const res = await db.recordPartnerDirectCashExpense(
        cf.partnerName, cf.date, cf.category, cf.amount, cf.description, cf.ref, cf.isCapex
      );
      setPartnerCashbook(prev => [res.cashbookEntry, ...(prev || [])]);
      if (res.isCapex) setCapitalItems(prev => [res.item, ...(prev || [])]);
      else setExpenses(prev => [res.item, ...(prev || [])]);
      setOpenCashModal(false);
      setCf({ partnerName: partners[0] || "", date: today(), amount: "", isCapex: false, category: "misc", description: "", ref: "" });
      alert(`Logged ${fmt(cf.amount)} paid by ${cf.partnerName}!`);
    } catch (err) { alert(err.message); }
  }

  // Quick Action Modal state (Direct Infusion / Drawing)
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

  // Transaction ledger rows: combines individual bank transactions for owner groups with manual cashbook records
  const ledgerRows = useMemo(() => {
    const list = [];

    // Bank transactions linked to owner groups
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

    // Off-book manual cashbook entries
    (partnerCashbook || []).forEach(e => {
      if (!e || e.bankTxnId) return; // avoid double counting if linked to a bank txn
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
          <button className={`pill-tab ${activeTab === "overview" ? "active" : ""}`} onClick={() => setActiveTab("overview")}>
            Partner Equity & Summaries
          </button>
          <button className={`pill-tab ${activeTab === "ledger" ? "active" : ""}`} onClick={() => setActiveTab("ledger")}>
            Transaction Ledger ({ledgerRows.length})
          </button>
        </div>

        <button className="btn btn-primary" onClick={() => setOpenCashModal(true)}>
          + Record Cash Spent by Partner
        </button>
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
                    <EmptyState 
                      icon="🤝" 
                      message="No partners detected yet" 
                      sub="Tag counterparties as 'Owner / Capital' in Bank Statement, or click '+ Record Cash Spent by Partner' above." 
                    />
                  </td>
                </tr>
              ) : (
                partnerSummaries.map(p => (
                  <tr key={p.name}>
                    <td style={{ fontWeight: 600 }}>{p.name}</td>
                    <td className="r mono" style={{ color: "var(--accent)", fontWeight: 700 }}>{fmt(p.capitalInjected)}</td>
                    <td className="r mono" style={{ color: "var(--green)" }}>{fmt(p.cashSpentOnSite)}</td>
                    <td className="r mono" style={{ color: "var(--red)" }}>{fmt(p.drawingsTaken)}</td>
                    <td className="r mono" style={{ fontWeight: 700, fontSize: 13, color: p.netStanding >= 0 ? "var(--teal)" : "var(--amber)" }}>
                      {fmt(p.netStanding)}
                    </td>
                    <td className="r">
                      <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                        <button 
                          className="btn btn-ghost btn-sm" 
                          onClick={() => setOpenActionModal({ partner: p.name, type: "Capital Infusion" })}
                        >
                          + Capital
                        </button>
                        <button 
                          className="btn btn-danger btn-sm" 
                          onClick={() => setOpenActionModal({ partner: p.name, type: "Drawings" })}
                        >
                          + Draw
                        </button>
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
                  <tr>
                    <td colSpan={7} style={{ textAlign: "center", color: "var(--text3)", padding: 32 }}>
                      <EmptyState icon="📖" message="No ledger transactions found" />
                    </td>
                  </tr>
                ) : (
                  ledgerRows.map(e => (
                    <tr key={e.id}>
                      <td className="mono">{e.entryDate}</td>
                      <td style={{ fontWeight: 600 }}>{e.partnerName}</td>
                      <td><BadgeComponent type={e.accountType === "capital" ? "accent" : "muted"}>{e.accountType}</BadgeComponent></td>
                      <td><BadgeComponent type={e.type === "Drawings" ? "red" : "green"}>{e.type}</BadgeComponent></td>
                      <td style={{ fontSize: 12 }}>{e.remarks}</td>
                      <td><BadgeComponent type="muted">{e.paymentMode}</BadgeComponent></td>
                      <td className="r mono" style={{ fontWeight: 700, color: e.type === "Drawings" ? "var(--red)" : "var(--green)" }}>
                        {e.type === "Drawings" ? "-" : "+"}{fmt(e.amount)}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* MODAL 1: Direct Cash Spent on Site by Partner */}
      {openCashModal && (
        <Modal 
          title="Record Site Cash Expense Paid by Partner" 
          onClose={() => setOpenCashModal(false)} 
          foot={<><button className="btn btn-ghost" onClick={() => setOpenCashModal(false)}>Cancel</button><button className="btn btn-primary" onClick={handleDirectCashSubmit}>Save Entry</button></>}
        >
          <div className="form-row cols-2">
            <FG label="Partner Name *">
              {partners.length > 0 ? (
                <input 
                  list="partner-suggestions" 
                  value={cf.partnerName} 
                  placeholder="Select or type partner name"
                  onChange={e => setCf({ ...cf, partnerName: e.target.value })} 
                />
              ) : (
                <input 
                  placeholder="Enter partner name" 
                  value={cf.partnerName} 
                  onChange={e => setCf({ ...cf, partnerName: e.target.value })} 
                />
              )}
              <datalist id="partner-suggestions">
                {partners.map(p => <option key={p} value={p} />)}
              </datalist>
            </FG>
            <FG label="Date *"><input type="date" value={cf.date} onChange={e => setCf({ ...cf, date: e.target.value })} /></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Amount Paid (₹) *"><input type="number" value={cf.amount} onChange={e => setCf({ ...cf, amount: e.target.value })} /></FG>
            <FG label="Expense Nature">
              <select value={cf.isCapex ? "capex" : "opex"} onChange={e => setCf({ ...cf, isCapex: e.target.value === "capex" })}>
                <option value="opex">Operational Overhead (Site Fuel, Spares, Cash Labor)</option>
                <option value="capex">Fixed Asset / Land (Civil Work, Machinery Advance)</option>
              </select>
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
          <div className="form-row">
            <FG label="Description *"><input placeholder="e.g. Paid cash for JCB land leveling or motor rewinding" value={cf.description} onChange={e => setCf({ ...cf, description: e.target.value })} /></FG>
          </div>
        </Modal>
      )}

      {/* MODAL 2: Quick Capital Infusion or Drawing */}
      {openActionModal && (
        <Modal 
          title={`${openActionModal.type}: ${openActionModal.partner}`} 
          onClose={() => setOpenActionModal(null)} 
          foot={<><button className="btn btn-ghost" onClick={() => setOpenActionModal(null)}>Cancel</button><button className="btn btn-primary" onClick={handleQuickActionSubmit}>Record</button></>}
        >
          <div className="form-row cols-2">
            <FG label="Date"><input type="date" value={quickForm.date} onChange={e => setQuickForm({ ...quickForm, date: e.target.value })} /></FG>
            <FG label="Amount (₹) *"><input type="number" value={quickForm.amount} onChange={e => setQuickForm({ ...quickForm, amount: e.target.value })} /></FG>
          </div>
          <div className="form-row">
            <FG label="Remarks / Purpose"><input placeholder="e.g. Cash drawings or capital injection" value={quickForm.remarks} onChange={e => setQuickForm({ ...quickForm, remarks: e.target.value })} /></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

  
// ── 8. BALANCE SHEET VIEW ─────────────────────────────────────────────────────
function BalanceSheetView({ capitalItems, bankTxns, partnerCashbook, purchases, expenses, weighments, grades, boulderReceipts, productionEntries, costConfig }) {
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

  const stock = useMemo(() => computeStock(grades, boulderReceipts, productionEntries, weighments), [grades, boulderReceipts, productionEntries, weighments]);
  const totalRMReceivedMT = (boulderReceipts || []).reduce((s, r) => s + (r.quantityMT || 0), 0);
  const totalRMCost = (purchases || []).reduce((s, p) => s + (p.taxableAmount || 0), 0);
  const avgBoulderRate = totalRMReceivedMT > 0 ? (totalRMCost / totalRMReceivedMT) : 1200;

  const boulderClosingMT = Object.values(stock).find(s => s.isBoulder)?.closing || 0;
  const rawMaterialValuation = Math.max(0, boulderClosingMT * avgBoulderRate);

  const finishedGoodsValuation = useMemo(() => {
    const cptVal = +costConfig.stdSellingRate || 1800;
    return Object.values(stock)
      .filter(s => !s.isBoulder)
      .reduce((s, g) => s + Math.max(0, (g.closing || 0) * cptVal), 0);
  }, [stock, costConfig]);

  const totalCurrentAssets = (bankBalance > 0 ? bankBalance : 0) + rawMaterialValuation + finishedGoodsValuation;
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

  const totalRevenue = (weighments || []).filter(w => w.purpose === "Sale").reduce((s, w) => s + ((w.netWeight || 0) / 1000) * (+costConfig.stdSellingRate || 1800), 0);
  const totalSpend = (purchases || []).reduce((s, p) => s + (p.taxableAmount || 0), 0) + (expenses || []).reduce((s, e) => s + (e.amount || 0), 0);
  const retainedEarnings = totalRevenue - totalSpend;
  const totalFinanced = totalLiabilities + totalPartnerEquity + retainedEarnings;

  return (
    <div>
      <div className="stats-grid">
        <StatCard label="Gross Fixed Assets" value={fmt(totalFixedAssets)} color="teal" sub="Land, Civil, Plant" />
        <StatCard label="Current Assets" value={fmt(totalCurrentAssets)} color="blue" sub="Yard Stock & Cash" />
        <StatCard label="Total Liabilities" value={fmt(totalLiabilities)} color="red" sub="Loans & Overdraft" />
        <StatCard label="Partner Net Worth" value={fmt(totalPartnerEquity + retainedEarnings)} color="accent" sub="Capital + Retained P&L" />
      </div>

      <div className="two-col">
        <div className="report-section">
          <h3>Assets (Gross Block & Real-Time Inventory)</h3>
          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"10px 0" }}>FIXED INFRASTRUCTURE ASSETS (CAPEX)</div>
          {Object.entries(fixedAssets).map(([cat, amt]) => (
            <div key={cat} className="cost-row">
              <span className="cost-label" style={{ textTransform:"capitalize" }}>{cat}</span>
              <span className="cost-val">{fmt(amt)}</span>
            </div>
          ))}

          <div style={{ fontWeight:600, fontSize:11, color:"var(--text3)", margin:"16px 0 8px" }}>CURRENT ASSETS & YARD VALUATION</div>
          <div className="cost-row">
            <span className="cost-label">Boulder Yard Stock ({fmtMT(boulderClosingMT)})</span>
            <span className="cost-val" style={{ color:"var(--teal)" }}>{fmt(rawMaterialValuation)}</span>
          </div>
          <div className="cost-row">
            <span className="cost-label">Finished Goods Yard Stock</span>
            <span className="cost-val" style={{ color:"var(--teal)" }}>{fmt(finishedGoodsValuation)}</span>
          </div>
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

// ── 9. COST SHEET & STANDARDS ─────────────────────────────────────────────────
function CostSheetView({ purchases, expenses, productionEntries, weighments, costConfig }) {
  const ym = today().slice(0, 7);
  const rm = (purchases || []).filter(p => p.date?.slice(0, 7) === ym).reduce((s, p) => s + p.taxableAmount, 0);
  const opex = (expenses || []).filter(e => e.date?.slice(0, 7) === ym).reduce((s, e) => s + e.amount, 0);
  const prod = (productionEntries || []).filter(p => p.date?.slice(0, 7) === ym).reduce((s, p) => s + p.totalOutputMT, 0);
  const cpt = prod > 0 ? (rm + opex) / prod : 0;

  return (
    <div>
      <div className="cpt-banner">
        <div><div className="cpt-sub">Production Cost per Tonne</div><div className="cpt-val">₹{cpt.toFixed(0)}</div></div>
        <div><div className="cpt-sub">Monthly Spend ({ym})</div><div className="cpt-val" style={{ color:"var(--blue)" }}>₹{(rm + opex).toLocaleString()}</div></div>
      </div>
    </div>
  );
}

function CostConfigView({ costConfig, setCostConfig }) {
  const [f, setF] = useState(costConfig);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    try {
      await db.saveCostConfig(f);
      setCostConfig(f);
      alert("Cost configuration saved!");
    } catch (err) { alert(err.message); }
  }

  return (
    <div className="config-card">
      <h3>Cost Configuration Standards</h3>
      <div className="config-row" style={{ marginTop: 14 }}>
        <FG label="Bag Cost (₹/tonne)"><input type="number" value={f.bagCostPerTonne || ""} onChange={e => set("bagCostPerTonne")(e.target.value)} /></FG>
        <FG label="Freight Rate (₹/MT)"><input type="number" value={f.freightPerTonne || ""} onChange={e => set("freightPerTonne")(e.target.value)} /></FG>
        <FG label="Std Valuation / Selling Rate (₹/MT)"><input type="number" value={f.stdSellingRate || ""} onChange={e => set("stdSellingRate")(e.target.value)} /></FG>
      </div>
      <button className="btn btn-primary" style={{ marginTop: 14 }} onClick={save}>Save Configurations</button>
    </div>
  );
}

// ── 10. MASTER TABLE HELPER ───────────────────────────────────────────────────
function MasterTableView({ title, noun, items, setItems, saveFn, delFn, fields, cols }) {
  const blank = Object.fromEntries(fields.map(f => [f.key, ""]));
  const [open, setOpen] = useState(false);
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    try {
      const res = await saveFn(f);
      setItems(xs => [...xs, res]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add {noun}</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr>{cols.map(c => <th key={c.label}>{c.label}</th>)}</tr></thead>
          <tbody>
            {(items || []).map(row => <tr key={row.id}>{cols.map(c => <td key={c.label}>{row[c.key]}</td>)}</tr>)}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title={`New ${noun}`} onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          {fields.map(fl => <FG key={fl.key} label={fl.label} span={fl.full ? 2 : undefined}><input value={f[fl.key] || ""} onChange={e => set(fl.key)(e.target.value)} /></FG>)}
        </Modal>
      )}
    </div>
  );
}

// ── 11. DASHBOARD OVERVIEW ────────────────────────────────────────────────────
function DashboardView({ weighments, boulderReceipts, productionEntries, expenses, capitalItems, bankTxns, partnerCashbook }) {
  const totalCapex = (capitalItems || []).reduce((s, c) => s + (c.amount || 0), 0);
  const bankBalance = (bankTxns || []).reduce((s, t) => s + ((t.credit || 0) - (t.debit || 0)), 0);
  const totalPartnerCapital = (partnerCashbook || []).filter(e => e.type === "Capital Infusion").reduce((s, e) => s + (e.amount || 0), 0);

  return (
    <div className="stats-grid-5">
      <StatCard label="Fixed Assets Gross" value={fmt(totalCapex)} color="teal" sub="Land & Plant" />
      <StatCard label="Bank Balance" value={fmt(bankBalance)} color={bankBalance >= 0 ? "green" : "red"} sub="Liquid Cash" />
      <StatCard label="Partner Capital" value={fmt(totalPartnerCapital)} color="accent" sub="Contributed" />
      <StatCard label="Boulder Inward" value={`${(boulderReceipts || []).reduce((s,r) => s + (r.quantityMT || 0), 0).toFixed(1)} MT`} color="blue" sub="Deliveries" />
      <StatCard label="Sales Dispatched" value={`${(weighments || []).filter(w => w.purpose === "Sale").reduce((s,w) => s + ((w.netWeight || 0)/1000), 0).toFixed(1)} MT`} color="purple" sub="Slips" />
    </div>
  );
}
