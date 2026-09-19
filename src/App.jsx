import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import * as db from "./db.js";

// ── Fonts + CSS ───────────────────────────────────────────────────────────────
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
.sidebar{width:220px;background:var(--bg2);border-right:1px solid var(--border);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;flex-shrink:0;overflow:hidden}
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
.output-rows{display:flex;flex-direction:column;gap:8px}
.output-row{display:grid;grid-template-columns:1fr 1fr auto;gap:8px;align-items:end}
.cpt-banner{background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r2);padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:20px}
.cpt-val{font-size:30px;font-weight:800;color:var(--accent);font-family:var(--mono)}
.cpt-sub{font-size:11px;color:var(--text3);margin-top:3px;font-family:var(--mono)}
.qc-param-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px}
.qc-param-card{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:10px 12px}
.qc-param-card label{font-size:9.5px;color:var(--text3);text-transform:uppercase;font-family:var(--mono);display:block;margin-bottom:6px}
.qc-param-card input{background:var(--bg4);border:1px solid var(--border2);border-radius:5px;padding:6px 9px;color:var(--text);font-family:var(--mono);font-size:13px;width:100%}
.lot-chip{display:inline-flex;align-items:center;gap:5px;background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r);padding:3px 10px;font-family:var(--mono);font-size:11px;color:var(--accent)}
.config-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:18px 20px;margin-bottom:16px}
.config-card h3{font-size:13px;font-weight:600;margin-bottom:14px}
.config-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
`;

// ── Helpers ───────────────────────────────────────────────────────────────────
const SHIFTS = ["Morning (6AM–2PM)","Afternoon (2PM–10PM)","Night (10PM–6AM)"];
const OPERATORS = ["Suresh Kumar","Manoj Singh","Ravi Patil","Dinesh Yadav"];
const fmtMT = v => `${(+v||0).toFixed(3)} MT`;
const fmtWt = kg => kg>=1000?`${(kg/1000).toFixed(3)} MT`:`${kg} kg`;
const fmt = n => `₹${(+n||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const today = () => new Date().toISOString().split("T")[0];
const nowTime = () => new Date().toTimeString().slice(0,5);
const addDays = (dStr, n) => { const d = new Date(dStr); d.setDate(d.getDate()+n); return d.toISOString().split("T")[0]; };
const calcNextDue = (lastDone, type, val) => {
  if (!lastDone) return today();
  const n = +val || 1; const d = new Date(lastDone);
  if (type === "daily") d.setDate(d.getDate() + n);
  if (type === "weekly") d.setDate(d.getDate() + n * 7);
  if (type === "monthly") d.setMonth(d.getMonth() + n);
  return d.toISOString().split("T")[0];
};

function computeStock(grades, receipts, entries, weighments) {
  const s = {};
  grades.forEach(g => { s[g.id] = { ...g, gradeId: g.id, in: 0, produced: 0, sold: 0, transferred: 0 }; });
  receipts.forEach(r => { const b = grades.find(g => g.isBoulder); if (b && s[b.id]) s[b.id].in += r.quantityMT; });
  entries.forEach(pe => {
    const b = grades.find(g => g.isBoulder);
    if (b && s[b.id]) s[b.id].produced -= pe.boulderConsumedMT;
    pe.outputs.forEach(o => { if (s[o.gradeId]) s[o.gradeId].produced += o.quantityMT; });
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

// ── UI Components ─────────────────────────────────────────────────────────────
function Badge({type,children}){return <span className={`badge badge-${type||"muted"}`}>{children}</span>;}
function StatCard({label,value,sub,color="blue"}){
  return <div className={`stat-card ${color}`}><div className="stat-label">{label}</div><div className={`stat-val ${color}`}>{value}</div>{sub&&<div className="stat-sub">{sub}</div>}</div>;
}
function FG({label,note,children,span}){
  return <div className="form-group" style={span?{gridColumn:`span ${span}`}:{}}>{label&&<label>{label}</label>}{children}{note&&<div className="form-note">{note}</div>}</div>;
}
function Modal({title,onClose,children,foot,size=""}){
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
function EmptyState({icon,message,sub}){
  return <div className="empty"><div className="empty-icon">{icon||"📭"}</div><p style={{fontWeight:600,marginBottom:4}}>{message}</p>{sub&&<p style={{fontSize:11.5,color:"var(--text3)",marginTop:4}}>{sub}</p>}</div>;
}

// ── PDF.JS LOADER (Vite Production Safe) ───────────────────────────────────────
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
  { id:"civil", label:"Land / Civil / Construction" },
  { id:"vehicle", label:"Vehicles" },
  { id:"license", label:"Licenses & Registration" },
  { id:"deposit", label:"Deposits & Utility Setup" },
  { id:"office", label:"Furniture & Office Setup" },
  { id:"other", label:"Other Capital Spend" },
];
const CP_TYPES = [
  { id:"unlabeled", label:"Unlabeled" }, { id:"vendor", label:"Vendor / Supplier" },
  { id:"customer", label:"Customer / Buyer" }, { id:"staff", label:"Staff / Labour" },
  { id:"owner", label:"Owner / Capital" }, { id:"loan", label:"Loan / Bank" },
  { id:"utility", label:"Utility / Rent" }, { id:"tax", label:"Tax / Government" },
  { id:"other", label:"Other" },
];
const QC_PARAMS = [
  { key:"whiteness", label:"Whiteness (L*)", unit:"L* value", placeholder:"92.5" },
  { key:"sio2", label:"SiO₂", unit:"%", placeholder:"99.2" },
  { key:"fe2o3", label:"Fe₂O₃", unit:"%", placeholder:"0.025" },
  { key:"al2o3", label:"Al₂O₃", unit:"%", placeholder:"0.12" },
  { key:"moisture", label:"Moisture", unit:"%", placeholder:"0.3" },
  { key:"loi", label:"LOI", unit:"%", placeholder:"0.4" },
  { key:"d50", label:"Particle D50", unit:"microns", placeholder:"45" },
  { key:"bulkDensity", label:"Bulk Density", unit:"g/cc", placeholder:"0.95" },
];

// ── NAVIGATION MAP ────────────────────────────────────────────────────────────
const NAV = [
  { id:"dashboard", label:"Dashboard", icon:"◈", group:null },
  { id:"weighment", label:"Weighment", icon:"⚖", group:"Operations" },
  { id:"customers", label:"Customers", icon:"◈", group:"Operations" },
  { id:"boulder", label:"Boulder In", icon:"⬡", group:"Operations" },
  { id:"production", label:"Production", icon:"⚙", group:"Operations" },
  { id:"stock", label:"Stock Ledger", icon:"▤", group:"Operations" },
  { id:"shift", label:"Shift Log", icon:"◷", group:"Operations" },
  { id:"maint", label:"Maint. Overview", icon:"⚡", group:"Maintenance" },
  { id:"tasks", label:"Task Schedule", icon:"📋", group:"Maintenance" },
  { id:"maintlog", label:"Activity Log", icon:"🔧", group:"Maintenance" },
  { id:"equipment", label:"Equipment", icon:"🔩", group:"Maintenance" },
  { id:"lots", label:"Lot Register", icon:"🏷", group:"Quality" },
  { id:"qc", label:"QC Tests", icon:"🔬", group:"Quality" },
  { id:"samples", label:"Samples", icon:"🧪", group:"Quality" },
  { id:"purchases", label:"Purchases", icon:"📥", group:"Reports" },
  { id:"opscosts", label:"Monthly Ops Costs", icon:"🧾", group:"Reports" },
  { id:"costs", label:"Cost Sheet", icon:"💰", group:"Reports" },
  { id:"costconfig", label:"Cost Config", icon:"⚙", group:"Reports" },
  { id:"expenses", label:"Expenses", icon:"💸", group:"Reports" },
  { id:"capital", label:"Capital / Infra", icon:"🏗", group:"Reports" },
  { id:"bankstatement", label:"Bank Statement", icon:"🏦", group:"Reports" },
  { id:"partners", label:"Partners & Capital", icon:"🤝", group:"Reports" },
  { id:"partnercashbook", label:"Partner Cashbook", icon:"📖", group:"Reports" },
  { id:"rpt-daily", label:"Daily Report", icon:"📅", group:"Reports" },
  { id:"rpt-stock", label:"Stock Report", icon:"📦", group:"Reports" },
  { id:"vehicles", label:"Vehicles", icon:"◉", group:"Masters" },
  { id:"suppliers", label:"Suppliers", icon:"◈", group:"Masters" },
  { id:"grades", label:"Grades", icon:"◆", group:"Masters" },
];
const TITLES = Object.fromEntries(NAV.map(n => [n.id, n.label]));
const GROUPS = [...new Set(NAV.filter(n => n.group).map(n => n.group))];

// ── MAIN APPLICATION SHELL ───────────────────────────────────────────────────
export default function App() {
  const [view, setView] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [loading, setLoading] = useState(true);

  // Entities
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
        setVehicles(d.vehicles); setCustomers(d.customers); setGrades(d.grades);
        setSuppliers(d.suppliers); setEquipment(d.equipment); setMaintTasks(d.maintTasks);
        setMaintLogs(d.maintLogs); setWeighments(d.weighments); setBoulderReceipts(d.boulderReceipts);
        setLots(d.lots); setProductionEntries(d.productionEntries); setShiftLogs(d.shiftLogs);
        setPurchases(d.purchases); setExpenses(d.expenses); setCapitalItems(d.capitalItems);
        setBankBatches(d.bankBatches); setBankTxns(d.bankTxns); setBankLabels(d.bankLabels);
        setCostConfig(d.costConfig); setQcTests(d.qcTests); setSamples(d.samples);
        setPartnerCashbook(d.partnerCashbook);
      } catch (err) {
        console.error("Failed to load records from Supabase:", err);
      } finally {
        setLoading(false);
      }
    }
    init();
  }, []);

  const goTo = id => { setView(id); setNavOpen(false); };

  if (loading) {
    return (
      <div style={{ display:"flex", alignItems:"center", justifyContent:"center", minHeight:"100vh", background:"var(--bg)", color:"var(--accent)", fontFamily:"monospace" }}>
        Connecting to Cloud Database...
      </div>
    );
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
          <div className="sidebar-foot">BallMill ERP · Cloud Sync</div>
        </aside>
        <div className="main">
          <div className="topbar">
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              <button className="hamburger-btn" onClick={() => setNavOpen(true)}>☰</button>
              <h2>{TITLES[view] || view}</h2>
            </div>
            <div className="topbar-right">
              <span className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{today()}</span>
              <Badge type="accent">Live</Badge>
            </div>
          </div>
          <div className="content">
            {view === "dashboard" && <Dashboard {...{ vehicles, customers, grades, weighments, boulderReceipts, productionEntries, expenses, lots, qcTests, samples, maintTasks, maintLogs }} />}
            {view === "weighment" && <WeighmentEntry {...{ vehicles, customers, grades, weighments, setWeighments }} />}
            {view === "boulder" && <BoulderReceipt {...{ suppliers, vehicles, boulderReceipts, setBoulderReceipts }} />}
            {view === "production" && <ProductionEntry {...{ grades, productionEntries, setProductionEntries, lots, setLots }} />}
            {view === "stock" && <StockLedger {...{ grades, boulderReceipts, productionEntries, weighments }} />}
            {view === "shift" && <ShiftLog {...{ grades, shiftLogs, setShiftLogs }} />}
            {view === "lots" && <LotRegister {...{ lots, setLots, qcTests, grades }} />}
            {view === "qc" && <QCTests {...{ lots, qcTests, setQcTests }} />}
            {view === "samples" && <SampleRegister {...{ lots, samples, setSamples }} />}
            {view === "costconfig" && <CostConfig {...{ costConfig, setCostConfig }} />}
            {view === "costs" && <MonthlyCostSheet {...{ purchases, expenses, productionEntries, weighments, grades, lots, costConfig }} />}
            {view === "maint" && <MaintDashboard {...{ maintTasks, maintLogs, equipment }} />}
            {view === "tasks" && <TaskSchedule {...{ maintTasks, setMaintTasks, equipment, maintLogs, setMaintLogs }} />}
            {view === "maintlog" && <MaintLog {...{ maintLogs, setMaintLogs, equipment }} />}
            {view === "equipment" && <EquipmentRegister {...{ equipment, setEquipment, maintLogs }} />}
            {view === "purchases" && <PurchaseRegister {...{ suppliers, purchases, setPurchases }} />}
            {view === "opscosts" && <MonthlyOpsCosts {...{ expenses, setExpenses }} />}
            {view === "expenses" && <ExpensesView {...{ expenses, setExpenses }} />}
            {view === "capital" && <CapitalRegister {...{ capitalItems, setCapitalItems }} />}
            {view === "bankstatement" && <BankStatementGrouping {...{ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, setCapitalItems, setExpenses, setPartnerCashbook }} />}
            {view === "partners" && <PartnersCapitalDashboard {...{ bankTxns, bankLabels }} />}
            {view === "partnercashbook" && <PartnerCashbookView {...{ partnerCashbook, bankLabels }} />}
            {view === "rpt-daily" && <DailyReport {...{ weighments, productionEntries, boulderReceipts, expenses, grades }} />}
            {view === "rpt-stock" && <StockReport {...{ grades, boulderReceipts, productionEntries, weighments }} />}
            {view === "vehicles" && <MasterPage title="Vehicles" noun="Vehicle" icon="🚛" items={vehicles} setItems={setVehicles} saveFn={db.saveVehicle} delFn={db.deleteVehicle} fields={[{ key:"vehicleNo", label:"Vehicle Number", required:true },{ key:"type", label:"Type", type:"select", options:["Tipper","Dumper","Truck","Other"] },{ key:"tareWeight", label:"Tare Weight (kg)", type:"number", required:true },{ key:"owner", label:"Owner" },{ key:"phone", label:"Phone" }]} cols={[{ label:"Vehicle #", key:"vehicleNo" },{ label:"Type", key:"type" },{ label:"Tare Wt", render:r=>fmtWt(r.tareWeight), r:true },{ label:"Owner", key:"owner" }]} />}
            {view === "customers" && <MasterPage title="Customers" noun="Customer" icon="🏢" items={customers} setItems={setCustomers} saveFn={db.saveCustomer} delFn={db.deleteCustomer} fields={[{ key:"name", label:"Company Name", required:true },{ key:"gstin", label:"GSTIN" },{ key:"contact", label:"Contact Person" },{ key:"phone", label:"Phone" },{ key:"state", label:"State" },{ key:"address", label:"Address", full:true }]} cols={[{ label:"Name", key:"name" },{ label:"GSTIN", key:"gstin" },{ label:"State", key:"state" },{ label:"Phone", key:"phone" }]} />}
            {view === "suppliers" && <MasterPage title="Suppliers" noun="Supplier" icon="⛏" items={suppliers} setItems={setSuppliers} saveFn={db.saveSupplier} delFn={db.deleteSupplier} fields={[{ key:"name", label:"Mine / Supplier Name", required:true },{ key:"contact", label:"Contact Person" },{ key:"phone", label:"Phone" },{ key:"address", label:"Region" }]} cols={[{ label:"Name", key:"name" },{ label:"Contact", key:"contact" },{ label:"Region", key:"address" }]} />}
            {view === "grades" && <MasterPage title="Material Grades" noun="Grade" icon="💎" items={grades} setItems={setGrades} saveFn={db.saveGrade} delFn={db.deleteGrade} fields={[{ key:"code", label:"Grade Code", required:true },{ key:"name", label:"Grade Name", required:true },{ key:"unit", label:"Unit", type:"select", options:["MT","KG"], default:"MT" },{ key:"isBoulder", label:"Raw Material / Boulder Grade", type:"checkbox" }]} cols={[{ label:"Code", key:"code" },{ label:"Name", key:"name" },{ label:"Type", render:r=><Badge type={r.isBoulder?"teal":"blue"}>{r.isBoulder?"Raw":"Finished"}</Badge> }]} />}
          </div>
        </div>
      </div>
    </>
  );
}

// ── VIEWS ─────────────────────────────────────────────────────────────────────

function WeighmentEntry({ vehicles, customers, grades, weighments, setWeighments }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), time:nowTime(), vehicleId:"", customerId:"", gradeId:"", lotNo:"", grossWeight:"", tareWeight:"", purpose:"Sale", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));
  const pickV = id => { const v = vehicles.find(x => String(x.id) === id); setF(x => ({ ...x, vehicleId:id, tareWeight: v ? String(v.tareWeight) : "" })); };
  const net = () => { const g = parseFloat(f.grossWeight), t = parseFloat(f.tareWeight); return (!isNaN(g) && !isNaN(t)) ? g - t : null; };

  async function save() {
    if (!f.vehicleId || !f.gradeId || !f.grossWeight || !f.tareWeight) return alert("Missing required fields");
    if (net() <= 0) return alert("Net weight must be positive");
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
            {weighments.map(w => (
              <tr key={w.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{w.slipNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{w.date} {w.time}</td>
                <td className="mono">{w.vehicleNo}</td><td>{w.customerName}</td>
                <td><Badge type="muted">{w.gradeCode}</Badge></td>
                <td className="r mono">{w.grossWeight} kg</td><td className="r mono">{w.tareWeight} kg</td>
                <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtWt(w.netWeight)}</td>
                <td><Badge type={w.purpose === "Sale" ? "green" : "blue"}>{w.purpose}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New Weighment Slip" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Slip</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Time"><input type="time" value={f.time} onChange={e => set("time")(e.target.value)} /></FG></div>
          <div className="form-row cols-2">
            <FG label="Vehicle *"><select value={f.vehicleId} onChange={e => pickV(e.target.value)}><option value="">Select...</option>{vehicles.map(v => <option key={v.id} value={v.id}>{v.vehicleNo}</option>)}</select></FG>
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
        </Modal>
      )}
    </div>
  );
}

function BoulderReceipt({ suppliers, vehicles, boulderReceipts, setBoulderReceipts }) {
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
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Receipt</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Receipt #</th><th>Date</th><th>Supplier</th><th>Vehicle</th><th>Royalty #</th><th className="r">Qty (MT)</th></tr></thead>
          <tbody>
            {boulderReceipts.map(r => (
              <tr key={r.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{r.receiptNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{r.date}</td>
                <td>{r.supplierName}</td><td className="mono">{r.vehicleNo}</td><td>{r.royaltyNo || "—"}</td>
                <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtMT(r.quantityMT)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New Boulder Receipt" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Supplier *"><select value={f.supplierId} onChange={e => set("supplierId")(e.target.value)}><option value="">Select...</option>{suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Vehicle"><select value={f.vehicleId} onChange={e => set("vehicleId")(e.target.value)}><option value="">Select...</option>{vehicles.map(v => <option key={v.id} value={v.id}>{v.vehicleNo}</option>)}</select></FG><FG label="Quantity (MT) *"><input type="number" step="0.001" value={f.quantityMT} onChange={e => set("quantityMT")(e.target.value)} /></FG></div>
          <div className="form-row cols-2"><FG label="Royalty #"><input value={f.royaltyNo} onChange={e => set("royaltyNo")(e.target.value)} /></FG><FG label="Challan #"><input value={f.challanNo} onChange={e => set("challanNo")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function ProductionEntry({ grades, productionEntries, setProductionEntries, lots, setLots }) {
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
            {productionEntries.map(pe => (
              <tr key={pe.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{pe.entryNo}</td>
                <td><span className="lot-chip">🏷 {pe.lotNo}</span></td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{pe.date}</td>
                <td><Badge type="muted">{(pe.shift || "").split(" ")[0]}</Badge></td>
                <td>{pe.operator}</td><td className="r mono">{fmtMT(pe.boulderConsumedMT)}</td>
                <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtMT(pe.totalOutputMT)}</td>
                <td>{pe.yieldPct ? <Badge type={pe.yieldPct >= 90 ? "green" : "amber"}>{pe.yieldPct}%</Badge> : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New Production Entry" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Entry</button></>}>
          <div className="form-row cols-3"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Shift"><select value={f.shift} onChange={e => set("shift")(e.target.value)}>{SHIFTS.map(s => <option key={s}>{s}</option>)}</select></FG><FG label="Operator"><select value={f.operator} onChange={e => set("operator")(e.target.value)}>{OPERATORS.map(o => <option key={o}>{o}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Boulder Consumed (MT) *"><input type="number" step="0.001" value={f.boulderConsumedMT} onChange={e => set("boulderConsumedMT")(e.target.value)} /></FG><FG label="Machine Hours"><input type="number" value={f.machineHours} onChange={e => set("machineHours")(e.target.value)} /></FG></div>
          <div style={{ marginBottom: 14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", marginBottom: 8 }}><label style={{ fontSize:11, color:"var(--text2)", textTransform:"uppercase", fontFamily:"var(--mono)" }}>Output Grades</label><button className="btn btn-ghost btn-sm" onClick={() => setF(x => ({ ...x, outputs:[...x.outputs, blankOut()] }))}>+ Add</button></div>
            <div className="output-rows">
              {f.outputs.map((row, i) => (
                <div key={i} className="output-row">
                  <div className="form-group"><select value={row.gradeId} onChange={e => setOut(i, "gradeId", e.target.value)}>{finished.map(g => <option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></div>
                  <div className="form-group"><input type="number" step="0.001" placeholder="MT" value={row.quantityMT} onChange={e => setOut(i, "quantityMT", e.target.value)} /></div>
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

function StockLedger({ grades, boulderReceipts, productionEntries, weighments }) {
  const stock = useMemo(() => computeStock(grades, boulderReceipts, productionEntries, weighments), [grades, boulderReceipts, productionEntries, weighments]);
  return (
    <div className="table-wrap">
      <div className="table-toolbar"><h3>Grade-wise Stock Inventory Position</h3></div>
      <table>
        <thead><tr><th>Grade</th><th>Classification</th><th className="r">Inflow / Produced</th><th className="r">Outflow / Sold</th><th className="r">Closing Bal</th><th>Status</th></tr></thead>
        <tbody>
          {Object.values(stock).map(s => {
            const st = s.closing < 0 ? "red" : s.closing < 10 ? "amber" : "green";
            return (
              <tr key={s.gradeId}>
                <td><Badge type={s.isBoulder ? "teal" : "muted"}>{s.code}</Badge> {s.name}</td>
                <td><Badge type={s.isBoulder ? "teal" : "blue"}>{s.isBoulder ? "Raw Material" : "Finished"}</Badge></td>
                <td className="r mono" style={{ color:"var(--teal)" }}>{fmtMT(s.isBoulder ? s.in : s.produced)}</td>
                <td className="r mono" style={{ color:"var(--amber)" }}>{fmtMT(s.isBoulder ? Math.abs(s.produced) : s.sold)}</td>
                <td className="r mono" style={{ fontWeight:700, color: st === "red" ? "var(--red)" : st === "amber" ? "var(--amber)" : "var(--accent)" }}>{fmtMT(s.closing)}</td>
                <td><Badge type={st}>{st === "red" ? "Negative" : st === "amber" ? "Low" : "Optimal"}</Badge></td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function ShiftLog({ grades, shiftLogs, setShiftLogs }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), shift:SHIFTS[0], operator:OPERATORS[0], machineStatus:"Running", boulderFed:"", production:[{ gradeId:grades[0]?.id||"", quantity:"" }], breakdowns:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));
  const setPr = (i, k, v) => setF(x => { const p = [...x.production]; p[i] = { ...p[i], [k]: v }; return { ...x, production: p }; });

  async function save() {
    if (!f.boulderFed) return alert("Boulder fed required.");
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
          <thead><tr><th>Log #</th><th>Date</th><th>Shift</th><th>Operator</th><th>Status</th><th className="r">Fed (MT)</th></tr></thead>
          <tbody>
            {shiftLogs.map(l => (
              <tr key={l.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{l.logNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{l.date}</td>
                <td><Badge type="muted">{(l.shift || "").split(" ")[0]}</Badge></td>
                <td>{l.operator}</td><td><Badge type={l.machineStatus === "Running" ? "green" : "red"}>{l.machineStatus}</Badge></td>
                <td className="r mono">{l.boulderFed}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New Shift Log" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Shift"><select value={f.shift} onChange={e => set("shift")(e.target.value)}>{SHIFTS.map(s => <option key={s}>{s}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Operator"><select value={f.operator} onChange={e => set("operator")(e.target.value)}>{OPERATORS.map(o => <option key={o}>{o}</option>)}</select></FG><FG label="Boulder Fed (MT) *"><input type="number" value={f.boulderFed} onChange={e => set("boulderFed")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Breakdowns"><input value={f.breakdowns} onChange={e => set("breakdowns")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function LotRegister({ lots, setLots, qcTests, grades }) {
  async function updateStatus(id, status) {
    try {
      await db.updateLotStatus(id, status);
      setLots(ls => ls.map(l => l.id === id ? { ...l, status } : l));
    } catch (err) { alert(err.message); }
  }
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th>Lot #</th><th>Date</th><th>Shift</th><th className="r">Total Output</th><th>QC Status</th><th>Action</th></tr></thead>
        <tbody>
          {lots.map(l => (
            <tr key={l.id}>
              <td><span className="lot-chip">🏷 {l.lotNo}</span></td>
              <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{l.date}</td>
              <td><Badge type="muted">{(l.shift || "").split(" ")[0]}</Badge></td>
              <td className="r mono" style={{ color:"var(--teal)", fontWeight:700 }}>{fmtMT(l.totalOutputMT)}</td>
              <td><Badge type={l.status}>{l.status}</Badge></td>
              <td>
                <select value={l.status} onChange={e => updateStatus(l.id, e.target.value)} style={{ background:"var(--bg3)", border:"1px solid var(--border2)", borderRadius:"var(--r)", padding:"4px 8px", color:"var(--text)", fontSize:11 }}>
                  {["in-stock","under-test","released","dispatched","rejected"].map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function QCTests({ lots, qcTests, setQcTests }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), lotNo:"", testedBy:"", labType:"In-house", certNo:"", result:"pass", remarks:"" };
  const blankParams = Object.fromEntries(QC_PARAMS.map(p => [p.key, ""]));
  const [f, setF] = useState(blank);
  const [params, setParams] = useState(blankParams);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.lotNo) return alert("Select a Lot.");
    try {
      const row = await db.saveQcTest({ ...f, params });
      setQcTests(ts => [row, ...ts]);
      setOpen(false); setF(blank); setParams(blankParams);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New QC Test</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Test #</th><th>Date</th><th>Lot #</th><th>Lab</th><th>Whiteness L*</th><th>Fe₂O₃ %</th><th>SiO₂ %</th><th>Result</th></tr></thead>
          <tbody>
            {qcTests.map(t => (
              <tr key={t.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{t.testNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{t.date}</td>
                <td><span className="lot-chip">🏷 {t.lotNo}</span></td>
                <td><Badge type="muted">{t.labType}</Badge></td>
                <td className="mono" style={{ fontWeight:700 }}>{t.params?.whiteness || "—"}</td>
                <td className="mono">{t.params?.fe2o3 || "—"}</td>
                <td className="mono">{t.params?.sio2 || "—"}</td>
                <td><Badge type={t.result}>{t.result}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New QC Test Entry" onClose={() => setOpen(false)} size="modal-lg" foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Test</button></>}>
          <div className="form-row cols-3">
            <FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG>
            <FG label="Lot # *"><select value={f.lotNo} onChange={e => set("lotNo")(e.target.value)}><option value="">Select...</option>{lots.map(l => <option key={l.id} value={l.lotNo}>{l.lotNo}</option>)}</select></FG>
            <FG label="Result"><select value={f.result} onChange={e => set("result")(e.target.value)}>{["pass","fail","conditional"].map(r => <option key={r}>{r}</option>)}</select></FG>
          </div>
          <div className="qc-param-grid">
            {QC_PARAMS.map(p => (
              <div key={p.key} className="qc-param-card">
                <label>{p.label}</label>
                <input type="number" step="0.001" placeholder={p.placeholder} value={params[p.key] || ""} onChange={e => setParams(x => ({ ...x, [p.key]: e.target.value }))} />
              </div>
            ))}
          </div>
        </Modal>
      )}
    </div>
  );
}

function SampleRegister({ lots, samples, setSamples }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), lotNo:"", location:"", collectedBy:"", quantity:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.lotNo || !f.location) return alert("Lot and Storage Location required.");
    try {
      const row = await db.saveSample({ ...f, disposalDate: addDays(f.date, 30) });
      setSamples(ss => [row, ...ss]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  async function markDisposed(id) {
    try {
      await db.markSampleDisposed(id);
      setSamples(ss => ss.map(s => s.id === id ? { ...s, status:"disposed", disposedDate:today() } : s));
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Sample</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Sample #</th><th>Lot #</th><th>Date</th><th>Location</th><th>Disposal Due</th><th>Status</th><th>Action</th></tr></thead>
          <tbody>
            {samples.map(s => (
              <tr key={s.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{s.sampleNo}</td>
                <td><span className="lot-chip">🏷 {s.lotNo}</span></td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{s.date}</td>
                <td>{s.location}</td><td className="mono">{s.disposalDate}</td>
                <td><Badge type={s.status}>{s.status}</Badge></td>
                <td>{s.status === "retained" && <button className="btn btn-danger btn-sm" onClick={() => markDisposed(s.id)}>Dispose</button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="Register New Sample" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Sample</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Lot # *"><select value={f.lotNo} onChange={e => set("lotNo")(e.target.value)}><option value="">Select...</option>{lots.map(l => <option key={l.id} value={l.lotNo}>{l.lotNo}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Storage Location *"><input value={f.location} onChange={e => set("location")(e.target.value)} /></FG><FG label="Collector"><input value={f.collectedBy} onChange={e => set("collectedBy")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function CostConfig({ costConfig, setCostConfig }) {
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
    <div>
      <div className="config-card">
        <h3>Operational Standard Rates</h3>
        <div className="config-row">
          <FG label="Bag Cost (₹/tonne)"><input type="number" value={f.bagCostPerTonne || ""} onChange={e => set("bagCostPerTonne")(e.target.value)} /></FG>
          <FG label="Freight Rate (₹/MT)"><input type="number" value={f.freightPerTonne || ""} onChange={e => set("freightPerTonne")(e.target.value)} /></FG>
          <FG label="Loading Labour (₹/MT)"><input type="number" value={f.loadingPerTonne || ""} onChange={e => set("loadingPerTonne")(e.target.value)} /></FG>
        </div>
      </div>
      <div className="config-card">
        <h3>Fixed Overheads</h3>
        <div className="config-row">
          <FG label="Monthly Testing (₹)"><input type="number" value={f.monthlyLabCost || ""} onChange={e => set("monthlyLabCost")(e.target.value)} /></FG>
          <FG label="Monthly Rent (₹)"><input type="number" value={f.monthlyRent || ""} onChange={e => set("monthlyRent")(e.target.value)} /></FG>
          <FG label="Std Selling Rate (₹/MT)"><input type="number" value={f.stdSellingRate || ""} onChange={e => set("stdSellingRate")(e.target.value)} /></FG>
        </div>
        <button className="btn btn-primary" style={{ marginTop: 14 }} onClick={save}>Save Configurations</button>
      </div>
    </div>
  );
}

function MonthlyCostSheet({ purchases, expenses, productionEntries, weighments, grades, lots, costConfig }) {
  const [ym, setYm] = useState(today().slice(0, 7));
  const inMonth = cat => expenses.filter(e => e.date?.slice(0, 7) === ym && e.category === cat).reduce((s, e) => s + e.amount, 0);
  const rmCost = purchases.filter(p => p.date?.slice(0, 7) === ym).reduce((s, p) => s + p.taxableAmount, 0);
  const salaries = inMonth("salary") + inMonth("labour");
  const power = inMonth("electric");
  const water = inMonth("water");
  const fuel = inMonth("fuel");
  const maint = inMonth("maint");
  const prodMT = productionEntries.filter(p => p.date?.slice(0, 7) === ym).reduce((s, p) => s + p.totalOutputMT, 0);
  const soldMT = weighments.filter(w => w.date?.slice(0, 7) === ym && w.purpose === "Sale").reduce((s, w) => s + (w.netWeight || 0)/1000, 0);
  const packingCost = soldMT * (+costConfig.bagCostPerTonne || 0);
  const freightCost = soldMT * (+costConfig.freightPerTonne || 0);
  const totalCost = rmCost + salaries + power + water + fuel + maint + packingCost + freightCost + (+costConfig.monthlyRent || 0);
  const cpt = prodMT > 0 ? totalCost / prodMT : 0;

  return (
    <div>
      <div className="filter-bar"><FG label="Select Month"><input type="month" value={ym} onChange={e => setYm(e.target.value)} /></FG></div>
      <div className="cpt-banner">
        <div><div className="cpt-sub">Cost per Tonne Produced</div><div className="cpt-val">₹{cpt.toLocaleString("en-IN", { maximumFractionDigits: 0 })}</div></div>
        <div><div className="cpt-sub">Total Operating Cost</div><div className="cpt-val" style={{ color:"var(--blue)" }}>₹{totalCost.toLocaleString("en-IN", { maximumFractionDigits: 0 })}</div></div>
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Cost Component</th><th className="r">Amount (₹)</th></tr></thead>
          <tbody>
            <tr><td>Raw Material Purchases</td><td className="r mono">₹{rmCost.toLocaleString()}</td></tr>
            <tr><td>Salaries & Casual Labour</td><td className="r mono">₹{salaries.toLocaleString()}</td></tr>
            <tr><td>Electricity Power Bill</td><td className="r mono">₹{power.toLocaleString()}</td></tr>
            <tr><td>Fuel & Equipment Maintenance</td><td className="r mono">₹{(fuel + maint).toLocaleString()}</td></tr>
            <tr><td>Packing & Outbound Freight</td><td className="r mono">₹{(packingCost + freightCost).toLocaleString()}</td></tr>
            <tr style={{ fontWeight:700, borderTop:"1px solid var(--border2)" }}><td>Total Production Cost</td><td className="r mono" style={{ color:"var(--accent)" }}>₹{totalCost.toLocaleString()}</td></tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function MaintDashboard({ maintTasks, maintLogs, equipment }) {
  const overdue = maintTasks.filter(t => t.active && t.nextDue < today());
  return (
    <div>
      <div className="stats-grid">
        <StatCard label="Overdue Tasks" value={overdue.length} color={overdue.length ? "red" : "green"} />
        <StatCard label="Active Tasks" value={maintTasks.filter(t => t.active).length} color="blue" />
        <StatCard label="Logs Recorded" value={maintLogs.length} color="purple" />
      </div>
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Overdue Equipment Maintenance</h3></div>
        <table>
          <thead><tr><th>Equipment</th><th>Task</th><th>Due Date</th><th>Action</th></tr></thead>
          <tbody>
            {overdue.map(t => (
              <tr key={t.id}>
                <td>{t.equipmentName}</td><td>{t.taskName}</td>
                <td className="mono" style={{ color:"var(--red)" }}>{t.nextDue}</td>
                <td><Badge type="red">Overdue</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function TaskSchedule({ maintTasks, setMaintTasks, equipment, maintLogs, setMaintLogs }) {
  const [open, setOpen] = useState(false);
  const blank = { equipmentId:"", taskName:"", category:"cleaning", intervalType:"weekly", intervalValue:"1", estimatedMins:"20", notes:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.equipmentId || !f.taskName) return alert("Required fields missing.");
    try {
      const row = await db.saveMaintTask({ ...f, nextDue: today(), active: true });
      setMaintTasks(ts => [...ts, row]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  async function logDone(task) {
    const d = today();
    const next = calcNextDue(d, task.intervalType, task.intervalValue);
    try {
      await db.updateMaintTaskDue(task.id, d, next);
      const log = await db.saveMaintLog({ equipmentId: task.equipmentId, taskId: task.id, taskName: task.taskName, date: d, type: "Preventive", durationHours: 1, cost: 0 });
      setMaintTasks(ts => ts.map(t => t.id === task.id ? { ...t, lastDone: d, nextDue: next } : t));
      setMaintLogs(ls => [log, ...ls]);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Scheduled Task</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Equipment</th><th>Task</th><th>Interval</th><th>Last Done</th><th>Next Due</th><th>Action</th></tr></thead>
          <tbody>
            {maintTasks.map(t => (
              <tr key={t.id}>
                <td>{t.equipmentName} ({t.code})</td><td>{t.taskName}</td>
                <td>{t.intervalValue} {t.intervalType}</td>
                <td className="mono">{t.lastDone || "Never"}</td><td className="mono">{t.nextDue}</td>
                <td><button className="btn btn-primary btn-sm" onClick={() => logDone(t)}>Mark Done</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="Schedule Maintenance Task" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Task</button></>}>
          <div className="form-row cols-2">
            <FG label="Equipment *"><select value={f.equipmentId} onChange={e => set("equipmentId")(e.target.value)}><option value="">Select...</option>{equipment.map(e => <option key={e.id} value={e.id}>{e.name}</option>)}</select></FG>
            <FG label="Task Name *"><input value={f.taskName} onChange={e => set("taskName")(e.target.value)} /></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Frequency"><input type="number" value={f.intervalValue} onChange={e => set("intervalValue")(e.target.value)} /></FG>
            <FG label="Unit"><select value={f.intervalType} onChange={e => set("intervalType")(e.target.value)}>{["daily","weekly","monthly"].map(u => <option key={u}>{u}</option>)}</select></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

function MaintLog({ maintLogs, equipment }) {
  return (
    <div className="table-wrap">
      <table>
        <thead><tr><th>Log #</th><th>Date</th><th>Equipment</th><th>Task</th><th>Type</th><th className="r">Duration</th><th className="r">Cost</th></tr></thead>
        <tbody>
          {maintLogs.map(l => (
            <tr key={l.id}>
              <td className="mono" style={{ color:"var(--accent)" }}>{l.logNo}</td>
              <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{l.date}</td>
              <td>{l.equipmentName}</td><td>{l.taskName}</td>
              <td><Badge type={l.type === "Breakdown" ? "red" : "green"}>{l.type}</Badge></td>
              <td className="r mono">{l.durationHours}h</td>
              <td className="r mono">{l.cost > 0 ? `₹${l.cost}` : "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function EquipmentRegister({ equipment, setEquipment, maintLogs }) {
  const [open, setOpen] = useState(false);
  const blank = { code:"", name:"", type:"Ball Mill", capacity:"", make:"", installDate:"", status:"active" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.code || !f.name) return alert("Code and Name required.");
    try {
      const row = await db.saveEquipment(f);
      setEquipment(es => [...es, row]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Equipment</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Code</th><th>Name</th><th>Type</th><th>Capacity</th><th>Status</th></tr></thead>
          <tbody>
            {equipment.map(e => (
              <tr key={e.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{e.code}</td>
                <td>{e.name}</td><td>{e.type}</td><td>{e.capacity || "—"}</td>
                <td><Badge type={e.status === "active" ? "green" : "red"}>{e.status}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="Add Equipment" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          <div className="form-row cols-2"><FG label="Code *"><input value={f.code} onChange={e => set("code")(e.target.value)} /></FG><FG label="Name *"><input value={f.name} onChange={e => set("name")(e.target.value)} /></FG></div>
          <div className="form-row cols-2"><FG label="Type"><input value={f.type} onChange={e => set("type")(e.target.value)} /></FG><FG label="Capacity"><input value={f.capacity} onChange={e => set("capacity")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function PurchaseRegister({ suppliers, purchases, setPurchases }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), supplierId:"", description:"", quantityMT:"", ratePerMT:"", gstPct:"5", invoiceNo:"", remarks:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.supplierId || !f.quantityMT || !f.ratePerMT) return alert("Required fields missing");
    try {
      const row = await db.savePurchase(f);
      setPurchases(ps => [row, ...ps]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ New Purchase</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>PO #</th><th>Date</th><th>Supplier</th><th className="r">Qty (MT)</th><th className="r">Rate</th><th className="r">Total</th></tr></thead>
          <tbody>
            {purchases.map(p => (
              <tr key={p.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{p.poNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{p.date}</td>
                <td>{p.supplierName}</td><td className="r mono">{fmtMT(p.quantityMT)}</td>
                <td className="r mono">₹{p.ratePerMT}</td><td className="r mono" style={{ fontWeight:700, color:"var(--amber)" }}>₹{p.totalAmount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="New Raw Material Purchase" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Supplier"><select value={f.supplierId} onChange={e => set("supplierId")(e.target.value)}><option value="">Select...</option>{suppliers.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}</select></FG></div>
          <div className="form-row cols-3"><FG label="Quantity (MT)"><input type="number" step="0.001" value={f.quantityMT} onChange={e => set("quantityMT")(e.target.value)} /></FG><FG label="Rate/MT"><input type="number" value={f.ratePerMT} onChange={e => set("ratePerMT")(e.target.value)} /></FG><FG label="GST %"><select value={f.gstPct} onChange={e => set("gstPct")(e.target.value)}>{[0,5,12,18].map(r => <option key={r} value={r}>{r}%</option>)}</select></FG></div>
        </Modal>
      )}
    </div>
  );
}

function MonthlyOpsCosts({ expenses, setExpenses }) {
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
      alert("Monthly costs updated successfully!");
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><FG label="Month"><input type="month" value={ym} onChange={e => setYm(e.target.value)} /></FG></div>
      <div className="config-card">
        <h3>Direct Monthly Expenditures</h3>
        <div className="config-row">
          {OPS_CATS.map(c => (
            <FG key={c} label={`${c.toUpperCase()} (₹)`}>
              <input type="number" value={vals[c] || ""} onChange={e => setVals({ ...vals, [c]: e.target.value })} />
            </FG>
          ))}
        </div>
        <button className="btn btn-primary" style={{ marginTop: 14 }} onClick={save}>Sync Monthly Totals</button>
      </div>
    </div>
  );
}

function ExpensesView({ expenses, setExpenses }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), category:"fuel", amount:"", description:"", paidTo:"", reference:"" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.amount || !f.description) return alert("Amount and description required");
    try {
      const row = await db.saveExpense(f);
      setExpenses(es => [row, ...es]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Expense</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Exp #</th><th>Date</th><th>Category</th><th>Description</th><th>Paid To</th><th className="r">Amount</th></tr></thead>
          <tbody>
            {expenses.map(e => (
              <tr key={e.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{e.expNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{e.date}</td>
                <td><Badge type="muted">{e.category}</Badge></td><td>{e.description}</td><td>{e.paidTo || "—"}</td>
                <td className="r mono" style={{ color:"var(--red)", fontWeight:700 }}>₹{e.amount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="Add Expense" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Category"><select value={f.category} onChange={e => set("category")(e.target.value)}>{EXP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Amount (₹)"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG><FG label="Paid To"><input value={f.paidTo} onChange={e => set("paidTo")(e.target.value)} /></FG></div>
          <div className="form-row"><FG label="Description"><input value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function CapitalRegister({ capitalItems, setCapitalItems }) {
  const [open, setOpen] = useState(false);
  const blank = { date:today(), category:"machinery", description:"", amount:"", paidTo:"", reference:"", fundedBy:"own" };
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    if (!f.amount || !f.description) return alert("Amount and description required");
    try {
      const row = await db.saveCapitalItem(f);
      setCapitalItems(cs => [row, ...cs]);
      setOpen(false); setF(blank);
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => setOpen(true)}>+ Add Capital Asset / Spend</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Cap #</th><th>Date</th><th>Category</th><th>Description</th><th>Funded By</th><th className="r">Amount</th></tr></thead>
          <tbody>
            {capitalItems.map(c => (
              <tr key={c.id}>
                <td className="mono" style={{ color:"var(--accent)" }}>{c.capNo}</td>
                <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{c.date}</td>
                <td><Badge type="muted">{c.category}</Badge></td><td>{c.description}</td>
                <td><Badge type={c.fundedBy === "loan" ? "amber" : "green"}>{c.fundedBy}</Badge></td>
                <td className="r mono" style={{ color:"var(--accent)", fontWeight:700 }}>₹{c.amount.toLocaleString()}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title="Record Capital / Infrastructure Investment" onClose={() => setOpen(false)} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Asset</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e => set("date")(e.target.value)} /></FG><FG label="Category"><select value={f.category} onChange={e => set("category")(e.target.value)}>{CAP_CATS.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></FG></div>
          <div className="form-row cols-2"><FG label="Amount (₹)"><input type="number" value={f.amount} onChange={e => set("amount")(e.target.value)} /></FG><FG label="Funding Source"><select value={f.fundedBy} onChange={e => set("fundedBy")(e.target.value)}><option value="own">Self Capital</option><option value="loan">Bank Loan / Debt</option></select></FG></div>
          <div className="form-row"><FG label="Description"><input value={f.description} onChange={e => set("description")(e.target.value)} /></FG></div>
        </Modal>
      )}
    </div>
  );
}

function BankStatementGrouping({ bankBatches, setBankBatches, bankTxns, setBankTxns, bankLabels, setBankLabels, setCapitalItems, setExpenses, setPartnerCashbook }) {
  const [sheet, setSheet] = useState(null);
  const [map, setMap] = useState({ dateCol:"", descCol:"", mode:"separate", debitCol:"", creditCol:"", amountCol:"" });
  const [expanded, setExpanded] = useState(null);

  async function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = "";

    if (file.name.toLowerCase().endsWith(".pdf")) {
      try {
        const parsedSheet = await parsePdfStatement(file);
        setSheet(parsedSheet);
        setMap({
          dateCol: "Date",
          descCol: "Narration",
          mode: "separate",
          debitCol: "Withdrawal Amt",
          creditCol: "Deposit Amt",
          amountCol: "",
        });
      } catch (err) {
        alert("PDF Error: " + err.message);
      }
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
      } catch (err) {
        alert("File read error: " + err.message);
      }
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
      alert(`Imported ${res.txns.length} bank transactions to Supabase!`);
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
      alert(`Promoted transaction ${t.description} to Capital / Infra!`);
    } catch (err) { alert(err.message); }
  }

  async function promoteToExpense(t) {
    try {
      const exp = await db.promoteTxnToExpense(t, "misc");
      setExpenses(es => [exp, ...es]);
      alert(`Promoted transaction ${t.description} to Expenses!`);
    } catch (err) { alert(err.message); }
  }

  async function logCashbook(t, partnerName, type) {
    try {
      const entry = await db.logPartnerCashbookEntry(t, partnerName, type);
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
    return Object.values(m).map(g => ({ ...g, label: bankLabels[g.key]?.label || "", type: bankLabels[g.key]?.type || "unlabeled" }));
  }, [bankTxns, bankLabels]);

  return (
    <div>
      <div className="config-card">
        <h3>Import Net-Banking Statement</h3>
        {!sheet ? (
          <label className="btn btn-primary" style={{ cursor:"pointer" }}>
            Choose Statement (PDF / CSV / XLSX)
            <input type="file" accept=".pdf,.csv,.xlsx" onChange={onFile} style={{ display:"none" }} />
          </label>
        ) : (
          <div>
            <div className="config-row">
              <FG label="Date Column"><select value={map.dateCol} onChange={e => setMap({ ...map, dateCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Description Column"><select value={map.descCol} onChange={e => setMap({ ...map, descCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Debit Column"><select value={map.debitCol} onChange={e => setMap({ ...map, debitCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
              <FG label="Credit Column"><select value={map.creditCol} onChange={e => setMap({ ...map, creditCol: e.target.value })}>{sheet.headers.map(h => <option key={h} value={h}>{h}</option>)}</select></FG>
            </div>
            <div style={{ display:"flex", gap:10, marginTop:14 }}>
              <button className="btn btn-ghost" onClick={() => setSheet(null)}>Cancel</button>
              <button className="btn btn-primary" onClick={importRows}>Save to Supabase</button>
            </div>
          </div>
        )}
      </div>

      <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
        {groups.map(g => (
          <div key={g.key} className="config-card" style={{ padding: 14 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap" }}>
              <div style={{ display:"flex", gap: 10, alignItems:"center" }}>
                <input placeholder={g.key} value={g.label} onChange={e => updateLabel([g.key], { label: e.target.value })} style={{ background:"var(--bg3)", border:"1px solid var(--border2)", borderRadius:"var(--r)", padding:"4px 8px", color:"var(--text)" }} />
                <select value={g.type} onChange={e => updateLabel([g.key], { type: e.target.value })} style={{ background:"var(--bg3)", border:"1px solid var(--border2)", borderRadius:"var(--r)", padding:"4px 8px", color:"var(--text)" }}>
                  {CP_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
                </select>
              </div>
              <div style={{ display:"flex", gap:14 }}>
                <span className="mono" style={{ color:"var(--red)" }}>Debit: {fmt(g.totalDebit)}</span>
                <span className="mono" style={{ color:"var(--green)" }}>Credit: {fmt(g.totalCredit)}</span>
                <button className="btn btn-ghost btn-sm" onClick={() => setExpanded(expanded === g.key ? null : g.key)}>Details ({g.txns.length})</button>
              </div>
            </div>
            {expanded === g.key && (
              <div className="table-wrap" style={{ marginTop: 10 }}>
                <table>
                  <thead><tr><th>Date</th><th>Narration</th><th className="r">Debit</th><th className="r">Credit</th><th className="r">Actions</th></tr></thead>
                  <tbody>
                    {g.txns.map(t => (
                      <tr key={t.id}>
                        <td className="mono">{t.date}</td><td>{t.description}</td>
                        <td className="r mono" style={{ color:"var(--red)" }}>{t.debit > 0 ? fmt(t.debit) : "—"}</td>
                        <td className="r mono" style={{ color:"var(--green)" }}>{t.credit > 0 ? fmt(t.credit) : "—"}</td>
                        <td className="r">
                          <div style={{ display:"flex", gap:4, justifyContent:"flex-end" }}>
                            {t.credit > 0 && g.type === "owner" && (
                              <button className="btn btn-primary btn-sm" onClick={() => logCashbook(t, g.label || g.key, "Capital Infusion")}>+ Equity</button>
                            )}
                            {t.debit > 0 && (
                              <>
                                <button className="btn btn-ghost btn-sm" title="Push to Capital" onClick={() => promoteToCapital(t)}>+ Cap</button>
                                <button className="btn btn-ghost btn-sm" title="Push to Expenses" onClick={() => promoteToExpense(t)}>+ Exp</button>
                                {g.type === "owner" && (
                                  <button className="btn btn-danger btn-sm" title="Log Drawing" onClick={() => logCashbook(t, g.label || g.key, "Drawings")}>+ Draw</button>
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
    </div>
  );
}

function PartnersCapitalDashboard({ bankTxns, bankLabels }) {
  const partners = useMemo(() => {
    const raw = {};
    bankTxns.forEach(t => {
      const type = bankLabels[t.key]?.type;
      if (type === "owner" && t.credit > 0) {
        const name = bankLabels[t.key]?.label || t.key;
        raw[name] = (raw[name] || 0) + t.credit;
      }
    });
    return Object.entries(raw).map(([name, credit]) => ({ name, credit })).sort((a,b) => b.credit - a.credit);
  }, [bankTxns, bankLabels]);

  const totalCap = partners.reduce((s, p) => s + p.credit, 0);

  return (
    <div>
      <div className="stats-grid">
        <StatCard label="Total Capital Deposited" value={fmt(totalCap)} color="accent" />
        <StatCard label="Partners Recorded" value={partners.length} color="blue" />
      </div>
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Partner Equity Ledger (All-Time)</h3></div>
        <table>
          <thead><tr><th>Partner Name</th><th className="r">Total Contributed Capital</th></tr></thead>
          <tbody>
            {partners.map(p => (
              <tr key={p.name}>
                <td style={{ fontWeight: 600 }}>{p.name}</td>
                <td className="r mono" style={{ color:"var(--accent)", fontWeight:700 }}>{fmt(p.credit)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PartnerCashbookView({ partnerCashbook, bankLabels }) {
  const [selPartner, setSelPartner] = useState("All");

  const partners = useMemo(() => {
    const set = new Set();
    Object.values(bankLabels).forEach(l => {
      if (l.type === "owner" && l.label) set.add(l.label);
    });
    partnerCashbook.forEach(e => set.add(e.partnerName));
    return ["All", ...Array.from(set)];
  }, [bankLabels, partnerCashbook]);

  const filtered = useMemo(() => {
    return selPartner === "All"
      ? partnerCashbook
      : partnerCashbook.filter(e => e.partnerName === selPartner);
  }, [partnerCashbook, selPartner]);

  const totalInfusions = filtered.filter(e => e.type === "Capital Infusion").reduce((s, e) => s + e.amount, 0);
  const totalDrawings = filtered.filter(e => e.type === "Drawings" || e.type === "Expense Reimbursement").reduce((s, e) => s + e.amount, 0);
  const netStanding = totalInfusions - totalDrawings;

  return (
    <div>
      <div className="filter-bar">
        <FG label="Partner Filter">
          <select value={selPartner} onChange={e => setSelPartner(e.target.value)}>
            {partners.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </FG>
      </div>
      <div className="stats-grid">
        <StatCard label="Total Infusions" value={fmt(totalInfusions)} color="accent" />
        <StatCard label="Total Withdrawals" value={fmt(totalDrawings)} color="red" />
        <StatCard label="Net Capital Standing" value={fmt(netStanding)} color={netStanding >= 0 ? "green" : "amber"} />
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Date</th><th>Partner</th><th>Type</th><th>Narration / Remarks</th><th className="r">Amount</th></tr></thead>
          <tbody>
            {filtered.length === 0 ? (
              <tr><td colSpan={5}><EmptyState icon="🤝" message="No cashbook entries logged yet" /></td></tr>
            ) : (
              filtered.map(e => (
                <tr key={e.id}>
                  <td className="mono" style={{ fontSize:11, color:"var(--text3)" }}>{e.entryDate}</td>
                  <td style={{ fontWeight: 600 }}>{e.partnerName}</td>
                  <td><Badge type={e.type === "Capital Infusion" ? "green" : "red"}>{e.type}</Badge></td>
                  <td>{e.remarks}</td>
                  <td className="r mono" style={{ fontWeight: 700, color: e.type === "Capital Infusion" ? "var(--green)" : "var(--red)" }}>
                    {e.type === "Capital Infusion" ? "+" : "-"}{fmt(e.amount)}
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DailyReport({ weighments, productionEntries, boulderReceipts, expenses, grades }) {
  const [date, setDate] = useState(today());
  const dayW = weighments.filter(w => w.date === date);
  const dayPE = productionEntries.filter(p => p.date === date);
  const dayBR = boulderReceipts.filter(r => r.date === date);
  const dayExp = expenses.filter(e => e.date === date);

  return (
    <div>
      <div className="filter-bar"><FG label="Select Day"><input type="date" value={date} onChange={e => setDate(e.target.value)} /></FG></div>
      <div className="stats-grid">
        <StatCard label="Boulder In" value={fmtMT(dayBR.reduce((s,r) => s + r.quantityMT, 0))} color="teal" />
        <StatCard label="Finished Produced" value={fmtMT(dayPE.reduce((s,p) => s + p.totalOutputMT, 0))} color="blue" />
        <StatCard label="Dispatched Sales" value={fmtMT(dayW.filter(w => w.purpose === "Sale").reduce((s,w) => s + (w.netWeight||0)/1000, 0))} color="accent" />
        <StatCard label="Expenses Logged" value={fmt(dayExp.reduce((s,e) => s + e.amount, 0))} color="red" />
      </div>
    </div>
  );
}

function StockReport({ grades, boulderReceipts, productionEntries, weighments }) {
  const stock = useMemo(() => computeStock(grades, boulderReceipts, productionEntries, weighments), [grades, boulderReceipts, productionEntries, weighments]);
  return (
    <div className="table-wrap">
      <div className="table-toolbar"><h3>Current Balance Inventory Summary</h3></div>
      <table>
        <thead><tr><th>Grade</th><th className="r">Available Stock</th></tr></thead>
        <tbody>
          {Object.values(stock).map(s => (
            <tr key={s.gradeId}>
              <td>{s.name}</td>
              <td className="r mono" style={{ color:"var(--accent)", fontWeight:700 }}>{fmtMT(s.closing)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function MasterPage({ title, noun, icon, fields, items, setItems, saveFn, delFn, cols }) {
  const blank = Object.fromEntries(fields.map(f => [f.key, f.default ?? ""]));
  const [open, setOpen] = useState(false);
  const [editing, setEdit] = useState(null);
  const [f, setF] = useState(blank);
  const set = k => v => setF(x => ({ ...x, [k]: v }));

  async function save() {
    const req = fields.find(fl => fl.required && !f[fl.key]);
    if (req) return alert(`${req.label} is required.`);
    try {
      const res = await saveFn(f, editing);
      if (editing) setItems(xs => xs.map(x => x.id === editing ? res : x));
      else setItems(xs => [...xs, res]);
      setOpen(false); setEdit(null); setF(blank);
    } catch (err) { alert(err.message); }
  }

  async function del(id) {
    if (!window.confirm(`Delete this ${noun}?`)) return;
    try {
      await delFn(id);
      setItems(xs => xs.filter(x => x.id !== id));
    } catch (err) { alert(err.message); }
  }

  return (
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={() => { setF(blank); setEdit(null); setOpen(true); }}>+ Add {noun}</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr>{cols.map(c => <th key={c.label} className={c.r ? "r" : ""}>{c.label}</th>)}<th>Actions</th></tr></thead>
          <tbody>
            {items.map(row => (
              <tr key={row.id}>
                {cols.map(c => <td key={c.label} className={c.r ? "r" : ""}>{c.render ? c.render(row) : row[c.key]}</td>)}
                <td>
                  <div style={{ display:"flex", gap:6 }}>
                    <button className="btn btn-ghost btn-sm" onClick={() => { setF(row); setEdit(row.id); setOpen(true); }}>Edit</button>
                    <button className="btn btn-danger btn-sm" onClick={() => del(row.id)}>Del</button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open && (
        <Modal title={`${editing ? "Edit" : "Add"} ${noun}`} onClose={() => { setOpen(false); setEdit(null); }} foot={<><button className="btn btn-ghost" onClick={() => setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>{editing ? "Update" : "Save"}</button></>}>
          <div className="form-row cols-2">
            {fields.map(fl => (
              <FG key={fl.key} label={fl.label} span={fl.full ? 2 : undefined}>
                {fl.type === "select" ? (
                  <select value={f[fl.key]} onChange={e => set(fl.key)(e.target.value)}>{fl.options.map(o => <option key={o}>{o}</option>)}</select>
                ) : fl.type === "checkbox" ? (
                  <input type="checkbox" checked={!!f[fl.key]} onChange={e => set(fl.key)(e.target.checked)} style={{ width:"auto" }} />
                ) : (
                  <input type={fl.type || "text"} value={f[fl.key] || ""} onChange={e => set(fl.key)(e.target.value)} />
                )}
              </FG>
            ))}
          </div>
        </Modal>
      )}
    </div>
  );
}

function Dashboard({ vehicles, customers, grades, weighments, boulderReceipts, productionEntries, expenses, lots, qcTests, samples, maintTasks, maintLogs }) {
  const stock = useMemo(() => computeStock(grades, boulderReceipts, productionEntries, weighments), [grades, boulderReceipts, productionEntries, weighments]);
  const stockArr = Object.values(stock);
  const boulderStk = stockArr.find(s => s.isBoulder)?.closing || 0;
  const totalRM = boulderReceipts.reduce((s, r) => s + r.quantityMT, 0);
  const totalProd = productionEntries.reduce((s, p) => s + p.totalOutputMT, 0);

  return (
    <div>
      <div className="stats-grid-5">
        <StatCard label="Boulder in Stock" value={`${boulderStk.toFixed(1)} MT`} color="teal" />
        <StatCard label="Total RM Received" value={`${totalRM.toFixed(1)} MT`} color="blue" />
        <StatCard label="Total Produced" value={`${totalProd.toFixed(1)} MT`} color="accent" />
        <StatCard label="Today Dispatches" value={weighments.filter(w => w.date === today() && w.purpose === "Sale").length} color="green" />
        <StatCard label="Pending Samples" value={samples.filter(s => s.status === "retained").length} color="purple" />
      </div>
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Live Production & Stock Ledger</h3></div>
        <table>
          <thead><tr><th>Grade</th><th>Type</th><th className="r">Inventory Balance</th></tr></thead>
          <tbody>
            {stockArr.map(s => (
              <tr key={s.gradeId}>
                <td>{s.name}</td>
                <td><Badge type={s.isBoulder ? "teal" : "blue"}>{s.isBoulder ? "Raw" : "Finished"}</Badge></td>
                <td className="r mono" style={{ color:"var(--accent)", fontWeight:700 }}>{fmtMT(s.closing)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
