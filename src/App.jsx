import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Fonts + CSS (exact Needle Point design system) ────────────────────────────
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
.badge-paid,.badge-green,.badge-ok,.badge-running,.badge-sale{background:#0a2016;color:var(--green);border:1px solid #1a4028}
.badge-partially_paid,.badge-amber,.badge-low,.badge-morning{background:#1a1500;color:var(--amber);border:1px solid #3a2e00}
.badge-overdue,.badge-red,.badge-negative,.badge-breakdown{background:#200808;color:var(--red);border:1px solid #401010}
.badge-sent,.badge-blue,.badge-inward,.badge-produced{background:#081828;color:var(--blue);border:1px solid #143050}
.badge-draft,.badge-muted,.badge-default{background:var(--bg3);color:var(--text3);border:1px solid var(--border)}
.badge-teal,.badge-raw,.badge-receipt{background:#051e1a;color:var(--teal);border:1px solid #0d3830}
.badge-purple,.badge-transfer{background:#100a28;color:var(--purple);border:1px solid #201848}
.badge-accent,.badge-igst{background:#0d1a00;color:var(--accent);border:1px solid #1e3a00}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.8);display:flex;align-items:center;justify-content:center;z-index:200;padding:20px;backdrop-filter:blur(2px)}
.modal{background:var(--bg2);border:1px solid var(--border2);border-radius:var(--r3);width:100%;max-width:600px;max-height:92vh;display:flex;flex-direction:column;box-shadow:var(--shadow2)}
.modal-lg{max-width:820px}.modal-xl{max-width:1060px}
.modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 22px;border-bottom:1px solid var(--border);flex-shrink:0}
.modal-head h3{font-size:15px;font-weight:600}
.modal-body{padding:20px 22px;overflow-y:auto;flex:1}
.modal-foot{display:flex;justify-content:flex-end;gap:8px;padding:14px 22px;border-top:1px solid var(--border);flex-shrink:0}
.form-row{display:grid;gap:14px;margin-bottom:14px}
.form-row.cols-2{grid-template-columns:1fr 1fr}
.form-row.cols-3{grid-template-columns:1fr 1fr 1fr}
.form-row.cols-4{grid-template-columns:1fr 1fr 1fr 1fr}
.form-group{display:flex;flex-direction:column;gap:5px}
.form-group label{font-size:10.5px;color:var(--text2);letter-spacing:.05em;text-transform:uppercase;font-family:var(--mono);font-weight:500}
.form-group input,.form-group select,.form-group textarea{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:8px 11px;color:var(--text);font-family:var(--font);font-size:13px;outline:none;transition:border-color .12s;width:100%}
.form-group input:focus,.form-group select:focus{border-color:var(--accent)}
.form-group input::placeholder{color:var(--text4)}
.form-note{font-size:10.5px;color:var(--text3);font-family:var(--mono);margin-top:2px}
.empty{text-align:center;padding:48px 24px;color:var(--text3)}
.empty-icon{font-size:32px;margin-bottom:12px}
.alert-strip{background:#120a00;border:1px solid #3a2000;border-radius:var(--r2);padding:14px 18px;margin-bottom:20px}
.alert-strip h4{font-size:11px;font-weight:700;color:var(--amber);text-transform:uppercase;letter-spacing:.08em;margin-bottom:10px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
.row-detail{background:#141414;border-top:1px solid var(--border);padding:12px 18px;display:flex;gap:32px;flex-wrap:wrap}
.row-detail-block label{font-size:9.5px;color:var(--text4);text-transform:uppercase;letter-spacing:.08em;font-family:var(--mono);display:block;margin-bottom:5px}
.row-detail-block p{font-size:12px;color:var(--text2)}
.output-rows{display:flex;flex-direction:column;gap:8px}
.output-row{display:grid;grid-template-columns:1fr 1fr auto;gap:8px;align-items:end}
.filter-bar{display:flex;gap:10px;margin-bottom:16px;align-items:center;flex-wrap:wrap}
.filter-bar input,.filter-bar select{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:7px 11px;color:var(--text);font-family:var(--font);font-size:12.5px;outline:none}
.filter-bar input{min-width:200px}
.err-msg{color:var(--red);font-size:12px;margin-top:8px;font-family:var(--mono)}
.bar-row{margin-bottom:12px}
.bar-row-header{display:flex;justify-content:space-between;margin-bottom:4px}
.bar-row-header span{font-size:12px;color:var(--text2)}
.bar-track{background:var(--bg4);border-radius:3px;height:4px}
.bar-fill{height:100%;border-radius:3px;transition:width .4s}
.ledger-foot td{font-size:13px;font-weight:700;border-top:1px solid var(--border2) !important;background:var(--bg3)}




/* ── Maintenance ── */
.maint-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:16px 18px;margin-bottom:10px;transition:border-color .12s}
.maint-card:hover{border-color:var(--border3)}
.maint-card.overdue{border-color:#401010;background:#120404}
.maint-card.due-soon{border-color:#3a2e00;background:#0e0c00}
.maint-card.ok{border-color:var(--border)}
.task-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
.task-name{font-weight:600;font-size:13px}
.task-meta{font-size:11px;color:var(--text3);font-family:var(--mono);margin-top:3px}
.due-badge{font-size:10px;font-family:var(--mono);padding:3px 8px;border-radius:20px;font-weight:600}
.due-badge.overdue{background:#200808;color:var(--red);border:1px solid #401010}
.due-badge.due-today{background:#1a1500;color:var(--amber);border:1px solid #3a2e00}
.due-badge.due-soon{background:#0a1500;color:#8aab20;border:1px solid #1e3a00}
.due-badge.ok{background:var(--bg3);color:var(--text4);border:1px solid var(--border)}
.maint-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:20px}
.log-param-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:10px}
.interval-chip{display:inline-flex;align-items:center;gap:4px;background:var(--bg3);border:1px solid var(--border2);border-radius:20px;padding:2px 8px;font-size:10px;font-family:var(--mono);color:var(--text3)}

/* ── Cost Sheet ── */
.cost-row{display:flex;justify-content:space-between;align-items:center;padding:9px 0;border-bottom:1px solid var(--border)}
.cost-row:last-child{border-bottom:none}
.cost-row.grand{border-top:2px solid var(--border2);margin-top:4px;padding-top:12px}
.cost-row.grand .cost-label{font-weight:700;color:var(--text);font-size:13px}
.cost-row.grand .cost-val{font-size:16px;font-weight:700;color:var(--accent)}
.cost-label{font-size:12.5px;color:var(--text2)}
.cost-val{font-family:var(--mono);font-size:12.5px;color:var(--text);text-align:right}
.cost-pct{font-size:10px;color:var(--text4);font-family:var(--mono);margin-left:6px}
.cpt-banner{background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r2);padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:20px}
.cpt-val{font-size:30px;font-weight:800;color:var(--accent);font-family:var(--mono);letter-spacing:-.5px}
.cpt-sub{font-size:11px;color:var(--text3);margin-top:3px;font-family:var(--mono)}
.margin-row{display:flex;align-items:center;gap:8px;padding:10px 14px;border-radius:var(--r);margin-bottom:6px}
.margin-row.profit{background:#051a0a;border:1px solid #0d3018}
.margin-row.loss{background:#1a0505;border:1px solid #3a1010}

/* ── QC / Lab ── */
.qc-param-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px}
.qc-param-card{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:10px 12px}
.qc-param-card label{font-size:9.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;font-family:var(--mono);display:block;margin-bottom:6px}
.qc-param-card input{background:var(--bg4);border:1px solid var(--border2);border-radius:5px;padding:6px 9px;color:var(--text);font-family:var(--mono);font-size:13px;outline:none;width:100%;transition:border-color .12s}
.qc-param-card input:focus{border-color:var(--accent)}
.qc-param-card .unit{font-size:10px;color:var(--text4);font-family:var(--mono);margin-top:3px}
.lot-chip{display:inline-flex;align-items:center;gap:5px;background:#0a1a00;border:1px solid #1e3a00;border-radius:var(--r);padding:3px 10px;font-family:var(--mono);font-size:11px;color:var(--accent)}
.sample-due{background:#1a0a00;border:1px solid #3a1a00;border-radius:var(--r);padding:10px 14px;display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}
.badge-released{background:#051e1a;color:var(--teal);border:1px solid #0d3830}
.badge-under-test{background:#081828;color:var(--blue);border:1px solid #143050}
.badge-rejected{background:#200808;color:var(--red);border:1px solid #401010}
.badge-in-stock{background:#0a2016;color:var(--green);border:1px solid #1a4028}
.badge-dispatched{background:#100a28;color:var(--purple);border:1px solid #201848}
.badge-retained{background:#051e1a;color:var(--teal);border:1px solid #0d3830}
.badge-disposed{background:var(--bg3);color:var(--text4);border:1px solid var(--border)}
.badge-pass{background:#0a2016;color:var(--green);border:1px solid #1a4028}
.badge-fail{background:#200808;color:var(--red);border:1px solid #401010}
.badge-conditional{background:#1a1500;color:var(--amber);border:1px solid #3a2e00}

/* ── Week 4: Reports & Expenses ── */
.pill-tabs{display:flex;gap:4px;background:var(--bg3);border-radius:var(--r2);padding:4px;margin-bottom:20px;width:fit-content}
.pill-tab{padding:6px 16px;border-radius:var(--r);font-size:12.5px;font-weight:500;cursor:pointer;border:none;background:none;color:var(--text3);font-family:var(--font);transition:all .12s}
.pill-tab.active{background:var(--bg2);color:var(--text);box-shadow:0 1px 4px rgba(0,0,0,.4)}
.report-section{background:var(--bg2);border:1px solid var(--border);border-radius:var(--r2);padding:20px 22px;margin-bottom:20px}
.report-section h3{font-size:13px;font-weight:600;margin-bottom:16px;color:var(--text)}
.report-controls{display:flex;gap:10px;align-items:flex-end;margin-bottom:20px;flex-wrap:wrap}
.report-controls .form-group{min-width:140px}
.mini-bar-row{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.mini-bar-row .label{font-size:11.5px;color:var(--text2);width:140px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0}
.mini-bar-row .track{flex:1;background:var(--bg4);border-radius:3px;height:6px}
.mini-bar-row .fill{height:100%;border-radius:3px}
.mini-bar-row .val{font-size:11px;font-family:var(--mono);color:var(--text3);width:80px;text-align:right;flex-shrink:0}
.expense-cat-dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:6px;flex-shrink:0}
.summary-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px}
.summary-card{background:var(--bg3);border:1px solid var(--border);border-radius:var(--r);padding:12px 16px}
.summary-card .s-label{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.07em;font-family:var(--mono);margin-bottom:6px}
.summary-card .s-val{font-size:17px;font-weight:700;letter-spacing:-.3px}
.print-hide{}.print-only{display:none}

/* ── Mobile / responsive ── */
.hamburger-btn{display:none;background:none;border:1px solid var(--border2);border-radius:var(--r);color:var(--text);width:34px;height:34px;align-items:center;justify-content:center;font-size:16px;cursor:pointer;flex-shrink:0}
.sidebar-close{display:none;background:none;border:none;color:var(--text3);font-size:16px;cursor:pointer;padding:4px}
.sidebar-backdrop{display:none}
.table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
table{min-width:640px}
@media (max-width:860px){
  .hamburger-btn{display:flex}
  .sidebar-close{display:block}
  .sidebar{position:fixed;left:0;top:0;z-index:210;width:260px;max-width:82vw;transform:translateX(-100%);transition:transform .2s ease;box-shadow:var(--shadow2)}
  .sidebar.open{transform:translateX(0)}
  .sidebar-backdrop{display:block;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:200}
  .topbar{padding:12px 14px}
  .topbar h2{font-size:15px}
  .topbar-date{display:none}
  .content{padding:14px 12px}
  .stats-grid,.stats-grid-5{grid-template-columns:repeat(2,1fr)}
  .form-row.cols-2,.form-row.cols-3,.form-row.cols-4{grid-template-columns:1fr}
  .maint-grid{grid-template-columns:1fr}
  .qc-param-grid{grid-template-columns:repeat(2,1fr)}
  .log-param-row{grid-template-columns:1fr}
  .summary-grid{grid-template-columns:1fr}
  .modal-overlay{padding:0}
  .modal{max-width:100%;width:100%;height:100%;max-height:100%;border-radius:0}
  .modal-body{padding:16px}
  .filter-bar input{min-width:0;flex:1 1 140px}
  .pill-tabs{overflow-x:auto;-webkit-overflow-scrolling:touch;max-width:100%}
  .pill-tab{white-space:nowrap}
  .btn{padding:9px 14px}
  .nav-item{padding:10px 12px}
}
@media (max-width:460px){
  .stats-grid,.stats-grid-5,.qc-param-grid{grid-template-columns:1fr}
  .stat-val{font-size:19px}
}
@media print{
  .sidebar,.topbar,.btn,.filter-bar,.pill-tabs,.print-hide{display:none!important}
  .print-only{display:block!important}
  body{background:#fff;color:#000}
  .report-section{border:1px solid #ddd;page-break-inside:avoid}
  table th,table td{border-bottom:1px solid #eee;color:#000}
}

`;

// ── Seed Data ─────────────────────────────────────────────────────────────────
const seedVehicles = [
  { id:1, vehicleNo:"MH12AB1234", type:"Tipper",  tareWeight:8200,  owner:"Ramesh Transport", phone:"9876543210" },
  { id:2, vehicleNo:"GJ05CD5678", type:"Dumper",  tareWeight:12500, owner:"Patel Carriers",   phone:"9123456789" },
];
const seedCustomers = [
  { id:1, name:"Sharma Constructions", gstin:"27AABCS1234A1Z5", contact:"Anil Sharma", phone:"9001234567", address:"Pune, MH",    state:"Maharashtra" },
  { id:2, name:"BuildRight Infra",     gstin:"24AABCB5678B1Z2", contact:"Priya Mehta", phone:"9009876543", address:"Surat, GJ",   state:"Gujarat" },
];
const seedGrades = [
  { id:1, code:"QD",  name:"Quartz Dust",          unit:"MT", description:"< 6mm fine dust",         isBoulder:false },
  { id:2, code:"Q6",  name:"Quartz Grit 6mm",       unit:"MT", description:"6–10mm grit",              isBoulder:false },
  { id:3, code:"Q20", name:"Quartz Aggregate 20mm",  unit:"MT", description:"20mm aggregate",          isBoulder:false },
  { id:4, code:"QBL", name:"Quartz Boulder",         unit:"MT", description:"Raw unprocessed boulder", isBoulder:true  },
];
const seedSuppliers = [
  { id:1, name:"Rajput Mines Pvt Ltd",  contact:"Deepak Rajput", phone:"9811122233", address:"Rajasthan" },
  { id:2, name:"Aravalli Stone Quarry", contact:"Hemant Gupta",  phone:"9922233344", address:"Rajasthan" },
];
const seedEquipment=[
  {id:1,code:"BM-01",name:"Ball Mill #1",       type:"Ball Mill",          capacity:"5 TPH",  make:"Laxmi Engineers",installDate:"2024-01-15",status:"active",notes:""},
  {id:2,code:"CL-01",name:"Air Classifier #1",  type:"Classifier",         capacity:"5 TPH",  make:"Laxmi Engineers",installDate:"2024-01-15",status:"active",notes:""},
  {id:3,code:"MS-01",name:"Magnetic Separator",  type:"Magnetic Separator", capacity:"5 TPH",  make:"Electro Flux",   installDate:"2024-01-15",status:"active",notes:""},
  {id:4,code:"JC-01",name:"Jaw Crusher #1",      type:"Jaw Crusher",        capacity:"10 TPH", make:"Puzzolana",      installDate:"2024-01-20",status:"active",notes:""},
  {id:5,code:"VS-01",name:"Vibrating Screen #1", type:"Vibrating Screen",   capacity:"8 TPH",  make:"AViTEQ",         installDate:"2024-01-20",status:"active",notes:""},
  {id:6,code:"AC-01",name:"Air Compressor",      type:"Air Compressor",     capacity:"200 CFM",make:"Elgi",           installDate:"2024-01-15",status:"active",notes:""},
];
const SHIFTS    = ["Morning (6AM–2PM)","Afternoon (2PM–10PM)","Night (10PM–6AM)"];
const OPERATORS = ["Suresh Kumar","Manoj Singh","Ravi Patil","Dinesh Yadav"];

const fmtMT  = v => `${(+v||0).toFixed(3)} MT`;
const fmtWt  = kg => kg>=1000?`${(kg/1000).toFixed(3)} MT`:`${kg} kg`;
const fmt    = n => `₹${(+n||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const today  = () => new Date().toISOString().split("T")[0];
const nowTime= () => new Date().toTimeString().slice(0,5);
const uid    = () => Date.now()+Math.random().toString(36).slice(2,6);

const nextLotNum=(lots,date)=>{
  const fy=date.slice(0,4);
  const existing=lots.filter(l=>l.lotNo.includes(fy));
  return `LOT-${fy}-${String(existing.length+1).padStart(4,"0")}`;
};
const addDays=(dateStr,n)=>{ const d=new Date(dateStr); d.setDate(d.getDate()+n); return d.toISOString().split("T")[0]; };
// Indian fiscal year: Apr 1 – Mar 31, labelled "2026-27" for the year starting Apr 2026.
const getFY=dateStr=>{
  if(!dateStr)return "";
  const d=new Date(dateStr); const y=d.getFullYear(), m=d.getMonth()+1;
  const startYear = m>=4 ? y : y-1;
  return `${startYear}-${String((startYear+1)%100).padStart(2,"0")}`;
};
const currentFY=()=>getFY(today());
const fyRange=fy=>{ const startYear=+fy.split("-")[0]; return {start:`${startYear}-04-01`,end:`${startYear+1}-03-31`}; };
const calcNextDue=(lastDone,intervalType,intervalValue)=>{
  if(!lastDone) return today();
  const n = +intervalValue||1;
  const d = new Date(lastDone);
  if(intervalType==="daily")   d.setDate(d.getDate()+n);
  if(intervalType==="weekly")  d.setDate(d.getDate()+n*7);
  if(intervalType==="monthly") d.setMonth(d.getMonth()+n);
  return d.toISOString().split("T")[0];
};
const daysDiff=(dateStr)=>{
  const diff=new Date(dateStr)-new Date(today());
  return Math.ceil(diff/(1000*60*60*24));
};




const seedMaintTasks=[
  {id:1,equipmentId:3,code:"MS-01",equipmentName:"Magnetic Separator",taskName:"Magnet Cleaning",category:"cleaning",intervalType:"daily",intervalValue:1,estimatedMins:20,lastDone:"",nextDue:today(),active:true,notes:"Clean collected ferrous material, log weight caught"},
  {id:2,equipmentId:1,code:"BM-01",equipmentName:"Ball Mill #1",taskName:"Resin Lining Inspection",category:"inspection",intervalType:"monthly",intervalValue:1,estimatedMins:45,lastDone:"",nextDue:today(),active:true,notes:"Check wear %, record thickness at 4 points"},
  {id:3,equipmentId:1,code:"BM-01",equipmentName:"Ball Mill #1",taskName:"Ball Charge Inspection",category:"inspection",intervalType:"monthly",intervalValue:1,estimatedMins:30,lastDone:"",nextDue:today(),active:true,notes:"Weigh remaining balls, check for cracked/deformed balls"},
  {id:4,equipmentId:1,code:"BM-01",equipmentName:"Ball Mill #1",taskName:"Bearing Lubrication",category:"lubrication",intervalType:"weekly",intervalValue:2,estimatedMins:15,lastDone:"",nextDue:today(),active:true,notes:"Apply EP-2 grease, check temperature"},
  {id:5,equipmentId:5,code:"VS-01",equipmentName:"Vibrating Screen #1",taskName:"Screen Mesh Inspection",category:"inspection",intervalType:"weekly",intervalValue:1,estimatedMins:20,lastDone:"",nextDue:today(),active:true,notes:"Check for holes, blinding, wear at edges"},
  {id:6,equipmentId:2,code:"CL-01",equipmentName:"Air Classifier #1",taskName:"Classifier Blade Inspection",category:"inspection",intervalType:"monthly",intervalValue:1,estimatedMins:30,lastDone:"",nextDue:today(),active:true,notes:"Check blade wear, rotor speed setting"},
  {id:7,equipmentId:4,code:"JC-01",equipmentName:"Jaw Crusher #1",taskName:"Jaw Plate Inspection",category:"inspection",intervalType:"weekly",intervalValue:2,estimatedMins:20,lastDone:"",nextDue:today(),active:true,notes:"Check wear, check toggle plate, grease toggle pins"},
  {id:8,equipmentId:6,code:"AC-01",equipmentName:"Air Compressor",taskName:"Compressor Oil & Filter",category:"service",intervalType:"monthly",intervalValue:1,estimatedMins:30,lastDone:"",nextDue:today(),active:true,notes:"Change oil, clean/replace air filter"},
  {id:9,equipmentId:1,code:"BM-01",equipmentName:"Ball Mill #1",taskName:"Drive Belt Inspection",category:"inspection",intervalType:"weekly",intervalValue:1,estimatedMins:15,lastDone:"",nextDue:today(),active:true,notes:"Check tension, alignment, and wear"},
  {id:10,equipmentId:3,code:"MS-01",equipmentName:"Magnetic Separator",taskName:"Magnet Strength Check",category:"inspection",intervalType:"monthly",intervalValue:1,estimatedMins:20,lastDone:"",nextDue:today(),active:true,notes:"Test with known weight of iron particles"},
];

// ── Stock Engine ──────────────────────────────────────────────────────────────
function computeStock(grades,boulderReceipts,productionEntries,weighments){
  const s={};
  grades.forEach(g=>{s[g.id]={...g,gradeId:g.id,in:0,produced:0,sold:0,transferred:0};});
  boulderReceipts.forEach(r=>{const b=grades.find(g=>g.isBoulder);if(b&&s[b.id])s[b.id].in+=r.quantityMT;});
  productionEntries.forEach(pe=>{
    const b=grades.find(g=>g.isBoulder);
    if(b&&s[b.id])s[b.id].produced-=pe.boulderConsumedMT;
    pe.outputs.forEach(o=>{if(s[o.gradeId])s[o.gradeId].produced+=o.quantityMT;});
  });
  weighments.forEach(w=>{
    const mt=(w.netWeight||0)/1000;
    if(w.purpose==="Sale"&&s[w.gradeId])s[w.gradeId].sold+=mt;
    if(w.purpose==="Inward (Raw)"&&s[w.gradeId])s[w.gradeId].in+=mt;
    if(w.purpose==="Stock Transfer"&&s[w.gradeId])s[w.gradeId].transferred+=mt;
  });
  Object.values(s).forEach(g=>{g.closing=g.isBoulder?g.in+g.produced:g.produced-g.sold-g.transferred;});
  return s;
}

// ── Shared UI ─────────────────────────────────────────────────────────────────
function Badge({type,children}){return <span className={`badge badge-${type||"muted"}`}>{children}</span>;}
function StatCard({label,value,sub,color="blue"}){
  return(<div className={`stat-card ${color}`}><div className="stat-label">{label}</div><div className={`stat-val ${color}`}>{value}</div>{sub&&<div className="stat-sub">{sub}</div>}</div>);
}
function FG({label,note,children,span}){
  return(<div className="form-group" style={span?{gridColumn:`span ${span}`}:{}}>{label&&<label>{label}</label>}{children}{note&&<div className="form-note">{note}</div>}</div>);
}
function Modal({title,onClose,children,foot,size=""}){
  return(
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
  return(<div className="empty"><div className="empty-icon">{icon||"📭"}</div><p style={{fontWeight:600,marginBottom:4}}>{message}</p>{sub&&<p style={{fontSize:11.5,color:"var(--text3)",marginTop:4}}>{sub}</p>}</div>);
}

// ── NAV ───────────────────────────────────────────────────────────────────────

// ── NAV ─────────────────────────────────────────────────────────────────────
const NAV=[
  {id:"dashboard",  label:"Dashboard",   icon:"◈", group:null},
  {id:"weighment",  label:"Weighment",   icon:"⚖",  group:"Operations"},
  {id:"customers",  label:"Customers",   icon:"◈", group:"Operations"},
  {id:"boulder",    label:"Boulder In",  icon:"⬡",  group:"Operations"},
  {id:"production", label:"Production",  icon:"⚙",  group:"Operations"},
  {id:"stock",      label:"Stock Ledger",icon:"▤",  group:"Operations"},
  {id:"shift",      label:"Shift Log",   icon:"◷",  group:"Operations"},
  {id:"maint",      label:"Maint. Overview",icon:"⚡", group:"Maintenance"},
  {id:"tasks",      label:"Task Schedule", icon:"📋", group:"Maintenance"},
  {id:"maintlog",   label:"Activity Log",  icon:"🔧", group:"Maintenance"},
  {id:"equipment",  label:"Equipment",     icon:"🔩", group:"Maintenance"},
  {id:"lots",       label:"Lot Register",icon:"🏷",  group:"Quality"},
  {id:"qc",         label:"QC Tests",    icon:"🔬",  group:"Quality"},
  {id:"samples",    label:"Samples",     icon:"🧪",  group:"Quality"},
  {id:"purchases",  label:"Purchases",   icon:"📥",  group:"Reports"},
  {id:"opscosts",   label:"Monthly Ops Costs",icon:"🧾", group:"Reports"},
  {id:"costs",      label:"Cost Sheet",  icon:"💰",  group:"Reports"},
  {id:"costconfig",  label:"Cost Config",  icon:"⚙",  group:"Reports"},
  {id:"expenses",   label:"Expenses",    icon:"💸",  group:"Reports"},
  {id:"capital",    label:"Capital / Infra",icon:"🏗", group:"Reports"},
  {id:"bankstatement",label:"Bank Statement",icon:"🏦", group:"Reports"},
  {id:"partners",     label:"Partners & Capital",icon:"🤝", group:"Reports"},
  {id:"rpt-daily",  label:"Daily Report",icon:"📅",  group:"Reports"},
  {id:"rpt-stock",  label:"Stock Report",icon:"📦",  group:"Reports"},
  {id:"vehicles",   label:"Vehicles",    icon:"◉", group:"Masters"},
  {id:"suppliers",  label:"Suppliers",   icon:"◈", group:"Masters"},
  {id:"grades",     label:"Grades",      icon:"◆", group:"Masters"},
];
const TITLES=Object.fromEntries(NAV.map(n=>[n.id,n.label]));
const GROUPS=[...new Set(NAV.filter(n=>n.group).map(n=>n.group))];

function WeighmentEntry({vehicles,customers,grades,weighments,setWeighments}){
  const [open,setOpen]=useState(false);const [search,setSearch]=useState("");
  const blank={date:today(),time:nowTime(),vehicleId:"",customerId:"",gradeId:"",lotNo:"",grossWeight:"",tareWeight:"",purpose:"Sale",remarks:""};
  const [f,setF]=useState(blank);const set=k=>v=>setF(x=>({...x,[k]:v}));
  const pickV=id=>{const v=vehicles.find(v=>String(v.id)===id);setF(x=>({...x,vehicleId:id,tareWeight:v?String(v.tareWeight):""}));};
  const net=()=>{const g=parseFloat(f.grossWeight),t=parseFloat(f.tareWeight);return(!isNaN(g)&&!isNaN(t))?g-t:null;};
  const netVal=net();
  function save(){
    if(!f.vehicleId||!f.gradeId||!f.grossWeight||!f.tareWeight)return alert("Vehicle, Grade, Gross and Tare required.");
    if(netVal<=0)return alert("Net weight must be positive.");
    const v=vehicles.find(v=>String(v.id)===String(f.vehicleId)),c=customers.find(c=>String(c.id)===String(f.customerId)),g=grades.find(g=>String(g.id)===String(f.gradeId));
    setWeighments(ws=>[{id:uid(),slipNo:`WS-${String(ws.length+1).padStart(4,"0")}`, ...f,grossWeight:+f.grossWeight,tareWeight:+f.tareWeight,netWeight:netVal,vehicleNo:v?.vehicleNo,customerName:c?.name||"—",gradeName:g?.name,gradeCode:g?.code,lotNo:f.lotNo||""},...ws]);
    setOpen(false);setF(blank);
  }
  const filtered=weighments.filter(w=>(w.vehicleNo||"").toLowerCase().includes(search.toLowerCase())||(w.customerName||"").toLowerCase().includes(search.toLowerCase())||(w.gradeCode||"").toLowerCase().includes(search.toLowerCase()));
  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Weighments" value={weighments.length} sub="all time" color="blue"/>
        <StatCard label="Today's Sales" value={weighments.filter(w=>w.date===today()&&w.purpose==="Sale").length} sub="dispatches" color="accent"/>
        <StatCard label="Today Net (MT)" value={fmtMT(weighments.filter(w=>w.date===today()&&w.purpose==="Sale").reduce((s,w)=>s+w.netWeight/1000,0))} sub="sold today" color="teal"/>
      </div>
      <div className="filter-bar">
        <input placeholder="Search vehicle, customer, grade…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New Weighment</button>
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Slip #</th><th>Date/Time</th><th>Vehicle</th><th>Customer</th><th>Grade</th><th className="r">Gross</th><th className="r">Tare</th><th className="r">Net</th><th>Purpose</th></tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9}><EmptyState icon="⚖" message="No weighment slips yet"/></td></tr>
            :filtered.map(w=>(
              <tr key={w.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{w.slipNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{w.date} {w.time}</td>
                <td className="mono" style={{fontWeight:600}}>{w.vehicleNo}</td>
                <td style={{color:"var(--text2)"}}>{w.customerName}</td>
                <td><Badge type="muted">{w.gradeCode}</Badge></td>
                <td className="r mono" style={{color:"var(--text3)"}}>{w.grossWeight?.toLocaleString()} kg</td>
                <td className="r mono" style={{color:"var(--text4)"}}>{w.tareWeight?.toLocaleString()} kg</td>
                <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtWt(w.netWeight)}</td>
                <td><Badge type={w.purpose==="Sale"?"green":w.purpose==="Inward (Raw)"?"blue":"muted"}>{w.purpose}</Badge></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open&&(
        <Modal title="New Weighment Slip" onClose={()=>setOpen(false)}
          foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Slip</button></>}>
          <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG><FG label="Time"><input type="time" value={f.time} onChange={e=>set("time")(e.target.value)}/></FG></div>
          <div className="form-row cols-2">
            <FG label="Vehicle *"><select value={f.vehicleId} onChange={e=>pickV(e.target.value)}><option value="">Select…</option>{vehicles.map(v=><option key={v.id} value={v.id}>{v.vehicleNo} — {v.type}</option>)}</select></FG>
            <FG label="Customer"><select value={f.customerId} onChange={e=>set("customerId")(e.target.value)}><option value="">Select…</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Grade *"><select value={f.gradeId} onChange={e=>set("gradeId")(e.target.value)}><option value="">Select…</option>{grades.map(g=><option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></FG>
            <FG label="Purpose"><select value={f.purpose} onChange={e=>set("purpose")(e.target.value)}>{["Sale","Inward (Raw)","Stock Transfer","Rejection"].map(p=><option key={p}>{p}</option>)}</select></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Gross Weight (kg) *"><input type="number" placeholder="20500" value={f.grossWeight} onChange={e=>set("grossWeight")(e.target.value)}/></FG>
            <FG label="Tare Weight (kg) *"><input type="number" placeholder="8200" value={f.tareWeight} onChange={e=>set("tareWeight")(e.target.value)}/></FG>
          </div>
          {netVal!==null&&<div className={`net-display ${netVal>0?"positive":"negative"}`} style={{marginBottom:14}}><span style={{fontSize:12,color:"var(--text3)",fontFamily:"var(--mono)"}}>Net Weight</span><span style={{fontFamily:"var(--mono)",fontWeight:700,fontSize:18,color:netVal>0?"var(--teal)":"var(--red)"}}>{netVal.toLocaleString()} kg</span></div>}
          <div className="form-row"><FG label="Remarks"><input placeholder="Optional" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG></div>
        </Modal>
      )}
    </div>
  );
}

function BoulderReceipt({suppliers,vehicles,boulderReceipts,setBoulderReceipts}){
  const [open,setOpen]=useState(false);
  const blank={date:today(),vehicleId:"",supplierId:"",quantityMT:"",royaltyNo:"",challanNo:"",remarks:""};
  const [f,setF]=useState(blank);const set=k=>v=>setF(x=>({...x,[k]:v}));
  function save(){
    if(!f.supplierId||!f.quantityMT)return alert("Supplier and Quantity required.");
    const s=suppliers.find(s=>String(s.id)===String(f.supplierId)),v=vehicles.find(v=>String(v.id)===String(f.vehicleId));
    setBoulderReceipts(rs=>[{id:uid(),receiptNo:`BR-${String(rs.length+1).padStart(4,"0")}`,...f,quantityMT:+f.quantityMT,supplierName:s?.name,vehicleNo:v?.vehicleNo||"—"},...rs]);
    setOpen(false);setF(blank);
  }
  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Received" value={`${boulderReceipts.reduce((s,r)=>s+r.quantityMT,0).toFixed(1)} MT`} sub={`${boulderReceipts.length} deliveries`} color="teal"/>
        <StatCard label="This Month" value={`${boulderReceipts.filter(r=>r.date?.slice(0,7)===today().slice(0,7)).reduce((s,r)=>s+r.quantityMT,0).toFixed(1)} MT`} sub="current month" color="blue"/>
        <StatCard label="Suppliers Used" value={[...new Set(boulderReceipts.map(r=>r.supplierId))].length} sub="unique" color="accent"/>
      </div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New Receipt</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Receipt #</th><th>Date</th><th>Supplier</th><th>Vehicle</th><th>Royalty #</th><th>Challan #</th><th className="r">Qty (MT)</th></tr></thead>
          <tbody>
            {boulderReceipts.length===0?<tr><td colSpan={7}><EmptyState icon="🪨" message="No boulder receipts yet"/></td></tr>
            :boulderReceipts.map(r=>(
              <tr key={r.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{r.receiptNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{r.date}</td>
                <td style={{fontWeight:500}}>{r.supplierName}</td>
                <td className="mono">{r.vehicleNo}</td>
                <td style={{color:"var(--text3)"}}>{r.royaltyNo||"—"}</td>
                <td style={{color:"var(--text3)"}}>{r.challanNo||"—"}</td>
                <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtMT(r.quantityMT)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open&&<Modal title="New Boulder Receipt" onClose={()=>setOpen(false)} foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
        <div className="form-row cols-2"><FG label="Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG><FG label="Supplier *"><select value={f.supplierId} onChange={e=>set("supplierId")(e.target.value)}><option value="">Select…</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select></FG></div>
        <div className="form-row cols-2"><FG label="Vehicle"><select value={f.vehicleId} onChange={e=>set("vehicleId")(e.target.value)}><option value="">Select…</option>{vehicles.map(v=><option key={v.id} value={v.id}>{v.vehicleNo}</option>)}</select></FG><FG label="Quantity (MT) *"><input type="number" step="0.001" placeholder="85.500" value={f.quantityMT} onChange={e=>set("quantityMT")(e.target.value)}/></FG></div>
        <div className="form-row cols-2"><FG label="Royalty / Gate Pass #"><input placeholder="RY-2024-XXXX" value={f.royaltyNo} onChange={e=>set("royaltyNo")(e.target.value)}/></FG><FG label="Challan / Invoice #"><input placeholder="CH-XXXX" value={f.challanNo} onChange={e=>set("challanNo")(e.target.value)}/></FG></div>
        <div className="form-row"><FG label="Remarks"><input placeholder="Optional" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG></div>
      </Modal>}
    </div>
  );
}

function ProductionEntry({grades,productionEntries,setProductionEntries,lots,setLots}){
  const finished=grades.filter(g=>!g.isBoulder);
  const blankOut=()=>({gradeId:finished[0]?.id||"",quantityMT:""});
  const blank={date:today(),shift:SHIFTS[0],operator:OPERATORS[0],boulderConsumedMT:"",machineHours:"",idleHours:"0",outputs:[blankOut()],remarks:""};
  const [open,setOpen]=useState(false);const [f,setF]=useState(blank);const [exp,setExp]=useState(null);
  const set=k=>v=>setF(x=>({...x,[k]:v}));
  const setOut=(i,k,v)=>setF(x=>{const o=[...x.outputs];o[i]={...o[i],[k]:v};return{...x,outputs:o};});
  const totalOut=f.outputs.reduce((s,o)=>s+(parseFloat(o.quantityMT)||0),0);
  const yieldPct=f.boulderConsumedMT>0?((totalOut/+f.boulderConsumedMT)*100).toFixed(1):null;
  function save(){
    if(!f.boulderConsumedMT)return alert("Boulder consumed required.");
    if(!f.outputs.every(o=>o.gradeId&&o.quantityMT))return alert("Fill all output rows.");
    const enriched=f.outputs.map(o=>{const g=grades.find(g=>String(g.id)===String(o.gradeId));return{...o,quantityMT:+o.quantityMT,gradeName:g?.name,gradeCode:g?.code};});
    const lotNo=nextLotNum(lots,f.date);
    const newLot={id:uid(),lotNo,date:f.date,shift:f.shift,operator:f.operator,gradeOutputs:enriched,totalOutputMT:totalOut,boulderConsumedMT:+f.boulderConsumedMT,yieldPct:yieldPct?+yieldPct:null,status:"in-stock",remarks:f.remarks};
    setLots(ls=>[newLot,...ls]);
    setProductionEntries(es=>[{id:uid(),entryNo:`PE-${String(es.length+1).padStart(4,"0")}`,lotNo,...f,boulderConsumedMT:+f.boulderConsumedMT,machineHours:+f.machineHours||0,idleHours:+f.idleHours||0,outputs:enriched,totalOutputMT:totalOut,yieldPct:yieldPct?+yieldPct:null},...es]);
    setOpen(false);setF(blank);
  }
  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Entries" value={productionEntries.length} sub="all time" color="blue"/>
        <StatCard label="Total Produced" value={`${productionEntries.reduce((s,p)=>s+p.totalOutputMT,0).toFixed(1)} MT`} sub="all grades" color="accent"/>
        <StatCard label="Avg Yield" value={productionEntries.filter(p=>p.yieldPct).length?`${(productionEntries.reduce((s,p)=>s+(p.yieldPct||0),0)/productionEntries.filter(p=>p.yieldPct).length).toFixed(1)}%`:"—"} sub="boulder→output" color="green"/>
        <StatCard label="Today" value={`${productionEntries.filter(p=>p.date===today()).reduce((s,p)=>s+p.totalOutputMT,0).toFixed(1)} MT`} sub="today" color="teal"/>
      </div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New Entry</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Entry #</th><th>Date</th><th>Shift</th><th>Operator</th><th className="r">Consumed</th><th className="r">Produced</th><th>Yield</th><th></th></tr></thead>
          <tbody>
            {productionEntries.length===0?<tr><td colSpan={8}><EmptyState icon="⚙" message="No production entries yet"/></td></tr>
            :productionEntries.map(pe=>(
              <>
              <tr key={pe.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{pe.entryNo}</td>
                <td><span className="lot-chip">🏷 {pe.lotNo||"--"}</span></td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{pe.date}</td>
                <td><Badge type="muted">{pe.shift.split(" ")[0]}</Badge></td>
                <td style={{color:"var(--text2)"}}>{pe.operator}</td>
                <td className="r mono">{fmtMT(pe.boulderConsumedMT)}</td>
                <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtMT(pe.totalOutputMT)}</td>
                <td>{pe.yieldPct?<Badge type={pe.yieldPct>=90?"green":pe.yieldPct>=75?"amber":"red"}>{pe.yieldPct}%</Badge>:<span style={{color:"var(--text4)"}}>—</span>}</td>
                <td><button className="btn btn-ghost btn-sm" onClick={()=>setExp(exp===pe.id?null:pe.id)}>{exp===pe.id?"▲":"▼"}</button></td>
              </tr>
              {exp===pe.id&&<tr key={pe.id+"d"}><td colSpan={8} style={{padding:0}}><div className="row-detail">{pe.outputs.map((o,i)=><div key={i} className="row-detail-block"><label>[{o.gradeCode}]</label><p>{fmtMT(o.quantityMT)}</p></div>)}{pe.remarks&&<div className="row-detail-block"><label>Remarks</label><p>{pe.remarks}</p></div>}</div></td></tr>}
              </>
            ))}
          </tbody>
        </table>
      </div>
      {open&&<Modal title="New Production Entry" onClose={()=>setOpen(false)} foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Entry</button></>}>
        <div className="form-row cols-3"><FG label="Date"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG><FG label="Shift"><select value={f.shift} onChange={e=>set("shift")(e.target.value)}>{SHIFTS.map(s=><option key={s}>{s}</option>)}</select></FG><FG label="Operator"><select value={f.operator} onChange={e=>set("operator")(e.target.value)}>{OPERATORS.map(o=><option key={o}>{o}</option>)}</select></FG></div>
        <div className="form-row cols-2"><FG label="Boulder Consumed (MT) *"><input type="number" step="0.001" placeholder="85.000" value={f.boulderConsumedMT} onChange={e=>set("boulderConsumedMT")(e.target.value)}/></FG><FG label="Machine Hours"><input type="number" placeholder="7.5" value={f.machineHours} onChange={e=>set("machineHours")(e.target.value)}/></FG></div>
        {yieldPct&&<div style={{background:"#0a1a00",border:"1px solid #1e3a00",borderRadius:"var(--r)",padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><span style={{fontSize:12,color:"var(--text3)",fontFamily:"var(--mono)"}}>Live Yield</span><span style={{fontFamily:"var(--mono)",fontWeight:700,color:"var(--accent)",fontSize:19}}>{yieldPct}%</span></div>}
        <div style={{marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><label style={{fontSize:10.5,color:"var(--text2)",textTransform:"uppercase",letterSpacing:".05em",fontFamily:"var(--mono)"}}>Output by Grade *</label><button className="btn btn-ghost btn-sm" onClick={()=>setF(x=>({...x,outputs:[...x.outputs,blankOut()]}))}>+ Add</button></div>
          <div className="output-rows">{f.outputs.map((row,i)=><div key={i} className="output-row"><div className="form-group"><select value={row.gradeId} onChange={e=>setOut(i,"gradeId",e.target.value)}><option value="">Grade…</option>{finished.map(g=><option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></div><div className="form-group"><input type="number" step="0.001" placeholder="MT" value={row.quantityMT} onChange={e=>setOut(i,"quantityMT",e.target.value)}/></div><button className="btn btn-danger btn-sm" onClick={()=>setF(x=>({...x,outputs:x.outputs.filter((_,idx)=>idx!==i)}))}>✕</button></div>)}</div>
        </div>
        <div className="form-row"><FG label="Remarks"><input placeholder="Notes…" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG></div>
      </Modal>}
    </div>
  );
}

function StockLedger({grades,boulderReceipts,productionEntries,weighments}){
  const [sel,setSel]=useState(null);const [dateFrom,setFrom]=useState("");const [dateTo,setTo]=useState(today());
  const stock=useMemo(()=>computeStock(grades,boulderReceipts,productionEntries,weighments),[grades,boulderReceipts,productionEntries,weighments]);
  const stockArr=Object.values(stock);
  const txns=useMemo(()=>{
    if(!sel)return[];const t=[];
    boulderReceipts.forEach(r=>{const b=grades.find(g=>g.isBoulder);if(b&&String(b.id)===sel)t.push({date:r.date,type:"receipt",desc:`From ${r.supplierName}`,qty:r.quantityMT,ref:r.receiptNo});});
    productionEntries.forEach(pe=>{const b=grades.find(g=>g.isBoulder);if(b&&String(b.id)===sel)t.push({date:pe.date,type:"consumed",desc:`${pe.entryNo}`,qty:-pe.boulderConsumedMT,ref:pe.entryNo});pe.outputs.forEach(o=>{if(String(o.gradeId)===sel)t.push({date:pe.date,type:"produced",desc:`${pe.entryNo}`,qty:o.quantityMT,ref:pe.entryNo});});});
    weighments.forEach(w=>{if(String(w.gradeId)===sel){const mt=(w.netWeight||0)/1000;if(w.purpose==="Sale")t.push({date:w.date,type:"sale",desc:`${w.slipNo} · ${w.customerName}`,qty:-mt,ref:w.slipNo});if(w.purpose==="Inward (Raw)")t.push({date:w.date,type:"receipt",desc:`${w.slipNo}`,qty:mt,ref:w.slipNo});}});
    return t.filter(x=>(!dateFrom||x.date>=dateFrom)&&(!dateTo||x.date<=dateTo)).sort((a,b)=>a.date.localeCompare(b.date));
  },[sel,grades,boulderReceipts,productionEntries,weighments,dateFrom,dateTo]);
  let running=0;const txnsWB=txns.map(t=>{running+=t.qty;return{...t,bal:running};});
  return(
    <div>
      <div className="table-wrap" style={{marginBottom:20}}>
        <div className="table-toolbar"><h3>Grade-wise Stock Position</h3><span style={{fontSize:11,color:"var(--text3)"}}>Click row to drill down</span></div>
        <table>
          <thead><tr><th>Grade</th><th>Type</th><th className="r">In / Produced</th><th className="r">Sold / Consumed</th><th className="r">Closing</th><th>Status</th></tr></thead>
          <tbody>
            {stockArr.map(s=>{const st=s.closing<0?"red":s.closing<10?"amber":"green";return(
              <tr key={s.gradeId} style={{cursor:"pointer",background:sel===String(s.gradeId)?"rgba(200,240,100,.04)":""}} onClick={()=>setSel(sel===String(s.gradeId)?null:String(s.gradeId))}>
                <td><Badge type={s.isBoulder?"teal":"muted"}>{s.code}</Badge> <span style={{marginLeft:6,fontSize:12,color:"var(--text2)"}}>{s.name}</span></td>
                <td><Badge type={s.isBoulder?"teal":"blue"}>{s.isBoulder?"Raw":"Finished"}</Badge></td>
                <td className="r mono" style={{color:"var(--teal)"}}>{fmtMT(s.isBoulder?s.in:s.produced)}</td>
                <td className="r mono" style={{color:"var(--amber)"}}>{fmtMT(s.isBoulder?Math.abs(s.produced):s.sold)}</td>
                <td className="r mono" style={{color:st==="red"?"var(--red)":st==="amber"?"var(--amber)":"var(--accent)",fontWeight:700}}>{fmtMT(s.closing)}</td>
                <td><Badge type={st}>{st==="red"?"⚠ Negative":st==="amber"?"Low":"OK"}</Badge></td>
              </tr>);
            })}
          </tbody>
        </table>
      </div>
      {sel&&<div className="table-wrap">
        <div className="table-toolbar">
          <div><h3>Ledger — {grades.find(g=>String(g.id)===sel)?.name}</h3><span style={{fontSize:11,color:"var(--text3)"}}>{txnsWB.length} txns</span></div>
          <div style={{display:"flex",gap:10,alignItems:"flex-end"}}>
            <div className="form-group"><label style={{fontSize:10,color:"var(--text3)",fontFamily:"var(--mono)"}}>From</label><input type="date" value={dateFrom} onChange={e=>setFrom(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"5px 9px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12,outline:"none"}}/></div>
            <div className="form-group"><label style={{fontSize:10,color:"var(--text3)",fontFamily:"var(--mono)"}}>To</label><input type="date" value={dateTo} onChange={e=>setTo(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"5px 9px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12,outline:"none"}}/></div>
            <button className="btn btn-ghost btn-sm" onClick={()=>{setFrom("");setTo(today());}}>Clear</button>
          </div>
        </div>
        <table>
          <thead><tr><th>Date</th><th>Type</th><th>Description</th><th>Ref</th><th className="r">Qty (MT)</th><th className="r">Balance (MT)</th></tr></thead>
          <tbody>
            {txnsWB.length===0?<tr><td colSpan={6}><EmptyState icon="📋" message="No transactions"/></td></tr>
            :txnsWB.map((t,i)=><tr key={i}><td className="mono" style={{color:"var(--text3)",fontSize:11}}>{t.date}</td><td><Badge type={t.type}>{t.type}</Badge></td><td style={{color:"var(--text2)",fontSize:12}}>{t.desc}</td><td className="mono" style={{color:"var(--text3)",fontSize:11}}>{t.ref}</td><td className="r mono" style={{color:t.qty>=0?"var(--teal)":"var(--amber)",fontWeight:700}}>{t.qty>=0?"+":""}{fmtMT(t.qty)}</td><td className="r mono" style={{color:t.bal<0?"var(--red)":"var(--accent)",fontWeight:700}}>{fmtMT(t.bal)}</td></tr>)}
          </tbody>
          {txnsWB.length>0&&<tfoot><tr className="ledger-foot"><td colSpan={5} style={{padding:"10px 16px",fontSize:12,color:"var(--text2)"}}>Closing Balance</td><td className="r mono" style={{padding:"10px 16px",fontSize:14,color:txnsWB.at(-1).bal<0?"var(--red)":"var(--accent)"}}>{fmtMT(txnsWB.at(-1).bal)}</td></tr></tfoot>}
        </table>
      </div>}
    </div>
  );
}

function ShiftLog({grades,shiftLogs,setShiftLogs}){
  const blankOut=()=>({gradeId:grades[0]?.id||"",quantity:""});
  const blank={date:today(),shift:SHIFTS[0],operator:OPERATORS[0],machineStatus:"Running",boulderFed:"",production:[blankOut()],breakdowns:"",remarks:""};
  const [open,setOpen]=useState(false);const [f,setF]=useState(blank);const [exp,setExp]=useState(null);
  const set=k=>v=>setF(x=>({...x,[k]:v}));
  const setPr=(i,k,v)=>setF(x=>{const p=[...x.production];p[i]={...p[i],[k]:v};return{...x,production:p};});
  function save(){
    if(!f.boulderFed)return alert("Boulder fed required.");
    if(!f.production.every(p=>p.gradeId&&p.quantity))return alert("Fill all rows.");
    const enriched=f.production.map(p=>{const g=grades.find(g=>String(g.id)===String(p.gradeId));return{...p,gradeName:g?.name,gradeCode:g?.code,quantity:+p.quantity};});
    setShiftLogs(ls=>[{id:uid(),logNo:`SL-${String(ls.length+1).padStart(4,"0")}`,...f,boulderFed:+f.boulderFed,production:enriched},...ls]);
    setOpen(false);setF(blank);
  }
  return(
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New Shift</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Log #</th><th>Date</th><th>Shift</th><th>Operator</th><th>Status</th><th className="r">Fed (MT)</th><th className="r">Out (MT)</th><th></th></tr></thead>
          <tbody>
            {shiftLogs.length===0?<tr><td colSpan={8}><EmptyState icon="◷" message="No shift logs yet"/></td></tr>
            :shiftLogs.map(log=>(
              <>
              <tr key={log.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{log.logNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{log.date}</td>
                <td><Badge type="muted">{log.shift.split(" ")[0]}</Badge></td>
                <td style={{color:"var(--text2)"}}>{log.operator}</td>
                <td><Badge type={log.machineStatus==="Running"?"green":log.machineStatus==="Breakdown"?"red":"amber"}>{log.machineStatus}</Badge></td>
                <td className="r mono">{log.boulderFed}</td>
                <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{log.production.reduce((s,p)=>s+(p.quantity||0),0).toFixed(1)}</td>
                <td><button className="btn btn-ghost btn-sm" onClick={()=>setExp(exp===log.id?null:log.id)}>{exp===log.id?"▲":"▼"}</button></td>
              </tr>
              {exp===log.id&&<tr key={log.id+"d"}><td colSpan={8} style={{padding:0}}><div className="row-detail">{log.production.map((p,i)=><div key={i} className="row-detail-block"><label>[{p.gradeCode}]</label><p style={{color:"var(--teal)",fontWeight:600}}>{p.quantity} MT</p></div>)}{log.breakdowns&&<div className="row-detail-block"><label>Breakdowns</label><p>{log.breakdowns}</p></div>}{log.remarks&&<div className="row-detail-block"><label>Remarks</label><p>{log.remarks}</p></div>}</div></td></tr>}
              </>
            ))}
          </tbody>
        </table>
      </div>
      {open&&<Modal title="New Shift Log" onClose={()=>setOpen(false)} foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save</button></>}>
        <div className="form-row cols-2"><FG label="Date"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG><FG label="Shift"><select value={f.shift} onChange={e=>set("shift")(e.target.value)}>{SHIFTS.map(s=><option key={s}>{s}</option>)}</select></FG></div>
        <div className="form-row cols-3"><FG label="Operator"><select value={f.operator} onChange={e=>set("operator")(e.target.value)}>{OPERATORS.map(o=><option key={o}>{o}</option>)}</select></FG><FG label="Machine Status"><select value={f.machineStatus} onChange={e=>set("machineStatus")(e.target.value)}>{["Running","Breakdown","Maintenance","Idle"].map(s=><option key={s}>{s}</option>)}</select></FG><FG label="Boulder Fed (MT) *"><input type="number" placeholder="85.5" value={f.boulderFed} onChange={e=>set("boulderFed")(e.target.value)}/></FG></div>
        <div style={{marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}><label style={{fontSize:10.5,color:"var(--text2)",textTransform:"uppercase",letterSpacing:".05em",fontFamily:"var(--mono)"}}>Output by Grade</label><button className="btn btn-ghost btn-sm" onClick={()=>setF(x=>({...x,production:[...x.production,blankOut()]}))}>+ Add</button></div>
          <div className="output-rows">{f.production.map((row,i)=><div key={i} className="output-row"><div className="form-group"><select value={row.gradeId} onChange={e=>setPr(i,"gradeId",e.target.value)}><option value="">Grade…</option>{grades.map(g=><option key={g.id} value={g.id}>[{g.code}] {g.name}</option>)}</select></div><div className="form-group"><input type="number" placeholder="MT" value={row.quantity} onChange={e=>setPr(i,"quantity",e.target.value)}/></div><button className="btn btn-danger btn-sm" onClick={()=>setF(x=>({...x,production:x.production.filter((_,idx)=>idx!==i)}))}>✕</button></div>)}</div>
        </div>
        <div className="form-row cols-2"><FG label="Breakdowns"><input placeholder="e.g. 30 min jaw crusher blockage" value={f.breakdowns} onChange={e=>set("breakdowns")(e.target.value)}/></FG><FG label="Remarks"><input value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG></div>
      </Modal>}
    </div>
  );
}

function MasterPage({title,noun,icon,fields,items,setItems,cols}){
  const blank=Object.fromEntries(fields.map(f=>[f.key,f.default??""]));
  const [open,setOpen]=useState(false);const [editing,setEdit]=useState(null);const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));
  function openAdd(){setF(blank);setEdit(null);setOpen(true);}
  function openEdit(x){setF({...x});setEdit(x.id);setOpen(true);}
  function closeModal(){setOpen(false);setEdit(null);setF(blank);}
  function save(){const req=fields.find(fl=>fl.required&&!f[fl.key]);if(req)return alert(`${req.label} required.`);if(editing!==null)setItems(xs=>xs.map(x=>x.id===editing?{...f,id:editing}:x));else setItems(xs=>[...xs,{...f,id:uid()}]);closeModal();}
  function del(id){if(window.confirm(`Delete this ${noun}?`))setItems(xs=>xs.filter(x=>x.id!==id));}
  return(
    <div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={openAdd}>+ Add {noun}</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr>{cols.map(c=><th key={c.label} className={c.r?"r":""}>{c.label}</th>)}<th>Actions</th></tr></thead>
          <tbody>
            {items.length===0?<tr><td colSpan={cols.length+1}><EmptyState icon={icon||"📋"} message={`No ${noun.toLowerCase()}s yet`}/></td></tr>
            :items.map(row=>(
              <tr key={row.id}>
                {cols.map(c=><td key={c.label} className={c.r?"r":""}>{c.render?c.render(row):row[c.key]}</td>)}
                <td><div style={{display:"flex",gap:6}}><button className="btn btn-ghost btn-sm" onClick={()=>openEdit(row)}>Edit</button><button className="btn btn-danger btn-sm" onClick={()=>del(row.id)}>Del</button></div></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {open&&<Modal title={`${editing?"Edit":"Add"} ${noun}`} onClose={closeModal} foot={<><button className="btn btn-ghost" onClick={closeModal}>Cancel</button><button className="btn btn-primary" onClick={save}>{editing?"Update":"Save"}</button></>}>
        <div className="form-row cols-2">
          {fields.map(fl=>fl.type==="checkbox"
            ?<FG key={fl.key} span={2}><label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,color:"var(--text2)",cursor:"pointer",textTransform:"none",letterSpacing:"normal"}}><input type="checkbox" checked={!!f[fl.key]} onChange={e=>set(fl.key)(e.target.checked)} style={{width:"auto"}}/>{fl.label}</label></FG>
            :<FG key={fl.key} label={`${fl.label}${fl.required?" *":""}`} span={fl.full?2:undefined}>{fl.type==="select"?<select value={f[fl.key]} onChange={e=>set(fl.key)(e.target.value)}>{fl.options.map(o=><option key={o}>{o}</option>)}</select>:<input type={fl.type||"text"} placeholder={fl.placeholder||""} value={f[fl.key]} onChange={e=>set(fl.key)(fl.upper?e.target.value.toUpperCase():e.target.value)}/>}</FG>
          )}
        </div>
      </Modal>}
    </div>
  );
}

function CustomerMaster({customers,setCustomers}){
  return <MasterPage title="Customers" noun="Customer" icon="🏢" items={customers} setItems={setCustomers}
    fields={[{key:"name",label:"Company Name",required:true,placeholder:"Sharma Constructions"},{key:"gstin",label:"GSTIN",placeholder:"27AABCS1234A1Z5",upper:true},{key:"contact",label:"Contact Person",placeholder:"Anil Sharma"},{key:"phone",label:"Phone",placeholder:"9001234567"},{key:"state",label:"State",placeholder:"e.g. Telangana"},{key:"address",label:"Address",placeholder:"City, State",full:true}]}
    cols={[{label:"Name",render:r=><span style={{fontWeight:600}}>{r.name}</span>},{label:"GSTIN",render:r=><span className="mono" style={{color:"var(--text3)",fontSize:11}}>{r.gstin||"—"}</span>},{label:"Contact",key:"contact"},{label:"State",render:r=><span style={{color:"var(--text2)"}}>{r.state||"—"}</span>},{label:"Phone",render:r=><span style={{color:"var(--text3)"}}>{r.phone}</span>}]}/>;
}
function VehicleMaster({vehicles,setVehicles}){
  return <MasterPage title="Vehicles" noun="Vehicle" icon="🚛" items={vehicles} setItems={setVehicles}
    fields={[{key:"vehicleNo",label:"Vehicle Number",required:true,placeholder:"MH12AB1234",upper:true},{key:"type",label:"Type",type:"select",options:["Tipper","Dumper","Truck","Tractor","Other"],default:"Tipper"},{key:"tareWeight",label:"Tare Weight (kg)",type:"number",placeholder:"8200",required:true},{key:"owner",label:"Owner",placeholder:"Ramesh Transport"},{key:"phone",label:"Phone",placeholder:"9876543210"}]}
    cols={[{label:"Vehicle #",render:r=><span className="mono" style={{color:"var(--accent)",fontWeight:700}}>{r.vehicleNo}</span>},{label:"Type",render:r=><Badge type="muted">{r.type}</Badge>},{label:"Tare Wt",render:r=><span className="mono">{fmtWt(+r.tareWeight)}</span>,r:true},{label:"Owner",key:"owner"},{label:"Phone",render:r=><span style={{color:"var(--text3)"}}>{r.phone}</span>}]}/>;
}
function SupplierMaster({suppliers,setSuppliers}){
  return <MasterPage title="Suppliers" noun="Supplier" icon="⛏" items={suppliers} setItems={setSuppliers}
    fields={[{key:"name",label:"Mine / Supplier Name",required:true,placeholder:"Rajput Mines Pvt Ltd"},{key:"contact",label:"Contact Person",placeholder:"Deepak Rajput"},{key:"phone",label:"Phone",placeholder:"9811122233"},{key:"address",label:"Region",placeholder:"Rajasthan"},{key:"gstin",label:"GSTIN",placeholder:"08XXXXX",upper:true}]}
    cols={[{label:"Name",render:r=><span style={{fontWeight:600}}>{r.name}</span>},{label:"Contact",key:"contact"},{label:"Phone",render:r=><span style={{color:"var(--text2)"}}>{r.phone}</span>},{label:"Region",render:r=><span style={{color:"var(--text3)"}}>{r.address}</span>}]}/>;
}
function GradesMaster({grades,setGrades}){
  return <MasterPage title="Material Grades" noun="Grade" icon="💎" items={grades} setItems={setGrades}
    fields={[{key:"code",label:"Grade Code",required:true,placeholder:"Q20",upper:true},{key:"name",label:"Grade Name",required:true,placeholder:"Quartz Aggregate 20mm"},{key:"unit",label:"Unit",type:"select",options:["MT","KG","CFT","CUM"],default:"MT"},{key:"description",label:"Description",placeholder:"e.g. 20mm aggregate",full:true},{key:"isBoulder",label:"This is a Raw Material / Boulder grade",type:"checkbox",default:false}]}
    cols={[{label:"Code",render:r=><Badge type={r.isBoulder?"teal":"muted"}>{r.code}</Badge>},{label:"Name",render:r=><span style={{fontWeight:600}}>{r.name}</span>},{label:"Unit",render:r=><span style={{color:"var(--text3)"}}>{r.unit}</span>},{label:"Type",render:r=><Badge type={r.isBoulder?"teal":"blue"}>{r.isBoulder?"Raw Material":"Finished Grade"}</Badge>}]}/>;
}

// ── DASHBOARD ─────────────────────────────────────────────────────────────────
function Dashboard({vehicles,customers,grades,weighments,boulderReceipts,productionEntries,expenses,lots,qcTests,samples,maintTasks,maintLogs}){
  const stock=useMemo(()=>computeStock(grades,boulderReceipts,productionEntries,weighments),[grades,boulderReceipts,productionEntries,weighments]);
  const stockArr=Object.values(stock);
  const todayStr=today();
  const boulderStk=stockArr.find(s=>s.isBoulder)?.closing||0;
  const totalRM=boulderReceipts.reduce((s,r)=>s+r.quantityMT,0);
  const totalProd=productionEntries.reduce((s,p)=>s+p.totalOutputMT,0);
  const totalSoldMT=weighments.filter(w=>w.purpose==="Sale").reduce((s,w)=>s+(w.netWeight||0)/1000,0);
  const todayProdMT=productionEntries.filter(p=>p.date===todayStr).reduce((s,p)=>s+p.totalOutputMT,0);
  const todaySales=weighments.filter(w=>w.date===todayStr&&w.purpose==="Sale").length;
  const totalExpenses=expenses.reduce((s,e)=>s+e.amount,0);
  const avgYield=productionEntries.filter(p=>p.yieldPct).length
    ?((productionEntries.reduce((s,p)=>s+(p.yieldPct||0),0))/productionEntries.filter(p=>p.yieldPct).length).toFixed(1):null;
  const alerts=stockArr.filter(s=>s.closing<0);
  const lowStock=stockArr.filter(s=>!s.isBoulder&&s.closing>=0&&s.closing<10);
  const barColors=["var(--accent)","var(--blue)","var(--purple)","var(--teal)"];

  return(
    <div>
      <div className="stats-grid-5">
        <StatCard label="Boulder in Stock"  value={`${boulderStk.toFixed(1)} MT`}  sub="in yard"           color={boulderStk<20?"red":boulderStk<50?"amber":"teal"}/>
        <StatCard label="Total RM Received" value={`${totalRM.toFixed(1)} MT`}      sub={`${boulderReceipts.length} receipts`} color="blue"/>
        <StatCard label="Total Produced"    value={`${totalProd.toFixed(1)} MT`}    sub={`${productionEntries.length} entries`} color="accent"/>
        <StatCard label="Avg Mill Yield"    value={avgYield?`${avgYield}%`:"--"}    sub="RM to output"      color={avgYield>=90?"green":avgYield>=80?"amber":"muted"}/>
        <StatCard label="Today Dispatches"  value={todaySales}                       sub={`${todayProdMT.toFixed(1)} MT produced today`} color="green"/>
      </div>


      {(()=>{
        const todayStr2=today();
        const overdueM=(maintTasks||[]).filter(t=>t.active&&t.nextDue<todayStr2);
        const dueDisposal=(samples||[]).filter(s=>s.status==="retained"&&s.disposalDate<=addDays(todayStr2,3));
        const untestedLots=(lots||[]).filter(l=>l.status==="in-stock"&&!(qcTests||[]).some(t=>t.lotNo===l.lotNo));
        if(overdueM.length===0&&dueDisposal.length===0&&untestedLots.length===0)return null;
        return(
          <div className="alert-strip" style={{borderColor:"#1a2800",background:"#0a1000"}}>
            <h4>QC Alerts</h4>
            {overdueM.slice(0,3).map(t=><div key={t.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="red">Maint Overdue</Badge><span style={{fontSize:12,color:"var(--text2)"}}>{t.taskName} — {t.equipmentName} ({daysDiff(t.nextDue)*-1}d overdue)</span></div>)}
            {untestedLots.map(l=><div key={l.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="amber">No QC Test</Badge><span style={{fontSize:12,color:"var(--text2)"}}>Lot <span className="lot-chip" style={{fontSize:10}}>🏷 {l.lotNo}</span> has no test recorded</span></div>)}
            {dueDisposal.map(s=><div key={s.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="red">Dispose Sample</Badge><span style={{fontSize:12,color:"var(--text2)"}}>Sample {s.sampleNo} (lot {s.lotNo}) due {s.disposalDate}</span></div>)}
          </div>
        );
      })()}
      {(alerts.length>0||lowStock.length>0||boulderStk<20)&&(
        <div className="alert-strip">
          <h4>Alerts</h4>
          {boulderStk<20&&<div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="red">Low Raw Material</Badge><span style={{fontSize:12,color:"var(--text2)"}}>Only {fmtMT(boulderStk)} in yard — reorder soon</span></div>}
          {alerts.map(s=><div key={s.gradeId} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="red">Negative</Badge><span style={{fontSize:12,color:"var(--text2)"}}>{s.name} — {fmtMT(s.closing)}</span></div>)}
          {lowStock.map(s=><div key={s.gradeId} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}><Badge type="amber">Low</Badge><span style={{fontSize:12,color:"var(--text2)"}}>{s.name} — only {fmtMT(s.closing)} remaining</span></div>)}
        </div>
      )}

      <div className="two-col">
        <div className="table-wrap" style={{marginBottom:0}}>
          <div className="table-toolbar"><h3>Current Stock Position</h3></div>
          <table>
            <thead><tr><th>Grade</th><th>Type</th><th className="r">Closing Stock</th><th>Status</th></tr></thead>
            <tbody>
              {stockArr.length===0?<tr><td colSpan={4}><EmptyState icon="📦" message="No stock data yet"/></td></tr>
              :stockArr.map(s=>{const st=s.closing<0?"red":s.closing<10?"amber":"green";return(
                <tr key={s.gradeId}>
                  <td><Badge type={s.isBoulder?"teal":"muted"}>{s.code}</Badge> <span style={{marginLeft:6,fontSize:12,color:"var(--text2)"}}>{s.name}</span></td>
                  <td><Badge type={s.isBoulder?"teal":"blue"}>{s.isBoulder?"Raw":"Finished"}</Badge></td>
                  <td className="r mono" style={{color:st==="red"?"var(--red)":st==="amber"?"var(--amber)":"var(--accent)",fontWeight:700}}>{fmtMT(s.closing)}</td>
                  <td><Badge type={st}>{st==="red"?"Negative":st==="amber"?"Low":"OK"}</Badge></td>
                </tr>);})}
            </tbody>
          </table>
        </div>
        <div className="table-wrap" style={{marginBottom:0,padding:"16px 20px"}}>
          <div style={{fontWeight:600,fontSize:13,marginBottom:16}}>Grade-wise Production (All Time)</div>
          {stockArr.filter(s=>!s.isBoulder).length===0?<EmptyState icon="⚙" message="No production yet"/>
          :stockArr.filter(s=>!s.isBoulder).sort((a,b)=>b.produced-a.produced).map((s,i)=>{
            const max=Math.max(...stockArr.filter(x=>!x.isBoulder).map(x=>x.produced),1);
            return(<div key={s.gradeId} className="bar-row"><div className="bar-row-header"><span>{s.name}</span><span className="mono" style={{fontSize:11}}>{fmtMT(s.produced)}</span></div><div className="bar-track"><div className="bar-fill" style={{width:`${(s.produced/max)*100}%`,background:barColors[i%4]}}/></div></div>);
          })}
        </div>
      </div>

      <div className="two-col">
        <div className="table-wrap" style={{marginBottom:0}}>
          <div className="table-toolbar"><h3>Recent Weighments</h3></div>
          <table>
            <thead><tr><th>Slip #</th><th>Date</th><th>Vehicle</th><th>Grade</th><th className="r">Net Wt</th><th>Purpose</th></tr></thead>
            <tbody>
              {weighments.length===0?<tr><td colSpan={6}><EmptyState icon="⚖" message="No weighments yet"/></td></tr>
              :weighments.slice(0,6).map(w=>(
                <tr key={w.id}>
                  <td className="mono" style={{color:"var(--accent)"}}>{w.slipNo}</td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{w.date}</td>
                  <td className="mono" style={{fontWeight:600}}>{w.vehicleNo}</td>
                  <td><Badge type="muted">{w.gradeCode}</Badge></td>
                  <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtWt(w.netWeight)}</td>
                  <td><Badge type={w.purpose==="Sale"?"green":w.purpose==="Inward (Raw)"?"blue":"muted"}>{w.purpose}</Badge></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="table-wrap" style={{marginBottom:0}}>
          <div className="table-toolbar"><h3>Recent Production</h3></div>
          <table>
            <thead><tr><th>Entry #</th><th>Date</th><th>Shift</th><th className="r">RM In</th><th className="r">Output</th><th>Yield</th></tr></thead>
            <tbody>
              {productionEntries.length===0?<tr><td colSpan={6}><EmptyState icon="⚙" message="No production yet"/></td></tr>
              :productionEntries.slice(0,6).map(pe=>(
                <tr key={pe.id}>
                  <td className="mono" style={{color:"var(--accent)"}}>{pe.entryNo}</td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{pe.date}</td>
                  <td><Badge type="muted">{pe.shift.split(" ")[0]}</Badge></td>
                  <td className="r mono">{fmtMT(pe.boulderConsumedMT)}</td>
                  <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtMT(pe.totalOutputMT)}</td>
                  <td>{pe.yieldPct?<Badge type={pe.yieldPct>=90?"green":pe.yieldPct>=80?"amber":"red"}>{pe.yieldPct}%</Badge>:"--"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ── EXPENSE CATEGORIES ────────────────────────────────────────────────────────
const EXP_CATS=[
  {id:"fuel",     label:"Fuel & Diesel",           color:"#f5a623"},
  {id:"salary",   label:"Staff Salaries",          color:"#5ca0ff"},
  {id:"labour",   label:"Contract / Casual Labour",color:"#7aa8ff"},
  {id:"maint",    label:"Maintenance",             color:"#a78bfa"},
  {id:"royalty",  label:"Royalty / Levy",          color:"#2dd4bf"},
  {id:"electric", label:"Electricity",             color:"#4ecb71"},
  {id:"water",    label:"Water",                   color:"#3fa9d6"},
  {id:"transport",label:"Transport",               color:"#ff5c5c"},
  {id:"misc",     label:"Miscellaneous",           color:"#999"},
];
// Categories the Monthly Ops Costs quick-entry screen manages as one running total per month.
// Each auto-entry is tagged opsAuto:true so re-saving a month updates it instead of duplicating.
const OPS_CATS=["electric","water","salary","labour","fuel","maint","transport","misc"];

// ── EXPENSES VIEW ─────────────────────────────────────────────────────────────
function ExpensesView({expenses,setExpenses}){
  const [open,setOpen]=useState(false);
  const [monthFilter,setMonthFilter]=useState(today().slice(0,7));
  const blank={date:today(),category:"fuel",amount:"",description:"",paidTo:"",reference:""};
  const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));

  function save(){
    if(!f.amount||!f.description)return alert("Amount and Description required.");
    setExpenses(es=>[{id:uid(),expNo:`EXP-${String(es.length+1).padStart(4,"0")}`,...f,amount:+f.amount},...es]);
    setOpen(false);setF(blank);
  }

  const filtered=expenses.filter(e=>!monthFilter||e.date?.slice(0,7)===monthFilter);
  const totalMonth=filtered.reduce((s,e)=>s+e.amount,0);
  const totalAll=expenses.reduce((s,e)=>s+e.amount,0);

  // Category breakdown for month
  const catBreak=EXP_CATS.map(c=>({...c,total:filtered.filter(e=>e.category===c.id).reduce((s,e)=>s+e.amount,0)})).filter(c=>c.total>0).sort((a,b)=>b.total-a.total);
  const maxCat=catBreak[0]?.total||1;

  const months=[...new Set(expenses.map(e=>e.date?.slice(0,7)))].sort((a,b)=>b.localeCompare(a));

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="This Month"   value={fmt(totalMonth)} sub={`${filtered.length} entries`}   color="red"/>
        <StatCard label="All Time"     value={fmt(totalAll)}   sub={`${expenses.length} entries`}   color="amber"/>
        <StatCard label="Avg / Month"  value={months.length?fmt(totalAll/months.length):"₹0"} sub="monthly avg" color="blue"/>
        <StatCard label="Top Category" value={catBreak[0]?.label||"—"} sub={catBreak[0]?fmt(catBreak[0].total):"no data"} color="purple"/>
      </div>

      <div className="two-col" style={{marginBottom:20}}>
        {/* Category Breakdown */}
        <div className="report-section" style={{marginBottom:0}}>
          <h3>Breakdown by Category {monthFilter&&`— ${monthFilter}`}</h3>
          {catBreak.length===0
            ?<EmptyState icon="📂" message="No expenses this month"/>
            :catBreak.map(c=>(
              <div key={c.id} className="mini-bar-row">
                <span className="label"><span className="expense-cat-dot" style={{background:c.color}}/>{c.label}</span>
                <div className="track"><div className="fill" style={{width:`${(c.total/maxCat)*100}%`,background:c.color}}/></div>
                <span className="val">{fmt(c.total)}</span>
              </div>
            ))
          }
        </div>
        {/* Monthly trend */}
        <div className="report-section" style={{marginBottom:0}}>
          <h3>Monthly Trend (Last 6 Months)</h3>
          {months.length===0
            ?<EmptyState icon="📈" message="No expense data yet"/>
            :months.slice(0,6).reverse().map(m=>{
              const mTotal=expenses.filter(e=>e.date?.slice(0,7)===m).reduce((s,e)=>s+e.amount,0);
              const maxM=Math.max(...months.slice(0,6).map(mm=>expenses.filter(e=>e.date?.slice(0,7)===mm).reduce((s,e)=>s+e.amount,0)),1);
              return(
                <div key={m} className="mini-bar-row">
                  <span className="label">{m}</span>
                  <div className="track"><div className="fill" style={{width:`${(mTotal/maxM)*100}%`,background:"var(--amber)"}}/></div>
                  <span className="val">{fmt(mTotal)}</span>
                </div>
              );
            })
          }
        </div>
      </div>

      <div className="filter-bar">
        <input type="month" value={monthFilter} onChange={e=>setMonthFilter(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"7px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12.5,outline:"none"}}/>
        <button className="btn btn-ghost btn-sm" onClick={()=>setMonthFilter("")}>All Time</button>
        <button className="btn btn-primary" onClick={()=>setOpen(true)}>+ Add Expense</button>
      </div>

      <div className="table-wrap">
        <table>
          <thead><tr><th>Exp #</th><th>Date</th><th>Category</th><th>Description</th><th>Paid To</th><th>Ref #</th><th className="r">Amount</th><th></th></tr></thead>
          <tbody>
            {filtered.length===0
              ?<tr><td colSpan={8}><EmptyState icon="💸" message="No expenses" sub={monthFilter?"Try clearing the month filter":"Add your first expense"}/></td></tr>
              :filtered.map(e=>{
                const cat=EXP_CATS.find(c=>c.id===e.category);
                return(
                  <tr key={e.id}>
                    <td className="mono" style={{color:"var(--accent)"}}>{e.expNo}</td>
                    <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{e.date}</td>
                    <td><span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:12}}><span className="expense-cat-dot" style={{background:cat?.color}}/>{cat?.label}</span></td>
                    <td style={{fontWeight:500}}>{e.description}{e.opsAuto&&<span style={{marginLeft:6}}><Badge type="blue">monthly total</Badge></span>}</td>
                    <td style={{color:"var(--text2)"}}>{e.paidTo||"—"}</td>
                    <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{e.reference||"—"}</td>
                    <td className="r mono" style={{color:"var(--red)",fontWeight:700}}>{fmt(e.amount)}</td>
                    <td><button className="btn btn-danger btn-sm" onClick={()=>setExpenses(es=>es.filter(x=>x.id!==e.id))}>Del</button></td>
                  </tr>
                );
              })
            }
          </tbody>
        </table>
      </div>

      {open&&(
        <Modal title="Add Expense" onClose={()=>setOpen(false)}
          foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Expense</button></>}>
          <div className="form-row cols-2">
            <FG label="Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
            <FG label="Category">
              <select value={f.category} onChange={e=>set("category")(e.target.value)}>
                {EXP_CATS.map(c=><option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Amount (₹) *"><input type="number" step="0.01" placeholder="5000" value={f.amount} onChange={e=>set("amount")(e.target.value)}/></FG>
            <FG label="Paid To"><input placeholder="Vendor / Person" value={f.paidTo} onChange={e=>set("paidTo")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Description *"><input placeholder="e.g. Diesel for crusher" value={f.description} onChange={e=>set("description")(e.target.value)}/></FG>
            <FG label="Reference / Bill #"><input placeholder="Bill no." value={f.reference} onChange={e=>set("reference")(e.target.value)}/></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── DAILY REPORT ──────────────────────────────────────────────────────────────
function DailyReport({weighments,productionEntries,boulderReceipts,expenses,grades}){
  const [date,setDate]=useState(today());

  const dayW    = weighments.filter(w=>w.date===date);
  const dayPE   = productionEntries.filter(p=>p.date===date);
  const dayBR   = boulderReceipts.filter(r=>r.date===date);
  const dayExp  = expenses.filter(e=>e.date===date);

  const salesW     = dayW.filter(w=>w.purpose==="Sale");
  const inwardW    = dayW.filter(w=>w.purpose==="Inward (Raw)");
  const netSaleMT  = salesW.reduce((s,w)=>s+(w.netWeight||0)/1000,0);
  const netInwardMT= inwardW.reduce((s,w)=>s+(w.netWeight||0)/1000,0);
  const totalProdMT= dayPE.reduce((s,p)=>s+p.totalOutputMT,0);
  const boulderIn  = dayBR.reduce((s,r)=>s+r.quantityMT,0);
  const totalExpAmt= dayExp.reduce((s,e)=>s+e.amount,0);
  const avgYield   = dayPE.filter(p=>p.yieldPct).length
    ?(dayPE.reduce((s,p)=>s+(p.yieldPct||0),0)/dayPE.filter(p=>p.yieldPct).length).toFixed(1):null;

  // Grade-wise production for the day
  const gradeMap={};
  dayPE.forEach(pe=>pe.outputs.forEach(o=>{gradeMap[o.gradeId]=(gradeMap[o.gradeId]||0)+o.quantityMT;}));

  // Grade-wise sales
  const salesMap={};
  salesW.forEach(w=>{salesMap[w.gradeId]=(salesMap[w.gradeId]||0)+(w.netWeight||0)/1000;});

  return(
    <div>
      <div className="report-controls">
        <div className="form-group"><label style={{fontSize:10.5,color:"var(--text2)",textTransform:"uppercase",letterSpacing:".05em",fontFamily:"var(--mono)"}}>Report Date</label>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"8px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:13,outline:"none"}}/>
        </div>
        <button className="btn btn-ghost btn-sm print-hide" onClick={()=>window.print()}>🖨 Print</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Boulder Received"  value={`${boulderIn.toFixed(2)} MT`}   sub={`${dayBR.length} deliveries`}  color="teal"/>
        <StatCard label="Total Produced"    value={`${totalProdMT.toFixed(2)} MT`} sub={`${dayPE.length} shifts`}       color="blue"/>
        <StatCard label="Sales Dispatched"  value={`${netSaleMT.toFixed(2)} MT`}   sub={`${salesW.length} loads`}       color="accent"/>
        <StatCard label="Expenses"          value={fmt(totalExpAmt)}               sub={`${dayExp.length} entries`}     color="red"/>
      </div>

      <div className="two-col">
        {/* Production Detail */}
        <div className="report-section" style={{marginBottom:0}}>
          <h3>Production by Grade</h3>
          {Object.keys(gradeMap).length===0
            ?<EmptyState icon="⚙" message="No production entries for this date"/>
            :Object.entries(gradeMap).map(([gid,qty])=>{
              const g=grades.find(g=>String(g.id)===gid);
              const maxQ=Math.max(...Object.values(gradeMap),1);
              return(
                <div key={gid} className="mini-bar-row">
                  <span className="label"><Badge type="muted">{g?.code||gid}</Badge></span>
                  <div className="track"><div className="fill" style={{width:`${(qty/maxQ)*100}%`,background:"var(--blue)"}}/></div>
                  <span className="val">{fmtMT(qty)}</span>
                </div>
              );
            })
          }
          {avgYield&&<div style={{marginTop:12,padding:"8px 12px",background:"var(--bg3)",borderRadius:"var(--r)",display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:"var(--text3)"}}>Avg Yield</span><span style={{fontFamily:"var(--mono)",color:"var(--accent)",fontWeight:700}}>{avgYield}%</span></div>}
        </div>

        {/* Sales Detail */}
        <div className="report-section" style={{marginBottom:0}}>
          <h3>Sales by Grade</h3>
          {Object.keys(salesMap).length===0
            ?<EmptyState icon="📤" message="No sales on this date"/>
            :Object.entries(salesMap).map(([gid,qty])=>{
              const g=grades.find(g=>String(g.id)===gid);
              const maxQ=Math.max(...Object.values(salesMap),1);
              return(
                <div key={gid} className="mini-bar-row">
                  <span className="label"><Badge type="muted">{g?.code||gid}</Badge> {g?.name}</span>
                  <div className="track"><div className="fill" style={{width:`${(qty/maxQ)*100}%`,background:"var(--accent)"}}/></div>
                  <span className="val">{fmtMT(qty)}</span>
                </div>
              );
            })
          }
        </div>
      </div>

      {/* Shift Log Table */}
      {dayPE.length>0&&(
        <div className="table-wrap">
          <div className="table-toolbar"><h3>Shift-wise Production</h3></div>
          <table>
            <thead><tr><th>Entry #</th><th>Shift</th><th>Operator</th><th className="r">Boulder In</th><th className="r">Total Out</th><th>Yield</th></tr></thead>
            <tbody>
              {dayPE.map(pe=>(
                <tr key={pe.id}>
                  <td className="mono" style={{color:"var(--accent)"}}>{pe.entryNo}</td>
                  <td><Badge type="muted">{pe.shift.split(" ")[0]}</Badge></td>
                  <td style={{color:"var(--text2)"}}>{pe.operator}</td>
                  <td className="r mono">{fmtMT(pe.boulderConsumedMT)}</td>
                  <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtMT(pe.totalOutputMT)}</td>
                  <td>{pe.yieldPct?<Badge type={pe.yieldPct>=90?"green":pe.yieldPct>=75?"amber":"red"}>{pe.yieldPct}%</Badge>:"—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Weighment Table */}
      {dayW.length>0&&(
        <div className="table-wrap">
          <div className="table-toolbar"><h3>Weighment Slips — {date}</h3></div>
          <table>
            <thead><tr><th>Slip #</th><th>Time</th><th>Vehicle</th><th>Customer</th><th>Grade</th><th className="r">Net Wt</th><th>Purpose</th></tr></thead>
          <tbody>
            {dayW.map(w=>(
              <tr key={w.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{w.slipNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{w.time}</td>
                <td className="mono" style={{fontWeight:600}}>{w.vehicleNo}</td>
                <td style={{color:"var(--text2)"}}>{w.customerName}</td>
                <td><Badge type="muted">{w.gradeCode}</Badge></td>
                <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtWt(w.netWeight)}</td>
                <td><Badge type={w.purpose==="Sale"?"green":w.purpose==="Inward (Raw)"?"blue":"muted"}>{w.purpose}</Badge></td>
              </tr>
            ))}
          </tbody>
          </table>
        </div>
      )}

      {/* Expenses Table */}
      {dayExp.length>0&&(
        <div className="table-wrap">
          <div className="table-toolbar"><h3>Expenses — {date}</h3></div>
          <table>
            <thead><tr><th>Category</th><th>Description</th><th>Paid To</th><th className="r">Amount</th></tr></thead>
            <tbody>
              {dayExp.map(e=>{
                const cat=EXP_CATS.find(c=>c.id===e.category);
                return(
                  <tr key={e.id}>
                    <td><span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:12}}><span className="expense-cat-dot" style={{background:cat?.color}}/>{cat?.label}</span></td>
                    <td style={{fontWeight:500}}>{e.description}</td>
                    <td style={{color:"var(--text2)"}}>{e.paidTo||"—"}</td>
                    <td className="r mono" style={{color:"var(--red)",fontWeight:700}}>{fmt(e.amount)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {dayW.length===0&&dayPE.length===0&&dayBR.length===0&&dayExp.length===0&&(
        <div className="report-section"><EmptyState icon="📅" message={`No data for ${date}`} sub="Select a date that has activity"/></div>
      )}
    </div>
  );
}

// ── STOCK REPORT ──────────────────────────────────────────────────────────────
function StockReport({grades,boulderReceipts,productionEntries,weighments}){
  const [asOf,setAsOf]=useState(today());

  // Filter all data up to asOf date
  const filtBR   = boulderReceipts.filter(r=>r.date<=asOf);
  const filtPE   = productionEntries.filter(p=>p.date<=asOf);
  const filtW    = weighments.filter(w=>w.date<=asOf);
  const stock    = useMemo(()=>computeStock(grades,filtBR,filtPE,filtW),[grades,filtBR,filtPE,filtW,asOf]);
  const stockArr = Object.values(stock);

  const totalBoulderIn  = filtBR.reduce((s,r)=>s+r.quantityMT,0);
  const totalProduced   = filtPE.reduce((s,p)=>s+p.totalOutputMT,0);
  const totalSold       = filtW.filter(w=>w.purpose==="Sale").reduce((s,w)=>s+(w.netWeight||0)/1000,0);
  const boulderStock    = stockArr.find(s=>s.isBoulder)?.closing||0;
  const avgYield        = filtPE.filter(p=>p.yieldPct).length
    ?(filtPE.reduce((s,p)=>s+(p.yieldPct||0),0)/filtPE.filter(p=>p.yieldPct).length).toFixed(1):null;

  // Grade-wise summary with all movements
  return(
    <div>
      <div className="report-controls">
        <div className="form-group">
          <label style={{fontSize:10.5,color:"var(--text2)",textTransform:"uppercase",letterSpacing:".05em",fontFamily:"var(--mono)"}}>Stock As Of</label>
          <input type="date" value={asOf} onChange={e=>setAsOf(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"8px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:13,outline:"none"}}/>
        </div>
        <button className="btn btn-ghost btn-sm print-hide" onClick={()=>window.print()}>🖨 Print</button>
      </div>

      <div className="stats-grid">
        <StatCard label="Boulder Received"  value={`${totalBoulderIn.toFixed(2)} MT`}  sub={`${filtBR.length} receipts`}  color="teal"/>
        <StatCard label="Total Produced"    value={`${totalProduced.toFixed(2)} MT`}   sub={`${filtPE.length} entries`}   color="blue"/>
        <StatCard label="Total Sold"        value={`${totalSold.toFixed(2)} MT`}       sub={`${filtW.filter(w=>w.purpose==="Sale").length} loads`} color="accent"/>
        <StatCard label="Avg Yield"         value={avgYield?`${avgYield}%`:"—"}        sub="boulder→output"               color="green"/>
      </div>

      <div className="table-wrap">
        <div className="table-toolbar"><h3>Stock Position as of {asOf}</h3></div>
        <table>
          <thead>
            <tr>
              <th>Grade</th><th>Type</th>
              <th className="r">Received / Produced (MT)</th>
              <th className="r">Sold / Consumed (MT)</th>
              <th className="r">Transferred (MT)</th>
              <th className="r">Closing Stock (MT)</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {stockArr.length===0
              ?<tr><td colSpan={7}><EmptyState icon="📦" message="No stock data"/></td></tr>
              :stockArr.map(s=>{
                const st=s.closing<0?"red":s.closing<10?"amber":"green";
                return(
                  <tr key={s.gradeId}>
                    <td><Badge type={s.isBoulder?"teal":"muted"}>{s.code}</Badge> <span style={{marginLeft:6,fontSize:12,color:"var(--text2)"}}>{s.name}</span></td>
                    <td><Badge type={s.isBoulder?"teal":"blue"}>{s.isBoulder?"Raw":"Finished"}</Badge></td>
                    <td className="r mono" style={{color:"var(--teal)"}}>{fmtMT(s.isBoulder?s.in:s.produced)}</td>
                    <td className="r mono" style={{color:"var(--amber)"}}>{fmtMT(s.isBoulder?Math.abs(s.produced):s.sold)}</td>
                    <td className="r mono" style={{color:"var(--text3)"}}>{fmtMT(s.transferred)}</td>
                    <td className="r mono" style={{color:st==="red"?"var(--red)":st==="amber"?"var(--amber)":"var(--accent)",fontWeight:700,fontSize:13}}>{fmtMT(s.closing)}</td>
                    <td><Badge type={st}>{st==="red"?"⚠ Negative":st==="amber"?"Low Stock":"OK"}</Badge></td>
                  </tr>
                );
              })
            }
          </tbody>
        </table>
      </div>

      {/* Movement Summary per grade */}
      <div className="report-section">
        <h3>Grade-wise Movement Summary — up to {asOf}</h3>
        {stockArr.filter(s=>!s.isBoulder).length===0
          ?<EmptyState icon="⚙" message="No finished goods yet"/>
          :<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:12}}>
            {stockArr.filter(s=>!s.isBoulder).map(s=>(
              <div key={s.gradeId} className="summary-card">
                <div className="s-label"><Badge type="muted">{s.code}</Badge> {s.name}</div>
                <div style={{display:"flex",flexDirection:"column",gap:6,marginTop:8,fontSize:12,fontFamily:"var(--mono)"}}>
                  <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--text3)"}}>Produced</span><span style={{color:"var(--teal)"}}>{fmtMT(s.produced)}</span></div>
                  <div style={{display:"flex",justifyContent:"space-between"}}><span style={{color:"var(--text3)"}}>Sold</span><span style={{color:"var(--amber)"}}>{fmtMT(s.sold)}</span></div>
                  <div style={{display:"flex",justifyContent:"space-between",borderTop:"1px solid var(--border)",paddingTop:5}}><span style={{color:"var(--text2)",fontWeight:600}}>In Stock</span><span style={{color:s.closing<0?"var(--red)":s.closing<10?"var(--amber)":"var(--accent)",fontWeight:700}}>{fmtMT(s.closing)}</span></div>
                </div>
              </div>
            ))}
          </div>
        }
      </div>

      {/* Recent movements */}
      {weighments.filter(w=>w.date<=asOf&&w.purpose==="Sale").length>0&&(
        <div className="table-wrap">
          <div className="table-toolbar"><h3>Recent Sale Dispatches (up to {asOf})</h3></div>
          <table>
            <thead><tr><th>Date</th><th>Slip #</th><th>Vehicle</th><th>Customer</th><th>Grade</th><th className="r">Net Wt</th></tr></thead>
            <tbody>
              {weighments.filter(w=>w.date<=asOf&&w.purpose==="Sale").slice(0,15).map(w=>(
                <tr key={w.id}>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{w.date}</td>
                  <td className="mono" style={{color:"var(--accent)"}}>{w.slipNo}</td>
                  <td className="mono" style={{fontWeight:600}}>{w.vehicleNo}</td>
                  <td style={{color:"var(--text2)"}}>{w.customerName}</td>
                  <td><Badge type="muted">{w.gradeCode}</Badge></td>
                  <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{fmtWt(w.netWeight)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}





// ── EQUIPMENT REGISTER ────────────────────────────────────────────────────────
const EQUIP_TYPES=["Ball Mill","Classifier","Magnetic Separator","Jaw Crusher","Vibrating Screen","Air Compressor","Bucket Elevator","Conveyor Belt","Other"];

function EquipmentRegister({equipment,setEquipment,maintenanceLogs}){
  const [open,setOpen]=useState(false);const [editing,setEdit]=useState(null);
  const blank={code:"",name:"",type:EQUIP_TYPES[0],capacity:"",make:"",installDate:"",status:"active",notes:""};
  const [f,setF]=useState(blank);const set=k=>v=>setF(x=>({...x,[k]:v}));
  function openAdd(){setF(blank);setEdit(null);setOpen(true);}
  function openEdit(e){setF({...e});setEdit(e.id);setOpen(true);}
  function close(){setOpen(false);setEdit(null);setF(blank);}
  function save(){
    if(!f.code||!f.name)return alert("Code and Name required.");
    if(editing!==null)setEquipment(es=>es.map(e=>e.id===editing?{...f,id:editing}:e));
    else setEquipment(es=>[...es,{...f,id:uid()}]);
    close();
  }
  function del(id){if(window.confirm("Delete equipment?"))setEquipment(es=>es.filter(e=>e.id!==id));}

  return(
    <div>
      <div className="stats-grid-3">
        <StatCard label="Total Equipment" value={equipment.length}                                        sub="registered"      color="blue"/>
        <StatCard label="Active"          value={equipment.filter(e=>e.status==="active").length}         sub="running"         color="green"/>
        <StatCard label="Under Maint."    value={equipment.filter(e=>e.status==="maintenance").length}    sub="offline"         color="amber"/>
      </div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={openAdd}>+ Add Equipment</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Code</th><th>Name</th><th>Type</th><th>Capacity</th><th>Make</th><th>Installed</th><th>Maint Logs</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            {equipment.length===0?<tr><td colSpan={9}><EmptyState icon="🔩" message="No equipment registered"/></td></tr>
            :equipment.map(eq=>{
              const mCount=(maintenanceLogs||[]).filter(m=>m.equipmentId===eq.id).length;
              const lastBD=(maintenanceLogs||[]).filter(m=>m.equipmentId===eq.id&&m.type==="Breakdown").sort((a,b)=>b.date.localeCompare(a.date))[0];
              return(
                <tr key={eq.id}>
                  <td className="mono" style={{color:"var(--accent)",fontWeight:700}}>{eq.code}</td>
                  <td style={{fontWeight:600}}>{eq.name}</td>
                  <td><Badge type="muted">{eq.type}</Badge></td>
                  <td style={{color:"var(--text2)"}}>{eq.capacity||"--"}</td>
                  <td style={{color:"var(--text3)"}}>{eq.make||"--"}</td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{eq.installDate||"--"}</td>
                  <td style={{fontSize:12,color:"var(--text2)"}}>{mCount} {lastBD&&<span style={{color:"var(--text4)",fontSize:10,fontFamily:"var(--mono)"}}>· last BD: {lastBD.date}</span>}</td>
                  <td><Badge type={eq.status==="active"?"green":eq.status==="maintenance"?"amber":"red"}>{eq.status}</Badge></td>
                  <td>
                    <div style={{display:"flex",gap:6}}>
                      <button className="btn btn-ghost btn-sm" onClick={()=>openEdit(eq)}>Edit</button>
                      <button className="btn btn-danger btn-sm" onClick={()=>del(eq.id)}>Del</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {open&&(
        <Modal title={editing?"Edit Equipment":"Add Equipment"} onClose={close}
          foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={close}>Cancel</button><button className="btn btn-primary" onClick={save}>{editing?"Update":"Save"}</button></div>}>
          <div className="form-row cols-2">
            <FG label="Equipment Code *"><input placeholder="BM-01" value={f.code} onChange={e=>set("code")(e.target.value.toUpperCase())}/></FG>
            <FG label="Name *"><input placeholder="Ball Mill #1" value={f.name} onChange={e=>set("name")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Type"><select value={f.type} onChange={e=>set("type")(e.target.value)}>{EQUIP_TYPES.map(t=><option key={t}>{t}</option>)}</select></FG>
            <FG label="Status"><select value={f.status} onChange={e=>set("status")(e.target.value)}>{["active","maintenance","breakdown","decommissioned"].map(s=><option key={s}>{s}</option>)}</select></FG>
          </div>
          <div className="form-row cols-3">
            <FG label="Capacity"><input placeholder="5 TPH" value={f.capacity} onChange={e=>set("capacity")(e.target.value)}/></FG>
            <FG label="Make"><input placeholder="Laxmi Engineers" value={f.make} onChange={e=>set("make")(e.target.value)}/></FG>
            <FG label="Install Date"><input type="date" value={f.installDate} onChange={e=>set("installDate")(e.target.value)}/></FG>
          </div>
          <div className="form-row"><FG label="Notes"><input placeholder="Serial number, location, any details" value={f.notes} onChange={e=>set("notes")(e.target.value)}/></FG></div>
        </Modal>
      )}
    </div>
  );
}

// ── MAINTENANCE DASHBOARD ─────────────────────────────────────────────────────
function MaintDashboard({maintTasks,maintLogs,equipment}){
  const todayStr=today();
  const overdue   = maintTasks.filter(t=>t.active&&t.nextDue<todayStr);
  const dueToday  = maintTasks.filter(t=>t.active&&t.nextDue===todayStr);
  const dueSoon   = maintTasks.filter(t=>t.active&&t.nextDue>todayStr&&daysDiff(t.nextDue)<=7);
  const last30days= maintLogs.filter(l=>l.date>=addDays(todayStr,-30));
  const totalCost = last30days.reduce((s,l)=>s+(+l.cost||0),0);
  const bdCount   = last30days.filter(l=>l.type==="Breakdown").length;
  const bdDowntime= last30days.filter(l=>l.type==="Breakdown").reduce((s,l)=>s+(+l.durationHours||0),0);

  function dueBadge(task){
    const d=daysDiff(task.nextDue);
    if(d<0)   return {cls:"overdue",  label:`${Math.abs(d)}d overdue`};
    if(d===0) return {cls:"due-today",label:"Due today"};
    if(d<=7)  return {cls:"due-soon", label:`Due in ${d}d`};
    return      {cls:"ok",           label:`${d}d left`};
  }
  const catColors={cleaning:"var(--teal)",inspection:"var(--blue)",lubrication:"var(--amber)",service:"var(--purple)",replacement:"var(--red)"};

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Overdue Tasks"   value={overdue.length}   sub="action needed now"           color={overdue.length>0?"red":"green"}/>
        <StatCard label="Due Today"       value={dueToday.length}  sub="scheduled for today"          color={dueToday.length>0?"amber":"muted"}/>
        <StatCard label="Due This Week"   value={dueSoon.length}   sub="next 7 days"                 color="blue"/>
        <StatCard label="Maint Cost (30d)"value={`Rs.${totalCost.toLocaleString("en-IN")}`} sub={`${last30days.length} activities`} color="purple"/>
      </div>

      {overdue.length>0&&(
        <div className="alert-strip">
          <h4>Overdue Maintenance ({overdue.length})</h4>
          {overdue.map(t=>(
            <div key={t.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span className="due-badge overdue">{daysDiff(t.nextDue)*-1}d overdue</span>
                <span style={{fontSize:12,color:"var(--text2)"}}><span style={{fontWeight:600}}>{t.taskName}</span> — {t.equipmentName} ({t.code})</span>
              </div>
              <span style={{fontSize:11,color:"var(--text3)",fontFamily:"var(--mono)"}}>Was due: {t.nextDue}</span>
            </div>
          ))}
        </div>
      )}

      <div className="maint-grid">
        {/* Due Today & Soon */}
        <div>
          <div style={{fontWeight:600,fontSize:13,marginBottom:12}}>Today & This Week</div>
          {dueToday.length===0&&dueSoon.length===0
            ?<div style={{color:"var(--text4)",fontSize:12,padding:"20px 0"}}>No tasks due in next 7 days</div>
            :[...dueToday,...dueSoon].map(t=>{
              const b=dueBadge(t);
              return(
                <div key={t.id} className={`maint-card ${b.cls==="overdue"?"overdue":b.cls==="due-today"?"due-soon":"ok"}`}>
                  <div className="task-header">
                    <div>
                      <div className="task-name">{t.taskName}</div>
                      <div className="task-meta">{t.equipmentName} · {t.code}</div>
                    </div>
                    <span className={`due-badge ${b.cls}`}>{b.label}</span>
                  </div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:4}}>
                    <span className="interval-chip">🔄 {t.intervalValue} {t.intervalType}</span>
                    <span className="interval-chip">⏱ {t.estimatedMins} min</span>
                    <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:catColors[t.category]||"var(--text3)",margin:"auto 0"}}/>
                    <span style={{fontSize:10,color:"var(--text3)"}}>{t.category}</span>
                  </div>
                  {t.notes&&<div style={{fontSize:11,color:"var(--text4)",marginTop:6,fontStyle:"italic"}}>{t.notes}</div>}
                </div>
              );
            })
          }
        </div>

        {/* Recent Logs */}
        <div>
          <div style={{fontWeight:600,fontSize:13,marginBottom:12}}>Recent Activity (30 days)</div>
          {last30days.length===0
            ?<div style={{color:"var(--text4)",fontSize:12,padding:"20px 0"}}>No maintenance logged in last 30 days</div>
            :last30days.slice(0,8).map(l=>(
              <div key={l.id} className="maint-card ok" style={{padding:"10px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div>
                    <div style={{fontSize:12,fontWeight:600}}>{l.taskName}</div>
                    <div style={{fontSize:10,color:"var(--text3)",fontFamily:"var(--mono)",marginTop:2}}>{l.equipmentCode} · {l.date} · {l.performedBy||"--"}</div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <Badge type={l.type==="Breakdown"?"red":l.type==="Preventive"?"green":"muted"}>{l.type}</Badge>
                    {l.cost>0&&<div style={{fontSize:10,color:"var(--amber)",fontFamily:"var(--mono)",marginTop:3}}>Rs.{l.cost.toLocaleString("en-IN")}</div>}
                  </div>
                </div>
              </div>
            ))
          }
          {last30days.length>0&&(
            <div style={{padding:"10px 14px",background:"var(--bg3)",borderRadius:"var(--r)",marginTop:8,display:"flex",justifyContent:"space-between",fontSize:12}}>
              <span style={{color:"var(--text3)"}}>30-day: {bdCount} breakdown{bdCount!==1?"s":""}, {bdDowntime.toFixed(1)}h downtime</span>
              <span style={{color:"var(--red)",fontFamily:"var(--mono)"}}>Rs.{totalCost.toLocaleString("en-IN")}</span>
            </div>
          )}
        </div>
      </div>

      {/* Equipment Status Overview */}
      <div className="table-wrap">
        <div className="table-toolbar"><h3>Equipment Status</h3></div>
        <table>
          <thead><tr><th>Code</th><th>Equipment</th><th>Type</th><th>Tasks Due</th><th>Overdue</th><th>Last Activity</th><th>Status</th></tr></thead>
          <tbody>
            {equipment.map(eq=>{
              const eqTasks=maintTasks.filter(t=>t.equipmentId===eq.id&&t.active);
              const eqOverdue=eqTasks.filter(t=>t.nextDue<todayStr).length;
              const eqDueSoon=eqTasks.filter(t=>t.nextDue>=todayStr&&daysDiff(t.nextDue)<=7).length;
              const lastLog=maintLogs.filter(l=>l.equipmentId===eq.id).sort((a,b)=>b.date.localeCompare(a.date))[0];
              return(
                <tr key={eq.id}>
                  <td className="mono" style={{color:"var(--accent)",fontWeight:700}}>{eq.code}</td>
                  <td style={{fontWeight:500}}>{eq.name}</td>
                  <td><Badge type="muted">{eq.type}</Badge></td>
                  <td style={{color:eqDueSoon>0?"var(--amber)":"var(--text3)"}}>{eqDueSoon} task{eqDueSoon!==1?"s":""} this week</td>
                  <td style={{color:eqOverdue>0?"var(--red)":"var(--text3)",fontWeight:eqOverdue>0?700:400}}>{eqOverdue>0?eqOverdue+" overdue":"--"}</td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{lastLog?.date||"No log"}</td>
                  <td><Badge type={eq.status==="active"?"green":eq.status==="maintenance"?"amber":"red"}>{eq.status}</Badge></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── TASK SCHEDULE ─────────────────────────────────────────────────────────────
const TASK_CATEGORIES=["cleaning","inspection","lubrication","service","replacement","other"];
const INTERVAL_TYPES=["daily","weekly","monthly"];

function TaskSchedule({maintTasks,setMaintTasks,equipment,maintLogs,setMaintLogs}){
  const [open,setOpen]=useState(false);
  const [logOpen,setLogOpen]=useState(null); // task being logged
  const [filter,setFilter]=useState("all");
  const todayStr=today();

  const blank={equipmentId:"",taskName:"",category:"cleaning",intervalType:"weekly",intervalValue:"1",estimatedMins:"20",notes:"",active:true};
  const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));

  function saveTask(){
    if(!f.equipmentId||!f.taskName)return alert("Equipment and Task Name required.");
    const eq=equipment.find(e=>String(e.id)===String(f.equipmentId));
    const nextDue=today();
    setMaintTasks(ts=>[...ts,{id:uid(),equipmentId:+f.equipmentId,code:eq?.code,equipmentName:eq?.name,...f,intervalValue:+f.intervalValue,estimatedMins:+f.estimatedMins,lastDone:"",nextDue,active:true}]);
    setOpen(false);setF(blank);
  }

  function deleteTask(id){if(window.confirm("Delete this scheduled task?"))setMaintTasks(ts=>ts.filter(t=>t.id!==id));}
  function toggleTask(id){setMaintTasks(ts=>ts.map(t=>t.id===id?{...t,active:!t.active}:t));}

  const catColors={cleaning:"var(--teal)",inspection:"var(--blue)",lubrication:"var(--amber)",service:"var(--purple)",replacement:"var(--red)",other:"var(--text3)"};

  const filtered=maintTasks.filter(t=>{
    if(filter==="overdue") return t.active&&t.nextDue<todayStr;
    if(filter==="due")     return t.active&&daysDiff(t.nextDue)<=7;
    if(filter==="inactive")return !t.active;
    return true;
  }).sort((a,b)=>a.nextDue.localeCompare(b.nextDue));

  function dueBadge(task){
    if(!task.active) return {cls:"ok",label:"Inactive"};
    const d=daysDiff(task.nextDue);
    if(d<0)   return {cls:"overdue",  label:`${Math.abs(d)}d overdue`};
    if(d===0) return {cls:"due-today",label:"Due today"};
    if(d<=7)  return {cls:"due-soon", label:`In ${d}d`};
    return      {cls:"ok",           label:`${d}d`};
  }

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Tasks"  value={maintTasks.length}                                    sub="scheduled"   color="blue"/>
        <StatCard label="Active"       value={maintTasks.filter(t=>t.active).length}                sub="running"     color="green"/>
        <StatCard label="Overdue"      value={maintTasks.filter(t=>t.active&&t.nextDue<todayStr).length} sub="overdue" color={maintTasks.filter(t=>t.active&&t.nextDue<todayStr).length>0?"red":"muted"}/>
        <StatCard label="Due This Week" value={maintTasks.filter(t=>t.active&&daysDiff(t.nextDue)<=7&&daysDiff(t.nextDue)>=0).length} sub="next 7 days" color="amber"/>
      </div>

      <div className="filter-bar">
        {[["all","All"],["overdue","Overdue"],["due","Due This Week"],["inactive","Inactive"]].map(([k,l])=>(
          <button key={k} className={`btn btn-sm ${filter===k?"btn-primary":"btn-ghost"}`} onClick={()=>setFilter(k)}>{l}</button>
        ))}
        <button className="btn btn-primary" style={{marginLeft:"auto"}} onClick={()=>setOpen(true)}>+ Add Task</button>
      </div>

      <div className="table-wrap">
        <table>
          <thead><tr><th>Equipment</th><th>Task</th><th>Category</th><th>Interval</th><th>Last Done</th><th>Next Due</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={8}><EmptyState icon="📋" message="No tasks match filter"/></td></tr>
            :filtered.map(t=>{
              const b=dueBadge(t);
              return(
                <tr key={t.id} style={{opacity:t.active?1:0.5}}>
                  <td><span className="mono" style={{color:"var(--accent)",fontSize:11}}>{t.code}</span> <span style={{fontSize:12,color:"var(--text2)"}}>{t.equipmentName}</span></td>
                  <td style={{fontWeight:500,fontSize:12.5}}>{t.taskName}</td>
                  <td><span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:11}}><span style={{width:7,height:7,borderRadius:"50%",background:catColors[t.category]||"var(--text3)",display:"inline-block"}}/>{t.category}</span></td>
                  <td><span className="interval-chip">{t.intervalValue} {t.intervalType}</span></td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{t.lastDone||"Never"}</td>
                  <td><span className={`due-badge ${b.cls}`}>{b.label}</span></td>
                  <td><Badge type={t.active?"green":"muted"}>{t.active?"Active":"Inactive"}</Badge></td>
                  <td>
                    <div style={{display:"flex",gap:4}}>
                      <button className="btn btn-primary btn-sm" onClick={()=>setLogOpen(t)}>Log Done</button>
                      <button className="btn btn-ghost btn-sm" onClick={()=>toggleTask(t.id)}>{t.active?"Pause":"Resume"}</button>
                      <button className="btn btn-danger btn-sm" onClick={()=>deleteTask(t.id)}>Del</button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Add Task Modal */}
      {open&&(
        <Modal title="Add Scheduled Task" onClose={()=>setOpen(false)}
          foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={saveTask}>Add Task</button></div>}>
          <div className="form-row cols-2">
            <FG label="Equipment *">
              <select value={f.equipmentId} onChange={e=>set("equipmentId")(e.target.value)}>
                <option value="">Select…</option>
                {equipment.map(e=><option key={e.id} value={e.id}>{e.code} — {e.name}</option>)}
              </select>
            </FG>
            <FG label="Task Name *"><input placeholder="e.g. Magnet Cleaning" value={f.taskName} onChange={e=>set("taskName")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-3">
            <FG label="Category">
              <select value={f.category} onChange={e=>set("category")(e.target.value)}>
                {TASK_CATEGORIES.map(c=><option key={c}>{c}</option>)}
              </select>
            </FG>
            <FG label="Every">
              <input type="number" min="1" value={f.intervalValue} onChange={e=>set("intervalValue")(e.target.value)}/>
            </FG>
            <FG label="Interval Type">
              <select value={f.intervalType} onChange={e=>set("intervalType")(e.target.value)}>
                {INTERVAL_TYPES.map(t=><option key={t}>{t}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Estimated Duration (minutes)"><input type="number" placeholder="20" value={f.estimatedMins} onChange={e=>set("estimatedMins")(e.target.value)}/></FG>
          </div>
          <div className="form-row"><FG label="Notes / Instructions"><input placeholder="What to check, tools needed, safety notes…" value={f.notes} onChange={e=>set("notes")(e.target.value)}/></FG></div>
        </Modal>
      )}

      {/* Log Done Modal */}
      {logOpen&&<LogDoneModal task={logOpen} equipment={equipment} onClose={()=>setLogOpen(null)}
        onSave={(logEntry)=>{
          setMaintLogs(ls=>[logEntry,...ls]);
          setMaintTasks(ts=>ts.map(t=>t.id===logOpen.id?{...t,lastDone:logEntry.date,nextDue:calcNextDue(logEntry.date,t.intervalType,t.intervalValue)}:t));
          setLogOpen(null);
        }}/>}
    </div>
  );
}

// ── LOG DONE MODAL (shared) ────────────────────────────────────────────────────
function LogDoneModal({task,equipment,onClose,onSave}){
  const blank={date:today(),type:"Preventive",performedBy:"",durationHours:"",cost:"",
    partsUsed:"",partsCost:"",findings:"",nextServiceDate:"",
    // Ball mill specific
    ballsAddedKg:"",magnetMaterialGrams:"",liningWearPct:"",meshCondition:"good",
    remarks:""};
  const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));
  const isBM = task.code?.startsWith("BM");
  const isMS = task.code?.startsWith("MS");
  const isVS = task.code?.startsWith("VS");
  const totalCost=(+f.cost||0)+(+f.partsCost||0);

  function save(){
    const logNo=`MT-${Date.now().toString().slice(-6)}`;
    onSave({
      id:uid(),logNo,equipmentId:task.equipmentId,equipmentCode:task.code,
      equipmentName:task.equipmentName,taskName:task.taskName,taskId:task.id,
      category:task.category,...f,cost:totalCost,
      durationHours:+f.durationHours||0,
    });
  }

  return(
    <Modal title={`Log: ${task.taskName} — ${task.code}`} onClose={onClose} size="modal-lg"
      foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={onClose}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Log</button></div>}>
      <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:"var(--r)",padding:"10px 14px",marginBottom:14,fontSize:12}}>
        <span className="mono" style={{color:"var(--accent)"}}>{task.code}</span> · {task.equipmentName} · <span style={{color:"var(--text3)"}}>{task.category} · every {task.intervalValue} {task.intervalType}</span>
        {task.notes&&<div style={{color:"var(--text4)",marginTop:4,fontSize:11,fontStyle:"italic"}}>{task.notes}</div>}
      </div>

      <div className="form-row cols-3">
        <FG label="Date Done *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
        <FG label="Type">
          <select value={f.type} onChange={e=>set("type")(e.target.value)}>
            {["Preventive","Breakdown","Inspection","Replacement","Lubrication","Cleaning","Other"].map(t=><option key={t}>{t}</option>)}
          </select>
        </FG>
        <FG label="Performed By"><input placeholder="Name / Technician" value={f.performedBy} onChange={e=>set("performedBy")(e.target.value)}/></FG>
      </div>
      <div className="form-row cols-3">
        <FG label="Duration (hours)"><input type="number" step="0.5" placeholder="1.5" value={f.durationHours} onChange={e=>set("durationHours")(e.target.value)}/></FG>
        <FG label="Service Cost (Rs.)"><input type="number" placeholder="500" value={f.cost} onChange={e=>set("cost")(e.target.value)}/></FG>
        <FG label="Parts Cost (Rs.)"><input type="number" placeholder="0" value={f.partsCost} onChange={e=>set("partsCost")(e.target.value)}/></FG>
      </div>
      <div className="form-row cols-2">
        <FG label="Parts Used"><input placeholder="e.g. 6305 bearing x2, EP-2 grease 200g" value={f.partsUsed} onChange={e=>set("partsUsed")(e.target.value)}/></FG>
        <FG label="Next Service Date"><input type="date" value={f.nextServiceDate} onChange={e=>set("nextServiceDate")(e.target.value)}/></FG>
      </div>

      {/* Ball mill specific fields */}
      {isBM&&(
        <div>
          <div style={{fontSize:10.5,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"var(--mono)",marginBottom:8}}>Ball Mill Specifics</div>
          <div className="log-param-row">
            <FG label="Balls Added (kg)" note="Leave 0 if none added"><input type="number" placeholder="0" value={f.ballsAddedKg} onChange={e=>set("ballsAddedKg")(e.target.value)}/></FG>
            <FG label="Lining Wear (%)" note="Estimated wear since last inspection"><input type="number" placeholder="15" value={f.liningWearPct} onChange={e=>set("liningWearPct")(e.target.value)}/></FG>
            <FG label="Belt Condition">
              <select value={f.meshCondition} onChange={e=>set("meshCondition")(e.target.value)}>
                {["good","fair","worn-replace-soon","replaced"].map(c=><option key={c}>{c}</option>)}
              </select>
            </FG>
          </div>
        </div>
      )}

      {/* Magnetic separator specific */}
      {isMS&&(
        <div>
          <div style={{fontSize:10.5,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"var(--mono)",marginBottom:8}}>Magnet Specifics</div>
          <div className="log-param-row">
            <FG label="Material Caught (grams)" note="Ferrous material removed from magnet"><input type="number" placeholder="45" value={f.magnetMaterialGrams} onChange={e=>set("magnetMaterialGrams")(e.target.value)}/></FG>
            <FG label="Magnet Condition">
              <select value={f.meshCondition} onChange={e=>set("meshCondition")(e.target.value)}>
                {["strong","normal","weak-check-soon","needs-replacement"].map(c=><option key={c}>{c}</option>)}
              </select>
            </FG>
          </div>
        </div>
      )}

      {/* Vibrating screen specific */}
      {isVS&&(
        <div>
          <div style={{fontSize:10.5,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"var(--mono)",marginBottom:8}}>Screen Mesh Specifics</div>
          <div className="log-param-row">
            <FG label="Mesh Condition">
              <select value={f.meshCondition} onChange={e=>set("meshCondition")(e.target.value)}>
                {["good","minor-wear","holes-detected","replaced"].map(c=><option key={c}>{c}</option>)}
              </select>
            </FG>
            <FG label="Mesh Size (microns/mm)"><input placeholder="e.g. 45 micron / 2mm" value={f.partsUsed} onChange={e=>set("partsUsed")(e.target.value)}/></FG>
          </div>
        </div>
      )}

      <div className="form-row">
        <FG label="Findings / Observations"><input placeholder="What was found, what was done, any concerns…" value={f.findings} onChange={e=>set("findings")(e.target.value)}/></FG>
      </div>
      <div className="form-row">
        <FG label="Remarks"><input placeholder="Additional notes" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG>
      </div>

      {totalCost>0&&(
        <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:"var(--r)",padding:"10px 14px",display:"flex",justifyContent:"space-between",fontSize:13}}>
          <span style={{color:"var(--text3)"}}>Total Cost This Activity</span>
          <span style={{fontFamily:"var(--mono)",color:"var(--amber)",fontWeight:700}}>Rs.{totalCost.toLocaleString("en-IN")}</span>
        </div>
      )}
    </Modal>
  );
}

// ── MAINTENANCE LOG (Full History) ────────────────────────────────────────────
function MaintLog({maintLogs,setMaintLogs,equipment,maintTasks,setMaintTasks}){
  const [logOpen,setLogOpen]=useState(false);
  const [eqFilter,setEqFilter]=useState("");
  const [typeFilter,setTypeFilter]=useState("");
  const [exp,setExp]=useState(null);

  const filtered=maintLogs
    .filter(l=>!eqFilter||String(l.equipmentId)===eqFilter)
    .filter(l=>!typeFilter||l.type===typeFilter)
    .sort((a,b)=>b.date.localeCompare(a.date));

  const totalCost=filtered.reduce((s,l)=>s+(+l.cost||0),0);
  const bdCount=filtered.filter(l=>l.type==="Breakdown").length;
  const bdDowntime=filtered.filter(l=>l.type==="Breakdown").reduce((s,l)=>s+(+l.durationHours||0),0);

  // For standalone log (not linked to a task)
  const blankLog={date:today(),equipmentId:"",taskName:"",type:"Preventive",performedBy:"",
    durationHours:"",cost:"",partsUsed:"",findings:"",remarks:""};
  const [lf,setLf]=useState(blankLog);
  const lset=k=>v=>setLf(x=>({...x,[k]:v}));

  function saveStandalone(){
    if(!lf.equipmentId||!lf.taskName)return alert("Equipment and Task required.");
    const eq=equipment.find(e=>String(e.id)===String(lf.equipmentId));
    setMaintLogs(ls=>[{id:uid(),logNo:`MT-${String(ls.length+1).padStart(4,"0")}`,equipmentId:+lf.equipmentId,equipmentCode:eq?.code,equipmentName:eq?.name,...lf,cost:+lf.cost||0,durationHours:+lf.durationHours||0},...ls]);
    setLogOpen(false);setLf(blankLog);
  }

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Logs"    value={maintLogs.length}  sub="all time"    color="blue"/>
        <StatCard label="Breakdowns"    value={bdCount}            sub="all time"    color={bdCount>0?"red":"green"}/>
        <StatCard label="Total Downtime" value={`${bdDowntime.toFixed(1)}h`} sub="from breakdowns" color="amber"/>
        <StatCard label="Total Maint Cost" value={`Rs.${maintLogs.reduce((s,l)=>s+(+l.cost||0),0).toLocaleString("en-IN")}`} sub="all time" color="purple"/>
      </div>
      <div className="filter-bar">
        <select value={eqFilter} onChange={e=>setEqFilter(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"7px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12.5,outline:"none"}}>
          <option value="">All Equipment</option>
          {equipment.map(e=><option key={e.id} value={e.id}>{e.code} — {e.name}</option>)}
        </select>
        <select value={typeFilter} onChange={e=>setTypeFilter(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"7px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12.5,outline:"none"}}>
          <option value="">All Types</option>
          {["Preventive","Breakdown","Inspection","Replacement","Lubrication","Cleaning","Other"].map(t=><option key={t}>{t}</option>)}
        </select>
        <button className="btn btn-primary" onClick={()=>setLogOpen(true)}>+ Log Activity</button>
      </div>

      {filtered.length>0&&(
        <div style={{display:"flex",gap:16,marginBottom:14,padding:"10px 14px",background:"var(--bg2)",border:"1px solid var(--border)",borderRadius:"var(--r2)",fontSize:12}}>
          <span style={{color:"var(--text3)"}}>Filtered: <span style={{color:"var(--text)",fontWeight:600}}>{filtered.length} entries</span></span>
          <span style={{color:"var(--text3)"}}>Cost: <span style={{color:"var(--amber)",fontFamily:"var(--mono)",fontWeight:600}}>Rs.{totalCost.toLocaleString("en-IN")}</span></span>
          {bdCount>0&&<span style={{color:"var(--text3)"}}>Breakdowns: <span style={{color:"var(--red)",fontWeight:600}}>{bdCount} · {bdDowntime.toFixed(1)}h downtime</span></span>}
        </div>
      )}

      <div className="table-wrap">
        <table>
          <thead><tr><th>Log #</th><th>Date</th><th>Equipment</th><th>Task</th><th>Type</th><th>By</th><th className="r">Duration</th><th className="r">Cost</th><th></th></tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9}><EmptyState icon="🔧" message="No maintenance logs yet" sub="Log activities from the Task Schedule or add manually here"/></td></tr>
            :filtered.map(l=>(
              <>
              <tr key={l.id}>
                <td className="mono" style={{color:"var(--accent)",fontSize:11}}>{l.logNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{l.date}</td>
                <td><span className="mono" style={{color:"var(--accent)",fontSize:10}}>{l.equipmentCode}</span> <span style={{fontSize:12}}>{l.equipmentName}</span></td>
                <td style={{fontWeight:500,fontSize:12.5}}>{l.taskName}</td>
                <td><Badge type={l.type==="Breakdown"?"red":l.type==="Replacement"?"amber":l.type==="Preventive"?"green":"muted"}>{l.type}</Badge></td>
                <td style={{color:"var(--text2)",fontSize:12}}>{l.performedBy||"--"}</td>
                <td className="r mono" style={{color:l.type==="Breakdown"?"var(--red)":"var(--text3)"}}>{l.durationHours?l.durationHours+"h":"--"}</td>
                <td className="r mono" style={{color:"var(--amber)"}}>{l.cost>0?`Rs.${l.cost.toLocaleString("en-IN")}`:"--"}</td>
                <td><button className="btn btn-ghost btn-sm" onClick={()=>setExp(exp===l.id?null:l.id)}>{exp===l.id?"▲":"▼"}</button></td>
              </tr>
              {exp===l.id&&<tr key={l.id+"d"}><td colSpan={9} style={{padding:0}}>
                <div className="row-detail">
                  {l.findings&&<div className="row-detail-block"><label>Findings</label><p>{l.findings}</p></div>}
                  {l.partsUsed&&<div className="row-detail-block"><label>Parts Used</label><p>{l.partsUsed}</p></div>}
                  {l.ballsAddedKg>0&&<div className="row-detail-block"><label>Balls Added</label><p style={{color:"var(--accent)"}}>{l.ballsAddedKg} kg</p></div>}
                  {l.magnetMaterialGrams>0&&<div className="row-detail-block"><label>Material Caught</label><p style={{color:"var(--teal)"}}>{l.magnetMaterialGrams}g ferrous</p></div>}
                  {l.liningWearPct&&<div className="row-detail-block"><label>Lining Wear</label><p>{l.liningWearPct}%</p></div>}
                  {l.meshCondition&&<div className="row-detail-block"><label>Condition</label><p>{l.meshCondition}</p></div>}
                  {l.nextServiceDate&&<div className="row-detail-block"><label>Next Service</label><p style={{color:"var(--amber)"}}>{l.nextServiceDate}</p></div>}
                  {l.remarks&&<div className="row-detail-block"><label>Remarks</label><p>{l.remarks}</p></div>}
                </div>
              </td></tr>}
              </>
            ))}
          </tbody>
        </table>
      </div>

      {logOpen&&(
        <Modal title="Log Maintenance Activity" onClose={()=>setLogOpen(false)} size="modal-lg"
          foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={()=>setLogOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={saveStandalone}>Save</button></div>}>
          <div className="form-row cols-2">
            <FG label="Date *"><input type="date" value={lf.date} onChange={e=>lset("date")(e.target.value)}/></FG>
            <FG label="Equipment *">
              <select value={lf.equipmentId} onChange={e=>lset("equipmentId")(e.target.value)}>
                <option value="">Select…</option>
                {equipment.map(e=><option key={e.id} value={e.id}>{e.code} — {e.name}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Task / Activity *"><input placeholder="e.g. Magnet cleaning, Belt replacement" value={lf.taskName} onChange={e=>lset("taskName")(e.target.value)}/></FG>
            <FG label="Type">
              <select value={lf.type} onChange={e=>lset("type")(e.target.value)}>
                {["Preventive","Breakdown","Inspection","Replacement","Lubrication","Cleaning","Other"].map(t=><option key={t}>{t}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-3">
            <FG label="Performed By"><input placeholder="Name" value={lf.performedBy} onChange={e=>lset("performedBy")(e.target.value)}/></FG>
            <FG label="Duration (hours)"><input type="number" step="0.5" placeholder="1.5" value={lf.durationHours} onChange={e=>lset("durationHours")(e.target.value)}/></FG>
            <FG label="Cost (Rs.)"><input type="number" placeholder="0" value={lf.cost} onChange={e=>lset("cost")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Parts Used"><input placeholder="e.g. Bearing 6305, EP-2 grease" value={lf.partsUsed} onChange={e=>lset("partsUsed")(e.target.value)}/></FG>
            <FG label="Findings"><input placeholder="What was found / done" value={lf.findings} onChange={e=>lset("findings")(e.target.value)}/></FG>
          </div>
          <div className="form-row"><FG label="Remarks"><input value={lf.remarks} onChange={e=>lset("remarks")(e.target.value)}/></FG></div>
        </Modal>
      )}
    </div>
  );
}

// ── PURCHASE REGISTER ─────────────────────────────────────────────────────────
function PurchaseRegister({suppliers,purchases,setPurchases}){
  const [open,setOpen]=useState(false);
  const blank={date:today(),supplierId:"",description:"",quantityMT:"",ratePerMT:"",gstPct:"5",invoiceNo:"",remarks:""};
  const [f,setF]=useState(blank);const set=k=>v=>setF(x=>({...x,[k]:v}));
  const taxable=(+f.quantityMT||0)*(+f.ratePerMT||0);
  const gstAmt=taxable*(+f.gstPct||0)/100;
  const totalAmt=taxable+gstAmt;
  function save(){
    if(!f.supplierId||!f.quantityMT||!f.ratePerMT)return alert("Supplier, Quantity and Rate required.");
    const s=suppliers.find(s=>String(s.id)===String(f.supplierId));
    setPurchases(ps=>[{id:uid(),poNo:`PO-${String(ps.length+1).padStart(4,"0")}`,...f,quantityMT:+f.quantityMT,ratePerMT:+f.ratePerMT,gstPct:+f.gstPct,taxableAmount:taxable,gstAmount:gstAmt,totalAmount:totalAmt,supplierName:s?.name},...ps]);
    setOpen(false);setF({...blank,date:f.date});
  }
  const totalSpend=purchases.reduce((s,p)=>s+p.totalAmount,0);
  const totalRM=purchases.reduce((s,p)=>s+p.quantityMT,0);
  const avgRate=totalRM>0?(purchases.reduce((s,p)=>s+p.taxableAmount,0)/totalRM).toFixed(0):null;
  const totalGST=purchases.reduce((s,p)=>s+p.gstAmount,0);
  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total RM Purchased" value={`${totalRM.toFixed(1)} MT`}  sub={`${purchases.length} orders`}    color="teal"/>
        <StatCard label="Total Spend"         value={`₹${(totalSpend/1000).toFixed(1)}K`} sub="incl. GST"             color="amber"/>
        <StatCard label="Avg Rate / MT"       value={avgRate?`₹${avgRate}`:"--"}  sub="excl. GST"                       color="blue"/>
        <StatCard label="Input GST Paid"      value={`₹${(totalGST/1000).toFixed(1)}K`} sub="ITC claimable in Tally"  color="accent"/>
      </div>
      <div className="filter-bar"><button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New Purchase</button></div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>PO #</th><th>Date</th><th>Supplier</th><th>Description</th><th className="r">Qty (MT)</th><th className="r">Rate/MT</th><th className="r">Taxable</th><th className="r">GST</th><th className="r">Total</th></tr></thead>
          <tbody>
            {purchases.length===0?<tr><td colSpan={9}><EmptyState icon="📥" message="No purchases yet" sub="Log raw material purchases to track RM cost"/></td></tr>
            :purchases.map(p=>(
              <tr key={p.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{p.poNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{p.date}</td>
                <td style={{fontWeight:500}}>{p.supplierName}</td>
                <td style={{color:"var(--text2)"}}>{p.description||"Raw Quartz"}</td>
                <td className="r mono">{(+p.quantityMT).toFixed(3)} MT</td>
                <td className="r mono" style={{color:"var(--text2)"}}>₹{p.ratePerMT}</td>
                <td className="r mono">₹{p.taxableAmount.toLocaleString("en-IN")}</td>
                <td className="r mono" style={{color:"var(--accent)",fontSize:11}}>₹{p.gstAmount.toLocaleString("en-IN")}</td>
                <td className="r mono" style={{fontWeight:700,color:"var(--amber)"}}>₹{p.totalAmount.toLocaleString("en-IN")}</td>
              </tr>
            ))}
          </tbody>
          {purchases.length>0&&<tfoot><tr className="ledger-foot">
            <td colSpan={4} style={{padding:"10px 16px",fontSize:12,color:"var(--text2)"}}>Totals</td>
            <td className="r mono" style={{padding:"10px 16px",color:"var(--teal)"}}>{totalRM.toFixed(3)} MT</td>
            <td style={{padding:"10px 16px"}}></td>
            <td className="r mono" style={{padding:"10px 16px"}}>₹{purchases.reduce((s,p)=>s+p.taxableAmount,0).toLocaleString("en-IN")}</td>
            <td className="r mono" style={{padding:"10px 16px",color:"var(--accent)"}}>₹{totalGST.toLocaleString("en-IN")}</td>
            <td className="r mono" style={{padding:"10px 16px",color:"var(--amber)",fontSize:14}}>₹{totalSpend.toLocaleString("en-IN")}</td>
          </tr></tfoot>}
        </table>
      </div>
      {open&&<Modal title="New Purchase Entry" onClose={()=>setOpen(false)}
        foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Purchase</button></div>}>
        <div className="form-row cols-2">
          <FG label="Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
          <FG label="Supplier *"><select value={f.supplierId} onChange={e=>set("supplierId")(e.target.value)}><option value="">Select…</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</select></FG>
        </div>
        <div className="form-row cols-2">
          <FG label="Description"><input placeholder="Raw Quartz Boulder" value={f.description} onChange={e=>set("description")(e.target.value)}/></FG>
          <FG label="Supplier Invoice #"><input placeholder="INV-XXXX" value={f.invoiceNo} onChange={e=>set("invoiceNo")(e.target.value)}/></FG>
        </div>
        <div className="form-row cols-3">
          <FG label="Quantity (MT) *"><input type="number" step="0.001" placeholder="100.000" value={f.quantityMT} onChange={e=>set("quantityMT")(e.target.value)}/></FG>
          <FG label="Rate per MT (Rs.) *"><input type="number" placeholder="1200" value={f.ratePerMT} onChange={e=>set("ratePerMT")(e.target.value)}/></FG>
          <FG label="GST %"><select value={f.gstPct} onChange={e=>set("gstPct")(e.target.value)}>{[0,5,12,18].map(r=><option key={r} value={r}>{r}%</option>)}</select></FG>
        </div>
        {taxable>0&&<div style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"12px 14px",marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}><span style={{color:"var(--text3)",fontFamily:"var(--mono)"}}>Taxable</span><span style={{fontFamily:"var(--mono)"}}>₹{taxable.toLocaleString("en-IN")}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}><span style={{color:"var(--text3)",fontFamily:"var(--mono)"}}>GST ({f.gstPct}%)</span><span style={{fontFamily:"var(--mono)",color:"var(--accent)"}}>₹{gstAmt.toLocaleString("en-IN")}</span></div>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:14,fontWeight:700,borderTop:"1px solid var(--border)",paddingTop:6}}><span style={{color:"var(--text2)",fontFamily:"var(--mono)"}}>Total</span><span style={{fontFamily:"var(--mono)",color:"var(--amber)"}}>₹{totalAmt.toLocaleString("en-IN")}</span></div>
        </div>}
        <div className="form-row"><FG label="Remarks"><input placeholder="Optional" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG></div>
      </Modal>}
    </div>
  );
}

// ── LOT REGISTER ──────────────────────────────────────────────────────────────
function LotRegister({lots,setLots,qcTests,grades}){
  const [search,setSearch]=useState("");
  const filtered=lots.filter(l=>
    l.lotNo.toLowerCase().includes(search.toLowerCase())||
    (l.gradeOutputs||[]).some(o=>o.gradeCode?.toLowerCase().includes(search.toLowerCase()))
  );
  function updateStatus(id,status){setLots(ls=>ls.map(l=>l.id===id?{...l,status}:l));}
  const statusColor={
    "in-stock":"in-stock","under-test":"under-test",
    "released":"released","dispatched":"dispatched","rejected":"rejected"
  };
  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Lots"     value={lots.length}                                             sub="all time"       color="blue"/>
        <StatCard label="In Stock"       value={lots.filter(l=>l.status==="in-stock").length}            sub="awaiting test"  color="amber"/>
        <StatCard label="Released"       value={lots.filter(l=>l.status==="released").length}            sub="QC passed"      color="green"/>
        <StatCard label="Rejected"       value={lots.filter(l=>l.status==="rejected").length}            sub="failed QC"      color="red"/>
      </div>
      <div className="filter-bar">
        <input placeholder="Search lot # or grade…" value={search} onChange={e=>setSearch(e.target.value)} style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"7px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12.5,outline:"none",minWidth:220}}/>
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Lot #</th><th>Date</th><th>Shift</th><th>Grade Output</th><th className="r">Total MT</th><th>Yield</th><th>QC Tests</th><th>Status</th><th>Action</th></tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9}><EmptyState icon="🏷" message="No lots yet" sub="Lots are auto-created when you save a Production Entry"/></td></tr>
            :filtered.map(lot=>{
              const tests=qcTests.filter(t=>t.lotNo===lot.lotNo);
              const latestTest=tests.sort((a,b)=>b.date.localeCompare(a.date))[0];
              return(
                <tr key={lot.id}>
                  <td><span className="lot-chip">🏷 {lot.lotNo}</span></td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{lot.date}</td>
                  <td><Badge type="muted">{(lot.shift||"").split(" ")[0]}</Badge></td>
                  <td style={{fontSize:12}}>
                    {(lot.gradeOutputs||[]).map((o,i)=>(
                      <span key={i} style={{display:"inline-block",marginRight:6}}><Badge type="muted">{o.gradeCode}</Badge> <span style={{color:"var(--text3)",fontSize:11}}>{(+o.quantityMT||0).toFixed(2)}MT</span></span>
                    ))}
                  </td>
                  <td className="r mono" style={{color:"var(--teal)",fontWeight:700}}>{(+lot.totalOutputMT||0).toFixed(3)} MT</td>
                  <td>{lot.yieldPct?<Badge type={lot.yieldPct>=90?"green":lot.yieldPct>=80?"amber":"red"}>{lot.yieldPct}%</Badge>:"--"}</td>
                  <td style={{fontSize:12}}>
                    {tests.length===0?<span style={{color:"var(--text4)"}}>No tests</span>
                    :<span>{tests.length} test{tests.length>1?"s":""} {latestTest&&<Badge type={latestTest.result}>{latestTest.result}</Badge>}</span>}
                  </td>
                  <td><Badge type={statusColor[lot.status]||"muted"}>{lot.status}</Badge></td>
                  <td>
                    <select value={lot.status} onChange={e=>updateStatus(lot.id,e.target.value)}
                      style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"4px 8px",color:"var(--text)",fontFamily:"var(--font)",fontSize:11.5,outline:"none",cursor:"pointer"}}>
                      {["in-stock","under-test","released","dispatched","rejected"].map(s=><option key={s} value={s}>{s}</option>)}
                    </select>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ── QC TEST ENTRY ─────────────────────────────────────────────────────────────
const QC_PARAMS=[
  {key:"whiteness",   label:"Whiteness (L*)",  unit:"L* value",  placeholder:"92.5"},
  {key:"sio2",        label:"SiO₂",         unit:"%",         placeholder:"99.2"},
  {key:"fe2o3",       label:"Fe₂O₃",    unit:"%",         placeholder:"0.025"},
  {key:"al2o3",       label:"Al₂O₃",    unit:"%",         placeholder:"0.12"},
  {key:"moisture",    label:"Moisture",        unit:"%",         placeholder:"0.3"},
  {key:"loi",         label:"LOI",             unit:"%",         placeholder:"0.4"},
  {key:"d50",         label:"Particle D50",   unit:"microns",   placeholder:"45"},
  {key:"bulkDensity", label:"Bulk Density",   unit:"g/cc",      placeholder:"0.95"},
];

function QCTests({lots,qcTests,setQcTests,setSamples}){
  const [open,setOpen]=useState(false);
  const blank={date:today(),lotNo:"",testedBy:"",labType:"In-house",certNo:"",result:"pass",remarks:""};
  const blankParams=Object.fromEntries(QC_PARAMS.map(p=>[p.key,""]));
  const [f,setF]=useState(blank);
  const [params,setParams]=useState(blankParams);
  const [exp,setExp]=useState(null);
  const set=k=>v=>setF(x=>({...x,[k]:v}));

  function save(){
    if(!f.lotNo)return alert("Select a Lot.");
    const testNo=`QC-${String(qcTests.length+1).padStart(4,"0")}`;
    const filledParams=Object.fromEntries(Object.entries(params).filter(([,v])=>v!==""));
    setQcTests(ts=>[{id:uid(),testNo,...f,params:filledParams},...ts]);
    // Auto-update lot status based on result
    setOpen(false);
    setF({...blank,date:f.date});
    setParams(blankParams);
  }

  const statusColor={pass:"pass",fail:"fail",conditional:"conditional"};

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Tests"  value={qcTests.length}                                    sub="all time"  color="blue"/>
        <StatCard label="Passed"       value={qcTests.filter(t=>t.result==="pass").length}       sub="released"  color="green"/>
        <StatCard label="Failed"       value={qcTests.filter(t=>t.result==="fail").length}       sub="rejected"  color="red"/>
        <StatCard label="Conditional"  value={qcTests.filter(t=>t.result==="conditional").length} sub="review"  color="amber"/>
      </div>
      <div className="filter-bar">
        <button className="btn btn-primary" onClick={()=>setOpen(true)}>+ New QC Test</button>
      </div>
      <div className="table-wrap">
        <table>
          <thead><tr><th>Test #</th><th>Date</th><th>Lot #</th><th>Lab</th><th>Tested By</th><th>Whiteness L*</th><th>Fe₂O₃ %</th><th>SiO₂ %</th><th>Result</th><th></th></tr></thead>
          <tbody>
            {qcTests.length===0?<tr><td colSpan={10}><EmptyState icon="🔬" message="No QC tests yet" sub="Record lab test results for each lot"/></td></tr>
            :qcTests.map(t=>(
              <>
              <tr key={t.id}>
                <td className="mono" style={{color:"var(--accent)"}}>{t.testNo}</td>
                <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{t.date}</td>
                <td><span className="lot-chip">🏷 {t.lotNo}</span></td>
                <td><Badge type="muted">{t.labType}</Badge></td>
                <td style={{color:"var(--text2)"}}>{t.testedBy||"--"}</td>
                <td className="mono" style={{color:+t.params?.whiteness>=90?"var(--green)":+t.params?.whiteness>=85?"var(--amber)":"var(--red)",fontWeight:700}}>{t.params?.whiteness||"--"}</td>
                <td className="mono" style={{color:+t.params?.fe2o3<=0.03?"var(--green)":+t.params?.fe2o3<=0.05?"var(--amber)":"var(--red)",fontWeight:700}}>{t.params?.fe2o3||"--"}</td>
                <td className="mono" style={{color:+t.params?.sio2>=99?"var(--green)":"var(--amber)"}}>{t.params?.sio2||"--"}</td>
                <td><Badge type={statusColor[t.result]||"muted"}>{t.result}</Badge></td>
                <td><button className="btn btn-ghost btn-sm" onClick={()=>setExp(exp===t.id?null:t.id)}>{exp===t.id?"▲":"▼"}</button></td>
              </tr>
              {exp===t.id&&<tr key={t.id+"d"}><td colSpan={10} style={{padding:0}}>
                <div className="row-detail">
                  {QC_PARAMS.filter(p=>t.params?.[p.key]).map(p=>(
                    <div key={p.key} className="row-detail-block">
                      <label>{p.label}</label>
                      <p style={{fontFamily:"var(--mono)",color:"var(--accent)"}}>{t.params[p.key]} <span style={{color:"var(--text4)",fontSize:10}}>{p.unit}</span></p>
                    </div>
                  ))}
                  {t.certNo&&<div className="row-detail-block"><label>Cert #</label><p>{t.certNo}</p></div>}
                  {t.remarks&&<div className="row-detail-block"><label>Remarks</label><p>{t.remarks}</p></div>}
                </div>
              </td></tr>}
              </>
            ))}
          </tbody>
        </table>
      </div>

      {open&&(
        <Modal title="New QC Test Entry" onClose={()=>setOpen(false)} size="modal-lg"
          foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Test</button></div>}>
          <div className="form-row cols-3">
            <FG label="Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
            <FG label="Lot # *">
              <select value={f.lotNo} onChange={e=>set("lotNo")(e.target.value)}>
                <option value="">Select lot…</option>
                {lots.map(l=><option key={l.id} value={l.lotNo}>{l.lotNo} — {l.date}</option>)}
              </select>
            </FG>
            <FG label="Lab Type">
              <select value={f.labType} onChange={e=>set("labType")(e.target.value)}>
                {["In-house","External Lab","Third Party"].map(t=><option key={t}>{t}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-3">
            <FG label="Tested By"><input placeholder="Name / Lab name" value={f.testedBy} onChange={e=>set("testedBy")(e.target.value)}/></FG>
            <FG label="Certificate #"><input placeholder="Cert / Report number" value={f.certNo} onChange={e=>set("certNo")(e.target.value)}/></FG>
            <FG label="Overall Result">
              <select value={f.result} onChange={e=>set("result")(e.target.value)}>
                {["pass","fail","conditional"].map(r=><option key={r}>{r}</option>)}
              </select>
            </FG>
          </div>

          <div style={{marginBottom:6,fontSize:10.5,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".06em",fontFamily:"var(--mono)"}}>Lab Parameters (fill only what was tested)</div>
          <div className="qc-param-grid">
            {QC_PARAMS.map(p=>(
              <div key={p.key} className="qc-param-card">
                <label>{p.label}</label>
                <input type="number" step="0.001" placeholder={p.placeholder} value={params[p.key]} onChange={e=>setParams(x=>({...x,[p.key]:e.target.value}))}/>
                <div className="unit">{p.unit}</div>
              </div>
            ))}
          </div>

          <div className="form-row">
            <FG label="Remarks"><input placeholder="Any notes, deviations, conditions…" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── SAMPLE REGISTER ───────────────────────────────────────────────────────────
function SampleRegister({lots,samples,setSamples}){
  const [open,setOpen]=useState(false);
  const [filter,setFilter]=useState("all");
  const blank={date:today(),lotNo:"",location:"",collectedBy:"",quantity:"",remarks:""};
  const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));

  function save(){
    if(!f.lotNo||!f.location)return alert("Lot and Storage Location required.");
    const sampleNo=`SMP-${String(samples.length+1).padStart(4,"0")}`;
    const disposalDate=addDays(f.date,30);
    setSamples(ss=>[{id:uid(),sampleNo,...f,disposalDate,status:"retained"},...ss]);
    setOpen(false);
    setF({...blank,date:f.date});
  }
  function markDisposed(id){setSamples(ss=>ss.map(s=>s.id===id?{...s,status:"disposed",disposedDate:today()}:s));}

  const todayStr=today();
  const dueToday=samples.filter(s=>s.status==="retained"&&s.disposalDate<=todayStr);
  const dueSoon=samples.filter(s=>s.status==="retained"&&s.disposalDate>todayStr&&s.disposalDate<=addDays(todayStr,7));

  const filtered=filter==="all"?samples
    :filter==="retained"?samples.filter(s=>s.status==="retained")
    :filter==="due"?samples.filter(s=>s.status==="retained"&&s.disposalDate<=addDays(todayStr,7))
    :samples.filter(s=>s.status==="disposed");

  return(
    <div>
      <div className="stats-grid">
        <StatCard label="Total Samples"  value={samples.length}                                           sub="all time"            color="blue"/>
        <StatCard label="Retained"       value={samples.filter(s=>s.status==="retained").length}          sub="in storage"          color="teal"/>
        <StatCard label="Due This Week"  value={dueSoon.length+dueToday.length}                           sub="for disposal"        color={dueToday.length>0?"red":"amber"}/>
        <StatCard label="Disposed"       value={samples.filter(s=>s.status==="disposed").length}          sub="cleared"             color="muted"/>
      </div>

      {dueToday.length>0&&(
        <div className="alert-strip">
          <h4>Disposal Due Today ({dueToday.length})</h4>
          {dueToday.map(s=>(
            <div key={s.id} className="sample-due">
              <div>
                <span className="mono" style={{color:"var(--accent)",fontSize:11}}>{s.sampleNo}</span>
                <span style={{margin:"0 8px",color:"var(--text3)"}}>·</span>
                <span className="lot-chip" style={{fontSize:10}}>🏷 {s.lotNo}</span>
                <span style={{margin:"0 8px",color:"var(--text3)"}}>·</span>
                <span style={{fontSize:12,color:"var(--text2)"}}>{s.location}</span>
              </div>
              <button className="btn btn-warn btn-sm" onClick={()=>markDisposed(s.id)}>Mark Disposed</button>
            </div>
          ))}
        </div>
      )}

      {dueSoon.length>0&&(
        <div style={{background:"#0e1000",border:"1px solid #2a2800",borderRadius:"var(--r2)",padding:"12px 16px",marginBottom:16}}>
          <div style={{fontSize:11,fontWeight:700,color:"var(--amber)",textTransform:"uppercase",letterSpacing:".08em",marginBottom:8}}>Due Within 7 Days ({dueSoon.length})</div>
          {dueSoon.map(s=>(
            <div key={s.id} style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:4}}>
              <span><span className="mono" style={{color:"var(--accent)",fontSize:11}}>{s.sampleNo}</span> · <span className="lot-chip" style={{fontSize:10}}>🏷 {s.lotNo}</span> · {s.location}</span>
              <span style={{fontFamily:"var(--mono)",color:"var(--amber)",fontSize:11}}>Due: {s.disposalDate}</span>
            </div>
          ))}
        </div>
      )}

      <div className="filter-bar">
        {["all","retained","due","disposed"].map(t=>(
          <button key={t} className={`btn btn-sm ${filter===t?"btn-primary":"btn-ghost"}`} onClick={()=>setFilter(t)}>
            {t==="all"?"All":t==="retained"?"Retained":t==="due"?"Due Soon":"Disposed"}
          </button>
        ))}
        <button className="btn btn-primary" style={{marginLeft:"auto"}} onClick={()=>setOpen(true)}>+ New Sample</button>
      </div>

      <div className="table-wrap">
        <table>
          <thead><tr><th>Sample #</th><th>Lot #</th><th>Collected</th><th>Storage Location</th><th>Collected By</th><th>Qty</th><th>Disposal Date</th><th>Status</th><th>Action</th></tr></thead>
          <tbody>
            {filtered.length===0?<tr><td colSpan={9}><EmptyState icon="🧪" message="No samples" sub="Collect and register a sample for every production lot"/></td></tr>
            :filtered.map(s=>{
              const isOverdue=s.status==="retained"&&s.disposalDate<=todayStr;
              const isDueSoon=s.status==="retained"&&s.disposalDate>todayStr&&s.disposalDate<=addDays(todayStr,7);
              return(
                <tr key={s.id} style={{background:isOverdue?"rgba(255,92,92,.04)":isDueSoon?"rgba(245,166,35,.03)":""}}>
                  <td className="mono" style={{color:"var(--accent)"}}>{s.sampleNo}</td>
                  <td><span className="lot-chip">🏷 {s.lotNo}</span></td>
                  <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{s.date}</td>
                  <td style={{fontWeight:500}}>{s.location}</td>
                  <td style={{color:"var(--text2)"}}>{s.collectedBy||"--"}</td>
                  <td style={{color:"var(--text3)"}}>{s.quantity||"--"}</td>
                  <td className="mono" style={{fontSize:11,color:isOverdue?"var(--red)":isDueSoon?"var(--amber)":"var(--text3)",fontWeight:isOverdue||isDueSoon?700:400}}>
                    {s.disposalDate} {isOverdue&&"⚠"} {isDueSoon&&"⏰"}
                  </td>
                  <td><Badge type={s.status==="retained"?"retained":"disposed"}>{s.status}</Badge></td>
                  <td>
                    {s.status==="retained"
                      ?<button className="btn btn-warn btn-sm" onClick={()=>markDisposed(s.id)}>Dispose</button>
                      :<span className="mono" style={{fontSize:10,color:"var(--text4)"}}>{s.disposedDate}</span>
                    }
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {open&&(
        <Modal title="Register New Sample" onClose={()=>setOpen(false)}
          foot={<div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Sample</button></div>}>
          <div className="form-row cols-2">
            <FG label="Collection Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
            <FG label="Lot # *">
              <select value={f.lotNo} onChange={e=>set("lotNo")(e.target.value)}>
                <option value="">Select lot…</option>
                {lots.map(l=><option key={l.id} value={l.lotNo}>{l.lotNo} — {l.date}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Storage Location *" note="e.g. Shelf A-3, Rack 2"><input placeholder="Shelf A-3" value={f.location} onChange={e=>set("location")(e.target.value)}/></FG>
            <FG label="Collected By"><input placeholder="Operator name" value={f.collectedBy} onChange={e=>set("collectedBy")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Sample Quantity" note="e.g. 500g, 1kg"><input placeholder="500g" value={f.quantity} onChange={e=>set("quantity")(e.target.value)}/></FG>
            <FG label="Disposal Date" note="Auto-set to 30 days from collection">
              <input type="date" value={f.date?addDays(f.date,30):""} readOnly style={{opacity:0.6,cursor:"not-allowed"}}/>
            </FG>
          </div>
          <div className="form-row">
            <FG label="Remarks"><input placeholder="Any additional notes" value={f.remarks} onChange={e=>set("remarks")(e.target.value)}/></FG>
          </div>
        </Modal>
      )}
    </div>
  );
}


// ── COST ENGINE (pure computation) ───────────────────────────────────────────
function computeMonthlyCosts(ym, purchases, expenses, productionEntries, weighments, costConfig){
  const inMonth = cat => expenses.filter(e=>e.date?.slice(0,7)===ym&&e.category===cat).reduce((s,e)=>s+e.amount,0);
  const rmCost    = purchases.filter(p=>p.date?.slice(0,7)===ym).reduce((s,p)=>s+p.taxableAmount,0);
  const salary    = inMonth("salary");
  const labour    = inMonth("labour");
  const salaries  = salary+labour; // combined labour cost (staff + contract)
  const power     = inMonth("electric");
  const water     = inMonth("water");
  const fuel      = inMonth("fuel");
  const maint     = inMonth("maint");
  const transport = inMonth("transport");
  const misc      = inMonth("misc");
  const prodMT    = productionEntries.filter(p=>p.date?.slice(0,7)===ym).reduce((s,p)=>s+p.totalOutputMT,0);
  const soldMT    = weighments.filter(w=>w.date?.slice(0,7)===ym&&w.purpose==="Sale").reduce((s,w)=>s+(w.netWeight||0)/1000,0);
  const packingCost = soldMT * (+costConfig.bagCostPerTonne||0);
  const freightCost = soldMT * (+costConfig.freightPerTonne||0);
  const labCost     = +costConfig.monthlyLabCost||0;
  const totalCost   = rmCost+salaries+power+water+fuel+maint+transport+misc+packingCost+freightCost+labCost;
  const cpt         = prodMT>0 ? totalCost/prodMT : 0;
  return{rmCost,salary,labour,salaries,power,water,fuel,maint,transport,misc,packingCost,freightCost,labCost,totalCost,prodMT,soldMT,cpt};
}

// ── MONTHLY OPERATING COSTS (quick entry) ────────────────────────────────────
// One screen to punch in this month's running-cost totals — electricity, water,
// salaries, contract labour, fuel, maintenance, transport, misc — instead of
// logging every bill as a separate Expense line. Saving writes/updates a single
// tagged (opsAuto:true) entry per category into the shared Expenses log, so the
// Cost Sheet, Expenses register and dashboards all read from one source of truth.
function MonthlyOpsCosts({expenses,setExpenses}){
  const now=new Date();
  const [ym,setYm]=useState(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`);
  const [vals,setVals]=useState({});

  useEffect(()=>{
    const next={};
    OPS_CATS.forEach(cat=>{
      const existing=expenses.filter(e=>e.opsAuto&&e.category===cat&&e.date?.slice(0,7)===ym);
      next[cat]=existing.length?String(existing.reduce((s,e)=>s+e.amount,0)):"";
    });
    setVals(next);
  },[ym]); // eslint-disable-line react-hooks/exhaustive-deps

  const set=cat=>v=>setVals(x=>({...x,[cat]:v}));
  const label=cat=>EXP_CATS.find(c=>c.id===cat)?.label||cat;
  // amount already logged this month via the ordinary Expenses form (not this screen) — shown so nothing gets double-counted
  const adhoc=cat=>expenses.filter(e=>!e.opsAuto&&e.category===cat&&e.date?.slice(0,7)===ym).reduce((s,e)=>s+e.amount,0);
  const noteFor=cat=>{const a=adhoc(cat);return a>0?`+ ₹${a.toLocaleString("en-IN")} already logged as individual entries this month`:undefined;};

  function save(){
    setExpenses(es=>{
      let next=es.filter(e=>!(e.opsAuto&&e.date?.slice(0,7)===ym&&OPS_CATS.includes(e.category)));
      OPS_CATS.forEach(cat=>{
        const amt=+vals[cat];
        if(amt>0){
          next=[{id:uid(),expNo:`EXP-OPS-${ym}-${cat}`,date:`${ym}-01`,category:cat,amount:amt,description:`Monthly ${label(cat)} total`,paidTo:"",reference:"",opsAuto:true},...next];
        }
      });
      return next;
    });
  }

  const enteredTotal=OPS_CATS.reduce((s,cat)=>s+(+vals[cat]||0),0);
  const adhocTotal=OPS_CATS.reduce((s,cat)=>s+adhoc(cat),0);

  const numInput=(cat,placeholder="0")=>(
    <input type="number" placeholder={placeholder} value={vals[cat]||""} onChange={e=>set(cat)(e.target.value)}/>
  );

  return(
    <div>
      <div style={{marginBottom:20}}>
        <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Monthly Operating Costs</h3>
        <p style={{fontSize:12,color:"var(--text3)"}}>Punch in this month's totals once — electricity, water, salaries, contract labour and the rest — instead of logging every bill as a separate expense. These write straight into Expenses and roll up into the Cost Sheet automatically.</p>
      </div>

      <div style={{marginBottom:20}}>
        <FG label="Month">
          <input type="month" value={ym} onChange={e=>setYm(e.target.value)}
            style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"8px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:13,outline:"none",maxWidth:200}}/>
        </FG>
      </div>

      <div className="config-card">
        <h3>Utilities</h3>
        <div className="config-row">
          <FG label="Electricity Bill (₹)" note={noteFor("electric")}>{numInput("electric","e.g. 42000")}</FG>
          <FG label="Water Bill (₹)" note={noteFor("water")}>{numInput("water","e.g. 3000")}</FG>
        </div>
      </div>

      <div className="config-card">
        <h3>Labour</h3>
        <div className="config-row">
          <FG label="Staff Salaries (₹)" note={noteFor("salary")}>{numInput("salary","e.g. 180000")}</FG>
          <FG label="Contract / Casual Labour (₹)" note={noteFor("labour")}>{numInput("labour","e.g. 25000")}</FG>
        </div>
      </div>

      <div className="config-card">
        <h3>Other Recurring Costs</h3>
        <div className="config-row">
          <FG label="Fuel & Diesel (₹)" note={noteFor("fuel")}>{numInput("fuel")}</FG>
          <FG label="Maintenance (₹)" note={noteFor("maint")}>{numInput("maint")}</FG>
        </div>
        <div className="config-row" style={{marginTop:14}}>
          <FG label="Transport (₹)" note={noteFor("transport")}>{numInput("transport")}</FG>
          <FG label="Miscellaneous (₹)" note={noteFor("misc")}>{numInput("misc")}</FG>
        </div>
      </div>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:14,background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:"var(--r2)",padding:"14px 18px"}}>
        <div>
          <div style={{fontSize:11,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".07em",fontFamily:"var(--mono)",marginBottom:4}}>Total — {ym}</div>
          <div style={{fontSize:20,fontWeight:700,fontFamily:"var(--mono)"}}>₹{(enteredTotal+adhocTotal).toLocaleString("en-IN",{maximumFractionDigits:0})}</div>
          {adhocTotal>0&&<div style={{fontSize:10.5,color:"var(--text3)",marginTop:3}}>Includes ₹{adhocTotal.toLocaleString("en-IN")} already logged individually in Expenses</div>}
        </div>
        <button className="btn btn-primary" onClick={save}>Save Monthly Totals</button>
      </div>
    </div>
  );
}

// ── COST CONFIG ───────────────────────────────────────────────────────────────
function CostConfig({costConfig,setCostConfig}){
  const set=k=>v=>setCostConfig(x=>({...x,[k]:v}));
  return(
    <div>
      <div style={{marginBottom:20}}>
        <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Cost Configuration</h3>
        <p style={{fontSize:12,color:"var(--text3)"}}>Set your per-unit variable costs. These are used to calculate Cost per Tonne and Lot P&L. Fixed costs like salaries and power are pulled automatically from the Expenses module.</p>
      </div>

      <div className="config-card">
        <h3>Packing & Freight</h3>
        <div className="config-row">
          <FG label="1-Tonne Bag Cost (₹ per bag)" note="FIBC / jumbo bag cost">
            <input type="number" placeholder="850" value={costConfig.bagCostPerTonne||""} onChange={e=>set("bagCostPerTonne")(e.target.value)}/>
          </FG>
          <FG label="Freight Rate (₹ per MT)" note="Average outbound freight per tonne">
            <input type="number" placeholder="400" value={costConfig.freightPerTonne||""} onChange={e=>set("freightPerTonne")(e.target.value)}/>
          </FG>
          <FG label="Loading Labour (₹ per MT)" note="Bag filling + loading cost">
            <input type="number" placeholder="80" value={costConfig.loadingPerTonne||""} onChange={e=>set("loadingPerTonne")(e.target.value)}/>
          </FG>
        </div>
      </div>

      <div className="config-card">
        <h3>Fixed Monthly Costs (if not tracked in Expenses)</h3>
        <div className="config-row">
          <FG label="Monthly Lab Testing (₹)" note="External lab fees if fixed monthly">
            <input type="number" placeholder="5000" value={costConfig.monthlyLabCost||""} onChange={e=>set("monthlyLabCost")(e.target.value)}/>
          </FG>
          <FG label="Plant Rent / Lease (₹/month)" note="Leave 0 if tracked in Expenses">
            <input type="number" placeholder="0" value={costConfig.monthlyRent||""} onChange={e=>set("monthlyRent")(e.target.value)}/>
          </FG>
          <FG label="Misc Fixed (₹/month)" note="Other fixed costs not in Expenses">
            <input type="number" placeholder="0" value={costConfig.miscFixed||""} onChange={e=>set("miscFixed")(e.target.value)}/>
          </FG>
        </div>
      </div>

      <div className="config-card">
        <h3>Grade-wise Selling Rates (₹ per MT)</h3>
        <p style={{fontSize:11.5,color:"var(--text3)",marginBottom:14}}>Used for margin calculations. These override the Grade Master rate card for P&L purposes.</p>
        <div className="config-row">
          <FG label="Standard Rate Override (₹/MT)" note="Global fallback if grade rate not set">
            <input type="number" placeholder="7500" value={costConfig.stdSellingRate||""} onChange={e=>set("stdSellingRate")(e.target.value)}/>
          </FG>
        </div>
        <div style={{padding:"10px 14px",background:"var(--bg3)",borderRadius:"var(--r)",fontSize:12,color:"var(--text3)"}}>
          Tip: Set grade-specific selling rates in <strong style={{color:"var(--text2)"}}>Masters → Grades</strong> for accurate per-grade margin calculations.
        </div>
      </div>

      <div style={{background:"var(--bg3)",border:"1px solid var(--border)",borderRadius:"var(--r2)",padding:"14px 18px"}}>
        <div style={{fontSize:11,fontWeight:600,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".07em",marginBottom:10,fontFamily:"var(--mono)"}}>Current Configuration</div>
        {[
          ["Bag Cost",         costConfig.bagCostPerTonne   ? `₹${costConfig.bagCostPerTonne}/tonne`    : "Not set"],
          ["Freight",          costConfig.freightPerTonne   ? `₹${costConfig.freightPerTonne}/MT`       : "Not set"],
          ["Loading Labour",   costConfig.loadingPerTonne   ? `₹${costConfig.loadingPerTonne}/MT`       : "Not set"],
          ["Monthly Lab Cost", costConfig.monthlyLabCost    ? `₹${(+costConfig.monthlyLabCost).toLocaleString("en-IN")}` : "Not set"],
          ["Plant Rent",       costConfig.monthlyRent       ? `₹${(+costConfig.monthlyRent).toLocaleString("en-IN")}/month` : "Not set"],
        ].map(([label,val])=>(
          <div key={label} className="cost-row">
            <span className="cost-label">{label}</span>
            <span className="cost-val" style={{color:val==="Not set"?"var(--text4)":"var(--accent)"}}>{val}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── MONTHLY COST SHEET ────────────────────────────────────────────────────────
function MonthlyCostSheet({purchases,expenses,productionEntries,weighments,grades,lots,costConfig}){
  const now=new Date();
  const [ym,setYm]=useState(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`);

  const c=computeMonthlyCosts(ym,purchases,expenses,productionEntries,weighments,costConfig);

  // Extra config costs
  const rentCost    = +costConfig.monthlyRent||0;
  const miscFixed   = +costConfig.miscFixed||0;
  const loadCost    = c.soldMT*(+costConfig.loadingPerTonne||0);
  const allFixed    = rentCost+miscFixed;
  const grandTotal  = c.totalCost+allFixed+loadCost;
  const cptFull     = c.prodMT>0 ? grandTotal/c.prodMT : 0;
  const cptSold     = c.soldMT>0  ? grandTotal/c.soldMT  : 0;

  // Grade-wise sold breakdown for the month
  const gradeMap={};
  weighments.filter(w=>w.date?.slice(0,7)===ym&&w.purpose==="Sale").forEach(w=>{
    const key=w.gradeCode||"Unknown";
    gradeMap[key]=(gradeMap[key]||{code:key,mt:0,revenue:0});
    const mt=(w.netWeight||0)/1000;
    const grade=grades.find(g=>g.code===key);
    const rate=+(grade?.stdRate||costConfig.stdSellingRate||0);
    gradeMap[key].mt+=mt;
    gradeMap[key].revenue+=mt*rate;
  });
  const gradeArr=Object.values(gradeMap);
  const totalRevenue=gradeArr.reduce((s,g)=>s+g.revenue,0);
  const grossMargin=totalRevenue-grandTotal;
  const marginPct=totalRevenue>0?((grossMargin/totalRevenue)*100).toFixed(1):null;

  // Lot P&L for this month
  const monthLots=lots.filter(l=>l.date?.slice(0,7)===ym);

  const costItems=[
    {label:"Raw Material (RM)",   val:c.rmCost,      color:"amber"},
    {label:"Staff Salaries",      val:c.salary,      color:"blue"},
    {label:"Contract / Casual Labour",val:c.labour,  color:"blue"},
    {label:"Electricity",         val:c.power,       color:"purple"},
    {label:"Water",                val:c.water,      color:"teal"},
    {label:"Fuel & Diesel",       val:c.fuel,        color:"orange"},
    {label:"Maintenance",         val:c.maint,       color:"red"},
    {label:"Transport (Expenses)",val:c.transport,   color:"teal"},
    {label:"Packing (1T Bags)",   val:c.packingCost, color:"green"},
    {label:"Loading Labour",      val:loadCost,      color:"green"},
    {label:"Freight (Outbound)",  val:c.freightCost, color:"blue"},
    {label:"Lab Testing",         val:c.labCost,     color:"purple"},
    {label:"Plant Rent",          val:rentCost,      color:"muted"},
    {label:"Misc Fixed",          val:miscFixed,     color:"muted"},
    {label:"Miscellaneous Exp.",  val:c.misc,        color:"muted"},
  ].filter(i=>i.val>0);

  const barColors={"amber":"var(--amber)","blue":"var(--blue)","purple":"var(--purple)",
    "orange":"var(--orange)","red":"var(--red)","teal":"var(--teal)",
    "green":"var(--green)","muted":"var(--text3)"};

  return(
    <div>
      <div style={{display:"flex",alignItems:"flex-end",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:10}}>
        <div>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Monthly Cost Sheet</h3>
          <p style={{fontSize:12,color:"var(--text3)"}}>All costs pulled from Purchase Register + Expenses + Config</p>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"flex-end"}}>
          <FG label="Month">
            <input type="month" value={ym} onChange={e=>setYm(e.target.value)}
              style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"8px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:13,outline:"none"}}/>
          </FG>
          <button className="btn btn-ghost btn-sm" onClick={()=>window.print()}>🖨 Print</button>
        </div>
      </div>

      {/* CPT Banner */}
      <div className="cpt-banner">
        <div>
          <div style={{fontSize:11,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"var(--mono)",marginBottom:6}}>Cost per Tonne Produced</div>
          <div className="cpt-val">₹{cptFull.toLocaleString("en-IN",{maximumFractionDigits:0})}</div>
          <div className="cpt-sub">Based on {c.prodMT.toFixed(2)} MT produced in {ym}</div>
        </div>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:11,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"var(--mono)",marginBottom:6}}>Cost per Tonne Sold</div>
          <div style={{fontSize:24,fontWeight:700,color:"var(--blue)",fontFamily:"var(--mono)"}}>₹{cptSold.toLocaleString("en-IN",{maximumFractionDigits:0})}</div>
          <div className="cpt-sub">Based on {c.soldMT.toFixed(2)} MT sold</div>
        </div>
        {marginPct&&(
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:11,color:"var(--text3)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"var(--mono)",marginBottom:6}}>Gross Margin</div>
            <div style={{fontSize:24,fontWeight:700,color:grossMargin>=0?"var(--green)":"var(--red)",fontFamily:"var(--mono)"}}>
              {marginPct}%
            </div>
            <div className="cpt-sub" style={{color:grossMargin>=0?"var(--green)":"var(--red)"}}>
              ₹{Math.abs(grossMargin).toLocaleString("en-IN",{maximumFractionDigits:0})} {grossMargin>=0?"profit":"loss"}
            </div>
          </div>
        )}
      </div>

      <div className="two-col">
        {/* Cost Breakdown */}
        <div className="table-wrap" style={{marginBottom:0,padding:"18px 20px"}}>
          <div style={{fontWeight:600,fontSize:13,marginBottom:16}}>Cost Breakdown — {ym}</div>
          {costItems.length===0
            ?<EmptyState icon="💰" message="No cost data" sub="Add purchases and expenses to see breakdown"/>
            :costItems.map(item=>{
              const pct=grandTotal>0?((item.val/grandTotal)*100).toFixed(1):0;
              return(
                <div key={item.label} style={{marginBottom:10}}>
                  <div className="cost-row" style={{border:"none",padding:"4px 0"}}>
                    <span className="cost-label">{item.label}</span>
                    <span className="cost-val">₹{item.val.toLocaleString("en-IN",{maximumFractionDigits:0})} <span className="cost-pct">({pct}%)</span></span>
                  </div>
                  <div style={{background:"var(--bg4)",borderRadius:3,height:4}}>
                    <div style={{width:`${pct}%`,height:"100%",borderRadius:3,background:barColors[item.color]||"var(--text3)",transition:"width .4s"}}/>
                  </div>
                </div>
              );
            })
          }
          <div className="cost-row grand" style={{marginTop:8}}>
            <span className="cost-label">Total Cost</span>
            <span className="cost-val accent">₹{grandTotal.toLocaleString("en-IN",{maximumFractionDigits:0})}</span>
          </div>
        </div>

        {/* Revenue & Margin */}
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          <div className="table-wrap" style={{marginBottom:0,padding:"18px 20px"}}>
            <div style={{fontWeight:600,fontSize:13,marginBottom:14}}>Revenue by Grade — {ym}</div>
            {gradeArr.length===0
              ?<EmptyState icon="📤" message="No sales this month"/>
              :gradeArr.map(g=>(
                <div key={g.code} className="cost-row">
                  <span className="cost-label"><Badge type="muted">{g.code}</Badge> <span style={{marginLeft:6}}>{g.mt.toFixed(3)} MT</span></span>
                  <span className="cost-val" style={{color:"var(--accent)"}}>₹{g.revenue.toLocaleString("en-IN",{maximumFractionDigits:0})}</span>
                </div>
              ))
            }
            {totalRevenue>0&&(
              <div className="cost-row grand">
                <span className="cost-label">Total Revenue</span>
                <span className="cost-val accent">₹{totalRevenue.toLocaleString("en-IN",{maximumFractionDigits:0})}</span>
              </div>
            )}
            {totalRevenue>0&&(
              <div className={`margin-row ${grossMargin>=0?"profit":"loss"}`} style={{marginTop:12}}>
                <span style={{fontSize:13,fontWeight:700,color:grossMargin>=0?"var(--green)":"var(--red)",fontFamily:"var(--mono)"}}>
                  {grossMargin>=0?"↑ Profit":"↓ Loss"}: ₹{Math.abs(grossMargin).toLocaleString("en-IN",{maximumFractionDigits:0})}
                </span>
                {marginPct&&<span style={{fontSize:11,color:"var(--text3)",fontFamily:"var(--mono)"}}>({marginPct}% margin)</span>}
              </div>
            )}
          </div>

          {/* Key Metrics */}
          <div className="table-wrap" style={{marginBottom:0,padding:"18px 20px"}}>
            <div style={{fontWeight:600,fontSize:13,marginBottom:14}}>Key Metrics</div>
            {[
              ["RM as % of Total Cost",   grandTotal>0?`${((c.rmCost/grandTotal)*100).toFixed(1)}%`:"--",     "amber"],
              ["Labour as % of Cost",     grandTotal>0?`${((c.salaries/grandTotal)*100).toFixed(1)}%`:"--",   "blue"],
              ["Power as % of Cost",      grandTotal>0?`${((c.power/grandTotal)*100).toFixed(1)}%`:"--",      "purple"],
              ["Packing + Freight / MT",  c.soldMT>0?`₹${(((c.packingCost+c.freightCost+loadCost)/c.soldMT)).toFixed(0)}/MT`:"--", "green"],
              ["RM Cost / MT Produced",   c.prodMT>0?`₹${(c.rmCost/c.prodMT).toFixed(0)}/MT`:"--",           "amber"],
              ["Fixed Cost / MT",         c.prodMT>0?`₹${((c.salaries+c.power+c.water+c.maint+rentCost)/c.prodMT).toFixed(0)}/MT`:"--","blue"],
            ].map(([label,val,color])=>(
              <div key={label} className="cost-row">
                <span className="cost-label">{label}</span>
                <span className="cost-val" style={{color:`var(--${color})`}}>{val}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* Lot P&L Table */}
      {monthLots.length>0&&(
        <div className="table-wrap">
          <div className="table-toolbar"><h3>Lot-wise P&L — {ym}</h3><span style={{fontSize:11,color:"var(--text3)"}}>CPT: ₹{cptFull.toFixed(0)}/MT</span></div>
          <table>
            <thead>
              <tr>
                <th>Lot #</th><th>Date</th><th className="r">Output (MT)</th>
                <th className="r">Est. Cost</th><th className="r">Est. Revenue</th>
                <th className="r">Gross Margin</th><th className="r">₹/MT</th><th>Status</th>
              </tr>
            </thead>
            <tbody>
              {monthLots.map(lot=>{
                const lotMT=+(lot.totalOutputMT||0);
                const lotCost=lotMT*cptFull;
                // Estimate revenue from rate card
                const lotRevenue=(lot.gradeOutputs||[]).reduce((s,o)=>{
                  const g=grades.find(gx=>gx.code===o.gradeCode);
                  const rate=+(g?.stdRate||costConfig.stdSellingRate||0);
                  return s+(+(o.quantityMT||0))*rate;
                },0);
                const lotMargin=lotRevenue-lotCost;
                const lotMarginPMT=lotMT>0?lotMargin/lotMT:0;
                const hasRevenue=lotRevenue>0;
                return(
                  <tr key={lot.id}>
                    <td><span className="lot-chip">🏷 {lot.lotNo}</span></td>
                    <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{lot.date}</td>
                    <td className="r mono">{lotMT.toFixed(3)}</td>
                    <td className="r mono" style={{color:"var(--red)"}}>{cptFull>0?`₹${lotCost.toLocaleString("en-IN",{maximumFractionDigits:0})}`:"--"}</td>
                    <td className="r mono" style={{color:"var(--accent)"}}>{hasRevenue?`₹${lotRevenue.toLocaleString("en-IN",{maximumFractionDigits:0})}`:<span style={{color:"var(--text4)"}}>Set rate card</span>}</td>
                    <td className="r mono" style={{color:hasRevenue?(lotMargin>=0?"var(--green)":"var(--red)"):"var(--text4)",fontWeight:700}}>
                      {hasRevenue?`₹${lotMargin.toLocaleString("en-IN",{maximumFractionDigits:0})}`:"--"}
                    </td>
                    <td className="r mono" style={{color:hasRevenue?(lotMarginPMT>=0?"var(--green)":"var(--red)"):"var(--text4)"}}>
                      {hasRevenue?`₹${lotMarginPMT.toFixed(0)}/MT`:"--"}
                    </td>
                    <td><Badge type={{"in-stock":"in-stock","under-test":"under-test","released":"released","dispatched":"dispatched","rejected":"rejected"}[lot.status]||"muted"}>{lot.status}</Badge></td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="ledger-foot">
                <td colSpan={2} style={{padding:"10px 16px",fontSize:12,color:"var(--text2)"}}>Month Total</td>
                <td className="r mono" style={{padding:"10px 16px"}}>{monthLots.reduce((s,l)=>s+(+l.totalOutputMT||0),0).toFixed(3)} MT</td>
                <td className="r mono" style={{padding:"10px 16px",color:"var(--red)"}}>₹{(monthLots.reduce((s,l)=>s+(+(l.totalOutputMT||0))*cptFull,0)).toLocaleString("en-IN",{maximumFractionDigits:0})}</td>
                <td colSpan={4} style={{padding:"10px 16px"}}></td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}

      {c.prodMT===0&&c.soldMT===0&&<div className="report-section"><EmptyState icon="💰" message={`No data for ${ym}`} sub="Add purchases and expenses to see cost breakdown"/></div>}
    </div>
  );
}

// ── CAPITAL / INFRA INVESTMENT REGISTER ──────────────────────────────────────
// Tracks one-time setup / capital spend since incorporation — machinery, civil
// work, vehicles, licenses, deposits — separately from the monthly operating
// costs above. Gives a running "total invested in the business" figure.
const CAP_CATS=[
  {id:"machinery", label:"Machinery & Equipment",   color:"#c8f064"},
  {id:"civil",      label:"Land / Civil / Construction",color:"#a78bfa"},
  {id:"vehicle",    label:"Vehicles",                color:"#5ca0ff"},
  {id:"license",    label:"Licenses & Registration",color:"#2dd4bf"},
  {id:"deposit",    label:"Deposits & Utility Setup",color:"#f5a623"},
  {id:"office",     label:"Furniture & Office Setup",color:"#ff5c5c"},
  {id:"other",      label:"Other Capital Spend",    color:"#999"},
];
function CapitalRegister({capitalItems,setCapitalItems}){
  const [open,setOpen]=useState(false);
  const blank={date:today(),category:"machinery",description:"",amount:"",paidTo:"",reference:"",fundedBy:"own"};
  const [f,setF]=useState(blank);
  const set=k=>v=>setF(x=>({...x,[k]:v}));

  function save(){
    if(!f.amount||!f.description)return alert("Amount and Description required.");
    setCapitalItems(cs=>[{id:uid(),capNo:`CAP-${String(cs.length+1).padStart(4,"0")}`,...f,amount:+f.amount},...cs]);
    setOpen(false);setF(blank);
  }

  const totalInvested=capitalItems.reduce((s,c)=>s+c.amount,0);
  const ownFunded=capitalItems.filter(c=>c.fundedBy==="own").reduce((s,c)=>s+c.amount,0);
  const loanFunded=capitalItems.filter(c=>c.fundedBy==="loan").reduce((s,c)=>s+c.amount,0);
  const catBreak=CAP_CATS.map(c=>({...c,total:capitalItems.filter(x=>x.category===c.id).reduce((s,x)=>s+x.amount,0)})).filter(c=>c.total>0).sort((a,b)=>b.total-a.total);
  const maxCat=catBreak[0]?.total||1;
  const sorted=[...capitalItems].sort((a,b)=>(a.date<b.date?1:-1));

  return(
    <div>
      <div style={{marginBottom:20}}>
        <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Capital / Infra Investment</h3>
        <p style={{fontSize:12,color:"var(--text3)"}}>One-time setup spend since incorporation — machinery, civil work, vehicles, licenses, deposits — kept separate from monthly operating costs. Share your bank statement and these entries can be reconstructed from it rather than typed in one by one.</p>
      </div>

      <div className="stats-grid">
        <StatCard label="Total Invested" value={fmt(totalInvested)}  sub={`${capitalItems.length} entries`} color="accent"/>
        <StatCard label="Own Capital"    value={fmt(ownFunded)}      sub="self-funded"                       color="green"/>
        <StatCard label="Loan / Credit"  value={fmt(loanFunded)}     sub="borrowed"                          color="amber"/>
        <StatCard label="Top Category"   value={catBreak[0]?.label||"—"} sub={catBreak[0]?fmt(catBreak[0].total):"no data"} color="purple"/>
      </div>

      <div className="report-section" style={{marginBottom:20}}>
        <h3>Breakdown by Category</h3>
        {catBreak.length===0
          ?<EmptyState icon="🏗" message="No capital spend logged yet"/>
          :catBreak.map(c=>(
            <div key={c.id} className="mini-bar-row">
              <span className="label"><span className="expense-cat-dot" style={{background:c.color}}/>{c.label}</span>
              <div className="track"><div className="fill" style={{width:`${(c.total/maxCat)*100}%`,background:c.color}}/></div>
              <span className="val">{fmt(c.total)}</span>
            </div>
          ))
        }
      </div>

      <div className="filter-bar">
        <button className="btn btn-primary" onClick={()=>setOpen(true)}>+ Add Capital Item</button>
      </div>

      <div className="table-wrap">
        <table>
          <thead><tr><th>Cap #</th><th>Date</th><th>Category</th><th>Description</th><th>Paid To</th><th>Funded By</th><th>Ref #</th><th className="r">Amount</th><th></th></tr></thead>
          <tbody>
            {sorted.length===0
              ?<tr><td colSpan={9}><EmptyState icon="🏗" message="No capital items" sub="Add machinery, construction, vehicles, licenses, deposits, etc."/></td></tr>
              :sorted.map(c=>{
                const cat=CAP_CATS.find(x=>x.id===c.category);
                return(
                  <tr key={c.id}>
                    <td className="mono" style={{color:"var(--accent)"}}>{c.capNo}</td>
                    <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{c.date}</td>
                    <td><span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:12}}><span className="expense-cat-dot" style={{background:cat?.color}}/>{cat?.label}</span></td>
                    <td style={{fontWeight:500}}>{c.description}</td>
                    <td style={{color:"var(--text2)"}}>{c.paidTo||"—"}</td>
                    <td><Badge type={c.fundedBy==="loan"?"amber":"green"}>{c.fundedBy==="loan"?"Loan":"Own"}</Badge></td>
                    <td className="mono" style={{color:"var(--text3)",fontSize:11}}>{c.reference||"—"}</td>
                    <td className="r mono" style={{color:"var(--accent)",fontWeight:700}}>{fmt(c.amount)}</td>
                    <td><button className="btn btn-danger btn-sm" onClick={()=>setCapitalItems(cs=>cs.filter(x=>x.id!==c.id))}>Del</button></td>
                  </tr>
                );
              })
            }
          </tbody>
          {sorted.length>0&&(
            <tfoot>
              <tr className="ledger-foot">
                <td colSpan={7} style={{padding:"10px 16px",fontSize:12,color:"var(--text2)"}}>Total Capital Invested</td>
                <td className="r mono" style={{padding:"10px 16px",color:"var(--accent)",fontWeight:700}}>{fmt(totalInvested)}</td>
                <td></td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {open&&(
        <Modal title="Add Capital Item" onClose={()=>setOpen(false)}
          foot={<><button className="btn btn-ghost" onClick={()=>setOpen(false)}>Cancel</button><button className="btn btn-primary" onClick={save}>Save Item</button></>}>
          <div className="form-row cols-2">
            <FG label="Date *"><input type="date" value={f.date} onChange={e=>set("date")(e.target.value)}/></FG>
            <FG label="Category">
              <select value={f.category} onChange={e=>set("category")(e.target.value)}>
                {CAP_CATS.map(c=><option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Amount (₹) *"><input type="number" step="0.01" placeholder="150000" value={f.amount} onChange={e=>set("amount")(e.target.value)}/></FG>
            <FG label="Paid To"><input placeholder="Vendor / Contractor" value={f.paidTo} onChange={e=>set("paidTo")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Description *"><input placeholder="e.g. Ball Mill #2 purchase" value={f.description} onChange={e=>set("description")(e.target.value)}/></FG>
            <FG label="Reference / Invoice #"><input placeholder="Invoice / bill no." value={f.reference} onChange={e=>set("reference")(e.target.value)}/></FG>
          </div>
          <div className="form-row cols-2">
            <FG label="Funded By">
              <select value={f.fundedBy} onChange={e=>set("fundedBy")(e.target.value)}>
                <option value="own">Own Capital</option>
                <option value="loan">Loan / Credit</option>
              </select>
            </FG>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ── BANK STATEMENT → COUNTERPARTY GROUPING ───────────────────────────────────
// Upload a statement (CSV/XLSX), map its columns once, and every transaction is
// bucketed by a cleaned-up version of its narration (UPI/NEFT/ref numbers, IFSCs,
// dates stripped out) so the same payee/payer's transactions land in one group.
// Naming a group sticks (bankLabels is keyed by that cleaned narration), so the
// next statement you import recognises payees you've already named.
const BANK_STOPWORDS=new Set(["UPI","NEFT","IMPS","RTGS","ACH","ACH DR","ACH CR","POS","ATM","DR","CR","WDL","TRF","TRANSFER","BIL","MOB","INB","ECOM","ECS","ONL","P2A","P2M","P2PM","TXN","REF","VPA","PAYMENT","AUTOPAY","MANDATE","NACH"]);
function normalizeDesc(desc){
  const raw=(desc||"").toUpperCase();
  const tokens=raw.split(/[-\/|,]+/).map(t=>t.trim()).filter(Boolean);
  const kept=tokens.filter(t=>{
    if(BANK_STOPWORDS.has(t))return false;
    if(/^\d+$/.test(t)&&t.length>=4)return false;           // account / ref / phone numbers
    if(t.length>=10){const dc=(t.match(/\d/g)||[]).length;if(dc/t.length>=0.55)return false;} // bank UTR/ref codes glued onto text (e.g. BARBR52025072300986781)
    if(/^[A-Z]{4}0[A-Z0-9]{6}$/.test(t))return false;        // IFSC code
    if(/^\d{1,2}[A-Z]{3}\d{2,4}$/.test(t))return false;      // e.g. 12JAN24
    if(/^\d+@[A-Z0-9.]+$/.test(t))return false;               // UPI VPA (phone@handle)
    if(t.length<=1)return false;
    return true;
  });
  const result=(kept.length?kept:tokens).join(" ").replace(/\s+/g," ").trim();
  return result||"UNKNOWN COUNTERPARTY";
}
const parseAmt=v=>{if(v===null||v===undefined||v==="")return 0;const n=parseFloat(String(v).replace(/[₹,\s]/g,""));return isNaN(n)?0:Math.abs(n);};
function parseBankDate(v){
  if(v===null||v===undefined||v==="")return "";
  if(typeof v==="number"){const ms=Date.UTC(1899,11,30)+v*86400000;return new Date(ms).toISOString().slice(0,10);}
  const s=String(v).trim();
  let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m)return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if(m){let[,d,mo,y]=m;if(y.length===2)y="20"+y;return `${y}-${mo.padStart(2,"0")}-${d.padStart(2,"0")}`;}
  return s.slice(0,10);
}
function guessCol(headers,patterns){
  const low=headers.map(h=>String(h).toLowerCase());
  for(const p of patterns){const i=low.findIndex(h=>h.includes(p));if(i>-1)return headers[i];}
  return "";
}
const CP_TYPES=[
  {id:"unlabeled",label:"Unlabeled",  color:"#999"},
  {id:"vendor",    label:"Vendor / Supplier", color:"#f5a623"},
  {id:"customer",  label:"Customer / Buyer",  color:"#4ecb71"},
  {id:"staff",     label:"Staff / Labour",    color:"#5ca0ff"},
  {id:"owner",     label:"Owner / Capital",   color:"#c8f064"},
  {id:"loan",      label:"Loan / Bank",       color:"#a78bfa"},
  {id:"utility",   label:"Utility / Rent",    color:"#2dd4bf"},
  {id:"tax",       label:"Tax / Government",  color:"#ff5c5c"},
  {id:"other",     label:"Other",             color:"#666"},
];

// ── PDF STATEMENT PARSING (coordinate-based) ─────────────────────────────────
// Bank statement PDFs have no real columns — just text positioned on a page —
// so instead of regex-scraping the reflowed text, this reads each word's x/y
// position (via pdfjs-dist), reconstructs table rows by clustering words with
// the same y-position, and derives column boundaries FROM the statement's own
// header row ("Date | Narration | Chq/Ref No | Value Dt | Withdrawal Amt |
// Deposit Amt | Closing Balance") rather than hardcoding pixel positions — so
// it adapts to slightly different page sizes/layouts as long as that header
// row is present. Multi-line narrations and transactions that wrap across a
// page break are handled by treating "a line starting with a DD/MM/YY date in
// the Date column" as the start of a new transaction, and everything after it
// (until the next such line) as that transaction's overflow.
// Validated against a real HDFC statement: reproduces the statement's own
// "Statement Summary" totals (txn count, debit/credit totals, closing balance)
// exactly. Other banks' layouts use the same general structure but different
// column labels — if a statement doesn't parse, its header row wording likely
// isn't recognised yet (see HEADER_KEYS below).
let _pdfjsPromise=null;
function loadPdfjs(){
  if(!_pdfjsPromise){
    _pdfjsPromise=import("pdfjs-dist").then(lib=>{
      lib.GlobalWorkerOptions.workerSrc=`https://cdn.jsdelivr.net/npm/pdfjs-dist@${lib.version}/build/pdf.worker.min.mjs`;
      return lib;
    });
  }
  return _pdfjsPromise;
}
async function parsePdfStatement(file){
  const pdfjsLib=await loadPdfjs();
  const buf=await file.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({data:buf}).promise;
  const allWords=[]; let pageWidth=612;
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const viewport=page.getViewport({scale:1});
    if(p===1)pageWidth=viewport.width;
    const content=await page.getTextContent();
    content.items.forEach(it=>{
      const t=(it.str||"").trim();
      if(!t)return;
      allWords.push({page:p,x0:it.transform[4],top:viewport.height-it.transform[5],text:t});
    });
  }
  const sorted=[...allWords].sort((a,b)=>a.page-b.page||a.top-b.top||a.x0-b.x0);
  const lines=[]; let cur=[];
  sorted.forEach(w=>{
    if(cur.length&&(w.page!==cur[0].page||Math.abs(w.top-cur[cur.length-1].top)>4.5)){lines.push(cur);cur=[];}
    cur.push(w);
  });
  if(cur.length)lines.push(cur);
  const lineText=line=>line.map(w=>w.text).join("").replace(/\s+/g,"");

  let colBounds=null;
  for(const line of lines){
    if(lineText(line).startsWith("DateNarrationChq")){
      // Header labels confirm this is a recognised layout, but their x-positions
      // aren't reliable column edges (e.g. the "Narration" label sits well right
      // of where narration text actually starts) — so scale known-good boundaries
      // (validated against real statements) to this page's width instead.
      const REF=[[0,65],[65,278],[278,358],[358,404],[404,489],[489,563],[563,638]];
      const order=["date","narr","ref","vdt","wd","dep","bal"];
      const k=pageWidth/638;
      colBounds={};
      order.forEach((c,i)=>{colBounds[c]=[REF[i][0]*k,i===order.length-1?pageWidth+50:REF[i][1]*k];});
      break;
    }
  }
  if(!colBounds)throw new Error("Couldn't find a recognised statement table header on this PDF.");

  const wordsInCol=(line,col)=>{const[lo,hi]=colBounds[col];return line.filter(w=>w.x0>=lo&&w.x0<hi).sort((a,b)=>a.x0-b.x0);};

  const filtered=[]; let inTable=false;
  for(const line of lines){
    const txt=lineText(line);
    if(txt.includes("Statementofaccount")){inTable=true;continue;}
    if(txt.includes("HDFCBANKLIMITED")||txt.toUpperCase().startsWith("STATEMENTSUMMARY")){inTable=false;continue;}
    if(txt.startsWith("DateNarrationChq"))continue;
    if(inTable)filtered.push(line);
  }

  const parseAmtLocal=s=>{const n=parseFloat(String(s).replace(/,/g,""));return isNaN(n)?null:n;};
  const txns=[]; let t=null;
  for(const line of filtered){
    const dwords=wordsInCol(line,"date");
    const isNew=dwords.length>0&&/^\d{2}\/\d{2}\/\d{2}$/.test(dwords[0].text);
    if(isNew){
      if(t)txns.push(t);
      t={date:dwords[0].text,narr:[],ref:[],vdt:"",wd:null,dep:null};
      wordsInCol(line,"narr").forEach(w=>t.narr.push(w.text));
      wordsInCol(line,"ref").forEach(w=>t.ref.push(w.text));
      const vd=wordsInCol(line,"vdt"); if(vd.length)t.vdt=vd[0].text;
      const wd=wordsInCol(line,"wd"); if(wd.length)t.wd=parseAmtLocal(wd[0].text);
      const dep=wordsInCol(line,"dep"); if(dep.length)t.dep=parseAmtLocal(dep[0].text);
    }else{
      if(!t)continue;
      wordsInCol(line,"narr").forEach(w=>t.narr.push(w.text));
      wordsInCol(line,"ref").forEach(w=>t.ref.push(w.text));
    }
  }
  if(t)txns.push(t);
  if(!txns.length)throw new Error("No transaction rows recognised in this PDF.");

  const rows=txns.map(tx=>({
    "Date":tx.date,
    "Narration":tx.narr.join(" "),
    "Chq/Ref No":tx.ref.join(""),
    "Value Dt":tx.vdt,
    "Withdrawal Amt":tx.wd!=null?tx.wd:"",
    "Deposit Amt":tx.dep!=null?tx.dep:"",
  }));
  return{headers:["Date","Narration","Chq/Ref No","Value Dt","Withdrawal Amt","Deposit Amt"],rows,fileName:file.name};
}

function BankStatementGrouping({bankTxns,setBankTxns,bankLabels,setBankLabels}){
  const [sheet,setSheet]=useState(null);       // {headers, rows, fileName}
  const [map,setMap]=useState({dateCol:"",descCol:"",mode:"separate",debitCol:"",creditCol:"",amountCol:"",typeCol:""});
  const [search,setSearch]=useState("");
  const [unlabeledOnly,setUnlabeledOnly]=useState(false);
  const [expanded,setExpanded]=useState(null);

  function onFile(e){
    const file=e.target.files[0]; if(!file)return;
    e.target.value="";
    if(file.name.toLowerCase().endsWith(".pdf")){
      parsePdfStatement(file).then(setSheet).catch(err=>{
        console.error(err);
        alert("Couldn't read this PDF automatically — "+(err.message||"it may use a different statement layout than the ones this parser recognises.")+" Try exporting the statement as CSV/XLSX from net-banking instead.");
      });
      return;
    }
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!rows.length){alert("No rows found — check the file has a header row and at least one transaction.");return;}
        const headers=Object.keys(rows[0]);
        setSheet({headers,rows,fileName:file.name});
        setMap({
          dateCol:guessCol(headers,["date"]),
          descCol:guessCol(headers,["narration","description","particular","remark","detail"]),
          mode:"separate",
          debitCol:guessCol(headers,["debit","withdrawal"]),
          creditCol:guessCol(headers,["credit","deposit"]),
          amountCol:guessCol(headers,["amount"]),
          typeCol:guessCol(headers,["type","dr/cr","dr / cr"]),
        });
      }catch(err){alert("Couldn't read that file. Export the statement as CSV or XLSX and try again.");}
    };
    reader.readAsArrayBuffer(file);
  }

  function importRows(){
    if(!sheet||!map.dateCol||!map.descCol)return alert("Map at least the Date and Description columns.");
    const batchId=uid();
    const parsed=sheet.rows.map(r=>{
      const date=parseBankDate(r[map.dateCol]);
      const description=String(r[map.descCol]||"").trim();
      let debit=0,credit=0;
      if(map.mode==="separate"){
        debit=parseAmt(r[map.debitCol]);
        credit=parseAmt(r[map.creditCol]);
      }else{
        const amt=parseAmt(r[map.amountCol]);
        const typeVal=String(r[map.typeCol]||"").toUpperCase();
        if(typeVal.includes("CR")||typeVal.includes("DEPOSIT"))credit=amt;
        else if(typeVal.includes("DR")||typeVal.includes("WITHDRAW"))debit=amt;
        else{
          const raw=r[map.amountCol];
          if(String(raw).trim().startsWith("-"))debit=amt; else credit=amt;
        }
      }
      return{date,description,debit,credit};
    }).filter(r=>r.date&&r.description&&(r.debit>0||r.credit>0));
    if(!parsed.length)return alert("No usable transactions after mapping — double-check the column choices against the preview.");
    setBankTxns(txns=>{
      const seen=new Set(txns.map(t=>`${t.date}|${t.description}|${t.debit}|${t.credit}`));
      const fresh=parsed.filter(r=>!seen.has(`${r.date}|${r.description}|${r.debit}|${r.credit}`))
        .map(r=>({id:uid(),batchId,source:sheet.fileName,importedAt:today(),key:normalizeDesc(r.description),...r}));
      return[...txns,...fresh];
    });
    setSheet(null);
    setMap({dateCol:"",descCol:"",mode:"separate",debitCol:"",creditCol:"",amountCol:"",typeCol:""});
  }

  const batches=useMemo(()=>{
    const m={};
    bankTxns.forEach(t=>{if(!m[t.batchId])m[t.batchId]={id:t.batchId,source:t.source,importedAt:t.importedAt,count:0};m[t.batchId].count++;});
    return Object.values(m).sort((a,b)=>b.importedAt.localeCompare(a.importedAt));
  },[bankTxns]);

  const groups=useMemo(()=>{
    const m={};
    bankTxns.forEach(t=>{
      if(!m[t.key])m[t.key]={key:t.key,txns:[],totalDebit:0,totalCredit:0};
      m[t.key].txns.push(t); m[t.key].totalDebit+=t.debit; m[t.key].totalCredit+=t.credit;
    });
    return Object.values(m).map(g=>({...g,label:bankLabels[g.key]?.label||"",type:bankLabels[g.key]?.type||"unlabeled",net:g.totalCredit-g.totalDebit}))
      .sort((a,b)=>(b.totalDebit+b.totalCredit)-(a.totalDebit+a.totalCredit));
  },[bankTxns,bankLabels]);

  // Two raw narration-groups (different keys) that end up with the SAME typed
  // name collapse into one card here — so labelling both "Ramesh Traders"
  // is how you merge a payee that split across two groups.
  const mergedGroups=useMemo(()=>{
    const m={};
    groups.forEach(g=>{
      const mergeKey=g.label.trim()?g.label.trim().toLowerCase():`__raw__${g.key}`;
      if(!m[mergeKey])m[mergeKey]={mergeKey,label:g.label,type:g.type,rawKeys:[],txns:[],totalDebit:0,totalCredit:0};
      m[mergeKey].rawKeys.push(g.key);
      m[mergeKey].txns.push(...g.txns);
      m[mergeKey].totalDebit+=g.totalDebit;
      m[mergeKey].totalCredit+=g.totalCredit;
    });
    return Object.values(m).map(mg=>({...mg,net:mg.totalCredit-mg.totalDebit}))
      .sort((a,b)=>(b.totalDebit+b.totalCredit)-(a.totalDebit+a.totalCredit));
  },[groups]);

  const filtered=mergedGroups.filter(g=>{
    if(unlabeledOnly&&g.label)return false;
    if(!search)return true;
    const s=search.toLowerCase();
    return g.rawKeys.some(k=>k.toLowerCase().includes(s))||g.label.toLowerCase().includes(s)||g.txns.some(t=>t.description.toLowerCase().includes(s));
  });

  // patch applies to every raw key folded into this card, so a merged
  // group's name/type stays in sync across all the narration variants it covers
  function updateLabel(rawKeys,patch){
    setBankLabels(bl=>{
      const next={...bl};
      rawKeys.forEach(k=>{next[k]={...next[k],...patch};});
      return next;
    });
  }

  const totalPaid=bankTxns.reduce((s,t)=>s+t.debit,0);
  const totalRecv=bankTxns.reduce((s,t)=>s+t.credit,0);
  const namedGroups=mergedGroups.filter(g=>g.label).length;

  return(
    <div>
      <div style={{marginBottom:20}}>
        <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Bank Statement — Counterparty Grouping</h3>
        <p style={{fontSize:12,color:"var(--text3)"}}>Upload a statement and transactions get bucketed by who they were to/from. Name a group once — "Ramesh Traders", "Staff Salary" — and it stays recognised on every future import.</p>
      </div>

      <div className="config-card" style={{marginBottom:20}}>
        <h3>Import Statement</h3>
        {!sheet ? (
          <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
            <label className="btn btn-primary" style={{cursor:"pointer"}}>
              Choose PDF / CSV / XLSX File
              <input type="file" accept=".pdf,.csv,.xlsx,.xls" onChange={onFile} style={{display:"none"}}/>
            </label>
            <span style={{fontSize:11.5,color:"var(--text3)"}}>Bank statement PDF (HDFC format tested) or a CSV/XLSX export from net-banking. First row must be the column headers for CSV/XLSX.</span>
          </div>
        ):(
          <div>
            <div style={{fontSize:12,color:"var(--text2)",marginBottom:12}}>{sheet.fileName} — {sheet.rows.length} rows found. Map the columns:</div>
            <div className="config-row">
              <FG label="Date Column *">
                <select value={map.dateCol} onChange={e=>setMap(m=>({...m,dateCol:e.target.value}))}>
                  <option value="">— select —</option>
                  {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </FG>
              <FG label="Description / Narration Column *">
                <select value={map.descCol} onChange={e=>setMap(m=>({...m,descCol:e.target.value}))}>
                  <option value="">— select —</option>
                  {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </FG>
            </div>
            <div className="config-row" style={{marginTop:14}}>
              <FG label="Amount Format">
                <select value={map.mode} onChange={e=>setMap(m=>({...m,mode:e.target.value}))}>
                  <option value="separate">Separate Debit & Credit columns</option>
                  <option value="single">Single Amount column</option>
                </select>
              </FG>
            </div>
            {map.mode==="separate" ? (
              <div className="config-row" style={{marginTop:14}}>
                <FG label="Debit / Withdrawal Column *">
                  <select value={map.debitCol} onChange={e=>setMap(m=>({...m,debitCol:e.target.value}))}>
                    <option value="">— select —</option>
                    {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </FG>
                <FG label="Credit / Deposit Column *">
                  <select value={map.creditCol} onChange={e=>setMap(m=>({...m,creditCol:e.target.value}))}>
                    <option value="">— select —</option>
                    {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </FG>
              </div>
            ):(
              <div className="config-row" style={{marginTop:14}}>
                <FG label="Amount Column *">
                  <select value={map.amountCol} onChange={e=>setMap(m=>({...m,amountCol:e.target.value}))}>
                    <option value="">— select —</option>
                    {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </FG>
                <FG label="Dr/Cr Indicator Column (optional)" note="Leave blank if a negative amount already means a withdrawal">
                  <select value={map.typeCol} onChange={e=>setMap(m=>({...m,typeCol:e.target.value}))}>
                    <option value="">— none —</option>
                    {sheet.headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </FG>
              </div>
            )}
            <div className="table-wrap" style={{marginTop:14,maxHeight:220,overflowY:"auto"}}>
              <table>
                <thead><tr>{sheet.headers.map(h=><th key={h}>{h}</th>)}</tr></thead>
                <tbody>{sheet.rows.slice(0,5).map((r,i)=>(<tr key={i}>{sheet.headers.map(h=><td key={h} style={{fontSize:11}}>{String(r[h])}</td>)}</tr>))}</tbody>
              </table>
            </div>
            <div style={{display:"flex",gap:10,marginTop:14}}>
              <button className="btn btn-ghost" onClick={()=>{setSheet(null);}}>Cancel</button>
              <button className="btn btn-primary" onClick={importRows}>Import Transactions</button>
            </div>
          </div>
        )}
      </div>

      {batches.length>0&&(
        <div className="report-section" style={{marginBottom:20}}>
          <h3>Imported Statements</h3>
          {batches.map(b=>(
            <div key={b.id} style={{display:"flex",alignItems:"center",gap:12,padding:"7px 0",borderBottom:"1px solid var(--border)",fontSize:12}}>
              <span style={{flex:1,color:"var(--text2)"}}>{b.source} <span style={{color:"var(--text3)"}}>· {b.importedAt}</span></span>
              <span style={{fontFamily:"var(--mono)",color:"var(--text3)"}}>{b.count} txns</span>
              <button className="btn btn-danger btn-sm" onClick={()=>setBankTxns(txns=>txns.filter(t=>t.batchId!==b.id))}>Remove</button>
            </div>
          ))}
        </div>
      )}

      <div className="stats-grid">
        <StatCard label="Transactions"  value={bankTxns.length} sub={`${batches.length} statement(s)`} color="blue"/>
        <StatCard label="Total Paid Out"value={fmt(totalPaid)}  sub="debits"   color="red"/>
        <StatCard label="Total Received"value={fmt(totalRecv)}  sub="credits"  color="green"/>
        <StatCard label="Groups Named"  value={`${namedGroups} / ${mergedGroups.length}`} sub="counterparties" color="purple"/>
      </div>

      {bankTxns.length>0&&(
        <>
          <div className="filter-bar">
            <input placeholder="Search by name or narration…" value={search} onChange={e=>setSearch(e.target.value)}
              style={{background:"var(--bg3)",border:"1px solid var(--border2)",borderRadius:"var(--r)",padding:"7px 11px",color:"var(--text)",fontFamily:"var(--font)",fontSize:12.5,outline:"none",minWidth:220}}/>
            <button className={`btn btn-sm ${unlabeledOnly?"btn-primary":"btn-ghost"}`} onClick={()=>setUnlabeledOnly(u=>!u)}>Unlabeled Only</button>
          </div>

          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {filtered.map(g=>{
              const typeInfo=CP_TYPES.find(t=>t.id===g.type);
              return(
                <div key={g.mergeKey} className="config-card" style={{padding:14}}>
                  <div style={{display:"flex",gap:12,flexWrap:"wrap",alignItems:"flex-end",justifyContent:"space-between"}}>
                    <div style={{display:"flex",gap:12,flexWrap:"wrap",flex:1,minWidth:280}}>
                      <FG label="Name this group">
                        <input placeholder={g.rawKeys[0]} value={g.label} onChange={e=>updateLabel(g.rawKeys,{label:e.target.value})}/>
                      </FG>
                      <FG label="Type">
                        <select value={g.type} onChange={e=>updateLabel(g.rawKeys,{type:e.target.value})}>
                          {CP_TYPES.map(t=><option key={t.id} value={t.id}>{t.label}</option>)}
                        </select>
                      </FG>
                    </div>
                    <div style={{display:"flex",gap:18,fontFamily:"var(--mono)",fontSize:12.5}}>
                      <div><div style={{color:"var(--text3)",fontSize:10}}>PAID</div><div style={{color:"var(--red)",fontWeight:700}}>{fmt(g.totalDebit)}</div></div>
                      <div><div style={{color:"var(--text3)",fontSize:10}}>RECEIVED</div><div style={{color:"var(--green)",fontWeight:700}}>{fmt(g.totalCredit)}</div></div>
                      <div><div style={{color:"var(--text3)",fontSize:10}}>TXNS</div><div style={{fontWeight:700}}>{g.txns.length}</div></div>
                    </div>
                  </div>
                  <div style={{marginTop:8,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontSize:11,color:"var(--text3)",fontFamily:"var(--mono)"}}>
                      {g.rawKeys.length>1?`Merged from ${g.rawKeys.length} narration groups`:g.rawKeys[0]}
                    </span>
                    <button className="btn btn-ghost btn-sm" onClick={()=>setExpanded(x=>x===g.mergeKey?null:g.mergeKey)}>{expanded===g.mergeKey?"Hide":"Show"} transactions</button>
                  </div>
                  {expanded===g.mergeKey&&(
                    <div className="table-wrap" style={{marginTop:10}}>
                      <table>
                        <thead><tr><th>Date</th><th>Narration</th><th className="r">Debit</th><th className="r">Credit</th></tr></thead>
                        <tbody>
                          {g.txns.sort((a,b)=>b.date.localeCompare(a.date)).map(t=>(
                            <tr key={t.id}>
                              <td className="mono" style={{fontSize:11,color:"var(--text3)"}}>{t.date}</td>
                              <td style={{fontSize:12}}>{t.description}</td>
                              <td className="r mono" style={{fontSize:11,color:"var(--red)"}}>{t.debit>0?fmt(t.debit):"—"}</td>
                              <td className="r mono" style={{fontSize:11,color:"var(--green)"}}>{t.credit>0?fmt(t.credit):"—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}

// ── PARTNERS & CAPITAL DASHBOARD ──────────────────────────────────────────────
// Built on top of the Bank Statement grouping — "partners" are whichever
// counterparties you've tagged as Owner / Capital there, "vendors" whichever
// you've tagged Vendor / Supplier. Capital figures are always all-time
// (contributions are a running total, not a per-period thing); vendor
// payments and the monthly trend respect the fiscal-year picker.
function PartnersCapitalDashboard({bankTxns,bankLabels}){
  const [fy,setFy]=useState(currentFY());

  const merged=useMemo(()=>{
    const raw={};
    bankTxns.forEach(t=>{
      if(!raw[t.key])raw[t.key]={key:t.key,txns:[],totalDebit:0,totalCredit:0};
      raw[t.key].txns.push(t); raw[t.key].totalDebit+=t.debit; raw[t.key].totalCredit+=t.credit;
    });
    const list=Object.values(raw).map(g=>({...g,label:bankLabels[g.key]?.label||"",type:bankLabels[g.key]?.type||"unlabeled"}));
    const m={};
    list.forEach(g=>{
      const mergeKey=g.label.trim()?g.label.trim().toLowerCase():`__raw__${g.key}`;
      if(!m[mergeKey])m[mergeKey]={mergeKey,label:g.label||g.key,type:g.type,rawKeys:[],txns:[],totalDebit:0,totalCredit:0};
      m[mergeKey].rawKeys.push(g.key);
      m[mergeKey].txns.push(...g.txns);
      m[mergeKey].totalDebit+=g.totalDebit;
      m[mergeKey].totalCredit+=g.totalCredit;
    });
    return Object.values(m);
  },[bankTxns,bankLabels]);

  const fyOptions=useMemo(()=>{
    const set=new Set(bankTxns.map(t=>getFY(t.date)).filter(Boolean));
    set.add(currentFY());
    return Array.from(set).sort().reverse();
  },[bankTxns]);

  const range=fy==="all"?null:fyRange(fy);
  const inRange=d=>!range||(d>=range.start&&d<=range.end);
  const periodLabel=fy==="all"?"All Time":`FY ${fy}`;

  // Capital — all-time, not period-filtered
  const partners=merged.filter(g=>g.type==="owner"&&g.totalCredit>0).sort((a,b)=>b.totalCredit-a.totalCredit);
  const totalCapital=partners.reduce((s,p)=>s+p.totalCredit,0);
  const maxPartner=partners[0]?.totalCredit||1;

  // Vendors — respects the FY picker
  const vendors=merged.filter(g=>g.type==="vendor")
    .map(g=>({...g,periodDebit:g.txns.filter(t=>inRange(t.date)).reduce((s,t)=>s+t.debit,0)}))
    .filter(g=>g.periodDebit>0).sort((a,b)=>b.periodDebit-a.periodDebit);
  const totalVendorPaid=vendors.reduce((s,v)=>s+v.periodDebit,0);
  const maxVendor=vendors[0]?.periodDebit||1;

  // Monthly trend — respects the FY picker
  const monthlyMap={};
  bankTxns.forEach(t=>{
    if(!inRange(t.date))return;
    const ym=t.date.slice(0,7);
    if(!monthlyMap[ym])monthlyMap[ym]={ym,credit:0,debit:0};
    monthlyMap[ym].credit+=t.credit; monthlyMap[ym].debit+=t.debit;
  });
  const monthly=Object.values(monthlyMap).sort((a,b)=>a.ym.localeCompare(b.ym));
  const maxMonthly=Math.max(1,...monthly.map(m=>Math.max(m.credit,m.debit)));

  if(bankTxns.length===0){
    return(
      <div>
        <div style={{marginBottom:20}}>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Partners & Capital</h3>
          <p style={{fontSize:12,color:"var(--text3)"}}>Built from your Bank Statement counterparties — tag some as "Owner / Capital" or "Vendor / Supplier" first.</p>
        </div>
        <EmptyState icon="🤝" message="No bank statement data yet" sub="Import a statement under Bank Statement first, then tag counterparties there."/>
      </div>
    );
  }

  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:14,marginBottom:20}}>
        <div>
          <h3 style={{fontSize:15,fontWeight:700,marginBottom:4}}>Partners & Capital</h3>
          <p style={{fontSize:12,color:"var(--text3)"}}>Capital figures are all-time. Vendor payments and the trend below follow the period picker.</p>
        </div>
        <FG label="Period">
          <select value={fy} onChange={e=>setFy(e.target.value)}>
            {fyOptions.map(f=><option key={f} value={f}>FY {f}</option>)}
            <option value="all">All Time</option>
          </select>
        </FG>
      </div>

      <div className="stats-grid">
        <StatCard label="Total Capital Invested" value={fmt(totalCapital)} sub="all-time, all partners" color="accent"/>
        <StatCard label="Partners"               value={partners.length}   sub="tagged Owner / Capital" color="purple"/>
        <StatCard label={`Paid to Vendors — ${periodLabel}`} value={fmt(totalVendorPaid)} sub={`${vendors.length} vendors`} color="red"/>
        <StatCard label="Net Cash Flow"          value={fmt(monthly.reduce((s,m)=>s+m.credit-m.debit,0))} sub={periodLabel} color="blue"/>
      </div>

      <div className="report-section" style={{marginBottom:20}}>
        <h3>Capital by Partner <span style={{fontWeight:400,color:"var(--text3)",fontSize:11}}>(all-time)</span></h3>
        {partners.length===0
          ?<EmptyState icon="🤝" message="No partners tagged yet" sub="Tag a counterparty as 'Owner / Capital' under Bank Statement."/>
          :partners.map(p=>(
            <div key={p.mergeKey} className="mini-bar-row">
              <span className="label">{p.label}</span>
              <div className="track"><div className="fill" style={{width:`${(p.totalCredit/maxPartner)*100}%`,background:"var(--accent)"}}/></div>
              <span className="val">{fmt(p.totalCredit)}</span>
            </div>
          ))
        }
      </div>

      <div className="report-section" style={{marginBottom:20}}>
        <h3>Vendor Payments <span style={{fontWeight:400,color:"var(--text3)",fontSize:11}}>({periodLabel})</span></h3>
        {vendors.length===0
          ?<EmptyState icon="📦" message="No vendor payments in this period" sub="Tag counterparties as 'Vendor / Supplier' under Bank Statement, or try a different period."/>
          :vendors.map(v=>(
            <div key={v.mergeKey} className="mini-bar-row">
              <span className="label">{v.label}</span>
              <div className="track"><div className="fill" style={{width:`${(v.periodDebit/maxVendor)*100}%`,background:"var(--red)"}}/></div>
              <span className="val">{fmt(v.periodDebit)}</span>
            </div>
          ))
        }
      </div>

      <div className="report-section">
        <h3>Monthly Cash In vs Out <span style={{fontWeight:400,color:"var(--text3)",fontSize:11}}>({periodLabel})</span></h3>
        {monthly.length===0
          ?<EmptyState icon="📊" message="No transactions in this period"/>
          :monthly.map(m=>(
            <div key={m.ym} style={{marginBottom:14,paddingBottom:14,borderBottom:"1px solid var(--border)"}}>
              <div style={{fontSize:11.5,color:"var(--text2)",fontFamily:"var(--mono)",marginBottom:6}}>{m.ym}</div>
              <div className="mini-bar-row">
                <span className="label">In</span>
                <div className="track"><div className="fill" style={{width:`${(m.credit/maxMonthly)*100}%`,background:"var(--green)"}}/></div>
                <span className="val">{fmt(m.credit)}</span>
              </div>
              <div className="mini-bar-row">
                <span className="label">Out</span>
                <div className="track"><div className="fill" style={{width:`${(m.debit/maxMonthly)*100}%`,background:"var(--red)"}}/></div>
                <span className="val">{fmt(m.debit)}</span>
              </div>
            </div>
          ))
        }
      </div>
    </div>
  );
}

// ── APP SHELL ────────────────────────────────────────────────────────────────
// dashboard gets lots/samples too for alerts
export default function App(){
  const [view,setView]=useState("dashboard");
  const [navOpen,setNavOpen]=useState(false);
  const goTo=id=>{setView(id);setNavOpen(false);};
  const [vehicles,          setVehicles]          =useState(seedVehicles);
  const [customers,         setCustomers]         =useState(seedCustomers);
  const [grades,            setGrades]            =useState(seedGrades);
  const [suppliers,         setSuppliers]         =useState(seedSuppliers);
  const [weighments,        setWeighments]        =useState([]);
  const [shiftLogs,         setShiftLogs]         =useState([]);
  const [boulderReceipts,   setBoulderReceipts]   =useState([]);
  const [productionEntries, setProductionEntries] =useState([]);
  const [purchases,         setPurchases]         =useState([]);
  const [expenses,          setExpenses]          =useState([]);
  const [capitalItems,      setCapitalItems]      =useState([]);
  const [bankTxns,          setBankTxns]          =useState([]);
  const [bankLabels,        setBankLabels]        =useState({});
  const [equipment,         setEquipment]         =useState(seedEquipment);
  const [maintTasks,        setMaintTasks]        =useState(seedMaintTasks);
  const [maintLogs,         setMaintLogs]         =useState([]);
  const [lots,              setLots]              =useState([]);
  const [costConfig,        setCostConfig]         =useState({bagCostPerTonne:"",freightPerTonne:"",loadingPerTonne:"",monthlyLabCost:"",monthlyRent:"",miscFixed:"",stdSellingRate:""});
  const [qcTests,           setQcTests]           =useState([]);
  const [samples,           setSamples]           =useState([]);

  return(
    <>
      <style>{FONTS}{CSS}</style>
      <div className="app">
        {navOpen&&<div className="sidebar-backdrop" onClick={()=>setNavOpen(false)}/>}
        <aside className={`sidebar${navOpen?" open":""}`}>
          <div className="sidebar-logo">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div><h1>BallMill ERP</h1><p>Nova Geo Resources</p></div>
              <button className="sidebar-close" onClick={()=>setNavOpen(false)} aria-label="Close menu">✕</button>
            </div>
          </div>
          <nav className="nav">
            {NAV.filter(n=>!n.group).map(n=>(
              <button key={n.id} className={`nav-item${view===n.id?" active":""}`} onClick={()=>goTo(n.id)}>
                <span className="icon">{n.icon}</span>{n.label}
              </button>
            ))}
            {GROUPS.map(grp=>(
              <div key={grp}>
                <div className="nav-label">{grp}</div>
                {NAV.filter(n=>n.group===grp).map(n=>(
                  <button key={n.id} className={`nav-item${view===n.id?" active":""}`} onClick={()=>goTo(n.id)}>
                    <span className="icon">{n.icon}</span>{n.label}
                  </button>
                ))}
              </div>
            ))}
          </nav>
          <div className="sidebar-foot">BallMill ERP v5 · Nova Geo</div>
        </aside>
        <div className="main">
          <div className="topbar">
            <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
              <button className="hamburger-btn" onClick={()=>setNavOpen(true)} aria-label="Open menu">☰</button>
              <h2>{TITLES[view]||view}</h2>
            </div>
            <div className="topbar-right">
              <span className="topbar-date" style={{fontSize:11,color:"var(--text3)",fontFamily:"var(--mono)"}}>{today()}</span>
              <Badge type="accent">v1.0</Badge>
            </div>
          </div>
          <div className="content">
            {view==="dashboard"  &&<Dashboard vehicles={vehicles} customers={customers} grades={grades} weighments={weighments} boulderReceipts={boulderReceipts} productionEntries={productionEntries} expenses={expenses} lots={lots} qcTests={qcTests} samples={samples} maintTasks={maintTasks} maintLogs={maintLogs}/>}
            {view==="weighment"  &&<WeighmentEntry vehicles={vehicles} customers={customers} grades={grades} weighments={weighments} setWeighments={setWeighments}/>}
            {view==="boulder"    &&<BoulderReceipt suppliers={suppliers} vehicles={vehicles} boulderReceipts={boulderReceipts} setBoulderReceipts={setBoulderReceipts}/>}
            {view==="production" &&<ProductionEntry grades={grades} productionEntries={productionEntries} setProductionEntries={setProductionEntries} lots={lots} setLots={setLots}/>}
            {view==="stock"      &&<StockLedger grades={grades} boulderReceipts={boulderReceipts} productionEntries={productionEntries} weighments={weighments}/>}
            {view==="shift"      &&<ShiftLog grades={grades} shiftLogs={shiftLogs} setShiftLogs={setShiftLogs}/>}
            {view==="lots"       &&<LotRegister lots={lots} setLots={setLots} qcTests={qcTests} grades={grades}/>}
            {view==="qc"         &&<QCTests lots={lots} qcTests={qcTests} setQcTests={setQcTests} setSamples={setSamples}/>}
            {view==="samples"    &&<SampleRegister lots={lots} samples={samples} setSamples={setSamples}/>}
            {view==="costconfig" &&<CostConfig costConfig={costConfig} setCostConfig={setCostConfig}/>}
            {view==="costs"      &&<MonthlyCostSheet purchases={purchases} expenses={expenses} productionEntries={productionEntries} weighments={weighments} grades={grades} lots={lots} costConfig={costConfig}/>}
            {view==="maint"     &&<MaintDashboard maintTasks={maintTasks} maintLogs={maintLogs} equipment={equipment}/>}
            {view==="tasks"     &&<TaskSchedule maintTasks={maintTasks} setMaintTasks={setMaintTasks} equipment={equipment} maintLogs={maintLogs} setMaintLogs={setMaintLogs}/>}
            {view==="maintlog"  &&<MaintLog maintLogs={maintLogs} setMaintLogs={setMaintLogs} equipment={equipment} maintTasks={maintTasks} setMaintTasks={setMaintTasks}/>}
            {view==="equipment" &&<EquipmentRegister equipment={equipment} setEquipment={setEquipment} maintenanceLogs={maintLogs}/>}
            {view==="purchases"  &&<PurchaseRegister suppliers={suppliers} purchases={purchases} setPurchases={setPurchases}/>}
            {view==="opscosts"   &&<MonthlyOpsCosts expenses={expenses} setExpenses={setExpenses}/>}
            {view==="expenses"   &&<ExpensesView expenses={expenses} setExpenses={setExpenses}/>}
            {view==="capital"    &&<CapitalRegister capitalItems={capitalItems} setCapitalItems={setCapitalItems}/>}
            {view==="bankstatement" &&<BankStatementGrouping bankTxns={bankTxns} setBankTxns={setBankTxns} bankLabels={bankLabels} setBankLabels={setBankLabels}/>}
            {view==="partners"   &&<PartnersCapitalDashboard bankTxns={bankTxns} bankLabels={bankLabels}/>}
            {view==="rpt-daily"  &&<DailyReport weighments={weighments} productionEntries={productionEntries} boulderReceipts={boulderReceipts} expenses={expenses} grades={grades}/>}
            {view==="rpt-stock"  &&<StockReport grades={grades} boulderReceipts={boulderReceipts} productionEntries={productionEntries} weighments={weighments}/>}
            {view==="vehicles"   &&<VehicleMaster vehicles={vehicles} setVehicles={setVehicles}/>}
            {view==="customers"  &&<CustomerMaster customers={customers} setCustomers={setCustomers}/>}
            {view==="suppliers"  &&<SupplierMaster suppliers={suppliers} setSuppliers={setSuppliers}/>}
            {view==="grades"     &&<GradesMaster grades={grades} setGrades={setGrades}/>}
          </div>
        </div>
      </div>
    </>
  );
}
