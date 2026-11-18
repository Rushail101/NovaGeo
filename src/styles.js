// Global styles for the finance app (moved out of App.jsx)
export const FONTS = `@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');`;
export const CSS = `
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

// Additions: toasts, callouts, user chip, loader, light theme, print, small-screen fixes
export const EXTRA_CSS = `
.toasts{position:fixed;right:18px;bottom:18px;display:flex;flex-direction:column;gap:8px;z-index:1000;max-width:min(420px,calc(100vw - 36px))}
.toast{padding:11px 14px;border-radius:var(--r2);font-size:13px;box-shadow:var(--shadow2);cursor:pointer;animation:toastIn .18s ease-out;border:1px solid var(--border2);background:var(--bg3);color:var(--text)}
.toast-ok{border-color:#1a4028;background:#0a2016;color:var(--green)}
.toast-err{border-color:#401010;background:#200808;color:#ff8a8a}
.toast-info{border-color:#143050;background:#081828;color:var(--blue)}
@keyframes toastIn{from{transform:translateY(8px);opacity:0}to{transform:none;opacity:1}}
.callout{border-radius:var(--r2);padding:12px 16px;margin-bottom:16px;font-size:12.5px;line-height:1.6;border:1px solid var(--border2);background:var(--bg2);color:var(--text2)}
.callout b{color:var(--text)}
.callout-amber{border-color:#3a2e00;background:#1a1500}
.callout-red{border-color:#401010;background:#200808}
.callout-green{border-color:#1a4028;background:#0a2016}
.callout-blue{border-color:#143050;background:#081828}
.inline-select{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:6px 9px;color:var(--text);font-family:var(--font);font-size:12.5px}
.search-input{background:var(--bg3);border:1px solid var(--border2);border-radius:var(--r);padding:7px 11px;color:var(--text);font-size:12.5px;min-width:260px;font-family:var(--font);outline:none}
.search-input:focus,.inline-select:focus{border-color:var(--accent)}
.table-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
.sql-pre{background:var(--bg);border:1px solid var(--border);border-radius:var(--r);padding:10px 12px;font-family:var(--mono);font-size:11px;color:var(--text2);max-height:260px;overflow:auto;margin-top:10px;white-space:pre-wrap}
.stat-card.clickable{cursor:pointer;transition:border-color .12s,transform .12s}
.stat-card.clickable:hover{border-color:var(--border3);transform:translateY(-1px)}
.user-chip{display:flex;align-items:center;gap:9px}
.user-avatar{width:28px;height:28px;border-radius:50%;background:var(--bg4);color:var(--accent);display:grid;place-items:center;font-weight:700;font-size:12px;flex-shrink:0}
.user-name{font-size:12px;color:var(--text);font-family:var(--font);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.user-role{font-size:9.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em}
.loader{display:flex;flex-direction:column;align-items:center;gap:14px;color:var(--text3)}
.loader span{width:34px;height:34px;border:3px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.chip{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:20px;font-size:11px;font-family:var(--mono);background:var(--bg3);border:1px solid var(--border2);color:var(--text2)}
.kv{display:grid;grid-template-columns:auto 1fr;gap:6px 16px;font-size:12.5px}
.kv dt{color:var(--text3);font-family:var(--mono);font-size:11px;text-transform:uppercase;letter-spacing:.05em}
.section-title{font-size:12px;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.08em;font-family:var(--mono);margin:22px 0 10px}
.grid-auto{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;margin-bottom:20px}
.due-row-overdue td{background:rgba(255,92,92,.05)}
.due-row-soon td{background:rgba(245,166,35,.05)}
tbody tr:focus-within td{background:rgba(255,255,255,.03)}
button:focus-visible,select:focus-visible,input:focus-visible,textarea:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.form-group textarea{resize:vertical;min-height:60px}
td.num,th.num{text-align:right;font-family:var(--mono);font-variant-numeric:tabular-nums}
@media (max-width:860px){.hide-sm{display:none}.search-input{min-width:0;width:100%}}
@media print{
  .sidebar,.topbar,.toasts,.btn,.filter-bar,.pill-tabs{display:none !important}
  body,.content{background:#fff !important;color:#000 !important}
  .table-wrap,.stat-card,.report-section,.config-card{border-color:#ccc !important;background:#fff !important}
  td,th{color:#000 !important;border-color:#ddd !important}
}
/* Light theme (toggle in sidebar) */
:root[data-theme="light"]{
  --bg:#f6f6f3;--bg2:#ffffff;--bg3:#f0f0ec;--bg4:#e6e6e0;
  --border:#e3e3dd;--border2:#d6d6cf;--border3:#c4c4bc;
  --text:#1b1b18;--text2:#55554f;--text3:#8a8a82;--text4:#b4b4ac;
  --accent:#4f7a00;--accent2:#3f6400;--accent3:#6a9e10;
  --red:#d03030;--amber:#b87300;--blue:#2563c9;--green:#1f8a46;--purple:#6d4fd6;--teal:#0f8f80;
  --shadow:0 2px 10px rgba(0,0,0,.06);--shadow2:0 6px 28px rgba(0,0,0,.12);
}
:root[data-theme="light"] .nav-item.active,:root[data-theme="light"] .btn-primary,:root[data-theme="light"] .auth-mark{color:#fff}
:root[data-theme="light"] .badge-green,:root[data-theme="light"] .toast-ok,:root[data-theme="light"] .callout-green,:root[data-theme="light"] .auth-ok{background:#e8f6ec;border-color:#bfe3cb}
:root[data-theme="light"] .badge-amber,:root[data-theme="light"] .callout-amber,:root[data-theme="light"] .due-badge{background:#fff5e0;border-color:#f0d9a8}
:root[data-theme="light"] .badge-red,:root[data-theme="light"] .toast-err,:root[data-theme="light"] .callout-red,:root[data-theme="light"] .auth-err{background:#fdeaea;border-color:#f3c2c2;color:var(--red)}
:root[data-theme="light"] .badge-blue,:root[data-theme="light"] .toast-info,:root[data-theme="light"] .callout-blue{background:#e9f0fc;border-color:#c6d7f5}
:root[data-theme="light"] .badge-teal{background:#e3f6f3;border-color:#b5e3dc}
:root[data-theme="light"] .badge-purple{background:#efeafd;border-color:#d5caf7}
:root[data-theme="light"] .badge-accent,:root[data-theme="light"] .lot-chip,:root[data-theme="light"] .cpt-banner{background:#eef6dc;border-color:#d3e6a8}
:root[data-theme="light"] .btn-danger{border-color:#f3c2c2}:root[data-theme="light"] .btn-danger:hover{background:#fdeaea}
:root[data-theme="light"] .btn-info{border-color:#c6d7f5}:root[data-theme="light"] .btn-info:hover{background:#e9f0fc}
:root[data-theme="light"] tr:hover td{background:rgba(0,0,0,.02)}
:root[data-theme="light"] .modal-overlay{background:rgba(20,20,15,.45)}
`;

export const MODULE_CSS = `
.modal-overlay{z-index:600 !important}
@media (max-width:600px){.stats-grid{grid-template-columns:1fr 1fr !important;gap:10px !important}.stat-val{font-size:20px !important}}
.user-chip{flex-wrap:wrap}.user-chip .btn{padding:4px 8px;font-size:11px;flex:1;justify-content:center}.user-chip>div:nth-child(2){flex-basis:calc(100% - 40px)}
.modal-xl{max-width:1080px}
.form-row.cols-4{grid-template-columns:repeat(4,1fr)}
.check{display:flex;align-items:center;gap:8px;font-size:12.5px;color:var(--text2);cursor:pointer}
.check input{width:16px;height:16px;accent-color:var(--accent)}
.chip-warn{color:var(--red);border-color:#401010;background:#200808}
:root[data-theme="light"] .chip-warn{background:#fdeaea;border-color:#f3c2c2}
.link-btn{background:none;border:none;color:var(--accent);cursor:pointer;font-size:11px;padding:0;font-family:var(--mono)}
.edit-table td{padding:5px 6px}
.edit-table input,.edit-table select{width:100%;background:var(--bg3);border:1px solid var(--border2);border-radius:6px;padding:6px 8px;color:var(--text);font-family:var(--font);font-size:12.5px;outline:none}
.edit-table input:focus,.edit-table select:focus{border-color:var(--accent)}
.invoice-totals{display:flex;flex-wrap:wrap;gap:8px 22px;justify-content:flex-end;margin:14px 0 4px;padding:12px 16px;background:var(--bg3);border-radius:var(--r2)}
.invoice-totals div{display:flex;flex-direction:column;align-items:flex-end;gap:2px}
.invoice-totals span{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;font-family:var(--mono)}
.invoice-totals b{font-family:var(--mono);font-size:14px}
.invoice-totals .grand b{font-size:19px;color:var(--accent)}
.cal-month{margin-bottom:18px}
.cal-month h4{font-size:12px;font-family:var(--mono);color:var(--text3);text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px}
.cal-item{display:grid;grid-template-columns:92px 1fr auto auto;gap:12px;align-items:center;padding:10px 14px;background:var(--bg2);border:1px solid var(--border);border-left:3px solid var(--border3);border-radius:var(--r);margin-bottom:6px}
.cal-item.overdue{border-left-color:var(--red)}
.cal-item.soon{border-left-color:var(--amber)}
.cal-item.done{border-left-color:var(--green);opacity:.72}
.cal-item .cal-date{font-family:var(--mono);font-size:12px}
.cal-item .cal-title{font-size:13px;font-weight:500}
.cal-item .cal-sub{font-size:11px;color:var(--text3);margin-top:2px}
.ret-table td.lbl{color:var(--text2)}
.ret-table tr.sub td{font-weight:700;border-top:1px solid var(--border2)}
.ret-table tr.grand td{font-weight:800;color:var(--accent);border-top:2px solid var(--border2)}
@media (max-width:860px){.form-row.cols-4{grid-template-columns:1fr 1fr}.cal-item{grid-template-columns:1fr auto}.cal-item .cal-date{grid-column:1/-1}}
`;
