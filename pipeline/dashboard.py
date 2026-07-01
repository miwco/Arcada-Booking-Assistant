"""Generate the self-contained interactive planning dashboard.

The bookings are placed by the scheduler, then the browser becomes a live
planning layer. Edits (move, auto-resolve, change room, remove teacher, approve,
drag-and-drop) are kept in localStorage with an undo history, and conflicts and
workload recompute client-side. The planner never blocks a placement — it warns
clearly when a move creates a double booking. The Excel files stay the source of
truth; exporting the finished plan to Excel is a later step.
"""
from __future__ import annotations

import json
import os

from .calendar_model import SPEC, build_events
from .dictionaries import OUTPUT_DIR, load_all
from .realized import build_realized
from .scheduler import SLOT_TIMES, schedule

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = OUTPUT_DIR

TEMPLATE = r"""<!DOCTYPE html>
<html lang="sv">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Booking Assistant — Film &amp; Media</title>
<style>
  :root{
    --bg:#0f1419; --panel:#1a212b; --panel2:#222c38; --line:#33414f;
    --txt:#e6edf3; --muted:#8b9aa8; --accent:#4da3ff;
    /* category colours — cool hues only (no red/yellow/orange/pink; those are conflicts) */
    --spec-f:#8b5cf6; --spec-l:#a855f7; --spec-m:#7c3aed; --spec-p:#c084fc;  /* film specializations: violet family, small differences */
    --spec-film:#6366f1;   /* Film (several inriktningar): indigo */
    --spec-o:#06b6d4;      /* Online media: cyan — clearly different from Film */
    --spec-alla:#2563eb;   /* Hela årskursen: vivid blue (vibrant, not gray) */
    --spec-other:#64748b;  /* another programme: slate */
    /* conflict / warning colours — alarming, reserved for conflicts only */
    --cf-group:#facc15; --cf-teacher:#ec4899; --cf-room:#f97316; --cf-studio:#ef4444; --ok:#34d399;
  }
  *{box-sizing:border-box}
  body{margin:0;font:14px/1.4 "Segoe UI",system-ui,sans-serif;background:var(--bg);color:var(--txt)}
  header{padding:12px 18px;background:var(--panel);border-bottom:1px solid var(--line);
    display:flex;gap:14px;align-items:center;flex-wrap:wrap;position:sticky;top:0;z-index:6}
  header h1{font-size:16px;margin:0;font-weight:600}
  .tabs{display:flex;gap:6px}
  .tab{padding:7px 14px;border-radius:7px;background:var(--panel2);cursor:pointer;border:1px solid var(--line)}
  .tab.active{background:var(--accent);color:#04121f;border-color:var(--accent);font-weight:600}
  .controls{display:flex;gap:8px;align-items:center;margin-left:auto;flex-wrap:wrap}
  select,input{background:var(--panel2);color:var(--txt);border:1px solid var(--line);
    border-radius:6px;padding:6px 9px;font:inherit}
  label.chk{display:inline-flex;align-items:center;gap:5px;color:var(--muted);font-size:12.5px;cursor:pointer}
  main{padding:16px 18px}
  .legend{display:flex;gap:14px;flex-wrap:wrap;margin:0 0 12px;color:var(--muted);font-size:12px;align-items:center}
  .legend span{display:inline-flex;align-items:center;gap:6px}
  .dot{width:12px;height:12px;border-radius:3px;display:inline-block}
  .sample{display:inline-block;width:20px;height:14px;border-radius:4px;background:var(--spec-alla)}
  #filters{position:sticky;top:var(--hdrH,84px);z-index:5;background:var(--bg);padding:6px 0}
  .filters{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap;align-items:center}
  .ctab,.sf{padding:5px 11px;border-radius:6px;background:var(--panel2);cursor:pointer;border:1px solid var(--line);font-size:13px}
  .ctab.active{background:var(--panel);border-color:var(--accent);color:var(--accent);font-weight:600}
  .sf.active{background:var(--accent);color:#04121f;border-color:var(--accent);font-weight:600}
  .sf.preset{font-weight:600}
  .sep{width:1px;height:22px;background:var(--line);margin:0 4px}
  table.cal{border-collapse:collapse;width:100%;table-layout:fixed}
  table.cal th,table.cal td{border:1px solid var(--line);vertical-align:top;padding:0}
  table.cal th{background:var(--panel);font-size:12px;padding:5px}
  /* keep the day-header row visible while scrolling weeks (planning calendar) */
  table.cal:not(.cmp) thead th{position:sticky;top:calc(var(--hdrH,84px) + var(--filtH,40px));z-index:4}
  .wkcell{white-space:nowrap;background:var(--panel);color:var(--muted);font-size:12px;font-weight:600;width:66px;padding:5px!important;text-align:center}
  .wkcell.prodweek{border-left:4px solid #a855f7}
  .prodlabel{color:#c084fc;font-weight:700;font-size:9px}
  td.holcell{background:repeating-linear-gradient(45deg,rgba(239,68,68,.12) 0 8px,transparent 8px 16px)}
  .holtag{color:var(--cf-studio);font-weight:700;font-size:10px}
  .sessionsel{display:inline-flex;align-items:center;gap:5px}
  .sessionsel select{padding:3px 6px;background:var(--panel);color:var(--txt);border:1px solid var(--line);border-radius:5px}
  .badge.hol{bottom:-7px;left:-5px;background:#fff;color:#b91c1c;border:1px solid #b91c1c}
  .slot{border-bottom:1px dashed var(--line);padding:4px;min-height:38px}
  .slot:last-child{border-bottom:none}
  .slot .sl{font-size:9px;color:var(--muted);letter-spacing:.04em}
  .slot.sugg{outline:2px solid var(--ok);outline-offset:-2px;background:rgba(52,211,153,.10)}
  .slot.dropok{outline:2px solid var(--ok);outline-offset:-2px;background:rgba(52,211,153,.18)}
  .slot.dropwarn{outline:2px solid var(--cf-room);outline-offset:-2px;background:rgba(251,146,60,.14)}
  .slot.dropbad{outline:2px dashed var(--cf-studio);outline-offset:-2px;background:rgba(239,68,68,.18)}
  .slot.flash{animation:fl .5s}
  @keyframes fl{0%,100%{background:transparent}50%{background:rgba(251,113,133,.5)}}
  .ev{position:relative;border-radius:5px;padding:4px 6px;margin:4px 1px;color:#06101a;font-size:12.5px;
    border:1.5px solid rgba(0,0,0,.30);cursor:grab}
  .ev .c{font-weight:700;line-height:1.25}
  .ev .t{opacity:.85;font-size:11px;line-height:1.2}
  /* AI-placed (system chose the day/time) = clearly DASHED; requested = solid */
  .ev{border:2px solid rgba(0,0,0,.35)}
  .ev.ai{border:2px dashed rgba(0,0,0,.72)}
  /* EVERY conflict = a red ring; the letter pills say which kinds (T/G/R/A211) */
  .ev.conflict{box-shadow:0 0 0 2.5px var(--cf-studio)}
  .ev.ok{box-shadow:0 0 0 2.5px var(--ok)}
  .cfbadges{position:absolute;top:-9px;left:-4px;display:flex;gap:2px}
  .cfp{background:var(--cf-studio);color:#fff;font-size:8.5px;font-weight:800;border-radius:6px;padding:0 4px;border:1px solid #fff;line-height:1.55}
  @keyframes pulse{0%,100%{box-shadow:0 0 0 2px var(--cf-studio),0 0 0 4px rgba(239,68,68,.25)}50%{box-shadow:0 0 0 2px var(--cf-studio),0 0 0 5px rgba(239,68,68,.6)}}
  .ev.ext{opacity:.4;filter:grayscale(.55);cursor:default}
  .ev.ctx{background-image:repeating-linear-gradient(45deg,rgba(0,0,0,.16) 0 6px,transparent 6px 12px);
    border-style:dashed;border-color:#94a3b8}
  .progtag{display:inline-block;background:#475569;color:#fff;border-radius:3px;padding:0 4px;font-size:9.5px;font-weight:800;vertical-align:middle}
  .ev.dragging{opacity:.45}
  .badge{position:absolute;font-size:8.5px;font-weight:800;border-radius:7px;padding:0 4px;line-height:1.5}
  .badge.ai{top:-7px;right:-5px;background:#0ea5e9;color:#04121f}
  .badge.cf{top:-7px;left:-5px;background:#0b0f14;color:#fff;border:1px solid #fff}
  .badge.okb{top:-7px;left:-5px;background:var(--ok);color:#04121f}
  .badge.cpu{bottom:-7px;right:-5px;background:#1f2937;color:#9ad}
  .badge.mv{bottom:-7px;left:-5px;background:#3730a3;color:#c7d2fe}
  .spec-f{background:var(--spec-f)}.spec-l{background:var(--spec-l)}.spec-m{background:var(--spec-m)}
  .spec-o{background:var(--spec-o)}.spec-p{background:var(--spec-p)}.spec-film{background:var(--spec-film);color:#e6edf3}
  .spec-alla{background:var(--spec-alla);color:#e6edf3}.spec-other{background:var(--spec-other);color:#e6edf3}
  .bar{height:16px;background:var(--accent);border-radius:3px}
  .tlink{color:var(--accent);cursor:pointer;text-decoration:underline dotted}
  .tlink:hover{text-decoration:underline}
  .ovctl{display:flex;gap:8px;align-items:center;margin-bottom:12px;flex-wrap:wrap}
  .kpi{display:inline-block;background:var(--panel2);border:1px solid var(--line);border-radius:9px;padding:8px 13px;margin:0 8px 8px 0;text-align:center}
  .kpi b{font-size:19px;display:block;color:var(--txt)} .kpi span{font-size:11px;color:var(--muted)}
  .chartrow{display:flex;align-items:center;gap:8px;margin:3px 0}
  .chartrow .nm{width:160px;flex:none;text-align:right;font-size:12.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .chartrow .barwrap{flex:1;background:var(--panel2);border-radius:4px;overflow:hidden;display:flex;height:18px;min-width:40px}
  .seg{height:100%}
  .chartrow .val{width:120px;flex:none;font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums}
  .seglegend{display:flex;gap:12px;flex-wrap:wrap;margin:6px 0 12px;font-size:11.5px;color:var(--muted)}
  .seglegend span{display:inline-flex;align-items:center;gap:5px}
  .wtable{border-collapse:collapse;width:100%;max-width:980px}
  .wtable th,.wtable td{border-bottom:1px solid var(--line);padding:8px 10px;text-align:left}
  .wtable th{color:var(--muted);font-weight:600}.wtable td.num{text-align:right;font-variant-numeric:tabular-nums}
  .conf{background:var(--panel);border:1px solid var(--line);border-left:3px solid var(--cf-teacher);
    border-radius:7px;padding:9px 12px;margin-bottom:9px}
  .conf h4{margin:0 0 5px;font-size:13px}.conf .who{color:var(--spec-o)}
  .muted{color:var(--muted)}
  .pill{display:inline-block;background:var(--panel2);border:1px solid var(--line);border-radius:20px;
    padding:1px 9px;font-size:11px;margin-left:6px;color:var(--muted)}
  .cmpwrap{display:flex;gap:14px;align-items:flex-start}
  .cmpcol{flex:1;min-width:0}.cmpcol h3{margin:2px 0 6px;font-size:14px;color:var(--accent)}
  table.cal.cmp{min-width:1180px}
  table.cal.cmp th{text-align:center}
  .cmpdiv{width:12px!important;background:var(--bg)!important;border-top:none!important;border-bottom:none!important;padding:0!important}
  #modal{position:fixed;inset:0;background:rgba(0,0,0,.55);display:none;z-index:20;align-items:center;justify-content:center}
  #modal.open{display:flex}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:12px;max-width:560px;width:92%;
    max-height:90vh;overflow:auto;padding:18px 20px}
  .card h3{margin:0 0 4px}.card .meta{color:var(--muted);font-size:13px;margin-bottom:10px}
  .card .row{margin:5px 0}
  .act{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}
  button{background:var(--panel2);color:var(--txt);border:1px solid var(--line);border-radius:7px;
    padding:7px 12px;cursor:pointer;font:inherit}
  button.primary{background:var(--accent);color:#04121f;border-color:var(--accent);font-weight:600}
  button.warn{background:var(--cf-teacher);color:#180a0d;border-color:var(--cf-teacher)}
  button.small{padding:3px 8px;font-size:12px}
  button:disabled{opacity:.4;cursor:default}
  .box{background:var(--panel2);border:1px solid var(--line);border-radius:8px;padding:10px;margin-top:10px}
  .x{float:right;cursor:pointer;color:var(--muted);font-size:20px;line-height:1}
  #msg{margin-top:8px;font-size:13px}
  #toast{position:fixed;left:50%;transform:translateX(-50%);bottom:18px;z-index:30;padding:11px 18px;
    border-radius:9px;font-size:13.5px;box-shadow:0 6px 26px rgba(0,0,0,.5);display:none;max-width:80%}
  #toast.show{display:block}
  #toast.bad{background:#5b1722;border:1px solid var(--cf-teacher);color:#ffe0e6}
  #toast.warn{background:#4a2f12;border:1px solid var(--cf-room);color:#ffe7cf}
  .imp{max-width:920px}
  .drop{border:2px dashed var(--line);border-radius:11px;padding:26px;text-align:center;color:var(--muted);cursor:pointer;background:var(--panel)}
  .drop.over{border-color:var(--accent);background:var(--panel2);color:var(--txt)}
  .fitem{display:flex;justify-content:space-between;align-items:center;padding:5px 9px;border-bottom:1px solid var(--line);font-size:13px}
  .ifind{border-left:3px solid var(--info);background:var(--panel2);border-radius:7px;padding:7px 10px;margin:5px 0}
  .ifind.warn{border-color:var(--cf-group)}.ifind.error{border-color:var(--cf-teacher)}
  .step{font-weight:700;color:var(--accent);margin:14px 0 6px}
  .aimsg{font-size:13px;color:var(--txt)}
  .airow{padding:4px 7px;margin:3px 0;border-radius:6px;background:var(--panel);border:1px solid var(--line)}
  .airow.best{border-color:var(--ok);background:rgba(52,211,153,.10)}
  .home,.manage{max-width:940px}
  .hero{background:linear-gradient(135deg,var(--panel),var(--panel2));border:1px solid var(--line);border-radius:12px;padding:18px 20px;margin-bottom:14px}
  .hero h2{margin:0 0 4px;font-size:22px}
  .steps{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin:14px 0}
  .stepcard{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px;cursor:pointer;transition:border-color .12s,transform .12s}
  .stepcard:hover{border-color:var(--accent);transform:translateY(-2px)}
  .stepcard .n{display:inline-flex;width:22px;height:22px;align-items:center;justify-content:center;border-radius:50%;background:var(--accent);color:#fff;font-size:12px;font-weight:700;margin-right:6px}
  .stepcard b{font-size:14px}.stepcard .d{color:var(--muted);font-size:12px;margin-top:5px}
  .statgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin-top:8px}
  .stat{background:var(--panel2);border:1px solid var(--line);border-radius:8px;padding:9px 11px}
  .stat b{font-size:19px;display:block;line-height:1.2}.stat span{color:var(--muted);font-size:12px}
  .mtabs{display:flex;gap:6px;margin-bottom:12px;flex-wrap:wrap}
  .mtab{padding:6px 13px;border:1px solid var(--line);border-radius:8px;cursor:pointer;background:var(--panel);font-size:13px}
  .mtab.active{background:var(--accent);color:#fff;border-color:var(--accent)}
  .gtable{width:100%;border-collapse:collapse;font-size:13px;margin-top:4px}
  .gtable th{color:var(--muted);font-weight:600;text-align:left;padding:4px 7px;border-bottom:1px solid var(--line)}
  .gtable td{padding:3px 7px;border-bottom:1px solid var(--line);vertical-align:middle}
  .gtable input{width:100%;background:var(--panel);border:1px solid var(--line);color:var(--txt);border-radius:5px;padding:5px 7px;box-sizing:border-box}
  .gtable input.sm{width:80px}
  .savebar{position:sticky;bottom:0;background:var(--bg);padding:10px 0;border-top:1px solid var(--line);margin-top:10px;display:flex;gap:8px;align-items:center}
  .mktable{width:100%;border-collapse:collapse;margin-top:6px}
  .mktable td{padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:middle;font-size:13px}
  .mktable .mkc{width:80px;text-align:center;white-space:nowrap}
  .mktable .mkn{font-weight:600;width:130px;white-space:nowrap}
  .bulkrow{display:flex;gap:8px;align-items:center;padding:6px 8px;margin:4px 0;background:var(--panel2);border:1px solid var(--line);border-radius:8px}
  .bulkrow>span:first-child{min-width:160px}
  .bulkrow button{margin-left:0}
  a.tmpl{display:inline-block;padding:6px 12px;border:1px solid var(--line);border-radius:8px;background:var(--panel);color:var(--txt);text-decoration:none;font-size:13px}
  a.tmpl:hover{border-color:var(--accent)}
  button:disabled{opacity:.4;cursor:not-allowed}
  #aiResolveBtn:not(:disabled){border-color:var(--accent)}
  #toast.ok{background:#0f3a2c;border:1px solid var(--ok);color:#d6fbee}
</style>
</head>
<body>
<header>
  <h1>📅 Arcada Booking Assistant</h1>
  <div class="tabs">
    <div class="tab active" data-view="home">🏠 Home</div>
    <div class="tab" data-view="import">📥 Import</div>
    <div class="tab" data-view="calendar">Planning calendar</div>
    <div class="tab" data-view="compare">Compare</div>
    <div class="tab" data-view="workload">Teacher overview</div>
    <div class="tab" data-view="realized">Realized bookings</div>
    <div class="tab" data-view="conflicts">Conflicts</div>
    <div class="tab" data-view="manage">⚙ Manage</div>
  </div>
  <div class="controls" id="controls">
    <button id="undoBtn" title="Undo (Ctrl+Z)">↶ Undo</button>
    <button id="resetBtn" title="Reset all changes to the original import">↺ Reset</button>
    <button id="aiResolveBtn" title="Use the AI to resolve conflicts (needs an API key — set one in Manage → Settings)">🤖 Resolve all (AI)</button>
    <button id="resolveBtn" title="Resolve conflicts with the built-in rules — no AI needed">⚙ Solve all without AI</button>
    <button id="exportBtn" class="primary" title="Export the batch chosen in the semester selector (autumn / spring / both) to Excel">⬇ Export Excel</button>
    <select id="semester"><option value="">All semesters</option>
      <option value="autumn 2026">Autumn 2026</option><option value="spring 2027">Spring 2027</option></select>
    <select id="teacher"><option value="">All teachers</option></select>
    <input id="search" placeholder="search course / room…" size="14">
    <button id="helpBtn" title="How to use this page" style="border-radius:50%;width:30px;padding:0;font-weight:700">ⓘ</button>
  </div>
</header>
<main>
  <div id="legend" class="legend"></div>
  <div id="filters"></div>
  <div id="view"></div>
</main>
<div id="modal"><div class="card" id="card"></div></div>
<div id="toast"></div>
<input type="file" id="impFile" accept=".xlsx" multiple style="display:none">
<input type="file" id="cfgFile" accept=".xlsx" style="display:none">
<script>
const RAW=__DATA__, WEEKDATES=__WEEKDATES__, SPEC=__SPEC__, TEAM=new Set(__TEAM__), SLOTTIMES=__SLOTS__, TARGETS=__TARGETS__;
const REALIZED=__REALIZED__, RYEARS=__RYEARS__; let realizedYear=RYEARS[0]||null;
const HOLIDAYS=__HOLIDAYS__, SESSION=__SESSION__; let PRODPERIODS=__PRODPERIODS__;
const WDAYS=["Mon","Tue","Wed","Thu","Fri"];
const WLABEL={Mon:"Må",Tue:"Ti",Wed:"On",Thu:"To",Fri:"Fr"};
const SPECF=[["F","Foto"],["L","Ljud"],["M","Manus"],["O","Online"],["P","Prod"]];
const ALLSPECS=["F","L","M","O","P"];
let view="home", cohort=null, specSel=new Set(ALLSPECS), showExt=false;
let cmpA=null, cmpB=null, curPopup=null, dragId=null;
let IMPUP=[], IMPVAL=null, IMPAPP={};   // import screen: uploaded files, validation findings, approvals
let AI={available:false};               // /ai/status — is the Claude API key configured

const $=s=>document.querySelector(s);
const esc=s=>(s==null?"":String(s)).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const first=n=>n.split(" ")[0];
const unitsOf=slot=>slot==="FULL"?["AM","PM"]:[slot];
const short=c=>c.length>40?c.slice(0,39)+"…":c;
const progShort=p=>(p||"").split(/[ /]/)[0]||"ext";
const SPEC_LETTERS=["F","L","M","O","P"];
// Media-YY (whole year) expands to all its specialization groups, so the general
// year code overlaps Media-YY-X. Two bookings share students iff atoms intersect.
function groupAtoms(s){const out=new Set();
  (s||"").split(/[;,]/).forEach(g=>{g=g.trim();if(!g)return;
    const m=g.match(/^Media-(\d{2})$/);
    if(m) SPEC_LETTERS.forEach(x=>out.add("Media-"+m[1]+"-"+x)); else out.add(g);});
  return [...out];}
const semFromWeek=w=>w>=30?"autumn 2026":"spring 2027";
const dateOf=(sem,wk,wd)=>(WEEKDATES[sem+"|"+wk]||{})[wd]||"";
const STUDIO="A211"; // the film studio — highest-priority room conflict
// A booking may need several rooms at once (multi-camera): "+", "&", "och" = all
// rooms simultaneously. ",", "/", "eller" = alternatives, so we take the first.
function roomList(raw){
  raw=(raw||"").trim(); if(!raw) return [];
  // "A309 (A211, A311)" = teacher is OK with any of these but needs one -> use the
  // primary room (before the parenthesis); the parenthetical alternatives are ignored.
  const p=raw.indexOf("("); const primary=(p>=0?raw.slice(0,p):raw).trim();
  const low=primary.toLowerCase();
  if(["online","online/teams","teams","inget rum","-",""].includes(low)) return [];
  // "+", "&", "och" = several rooms AT ONCE (e.g. A211+A210 studio + sound).
  let parts = /[+&]| och /i.test(primary) ? primary.split(/\s*[+&]\s*|\s+och\s+/i)
                                          : [primary.split(/\s*[,/]\s*|\s+eller\s+/i)[0]];
  return [...new Set(parts.map(p=>p.trim().toUpperCase())
    .filter(p=>p&&!["ONLINE","TEAMS","INGET RUM"].includes(p)))];
}
const isStudio=r=>r===STUDIO;
const KNOWN_ROOMS=[...new Set(RAW.flatMap(e=>roomList(e.room)))].sort();
const SEM_WEEKS={}; Object.keys(WEEKDATES).forEach(k=>{const i=k.lastIndexOf("|");const s=k.slice(0,i),w=+k.slice(i+1);
  (SEM_WEEKS[s]||(SEM_WEEKS[s]=[])).push(w);}); for(const s in SEM_WEEKS) SEM_WEEKS[s].sort((a,b)=>a-b);

/* ---- overrides + undo (localStorage) ---------------------------------- */
const OVKEY="ba_overrides_v1"; let OV=load(); let UNDO=[];
function load(){try{return JSON.parse(localStorage.getItem(OVKEY))||{};}catch(e){return {};}}
function save(){localStorage.setItem(OVKEY,JSON.stringify(OV));}
function ovGet(k){return OV[k]||(OV[k]={});}
function pushUndo(){UNDO.push(JSON.stringify(OV)); if(UNDO.length>200)UNDO.shift(); refreshUndo();}
function undo(){ if(!UNDO.length){toast("Nothing to undo","ok");return;} OV=JSON.parse(UNDO.pop());
  save(); render(); if(curPopup&&$("#modal").classList.contains("open")) openPopup(curPopup); refreshUndo(); toast("Undone.","ok");}
function refreshUndo(){const b=$("#undoBtn"); if(b) b.disabled=!UNDO.length;}
function toast(html,type){const t=$("#toast"); t.className=type||""; t.innerHTML=html; t.classList.add("show");
  clearTimeout(toast._t); toast._t=setTimeout(()=>t.classList.remove("show"), type==="bad"?5200:2600);}

/* ---- working model + occupancy + conflicts ----------------------------- */
let MODEL=[], OCC={};
function buildModel(){
  const moved=ovGet("moved"),removed=ovGet("removed"),approved=ovGet("approved"),rooms=ovGet("room"),gone=ovGet("removed_lectures");
  const out=[];
  for(const e of RAW){
    if(gone[e.id]) continue;                       // whole lecture removed from the plan
    const m=moved[e.id];
    const wk=m&&m.week?+m.week:e.week;
    const wd=m?m.weekday:e.placed_weekday;
    const slot=m?m.slot:e.slot;
    const sem=(m&&m.week)?semFromWeek(wk):e.semester;
    const room=(e.id in rooms)?rooms[e.id]:e.room;
    const rm=removed[e.id]||[]; const tlist=e.teachers.split("; ").filter(t=>t&&!rm.includes(t));
    // 'moved' only if the current placement differs from the original imported one
    const moved_flag=(wk!==e.week)||(wd!==e.placed_weekday)||(slot!==e.slot);
    out.push({...e,week:wk,semester:sem,wd,slot,room,tlist,teachers:tlist.join("; "),
      moved:moved_flag,placed_date:dateOf(sem,wk,wd),approvedFlag:!!approved[e.id]||e.pre_ok});
  }
  return out;
}
function buildOcc(model){
  const occ={}; const add=(k,c)=>{(occ[k]||(occ[k]=new Set())).add(c);};
  for(const e of model){ if(e.external||!e.placed_date) continue; const ck=e.course_code+"|"+e.cohort;
    for(const u of unitsOf(e.slot)){
      groupAtoms(e.groups).forEach(g=>add("g|"+g+"|"+e.placed_date+"|"+u,ck));
      e.tlist.forEach(t=>add("t|"+t+"|"+e.placed_date+"|"+u,ck));
      roomList(e.room).forEach(r=>add("r|"+r+"|"+e.placed_date+"|"+u,ck));
    }}
  return occ;
}
function computeConflicts(model,occ){
  for(const e of model){ e.kinds=[]; e.state="clean"; if(e.external||!e.placed_date) continue;
    const ck=e.course_code+"|"+e.cohort;
    const has=(pre,name)=>{for(const u of unitsOf(e.slot)){const s=occ[pre+"|"+name+"|"+e.placed_date+"|"+u];
      if(s&&[...s].some(k=>k!==ck)) return true;} return false;};
    if(groupAtoms(e.groups).some(g=>has("g",g))) e.kinds.push("group");
    if(e.tlist.some(t=>has("t",t))) e.kinds.push("teacher");
    const clashRooms=roomList(e.room).filter(r=>has("r",r));
    if(clashRooms.some(isStudio)) e.kinds.push("studio");        // A211 — hard
    if(clashRooms.some(r=>!isStudio(r))) e.kinds.push("room");   // other rooms — soft
    // severity: group + studio (A211) are hard; teacher is an approvable conflict;
    // a non-A211 room clash on its own is only a soft warning.
    // every clash (incl. room) is a visible conflict; A211 studio + group are the hard ones
    e.state = !e.kinds.length ? "clean" : e.approvedFlag ? "ok" : "conflict";
  }
}
function ringClass(e){ if(e.state==="ok") return "ok"; return e.state==="conflict"?"conflict":""; }
const KIND_LETTER={group:"G",teacher:"T",room:"R",studio:"A211"};   // A211 = studio conflict
function kindBadge(e){return ["group","studio","teacher","room"].filter(k=>e.kinds.includes(k))
  .map(k=>KIND_LETTER[k]).join("+");}
function kindBadges(e){ if(e.state!=="conflict"||!e.kinds||!e.kinds.length) return "";
  return `<span class="cfbadges">`+["group","studio","teacher","room"].filter(k=>e.kinds.includes(k))
    .map(k=>`<span class="cfp" title="${k} conflict">${KIND_LETTER[k]}</span>`).join("")+`</span>`;}

/* ---- slot evaluation + auto-move (cross-week, severity-aware) ----------
   level 0 = fully clean; +1 soft room clash; +2 teacher clash (approvable).
   level Infinity = HARD block (group, A211 studio, or same course twice/day) —
   the auto-mover never lands here, but you can still drop manually (with a warning). */
function slotEval(e,wk,wd,slot){
  const sem=semFromWeek(wk); const date=dateOf(sem,wk,wd);
  if(!date) return {level:Infinity,reasons:["no date for that week/day"]};
  const reasons=[]; const ck=e.course_code+"|"+e.cohort;
  let group=false,teacher=false,studio=false,room=false,sameDay=false;
  if(MODEL.some(o=>o.session_group===e.session_group&&o.id!==e.id&&o.week===wk&&o.wd===wd&&!o.external)){
    sameDay=true; reasons.push("same course already on "+wd+" wk"+wk);}
  const busy=(pre,name)=>unitsOf(slot).some(u=>{const s=OCC[pre+"|"+name+"|"+date+"|"+u];return s&&[...s].some(k=>k!==ck);});
  for(const g of groupAtoms(e.groups)) if(busy("g",g)){group=true;reasons.push("group "+g+" busy");}
  for(const t of e.tlist) if(busy("t",t)){teacher=true;reasons.push("teacher "+t+" busy");}
  for(const r of roomList(e.room)) if(busy("r",r)){ if(isStudio(r)){studio=true;reasons.push("🎬 A211 studio busy");} else {room=true;reasons.push("room "+r+" busy");}}
  const hard=group||studio||sameDay;
  const level=hard?Infinity:(teacher?2:0)+(room?1:0);
  return {level,reasons:[...new Set(reasons)],group,teacher,studio,room,sameDay,hard};
}
// preferred teaching days: Tue/Wed/Thu best, then Mon, then Fri (Thu better than Fri)
const WD_PREF={Tue:0,Wed:0,Thu:0,Mon:1,Fri:2};
function candidateSlots(e){
  const sem=semFromWeek(e.week); const weeks=(SEM_WEEKS[sem]||[e.week]);
  const slots=e.slot==="FULL"?["FULL"]:["AM","PM"]; const out=[];
  for(const wk of weeks) for(const wd of WDAYS) for(const s of slots) out.push({week:wk,wd,slot:s});
  out.sort((a,b)=> (Math.abs(a.week-e.week)-Math.abs(b.week-e.week)) ||      // keep near the requested week
    ((a.slot===e.slot?0:1)-(b.slot===e.slot?0:1)) ||                          // keep AM/PM
    ((WD_PREF[a.wd]||0)-(WD_PREF[b.wd]||0)) || (WDAYS.indexOf(a.wd)-WDAYS.indexOf(b.wd)));  // prefer Tue/Wed/Thu
  return out;
}
function autoMove(e){
  // candidates are pre-sorted by least disruption; pick the lowest-severity slot.
  let best=null; const tried=[];
  for(const c of candidateSlots(e)){
    if(c.week===e.week&&c.wd===e.wd&&c.slot===e.slot) continue;
    const v=slotEval(e,c.week,c.wd,c.slot);
    if(v.level===Infinity){tried.push(...v.reasons);continue;}
    if(!best||v.level<best.v.level){best={c,v}; if(v.level===0) break;}
  }
  if(best){const w=[]; if(best.c.week!==e.week)w.push("moved to week "+best.c.week);
    if(best.v.teacher)w.push("still a teacher clash (approvable)"); if(best.v.room)w.push("minor room clash remains");
    if(e.hard_time)w.push("changes a requested 'absolut' time");
    return {ok:true,week:best.c.week,wd:best.c.wd,slot:best.c.slot,warn:w.join("; "),level:best.v.level};}
  return {ok:false,reasons:[...new Set(tried)].slice(0,6)};
}
function freeSlots(e,n){ // fully clean slots only (level 0), same week first then nearest
  const out=[]; for(const c of candidateSlots(e)){ if(c.week===e.week&&c.wd===e.wd&&c.slot===e.slot) continue;
    if(slotEval(e,c.week,c.wd,c.slot).level===0){out.push(c); if(out.length>=(n||8)) break;} } return out;
}
function freeRooms(e){ // known rooms not occupied at e's slot
  const ck=e.course_code+"|"+e.cohort; const mine=roomList(e.room);
  return KNOWN_ROOMS.filter(r=>!mine.includes(r)&&unitsOf(e.slot).every(u=>{
    const s=OCC["r|"+r+"|"+e.placed_date+"|"+u]; return !(s&&[...s].some(k=>k!==ck));}));
}

/* ---- filtering --------------------------------------------------------- */
function passSpec(e){ return e.specs.length?e.specs.some(s=>specSel.has(s)):specSel.size===5; }
function filtered(model){
  const sem=$("#semester").value,tea=$("#teacher").value,q=$("#search").value.toLowerCase();
  return model.filter(e=>((view==="realized")||!sem||e.semester===sem)&&(!tea||e.tlist.includes(tea))&&
    (!q||(e.course+" "+e.room+" "+(e.course_code||"")).toLowerCase().includes(q))&&passSpec(e)&&(showExt||!e.external));
}
// Source for the analytics views: planning MODEL, or realized for the chosen year.
function analyticsData(){
  if(view==="realized"){
    const all=REALIZED[realizedYear]||[];
    const course=filtered(all.filter(e=>e.kind==="course"));
    const adminH={}; all.filter(e=>e.kind==="admin").forEach(e=>e.tlist.filter(t=>TEAM.has(t))
      .forEach(t=>(adminH[t]=adminH[t]||[]).push(e)));
    for(const t in adminH) adminH[t]=dedupHours(adminH[t]);
    return {evs:course, realized:true, adminH};
  }
  return {evs:filtered(MODEL).filter(e=>!e.external&&e.placed_date), realized:false, adminH:{}};
}
const cohorts=()=>{const o=["Media-26","Media-25","Media-24","Media-23"];
  return [...new Set(RAW.map(e=>e.cohort))].sort((a,b)=>
    (o.indexOf(a)<0?9:o.indexOf(a))-(o.indexOf(b)<0?9:o.indexOf(b))||a.localeCompare(b));};
const weekKey=w=>w>=30?w:w+100;
function stickyTops(){ // measure the sticky header + filters so the calendar day-row sits right below them
  const h=document.querySelector("header"), f=$("#filters"), root=document.documentElement;
  if(h) root.style.setProperty("--hdrH",h.offsetHeight+"px");
  root.style.setProperty("--filtH",((f&&f.offsetHeight)||0)+"px");
}
addEventListener("resize",stickyTops);

/* ---- render ------------------------------------------------------------ */
function render(){
  MODEL=buildModel(); OCC=buildOcc(MODEL); computeConflicts(MODEL,OCC);
  renderLegend(); renderFilters(); refreshUndo();
  const ctl=$("#controls"); if(ctl) ctl.style.display=(view==="home"||view==="manage"||view==="import")?"none":"flex";
  updateSolverButtons(); stickyTops();
  if(view==="home") renderHome();
  else if(view==="manage") renderManage();
  else if(view==="import") renderImport();
  else if(view==="calendar") renderCalendar();
  else if(view==="compare") renderCompare();
  else if(view==="workload"||view==="realized") renderWorkload();
  else renderConflicts();
}
function renderLegend(){
  if(view!=="calendar"&&view!=="compare"){$("#legend").innerHTML="";return;}
  $("#legend").innerHTML=`<span class="muted" style="font-weight:600">Categories:</span>`+
    Object.entries(SPEC).map(([k,v])=>`<span><i class="dot ${v[1]}"></i>${esc(v[0])}</span>`).join("")
    +`<span class="sep"></span><span class="muted" style="font-weight:600">Conflicts:</span>`
    +`<span><i class="sample" style="background:transparent;box-shadow:0 0 0 2.5px var(--cf-studio)"></i>Conflict (red)</span>`
    +`<span class="cfp" style="position:static">G</span>Group`
    +`<span class="cfp" style="position:static">T</span>Teacher`
    +`<span class="cfp" style="position:static">R</span>Room`
    +`<span class="cfp" style="position:static">A211</span>Studio`
    +`<span><i class="sample" style="background:transparent;box-shadow:0 0 0 2.5px var(--ok)"></i>OK / approved</span>`
    +`<span class="sep"></span><span><i class="sample" style="border:2px dashed rgba(0,0,0,.72);background:var(--panel2)"></i>🤖 AI-placed (dashed)</span>`
    +`<span><i class="sample" style="border:2px solid rgba(0,0,0,.5);background:var(--panel2)"></i>requested (solid)</span>`;
}
/* ---- sessions (academic year + programme) + production periods ---------- */
function yearOpts(){ const cur=(SESSION&&SESSION.year)||"2026-2027"; const b=parseInt(cur.slice(2,4))||26; let o="";
  for(let y=b-3;y<=b+1;y++){ const v="20"+String(y).padStart(2,"0")+"-20"+String(y+1).padStart(2,"0"); o+=`<option ${v===cur?"selected":""}>${v}</option>`; } return o; }
function sessionSelector(){
  const s=SESSION||{}; const list=s.list||[];
  const years=[...new Set(list.map(x=>x.year).concat(s.year?[s.year]:[]))].filter(Boolean).sort();
  const progs=[...new Set(list.map(x=>x.programme).concat(s.programme?[s.programme]:[]))].filter(Boolean);
  return `<span class="sessionsel"><b>Session:</b> `+
    `<select id="selProg" onchange="switchSession()">`+progs.map(p=>`<option ${p===s.programme?"selected":""}>${esc(p)}</option>`).join("")+`</select> `+
    `<select id="selYear" onchange="switchSession()">`+years.map(y=>`<option ${y===s.year?"selected":""}>${esc(y)}</option>`).join("")+`</select> `+
    `<button class="small" onclick="prodPeriodsModal()" title="Mark production periods">🎬 Periods</button></span><span class="sep"></span>`;
}
function switchSession(){ const p=$("#selProg").value,y=$("#selYear").value;
  if(p===SESSION.programme&&y===SESSION.year) return;
  if(!(SESSION.list||[]).some(s=>s.programme===p&&s.year===y)){ toast("No saved session for "+p+" "+y+" — import that batch first.","warn"); render(); return; }
  toast("Loading "+p+" "+y+"…","ok");
  fetch("/session/select",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({programme:p,year:y})})
    .then(()=>{try{localStorage.removeItem("ba_overrides_v1");}catch(e){} location.href="/dashboard.html?view="+view+"&t="+Date.now();});
}
function ppRow(p){p=p||{};return `<tr><td><input data-f="name" value="${esc(p.name||"")}" placeholder="Spring production"></td>`+
  `<td><input type="date" data-f="start" value="${esc(p.start||"")}"></td>`+
  `<td><input type="date" data-f="end" value="${esc(p.end||"")}"></td>`+
  `<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`;}
function prodPeriodsModal(){ curPopup=null;
  if(!SESSION||!SESSION.year){ toast("Import a batch first to create a session.","warn"); return; }
  $("#card").innerHTML=`<span class="x" onclick="closePopup()">×</span><h3>🎬 Production periods</h3>`+
    `<div class="row muted">Weeks when students are making their productions — highlighted in the calendar. Add as many as you like (e.g. a longer one in spring, a shorter one in autumn). Session: <b>${esc(SESSION.programme)} ${esc(SESSION.year)}</b>.</div>`+
    `<table class="gtable"><thead><tr><th>Name</th><th>Start</th><th>End</th><th></th></tr></thead><tbody id="tbPP">${(PRODPERIODS||[]).map(ppRow).join("")}</tbody></table>`+
    `<div class="act"><button onclick="$('#tbPP').insertAdjacentHTML('beforeend',ppRow({}))">+ Add period</button></div>`+
    `<div class="act"><button class="primary" onclick="savePeriods()">💾 Save &amp; show</button><button onclick="closePopup()">Cancel</button></div>`;
  $("#modal").classList.add("open");
}
async function savePeriods(){
  const periods=[...$("#tbPP").querySelectorAll("tr")].map(tr=>{const o={};tr.querySelectorAll("input").forEach(i=>o[i.dataset.f]=i.value);return o;})
    .filter(p=>p.name&&p.start&&p.end);
  toast("Saving production periods…","ok");
  try{ await fetch("/session/periods",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({periods})}); }catch(e){}
  location.href="/dashboard.html?view="+view+"&t="+Date.now();
}
function renderFilters(){
  if(view==="conflicts"||view==="import"||view==="home"||view==="manage"){$("#filters").innerHTML="";return;}
  if(view==="realized"){
    $("#filters").innerHTML=`<div class="filters"><b>Realized year:</b> <select id="ryear">`+
      RYEARS.map(y=>`<option ${y===realizedYear?"selected":""}>${esc(y)}</option>`).join("")+`</select>`+
      `<span class="sep"></span>`+
      `<div class="sf preset ${specSel.size===5?'active':''}" data-main="all">All</div>`+
      `<div class="sf preset ${specSel.size===4&&!specSel.has("O")?'active':''}" data-main="film">Film</div>`+
      `<div class="sf preset ${specSel.size===1&&specSel.has("O")?'active':''}" data-main="online">Online</div><span class="sep"></span>`+
      SPECF.map(([k,l])=>`<div class="sf ${specSel.has(k)?'active':''}" data-s="${k}">${l}</div>`).join("")+
      `<span class="muted" style="margin-left:8px">actual bookings from the booking system</span></div>`;
    bindFilters(); const ry=$("#ryear"); if(ry) ry.onchange=e=>{realizedYear=e.target.value;render();};
    return;
  }
  const pAll=specSel.size===5, pFilm=specSel.size===4&&!specSel.has("O")&&["F","L","M","P"].every(s=>specSel.has(s)),
        pOnline=specSel.size===1&&specSel.has("O");
  let html=`<div class="filters">`+sessionSelector();
  if(view==="calendar") html+=cohorts().map(c=>`<div class="ctab ${c===cohort?'active':''}" data-c="${esc(c)}">${esc(c)}</div>`).join("")+`<span class="sep"></span>`;
  html+=`<div class="sf preset ${pAll?'active':''}" data-main="all">All</div>`
    +`<div class="sf preset ${pFilm?'active':''}" data-main="film">Film</div>`
    +`<div class="sf preset ${pOnline?'active':''}" data-main="online">Online</div><span class="sep"></span>`
    +SPECF.map(([k,l])=>`<div class="sf ${specSel.has(k)?'active':''}" data-s="${k}">${l}</div>`).join("")
    +`<span class="sep"></span><label class="chk"><input type="checkbox" id="extchk" ${showExt?'checked':''}> external</label>`;
  $("#filters").innerHTML=html+`</div>`; bindFilters();
}
// dates / Finnish holidays / production periods
const dateISO=(wk,wd)=>dateOf(semFromWeek(wk),wk,wd);
function holidayOn(wk,wd){ const iso=dateISO(wk,wd); return iso?(HOLIDAYS[iso]||null):null; }
function prodPeriodFor(wk){ const mon=dateISO(wk,"Mon"),fri=dateISO(wk,"Fri"); if(!mon)return null;
  for(const p of (PRODPERIODS||[])){ if(p.start&&p.end&&!(fri<p.start||mon>p.end)) return p; } return null; }
// one AM (or PM) cell — each week is rendered as TWO table rows (an AM row and a PM
// row) so the AM band and PM band line up across every day of the week.
function unitCell(evs,cohortName,wk,d,u){
  const hol=holidayOn(wk,d);
  const inside=evs.filter(e=>e.wd===d&&e.week===wk&&unitsOf(e.slot).includes(u)).map(evChip).join("");
  return `<td class="${hol?"holcell":""}"><div class="slot" data-cohort="${esc(cohortName)}" data-week="${wk}" data-wd="${d}" data-unit="${u}"${hol?` title="Public holiday: ${esc(hol)}"`:""}>`+
    `<span class="sl">${u} ${SLOTTIMES[u]||""}${(hol&&u==="AM")?` · <span class="holtag">🎌 ${esc(hol)}</span>`:""}</span>`+inside+`</div></td>`;
}
function wkCell(wk,span){const pp=prodPeriodFor(wk);
  return `<td class="wkcell${pp?" prodweek":""}"${span?' rowspan="2"':''}${pp?` title="Production period: ${esc(pp.name)}"`:""}>`+
    `W${wk}<br><span class="muted">${semFromWeek(wk).split(" ")[0]}</span>${pp?`<br><span class="prodlabel">🎬 ${esc(pp.name.slice(0,12))}</span>`:""}</td>`;}
function buildGrid(evs,cohortName){
  const weeks=[...new Set(evs.map(e=>e.week))].sort((a,b)=>weekKey(a)-weekKey(b));
  if(!weeks.length) return `<p class="muted">No bookings match.</p>`;
  let html=`<table class="cal"><thead><tr><th class="wkcell">Wk</th>`+WDAYS.map(d=>`<th>${WLABEL[d]}</th>`).join("")+`</tr></thead><tbody>`;
  for(const wk of weeks){
    html+=`<tr>`+wkCell(wk,true)+WDAYS.map(d=>unitCell(evs,cohortName,wk,d,"AM")).join("")+`</tr>`;
    html+=`<tr>`+WDAYS.map(d=>unitCell(evs,cohortName,wk,d,"PM")).join("")+`</tr>`;
  }
  return html+`</tbody></table>`;
}
// Compare: ONE shared table so each week is a single physical row — both
// cohorts' days sit on the exact same line, guaranteeing visual alignment.
function buildCompareGrid(fa,ca,fb,cb){
  const weeks=[...new Set([...fa,...fb].map(e=>e.week))].sort((a,b)=>weekKey(a)-weekKey(b));
  if(!weeks.length) return `<p class="muted">No bookings match.</p>`;
  const days=WDAYS.map(d=>`<th>${WLABEL[d]}</th>`).join("");
  let html=`<div style="overflow-x:auto"><table class="cal cmp"><thead>`+
    `<tr><th class="wkcell" rowspan="2">Wk</th><th colspan="5" style="color:var(--accent)">${esc(ca)}</th>`+
    `<th class="cmpdiv" rowspan="2"></th><th colspan="5" style="color:var(--accent)">${esc(cb)}</th></tr>`+
    `<tr>${days}${days}</tr></thead><tbody>`;
  for(const wk of weeks){
    html+=`<tr>`+wkCell(wk,true)+WDAYS.map(d=>unitCell(fa,ca,wk,d,"AM")).join("")+
      `<td class="cmpdiv" rowspan="2"></td>`+WDAYS.map(d=>unitCell(fb,cb,wk,d,"AM")).join("")+`</tr>`+
      `<tr>`+WDAYS.map(d=>unitCell(fa,ca,wk,d,"PM")).join("")+WDAYS.map(d=>unitCell(fb,cb,wk,d,"PM")).join("")+`</tr>`;
  }
  return html+`</tbody></table></div>`;
}
function evChip(e){
  if(e.external) return `<div class="ev ext" title="External booking (context)"><span class="badge cpu">ext</span><div class="c">${esc(short(e.course))}</div><div class="t">${esc(e.teachers)}</div></div>`;
  const cls=ringClass(e)+(e.ai_placed?" ai":"")+(e.context?" ctx":"");
  const prog=e.context?`<span class="progtag" title="${esc(e.program)} — context only, not exported">${esc(progShort(e.program))}</span> `:"";
  return `<div class="ev ${e.spec_class} ${cls}" data-id="${e.id}" draggable="true" title="${e.context?esc(e.program)+' (context · not exported) · ':''}drag to move · click for details">`+
    kindBadges(e)+(e.state==="ok"?`<span class="badge okb">OK</span>`:"")+
    (HOLIDAYS[e.placed_date]?`<span class="badge hol" title="On a public holiday: ${esc(HOLIDAYS[e.placed_date])}">🎌</span>`:"")+
    (e.ai_placed?`<span class="badge ai">AI</span>`:"")+(e.moved?`<span class="badge mv">moved</span>`:"")+
    (e.needs_computer?`<span class="badge cpu">🖥</span>`:"")+
    `<div class="c">${prog}${esc(short(e.course))}</div><div class="t">${esc(e.tlist.map(first).join(", "))}${e.room?" · "+esc(e.room):""}</div></div>`;
}
function renderCalendar(){ const cohs=cohorts(); if(!cohort||!cohs.includes(cohort)) cohort=cohs[0];
  $("#view").innerHTML=buildGrid(filtered(MODEL).filter(e=>e.cohort===cohort),cohort);}
function renderCompare(){
  const cohs=cohorts(); if(!cmpA) cmpA=cohs[0]; if(!cmpB) cmpB=cohs[1]||cohs[0];
  const sel=id=>`<select id="${id}">`+cohs.map(c=>`<option ${c===(id==='cmpA'?cmpA:cmpB)?'selected':''}>${esc(c)}</option>`).join("")+`</select>`;
  const fa=filtered(MODEL).filter(e=>e.cohort===cmpA), fb=filtered(MODEL).filter(e=>e.cohort===cmpB);
  $("#view").innerHTML=`<div class="filters">Compare ${sel("cmpA")} vs ${sel("cmpB")} <span class="muted">— weeks aligned on one shared row; teacher/room clashes between cohorts show as conflicts in both.</span></div>`+
    buildCompareGrid(fa,cmpA,fb,cmpB);
  $("#cmpA").onchange=e=>{cmpA=e.target.value;render();}; $("#cmpB").onchange=e=>{cmpB=e.target.value;render();};
}

/* ---- popup + actions --------------------------------------------------- */
function involved(e){ const mine=roomList(e.room);
  return MODEL.filter(o=>o!==e&&!o.external&&o.placed_date===e.placed_date&&o.course_code!==e.course_code&&
    unitsOf(o.slot).some(u=>unitsOf(e.slot).includes(u))&&
    (groupAtoms(o.groups).some(g=>groupAtoms(e.groups).includes(g))||o.tlist.some(t=>e.tlist.includes(t))||
     roomList(o.room).some(r=>mine.includes(r))));}
function sharedTeachers(list){const c={};list.forEach(e=>e.tlist.forEach(t=>c[t]=(c[t]||0)+1));return Object.keys(c).filter(t=>c[t]>1);}
function sharedGroups(list){const c={};list.forEach(e=>groupAtoms(e.groups).forEach(g=>c[g]=(c[g]||0)+1));return Object.keys(c).filter(g=>c[g]>1);}
function openPopup(id){
  const e=MODEL.find(x=>x.id==id); if(!e||e.external) return; curPopup=id;
  const ps=involved(e); const all=[e,...ps];
  let h=`<span class="x" onclick="closePopup()">×</span><h3>${esc(e.course)} <span class="pill">${esc(e.course_code)}</span></h3>`+
    `<div class="meta">${esc(e.cohort)} · W${e.week} ${esc(e.placed_date)} ${e.wd} · ${e.slot} ${esc(SLOTTIMES[e.slot]||"")} · ${e.minutes} min · ${e.ai_placed?"🤖 AI":"requested"}${e.moved?" · ✎ moved":""}</div>`+
    `<div class="row"><b>Teachers:</b> ${teacherSpans(e.teachers)}${e.examiner?` · <b>Examiner:</b> ${teacherSpans(e.examiner)}`:""}</div>`+
    `<div class="row"><b>Groups:</b> ${esc(e.groups)}</div>`+
    `<div class="row"><b>Room:</b> ${esc(e.room)||"—"}${e.needs_computer?" · 🖥 computer room needed":""}</div>`+
    (e.time_range?`<div class="row"><b>Requested time:</b> ${esc(e.time_range)}${e.hard_time?' <span class="pill">absolut/hard</span>':''}</div>`:"")+
    (e.comments?`<div class="row muted">💬 ${esc(e.comments)}</div>`:"");

  if(e.state==="conflict"){
    const expl=[];
    if(e.kinds.includes("teacher")) expl.push("👤 <b>Teacher clash</b>: "+sharedTeachers(all).map(esc).join(", ")+" teach two courses here.");
    if(e.kinds.includes("group")) expl.push("👥 <b>Group clash</b>: "+sharedGroups(all).map(esc).join(", ")+" booked twice.");
    if(e.kinds.includes("studio")) expl.push("🎬 <b>FILM STUDIO (A211) clash</b> — two classes in A211 at once; needs your explicit approval.");
    if(e.kinds.includes("room")) expl.push("🚪 <b>Room clash</b> (minor): "+esc(roomList(e.room).filter(r=>!isStudio(r)).join(", "))+" used twice.");
    h+=`<div class="box"><b>⚠ Conflict</b><div class="row">${expl.join("<br>")}</div><div class="muted">Involved:</div>`+
      all.map(x=>`<div class="muted">• ${esc(x.course)} <span class="pill">${esc(x.cohort)}</span> ${teacherSpans(x.teachers)} · ${esc(x.room)||"—"}${x.examiner?` · exam ${teacherSpans(x.examiner)}`:""}</div>`).join("")+
      `<div class="act">`+all.map(x=>`<button class="primary" onclick="moveCourse(${x.id})">↪ Move ${esc(short(x.course))} (${esc(x.cohort)})</button>`).join("")+
      `<button onclick="autoResolve(${id})">🤖 Auto-resolve</button><button onclick="approve(${id})">✓ Approve / keep</button></div><div id="msg"></div></div>`;
  } else if(e.state==="ok"){
    h+=`<div class="box"><b>✓ Conflict kept (approved${e.pre_ok?" — request allows double-booking":""})</b>`+
      (!e.pre_ok?`<div class="act"><button onclick="unapprove(${id})">Un-approve</button></div>`:"")+`</div>`;
  }
  // AI assist (optional; proposes & explains, never applies on its own)
  let aibtns="";
  if(e.state==="conflict") aibtns+=`<button onclick="aiSuggest(${id})">🤖 Ask AI to rank fixes</button>`;
  if(e.comments) aibtns+=`<button onclick="aiInterpret(${id})">🤖 Read the comment</button>`;
  if(aibtns) h+=`<div class="box"><b>🤖 AI assist</b> <span class="muted">— suggests &amp; explains; you approve, the planner applies</span>`+
    `<div class="act">${aibtns}</div><div id="aimsg" class="muted" style="margin-top:6px">`+
    (AI.available?"":`AI is off — copy <code>.env.example</code> to <code>.env</code>, add your key, restart.`)+`</div></div>`;
  // manual move (week + day + slot) + free slots + remove
  const fs=freeSlots(e,8);
  const slotOpts=e.slot==="FULL"?["FULL"]:["AM","PM","FULL"];
  h+=`<div class="box"><b>Move</b>${e.slot==="FULL"?' <span class="muted">(full day — AM+PM move together)</span>':''}<div class="act"><select id="mvw">`+(SEM_WEEKS[semFromWeek(e.week)]||[e.week]).map(w=>`<option ${w===e.week?"selected":""}>${w}</option>`).join("")+`</select>`+
    `<select id="mvd">`+WDAYS.map(d=>`<option ${d===e.wd?"selected":""}>${d}</option>`).join("")+`</select>`+
    `<select id="mvs">`+slotOpts.map(s=>`<option ${s===e.slot?"selected":""}>${s}</option>`).join("")+`</select>`+
    `<button class="primary" onclick="manualMove(${id})">Move</button>`+(e.moved?`<button onclick="resetMove(${id})">reset placement</button>`:"")+
    `<button class="warn" onclick="removeLecture(${id})">🗑 Remove lecture</button></div>`+
    (fs.length?`<div class="muted" style="margin-top:6px">Free slots: `+fs.map(c=>`<button class="small" onclick="doMove(${id},${c.week},'${c.wd}','${c.slot}')">${c.week===e.week?"":"W"+c.week+" "}${c.wd} ${c.slot}</button>`).join(" ")+`</div>`:`<div class="muted" style="margin-top:6px">No clash-free slot in this semester — you can still place it manually (it will warn).</div>`)+`</div>`;
  // room
  const fr=freeRooms(e).slice(0,8);
  h+=`<div class="box"><b>Room</b> <span class="muted">current: ${esc(e.room)||"—"}</span><div class="act">`+
    `<input id="rmInput" size="10" value="${esc(e.room)}"><button onclick="changeRoom(${id})">Set</button>`+
    ((e.id in ovGet("room"))?`<button onclick="resetRoom(${id})">reset</button>`:"")+`</div>`+
    (fr.length?`<div class="muted" style="margin-top:6px">Free here: `+fr.map(r=>`<button class="small" onclick="setRoom(${id},'${esc(r)}')">${esc(r)}</button>`).join(" ")+`</div>`:"")+`</div>`;
  // remove teacher
  if(e.tlist.length){
    h+=`<div class="box"><b>Remove teacher</b> <span class="muted">(examiner kept by default)</span><div class="act">`+
      e.tlist.map(t=>`<button class="warn small" onclick="rmTeacher(${id},'${esc(t).replace(/'/g,"\\'")}')">remove ${esc(t)}${t===e.examiner?" ⭐exam":""}</button>`).join("")+`</div>`+
      ((ovGet("removed")[id]||[]).length?`<div class="muted" style="margin-top:6px">removed: ${esc((ovGet("removed")[id]||[]).join(", "))} <button class="small" onclick="resetTeachers(${id})">undo</button></div>`:"")+`</div>`;
  }
  $("#card").innerHTML=h; $("#modal").classList.add("open");
}
function closePopup(){$("#modal").classList.remove("open");curPopup=null;}
function msg(t){const m=$("#msg");if(m)m.innerHTML=t;}

function doMove(id,week,wd,slot){
  pushUndo(); ovGet("moved")[id]={week:+week,weekday:wd,slot}; save(); render();
  const e=MODEL.find(x=>x.id==id);
  if(e&&e.kinds&&e.kinds.includes("studio")) toast("🎬 FILM STUDIO (A211) double-booked on "+e.placed_date+" "+e.slot+"! Kept — approve only if intended.","bad");
  else if(e&&e.state==="conflict") toast("⚠ Double booking: "+e.kinds.join(" + ")+" clash on "+e.placed_date+" "+e.slot+". Kept anyway.","bad");
  else toast("Moved"+(e&&e.week!=RAW.find(x=>x.id==id).week?" to week "+e.week:"")+".","ok");
  if(curPopup) openPopup(curPopup);
}
function manualMove(id){ doMove(id,$("#mvw").value,$("#mvd").value,$("#mvs").value); }

/* ---- AI assist: propose-not-apply (deterministic solver stays in charge) - */
function aiBox(html){const m=$("#aimsg"); if(m) m.innerHTML=html;}
function aiErr(j){ return j.error==="no_key"
  ? `AI is off — copy <code>.env.example</code> to <code>.env</code>, add your key, restart.`
  : `<span style="color:var(--cf-room)">🤖 ${esc(j.message||j.error||"AI request failed")}</span>`; }
function aiCandidates(e){   // only the *legal* moves the deterministic solver allows
  const out=[];
  for(const c of candidateSlots(e)){
    if(c.week===e.week&&c.wd===e.wd&&c.slot===e.slot) continue;
    const v=slotEval(e,c.week,c.wd,c.slot);
    if(v.level===Infinity) continue;
    out.push({index:out.length, week:c.week, wd:c.wd, slot:c.slot,
      label:(c.week===e.week?"":"W"+c.week+" ")+c.wd+" "+c.slot,
      quality:v.level===0?"clean":(v.level===1?"minor room clash":"teacher clash (needs approval)"),
      reasons:v.reasons});
    if(out.length>=8) break;
  }
  return out;
}
async function aiSuggest(id){
  const e=MODEL.find(x=>x.id==id); if(!e) return;
  const cands=aiCandidates(e);
  if(!cands.length){ aiBox(`<span style="color:var(--cf-room)">No legal alternative slot this semester — place it manually.</span>`); return; }
  aiBox(`🤖 Asking Claude to rank ${cands.length} legal option(s)…`);
  const payload={booking:{course:e.course, code:e.course_code, cohort:e.cohort, teachers:e.teachers,
    examiner:e.examiner, groups:e.groups, week:e.week, day:e.wd, slot:e.slot, room:e.room,
    requested_time:e.time_range||"", hard_time:!!e.hard_time, comment:e.comments||""},
    conflicts:e.kinds||[], candidates:cands};
  let j; try{ j=await (await fetch("/ai/suggest",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)})).json(); }
  catch(err){ aiBox(`<span style="color:var(--cf-room)">AI unavailable (is the local server running?).</span>`); return; }
  if(!j.ok){ aiBox(aiErr(j)); return; }
  const rec=(j.recommend_index===undefined?-1:j.recommend_index);
  const rank=(j.ranking&&j.ranking.length)?j.ranking:cands.map(c=>({index:c.index,why:c.quality}));
  let html=`<div class="aimsg"><b>🤖 Claude</b> <span class="pill">${esc(j.confidence||"?")} confidence</span> `+
    `<span class="muted">source: ${esc(j.source||"")}${j.cached?" · cached":""}</span>`+
    `<div class="row">${esc(j.explanation||"")}</div>`+
    rank.map(r=>{const c=cands[r.index]; if(!c) return ""; const best=r.index===rec;
      return `<div class="airow${best?" best":""}">${best?"⭐ ":""}<b>${esc(c.label)}</b> `+
        `<span class="muted">${esc(c.quality)}</span> — ${esc(r.why||"")} `+
        `<button class="small primary" onclick="doMove(${id},${c.week},'${c.wd}','${c.slot}')">Apply</button></div>`;}).join("")+
    (j.caveat?`<div class="row" style="color:var(--cf-room)">⚠ ${esc(j.caveat)}</div>`:"")+
    `<div class="muted" style="margin-top:4px">AI ranks only legal moves; you approve, the planner applies.</div></div>`;
  aiBox(html);
}
async function aiInterpret(id){
  const e=MODEL.find(x=>x.id==id); if(!e||!e.comments) return;
  aiBox(`🤖 Reading the comment…`);
  let j; try{ j=await (await fetch("/ai/interpret",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({comment:e.comments,course:e.course})})).json(); }
  catch(err){ aiBox(`<span style="color:var(--cf-room)">AI unavailable (is the local server running?).</span>`); return; }
  if(!j.ok){ aiBox(aiErr(j)); return; }
  const rows=[["Summary",j.summary],["Time",j.time_range],["Preferred slot",j.preferred_slot],["Room",j.room],
    ["Frequency",j.frequency],["Double-booking ok",j.double_booking_ok],["Computer room",j.needs_computer]]
    .filter(r=>r[1]!=null&&r[1]!=="");
  aiBox(`<div class="aimsg"><b>🤖 Claude read the comment</b> <span class="pill">${esc(j.confidence||"?")} confidence</span> `+
    `<span class="muted">source: Other comments${j.cached?" · cached":""}</span>`+
    rows.map(r=>`<div class="row"><b>${esc(r[0])}:</b> ${esc(String(r[1]))}</div>`).join("")+
    (j.reason?`<div class="row muted">Why: ${esc(j.reason)}</div>`:"")+
    `<div class="muted">Informational only — nothing changed.</div></div>`);
}
function moveCourse(id){ const e=MODEL.find(x=>x.id==id); const r=autoMove(e);
  if(r.ok){ doMove(id,r.week,r.wd,r.slot); msg(`✓ Moved <b>${esc(short(e.course))}</b> to ${r.week!==e.week?"week "+r.week+" ":""}${r.wd} ${r.slot}.${r.warn?' ⚠ '+esc(r.warn):''}`);}
  else suggestAlternatives(e,r.reasons);}
function autoResolve(id){ const e=MODEL.find(x=>x.id==id); for(const x of [e,...involved(e)]){const r=autoMove(x);
  if(r.ok){ doMove(x.id,r.week,r.wd,r.slot); msg(`✓ Auto-resolved by moving <b>${esc(short(x.course))}</b> to ${r.week!==x.week?"week "+r.week+" ":""}${r.wd} ${r.slot}.${r.warn?' ⚠ '+esc(r.warn):''}`); return;}}
  suggestAlternatives(e,autoMove(e).reasons);}
function suggestAlternatives(e,reasons){
  const all=[e,...involved(e)]; const shared=sharedTeachers(all); const sug=[];
  for(const t of shared){const examC=all.find(x=>x.examiner===t); all.filter(x=>x.examiner!==t).forEach(o=>sug.push({t,course:o,keep:examC}));}
  let h=`<div class="box"><b>No clash-free slot found</b><div class="muted">${(reasons||[]).map(esc).join("; ")||"every slot is busy"}</div>`;
  if(sug.length) h+=`<div class="row">Or remove a non-examiner teacher:</div><div class="act">`+
    sug.slice(0,4).map(s=>`<button class="warn small" onclick="rmTeacher(${s.course.id},'${esc(s.t).replace(/'/g,"\\'")}')">remove ${esc(s.t)} from ${esc(short(s.course.course))}${s.keep?` (keep on ${esc(short(s.keep.course))})`:""}</button>`).join("")+`</div>`;
  h+=`<div class="act"><button onclick="approve(${e.id})">✓ Keep conflict</button></div></div>`; msg(h);
}

/* ---- Resolve all conflicts (greedy solver) + learned rules ------------- */
let RULES=loadRules();
function loadRules(){try{return JSON.parse(localStorage.getItem("ba_rules"))||{};}catch(e){return {};}}
function saveRules(){localStorage.setItem("ba_rules",JSON.stringify(RULES));}
const usesStudio=e=>roomList(e.room).includes("A211");
function remainingConflicts(){
  const evs=MODEL.filter(e=>!e.external&&!e.context&&e.placed_date&&e.state==="conflict");
  const seen=new Set(),out=[];
  for(const e of evs){const k=e.course_code+e.cohort+e.placed_date+e.slot; if(!seen.has(k)){seen.add(k);out.push(e);}}
  return out;
}
function updateSolverButtons(){const b=$("#aiResolveBtn"); if(!b)return;
  b.disabled=!AI.available;
  b.title=AI.available?"Use the AI to resolve conflicts (model + rules from Manage → Settings)"
    :"Add an AI API key in Manage → Settings to enable the AI solver";}
/* AI-powered resolve-all: ask the AI to choose among the LEGAL options for each
   conflict (move to its pick, or approve if your house rules allow it). */
async function aiResolveAll(){
  if(!AI.available){toast("Add an AI API key in Manage → Settings first.","warn");return;}
  const conf=remainingConflicts();
  if(!conf.length){toast("No conflicts to resolve.","ok");return;}
  const MAX=50, n=Math.min(conf.length,MAX);
  exportModal(`<h3>🤖 Resolve with AI</h3>`+
    `<div class="row">The AI will look at <b>${n}</b> conflict${n>1?"s":""} and, for each, pick the best legal slot (or keep it if your rules allow). It uses your chosen model and custom instructions from Settings.</div>`+
    `<div class="row muted">This makes up to ${n} AI requests and stops at your spending cap. You approve nothing is exported automatically.</div>`+
    `<div class="act"><button class="primary" onclick="closePopup();aiResolveRun()">Start</button><button onclick="closePopup()">Cancel</button></div>`);
}
async function aiResolveRun(){
  pushUndo();
  let solved=0,moved=0,approved=0,asked=0; const MAX=50;
  toast("🤖 AI is resolving conflicts…","ok");
  for(let guard=0;guard<MAX;guard++){
    MODEL=buildModel();OCC=buildOcc(MODEL);computeConflicts(MODEL,OCC);
    const conf=remainingConflicts(); if(!conf.length) break;
    const e=conf[0]; const cands=aiCandidates(e);
    if(!cands.length){ [e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true); approved++; solved++; save(); continue; }
    asked++;
    let j; try{
      j=await (await fetch("/ai/suggest",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({
        booking:{course:e.course,code:e.course_code,cohort:e.cohort,teachers:e.teachers,examiner:e.examiner,
          groups:e.groups,week:e.week,day:e.wd,slot:e.slot,room:e.room,requested_time:e.time_range||"",
          hard_time:!!e.hard_time,comment:e.comments||""},conflicts:e.kinds||[],candidates:cands})})).json();
    }catch(err){ toast("AI request failed — stopping.","warn"); break; }
    if(j.error==="cap_exceeded"){ toast("AI spending cap reached — stopping.","warn"); break; }
    if(j.error==="no_key"){ toast("No AI key — stopping.","warn"); break; }
    if(!j.ok){ // can't get a suggestion; approve to make progress
      [e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true); approved++; solved++; save(); continue; }
    if(j.action==="approve"||j.recommend_index<0||!cands[j.recommend_index]){
      [e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true); approved++;
    }else{ const c=cands[j.recommend_index]; ovGet("moved")[e.id]={week:c.week,weekday:c.wd,slot:c.slot}; moved++; }
    solved++; save();
  }
  MODEL=buildModel();OCC=buildOcc(MODEL);computeConflicts(MODEL,OCC); render();
  try{ AI=await (await fetch("/ai/status")).json(); }catch(e){} updateSolverButtons();
  const rem=remainingConflicts();
  exportModal(`<h3>🤖 AI resolve — done</h3>`+
    `<div class="row">Resolved <b>${solved}</b> conflict(s): ${moved} moved, ${approved} kept (allowed). ${asked} AI request(s).</div>`+
    (rem.length?`<div class="row muted">${rem.length} still need your decision — open the Conflicts tab.</div>`:`<div class="row" style="color:var(--ok)">All clear — ready to export.</div>`)+
    `<div class="act"><button class="primary" onclick="closePopup()">Done</button></div>`);
}
function resolveAll(){
  pushUndo();
  let moves=0,guard=0;
  while(guard++<1200){
    MODEL=buildModel();OCC=buildOcc(MODEL);computeConflicts(MODEL,OCC);
    const conf=MODEL.filter(e=>!e.external&&!e.context&&e.placed_date&&e.state==="conflict");
    if(!conf.length) break;
    let did=false;
    for(const e of conf){
      // try to move the least-disruptive lecture in this clash to a fully-clean slot.
      // keep studio (A211) courses & fixed-time courses put; prefer moving AI-placed ones.
      const cands=[e,...involved(e)].filter(x=>x.placed_date)
        .sort((a,b)=>(usesStudio(a)?1:0)-(usesStudio(b)?1:0)||(a.hard_time?1:0)-(b.hard_time?1:0)||(a.ai_placed?0:1)-(b.ai_placed?0:1));
      for(const x of cands){const r=autoMove(x); if(r.ok&&r.level===0){ovGet("moved")[x.id]={week:r.week,weekday:r.wd,slot:r.slot};moves++;did=true;break;}}
      if(did) break;
    }
    if(!did) break;
  }
  // learned rule: auto-approve teacher-only clashes if the user asked us to remember
  if(RULES.autoApproveTeacher){
    MODEL=buildModel();OCC=buildOcc(MODEL);computeConflicts(MODEL,OCC);
    MODEL.filter(e=>!e.external&&!e.context&&e.state==="conflict"&&e.kinds.length===1&&e.kinds[0]==="teacher")
      .forEach(e=>[e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true));
  }
  save();render();
  const rem=remainingConflicts();
  if(!rem.length){closePopup();toast("✓ Resolved everything by moving "+moves+" lecture(s). Ready to export.","ok");}
  else resolveAllModal(moves,rem);
}
function resolveAllModal(moves,rem){
  curPopup=null;
  const byKind={}; rem.forEach(e=>{const t=e.kinds.includes("group")?"group":e.kinds.includes("studio")?"studio":e.kinds.includes("teacher")?"teacher":"room";byKind[t]=(byKind[t]||0)+1;});
  const hard=rem.filter(e=>e.kinds.includes("group")||e.kinds.includes("studio")).length;
  let h=`<span class="x" onclick="closePopup()">×</span><h3>🪄 Auto-resolve</h3>`+
    (moves!=null?`<div class="row">Moved <b>${moves}</b> lecture(s) to clear conflicts.</div>`:"")+
    `<div class="row"><b>${rem.length}</b> conflict(s) need your decision: <span class="muted">`+Object.entries(byKind).map(([k,n])=>`${n} ${k}`).join(" · ")+`</span></div>`+
    `<div class="box"><div class="act"><button class="primary" onclick="approveAllRemaining()">✓ Approve all remaining → unblock export</button>`+
    `<label class="chk"><input type="checkbox" id="remRule" ${RULES.autoApproveTeacher?"checked":""}> remember: auto-approve teacher clashes</label></div>`+
    (hard?`<div class="row muted">⚠ ${hard} of these are <b>group</b> or <b>A211 studio</b> clashes — approving means a real student/studio double-booking. Prefer to move those (open & "Move course").</div>`:`<div class="row muted">These are teacher/room clashes (a teacher can sometimes be double-booked if you approve it).</div>`)+`</div>`+
    rem.slice(0,40).map(e=>`<div class="find ${e.kinds.includes("group")||e.kinds.includes("studio")?"warn":"info"}">`+
      `<b>${esc(short(e.course))}</b> <span class="pill">${esc(e.cohort)}</span> ${esc(e.placed_date)} ${e.slot} — ${esc(kindBadge(e))} clash`+
      `<div class="act"><button class="small" onclick="openPopup(${e.id})">resolve…</button><button class="small" onclick="approveOne(${e.id})">approve</button></div></div>`).join("")+
    `<div class="act"><button onclick="closePopup()">Close</button></div>`;
  $("#card").innerHTML=h;$("#modal").classList.add("open");
}
function approveAllRemaining(){pushUndo();
  if($("#remRule")&&$("#remRule").checked){RULES.autoApproveTeacher=true;saveRules();}
  remainingConflicts().forEach(e=>[e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true));
  save();render();closePopup();toast("Approved remaining conflicts — export unblocked.","ok");}
function approveOne(id){pushUndo();const e=MODEL.find(x=>x.id==id);[e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true);
  save();render();const rem=remainingConflicts();rem.length?resolveAllModal(null,rem):(closePopup(),toast("All clear — ready to export.","ok"));}
function resetMove(id){pushUndo();delete ovGet("moved")[id];save();render();openPopup(id);toast("Placement reset.","ok");}
function approve(id){pushUndo();const e=MODEL.find(x=>x.id==id);[e,...involved(e)].forEach(p=>ovGet("approved")[p.id]=true);save();render();openPopup(id);}
function unapprove(id){pushUndo();delete ovGet("approved")[id];save();render();openPopup(id);}
function rmTeacher(id,t){pushUndo();const a=ovGet("removed")[id]||(ovGet("removed")[id]=[]);if(!a.includes(t))a.push(t);save();render();openPopup(curPopup||id);toast("Removed "+t+".","ok");}
function resetTeachers(id){pushUndo();delete ovGet("removed")[id];save();render();openPopup(id);}
function setRoom(id,r){pushUndo();ovGet("room")[id]=r;save();render();
  const e=MODEL.find(x=>x.id==id);
  if(e&&e.kinds.includes("studio")) toast("🎬 A211 still double-booked here.","bad");
  else if(e&&e.kinds.includes("room")) toast("Room "+r+" still clashes (minor).","warn");
  else toast("Room set to "+(r||"—")+".","ok");
  openPopup(id);}
function changeRoom(id){setRoom(id,$("#rmInput").value.trim());}
function resetRoom(id){pushUndo();delete ovGet("room")[id];save();render();openPopup(id);}
function removeLecture(id){const e=MODEL.find(x=>x.id==id);pushUndo();ovGet("removed_lectures")[id]=true;save();closePopup();render();
  toast("Removed: "+esc(e?e.course:"lecture")+". Undo with Ctrl+Z.","bad");}

/* ---- help + reset ------------------------------------------------------ */
function helpModal(){curPopup=null;
  $("#card").innerHTML=`<span class="x" onclick="closePopup()">×</span><h3>ⓘ How to use the planner</h3>`+
    `<div class="row muted">The dashboard is a planning layer over the imported booking requests. Your edits are saved in this browser; <b>Export Excel</b> writes the final files, <b>Reset</b> reverts to the import.</div>`+
    `<div class="box"><b>On the calendar</b>`+
    `<div class="row muted">• <b>Drag</b> a lecture to another slot — green = fits, red = clash (still allowed, you'll get a warning).</div>`+
    `<div class="row muted">• <b>Click</b> a lecture for details and actions: move, auto-resolve a conflict, change room, remove a teacher, remove the whole lecture, or approve/keep a conflict.</div>`+
    `<div class="row muted">• A <b>full-day</b> lecture shows in both AM and PM and moves as one.</div></div>`+
    `<div class="box"><b>Markers</b>`+
    `<table class="mktable">`+
    [["<span class='sample' style='border:2px dashed rgba(0,0,0,.72);background:var(--panel2);width:26px'></span> <span class='badge ai' style='position:static'>AI</span>",
      "Dashed border / AI badge","The system chose the day &amp; time (the teacher didn't specify one)."],
     ["<span class='sample' style='border:2px solid rgba(0,0,0,.5);background:var(--panel2);width:26px'></span>",
      "Solid border","The teacher requested this exact day &amp; time."],
     ["<span class='badge mv' style='position:static'>moved</span>","Moved","You moved it from its imported slot (clears if you put it back)."],
     ["<span class='sample' style='box-shadow:0 0 0 2.5px var(--cf-studio);background:transparent;width:22px'></span>",
      "Red ring","A conflict — two courses in the same slot. The letters say which kind(s)."],
     ["<span class='cfp' style='position:static'>G</span>","G — Group","The same student group is double-booked (never allowed)."],
     ["<span class='cfp' style='position:static'>T</span>","T — Teacher","The same teacher is double-booked (can be approved if intended)."],
     ["<span class='cfp' style='position:static'>R</span>","R — Room","The same room is booked twice."],
     ["<span class='cfp' style='position:static'>A211</span>","A211 — Studio","Two courses in the film studio A211 at once (needs a decision)."],
     ["<span class='badge okb' style='position:static'>OK</span>","OK","A conflict you approved / chose to keep."],
     ["🖥","Computer","A computer room is needed."]]
      .map(m=>`<tr><td class="mkc">${m[0]}</td><td class="mkn">${m[1]}</td><td class="muted">${m[2]}</td></tr>`).join("")+
    `</table></div>`+
    `<div class="box"><b>Filters & tools</b><div class="row muted">Cohort tabs · All / Film / Online + specialization toggles · semester · teacher · search · "show external". Undo = Ctrl+Z. The other tabs cover side-by-side compare, teacher workload, realized (past-year) bookings, and a conflicts list.</div></div>`+
    `<div class="box"><b>Trust</b><div class="row muted">Export shows a <b>review summary</b> of every change first and saves a <b>decision log</b> next to the files. The AI only suggests (with reason + confidence + source); you approve and the planner applies. Hard rules stay enforced on write.</div></div>`+
    `<div class="act"><button class="primary" onclick="closePopup()">Got it</button><button onclick="rulesModal()">🧠 Learned rules</button></div>`;
  $("#modal").classList.add("open");}
const RULE_LABELS={autoApproveTeacher:"Auto-approve teacher-only clashes during 🪄 Resolve all"};
function rulesModal(){curPopup=null;
  const keys=Object.keys(RULES).filter(k=>RULES[k]);
  $("#card").innerHTML=`<span class="x" onclick="closePopup()">×</span><h3>🧠 Learned rules</h3>`+
    `<div class="row muted">Rules you taught the assistant (by ticking “remember”). They are applied transparently — remove any you no longer want. The hard safety rules (students never double-booked, A211 protected, group hierarchy, the export gate) are built in and not editable here.</div>`+
    (keys.length?keys.map(k=>`<div class="ifind info"><b>${esc(RULE_LABELS[k]||k)}</b>`+
        `<div class="act"><button class="small warn" onclick="removeRule('${esc(k)}')">Remove</button></div></div>`).join("")
      :`<div class="row muted">No learned rules yet. They appear here when you tick “remember: auto-approve teacher clashes” in 🪄 Resolve all.</div>`)+
    `<div class="act"><button onclick="closePopup()">Close</button></div>`;
  $("#modal").classList.add("open");}
function removeRule(k){delete RULES[k];saveRules();rulesModal();toast("Rule removed.","ok");}
function resetConfirm(){curPopup=null;
  $("#card").innerHTML=`<span class="x" onclick="closePopup()">×</span><h3>↺ Reset all changes?</h3>`+
    `<div class="row">Are you sure you want to reset all changes? This reverts the planning view to the original imported booking data (moves, removals, approvals and room changes are discarded).</div>`+
    `<div class="row muted">You can still undo this with Ctrl+Z afterwards.</div>`+
    `<div class="act"><button class="warn" onclick="resetAll()">Yes, reset everything</button><button onclick="closePopup()">Cancel</button></div>`;
  $("#modal").classList.add("open");}
function resetAll(){pushUndo();OV={};save();closePopup();render();toast("Reset to the original imported plan.","ok");}

/* ---- teacher analytics ------------------------------------------------- */
const H_PER_ECTS=TARGETS.hours_per_ects_coursework||20; // course-work hours per ECTS
const INCLASS=TARGETS.inclass_per_ects||{warn_low:8,normal_min:10,normal_max:12,warn_high:14};
const YEARLY=TARGETS.yearly_coursework||{target_low:800,target_high:1200};
const RBOOKED=TARGETS.realized_booked||{target_low:400,target_high:800};
function bookedFlag(h){ // realized: booked course hours vs the booked-hours target
  if(h<RBOOKED.target_low) return {level:"low",msg:"This teacher may be under the yearly booked-hours target."};
  if(h>RBOOKED.target_high) return {level:"high",msg:"This teacher may be over the yearly booked-hours target."};
  return null;}
function yearlyTarget(name){const t=(TARGETS.teachers||{})[name]||{};
  if(t.yearly_low!=null) return {low:t.yearly_low,high:t.yearly_high};
  const f=t.fte!=null?t.fte:1; return {low:YEARLY.target_low*f,high:YEARLY.target_high*f};}
// course-level in-class hours (each session once, regardless of #teachers)
function courseInclass(list){const seen={};let m=0;
  for(const e of list){const k=e.id; if(seen[k])continue; seen[k]=1;}
  return dedupHours([...new Map(list.map(e=>[e.id,e])).values()]);}
function inclassFlag(ects,hrs){ if(!ects) return null; const r=hrs/ects;
  if(r<INCLASS.warn_low) return {level:"low",msg:"This course may have too few in-class hours compared to its ECTS."};
  if(r>INCLASS.warn_high) return {level:"high",msg:"This course may have too many in-class hours compared to its ECTS."};
  return null;}
function yearlyFlag(name,coursework){const t=yearlyTarget(name);
  if(coursework<t.low) return {level:"low",msg:"This teacher may be under the yearly course-workload target."};
  if(coursework>t.high) return {level:"high",msg:"This teacher may be over the yearly course-workload target."};
  return null;}
const PALETTE=["#4da3ff","#22c55e","#a855f7","#f59e0b","#ef4444","#14b8a6","#eab308","#ec4899","#64748b","#84cc16"];
function teacherSpans(str){ if(!str) return "—";
  return [...new Set(str.split(/;|,/).map(s=>s.trim()).filter(Boolean))]
    .map(n=>`<span class="tlink" data-t="${esc(n)}">${esc(n)}</span>`).join(", ");}
// effective hours: overlapping sessions on a day count once (the longer)
function dedupHours(list){const byDate={};
  for(const e of list){const D=byDate[e.placed_date]||(byDate[e.placed_date]={full:false,am:0,pm:0,max:0});
    const us=unitsOf(e.slot); if(us.length>1)D.full=true;
    if(us.includes("AM"))D.am=Math.max(D.am,e.minutes); if(us.includes("PM"))D.pm=Math.max(D.pm,e.minutes); D.max=Math.max(D.max,e.minutes);}
  let m=0; for(const d in byDate){const D=byDate[d]; m+=D.full?D.max:(D.am+D.pm);} return m/60;}
function courseIndex(evs){ // course(code|cohort) -> {ects, teamTeachers:Set}
  const idx={}; for(const e of evs){const k=e.course_code+"|"+e.cohort;
    const o=idx[k]||(idx[k]={ects:+e.ects||0,team:new Set()});
    e.tlist.filter(t=>TEAM.has(t)).forEach(t=>o.team.add(t));} return idx;}
function teacherStats(name,evs){
  const teaching=evs.filter(e=>e.tlist.includes(name));
  const examCourses=new Set(evs.filter(e=>e.examiner===name).map(e=>e.course_code+"|"+e.cohort));
  const idx=courseIndex(evs);
  const byCourse={};
  for(const e of teaching){const k=e.course_code+"|"+e.cohort;
    const c=byCourse[k]||(byCourse[k]={code:e.course_code,name:e.course,cohort:e.cohort,spec:e.spec_label,
      ects:+e.ects||0,sessions:0,list:[],confl:0,appr:0});
    c.sessions++; c.list.push(e); if(e.state==="conflict")c.confl++; if(e.state==="ok")c.appr++;}
  const courses=Object.entries(byCourse).map(([k,c])=>({...c,hours:dedupHours(c.list),
    examiner:examCourses.has(k), share:(c.ects*H_PER_ECTS)/Math.max(1,(idx[k]?idx[k].team.size:1))}));
  // examiner-only courses (examiner but doesn't teach a session)
  for(const k of examCourses){ if(!byCourse[k]){const any=evs.find(e=>e.course_code+"|"+e.cohort===k);
    courses.push({code:any.course_code,name:any.course,cohort:any.cohort,spec:any.spec_label,
      ects:+any.ects||0,sessions:0,hours:0,list:[],confl:0,appr:0,examiner:true,share:0});}}
  courses.sort((a,b)=>b.hours-a.hours||b.ects-a.ects);
  return {name,courses,hours:dedupHours(teaching),sessions:teaching.length,
    nCourses:new Set(teaching.map(e=>e.course_code+"|"+e.cohort)).size,
    nExam:examCourses.size, conflicts:teaching.filter(e=>e.state==="conflict"),
    approved:teaching.filter(e=>e.state==="ok"),
    expected:courses.reduce((s,c)=>s+c.share,0)};
}
function openTeacher(name){
  closePopup(); curPopup=null;
  const {evs,realized,adminH}=analyticsData();
  const st=teacherStats(name,evs); const onTeam=TEAM.has(name); const admin=adminH[name]||0;
  const sem=realized?realizedYear+" (realized)":($("#semester").value||"both semesters");
  let h=`<span class="x" onclick="closePopup()">×</span><h3>👤 ${esc(name)}${onTeam?"":' <span class="pill">not on team CSV</span>'}</h3>`;
  h+=`<div class="meta">${esc(sem)} · counts follow the top filters</div><div style="margin:8px 0">`;
  if(realized){
    const bf=bookedFlag(st.hours);
    h+=`<span class="kpi"><b>${st.hours.toFixed(0)}</b><span>course hours (booked)</span></span>`+
      `<span class="kpi"><b>${admin.toFixed(0)}</b><span>admin hours</span></span>`+
      `<span class="kpi"><b>${st.sessions}</b><span>sessions</span></span>`+
      `<span class="kpi"><b>${st.nCourses}</b><span>courses</span></span></div>`+
      `<div class="row muted">Booked-hours target ${RBOOKED.target_low}–${RBOOKED.target_high} h (course hours only; admin separate)${bf?` · <span style="color:var(--cf-room)">${bf.level==="low"?"🔽":"🔼"} ${esc(bf.msg)}</span>`:" · within target"}. Booked hours only; overlaps count once.</div>`;
  } else {
    const tgt=yearlyTarget(name); const yf=yearlyFlag(name,st.expected);
    h+=`<span class="kpi"><b>${st.hours.toFixed(0)}</b><span>in-class hours</span></span>`+
      `<span class="kpi"><b>${st.expected.toFixed(0)}</b><span>course-work (ECTS)</span></span>`+
      `<span class="kpi"><b>${Math.max(0,st.expected-st.hours).toFixed(0)}</b><span>other course-work</span></span>`+
      `<span class="kpi"><b>${st.sessions}</b><span>sessions</span></span>`+
      `<span class="kpi"><b>${st.nCourses}</b><span>courses</span></span>`+
      `<span class="kpi"><b>${st.nExam}</b><span>examiner of</span></span>`+
      `<span class="kpi"><b>${st.conflicts.length}</b><span>conflicts</span></span></div>`+
      `<div class="row muted">Yearly course-work target ${tgt.low.toFixed(0)}–${tgt.high.toFixed(0)} h${yf?` · <span style="color:var(--cf-room)">${yf.level==="low"?"🔽":"🔼"} ${esc(yf.msg)}</span>`:" · within target"}. <span class="pill">in-class ≠ course-work</span></div>`;
  }
  const maxh=Math.max(1,...st.courses.map(c=>c.hours));
  h+=`<div class="box"><b>Courses</b><table class="wtable" style="margin-top:6px"><thead><tr>`+
    `<th>Course</th><th>Role</th><th class="num">ECTS</th><th class="num">Sess.</th><th class="num">In-class h</th><th style="width:120px"></th></tr></thead><tbody>`+
    st.courses.map(c=>{const fl=inclassFlag(c.ects,c.hours);
      return `<tr title="${fl?esc(fl.msg):''}"><td>${esc(c.name)} <span class="pill">${esc(c.cohort)}</span>${fl?(fl.level==="low"?" 🔽":" 🔼"):""}</td>`+
      `<td>${c.sessions?"teacher":""}${c.examiner?(c.sessions?" + examiner ⭐":"examiner ⭐"):""}${c.confl?' <span class="pill">⚠'+c.confl+'</span>':''}${c.appr?' <span class="pill">✓'+c.appr+'</span>':''}</td>`+
      `<td class="num">${c.ects||""}</td><td class="num">${c.sessions||""}</td><td class="num">${c.hours?c.hours.toFixed(1):""}</td>`+
      `<td><div class="bar" style="width:${(c.hours/maxh*100).toFixed(0)}%"></div></td></tr>`;}).join("")+`</tbody></table></div>`;
  const cf=[...st.conflicts,...st.approved];
  if(cf.length) h+=`<div class="box"><b>Conflicts involving ${esc(name)}</b>`+
    cf.map(e=>`<div class="muted">${e.state==="ok"?"✓ approved":"⚠ unresolved"} · ${esc(e.course)} <span class="pill">${esc(e.cohort)}</span> ${esc(e.placed_date)} ${e.slot} · ${esc(e.kinds.join("+"))} <button class="small" onclick="openPopup(${e.id})">open</button></div>`).join("")+`</div>`;
  $("#card").innerHTML=h; $("#modal").classList.add("open");
}

/* ---- teacher overview (flexible, multi-mode) -------------------------- */
let ovMode="bars", ovMetric="inclass", ovBreak="none", ovX="coursework", ovY="inclass";
const METRICS={inclass:"In-class hours (scheduled)",coursework:"Course-work target (ECTS)",
  other:"Other course-work (target − in-class)",sessions:"Sessions",courses:"Courses",
  exam:"Examiner of",conflicts:"Conflicts",approved:"Approved conflicts"};
const AXIS={inclass:"In-class hours",coursework:"Course-work target",sessions:"Sessions",
  courses:"Courses",exam:"Examiner of",conflicts:"Conflicts"};
const BREAKS={none:"(no breakdown)",spec:"by specialization",semester:"by semester",cohort:"by year/cohort"};
const HOURKEY=k=>k==="inclass"||k==="coursework"||k==="other";
function segKey(e,b){return b==="spec"?e.spec_label:b==="semester"?e.semester:b==="cohort"?e.cohort:"all";}

function teamStats(evs){
  const idx=courseIndex(evs);
  const T={}; const get=t=>T[t]||(T[t]={teach:[],exam:new Set()});
  for(const e of evs){ e.tlist.filter(t=>TEAM.has(t)).forEach(t=>get(t).teach.push(e));
    if(TEAM.has(e.examiner)) get(e.examiner).exam.add(e.course_code+"|"+e.cohort); }
  return Object.entries(T).map(([t,o])=>{
    const courses=new Set(o.teach.map(e=>e.course_code+"|"+e.cohort));
    let coursework=0; courses.forEach(k=>{const ix=idx[k]; if(ix)coursework+=ix.ects*H_PER_ECTS/Math.max(1,ix.team.size);});
    const inclass=dedupHours(o.teach);
    return {t,teach:o.teach,inclass,coursework,other:Math.max(0,coursework-inclass),
      sessions:o.teach.length,courses:courses.size,exam:o.exam.size,
      conflicts:o.teach.filter(e=>e.state==="conflict").length,approved:o.teach.filter(e=>e.state==="ok").length,
      yflag:yearlyFlag(t,coursework),target:yearlyTarget(t)};});
}
function courseStats(evs){
  const by={}; for(const e of evs){const k=e.course_code+"|"+e.cohort;
    const c=by[k]||(by[k]={code:e.course_code,name:e.course,cohort:e.cohort,spec:e.spec_label,ects:+e.ects||0,list:[],team:new Set()});
    c.list.push(e); e.tlist.filter(t=>TEAM.has(t)).forEach(t=>c.team.add(t));}
  return Object.values(by).map(c=>{const inclass=courseInclass(c.list);
    return {...c,inclass,coursework:c.ects*H_PER_ECTS,flag:inclassFlag(c.ects,inclass)};});
}
function warningsBox(stats,courses,realized){
  const tw=realized
    ? stats.filter(s=>bookedFlag(s.inclass)).map(s=>({n:s.t,f:bookedFlag(s.inclass),v:s.inclass,t:RBOOKED,lab:"booked"}))
    : stats.filter(s=>s.yflag).map(s=>({n:s.t,f:s.yflag,v:s.coursework,t:{low:s.target.low,high:s.target.high},lab:"course-work"}));
  const cw=courses.filter(c=>c.flag);
  if(!tw.length&&!cw.length) return `<div class="box"><b>⚖ Workload check</b> <span class="muted">No teachers or courses outside the configured targets. (Targets are planning guides, not absolute.)</span></div>`;
  let h=`<div class="box"><b>⚖ Possible workload issues</b> <span class="muted">(planning guides, not final judgments — see config/workload_targets.json)</span>`;
  if(tw.length) h+=`<div class="row" style="margin-top:6px"><b>Teachers</b></div>`+tw.map(x=>
    `<div class="muted">${x.f.level==="low"?"🔽":"🔼"} <span class="tlink" data-t="${esc(x.n)}">${esc(x.n)}</span> — ${esc(x.f.msg)} (${x.lab} ≈ ${x.v.toFixed(0)} h; target ${x.t.target_low!=null?x.t.target_low:x.t.low}–${x.t.target_high!=null?x.t.target_high:x.t.high} h)</div>`).join("");
  if(cw.length) h+=`<div class="row" style="margin-top:6px"><b>Courses</b></div>`+cw.map(c=>
    `<div class="muted">${c.flag.level==="low"?"🔽":"🔼"} ${esc(c.name)} <span class="pill">${esc(c.cohort)}</span> — ${esc(c.flag.msg)} (${c.inclass.toFixed(0)} in-class h for ${c.ects} ECTS)</div>`).join("");
  return h+`</div>`;
}

function renderWorkload(){ // "Teacher overview" + "Realized bookings"
  const {evs,realized,adminH}=analyticsData();
  const stats=teamStats(evs), courses=courseStats(evs);
  stats.forEach(s=>s.admin=adminH[s.t]||0);
  const head=realized?`<div class="box" style="border-left:3px solid var(--ok)"><b>📁 Realized bookings · ${esc(realizedYear)}</b> <span class="muted">— actual sessions teachers were really booked for (a follow-up view; kept simple for now).</span>`+
    `<div class="row muted" style="margin-top:6px"><b>How this will work:</b> go to <b>Arcada Timetabler</b> (<a href="https://timetable.arcada.fi/Timetable" target="_blank" rel="noopener" style="color:var(--accent)">timetable.arcada.fi/Timetable</a>), download the teachers' actual bookings, then import them here to compare planned vs. realized.</div>`+
    `<div class="act"><button onclick="realizedImport()">📥 Import realized bookings</button></div></div>`:"";
  // realized: booked hours only — hide ECTS-based metrics and always-zero ones
  const mEntries=realized?[["inclass","Course hours (booked)"],["sessions","Sessions"],["courses","Courses"]]:Object.entries(METRICS);
  const aEntries=realized?[["inclass","Course hours"],["sessions","Sessions"],["courses","Courses"]]:Object.entries(AXIS);
  if(realized){ const ok=new Set(["inclass","sessions","courses"]);
    if(!ok.has(ovMetric))ovMetric="inclass"; if(!ok.has(ovX))ovX="inclass"; if(!ok.has(ovY))ovY="sessions"; }
  const ctl=head+`<div class="ovctl"><label>View: <select id="ovmode">`+
    [["bars","Ranked bars"],["scatter","Scatter (teacher vs teacher)"],["weekly",realized?"Weekly booked trend":"Weekly in-class trend"],["courses",realized?"Course comparison":"Course comparison (in-class vs ECTS)"]]
      .map(([k,v])=>`<option value="${k}" ${k===ovMode?"selected":""}>${v}</option>`).join("")+`</select></label>`+
    (ovMode==="bars"?`<label>Metric: <select id="ovm">`+mEntries.map(([k,v])=>`<option value="${k}" ${k===ovMetric?"selected":""}>${v}</option>`).join("")+`</select></label>`+
      `<label>Breakdown: <select id="ovb">`+Object.entries(BREAKS).map(([k,v])=>`<option value="${k}" ${k===ovBreak?"selected":""}>${v}</option>`).join("")+`</select></label>`:"")+
    (ovMode==="scatter"?`<label>X: <select id="ovx">`+aEntries.map(([k,v])=>`<option value="${k}" ${k===ovX?"selected":""}>${v}</option>`).join("")+`</select></label>`+
      `<label>Y: <select id="ovy">`+aEntries.map(([k,v])=>`<option value="${k}" ${k===ovY?"selected":""}>${v}</option>`).join("")+`</select></label>`:"")+
    `<span class="muted">${realized?"team only · booked hours · overlaps count once (longer wins) · admin shown separately":"team only · in-class = scheduled teaching · course-work = ECTS×"+H_PER_ECTS+"h ÷ team teachers"}</span></div>`;

  let body="";
  if(ovMode==="bars") body=barsView(stats);
  else if(ovMode==="scatter") body=scatterView(stats);
  else if(ovMode==="weekly") body=weeklyView(evs);
  else body=coursesView(courses);
  // metrics table (always)
  const table=`<div class="box" style="margin-top:14px"><b>All metrics</b> <span class="muted">click a name for the full profile</span>`+
    (realized
     ? `<table class="wtable" style="margin-top:6px"><thead><tr><th>Teacher</th><th class="num">Course h (booked)</th>`+
       `<th class="num">Admin h</th><th class="num">Target ${RBOOKED.target_low}–${RBOOKED.target_high}</th>`+
       `<th class="num">Sessions</th><th class="num">Courses</th></tr></thead><tbody>`+
       stats.slice().sort((a,b)=>b.inclass-a.inclass).map(s=>{const f=bookedFlag(s.inclass);
         return `<tr><td><span class="tlink" data-t="${esc(s.t)}">${esc(s.t)}</span></td>`+
         `<td class="num">${s.inclass.toFixed(0)}</td><td class="num">${s.admin.toFixed(0)}</td>`+
         `<td class="num" style="color:${f?'var(--cf-room)':'var(--muted)'}">${f?(f.level==='low'?'🔽 low':'🔼 high'):'ok'}</td>`+
         `<td class="num">${s.sessions}</td><td class="num">${s.courses}</td></tr>`;}).join("")
     : `<table class="wtable" style="margin-top:6px"><thead><tr><th>Teacher</th><th class="num">In-class h</th>`+
       `<th class="num">Course-work</th><th class="num">Target</th><th class="num">Sessions</th><th class="num">Courses</th>`+
       `<th class="num">Examiner</th><th class="num">Confl.</th></tr></thead><tbody>`+
       stats.slice().sort((a,b)=>b.inclass-a.inclass).map(s=>`<tr><td><span class="tlink" data-t="${esc(s.t)}">${esc(s.t)}</span></td>`+
         `<td class="num">${s.inclass.toFixed(1)}</td><td class="num">${s.coursework.toFixed(0)}</td>`+
         `<td class="num" style="color:${s.yflag?'var(--cf-room)':'var(--muted)'}">${s.target.low.toFixed(0)}–${s.target.high.toFixed(0)}${s.yflag?(s.yflag.level==='low'?' 🔽':' 🔼'):''}</td>`+
         `<td class="num">${s.sessions}</td><td class="num">${s.courses}</td><td class="num">${s.exam||""}</td>`+
         `<td class="num">${s.conflicts||""}</td></tr>`).join(""))+`</tbody></table></div>`;
  $("#view").innerHTML=ctl+warningsBox(stats,courses,realized)+body+table;
  const on=(id,fn)=>{const el=$("#"+id); if(el) el.onchange=e=>{fn(e.target.value);render();};};
  on("ovmode",v=>ovMode=v); on("ovm",v=>ovMetric=v); on("ovb",v=>ovBreak=v); on("ovx",v=>ovX=v); on("ovy",v=>ovY=v);
}
function barsView(stats){
  const breakdownOK=(ovMetric==="inclass"||ovMetric==="sessions")&&ovBreak!=="none";
  const segs=new Set();
  stats.forEach(s=>{s.seg={}; if(breakdownOK){const g={}; s.teach.forEach(e=>(g[segKey(e,ovBreak)]||(g[segKey(e,ovBreak)]=[])).push(e));
    for(const k in g){s.seg[k]=ovMetric==="inclass"?dedupHours(g[k]):g[k].length; segs.add(k);}}});
  const segList=[...segs].sort(); const segColor={}; segList.forEach((k,i)=>segColor[k]=PALETTE[i%PALETTE.length]);
  const arr=stats.slice().sort((a,b)=>(b[ovMetric]||0)-(a[ovMetric]||0));
  const max=Math.max(1,...arr.map(s=>s[ovMetric]||0));
  const fmt=v=>HOURKEY(ovMetric)?(+v).toFixed(0):v;
  let h=breakdownOK?`<div class="seglegend">`+segList.map(k=>`<span><i class="dot" style="background:${segColor[k]}"></i>${esc(k)}</span>`).join("")+`</div>`:"";
  h+=arr.map(s=>{const total=s[ovMetric]||0;
    const bar=breakdownOK?segList.filter(k=>s.seg[k]).map(k=>`<div class="seg" style="width:${(s.seg[k]/max*100).toFixed(1)}%;background:${segColor[k]}" title="${esc(k)}: ${fmt(s.seg[k])}"></div>`).join("")
      :`<div class="seg" style="width:${(total/max*100).toFixed(1)}%;background:var(--accent)"></div>`;
    return `<div class="chartrow"><div class="nm"><span class="tlink" data-t="${esc(s.t)}">${esc(s.t)}</span></div><div class="barwrap">${bar}</div><div class="val">${fmt(total)}${HOURKEY(ovMetric)?" h":""}</div></div>`;}).join("");
  return h;
}
function scatterView(stats){
  const W=560,H=380,P=46; const xs=stats.map(s=>s[ovX]||0), ys=stats.map(s=>s[ovY]||0);
  const mx=Math.max(1,...xs), my=Math.max(1,...ys);
  const px=v=>P+(v/mx)*(W-2*P), py=v=>H-P-(v/my)*(H-2*P);
  const diag=(HOURKEY(ovX)&&HOURKEY(ovY))?`<line x1="${px(0)}" y1="${py(0)}" x2="${px(Math.min(mx,my))}" y2="${py(Math.min(mx,my))}" stroke="#33414f" stroke-dasharray="4 4"/>`:"";
  const pts=stats.map(s=>`<g class="tlink" data-t="${esc(s.t)}" style="cursor:pointer"><circle cx="${px(s[ovX]||0).toFixed(0)}" cy="${py(s[ovY]||0).toFixed(0)}" r="6" fill="var(--accent)" opacity="0.85"><title>${esc(s.t)}: ${AXIS[ovX]} ${(s[ovX]||0).toFixed(0)}, ${AXIS[ovY]} ${(s[ovY]||0).toFixed(0)}</title></circle>`+
    `<text x="${(px(s[ovX]||0)+9).toFixed(0)}" y="${(py(s[ovY]||0)+4).toFixed(0)}" fill="var(--muted)" font-size="10">${esc(first(s.t))}</text></g>`).join("");
  return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:640px;background:var(--panel);border:1px solid var(--line);border-radius:8px">`+
    `<line x1="${P}" y1="${H-P}" x2="${W-P}" y2="${H-P}" stroke="#5b6b7b"/><line x1="${P}" y1="${P}" x2="${P}" y2="${H-P}" stroke="#5b6b7b"/>`+diag+pts+
    `<text x="${W/2}" y="${H-12}" text-anchor="middle" fill="var(--muted)" font-size="12">${esc(AXIS[ovX])}</text>`+
    `<text x="14" y="${H/2}" transform="rotate(-90 14 ${H/2})" text-anchor="middle" fill="var(--muted)" font-size="12">${esc(AXIS[ovY])}</text></svg>`+
    (HOURKEY(ovX)&&HOURKEY(ovY)?`<div class="muted" style="font-size:12px">Dashed line = equal X and Y (e.g. in-class equal to course-work target).</div>`:"");
}
function weeklyView(evs){
  const sel=$("#teacher")?$("#teacher").value:"";
  const byWk={}; for(const e of evs){ if(sel&&!e.tlist.includes(sel))continue;
    (byWk[e.week]||(byWk[e.week]=[])).push(e); }
  const weeks=Object.keys(byWk).map(Number).sort((a,b)=>weekKey(a)-weekKey(b));
  const vals=weeks.map(w=>({w,h:dedupHours(byWk[w])}));
  const max=Math.max(1,...vals.map(v=>v.h));
  return `<div class="muted" style="margin-bottom:6px">In-class hours per week${sel?` for <b>${esc(sel)}</b>`:" (team total)"} — pick a teacher in the top filter to focus.</div>`+
    vals.map(v=>`<div class="chartrow"><div class="nm" style="width:90px">W${v.w} <span class="muted">${semFromWeek(v.w).split(" ")[0]}</span></div>`+
      `<div class="barwrap"><div class="seg" style="width:${(v.h/max*100).toFixed(1)}%;background:var(--accent)"></div></div><div class="val">${v.h.toFixed(0)} h</div></div>`).join("");
}
function coursesView(courses){
  const arr=courses.slice().sort((a,b)=>(b.ects-a.ects)||(b.inclass-a.inclass));
  const max=Math.max(1,...arr.map(c=>c.inclass));
  return `<div class="muted" style="margin-bottom:6px">Per course: scheduled in-class hours vs ECTS. Normal ≈ ${INCLASS.normal_min}–${INCLASS.normal_max} h/ECTS; flagged if &lt; ${INCLASS.warn_low} or &gt; ${INCLASS.warn_high} h/ECTS.</div>`+
    `<table class="wtable"><thead><tr><th>Course</th><th class="num">ECTS</th><th class="num">In-class h</th><th class="num">h/ECTS</th><th class="num">Course-work</th><th style="width:160px">In-class</th></tr></thead><tbody>`+
    arr.map(c=>{const r=c.ects?c.inclass/c.ects:0; const col=c.flag?(c.flag.level==="low"?"var(--cf-room)":"var(--cf-teacher)"):"var(--ok)";
      return `<tr title="${c.flag?esc(c.flag.msg):''}"><td>${esc(c.name)} <span class="pill">${esc(c.cohort)}</span>${c.flag?(c.flag.level==="low"?" 🔽":" 🔼"):""}</td>`+
      `<td class="num">${c.ects||""}</td><td class="num">${c.inclass.toFixed(0)}</td><td class="num" style="color:${col}">${c.ects?r.toFixed(1):""}</td>`+
      `<td class="num">${c.coursework.toFixed(0)}</td><td><div class="bar" style="width:${(c.inclass/max*100).toFixed(0)}%;background:${col}"></div></td></tr>`;}).join("")+`</tbody></table>`;
}
function renderConflicts(){
  const evs=filtered(MODEL).filter(e=>!e.external&&e.placed_date&&e.state==="conflict"); const cells={};
  for(const e of evs) for(const u of unitsOf(e.slot)) (cells[e.placed_date+"|"+u]||(cells[e.placed_date+"|"+u]=[])).push(e);
  const cards=[];
  for(const key of Object.keys(cells).sort()){const [date,unit]=key.split("|");const list=cells[key];
    if(new Set(list.map(e=>e.course_code+e.cohort)).size<2) continue; const who=new Set(),seen={};
    for(const e of list)for(const kind of ["groups","tlist"]){const arr=kind==="groups"?groupAtoms(e.groups):e.tlist;
      for(const r of arr){const k=kind+r;(seen[k]||(seen[k]=new Set())).add(e.course_code);if(seen[k].size>1)who.add((kind==="groups"?"👥 ":"👤 ")+r);}}
    if(who.size) cards.push({date,unit,who:[...who],items:[...new Map(list.map(e=>[e.course_code+e.cohort,e])).values()]});}
  $("#view").innerHTML=`<p class="muted">Unresolved slot collisions (approved hidden). ${cards.length} clash(es).</p>`+
    (cards.length?cards.map(c=>`<div class="conf"><h4>${c.date} · ${c.unit} · <span class="who">${esc(c.who.join(", "))}</span></h4>`+
      c.items.map(e=>`<div class="muted">• ${esc(e.course)} <span class="pill">${esc(e.cohort)}</span> ${teacherSpans(e.teachers)} <button class="small" onclick="openPopup(${e.id})">resolve…</button></div>`).join("")+`</div>`).join("")
      :`<p class="muted">No unresolved conflicts 🎉</p>`);
}

/* ---- drag & drop (cross-week, never blocks) --------------------------- */
function clearDrag(){document.querySelectorAll(".slot.sugg,.slot.dropok,.slot.dropbad").forEach(s=>s.classList.remove("sugg","dropok","dropbad"));
  document.querySelectorAll(".ev.dragging").forEach(x=>x.classList.remove("dragging"));}
document.addEventListener("dragstart",ev=>{const c=ev.target.closest(".ev[data-id]");if(!c)return;
  dragId=+c.dataset.id;c.classList.add("dragging");const e=MODEL.find(x=>x.id===dragId);if(!e)return;
  document.querySelectorAll(".slot[data-week]").forEach(z=>{ if(z.dataset.cohort!==e.cohort) return;
    if(+z.dataset.week!==e.week) return; // green suggestions: same week, fully clean
    const slot=e.slot==="FULL"?"FULL":z.dataset.unit;
    if(slotEval(e,+z.dataset.week,z.dataset.wd,slot).level===0) z.classList.add("sugg");});});
document.addEventListener("dragend",clearDrag);
document.addEventListener("dragover",ev=>{const z=ev.target.closest(".slot[data-week]");if(!z||dragId==null)return;ev.preventDefault();
  const e=MODEL.find(x=>x.id===dragId);if(!e||z.dataset.cohort!==e.cohort)return;
  const slot=e.slot==="FULL"?"FULL":z.dataset.unit; const lvl=slotEval(e,+z.dataset.week,z.dataset.wd,slot).level;
  z.classList.remove("dropok","dropwarn","dropbad");
  z.classList.add(lvl===0?"dropok":lvl===Infinity?"dropbad":"dropwarn");});
document.addEventListener("dragleave",ev=>{const z=ev.target.closest(".slot[data-week]");if(z){z.classList.remove("dropok","dropwarn","dropbad");}});
document.addEventListener("drop",ev=>{const z=ev.target.closest(".slot[data-week]");if(!z||dragId==null)return;ev.preventDefault();
  const e=MODEL.find(x=>x.id===dragId); const id=dragId; dragId=null;
  if(!e||z.dataset.cohort!==e.cohort){clearDrag();toast("Can only move within the same cohort.","bad");return;}
  const slot=e.slot==="FULL"?"FULL":z.dataset.unit; const lvl=slotEval(e,+z.dataset.week,z.dataset.wd,slot).level;
  if(lvl===Infinity){z.classList.add("flash");setTimeout(()=>z.classList.remove("flash"),520);}
  clearDrag(); doMove(id,+z.dataset.week,z.dataset.wd,slot);});

/* ---- Excel export ------------------------------------------------------ */
function conflictNote(e){const ps=involved(e); const all=[e,...ps]; const bits=[];
  if(e.kinds.includes("studio")) bits.push("A211 studio also used by "+ps.map(p=>p.course).join(" / "));
  if(e.kinds.includes("group")) bits.push("group "+sharedGroups(all).join(", ")+" also in "+ps.map(p=>p.course).join(" / "));
  if(e.kinds.includes("teacher")) bits.push("teacher "+sharedTeachers(all).join(", ")+" also teaches "+ps.map(p=>p.course).join(" / "));
  if(e.pre_ok) bits.push("request allows double-booking");
  return bits.join("; ");}
function buildPlanPayload(){
  return MODEL.filter(e=>!e.external&&!e.context&&e.placed_date).map(e=>({   // context (other-programme) not exported
    id:e.id,cohort:e.cohort,semester:e.semester,week:e.week,weekday:e.wd,slot:e.slot,placed_date:e.placed_date,
    course_code:e.course_code,course:e.course,groups:e.groups,examiner:e.examiner,teachers:e.teachers,
    room:e.room,minutes:e.minutes,type:e.type,content:e.content,comments:e.comments,
    num_students:e.num_students,language:e.language,programme:e.programme,
    state:e.state,approvedFlag:e.approvedFlag,kinds:e.kinds,
    conflictNote:(e.approvedFlag&&e.kinds.length)?conflictNote(e):"",external:false}));
}
let pendingExportSem=null;                   // batch chosen for the export under review
function exportExcel(){
  const sem=$("#semester").value;            // export the batch chosen in the semester filter ("" = both)
  const batch=sem||"both semesters";
  const unresolved=MODEL.filter(e=>!e.external&&!e.context&&e.placed_date&&e.state==="conflict"&&(!sem||e.semester===sem));
  if(unresolved.length){return exportModal(`<h3>⚠ Cannot export ${esc(batch)} yet</h3>`+
    `<div class="row">${unresolved.length} unresolved conflict(s) in ${esc(batch)} — resolve or approve each first:</div>`+
    unresolved.slice(0,50).map(e=>`<div class="muted">• ${esc(e.course)} <span class="pill">${esc(e.cohort)}</span> ${esc(e.kinds.join("+"))} on ${esc(e.placed_date)} ${esc(e.slot)} <button class="small" onclick="openPopup(${e.id})">resolve…</button></div>`).join("")+
    `<div class="act"><button onclick="closePopup()">Close</button></div>`);}
  pendingExportSem=sem;
  reviewExportModal(sem);                     // trust gate: show every change before writing files
}
// human-readable record of every planning change for this batch (also saved as the decision log)
function fmtPlace(wk,wd,slot){return "W"+wk+" "+(wd||"?")+" "+slot;}
function buildDecisionLog(sem){
  const log=[]; const inSem=e=>!sem||e.semester===sem;
  const moved=ovGet("moved"),removed=ovGet("removed"),approved=ovGet("approved"),rooms=ovGet("room"),gone=ovGet("removed_lectures");
  const raw=id=>RAW.find(x=>x.id==id), cur=id=>MODEL.find(x=>x.id==id);
  for(const id in moved){const o=raw(id),c=cur(id); if(!o||!c||!inSem(c))continue;
    log.push({type:"move",text:`${o.course} (${o.cohort}): ${fmtPlace(o.week,o.placed_weekday,o.slot)} → ${fmtPlace(c.week,c.wd,c.slot)}${o.ai_placed?" (was system-placed)":""}`});}
  for(const id in gone){const o=raw(id); if(!o||!inSem(o))continue;
    log.push({type:"remove-lecture",text:`Removed lecture: ${o.course} (${o.cohort})`});}
  for(const id in removed){const o=raw(id),c=cur(id); if(!o||!(removed[id]||[]).length||(c&&!inSem(c)))continue;
    log.push({type:"remove-teacher",text:`${o.course} (${o.cohort}): removed teacher ${removed[id].join(", ")}`});}
  for(const id in rooms){const o=raw(id),c=cur(id); if(!o||!c||!inSem(c))continue;
    log.push({type:"room",text:`${o.course} (${o.cohort}): room ${o.room||"—"} → ${rooms[id]||"—"}`});}
  for(const id in approved){const c=cur(id); if(!c||c.state!=="ok"||c.pre_ok||!inSem(c))continue;
    log.push({type:"approve",text:`${c.course} (${c.cohort}) ${c.placed_date} ${c.slot}: approved double-booking${c.kinds&&c.kinds.length?" ("+c.kinds.join("+")+")":""}`});}
  return log;
}
function reviewExportModal(sem){
  const batch=sem||"both semesters";
  const log=buildDecisionLog(sem);
  const byType={}; log.forEach(d=>byType[d.type]=(byType[d.type]||0)+1);
  const sysPlaced=MODEL.filter(e=>!e.external&&!e.context&&e.placed_date&&e.ai_placed&&!e.moved&&(!sem||e.semester===sem)).length;
  const approvedDb=MODEL.filter(e=>!e.external&&!e.context&&e.state==="ok"&&!e.pre_ok&&(!sem||e.semester===sem)).length;
  let h=`<h3>Review before export — ${esc(batch)}</h3>`+
    `<div class="row muted">Everything below goes into the booker Excel and a saved decision log. Check it, then confirm. Hard rules are enforced again on write.</div>`+
    `<div class="box"><b>${log.length} change(s)</b> `+
      (Object.entries(byType).map(([k,n])=>`<span class="pill">${n} ${esc(k)}</span>`).join(" ")||`<span class="muted">none</span>`)+
      (approvedDb?` <span class="pill" style="background:rgba(236,72,153,.18)">⚠ ${approvedDb} approved double-booking</span>`:"")+
      (sysPlaced?`<div class="row muted">${sysPlaced} lecture(s) kept at their system-chosen slot (no specific day/time was requested).</div>`:"")+`</div>`+
    (log.length
      ? `<div class="box" style="max-height:300px;overflow:auto">`+log.map(d=>
          `<div class="ifind ${(d.type==="approve"||d.type==="remove-lecture")?"warn":"info"}"><b>${esc(d.type)}</b> — ${esc(d.text)}</div>`).join("")+`</div>`
      : `<div class="row muted">No manual changes — the plan exports exactly as imported.</div>`)+
    `<div class="act"><button class="primary" onclick="doExport()">✓ Confirm &amp; write Excel</button><button onclick="closePopup()">Cancel</button></div>`;
  exportModal(h);
}
async function doExport(){
  const sem=pendingExportSem||"", batch=sem||"both semesters";
  toast("Exporting "+batch+" to Excel…","ok");
  try{
    const res=await fetch("/export",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({bookings:buildPlanPayload(),semester:sem||null,decisions:buildDecisionLog(sem)})});
    const j=await res.json();
    if(!j.ok){ if(j.error==="unresolved_conflicts") return exportModal(`<h3>⚠ Export blocked</h3>`+j.unresolved.slice(0,50).map(x=>`<div class="muted">• ${esc(x)}</div>`).join("")+`<div class="act"><button onclick="closePopup()">Close</button></div>`);
      return toast("Export failed: "+esc(j.message||j.error),"bad"); }
    exportModal(`<h3>✓ Excel exported (${esc(j.semester||batch)})</h3><div class="row">${j.total_rows} sessions written to <b>${esc(j.export_dir)}</b>.</div>`+
      (j.approved_double_bookings?`<div class="row">${j.approved_double_bookings} approved double-booking(s) noted in the comments.</div>`:"")+
      `<div class="box"><b>Files</b>`+j.files.map(f=>`<div class="muted">• ${esc(f.file)} — ${f.rows} rows</div>`).join("")+
      (j.decision_log?`<div class="row muted">📝 Decision log: ${esc(j.decision_log)}</div>`:"")+`</div>`+
      (j.warnings&&j.warnings.length?`<div class="box"><b>Warnings to check</b>`+j.warnings.map(w=>`<div class="muted">• ${esc(w)}</div>`).join("")+`</div>`:"")+
      `<div class="act"><button class="primary" onclick="closePopup()">Done</button></div>`);
  }catch(err){ toast("Export failed — is the server running (py scripts/serve.py)? "+esc(""+err),"bad"); }
}
function exportModal(html){curPopup=null;$("#card").innerHTML=`<span class="x" onclick="closePopup()">×</span>`+html;$("#modal").classList.add("open");}

/* ---- import screen (upload → validate → approve → load) ----------------- */
async function renderImport(){
  try{ const r=await fetch("/uploads"); IMPUP=(await r.json()).files||[]; }
  catch(e){ IMPUP=null; }
  if(IMPUP===null){
    $("#view").innerHTML=`<div class="imp box">⚠ The Import screen needs the local server.<br>`+
      `Start the app with <b>run.bat</b> (or <code>py scripts/serve.py</code>) and open `+
      `<b>http://localhost:8765/dashboard.html</b>.</div>`;
    return;
  }
  let h=`<div class="imp">`+
    `<div class="step">1 · Upload the teachers' booking Excel files</div>`+
    `<div class="drop" id="drop" onclick="document.getElementById('impFile').click()">`+
      `Drag &amp; drop <b>.xlsx</b> files here, or click to choose.<br>`+
      `<span class="muted">One file per examiner. Blank template tabs are ignored.</span></div>`+
    (IMPUP.length
      ? `<div class="box" style="margin-top:8px"><b>Uploaded (${IMPUP.length})</b>`+
        IMPUP.map(f=>`<div class="fitem"><span>📄 ${esc(f)}</span></div>`).join("")+
        `<div class="act"><button class="primary" onclick="impValidate()">2 · Validate &amp; review →</button>`+
        `<button onclick="impClear()">Clear all</button></div></div>`
      : `<div class="muted" style="margin-top:8px">No files uploaded yet.</div>`);
  if(IMPVAL){
    let nf=0,withFix=0,appr=0; const byField={};
    IMPVAL.forEach((f,fi)=>f.courses.forEach((c,ci)=>c.findings.forEach((fd,k)=>{
      nf++; if(fd.suggested){withFix++; byField[fd.field]=(byField[fd.field]||0)+1; if(IMPAPP[fi+"|"+ci+"|"+k]==="yes")appr++;}})));
    const bulk=Object.entries(byField).map(([fld,n])=>
      `<div class="bulkrow"><span><b>${esc(fld)}</b> <span class="muted">(${n})</span></span>`+
      `<button class="small primary" onclick="impBulk('${esc(fld)}','yes')">✓ Approve all</button>`+
      `<button class="small warn" onclick="impBulk('${esc(fld)}','no')">✗ Reject all</button></div>`).join("");
    h+=`<div class="step">2 · Review suggested corrections</div>`+
      `<div class="muted">${nf} findings · ${withFix} with a fix · `+
      `<b style="color:var(--ok)">${appr} approved</b>. `+
      `Examiner = each course's Examinator field. Approve the fixes you trust — or decide in bulk:</div>`+
      (bulk?`<div class="row" style="margin:4px 0">${bulk}</div>`:"")+
      IMPVAL.map((f,fi)=>
        `<div class="step" style="font-size:14px">📑 ${esc(f.owner||f.file)} `+
          `<span class="muted">— ${esc(f.file)}</span></div>`+
        f.courses.map((c,ci)=>
          `<div class="box"><b>${esc(c.name||c.sheet)}</b> `+
          `<span class="pill">${esc(c.code||"no code")}</span> `+
          `<span class="pill">${c.n_rows} rows</span>`+
          (c.findings.length
            ? c.findings.map((fd,k)=>{const id=fi+"|"+ci+"|"+k,a=IMPAPP[id];
              return `<div class="ifind ${esc(fd.severity||"info")}"><b>${esc(fd.field)}</b> `+
                `<span class="muted">${esc(fd.confidence||"")}</span> — ${esc(fd.issue)}`+
                (fd.suggested
                  ? `<div style="margin-top:3px"><span style="color:var(--cf-teacher)">${esc(fd.current||"—")}</span>`+
                    ` → <span style="color:var(--ok);font-weight:600">${esc(fd.suggested)}</span></div>`+
                    `<div class="act"><button class="${a==="yes"?"primary":""}" onclick="impApprove('${id}','yes')">`+
                      `${a==="yes"?"✓ Approved":"Approve"}</button>`+
                    `<button class="${a==="no"?"warn":""}" onclick="impApprove('${id}','no')">`+
                      `${a==="no"?"✗ Rejected":"Reject"}</button></div>`
                  : "")+`</div>`;}).join("")
            : `<div class="muted" style="color:var(--ok)">✓ no issues</div>`)+
          `</div>`).join("")).join("")+
      `<div class="step">3 · Load into the planner</div>`+
      `<div class="box"><div class="row"><b>Academic year</b> <select id="impYear">${yearOpts()}</select>`+
        ` &nbsp; <b>Programme</b> <select id="impProg">${((SESSION&&SESSION.programmes)||["Media"]).map(p=>`<option ${p===(SESSION&&SESSION.programme)?"selected":""}>${esc(p)}</option>`).join("")}</select></div>`+
        `<div class="muted">Saved as a session (year + programme) you can switch to later. Re-importing the same year + programme overwrites it.</div>`+
        `<div class="muted" style="margin-top:4px">Replaces the current planner with these files (approved corrections applied). Conflicts are detected after import — use 🪄 Resolve all, then Export.</div>`+
      `<div class="act"><button class="primary" onclick="impApply()">⬇ Import approved into planner</button></div></div>`;
  }
  $("#view").innerHTML=h+`</div>`;
  const dz=$("#drop");
  if(dz){
    dz.ondragover=e=>{e.preventDefault();dz.classList.add("over");};
    dz.ondragleave=()=>dz.classList.remove("over");
    dz.ondrop=e=>{e.preventDefault();dz.classList.remove("over");impUpload(e.dataTransfer.files);};
  }
}
function impApprove(id,v){IMPAPP[id]=(IMPAPP[id]===v?undefined:v);renderImport();}
function impBulk(field,v){ // one decision for every finding of a field (e.g. all minute fixes)
  IMPVAL.forEach((f,fi)=>f.courses.forEach((c,ci)=>c.findings.forEach((fd,k)=>{
    if(fd.field===field&&fd.suggested)IMPAPP[fi+"|"+ci+"|"+k]=v;})));
  renderImport();
  toast((v==="yes"?"Approved":"Rejected")+" all "+field+" fixes.","ok");
}
async function impUpload(fileList){
  const files=[...fileList].filter(f=>f.name.toLowerCase().endsWith(".xlsx"));
  if(!files.length){toast("Only .xlsx files are accepted.","warn");return;}
  toast("Uploading "+files.length+" file(s)…","ok");
  const payload=[];
  for(const f of files){
    const b64=await new Promise(res=>{const r=new FileReader();r.onload=()=>res(r.result);r.readAsDataURL(f);});
    payload.push({name:f.name,b64});
  }
  const r=await fetch("/upload",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({files:payload})});
  const j=await r.json(); IMPVAL=null;
  toast("Uploaded "+((j.saved||[]).length)+" file(s).","ok"); renderImport();
}
async function impClear(){await fetch("/clear-uploads",{method:"POST"});IMPUP=[];IMPVAL=null;IMPAPP={};renderImport();}
async function impValidate(){
  toast("Validating…","ok");
  const r=await fetch("/validate",{method:"POST"});const j=await r.json();
  if(!j.ok){toast(j.message||"Validation failed","warn");return;}
  IMPVAL=j.files; IMPAPP={}; renderImport();
}
async function impApply(){
  const approved=[];
  IMPVAL.forEach((f,fi)=>f.courses.forEach((c,ci)=>c.findings.forEach((fd,k)=>{
    if(IMPAPP[fi+"|"+ci+"|"+k]==="yes"&&fd.suggested)
      approved.push({file:f.file,sheet:c.sheet,field:fd.field,to:fd.suggested});})));
  toast("Importing into the planner…","ok");
  const programme=($("#impProg")&&$("#impProg").value)||"Media";
  const year=($("#impYear")&&$("#impYear").value)||"";
  const r=await fetch("/apply",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({approved,programme,year})});
  const j=await r.json();
  if(!j.ok){toast("Import failed: "+(j.message||j.error||"error"),"warn");return;}
  // new data => previous planning edits (moved/removed/approved, keyed by booking id)
  // no longer match and would hide or misplace the freshly imported bookings. Clear them.
  try{ localStorage.removeItem("ba_overrides_v1"); }catch(e){}
  const gen=j.generated||{};
  const genLine=Object.keys(gen).length
    ? `<div class="row muted">Auto-generated from your files: `+
      ["teachers","courses","groups"].filter(k=>gen[k]).map(k=>`${gen[k].added||0} ${k} added`+(gen[k].flagged?`, ${gen[k].flagged} to review`:"")).join(" · ")+`.</div>`
    : "";
  const rep=j.report||{};
  const det=Object.entries(rep.detected||{}).sort((a,b)=>b[1]-a[1])
    .map(([c,n])=>`<span class="pill">${esc(c)}: ${n}</span>`).join(" ");
  const unm=(rep.unmatched||[]).slice(0,20).map(u=>
    `<div class="ifind warn"><b>${esc(u.sheet||"")}</b> ${esc(u.code||"")} — group “${esc(u.raw)}” not matched`+
    (u.suggest?` → suggest <b style="color:var(--ok)">${esc(u.suggest)}</b>`:"")+`</div>`).join("");
  const skp=(rep.skipped_list||[]).slice(0,20).map(s=>
    `<div class="ifind ${(s.reason||"").includes("recognised")?"warn":"info"}">${esc(s.sheet||"")} ${esc(s.code||"")}: ${esc(s.reason||"")}`+
    (s.groups?` <span class="muted">(groups: ${esc(s.groups)})</span>`:"")+`</div>`).join("");
  const details=`<details style="margin-top:6px"><summary style="cursor:pointer"><b>📊 Import details</b> — rows read ${rep.rows_read||0}, bookings created ${rep.bookings_created||0}, skipped ${rep.skipped||0} (${rep.files||0} file(s))</summary>`+
    `<div class="box"><div class="row"><b>Group codes detected:</b> ${det||"—"}</div>`+
    (unm?`<div class="row"><b>Group codes that failed to match</b> (not dropped — placed under “Unassigned”):</div><div style="max-height:200px;overflow:auto">${unm}</div>`:"")+
    (skp?`<div class="row muted" style="margin-top:6px">Rows needing attention:</div><div style="max-height:160px;overflow:auto">${skp}</div>`:"")+
    `</div></details>`;
  exportModal(`<h3>✓ Imported into the planner</h3>`+
    `<div class="row">${j.sessions} session(s) loaded`+(j.context?` (+${j.context} context)`:"")+
      ` across ${(j.cohorts||[]).length} cohort(s) · ${approved.length} correction(s) applied.</div>`+
    `<div class="row muted">Cohorts: ${(j.cohorts||[]).map(c=>esc(c)).join(", ")||"—"}</div>`+
    genLine+
    (j.skipped?`<div class="row muted">${j.skipped} row(s) need attention (unrecognised group / no week) — shown below and under the “Unassigned” cohort.</div>`:"")+
    details+
    `<div class="act"><button class="primary" onclick="goPlanner()">Open the planner →</button></div>`);
}
function goPlanner(){ try{ localStorage.removeItem("ba_overrides_v1"); }catch(e){}
  // cache-busting reload that opens directly on the planning calendar (not Home)
  location.href="/dashboard.html?view=calendar&t="+Date.now(); }

/* ---- Home (start screen) ----------------------------------------------- */
function go(v){document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("active",x.dataset.view===v));view=v;render();}
async function renderHome(){
  let s={}; try{ s=await (await fetch("/status")).json(); }catch(e){ s={offline:true}; }
  const steps=[
    ["📥","Import files","Upload the teachers' Excel and load them","go('import')"],
    ["🔎","Review &amp; fix","Check detected data, approve corrections","go('import')"],
    ["📅","Plan timetable","See the weekly calendar, drag lectures","go('calendar')"],
    ["⚠️","Resolve conflicts","Auto-resolve or fix clashes","go('conflicts')"],
    ["⬇","Export Excel","Write the final booker files","exportExcel()"]];
  let h=`<div class="home"><div class="hero"><h2>Booking Assistant</h2>`+
    `<div class="muted">Turn teachers' booking requests into clean, conflict-free Excel files — all in one place, no spreadsheets to edit by hand.</div></div>`;
  if(s.offline){
    h+=`<div class="box">⚠ Start the app with <b>BookingAssistant.exe</b> (or run.bat) to enable the full workflow.</div>`;
  }else{
    h+=`<b>How it works</b><div class="steps">`+steps.map((st,i)=>
      `<div class="stepcard" onclick="${st[3]}"><div><span class="n">${i+1}</span><b>${st[0]} ${st[1]}</b></div><div class="d">${st[2]}</div></div>`).join("")+`</div>`;
    h+=`<b>Status</b><div class="statgrid">`+
      `<div class="stat"><b>${s.bookings||0}</b><span>sessions loaded</span></div>`+
      `<div class="stat"><b>${s.teachers||0}</b><span>teachers</span></div>`+
      `<div class="stat"><b>${s.courses||0}</b><span>courses</span></div>`+
      `<div class="stat"><b>${s.ai?"On":"Off"}</b><span>AI assist</span></div></div>`+
      (!s.bookings?`<div class="box" style="margin-top:8px">No bookings loaded yet. Start with <b>📥 Import</b> to add the teachers' booking files, or set up teachers/courses/groups in <b>⚙ Manage</b>.</div>`:"")+
      `<div class="box" style="margin-top:10px"><b>Where files are saved</b>`+
        `<div class="muted">Final Excel → <code>${esc(s.export_dir||"export")}</code> &nbsp;·&nbsp; data source → <code>${esc(s.data_dir||"")}</code></div></div>`+
      `<div class="act" style="margin-top:10px"><button class="primary" onclick="go('import')">📥 Start: import files</button>`+
        `<button onclick="go('calendar')">📅 Open planner</button>`+
        `<button onclick="exportExcel()">⬇ Export Excel</button>`+
        `<button onclick="go('manage')">⚙ Manage teachers &amp; courses</button></div>`;
  }
  $("#view").innerHTML=h+`</div>`;
}

/* ---- Manage (teachers / courses / groups / settings) ------------------- */
let mtab="teachers";
function mrow(fields,vals){return fields.map(f=>`<td><input class="${f.cls||""}" data-f="${f.f}" value="${esc(vals[f.f]||"")}" placeholder="${f.ph||""}"></td>`).join("");}
function mgrAppendRow(tbId,fields,vals){const tb=$("#"+tbId);if(!tb)return;const tr=document.createElement("tr");
  tr.innerHTML=mrow(fields,vals||{})+`<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td>`;
  tb.appendChild(tr);return tr;}
function mgrAddRow(tbId,fields){const tr=mgrAppendRow(tbId,fields,{});const inp=tr&&tr.querySelector("input");if(inp)inp.focus();}
function mgrCollect(tbId){const tb=$("#"+tbId);if(!tb)return [];
  return [...tb.querySelectorAll("tr")].map(tr=>{const o={};
    tr.querySelectorAll("input,select").forEach(i=>o[i.dataset.f]=(i.type==="checkbox")?i.checked:i.value.trim());
    return o;}).filter(o=>Object.keys(o).some(k=>o[k]!==""&&o[k]!==false&&o[k]!==undefined));}
function fileB64(f){return new Promise(res=>{const r=new FileReader();r.onload=()=>res(r.result);r.readAsDataURL(f);});}
function impBar(kind){return `<div class="act"><button onclick="mgrImport('${kind}')">📥 Import from Excel</button>`+
  `<a href="/template/${kind}" download class="tmpl">⬇ Download template</a></div>`;}
let cfgImportKind=null;
function mgrImport(kind){cfgImportKind=kind;$("#cfgFile").click();}
async function cfgFilePicked(f){
  if(!f)return; toast("Reading "+f.name+"…","ok");
  let d; try{ const b64=await fileB64(f);
    d=await (await fetch("/import/"+cfgImportKind,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({b64})})).json();
  }catch(e){ toast("Could not read that file.","warn"); return; }
  if(!d.ok){ toast("Import failed: "+esc(d.message||d.error||""),"warn"); return; }
  importPreview(cfgImportKind,d);
}
function importPreview(kind,d){
  const cols=Object.entries(d.matched_columns||{}).map(([k,v])=>`<span class="pill">${esc(k)} ← ${esc(v)}</span>`).join(" ")||`<span class="muted">columns guessed by position</span>`;
  let body="",n=0;
  if(kind==="groups"){
    n=(d.groups||[]).length+(d.programs||[]).length;
    body=`<div class="muted">${(d.programs||[]).length} programme(s), ${(d.groups||[]).length} group code(s).</div>`+
      (d.programs||[]).map(p=>`<div class="ifind info"><b>${esc(p.code)}</b> ${esc(p.name||"")}</div>`).join("")+
      (d.groups||[]).map(g=>`<div class="ifind ${g.issue?"warn":"info"}"><b>${esc(g.code)}</b> ${g.issue?`<span class="muted">— ${esc(g.issue)}</span>`:""}</div>`).join("");
  }else{
    const items=d.items||[]; n=items.length;
    body=items.slice(0,200).map(it=>{const bad=it.issue;
      const label=kind==="teachers"?`${esc(it.name)} ${it.aliases&&it.aliases.length?`<span class="muted">(${esc(it.aliases.join("; "))})</span>`:""}`
        :`<b>${esc(it.code)}</b> ${esc(it.name||"")} ${it.ects?`<span class="muted">${esc(it.ects)} ECTS</span>`:""}`;
      return `<div class="ifind ${bad?"warn":"info"}">${label}${bad?` <span style="color:var(--cf-room)">— ${esc(bad)}</span>`:""}</div>`;}).join("");
  }
  exportModal(`<h3>Review import — ${esc(kind)}</h3><div class="row muted">Detected columns: ${cols}</div>`+
    `<div class="box" style="max-height:340px;overflow:auto">${body||"<span class='muted'>Nothing found.</span>"}</div>`+
    `<div class="act"><button class="primary" onclick="importMerge('${kind}')">Add ${n} to the list</button><button onclick="closePopup()">Cancel</button></div>`);
  window._imp=d;
}
function importMerge(kind){
  const d=window._imp||{};
  if(kind==="teachers")(d.items||[]).forEach(t=>{ if(!mgrCollect("tbT").some(x=>x.name.toLowerCase()===t.name.toLowerCase()))
    mgrAppendRow("tbT",[{f:"name"},{f:"aliases"}],{name:t.name,aliases:(t.aliases||[]).join("; ")}); });
  else if(kind==="courses")(d.items||[]).forEach(c=>{ if(!mgrCollect("tbC").some(x=>x.code.toUpperCase()===c.code.toUpperCase()))
    mgrAppendRow("tbC",[{f:"code",cls:"sm"},{f:"name"},{f:"ects",cls:"sm"},{f:"examiner"},{f:"notes"}],{code:c.code,name:c.name,ects:c.ects,examiner:"",notes:c.program?("programme: "+c.program):""}); });
  else if(kind==="groups"){
    (d.programs||[]).forEach(p=>{ if($("#tbP")&&!mgrCollect("tbP").some(x=>(x.code||"").toUpperCase()===p.code.toUpperCase()))
      mgrAppendRow("tbP",[{f:"code",cls:"sm"},{f:"name"}],{code:p.code,name:p.name}).querySelector('td:nth-child(3)').innerHTML='<input type="checkbox" data-f="active" checked>'; });
    (d.groups||[]).forEach(g=>{ if($("#tbX")&&!mgrCollect("tbX").some(x=>(x.code||"").toUpperCase()===g.code.toUpperCase()))
      mgrAppendRow("tbX",[{f:"code"}],{code:g.code}); });
  }
  closePopup(); toast("Added — review, then Save & apply.","ok");
}
function realizedImport(){
  const inp=document.createElement("input"); inp.type="file"; inp.accept=".xlsx,.csv"; inp.multiple=true;
  inp.onchange=async()=>{ const files=[...inp.files]; if(!files.length) return;
    toast("Uploading "+files.length+" file(s)…","ok");
    const payload=[]; for(const f of files) payload.push({name:f.name,b64:await fileB64(f)});
    try{ const j=await (await fetch("/import/realized",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({files:payload})})).json();
      exportModal(`<h3>📁 Realized bookings imported</h3><div class="row">Saved <b>${j.saved||0}</b> file(s) to <code>${esc(j.dir||"")}</code>.</div>`+
        `<div class="row muted">Stored for now — reading them into the analysis is a planned next step.</div>`+
        `<div class="act"><button class="primary" onclick="closePopup()">OK</button></div>`);
    }catch(e){ toast("Upload failed — is the server running?","warn"); }
  };
  inp.click();
}
async function renderManage(){
  $("#view").innerHTML=`<div class="manage"><div class="mtabs">`+
    [["teachers","👤 Teachers"],["courses","📚 Courses"],["groups","🎓 Groups & programmes"],["settings","⚙ Settings"]]
      .map(([k,l])=>`<div class="mtab ${mtab===k?"active":""}" onclick="mtab='${k}';renderManage()">${l}</div>`).join("")+
    `</div><div id="mbody" class="muted">Loading…</div></div>`;
  try{
    if(mtab==="teachers") await mgrTeachers();
    else if(mtab==="courses") await mgrCourses();
    else if(mtab==="groups") await mgrGroups();
    else await mgrSettings();
  }catch(e){ $("#mbody").innerHTML=`<div class="box">⚠ Needs the local server (BookingAssistant.exe / run.bat).</div>`; }
}
async function mgrTeachers(){
  const d=await (await fetch("/config/teachers")).json();
  const tf=[{f:"name",ph:"Full name"},{f:"aliases",ph:"nicknames; spellings"}];
  const yf=[{f:"wrong",ph:"misspelling in files"},{f:"correct",ph:"correct name"}];
  $("#mbody").innerHTML=`<div class="muted">Your teaching team. Import from an Excel file, or add manually. Aliases catch nicknames or alternate spellings used in the booking files.</div>`+
    impBar("teachers")+
    `<table class="gtable"><thead><tr><th>Name</th><th>Aliases (optional)</th><th></th></tr></thead><tbody id="tbT">`+
      d.teachers.map(t=>`<tr>${mrow(tf,{name:t.name,aliases:(t.aliases||[]).join("; ")})}<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("")+
    `</tbody></table><div class="act"><button onclick="mgrAddRow('tbT',[{f:'name',ph:'Full name'},{f:'aliases',ph:'nicknames; spellings'}])">+ Add teacher</button></div>`+
    `<div class="step" style="font-size:14px">Name corrections <span class="muted">(auto-fix a known misspelling → correct name)</span></div>`+
    `<table class="gtable"><thead><tr><th>Wrong spelling</th><th>Correct name</th><th></th></tr></thead><tbody id="tbTy">`+
      d.typos.map(t=>`<tr>${mrow(yf,t)}<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("")+
    `</tbody></table><div class="act"><button onclick="mgrAddRow('tbTy',[{f:'wrong',ph:'misspelling'},{f:'correct',ph:'correct name'}])">+ Add correction</button></div>`+
    mgrSaveBar("mgrSaveTeachers()");
}
async function mgrSaveTeachers(){
  const teachers=mgrCollect("tbT").map(o=>({name:o.name,aliases:(o.aliases||"").split(";").map(s=>s.trim()).filter(Boolean)}));
  const typos=mgrCollect("tbTy").filter(o=>o.wrong&&o.correct);
  if(!teachers.length){toast("Add at least one teacher.","warn");return;}
  await mgrApply("/config/teachers",{teachers,typos});
}
async function mgrCourses(){
  const d=await (await fetch("/config/courses")).json();
  const cf=[{f:"code",cls:"sm",ph:"AS-1-044"},{f:"name",ph:"Course name"},{f:"ects",cls:"sm",ph:"5"},{f:"examiner",ph:"examiner"},{f:"notes",ph:"optional"}];
  $("#mbody").innerHTML=`<div class="muted">Courses (code → name, ECTS, examiner). Generated automatically from your imported files; import a list, or add manually. The <b>examiner</b> is the responsible teacher for each course.</div>`+
    impBar("courses")+
    `<table class="gtable"><thead><tr><th>Code</th><th>Name</th><th>ECTS</th><th>⭐ Examiner</th><th>Notes</th><th></th></tr></thead><tbody id="tbC">`+
      d.courses.map(c=>`<tr>${mrow(cf,c)}<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("")+
    `</tbody></table><div class="act"><button onclick="mgrAddRow('tbC',[{f:'code',cls:'sm',ph:'AS-1-044'},{f:'name',ph:'Course name'},{f:'ects',cls:'sm',ph:'5'},{f:'examiner',ph:'examiner'},{f:'notes',ph:'optional'}])">+ Add course</button></div>`+
    mgrSaveBar("mgrSaveCourses()");
}
async function mgrSaveCourses(){
  const courses=mgrCollect("tbC").filter(c=>c.code);
  if(!courses.length){toast("Add at least one course.","warn");return;}
  await mgrApply("/config/courses",{courses});
}
async function mgrGroups(){
  const d=await (await fetch("/config/programs")).json();
  const ref=(d.reference||[]).map(p=>`<option value="${esc(p.code)}|${esc(p.name)}">${esc(p.code)} — ${esc(p.name)}</option>`).join("");
  const progRows=(d.programs||[]).map(p=>`<tr><td><input class="sm" data-f="code" value="${esc(p.code)}"></td>`+
    `<td><input data-f="name" value="${esc(p.name)}"></td>`+
    `<td style="text-align:center"><input type="checkbox" data-f="active" ${p.active?"checked":""}></td>`+
    `<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("");
  const trackRows=Object.entries(d.tracks||{}).map(([k,v])=>`<tr>${mrow([{f:"letter",cls:"sm"},{f:"label"}],{letter:k,label:v})}<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("");
  const extraRows=(d.extra||[]).map(g=>`<tr>${mrow([{f:"code",cls:"sm"}],{code:g})}<td><button class="small warn" onclick="this.closest('tr').remove()">✕</button></td></tr>`).join("");
  $("#mbody").innerHTML=
    `<div class="muted">Student groups are generated as <b>CODE-YY</b> for the active programmes and intake years (e.g. <b>Media-26</b>, <b>FT-26</b>, <b>KP-26</b>). Supports more than Film &amp; Media.</div>`+
    impBar("groups")+
    `<div class="box"><b>Active intake years</b> <span class="muted">newest cohorts only — older years drop off automatically each year</span>`+
      `<div class="row"><b>Newest year</b> <input class="sm" id="gBase" value="${esc(d.base_year)}"> <b style="margin-left:8px">show</b> <input class="sm" id="gWin" value="${esc(d.window)}"> <b>years</b> `+
      `<span class="muted">→ currently <b>${(d.years||[]).join(", ")}</b></span></div></div>`+
    `<div class="box"><b>Programmes</b> <span class="muted">tick the ones you plan; codes are generated for ticked programmes</span>`+
      `<table class="gtable"><thead><tr><th>Code</th><th>Name</th><th>Active</th><th></th></tr></thead><tbody id="tbP">${progRows}</tbody></table>`+
      `<div class="act"><select id="refSel"><option value="">— add from list —</option>${ref}</select>`+
      `<button onclick="addProgFromRef()">+ Add programme</button>`+
      `<button onclick="mgrAppendRow('tbP',[{f:'code',cls:'sm'},{f:'name'}],{}).querySelector('td:nth-child(3)').innerHTML='&lt;input type=checkbox data-f=active checked&gt;'">+ Add custom</button></div></div>`+
    `<div class="box"><b>Media tracks</b> <span class="muted">extra Media-YY-X codes (Foto, Ljud, …)</span>`+
      `<table class="gtable"><thead><tr><th>Letter</th><th>Label</th><th></th></tr></thead><tbody id="tbMT">${trackRows}</tbody></table>`+
      `<div class="act"><button onclick="mgrAddRow('tbMT',[{f:'letter',cls:'sm'},{f:'label'}])">+ Add track</button></div></div>`+
    `<div class="box"><b>Extra group codes</b> <span class="muted">any specific codes to add by hand or from Excel</span>`+
      `<table class="gtable"><thead><tr><th>Group code</th><th></th></tr></thead><tbody id="tbX">${extraRows}</tbody></table>`+
      `<div class="act"><button onclick="mgrAddRow('tbX',[{f:'code',cls:'sm'}])">+ Add group</button></div></div>`+
    `<div class="box"><b>Generated group codes</b> <span class="muted">(updates after Save &amp; apply)</span>`+
      `<div class="row" style="max-height:120px;overflow:auto">${(d.generated||[]).map(c=>`<span class="pill">${esc(c)}</span>`).join(" ")}</div></div>`+
    mgrSaveBar("mgrSaveGroups()");
}
function addProgFromRef(){const v=$("#refSel").value;if(!v)return;const [code,name]=v.split("|");
  if(mgrCollect("tbP").some(x=>(x.code||"").toUpperCase()===code.toUpperCase()))return;
  mgrAppendRow("tbP",[{f:"code",cls:"sm"},{f:"name"}],{code,name}).querySelector("td:nth-child(3)").innerHTML='<input type="checkbox" data-f="active" checked>';}
async function mgrSaveGroups(){
  const programs=mgrCollect("tbP").filter(p=>p.code).map(p=>({code:p.code,name:p.name,active:p.active!==false}));
  const tracks={}; mgrCollect("tbMT").forEach(o=>{if(o.letter)tracks[o.letter]=o.label||o.letter;});
  const extra=mgrCollect("tbX").map(o=>o.code).filter(Boolean);
  await mgrApply("/config/programs",{programs,tracks,extra,base_year:+$("#gBase").value||26,window:+$("#gWin").value||4});
}
async function mgrSettings(){
  const [s,ai,rl]=await Promise.all([fetch("/config/settings").then(r=>r.json()),fetch("/config/ai").then(r=>r.json()),fetch("/config/rules").then(r=>r.json())]);
  const fld=(id,label,val,hint)=>`<div class="row"><b>${label}</b> <input id="${id}" class="sm" value="${esc(val)}"> <span class="muted">${hint||""}</span></div>`;
  const modelOpts=(ai.models||[]).map(m=>`<option value="${esc(m.id)}" ${m.id===ai.model?"selected":""}>${esc(m.label)}</option>`).join("");
  const rec=(ai.models||[]).map(m=>`<div class="airow"><b>${esc(m.label)}</b> <span class="muted">— ${esc(m.note)}</span></div>`).join("");
  $("#mbody").innerHTML=
    `<div class="box"><b>🤖 AI conflict solver</b> <span class="muted">— optional. Powers the <b>Resolve all (AI)</b> button and per-conflict suggestions.</span>`+
      `<div class="row"><b>API key</b> <input id="aiKey" type="password" autocomplete="off" placeholder="${ai.has_key?"•••••• saved — leave blank to keep":"sk-ant-…"}" style="width:300px"> <span class="muted">${ai.has_key?"✓ a key is saved on this computer":"stored locally in .env, next to the app"}</span></div>`+
      `<div class="row"><b>Model</b> <select id="aiModel">${modelOpts}</select></div>`+
      `<div class="box" style="margin:6px 0;background:var(--panel2)"><div class="muted" style="margin-bottom:4px">Which model? (prices are rough estimates)</div>${rec}</div>`+
      `<div class="row"><b>Spending cap</b> $ <input id="aiCap" class="sm" value="${esc(ai.cap_usd||0)}"> <span class="muted">per month — 0 means no cap. Used this month ≈ $${(ai.spent_usd||0).toFixed(2)} (estimate)</span></div>`+
      `<div class="row"><b>Custom instructions for solving conflicts</b></div>`+
      `<textarea id="aiInstr" rows="6" style="width:100%;box-sizing:border-box;background:var(--panel);border:1px solid var(--line);color:var(--txt);border-radius:6px;padding:7px">${esc(ai.instructions||"")}</textarea>`+
      `<div class="muted" style="margin-top:4px">For example: which groups, teachers or rooms may sometimes be double-booked · which clashes are never allowed · high-priority courses · preferred days or times.</div>`+
      `<div class="act"><button class="primary" onclick="mgrSaveAI()">💾 Save AI settings</button><span id="aisaved" class="muted"></span></div></div>`+
    `<div class="box"><b>📜 AI rules file</b> <span class="muted">— the project rules the AI follows (group logic, conflicts, studio priority, validation, export, planning). Edit here or in <code>${esc((rl.path||"booking-assistant-rules.md"))}</code>.</span>`+
      `<textarea id="aiRules" rows="12" style="width:100%;box-sizing:border-box;font-family:monospace;font-size:12px;background:var(--panel);border:1px solid var(--line);color:var(--txt);border-radius:6px;padding:7px">${esc(rl.rules||"")}</textarea>`+
      `<div class="act"><button class="primary" onclick="mgrSaveRules()">💾 Save rules</button><span id="rulesaved" class="muted"></span></div></div>`+
    `<div class="box"><b>Folders</b> <span class="muted">(all next to the app — created automatically)</span>`+
      `<div class="row"><b>Data source</b> <input id="sData" value="${esc(s.data_dir)}" style="width:240px"> <span class="muted">${esc(s.data_dir_full||"")}</span></div>`+
      `<div class="row"><b>📤 Export (final Excel)</b> <code>${esc(s.export_dir)}</code></div>`+
      `<div class="row"><b>📥 Import / uploads</b> <code>${esc(s.import_dir)}</code></div>`+
      `<div class="row"><b>📄 Templates</b> <code>${esc(s.templates_dir)}</code></div></div>`+
    `<div class="box"><b>Workload targets</b> <span class="muted">(planning guides only — used for the "possible issues" flags)</span>`+
      fld("sHpe","Hours per ECTS (course-work)",s.hours_per_ects_coursework,"5 ECTS ≈ 100 h")+
      fld("sYlo","Yearly course-work — low",s.yearly_low,"warn below")+
      fld("sYhi","Yearly course-work — high",s.yearly_high,"warn above")+
      fld("sIlo","In-class h per ECTS — warn low",s.inclass_warn_low,"")+
      fld("sIhi","In-class h per ECTS — warn high",s.inclass_warn_high,"")+`</div>`+
    mgrSaveBar("mgrSaveSettings()");
}
async function mgrSaveSettings(){
  await mgrApply("/config/settings",{data_dir:$("#sData").value.trim(),
    hours_per_ects_coursework:+$("#sHpe").value||20,yearly_low:+$("#sYlo").value||800,yearly_high:+$("#sYhi").value||1200,
    inclass_warn_low:+$("#sIlo").value||8,inclass_warn_high:+$("#sIhi").value||14});
}
async function mgrSaveAI(){
  const sv=$("#aisaved"); if(sv)sv.textContent=" saving…";
  const payload={model:$("#aiModel").value,cap_usd:+$("#aiCap").value||0,instructions:$("#aiInstr").value};
  const k=$("#aiKey").value.trim(); if(k)payload.api_key=k;
  try{
    await fetch("/config/ai",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
    AI=await (await fetch("/ai/status")).json(); updateSolverButtons();
    if(sv)sv.textContent=" ✓ saved"; toast("AI settings saved"+(k?" — AI is now enabled.":"."),"ok");
  }catch(e){ toast("Could not save AI settings.","warn"); if(sv)sv.textContent=""; }
}
async function mgrSaveRules(){
  const sv=$("#rulesaved"); if(sv)sv.textContent=" saving…";
  try{ await fetch("/config/rules",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({rules:$("#aiRules").value})});
    if(sv)sv.textContent=" ✓ saved"; toast("AI rules saved.","ok");
  }catch(e){ toast("Could not save rules.","warn"); if(sv)sv.textContent=""; }
}
function mgrSaveBar(saveFn){
  return `<div class="savebar"><button class="primary" onclick="${saveFn}">💾 Save &amp; apply</button>`+
    `<span class="muted">Saves your changes and rebuilds the planner so they take effect.</span><span id="msaved"></span></div>`;
}
async function mgrApply(endpoint,payload){
  const sv=$("#msaved"); if(sv)sv.textContent=" saving…";
  try{
    const r=await fetch(endpoint,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
    const j=await r.json(); if(!j.ok&&j.ok!==undefined){throw new Error(j.message||"save failed");}
    if(sv)sv.textContent=" rebuilding…";
    await fetch("/rebuild",{method:"POST"});
    toast("Saved and applied. Reloading…","ok");
    setTimeout(()=>location.reload(),500);
  }catch(e){ toast("Could not save: "+esc(""+(e.message||e)),"warn"); if(sv)sv.textContent=""; }
}

/* ---- bindings ---------------------------------------------------------- */
function bindFilters(){
  document.querySelectorAll(".ctab").forEach(t=>t.onclick=()=>{cohort=t.dataset.c;render();});
  document.querySelectorAll(".sf[data-main]").forEach(t=>t.onclick=()=>{const m=t.dataset.main;
    specSel=m==="all"?new Set(ALLSPECS):m==="film"?new Set(["F","L","M","P"]):new Set(["O"]);render();});
  document.querySelectorAll(".sf[data-s]").forEach(t=>t.onclick=()=>{const k=t.dataset.s;
    specSel.has(k)?specSel.delete(k):specSel.add(k); if(!specSel.size)specSel=new Set(ALLSPECS); render();});
  const ex=$("#extchk"); if(ex) ex.onchange=()=>{showExt=ex.checked;render();};
}
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>{document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
  t.classList.add("active");view=t.dataset.view;render();});
["semester","teacher","search"].forEach(id=>$("#"+id).addEventListener("input",render));
$("#undoBtn").onclick=undo;
$("#exportBtn").onclick=exportExcel;
$("#resetBtn").onclick=resetConfirm;
$("#resolveBtn").onclick=resolveAll;
$("#aiResolveBtn").onclick=aiResolveAll;
$("#helpBtn").onclick=helpModal;
$("#impFile").onchange=e=>{impUpload(e.target.files);e.target.value="";};
$("#cfgFile").onchange=e=>{const f=e.target.files[0];e.target.value="";cfgFilePicked(f);};
document.addEventListener("keydown",ev=>{ if((ev.ctrlKey||ev.metaKey)&&(ev.key==="z"||ev.key==="Z")){ev.preventDefault();undo();}});
document.addEventListener("click",e=>{const t=e.target.closest(".tlink");if(t){e.stopPropagation();openTeacher(t.dataset.t);}},true);
$("#view").addEventListener("click",e=>{const c=e.target.closest(".ev[data-id]");if(c&&!c.classList.contains("dragging"))openPopup(c.dataset.id);});
$("#modal").addEventListener("click",e=>{if(e.target.id==="modal")closePopup();});
// only the FM team (config/teacher_aliases.csv), not guest/visiting names
[...TEAM].sort().forEach(t=>{const o=document.createElement("option");o.value=o.textContent=t;$("#teacher").appendChild(o);});
fetch("/ai/status").then(r=>r.json()).then(s=>{AI=s;updateSolverButtons();}).catch(()=>{});  // optional; AI assist if a key is set
(function(){ const v=new URLSearchParams(location.search).get("view");  // open on a specific tab (e.g. after import)
  if(v){ const tab=[...document.querySelectorAll(".tab")].find(t=>t.dataset.view===v);
    if(tab){ document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active")); tab.classList.add("active"); view=v; } } })();
render();
</script>
</body>
</html>"""


def _load_targets():
    from .dictionaries import cfg_path
    path = cfg_path("workload_targets.json")
    try:
        with open(path, encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return {"hours_per_ects_coursework": 20,
                "inclass_per_ects": {"warn_low": 8, "normal_min": 10, "normal_max": 12, "warn_high": 14},
                "yearly_coursework": {"target_low": 800, "target_high": 1200, "working_year_hours": 1600},
                "student_hours_per_ects": 27, "teachers": {}}


def generate(bookings_csv=None):
    bookings_csv = bookings_csv or os.path.join(OUT, "bookings_2026_2027.csv")
    d = load_all()
    events, week_dates = build_events(bookings_csv, team=d.teachers)
    schedule(events)
    realized = build_realized(d)
    # session + holidays + production periods
    from . import sessions as _sess
    from .dictionaries import current_academic_year, programs
    from .holidays import holidays_for_academic_year
    act = _sess.active()
    ay = act.get("year") or current_academic_year()
    session_info = {"programme": act.get("programme") or "Media", "year": ay,
                    "list": _sess.list_sessions(), "programmes": list(programs().keys())}
    html = (TEMPLATE
            .replace("__DATA__", json.dumps(events, ensure_ascii=False))
            .replace("__WEEKDATES__", json.dumps(week_dates, ensure_ascii=False))
            .replace("__SPEC__", json.dumps(SPEC, ensure_ascii=False))
            .replace("__TEAM__", json.dumps(d.teachers, ensure_ascii=False))
            .replace("__SLOTS__", json.dumps(SLOT_TIMES, ensure_ascii=False))
            .replace("__TARGETS__", json.dumps(_load_targets(), ensure_ascii=False))
            .replace("__REALIZED__", json.dumps(realized, ensure_ascii=False))
            .replace("__RYEARS__", json.dumps(sorted(realized.keys()), ensure_ascii=False))
            .replace("__HOLIDAYS__", json.dumps(holidays_for_academic_year(ay), ensure_ascii=False))
            .replace("__PRODPERIODS__", json.dumps(_sess.meta().get("production_periods", []), ensure_ascii=False))
            .replace("__SESSION__", json.dumps(session_info, ensure_ascii=False)))
    out_path = os.path.join(OUT, "dashboard.html")
    os.makedirs(OUT, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path, len(events)


if __name__ == "__main__":
    p, n = generate()
    print(f"Wrote {p} ({n} calendar events)")
