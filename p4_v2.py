
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PRODUCTION ORC — DODO</title>
<style>
:root{
  --navy:#1a1f5e;--navy2:#2d3480;--green:#16a34a;--red:#dc2626;
  --amber:#d97706;--purple:#7c3aed;--blue:#0891b2;
  --bg:#f0f2f8;--card:#fff;--border:#dde4ef;--text:#1e293b;--gray:#64748b;
  --lgray:#e2e8f0;--radius:10px;--shadow:0 2px 10px rgba(0,0,0,.07);
  --hdr-h:52px;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);font-size:calc(13px*var(--zf,1));height:100vh;overflow:hidden;display:flex;flex-direction:column}

/* ── STOP ACTIVE THEME ── */
body.stop-on #app-hdr{background:#7f0000!important;border-color:#b91c1c}

/* ── HEADER ── */
#app-hdr{height:var(--hdr-h);background:var(--navy);display:flex;align-items:flex-end;padding:0 14px;gap:8px;flex-shrink:0;border-bottom:3px solid var(--navy2)}
.hdr-logo{color:#fff;font-weight:800;font-size:calc(15px*var(--zf,1));letter-spacing:1px;margin-right:10px;white-space:nowrap;align-self:center}
.hdr-tabs{display:flex;gap:3px;flex:1;align-self:flex-end}
.htab{background:rgba(255,255,255,.08);border:1.5px solid rgba(255,255,255,.18);border-bottom:3px solid var(--navy);color:rgba(255,255,255,.6);padding:6px 15px 8px;border-radius:8px 8px 0 0;cursor:pointer;font-size:calc(12px*var(--zf,1));font-weight:700;white-space:nowrap;transition:all .15s;position:relative;top:3px}
.htab:hover{background:rgba(255,255,255,.16);border-color:rgba(255,255,255,.35);border-bottom-color:var(--navy);color:#fff}
.htab.on{background:#fff;color:#1e3a8a;font-weight:800;border-color:rgba(255,255,255,.3);border-bottom:3px solid #fff;box-shadow:0 -3px 8px rgba(0,0,0,.12)}
.htab.prod-on{background:var(--green)!important;color:#fff!important;font-weight:700;animation:pt 2s infinite;border-color:var(--green)!important;border-bottom-color:var(--green)!important;box-shadow:none!important}
#ht-prod{display:none}
#ht-prod.prod-visible{display:inline-block!important}
@keyframes pt{0%,100%{opacity:1}50%{opacity:.75}}
#hdr-right{display:flex;align-items:center;gap:8px;margin-left:auto;font-size:calc(11px*var(--zf,1));color:rgba(255,255,255,.75);align-self:center}
#hdr-pilot-lbl{font-weight:800;color:#fff;font-size:calc(18px*var(--zf,1));letter-spacing:.3px}
#ht-guest-badge{display:none!important}
#main-prod-banner{display:none!important}

/* ── ALERT STRIP ── */
#alert-strip{background:#b91c1c;color:#fff;text-align:center;padding:4px;font-weight:700;font-size:calc(12px*var(--zf,1));flex-shrink:0;display:none;animation:blink .85s step-start infinite}
#alert-strip.on{display:block}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.35}}

/* ── VIEWS ── */
.view{display:none;flex:1;flex-direction:column;overflow:hidden}
.view.on{display:flex}

/* ── LOGIN ── */
#v-login{background:linear-gradient(135deg,#1a1f5e,#2d3480,#1e3a8a);align-items:center;justify-content:center}
.login-card{background:#fff;border-radius:14px;padding:32px;width:100%;max-width:380px;box-shadow:0 20px 60px rgba(0,0,0,.35)}
.lc-h1{color:var(--navy);font-size:calc(24px*var(--zf,1));font-weight:800;margin-bottom:2px;text-align:center}
.lc-sub{color:#64748b;text-align:center;margin-bottom:20px;font-size:calc(12px*var(--zf,1))}
.lf{margin-bottom:12px}
.lf label{display:block;font-weight:600;margin-bottom:3px;color:#374151;font-size:calc(11px*var(--zf,1));text-transform:uppercase;letter-spacing:.5px}
.lf select,.lf input{width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(14px*var(--zf,1));outline:none;transition:border .2s}
.lf select:focus,.lf input:focus{border-color:var(--navy)}
.btn-login{width:100%;padding:11px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:calc(14px*var(--zf,1));font-weight:700;cursor:pointer;margin-top:4px}
.btn-login:hover{opacity:.88}
.ln-err{color:#dc2626;text-align:center;margin-top:6px;font-size:calc(12px*var(--zf,1));min-height:16px}

/* ── MAIN VIEW (Déclarations / Évts) ── */
#v-main{overflow:hidden}
.main-hdr{background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap}
.main-hdr .mtabs{display:flex;gap:3px}
.mtab{background:none;border:none;border-bottom:2px solid transparent;padding:5px 12px;cursor:pointer;font-size:calc(12px*var(--zf,1));font-weight:600;color:var(--gray);transition:all .15s}
.mtab.on{border-color:var(--navy);color:var(--navy)}
.main-hdr .mbtns{display:flex;gap:6px;margin-left:auto}
.btn-sm{padding:6px 12px;border:none;border-radius:6px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer;transition:all .15s}
.btn-green{background:var(--green);color:#fff}
.btn-green:hover{filter:brightness(.9)}
.btn-ghost{background:var(--lgray);color:var(--text)}
.btn-ghost:hover{filter:brightness(.93)}
.table-wrap{flex:1;overflow-y:auto}
.ktbl{width:100%;border-collapse:collapse;font-size:calc(12px*var(--zf,1))}
.ktbl th{text-align:left;padding:7px 10px;background:var(--navy);color:#fff;font-size:calc(10px*var(--zf,1));text-transform:uppercase;letter-spacing:.5px;position:sticky;top:0}
.ktbl td{padding:6px 10px;border-bottom:1px solid var(--border)}
.ktbl tr:hover td{background:var(--lgray)}
.tg{color:var(--green);font-weight:700}
.tm{color:var(--amber);font-weight:700}
.tb{color:var(--red);font-weight:700}
.btn-tbl{padding:3px 8px;border:none;border-radius:4px;font-size:calc(10px*var(--zf,1));cursor:pointer;font-weight:600}

/* ── PRODUCTION VIEW ── */
#v-prod{overflow:hidden}
/* pilot/OF banner */
.pob{background:var(--navy);color:#fff;padding:6px 14px;display:flex;align-items:center;gap:24px;flex-shrink:0}
.pob-item{display:flex;flex-direction:column;gap:1px}
.pob-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.7px;opacity:.7;font-weight:600}
.pob-val{font-size:calc(22px*var(--zf,1));font-weight:800;line-height:1}
.pob-item.of .pob-val{font-size:calc(20px*var(--zf,1));color:#93c5fd}
.pob-item.trs .pob-val{color:#86efac}
/* status bar */
.sbar{display:flex;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.sc{flex:1;padding:6px 14px;border-right:1px solid var(--border);text-align:center}
.sc:last-child{border:none}
.sc-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.6px}
.sc-val{font-size:calc(22px*var(--zf,1));font-weight:800;font-variant-numeric:tabular-nums;color:var(--navy);margin-top:1px}
.sc-val.green{color:var(--green)}
.sc-val.red{color:var(--red)}
.sc-val.amber{color:var(--amber)}
/* main 2-col layout */
.prod-body{display:flex;flex:1;overflow:hidden}
/* CENTER form col (now left) */
.form-col{flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:6px}
/* Action buttons row below timeline */
.prod-act-row{display:flex;gap:6px;flex-wrap:wrap;padding:6px 0 2px;border-top:1px solid var(--border);margin-top:2px;position:sticky;bottom:0;background:var(--card);z-index:10}
.act-btn{flex:1;min-width:100px;border:none;border-radius:8px;padding:24px 6px;cursor:pointer;font-size:calc(16px*var(--zf,1));font-weight:700;text-align:center;transition:all .15s;white-space:nowrap;min-height:80px;display:flex;align-items:center;justify-content:center;gap:4px;flex-direction:column;line-height:1.3}
.act-btn:hover{filter:brightness(.9)}
.act-stop{background:linear-gradient(135deg,#b91c1c,#7f0000);color:#fff;font-size:calc(16px*var(--zf,1));font-weight:800;box-shadow:0 3px 8px rgba(185,28,28,.3)}
.act-nett{background:#e0f2fe;color:var(--blue)}
.act-pause{background:#f3e8ff;color:var(--purple)}
.act-cancel{background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1}
.act-endprod{background:linear-gradient(135deg,#16a34a,#15803d);color:#fff;font-weight:800;box-shadow:0 3px 8px rgba(22,163,74,.35)}
/* 3-col form zones */
.form-3col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.fzone{border-radius:7px;padding:8px}
.fzone h4{font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.7px;font-weight:700;margin-bottom:6px;padding-bottom:3px;border-bottom:1px solid rgba(0,0,0,.08)}
.zi{background:#eef2ff;border:1px solid #c7d2fe}.zi h4{color:#3730a3}
.zp{background:#f0fdf4;border:1px solid #bbf7d0}.zp h4{color:#166534}
.zq{background:#fff7ed;border:1px solid #fed7aa}.zq h4{color:#9a3412}
.fr{display:flex;flex-direction:column;margin-bottom:4px}
.fr label{font-size:calc(9px*var(--zf,1));font-weight:700;color:var(--gray);margin-bottom:2px;text-transform:uppercase;letter-spacing:.2px}
.fr input,.fr select,.fr textarea{padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1));background:var(--card);color:var(--text);width:100%;outline:none;transition:border .15s}
.fr input:focus,.fr select:focus{border-color:#6366f1}
.fr textarea{resize:none;height:42px}
.fr.comment-big textarea{height:80px;font-size:calc(13px*var(--zf,1));border:2px solid #f59e0b;background:#fffbeb;font-weight:500}
.fr.comment-big label{color:#d97706;font-size:calc(10px*var(--zf,1))}
input[type=checkbox]{cursor:pointer}
input:not([type=checkbox]):not([type=radio]):not([type=button]):not([type=submit]),textarea{cursor:text!important}
input,select,textarea{cursor:auto}
input[type=text],input[type=number],input[type=password],input[type=time],input[type=date],textarea{cursor:text!important}
.tl-legend{display:flex;gap:12px;padding:2px 4px;font-size:calc(10px*var(--zf,1));color:var(--gray);flex-wrap:wrap;align-items:center}
.tl-legend span{display:flex;align-items:center;gap:3px}
.tl-legend i{display:inline-block;width:12px;height:10px;border-radius:2px;flex-shrink:0}
select{cursor:default}
.fr.big input{font-size:calc(16px*var(--zf,1));font-weight:700;padding:5px 6px;color:var(--green)}
.fr.ro input{background:#f8fafc;color:var(--gray)}
/* Timeline */
.tl-wrap{background:var(--card);border-radius:7px;padding:7px 8px;border:1px solid var(--border)}
.tl-wrap h5{font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;letter-spacing:.7px;margin-bottom:4px}
/* RIGHT recap col (wider) */
.recap-col{width:280px;flex-shrink:0;border-left:1px solid var(--border);background:var(--card);display:flex;flex-direction:column;overflow:hidden}
.recap-hdr{font-size:calc(10px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.6px;padding:8px 8px 4px}
.recap-body{flex:1;overflow-y:auto;padding:0 6px 6px}
.si{display:flex;align-items:center;gap:5px;padding:3px 0;border-bottom:1px solid var(--border);font-size:calc(11px*var(--zf,1))}
.si:last-child{border:none}
.sdot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.si-nm{flex:1;font-weight:600;font-size:calc(10px*var(--zf,1));line-height:1.2}
.si-dur{font-size:calc(9px*var(--zf,1));color:var(--gray);white-space:nowrap}
.btn-edit{background:none;border:none;cursor:pointer;font-size:calc(11px*var(--zf,1));color:#6366f1;padding:1px 3px;border-radius:2px}
/* TRS gauge */
.gauge-box{padding:6px;border-top:1px solid var(--border);text-align:center}
.gauge-lbl{font-size:calc(11px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:3px}
.cs{width:calc(36px*var(--zf,1));height:calc(46px*var(--zf,1));border:2px solid #cbd5e1;border-radius:7px;display:flex;align-items:center;justify-content:center;font-size:calc(22px*var(--zf,1));font-weight:900;font-family:monospace;color:#c8d5e0;background:#f8fafc;transition:all .12s;user-select:none;cursor:default}
.cs.cs-filled{color:#2563eb;border-color:#3b82f6;background:#eff6ff}
.cs.cs-active{border-color:#4f46e5;box-shadow:0 0 0 3px rgba(99,102,241,.2)}
.cs.cs-done{border-color:#16a34a;background:#f0fdf4}
/* Active stops bottom bar — chips */
#stop-bottom,#stop-bottom-main{display:none;background:#7f0000;color:#fff;padding:8px 14px;align-items:center;gap:8px;flex-shrink:0;border-top:2px solid #b91c1c;flex-wrap:wrap}
#stop-bottom.on,#stop-bottom-main.on{display:flex}
.stop-chip{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.28);border-radius:8px;padding:6px 10px;border:1px solid rgba(255,255,255,.2)}
.chip-lbl{font-weight:800;font-size:calc(13px*var(--zf,1));white-space:nowrap}
.chip-tim{font-size:calc(18px*var(--zf,1));font-weight:800;font-variant-numeric:tabular-nums;min-width:52px;text-align:right}
.btn-endstop{background:#16a34a;color:#fff;border:none;border-radius:6px;padding:6px 12px;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer;white-space:nowrap}
.btn-endstop:hover{filter:brightness(.9)}
/* KPI shift cards */
.shift-kpis{display:flex;gap:8px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0}
.skpi{flex:1;background:var(--bg);border-radius:var(--radius);padding:10px;border:1px solid var(--border);text-align:center}
.skpi.current{flex:2;background:var(--navy);color:#fff;border-color:var(--navy2);box-shadow:var(--shadow)}
.skpi.current .sk-lbl{color:rgba(255,255,255,.7)}.skpi.current .sk-val{color:#93c5fd;font-size:calc(26px*var(--zf,1))}.skpi.current .sk-sub{color:rgba(255,255,255,.75)}
.sk-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);letter-spacing:.7px;margin-bottom:3px}
.sk-val{font-size:calc(20px*var(--zf,1));font-weight:800;color:var(--navy);line-height:1}.sk-sub{font-size:calc(10px*var(--zf,1));color:var(--gray);margin-top:3px}
/* Merged table row types */
.row-prod td{background:#f0fdf4}.row-evt td{background:#fff7ed}
.row-prod:hover td,.row-evt:hover td{filter:brightness(.96)}
.row-tag{display:inline-block;padding:1px 6px;border-radius:4px;font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase}
.tag-p{background:#bbf7d0;color:#166534}.tag-e{background:#fed7aa;color:#9a3412}.tag-n{background:#bfdbfe;color:#1e40af}

/* ── FIN DE POSTE ── */
#v-finposte{padding:0}
.fp-scroll{flex:1;overflow-y:auto;padding:14px}
.fp-top{text-align:center;padding-bottom:10px}
.fp-top h2{font-size:calc(22px*var(--zf,1));font-weight:800;color:var(--navy)}
.fp-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:14px}
.fp-card{background:var(--card);border-radius:var(--radius);padding:12px;text-align:center;box-shadow:var(--shadow);border:1px solid var(--border)}
.fp-big{font-size:calc(28px*var(--zf,1));font-weight:800;color:var(--navy)}
.fp-lbl{font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-top:3px}
.fp-acts{display:flex;justify-content:center;gap:10px;padding:14px 0}
.fp-tbl{width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1));margin-bottom:12px}
.fp-tbl th{background:var(--navy);color:#fff;padding:5px 8px;font-size:calc(9px*var(--zf,1));text-align:left;text-transform:uppercase}
.fp-tbl td{padding:5px 8px;border-bottom:1px solid var(--border)}

/* ── HISTORY ── */
#v-history{padding:0}

/* ── SETTINGS ── */
#v-settings{padding:0;overflow:hidden}
#settings-lock{flex:1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px}
.lock-card{background:var(--card);border-radius:12px;padding:28px;width:100%;max-width:340px;text-align:center;box-shadow:var(--shadow)}
.lock-card h3{color:var(--navy);font-size:calc(18px*var(--zf,1));font-weight:800;margin-bottom:8px}
.lock-card p{color:var(--gray);font-size:calc(12px*var(--zf,1));margin-bottom:16px}
#v-settings-content{flex:1;overflow-y:auto;padding:14px;display:none}
.ss{background:var(--card);border-radius:var(--radius);padding:14px;margin-bottom:12px;box-shadow:var(--shadow);border:1px solid var(--border)}
.ss h3{font-size:calc(12px*var(--zf,1));font-weight:700;margin-bottom:10px;color:var(--navy)}
.pr{display:flex;align-items:center;gap:6px;margin-bottom:6px;padding:5px;border-radius:5px;background:var(--bg)}
.pr .pn{font-weight:600;min-width:110px;font-size:calc(11px*var(--zf,1))}
.pr input{flex:1;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1))}
.btn-eye{background:none;border:none;cursor:pointer;color:var(--gray);font-size:calc(12px*var(--zf,1));padding:2px}
.day-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:6px}
.day-box{background:var(--bg);border-radius:4px;padding:4px;text-align:center}
.day-lbl{font-size:calc(8px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px}
.day-box input{width:100%;padding:2px;border:1px solid var(--border);border-radius:3px;font-size:calc(10px*var(--zf,1));text-align:center}
.model-card{border:1px solid var(--border);border-radius:7px;padding:10px;margin-bottom:8px}
.mch{display:flex;align-items:center;gap:6px;margin-bottom:6px}
.mch input{flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;padding:4px 6px;border:1px solid var(--border);border-radius:4px}

/* ── MODALS ── */
.modal{display:none}
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:300;align-items:center;justify-content:center}
.overlay.on{display:flex}
.mbox{background:var(--card);border-radius:12px;width:90%;max-width:520px;box-shadow:0 20px 60px rgba(0,0,0,.3);max-height:92vh;display:flex;flex-direction:column}
.mbox.wide{max-width:1100px}
.mhdr{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.mhdr.red{background:#b91c1c;border-radius:12px 12px 0 0}
.mhdr h2{font-size:calc(15px*var(--zf,1));font-weight:700;color:var(--navy)}
.mhdr.red h2,.mhdr.red button{color:#fff}
.mbody{flex:1;overflow-y:auto;padding:14px}
.mftr{display:flex;gap:8px;justify-content:flex-end;padding:12px 14px;border-top:1px solid var(--border);flex-shrink:0}
.btn{padding:7px 14px;border:none;border-radius:6px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer;transition:all .15s;display:inline-flex;align-items:center;gap:4px}
.btn:hover{filter:brightness(.9)}
.btn-prim{background:var(--navy);color:#fff}
.btn-danger{background:var(--red);color:#fff}
.btn-ok{background:var(--green);color:#fff}
.btn-sec{background:var(--lgray);color:var(--text)}
.btn-amber{background:var(--amber);color:#fff}
.btn-lg{font-size:calc(14px*var(--zf,1));padding:10px 20px}

/* Stop modal */
.stop-section-lbl{font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;letter-spacing:.8px;margin:10px 0 5px;color:var(--navy)}
.stop-section-lbl:first-child{margin-top:0}
.stops-grid{display:grid;gap:5px;margin-bottom:6px}
.stops-grid.ratt{grid-template-columns:repeat(3,1fr)}
.stops-grid.pb{grid-template-columns:repeat(4,1fr)}
.stop-btn{border:none;border-radius:8px;padding:8px 5px;cursor:pointer;font-size:calc(11px*var(--zf,1));font-weight:700;color:#fff;text-align:center;transition:all .12s;box-shadow:0 2px 0 rgba(0,0,0,.2)}
.stop-btn:active{transform:translateY(2px);box-shadow:none}
.stop-btn.ratt{background:#7c3aed}
.stop-btn.pb{background:#b91c1c}
.custom-row{display:flex;gap:6px;margin-top:4px}
.custom-row input{flex:1;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1));outline:none}
.custom-row input:focus{border-color:var(--navy)}

/* End prod modal */
.ep-grid{display:flex;flex-direction:column;gap:4px;margin-bottom:12px}
.ep-stat{display:flex;align-items:center;justify-content:space-between;padding:7px 12px;background:var(--bg);border-radius:6px;border-left:3px solid #c7d2fe}
.ep-stat .val{font-size:calc(16px*var(--zf,1));font-weight:800;color:var(--navy);white-space:nowrap}
.ep-stat .lbl{font-size:calc(10px*var(--zf,1));text-transform:uppercase;color:var(--gray);font-weight:700;margin-right:10px}
.ep-tbl{width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1))}
.ep-tbl th{text-align:left;padding:4px 6px;background:var(--bg);font-size:calc(9px*var(--zf,1));text-transform:uppercase;color:var(--gray)}
.ep-tbl td{padding:4px 6px;border-bottom:1px solid var(--border)}

/* Utils */
.hidden{display:none!important}
.flex{display:flex;align-items:center;gap:7px;flex-wrap:wrap}
.mt8{margin-top:8px}
.card{background:var(--card);border-radius:var(--radius);padding:12px;box-shadow:var(--shadow);border:1px solid var(--border)}

/* ── KPI VIEW ── */
#v-kpi{background:var(--bg)}
.kpi-prev-card{flex:1;background:var(--card);border:1px solid var(--border);border-radius:var(--radius);text-align:center;padding:6px 8px;min-width:0}

@media(max-width:900px){.form-3col{grid-template-columns:1fr 1fr}.recap-col{width:180px}}
@media(max-width:650px){.form-3col{grid-template-columns:1fr}.prod-body{flex-direction:column}.recap-col{width:100%}}
</style>
</head>
<body>

<!-- ════ LOGIN ════ -->
<div id="v-login" class="view on">
  <div class="login-card">
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;margin-bottom:16px">
      <div style="background:#fff;border-radius:11px;padding:3px 8px;box-shadow:0 2px 10px rgba(0,0,0,.35);display:inline-flex;align-items:center"><img src="data:image/png;base64,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" height="46" style="display:block;border-radius:8px" alt="DODO"></div>
      <div class="lc-h1" style="margin-bottom:0;white-space:nowrap;font-size:calc(22px*var(--zf,1));letter-spacing:.5px">PRODUCTION ORC1</div>
    </div>
    <div class="lf">
      <label>Pilote</label>
      <select id="ln-pilot"><option value="">-- Choisir --</option></select>
    </div>
    <div class="lf">
      <label>Modèle horaire (Poste)</label>
      <select id="ln-model" onchange="onLoginModelChange()"><option value="">-- Choisir --</option></select>
    </div>
    <!-- Horaires du jour -->
    <div id="ln-model-info" style="display:none;background:#f0f9ff;border:1px solid #bae6fd;border-radius:7px;padding:8px 10px;font-size:calc(12px*var(--zf,1));margin-bottom:6px">
      <div style="color:#0369a1;font-weight:700;margin-bottom:3px">Horaires aujourd'hui :</div>
      <div id="ln-model-times" style="font-size:calc(15px*var(--zf,1));font-weight:800;color:#0c4a6e;margin-bottom:5px"></div>
      <div id="ln-model-modify" style="display:none;margin-bottom:5px">
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <label style="font-size:calc(11px*var(--zf,1));font-weight:600">Début <input type="time" id="ln-new-debut" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:calc(12px*var(--zf,1))"></label>
          <label style="font-size:calc(11px*var(--zf,1));font-weight:600">Fin <input type="time" id="ln-new-fin" style="padding:3px 6px;border:1.5px solid #bae6fd;border-radius:5px;font-size:calc(12px*var(--zf,1))"></label>
        </div>
      </div>
      <button style="font-size:calc(11px*var(--zf,1));padding:3px 8px;background:none;border:1px solid #0369a1;border-radius:5px;color:#0369a1;cursor:pointer" onclick="toggleLoginModelModify()">✏ Modifier les horaires d'aujourd'hui</button>
    </div>
    <div class="lf">
      <label>Mot de passe</label>
      <input type="password" id="ln-pw" placeholder="••••" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="btn-login" onclick="doLogin()">Valider</button>
    <div class="ln-err" id="ln-err"></div>
    <div style="margin-top:12px;border-top:1px solid rgba(255,255,255,.15);padding-top:12px;text-align:center">
      <button onclick="doGuestLogin()" style="background:rgba(255,255,255,.12);color:#cbd5e1;border:1px solid rgba(255,255,255,.25);border-radius:7px;padding:7px 18px;font-size:calc(12px*var(--zf,1));cursor:pointer;width:100%;font-weight:600">👁 Consulter en tant qu'invité</button>
    </div>
    <div style="margin-top:10px;text-align:center;font-size:calc(11px*var(--zf,1));color:#94a3b8">
      Prod bloquée ? <a href="/reset" style="color:#dc2626;font-weight:700">Cliquer ici pour réinitialiser</a>
    </div>
  </div>
</div>

<!-- ════ APP ════ -->
<div id="app" class="hidden" style="display:none;flex:1;flex-direction:column;overflow:hidden">
  <div id="app-hdr">
    <div style="background:#fff;border-radius:12px;padding:4px 10px;box-shadow:0 3px 12px rgba(0,0,0,.2);display:inline-flex;align-items:center;align-self:center"><img src="data:image/png;base64,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" height="34" style="display:block;border-radius:6px" alt="DODO"></div>
    <div class="hdr-logo">PRODUCTION ORC</div>
    <div class="hdr-tabs">
      <button class="htab on" id="ht-main" onclick="goTab('main')">Accueil</button>
      <button class="htab prod-on" id="ht-prod" onclick="goTab('prod')">▶ Prod en cours</button>
      <span id="ht-guest-badge" style="display:none;font-size:calc(11px*var(--zf,1));font-weight:700;color:#94a3b8;padding:4px 10px;border:1px solid #94a3b8;border-radius:12px;margin-left:4px">👁 Invité</span>
      <button class="htab" id="ht-hist" onclick="goTab('history')">Historique</button>
      <button class="htab" id="ht-rapports" onclick="goTab('rapports')">📋 Rapports</button>
      <button class="htab" id="ht-kpi" onclick="goTab('kpi')">📊 KPI</button>
    </div>
    <div id="hdr-right">
      <span id="hdr-pilot-lbl"></span>
      <button class="btn-sm btn-ghost" onclick="toggleZoomPop()" id="zoom-btn" style="font-size:calc(11px*var(--zf,1));display:flex;align-items:center;gap:4px" title="Zoom texte">🔍 Zoom</button>
      <button class="btn-sm btn-ghost" onclick="doLogout()" style="font-size:calc(11px*var(--zf,1))">Déconnexion</button>
      <button class="btn-sm btn-ghost" id="ht-cfg" onclick="goTab('settings')" style="font-size:calc(18px*var(--zf,1));padding:4px 8px;line-height:1" title="Paramètres">⚙</button>
    </div>
  </div>
  <!-- Zoom popover -->
  <div id="zoom-pop" style="display:none;position:fixed;top:44px;right:80px;z-index:9000;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:12px 16px;box-shadow:0 4px 20px rgba(0,0,0,.4);min-width:220px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
      <span style="color:#fff;font-size:calc(12px*var(--zf,1));font-weight:700">🔍 Zoom texte</span>
      <button onclick="resetZoom()" style="background:none;border:1px solid #475569;border-radius:5px;color:#94a3b8;font-size:calc(10px*var(--zf,1));padding:2px 7px;cursor:pointer">Reset</button>
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      <button onclick="stepZoom(-5)" style="background:#334155;border:none;border-radius:5px;color:#fff;font-size:calc(14px*var(--zf,1));width:26px;height:26px;cursor:pointer;line-height:1">−</button>
      <input type="range" id="zoom-slider" min="70" max="150" value="135" step="5"
        oninput="applyZoom(+this.value)"
        style="flex:1;accent-color:#3b82f6;cursor:pointer">
      <button onclick="stepZoom(+5)" style="background:#334155;border:none;border-radius:5px;color:#fff;font-size:calc(14px*var(--zf,1));width:26px;height:26px;cursor:pointer;line-height:1">+</button>
    </div>
    <div style="text-align:center;margin-top:6px;color:#93c5fd;font-size:calc(13px*var(--zf,1));font-weight:700" id="zoom-val">135%</div>
  </div>
  <div id="alert-strip"></div>

  <!-- ════ MAIN VIEW ════ -->
  <div id="v-main" class="view" style="flex-direction:column">
    <!-- Bannière prod en cours (visible si prod_active mais sur vue accueil) -->
    <div id="main-prod-banner" style="display:none;background:#1e293b;color:#fff;padding:8px 14px;font-size:calc(12px*var(--zf,1));align-items:center;gap:16px;cursor:pointer" onclick="goTab('prod')">
      <span style="font-weight:800;color:#86efac">▶ Prod en cours</span>
      <span>OF : <span id="mpb-of" style="font-weight:700">—</span></span>
      <span>Durée : <span id="mpb-dur" style="color:#67e8f9;font-weight:700">—</span></span>
      <span>Arrêts : <span id="mpb-stops" style="color:#fca5a5;font-weight:700">—</span></span>
    </div>
    <!-- Barre Excel occupé -->
    <div id="excel-busy-bar" style="display:none;background:#92400e;color:#fef3c7;padding:5px 14px;font-size:calc(11px*var(--zf,1));font-weight:700;text-align:center">
      ⚠ Fichier Excel ouvert par un autre programme — impossible de lire/écrire les données
    </div>
    <div class="main-hdr">
      <div class="mbtns" style="margin-left:0" id="main-action-btns">
        <button class="btn btn-green" id="btn-start" onclick="doStartProd()" style="font-size:calc(15px*var(--zf,1));padding:18px 24px;font-weight:800;min-height:64px">▶ Démarrer production</button>
        <button class="btn btn-danger" onclick="openStopModal()" style="font-size:calc(15px*var(--zf,1));padding:18px 24px;font-weight:800;min-height:64px">⛔ Déclarer un arrêt</button>
        <button class="btn btn-amber" onclick="doFinPoste()" style="font-size:calc(15px*var(--zf,1));padding:18px 24px;font-weight:800;min-height:64px">🏁 Fin de poste</button>

      </div>
    </div>
    <!-- KPI accueil — POSTE ACTUEL -->
    <div style="background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;padding:4px 8px;display:flex;gap:6px;align-items:stretch;flex-wrap:wrap">

      <!-- POSTE ACTUEL encart principal -->
      <div style="flex:3;min-width:260px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:4px 10px;display:flex;flex-direction:column;gap:3px">
        <!-- Titre + TRS jauge + valeur -->
        <div style="display:flex;align-items:center;gap:8px">
          <div style="flex-shrink:0;text-align:center">
            <svg viewBox="0 0 100 58" style="width:100px;display:block;margin:0 auto">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="rgba(0,0,0,.12)" stroke-width="11" stroke-linecap="round"/>
              <path id="gauge-poste-acc-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="11" stroke-linecap="round" stroke-dasharray="0,132"/>
              <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#15803d" id="gauge-poste-acc-pct">—</text>
            </svg>
          </div>
          <div style="flex:1;min-width:0">
            <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#0369a1;text-transform:uppercase;letter-spacing:.5px">TRS du Poste</div>
            <div style="font-size:calc(13px*var(--zf,1));font-weight:700;color:#0369a1;margin-bottom:2px" id="gauge-poste-acc-lbl">—</div>
            <!-- Stats en colonne -->
            <div style="display:flex;flex-direction:column;gap:1px;margin-top:2px">
              <div style="display:flex;align-items:baseline;gap:5px">
                <span style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;white-space:nowrap">Nb OF</span>
                <span style="font-size:calc(17px*var(--zf,1));font-weight:900;color:#1e40af;line-height:1" id="acc-nb-of">0</span>
              </div>
              <div style="display:flex;align-items:baseline;gap:5px">
                <span style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#dc2626;white-space:nowrap">Arrêts</span>
                <span style="font-size:calc(15px*var(--zf,1));font-weight:900;color:#b91c1c;line-height:1" id="main-stat-arrets">0 min</span>
              </div>
              <div style="display:flex;align-items:baseline;gap:5px">
                <span style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;white-space:nowrap">Prod</span>
                <span style="font-size:calc(15px*var(--zf,1));font-weight:900;color:#15803d;line-height:1" id="main-stat-prod">0 min</span>
              </div>
            </div>
          </div>
          <!-- Répartition temps (pie) -->
          <div style="flex-shrink:0;text-align:center;border-left:1px solid #bae6fd;padding-left:8px;min-width:110px">
            <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px">Répartition temps</div>
            <svg id="pie-poste-acc" viewBox="0 0 130 130" style="width:160px;height:auto;display:block;margin:0 auto"></svg>
          </div>
        </div>
        <!-- Horaire temporaire -->
        <div id="acc-model-info" style="display:none;border-top:1px solid #bae6fd;padding-top:5px;margin-top:2px;font-size:calc(11px*var(--zf,1));color:#0369a1"></div>
      </div>

      <!-- Arrêts prévus — barres budget -->
      <div style="flex:2;min-width:190px;background:#fefce8;border:1px solid #fde68a;border-radius:10px;padding:4px 10px;display:flex;flex-direction:column">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">⏱ Arrêts prévus du poste en cours</div>
        <div id="budget-bars-acc" style="flex:1"></div>
      </div>

      <!-- Poste précédent -->
      <div class="skpi" style="flex:1;min-width:75px;max-width:120px">
        <div class="sk-lbl" id="kpi1-lbl">Poste précédent</div>
        <div class="sk-val" id="kpi1-trs">--%</div>
        <div class="sk-sub" id="kpi1-date" style="font-size:calc(10px*var(--zf,1));opacity:.85"></div>
        <div class="sk-sub" id="kpi1-sub">0 OF</div>
      </div>
      <div class="skpi" style="flex:1;min-width:75px;max-width:120px">
        <div class="sk-lbl" id="kpi2-lbl">Avant-dernier</div>
        <div class="sk-val" id="kpi2-trs">--%</div>
        <div class="sk-sub" id="kpi2-date" style="font-size:calc(10px*var(--zf,1));opacity:.85"></div>
        <div class="sk-sub" id="kpi2-sub">0 OF</div>
      </div>

      <!-- IDs cachés compatibles JS existant -->
      <div style="display:none">
        <span id="kpi0-trs"></span><span id="kpi0-sub"></span>
        <span id="kpi0-date"></span><span id="kpi0-heure"></span>
        <span id="main-stat-pause"></span><span id="main-stat-nett"></span>
        <span id="main-stat-reunion"></span>
      </div>
    </div>
    <div class="table-wrap">
      <table class="ktbl">
        <thead><tr>
          <th>Type</th><th>OF</th><th>Fibre</th><th>Date</th><th>Poste</th><th>Pilote</th>
          <th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Commentaire</th><th>Actions</th>
        </tr></thead>
        <tbody id="main-body"></tbody>
      </table>
    </div>
    <!-- Active stops bottom bar — chips (accueil) -->
    <div id="stop-bottom-main">
      <div style="font-size:calc(10px*var(--zf,1));opacity:.7;font-weight:700;white-space:nowrap">EN COURS :</div>
      <div id="stop-chips-main" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1"></div>
    </div>
  </div>

  <!-- ════ PRODUCTION VIEW ════ -->
  <div id="v-prod" class="view" style="flex-direction:column">
    <!-- Pilot/OF banner -->
    <div class="pob">
      <div class="pob-item">
        <div class="pob-lbl">Pilote</div>
        <div class="pob-val" id="pob-pilot">—</div>
      </div>
      <div class="pob-item of">
        <div class="pob-lbl">N° OF</div>
        <div class="pob-val" id="pob-of">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Poste</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1))" id="pob-poste">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Départ OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#fbbf24" id="pob-of-start">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Fin OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#94a3b8" id="rc-fin">—</div>
      </div>
      <div class="pob-item">
        <div class="pob-lbl">Durée OF</div>
        <div class="pob-val" style="font-size:calc(16px*var(--zf,1));color:#0891b2" id="rc-duree">—</div>
      </div>
      <div class="pob-item trs">
        <div class="pob-lbl">TRS estimé</div>
        <div class="pob-val" id="pob-trs">—</div>
      </div>
    </div>
    <!-- Status bar -->
    <div class="sbar">
      <div class="sc"><div class="sc-lbl">⏱ Durée OF</div><div class="sc-val green" id="sc-of">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⛔ Arrêts</div><div class="sc-val red" id="sc-stops">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">⏸ Pauses</div><div class="sc-val amber" id="sc-pause">00:00:00</div></div>
      <div class="sc"><div class="sc-lbl">🎯 Pièces théo.</div><div class="sc-val" id="sc-theo" style="color:var(--blue)">—</div></div>
    </div>
    <!-- Body -->
    <div class="prod-body">
      <!-- LEFT: form (now full center, no act-col) -->
      <div class="form-col">
        <div class="form-3col">
          <!-- Zone Identification -->
          <div class="fzone zi">
            <h4>📋 Identification</h4>
            <div class="fr"><label>N° OF *</label><input id="f-of_num" oninput="scheduleAutoSave()" onfocus="openCodeInput('of_num','N° OF','9')"></div>
            <div class="fr ro"><label>Date</label><input id="f-date" readonly></div>
            <div class="fr ro"><label>Poste</label><input id="f-poste" readonly></div>
            <div class="fr ro"><label>Pilote</label><input id="f-pilote" readonly></div>
            <div class="fr"><label>Co-Pilote</label><select id="f-copilote" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Nb Personnes</label><input id="f-nb_pers" type="number" min="1" value="10" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Code Produit</label><input id="f-code_prod" oninput="scheduleAutoSave()" onfocus="openCodeInput('code_prod','Code Produit')"></div>
            <div class="fr"><label>Type Produit</label><select id="f-type_prod" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
          </div>
          <!-- Zone Production -->
          <div class="fzone zp">
            <h4>🏭 Production</h4>
            <div class="fr big"><label>Qté Fabriquée *</label><input id="f-qte_fab" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr big"><label>Qté Emballée</label><input id="f-qte_emb" type="number" min="0" placeholder="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Poids Garnissage (g)</label><input id="f-poids" type="number" min="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Taille</label><select id="f-taille" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Fibre</label><select id="f-fibre" onchange="scheduleAutoSave()"><option value="">--</option></select></div>
            <div class="fr"><label>Traca Fibre</label><input type="text" id="f-traca" oninput="scheduleAutoSave()" placeholder="n° de traca"></div>
            <div class="fr"><label>Code Taie</label><input id="f-ref_taie" oninput="scheduleAutoSave()" onfocus="openCodeInput('ref_taie','Code Taie')"></div>
            <div class="fr"><label>Lots de 2</label><select id="f-kit" onchange="scheduleAutoSave()"><option value="">Non</option><option value="oui">Oui</option></select></div>
            <div class="fr" style="display:none"><input id="f-of_taie" oninput="scheduleAutoSave()"></div>
            <div class="fr" style="display:none"><input id="f-duree_mq_mp" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>MQ PERSONNEL (Seulement si arrêt d'une partie de la ligne) (min)</label><input id="f-manquant_pers" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
          </div>
          <!-- Zone Qualité -->
          <div class="fzone zq">
            <h4>✅ Qualité</h4>
            <div class="fr" style="display:none"><input id="f-qte_init_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Taie 2nd Choix</label><input id="f-nb_taie2_choix" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb Défaut Couture</label><input id="f-nb_def_cout" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Taie</label><input id="f-mq_taie" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Mq Housse/Encart</label><input id="f-mq_housse_encart" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr"><label>Nb PP Cousue</label><input id="f-nb_pp_cousue" type="number" min="0" value="0" oninput="scheduleAutoSave()"></div>
            <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="f-comment" oninput="scheduleAutoSave()" placeholder="Commentaire libre…"></textarea></div>
          </div>
        </div>
        <!-- Timeline 4h -->
        <div class="tl-wrap">
          <h5>Timeline — 4 dernières heures</h5>
          <svg id="tl-svg" viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">
            <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
          </svg>
          <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
        </div>
        <!-- Action buttons row (below timeline) -->
        <div class="prod-act-row">
          <button class="act-btn act-stop" onclick="openStopModal()">⛔ Déclarer un arrêt</button>
          <button class="act-btn act-nett" onclick="doNettoyage()">🧹 Nettoyage</button>
          <button class="act-btn act-pause" id="btn-pause" onclick="doPause()">⏸ Pause</button>
          <button class="act-btn" id="btn-reunion" onclick="doReunion()" style="background:var(--card);border:1.5px solid #8b5cf6;color:#7c3aed;font-weight:700;cursor:pointer">👥 Réunion</button>
          <button class="act-btn act-cancel" onclick="doCancelProd()">✖ Annuler prod</button>
          <button class="act-btn act-endprod" onclick="doEndProdPreview()">🏁 Fin d'OF/prod</button>
        </div>
      </div>
      <!-- RIGHT: recap arrêts + gauges + pie charts -->
      <div class="recap-col" style="width:310px">
        <div class="recap-hdr">Arrêts / pauses</div>
        <div class="recap-body" id="recap-list" style="max-height:120px;flex:none;overflow-y:auto"></div>
        <!-- Budget arrêts prévus -->
        <div style="padding:5px 8px;border-top:1px solid var(--border);flex-shrink:0;background:#fffbeb">
          <div style="font-size:calc(9px*var(--zf,1));font-weight:700;color:#92400e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px">⏱ Arrêts prévus</div>
          <div id="budget-bars-prod"></div>
        </div>
        <!-- TRS OF gauge -->
        <div class="gauge-box" style="padding:8px 4px 4px;border-top:1px solid var(--border);flex-shrink:0">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;align-items:center">
            <div style="text-align:center">
              <svg viewBox="0 0 100 56" style="width:100%;max-width:140px">
                <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
                <path id="gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
                <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="gauge-pct">—</text>
              </svg>
              <div class="gauge-lbl">TRS OF</div>
            </div>
            <div style="text-align:center">
              <svg viewBox="0 0 100 56" style="width:100%;max-width:140px">
                <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="10" stroke-linecap="round"/>
                <path id="gauge-poste-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#0891b2" stroke-width="10" stroke-linecap="round" stroke-dasharray="0,1000"/>
                <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#0c4a6e" id="gauge-poste-pct">—</text>
              </svg>
              <div class="gauge-lbl" id="gauge-poste-lbl">TRS Poste</div>
            </div>
          </div>
        </div>
        <!-- Pie charts -->
        <div style="padding:4px;border-top:1px solid var(--border);display:grid;grid-template-columns:1fr 1fr;gap:4px">
          <div style="text-align:center">
            <div style="font-size:calc(8px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">Poste — Prod/Arrêts</div>
            <svg id="pie-poste" viewBox="0 0 130 115" style="width:100%;height:auto;display:block"></svg>
          </div>
          <div style="text-align:center">
            <div style="font-size:calc(8px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:2px">OF en cours</div>
            <svg id="pie-of" viewBox="0 0 130 115" style="width:100%;height:auto;display:block"></svg>
          </div>
        </div>
      </div>
    </div>
    <!-- Active stops bottom bar — chips -->
    <div id="stop-bottom">
      <div style="font-size:calc(10px*var(--zf,1));opacity:.7;font-weight:700;white-space:nowrap">EN COURS :</div>
      <div id="stop-chips" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1"></div>
    </div>
  </div>

  <!-- ════ FIN DE POSTE ════ -->
  <div id="v-finposte" class="view" style="flex-direction:column;overflow:hidden">
    <!-- Barre titre -->
    <div style="background:var(--navy);color:#fff;padding:7px 14px;flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
      <div>
        <div style="font-size:calc(15px*var(--zf,1));font-weight:800">🏁 Fin de poste</div>
        <div id="fp-who" style="font-size:calc(11px*var(--zf,1));opacity:.8"></div>
      </div>
      <div style="text-align:right">
        <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:#fbbf24" id="fp-horaire-display">—</div>
        <div id="fp-date" style="font-size:calc(11px*var(--zf,1));font-weight:700;opacity:.8"></div>
      </div>
    </div>
    <!-- Graphiques + KPI (en haut, compact) -->
    <div style="display:flex;gap:12px;padding:10px 14px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0;align-items:center;flex-wrap:wrap">
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">TRS Poste</div>
        <svg id="fp-gauge" viewBox="0 0 100 58" style="width:200px;display:block;margin:0 auto">
          <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
          <path id="fp-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
          <text x="50" y="46" text-anchor="middle" font-size="14" font-weight="800" fill="#1a1f5e" id="fp-gauge-pct">--%</text>
        </svg>
        <div style="font-size:calc(12px*var(--zf,1));color:var(--gray);margin-top:2px" id="fp-shift-hours">—</div>
      </div>
      <div style="text-align:center;flex-shrink:0">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Répartition</div>
        <svg id="fp-pie" viewBox="0 0 130 130" style="width:200px;height:200px;display:block;margin:0 auto"></svg>
      </div>
      <div style="flex:1;display:flex;flex-direction:column;gap:5px">
        <div style="display:none"><span id="fp-trs"></span><span id="fp-trs-of"></span></div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:5px">
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#059669;font-weight:900" id="fp-pieces">--</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Nb pièces prod.</div></div>
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0891b2;font-weight:900" id="fp-eq">0</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Équivalence</div></div>
          <div class="fp-card" style="padding:9px"><div class="fp-big" style="font-size:calc(26px*var(--zf,1));color:#0369a1;font-weight:900" id="fp-cadence">--</div><div class="fp-lbl" style="font-size:calc(12px*var(--zf,1))">Cadence/h</div></div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(6,1fr);gap:5px">
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1))" id="fp-ouverture">--</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée ouverture</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#16a34a" id="fp-prod-t">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Durée prod</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#dc2626" id="fp-stop-t">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts total</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#16a34a" id="fp-ded">0 min</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Arrêts prévus</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#7c3aed" id="fp-nof">0</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Nb OF</div></div>
          <div class="fp-card" style="padding:7px"><div class="fp-big" style="font-size:calc(14px*var(--zf,1));color:#8b5cf6" id="fp-chg-fibre">--</div><div class="fp-lbl" style="font-size:calc(11px*var(--zf,1))">Chg. fibre</div></div>
        </div>
      </div>
    </div>
    <!-- Timeline compact -->
    <div style="padding:5px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:3px">Timeline du poste</div>
      <svg id="fp-tl" viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">
        <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
      </svg>
      <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
    </div>
    <!-- Corps défilant : productions + arrêts côte à côte -->
    <div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div class="card" style="padding:8px">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Productions</div>
        <table class="fp-tbl" style="font-size:calc(10px*var(--zf,1))">
          <thead><tr><th>OF</th><th>Début</th><th>Fin</th><th>Taille</th><th>Qté</th><th>Éq</th><th>Durée</th><th>TRS%</th></tr></thead>
          <tbody id="fp-prods"></tbody>
        </table>
      </div>
      <div class="card" style="padding:8px">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:5px">Arrêts du poste</div>
        <div id="fp-stops-list" style="font-size:calc(11px*var(--zf,1))"></div>
      </div>
    </div>
    <!-- Boutons -->
    <div style="padding:8px 12px;background:var(--card);border-top:1px solid var(--border);flex-shrink:0;display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-sec" onclick="goTab('main')">← Retour</button>
      <button onclick="confirmFinPoste()" style="background:#16a34a;color:#fff;border:none;border-radius:10px;font-size:calc(18px*var(--zf,1));font-weight:900;padding:18px 40px;cursor:pointer;box-shadow:0 4px 18px rgba(22,163,74,.4);letter-spacing:.3px;transition:all .15s" onmouseover="this.style.filter='brightness(.9)'" onmouseout="this.style.filter=''">✅ Confirmer fin de poste &amp; Déconnexion</button>
    </div>
  </div>

  <!-- ════ MODAL : horaires de poste ════ -->
  <div id="m-fp-horaires" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;align-items:center;justify-content:center">
    <div class="card" style="width:340px;padding:20px;background:#fff;border-radius:12px">
      <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:14px">✏ Horaires de mon poste</div>
      <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Ces horaires servent uniquement au calcul du TRS de poste (non sauvegardés).</div>
      <div class="lf" style="margin-bottom:10px">
        <label>Début de poste</label>
        <input type="datetime-local" id="fp-debut-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div class="lf" style="margin-bottom:14px">
        <label>Fin de poste</label>
        <input type="datetime-local" id="fp-fin-dt" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="closeM('m-fp-horaires')">Annuler</button>
        <button class="btn btn-prim" onclick="applyFPHoraires()">✓ Appliquer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL PRÉ-POSTE (gaps non déclarés avant 1er OF ou entre OFs) ════ -->
  <div id="m-preshift" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:601;align-items:center;justify-content:center">
    <div class="card" style="width:min(880px,95vw);padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--red)">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
        <div id="ps-title" style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy)">⚠ Période non déclarée</div>
        <div id="ps-counter" style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#94a3b8;background:#f1f5f9;border-radius:8px;padding:3px 10px"></div>
      </div>
      <div id="ps-text" style="font-size:calc(13px*var(--zf,1));color:var(--red);font-weight:700;margin-bottom:14px"></div>
      <input type="hidden" id="ps-start-iso">
      <input type="hidden" id="ps-gap-s">
      <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:var(--gray);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Déclarer comme :</div>
      <div id="ps-stop-btns" style="margin-bottom:10px"></div>
      <div style="margin-bottom:10px;display:flex;gap:6px">
        <input id="ps-custom" placeholder="Ou saisir librement…" style="flex:1;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))" onkeydown="if(event.key==='Enter')confirmPsAsStop()">
        <button class="btn btn-prim" onclick="confirmPsAsStop()" style="flex-shrink:0">✓ Valider</button>
      </div>
      <hr style="border:none;border-top:1px solid var(--border);margin-bottom:14px">
      <div style="display:flex;gap:8px;align-items:center;justify-content:space-between">
        <div id="ps-backdate-row" style="flex:1">
          <button class="btn btn-green" style="width:100%;text-align:left;padding:10px 14px;font-size:calc(13px*var(--zf,1))" onclick="psChooseBackdate()">
            ↩ Rétrodater le début de cet OF à <span id="ps-backdate-time" style="font-weight:800">--h--</span>
          </button>
        </div>
        <button class="btn btn-sec" style="flex-shrink:0;padding:10px 18px;font-size:calc(13px*var(--zf,1))" onclick="psIgnorer()">Ignorer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL INTERPOSTE ════ -->
  <div id="m-interposte" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:min(880px,95vw);padding:20px;background:#fff;border-radius:12px;border-top:4px solid var(--amber)">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:4px">⏱ Temps hors production</div>
      <div id="ip-duration" style="font-size:calc(13px*var(--zf,1));color:var(--amber);font-weight:700;margin-bottom:8px"></div>
      <div style="display:flex;gap:10px;margin-bottom:12px;align-items:flex-end">
        <div style="flex:1">
          <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Début</label>
          <input type="time" id="ip-debut" style="width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
        </div>
        <div style="flex:1">
          <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Fin</label>
          <input type="time" id="ip-fin" style="width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
        </div>
      </div>
      <div style="font-size:calc(12px*var(--zf,1));color:var(--gray);margin-bottom:10px">Que s'est-il passé pendant cette période ?</div>
      <div id="ip-arrprev-group" style="display:none;margin-bottom:10px"></div>
      <div id="ip-btns" style="display:flex;flex-wrap:wrap;gap:7px;margin-bottom:12px"></div>
      <div style="margin-bottom:10px">
        <input id="ip-custom" placeholder="Ou saisir librement…" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))">
      </div>
      <div style="margin-bottom:10px">
        <input id="ip-comment" placeholder="Commentaire (optionnel)" style="width:100%;padding:7px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(12px*var(--zf,1))">
      </div>
      <!-- Formulaire modification plage horaire (caché par défaut) -->
      <div id="ip-model-form" style="display:none;background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:8px;padding:10px;margin-bottom:10px">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#0369a1;margin-bottom:8px;text-transform:uppercase">Modifier la plage horaire du poste</div>
        <div style="display:flex;gap:10px;margin-bottom:8px">
          <div style="flex:1">
            <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Début modèle</label>
            <input type="time" id="ip-model-debut" style="width:100%;padding:6px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
          </div>
          <div style="flex:1">
            <label style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);display:block;margin-bottom:3px">Fin modèle</label>
            <input type="time" id="ip-model-fin" style="width:100%;padding:6px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700">
          </div>
        </div>
        <div style="display:flex;gap:6px;justify-content:flex-end">
          <button class="btn btn-sec" style="font-size:calc(11px*var(--zf,1));padding:4px 10px" onclick="document.getElementById('ip-model-form').style.display='none'">Annuler</button>
          <button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:4px 10px" onclick="ipConfirmModifyModel()">✓ Appliquer</button>
        </div>
      </div>
      <div style="display:flex;gap:8px;justify-content:space-between;align-items:center">
        <button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:4px 10px;color:#0369a1;border-color:#bae6fd" onclick="ipShowModifyModel()">✏ Modifier plage horaire</button>
        <div style="display:flex;gap:8px">
          <button class="btn btn-sec" onclick="skipInterposte()">Ignorer</button>
          <button class="btn btn-prim" onclick="confirmInterposte()">✓ Valider</button>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ MODAL ARRÊTS MANQUANTS ════ -->
  <div id="m-missing-decl" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:480px;max-height:85vh;overflow-y:auto;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #f59e0b">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:#92400e;margin-bottom:4px">⚠ Arrêts non déclarés</div>
      <div style="font-size:calc(12px*var(--zf,1));color:#78350f;margin-bottom:14px">Vous n'avez pas déclaré les arrêts prévus suivants. Souhaitez-vous les ajouter avant de terminer le poste ?</div>
      <div id="md-rows" style="display:flex;flex-direction:column;gap:10px;margin-bottom:14px"></div>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="skipMissingDecl()">Ignorer et continuer</button>
        <button class="btn btn-prim" style="background:#f59e0b;border-color:#f59e0b" onclick="skipMissingDecl()">✓ Continuer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL ÉCART FIN DE POSTE ════ -->
  <div id="m-ecart-poste" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:600;align-items:center;justify-content:center">
    <div class="card" style="width:min(1160px,95vw);max-height:92vh;overflow-y:auto;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #dc2626">
      <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:6px">📊 Réconciliation fin de poste</div>
      <div id="ecart-guide" style="font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;line-height:1.5"></div>
      <div id="ecart-info" style="background:#f8fafc;border:1px solid var(--border);border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:calc(12px*var(--zf,1))"></div>
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);letter-spacing:.06em;margin-bottom:6px">Plages non justifiées</div>
      <div id="ecart-gaps" style="margin-bottom:14px"></div>
      <div id="ecart-of-panel" style="margin-bottom:10px">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);letter-spacing:.06em;margin-bottom:6px">Modifier les horaires des OFs</div>
        <div id="ecart-of-list" style="display:flex;flex-direction:column;gap:6px"></div>
      </div>
      <!-- Modifier la plage horaire du poste -->
      <div id="ecart-plage-panel" style="margin-bottom:14px">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#0369a1;letter-spacing:.06em;margin-bottom:6px">Modifier la plage horaire de mon poste</div>
        <div style="background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:8px;padding:10px;display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap">
          <div>
            <label style="font-size:calc(10px*var(--zf,1));color:var(--gray);font-weight:600;display:block;margin-bottom:3px">Heure de début</label>
            <input type="time" id="ecart-plage-debut" style="padding:5px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700;width:110px">
          </div>
          <div>
            <label style="font-size:calc(10px*var(--zf,1));color:var(--gray);font-weight:600;display:block;margin-bottom:3px">Heure de fin</label>
            <input type="time" id="ecart-plage-fin" style="padding:5px 8px;border:1.5px solid #bae6fd;border-radius:6px;font-size:calc(13px*var(--zf,1));font-weight:700;width:110px">
          </div>
          <button class="btn btn-prim" style="font-size:calc(12px*var(--zf,1));padding:6px 16px;background:#0369a1;border-color:#0369a1" onclick="saveEcartPlageHoraire()">✓ Appliquer</button>
          <span style="font-size:calc(10px*var(--zf,1));color:#64748b;align-self:center">Modifie uniquement pour ce poste aujourd'hui</span>
        </div>
      </div>
      <div style="display:flex;gap:8px;justify-content:flex-end;align-items:center">
        <button class="btn btn-ghost" style="font-size:calc(12px*var(--zf,1))" onclick="closeM('m-ecart-poste')">Annuler</button>
        <button id="ecart-valider-btn" class="btn btn-sec" onclick="skipEcartPoste()">Valider et terminer</button>
      </div>
    </div>
  </div>

  <!-- ════ MODAL : "Autre" type d'arrêt ════ -->
  <div id="m-ecart-autre" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:700;align-items:center;justify-content:center">
    <div class="card" style="width:340px;padding:20px;background:#fff;border-radius:12px;border-top:4px solid #64748b">
      <div style="font-size:calc(14px*var(--zf,1));font-weight:800;color:var(--navy);margin-bottom:10px">✏ Saisir le type d'arrêt</div>
      <input id="ecart-autre-label" placeholder="Ex : Changement outillage, Défaut matière…" style="width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(13px*var(--zf,1));outline:none;margin-bottom:14px" onkeydown="if(event.key==='Enter')confirmEcartAutre()">
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-sec" onclick="closeM('m-ecart-autre')">Annuler</button>
        <button class="btn btn-prim" onclick="confirmEcartAutre()">✓ Valider</button>
      </div>
    </div>
  </div>

<!-- Modal saisie code formaté (DDDDDD_DDD) -->
<div id="m-code-input" class="modal" style="position:fixed;inset:0;z-index:9999;backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);background:rgba(0,0,0,0.75);align-items:center;justify-content:center" onclick="if(event.target===this)closeM('m-code-input')">
  <div style="background:#fff;border-radius:14px;padding:24px 28px;width:90%;max-width:480px;text-align:center;box-shadow:0 30px 80px rgba(0,0,0,.5)">
    <div style="font-size:calc(13px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:8px;letter-spacing:.05em" id="code-input-lbl">CODE</div>
    <div id="code-slots" tabindex="0" onkeydown="_csKeydown(event)"
      style="display:flex;align-items:center;justify-content:center;gap:5px;margin-bottom:16px;outline:none;cursor:text">
      <div class="cs" id="cs-0">X</div><div class="cs" id="cs-1">X</div><div class="cs" id="cs-2">X</div>
      <div class="cs" id="cs-3">X</div><div class="cs" id="cs-4">X</div><div class="cs" id="cs-5">X</div>
      <div id="cs-sep" style="font-size:calc(26px*var(--zf,1));font-weight:900;color:#94a3b8;line-height:1;align-self:center;margin:0 3px">_</div>
      <div class="cs" id="cs-6">X</div><div class="cs" id="cs-7">X</div><div class="cs" id="cs-8">X</div>
    </div>
    <div id="cs-hint" style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:16px">Tapez les 6 premiers puis les 3 derniers chiffres</div>
    <div style="display:flex;gap:8px;justify-content:center">
      <button class="btn btn-ghost" onclick="closeM('m-code-input')">Annuler</button>
      <button class="btn btn-primary" onclick="_codeInputConfirm()">Confirmer ✓</button>
    </div>
  </div>
</div>

<!-- Modal détail OF -->
<div id="m-of-detail" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:650;align-items:center;justify-content:center" onclick="if(event.target===this)closeM('m-of-detail')">
  <div style="width:min(860px,96vw);max-height:88vh;background:#fff;border-radius:14px;box-shadow:0 24px 80px rgba(0,0,0,.45);display:flex;flex-direction:column;overflow:hidden">
    <div id="of-detail-content" style="display:flex;flex-direction:column;min-height:0;flex:1;overflow:hidden"></div>
  </div>
</div>

  <!-- ════ RAPPORTS DES POSTES ════ -->
  <div id="v-rapports" class="view" style="flex-direction:column;overflow:hidden">
    <div style="display:flex;flex:1;overflow:hidden;min-height:0">
      <!-- Panneau gauche : liste OU résumé KPI selon mode -->
      <div id="rpt-left-exe" style="width:280px;min-width:0;flex-shrink:0;border-right:1px solid var(--border);display:flex;flex-direction:column;background:#f8fafc;overflow:hidden;transition:width .25s ease">
        <!-- Mode liste (par défaut) -->
        <div id="rpt-left-list" style="display:flex;flex-direction:column;flex:1;overflow:hidden">
          <div style="padding:10px 14px;font-size:calc(13px*var(--zf,1));font-weight:800;color:var(--navy);border-bottom:1px solid var(--border);flex-shrink:0;display:flex;align-items:center;justify-content:space-between">
            <span>📋 Tous les postes</span>
            <button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:3px 8px" onclick="reloadAndLoadRapports()">↺</button>
          </div>
          <div id="rpt-list" style="flex:1;overflow-y:auto">
            <div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Chargement…</div>
          </div>
        </div>
        <!-- Mode KPI (caché jusqu'à sélection d'un poste) -->
        <div id="rpt-left-kpi" style="display:none;flex-direction:column;flex:1;overflow-y:auto"></div>
      </div>
      <!-- Détail du rapport : tables + timeline -->
      <div id="rpt-detail" style="overflow-y:auto;flex:1;padding:0">
        <div style="padding:60px;text-align:center;color:var(--gray)">
          <div style="font-size:calc(40px*var(--zf,1));margin-bottom:12px">📋</div>
          <div style="font-size:calc(14px*var(--zf,1));font-weight:600">Sélectionner un poste dans la liste</div>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ HISTORY ════ -->
  <div id="v-history" class="view" style="flex-direction:column;overflow:hidden">
    <div style="background:var(--card);border-bottom:1px solid var(--border);padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap">
      <span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:var(--navy)">Historique</span>
      <label style="font-size:calc(11px*var(--zf,1));font-weight:600;color:var(--gray)">Du <input type="date" id="hist-from" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));margin-left:4px"></label>
      <label style="font-size:calc(11px*var(--zf,1));font-weight:600;color:var(--gray)">Au <input type="date" id="hist-to" style="padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));margin-left:4px"></label>
      <button class="btn btn-primary" onclick="loadHist()" style="padding:5px 12px;font-size:calc(12px*var(--zf,1))">Charger</button>
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <input id="hist-search" type="text" placeholder="🔍 Rechercher N° OF…" style="padding:4px 9px;border:1.5px solid var(--border);border-radius:7px;font-size:calc(12px*var(--zf,1));min-width:160px" oninput="filterHistBySearch()">
      <div style="width:1px;height:22px;background:var(--border);flex-shrink:0"></div>
      <div style="display:flex;gap:5px;flex-wrap:wrap;align-items:center">
        <span style="font-size:calc(10px*var(--zf,1));font-weight:700;color:var(--gray);text-transform:uppercase">Filtrer :</span>
        <button class="hf-btn" data-hf="production" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #16a34a;color:#16a34a;background:none;cursor:pointer;font-weight:700;transition:all .15s">🏭 Production</button>
        <button class="hf-btn" data-hf="arret" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #dc2626;color:#dc2626;background:none;cursor:pointer;font-weight:700;transition:all .15s">⛔ Arrêts</button>
        <button class="hf-btn" data-hf="nettoyage" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #0891b2;color:#0891b2;background:none;cursor:pointer;font-weight:700;transition:all .15s">🧹 Nettoyage</button>
        <button class="hf-btn" data-hf="pause" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #f59e0b;color:#f59e0b;background:none;cursor:pointer;font-weight:700;transition:all .15s">⏸ Pause</button>
        <button class="hf-btn" data-hf="reunion" onclick="toggleHistFilter(this)" style="font-size:calc(11px*var(--zf,1));padding:3px 9px;border-radius:12px;border:1.5px solid #8b5cf6;color:#8b5cf6;background:none;cursor:pointer;font-weight:700;transition:all .15s">👥 Réunion</button>
      </div>
    </div>
    <div style="flex:1;overflow-y:auto">
      <table class="ktbl"><thead><tr id="hist-hd"></tr></thead><tbody id="hist-bd"></tbody></table>
    </div>
  </div>

  <!-- ════ KPI VIEW ════ -->
  <div id="v-kpi" class="view" style="flex-direction:column;overflow:hidden;background:#f1f5f9">
    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1e3a8a,#1e40af);color:#fff;padding:9px 16px;flex-shrink:0;display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div style="font-size:calc(18px*var(--zf,1));font-weight:900;letter-spacing:.4px;text-shadow:0 1px 4px rgba(0,0,0,.2)">📊 Tableau de Bord Performance</div>
      <div style="display:flex;align-items:center;gap:6px;margin-left:auto;flex-wrap:wrap">
        <label style="font-size:calc(12px*var(--zf,1));opacity:.85;font-weight:600">Du</label>
        <input type="date" id="kpi-from" style="padding:4px 8px;border:1px solid rgba(255,255,255,.35);border-radius:6px;font-size:calc(12px*var(--zf,1));background:rgba(255,255,255,.18);color:#fff;outline:none;font-weight:600">
        <label style="font-size:calc(12px*var(--zf,1));opacity:.85;font-weight:600">au</label>
        <input type="date" id="kpi-to" style="padding:4px 8px;border:1px solid rgba(255,255,255,.35);border-radius:6px;font-size:calc(12px*var(--zf,1));background:rgba(255,255,255,.18);color:#fff;outline:none;font-weight:600">
        <button onclick="loadKPI()" style="padding:5px 14px;background:rgba(255,255,255,.22);border:1px solid rgba(255,255,255,.45);border-radius:7px;color:#fff;font-size:calc(12px*var(--zf,1));font-weight:700;cursor:pointer">↺ Actualiser</button>
      </div>
    </div>
    <!-- KPI Cards row -->
    <div id="kpi-cards" style="display:flex;flex-shrink:0;border-bottom:2px solid #e2e8f0;background:#fff;overflow-x:auto;min-height:68px"></div>
    <!-- 3 courbes côte à côte (compact) -->
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;height:128px;flex-shrink:0;background:#fff;border-bottom:1px solid #e2e8f0">
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px;border-right:1px solid #f1f5f9">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">📈 TRS par poste (%)</div>
        <div id="kpi-trs-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px;border-right:1px solid #f1f5f9;background:#fafafa">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#dc2626;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">🛑 Arrêts cumulés (min)</div>
        <div id="kpi-arr-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
      <div style="display:flex;flex-direction:column;overflow:hidden;padding:6px 10px 4px">
        <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#6366f1;letter-spacing:.5px;margin-bottom:2px;flex-shrink:0">📦 Nombre d'OF par poste</div>
        <div id="kpi-of-chart" style="flex:1;min-height:0;overflow:hidden"></div>
      </div>
    </div>
    <!-- Corps principal -->
    <div style="flex:1;overflow:hidden;display:grid;grid-template-columns:1fr 210px;min-height:0">
      <!-- Gauche : cadence + 2 évolutions -->
      <div style="display:flex;flex-direction:column;overflow:hidden;border-right:1px solid #e2e8f0;min-height:0">
        <div style="flex:2;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#fff;border-bottom:1px solid #f1f5f9;min-height:0">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#f59e0b;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">⚡ Évolution cadence (éq./h)</div>
          <div id="kpi-cad-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
        <div style="flex:1;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#fff;border-bottom:1px solid #f1f5f9;min-height:0">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#8b5cf6;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">🧵 Changements de fibre par poste</div>
          <div id="kpi-fibre-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
        <div style="flex:1;padding:6px 10px;overflow:hidden;display:flex;flex-direction:column;background:#f8fafc;min-height:0">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#16a34a;letter-spacing:.5px;margin-bottom:3px;flex-shrink:0">📦 Pièces fab. / équivalence</div>
          <div id="kpi-qte-chart" style="flex:1;min-height:0;overflow:hidden"></div>
        </div>
      </div>
      <!-- Droite : donut + pareto -->
      <div style="display:flex;flex-direction:column;overflow:hidden;background:#fff;min-height:0">
        <div style="padding:7px 10px;border-bottom:1px solid #f1f5f9;flex-shrink:0;display:flex;flex-direction:column;align-items:center">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;letter-spacing:.5px;margin-bottom:3px;align-self:flex-start">Prod vs Arrêts</div>
          <div style="display:flex;align-items:center;gap:8px">
            <svg id="kpi-donut" viewBox="0 0 100 100" style="width:78px;height:78px;flex-shrink:0"></svg>
            <div id="kpi-donut-legend" style="display:flex;flex-direction:column;gap:3px;font-size:calc(10px*var(--zf,1))"></div>
          </div>
        </div>
        <div style="flex:1;overflow-y:auto;padding:8px 10px;display:flex;flex-direction:column;min-height:0">
          <div style="font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;letter-spacing:.5px;margin-bottom:6px;flex-shrink:0">Pareto arrêts</div>
          <div id="kpi-pareto-new" style="display:flex;flex-direction:column;gap:5px"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ════ SETTINGS ════ -->
  <div id="v-settings" class="view" style="flex-direction:column">
    <div id="settings-lock">
      <div class="lock-card">
        <h3>🔒 Paramètres</h3>
        <p>Entrez le mot de passe administrateur pour accéder aux paramètres.</p>
        <input type="password" id="lock-pw" placeholder="MDP admin" style="width:100%;padding:9px 11px;border:2px solid #e5e7eb;border-radius:7px;font-size:calc(14px*var(--zf,1));outline:none;margin-bottom:10px" onkeydown="if(event.key==='Enter')unlockSettings()">
        <button onclick="unlockSettings()" style="width:100%;padding:10px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:calc(14px*var(--zf,1));font-weight:700;cursor:pointer">Déverrouiller</button>
        <div id="lock-err" style="color:#dc2626;margin-top:6px;font-size:calc(12px*var(--zf,1));text-align:center;min-height:14px"></div>
      </div>
    </div>
    <div id="v-settings-content">
      <div class="ss">
        <h3>📂 Fichier Excel de données</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Chemin complet vers le fichier Excel (.xlsx) contenant les onglets Declarations et Listes.</div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:6px">
          <input id="cfg-db-path" placeholder="C:\chemin\vers\fichier.xlsx" style="flex:1;min-width:200px;padding:7px 10px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-prim" onclick="setDbPath()">💾 Enregistrer</button>
        </div>
        <div id="cfg-db-status" style="font-size:calc(11px*var(--zf,1));color:var(--gray)"></div>
      </div>
      <div class="ss">
        <h3>🔐 Mots de passe pilotes</h3>
        <div id="pwd-list"></div>
        <div class="flex mt8">
          <input id="np-name" placeholder="Nom pilote" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <input id="np-pw" type="password" placeholder="MDP" style="flex:1;padding:5px 8px;border:1px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <button class="btn btn-prim" onclick="addPilot()">+ Ajouter</button>
        </div>
        <div class="flex mt8">
          <button class="btn btn-ok" onclick="savePwds()">💾 Enregistrer MDP</button>
        </div>
      </div>
      <div class="ss">
        <h3>🕐 Modèles horaires (par poste &amp; par jour)</h3>
        <div id="models-list"></div>
        <button class="btn btn-sec mt8" onclick="addModel()">+ Nouveau modèle</button>
        <div class="flex mt8">
          <button class="btn btn-ok" onclick="saveModels()">💾 Enregistrer modèles</button>
        </div>
      </div>
      <div class="ss">
        <h3>⚙ Référence production 8h</h3>
        <div class="flex">
          <label style="font-weight:600;font-size:calc(12px*var(--zf,1))">Prod ref (unités/8h):</label>
          <input id="cfg-pr" type="number" style="width:90px;padding:5px;border:1px solid var(--border);border-radius:5px">
          <button class="btn btn-ok" onclick="saveProdRef()">Enregistrer</button>
        </div>
      </div>
      <div class="ss">
        <h3>⛔ Liste des arrêts configurables</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:8px">Ajouter, modifier ou supprimer les boutons d'arrêt disponibles en production. Pris en compte immédiatement.</div>
        <div id="events-list-ui" style="margin-bottom:10px"></div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;background:#f8fafc;padding:8px;border-radius:7px;border:1px solid var(--border)">
          <input id="ev-new-label" placeholder="Nom de l'arrêt" style="flex:1;min-width:120px;padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
          <select id="ev-new-cat" style="padding:6px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1))">
            <option value="pb">🔴 Panne</option>
            <option value="ratt">🟠 Rattrapage</option>
            <option value="organisation">🔵 Organisation</option>
            <option value="autre">⚫ Autre</option>
          </select>
          <button class="btn btn-green" style="font-size:calc(11px*var(--zf,1));padding:5px 12px" onclick="addEvtItem()">+ Ajouter</button>
        </div>
        <button class="btn btn-prim" style="margin-top:8px;font-size:calc(12px*var(--zf,1))" onclick="saveEvtList()">💾 Enregistrer la liste</button>
      </div>
      <div class="ss" style="display:none"><div id="interposte-list-ui"></div></div>
      <div class="ss">
        <h3>⏱ Arrêts prévus (budget planifié)</h3>
        <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-bottom:10px">Les durées planifiées sont <b>déduites du temps de référence TRS</b> si le pilote les a réellement déclarées. Tout dépassement reste impactant. Mettre 0 pour désactiver.</div>
        <div style="display:flex;flex-direction:column;gap:6px">
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage court</span>
            <input type="number" id="ap-clean-short" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage long</span>
            <input type="number" id="ap-clean-long" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">🧹 Nettoyage très long</span>
            <input type="number" id="ap-clean-grand" min="0" max="240" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">📋 Réunion quotidienne</span>
            <input type="number" id="ap-meeting" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1px solid var(--border);border-radius:7px;padding:7px 10px">
            <span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:600;color:#374151">⏸ Pause</span>
            <input type="number" id="ap-pause" min="0" max="120" style="width:70px;padding:5px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(13px*var(--zf,1));text-align:right">
            <span style="font-size:calc(11px*var(--zf,1));color:var(--gray)">min</span>
          </div>
        </div>
        <div style="font-size:calc(10px*var(--zf,1));color:var(--gray);margin-top:6px">Pour l'affectation automatique : le libellé doit contenir "nettoyage court/long/très long", "réunion" ou "pause".</div>
        <button class="btn btn-prim" style="margin-top:10px;font-size:calc(12px*var(--zf,1))" onclick="saveArretsPrevus()">💾 Enregistrer arrêts prévus</button>
      </div>
    </div>
  </div>

</div><!-- /app -->

<!-- ════ MODAL: Choix type nettoyage ════ -->
<div class="overlay" id="m-nett-type" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:700;align-items:center;justify-content:center">
  <div class="card" style="width:360px;padding:22px;background:#fff;border-radius:14px;border-top:4px solid #f59e0b;box-shadow:0 8px 32px rgba(0,0,0,.18)">
    <div style="font-size:calc(15px*var(--zf,1));font-weight:800;color:#78350f;margin-bottom:16px">🧹 Type de nettoyage</div>
    <div style="display:flex;flex-direction:column;gap:10px">
      <button class="btn" id="nett-btn-court" style="text-align:left;padding:12px 16px;border-radius:10px;border:2px solid #fde68a;background:#fffbeb;font-size:calc(13px*var(--zf,1));font-weight:700;color:#92400e;transition:all .15s"
        onclick="doStartNettoyage('court')">
        🧹 Nettoyage court
        <span id="nett-lbl-court" style="font-size:calc(11px*var(--zf,1));font-weight:400;color:#b45309;display:block;margin-top:2px"></span>
      </button>
      <button class="btn" id="nett-btn-long" style="text-align:left;padding:12px 16px;border-radius:10px;border:2px solid #fcd34d;background:#fefce8;font-size:calc(13px*var(--zf,1));font-weight:700;color:#78350f;transition:all .15s"
        onclick="doStartNettoyage('long')">
        🧹 Nettoyage long
        <span id="nett-lbl-long" style="font-size:calc(11px*var(--zf,1));font-weight:400;color:#b45309;display:block;margin-top:2px"></span>
      </button>
      <button class="btn" id="nett-btn-grand" style="text-align:left;padding:12px 16px;border-radius:10px;border:2px solid #f59e0b;background:#fff7ed;font-size:calc(13px*var(--zf,1));font-weight:700;color:#7c2d12;transition:all .15s"
        onclick="doStartNettoyage('grand')">
        🧹 Nettoyage très long
        <span id="nett-lbl-grand" style="font-size:calc(11px*var(--zf,1));font-weight:400;color:#b45309;display:block;margin-top:2px"></span>
      </button>
    </div>
    <button class="btn btn-sec" style="margin-top:14px;width:100%;font-size:calc(12px*var(--zf,1))" onclick="closeM('m-nett-type')">Annuler</button>
  </div>
</div>

<!-- ════ MODAL: Déclarer un arrêt ════ -->
<div class="overlay" id="m-stop">
  <div class="mbox">
    <div class="mhdr red">
      <h2>⛔ Déclarer un arrêt</h2>
      <button style="background:none;border:none;cursor:pointer;color:#fff;font-size:calc(16px*var(--zf,1))" onclick="closeM('m-stop')">✕</button>
    </div>
    <div class="mbody">
      <div class="stop-section-lbl" style="color:#92400e">⏱ Arrêts prévus</div>
      <div class="stops-grid" style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:4px">
        <button style="background:#64748b;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer" onclick="closeM('m-stop');doPause()">⏸ Pause</button>
        <button style="background:#64748b;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer" onclick="closeM('m-stop');doReunion()">👥 Réunion</button>
        <button style="background:#78350f;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer" onclick="closeM('m-stop');doStartNettoyage('court')">🧹 Nettoyage court</button>
        <button style="background:#92400e;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer" onclick="closeM('m-stop');doStartNettoyage('long')">🧹 Nettoyage long</button>
        <button style="background:#a16207;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer" onclick="closeM('m-stop');doStartNettoyage('grand')">🧹 Nettoyage très long</button>
      </div>
      <div class="stop-section-lbl">🔄 Rattrapage</div>
      <div class="stops-grid ratt" id="sgrid-ratt"></div>
      <div class="stop-section-lbl">🔧 PB Technique</div>
      <div class="stops-grid pb" id="sgrid-pb"></div>
      <div class="stop-section-lbl" id="sgrid-nett-lbl" style="display:none">🟡 Nettoyage</div>
      <div class="stops-grid" id="sgrid-nettoyage" style="display:none"></div>
      <div class="stop-section-lbl" id="sgrid-org-lbl" style="display:none">🔵 Organisation</div>
      <div class="stops-grid" id="sgrid-organisation" style="display:none"></div>
      <div class="stop-section-lbl">✏ Arrêt libre / autre</div>
      <div class="custom-row">
        <input id="custom-stop-input" placeholder="Nom de l'arrêt…" maxlength="60">
        <button class="btn btn-amber" onclick="declareCustomStop()">Déclarer</button>
      </div>
    </div>
    <div class="mftr"><button class="btn btn-sec" onclick="closeM('m-stop')">Annuler</button></div>
  </div>
</div>

<!-- ════ MODAL: Fin d'OF/prod ════ -->
<div class="overlay" id="m-endprod">
  <div class="mbox wide">
    <div class="mhdr"><h2 id="ep-title">⏹ Fin d'OF/prod</h2></div>
    <div class="mbody">
      <div style="display:flex;gap:10px;margin-bottom:8px;align-items:flex-start;flex-wrap:wrap">
        <div style="flex:1;min-width:160px"><div class="ep-grid" id="ep-stats"></div></div>
        <div style="display:flex;gap:8px;flex-shrink:0">
          <div style="text-align:center">
            <div style="font-size:calc(9px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">Répartition</div>
            <svg id="ep-pie" viewBox="0 0 130 115" style="width:160px;height:142px;display:block"></svg>
          </div>
          <div style="text-align:center">
            <div style="font-size:calc(11px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:2px">TRS</div>
            <svg id="ep-gauge" viewBox="0 0 100 58" style="width:130px;display:block;margin:0 auto;margin-top:8px">
              <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
              <path id="ep-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
              <text x="50" y="46" text-anchor="middle" font-size="13" font-weight="800" fill="#1a1f5e" id="ep-gauge-pct">--%</text>
            </svg>
          </div>
        </div>
      </div>
      <div style="margin-bottom:10px">
        <div style="font-size:calc(10px*var(--zf,1));text-transform:uppercase;font-weight:700;color:var(--gray);margin-bottom:5px">Arrêts &amp; pauses</div>
        <table class="ep-tbl"><thead><tr><th>Type</th><th>Durée</th><th>%</th></tr></thead><tbody id="ep-stops"></tbody></table>
      </div>
      <div class="card" style="padding:8px;margin-bottom:0">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:4px">Timeline</div>
        <svg id="ep-tl" viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">
          <rect x="0" y="4" width="800" height="28" fill="#e2e8f0" rx="4"/>
        </svg>
        <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span></div>
      </div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-endprod')">Annuler</button>
      <button class="btn btn-danger btn-lg" onclick="confirmEndProd()">✓ Confirmer fin de production</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Éditer arrêt ════ -->
<div class="overlay" id="m-editstop">
  <div class="mbox">
    <div class="mhdr"><h2>✏ Modifier l'arrêt</h2></div>
    <div class="mbody">
      <input type="hidden" id="es-key">
      <div class="fr" style="margin-bottom:8px"><label>Type</label><select id="es-type"></select></div>
      <div class="fr" style="margin-bottom:8px"><label>Heure début</label><input type="time" id="es-deb" step="60"></div>
      <div class="fr" style="margin-bottom:8px"><label>Heure fin</label><input type="time" id="es-fin" step="60"></div>
      <div class="fr"><label>Commentaire</label><input type="text" id="es-cmt"></div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-editstop')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteStop()">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditStop()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Commentaire fin d'arrêt ════ -->
<div class="overlay" id="m-stopcmt">
  <div class="mbox" style="max-width:400px">
    <div class="mhdr red"><h2>✓ Terminer l'arrêt</h2></div>
    <div class="mbody">
      <input type="hidden" id="cmt-stop-key">
      <div id="cmt-stop-lbl" style="font-size:calc(13px*var(--zf,1));font-weight:700;color:var(--navy);margin-bottom:10px"></div>
      <div class="fr"><label>Commentaire (optionnel)</label><textarea id="cmt-stop-text" style="height:70px;resize:none;width:100%;padding:6px 8px;border:1.5px solid var(--border);border-radius:6px;font-size:calc(13px*var(--zf,1))" placeholder="Description de l'arrêt…"></textarea></div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-stopcmt')">Annuler</button>
      <button class="btn btn-ok btn-lg" onclick="confirmEndStop()">✓ Confirmer fin d'arrêt</button>
    </div>
  </div>
</div>

<!-- ════ MODAL: Éditer ligne ════ -->
<div class="overlay" id="m-editrow">
  <div class="mbox wide">
    <div class="mhdr">
      <h2 id="er-title">✏ Modifier la ligne</h2>
      <button style="background:none;border:none;cursor:pointer;font-size:calc(16px*var(--zf,1))" onclick="closeM('m-editrow')">✕</button>
    </div>
    <div class="mbody">
      <input type="hidden" id="er-rownum"><input type="hidden" id="er-rowtype">
      <div id="er-prod-fields">
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:6px">
          <div class="fr"><label>N° OF</label><input id="er-of"></div>
          <div class="fr"><label>Heure Début</label><input type="time" id="er-deb" step="60"></div>
          <div class="fr"><label>Heure Fin</label><input type="time" id="er-fin" step="60"></div>
          <div class="fr"><label>Poste</label><input id="er-poste"></div>
          <div class="fr"><label>Pilote</label><input id="er-pilote"></div>
          <div class="fr"><label>Co-Pilote</label><input id="er-copilote"></div>
          <div class="fr"><label>Nb Personnes</label><input type="number" id="er-nbpers"></div>
          <div class="fr"><label>Taille</label><select id="er-taille"><option value="">--</option></select></div>
          <div class="fr"><label>Code Produit</label><input id="er-codeprod"></div>
          <div class="fr"><label>Type Produit</label><select id="er-typeprod"><option value="">--</option></select></div>
          <div class="fr"><label>Qté Fabriquée</label><input type="number" id="er-qtefab"></div>
          <div class="fr"><label>Qté Emballée</label><input type="number" id="er-qteemb"></div>
          <div class="fr"><label>Poids Garnissage (g)</label><input type="number" id="er-poids"></div>
          <div class="fr"><label>Fibre</label><select id="er-fibre"><option value="">--</option></select></div>
          <div class="fr"><label>OF Taie</label><input id="er-oftaie"></div>
          <div class="fr"><label>Traca Fibre</label><select id="er-traca"><option value="">--</option></select></div>
          <div class="fr"><label>Réf Taie</label><input id="er-reftaie"></div>
          <div class="fr"><label>Lots de 2</label><select id="er-kit"><option value="">Non</option><option value="oui">Oui</option></select></div>
          <div class="fr"><label>Qté Init Taie</label><input type="number" id="er-qteinit"></div>
          <div class="fr"><label>Nb Taie 2nd Choix</label><input type="number" id="er-nbtaie2"></div>
          <div class="fr"><label>Nb Défaut Couture</label><input type="number" id="er-nbdef"></div>
          <div class="fr"><label>Mq Taie</label><input type="number" id="er-mqtaie"></div>
          <div class="fr"><label>Mq Housse/Encart</label><input type="number" id="er-mqhousse"></div>
          <div class="fr"><label>Nb PP Cousue</label><input type="number" id="er-nbpp"></div>
          <div class="fr"><label>Duree MQ MP (min)</label><input type="number" id="er-dureemq"></div>
          <div class="fr"><label>Manquant Personnel (min)</label><input type="number" id="er-manqpers"></div>
        </div>
        <div class="fr comment-big"><label>💬 Commentaire</label><textarea id="er-comment-prod" style="height:60px;resize:none;width:100%;padding:4px 6px;border:2px solid #f59e0b;border-radius:4px;font-size:calc(12px*var(--zf,1));background:#fffbeb"></textarea></div>
      </div>
      <div id="er-evt-fields" style="display:none">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:6px">
          <div class="fr"><label>Type d'arrêt</label><select id="er-evttype"><option value="">--</option></select></div>
          <div class="fr"><label>N° OF</label><input id="er-evtof"></div>
          <div class="fr"><label>Heure Début</label><input type="time" id="er-evtdeb" step="60"></div>
          <div class="fr"><label>Heure Fin</label><input type="time" id="er-evtfin" step="60"></div>
        </div>
        <div class="fr"><label>Commentaire</label><textarea id="er-evtcomment" style="height:44px;resize:none;width:100%;padding:4px 6px;border:1px solid var(--border);border-radius:4px;font-size:calc(12px*var(--zf,1))"></textarea></div>
        <div class="fr" style="margin-top:6px"><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-weight:600;font-size:calc(12px*var(--zf,1))"><input type="checkbox" id="er-horstrs"> Hors TRS</label></div>
      </div>
    </div>
    <div class="mftr">
      <button class="btn btn-sec" onclick="closeM('m-editrow')">Annuler</button>
      <button class="btn btn-danger" onclick="deleteRow(null,'er-rownum')">🗑 Supprimer</button>
      <button class="btn btn-ok" onclick="saveEditRow()">💾 Enregistrer</button>
    </div>
  </div>
</div>

<div id="m-pw-action" class="overlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:5000;align-items:center;justify-content:center">
  <div class="mbox" style="max-width:320px;padding:20px" onclick="event.stopPropagation()">
    <div class="mhdr" style="margin:-20px -20px 14px;padding:14px 16px;border-radius:12px 12px 0 0"><h2 id="pwa-title">🔒 Mot de passe Admin</h2></div>
    <div class="fr" style="margin-bottom:14px"><label>Mot de passe</label><input type="password" id="pwa-pw" placeholder="••••" style="padding:6px 10px;border:1.5px solid #cbd5e1;border-radius:6px;font-size:calc(14px*var(--zf,1));width:100%" onkeydown="if(event.key==='Enter')_pwaConfirm()"></div>
    <div style="display:flex;gap:8px">
      <button class="btn btn-prim" onclick="_pwaConfirm()">✓ Confirmer</button>
      <button class="btn btn-sec" onclick="closeM('m-pw-action')">Annuler</button>
    </div>
  </div>
</div>

<script>
// ── EVENTS definition (matches Python EVENTS list) ──
const EVENTS = [
  ["Pochon / Fibre","ratt_pochon","ratt"],
  ["Couture","ratt_couture","ratt"],
  ["Emballage","ratt_emb","ratt"],
  ["Presse Souder","ratt_presse_soud","ratt"],
  ["Presse ZIP","ratt_presse_zip","ratt"],
  ["Réunion","arret_reunion","ratt"],
  ["Chargeuse","pb_chargeuse","pb"],
  ["Carde","pb_carde","pb"],
  ["Étaleur / Tour","pb_etaleur","pb"],
  ["Coupe / Circ.","pb_coupe","pb"],
  ["Tapis Bascule","pb_tapis1","pb"],
  ["Enrouleur Pochon","pb_enrouleur","pb"],
  ["Pesée / Tapis 2","pb_pesee","pb"],
  ["Déviation / Table","pb_deviation","pb"],
  ["Enfileur Pochon","pb_enfileur","pb"],
  ["Kinna / Stroebel","pb_kinna","pb"],
  ["Tapeuse","pb_tapeuse","pb"],
  ["Table Rot. / Twin","pb_table_rot","pb"],
  ["Enfileuse H100","pb_h100","pb"],
  ["Enfileuse Traversin","pb_traversin","pb"],
  ["Presse ORC","pb_presse_orc","pb"],
  ["Presse Housse ZIP","pb_presse_zip2","pb"],
  ["Cercleuse","pb_cercleuse","pb"],
  ["Enrouleuse Traversin","pb_enrouleuse","pb"],
  ["Matière première","arret_mp","pb"],
];

const STOP_COL = {
  ratt:"#dc2626",pb:"#dc2626",autre:"#64748b",
  nettoyage:"#f59e0b",
  organisation:"#3b82f6",
  "_pause":"#94a3b8","Pause":"#94a3b8"
};

function getStopColor(key, cat) {
  if (cat) return STOP_COL[cat]||'#64748b';
  if (key==='nettoyage') return STOP_COL.nettoyage;
  if (key==='_pause'||key==='Pause') return STOP_COL['Pause'];
  return '#94a3b8';
}

const FORM_FIELDS = ["of_num","copilote","nb_pers","taille","code_prod","type_prod","poids","fibre","of_taie","traca","ref_taie","kit","qte_fab","qte_emb","qte_init_taie","nb_taie2_choix","nb_def_cout","mq_taie","mq_housse_encart","nb_pp_cousue","duree_mq_mp","manquant_pers","comment"];

// ── State ──
let ST = {};
let gEvts = [];
let _curStopKey = null;
let _curStopElap = 0;
let _ofElapAtPoll = 0;
let _stopWallAtPoll = 0;
let _pauseTotalAtPoll = 0;
let _pauseElapAtPoll = 0;
let _pauseBaseS = 0;
let _pauseStartMs = 0;
let _lastPoll = Date.now();
let _ticker = null;
let _autoSaveTimer = null;
let _curTab = 'main';
let _cfgPwds = {};
let _cfgModels = [];
let _cfgModelsBase = [];
let _settingsUnlocked = false;
let _adminPw = '';
window._evMap = {};

// ── INIT ──
document.addEventListener('DOMContentLoaded', async () => {
  await loadEvtsList(); // charge la liste dynamique des arrêts avant de construire les grilles
  await loadInterposteCfg();
  buildEditStopOpts();
  await loadLists();
  const s = await apiFetch('/api/state');
  if (s && (s.pilot || s.prod_active)) {
    document.getElementById('v-login').classList.remove('on');
    showApp(s);
    if (s.prod_active && !s.pilot) {
      toast('Prod en cours restaurée — vérifiez les infos','warn');
    }
  }
  setInterval(pollState, 5000);
  setInterval(pollEvts, 8000);
  setInterval(loadLists, 300000); // sync Excel lists every 5 min
  startTicker();
});

async function loadLists() {
  const d = await apiFetch('/api/lists');
  if (!d) return;
  popSel('f-taille', d.tailles||[]);
  popSel('f-type_prod', d.types_prod||[]);
  popSel('f-fibre', d.fibres||[]);
  popSel('f-copilote', d.copilotes||[]);
  popSel('er-taille', d.tailles||[]);
  popSel('er-typeprod', d.types_prod||[]);
  popSel('er-fibre', d.fibres||[]);
  popSel('er-traca', d.tracas||[]);
  // Build equivalence coef map: type_prod -> coef
  const eqs=d.equivalences||[]; const tps=d.types_prod||[];
  window._equivCoefs={};
  tps.forEach((t,i)=>{if(i<eqs.length&&eqs[i])window._equivCoefs[t]=parseFloat(String(eqs[i]).replace(',','.'))||1;});
  const pil = d.pilotes||[];
  const sel = document.getElementById('ln-pilot');
  // Clear existing options (except first placeholder)
  while(sel.options.length>1) sel.remove(1);
  pil.forEach(p => { const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o); });
  // If no pilots yet, retry after 1.5s (Excel may still be loading)
  if(!pil.length){
    setTimeout(async()=>{
      const d2=await apiFetch('/api/lists');
      if(!d2) return;
      const pil2=d2.pilotes||[];
      while(sel.options.length>1) sel.remove(1);
      pil2.forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o);});
    },1500);
  }
  // Load models for login select
  await loadModelsForLogin();
}

async function loadModelsForLogin() {
  const d = await apiFetch('/api/config');
  if (!d) return;
  _cfgModels = d.modeles_horaires||[];
  _cfgModelsBase = d.modeles_horaires_base||JSON.parse(JSON.stringify(_cfgModels));
  _cfgPwds = d.pilot_passwords||{};
  const sel = document.getElementById('ln-model');
  while (sel.options.length>1) sel.remove(1);
  _cfgModels.forEach(m => {
    const o=document.createElement('option'); o.value=m.nom||''; o.textContent=m.nom||''; sel.appendChild(o);
  });
  // Fallback: if no models, show POSTES
  if (!_cfgModels.length) {
    ["Matin","Après-midi","Nuit","Jour"].forEach(p => {
      const o=document.createElement('option'); o.value=p; o.textContent=p; sel.appendChild(o);
    });
  }
}

function popSel(id, vals) {
  const s = document.getElementById(id);
  if (!s) return;
  const cur = s.value;
  while (s.options.length>1) s.remove(1);
  vals.forEach(v => { const o=document.createElement('option'); o.value=v; o.textContent=v; s.appendChild(o); });
  if (cur) s.value = cur;
}

// ── Login model horaire info ──
function onLoginModelChange(){
  const nom=document.getElementById('ln-model').value;
  const info=document.getElementById('ln-model-info');
  const timesEl=document.getElementById('ln-model-times');
  const modifyDiv=document.getElementById('ln-model-modify');
  if(!nom){if(info)info.style.display='none';return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const todayKey=DAY_KEYS[new Date().getDay()];
  const model=_cfgModelsBase.find(m=>m.nom===nom);
  let debut='',fin='';
  if(model&&model.jours&&model.jours[todayKey]){
    debut=model.jours[todayKey].debut||'';
    fin=model.jours[todayKey].fin||'';
  }
  if(timesEl) timesEl.textContent=debut&&fin?`${debut} → ${fin}`:'Aucun horaire défini pour aujourd\'hui';
  // Pre-fill modify inputs
  const nd=document.getElementById('ln-new-debut'),nf=document.getElementById('ln-new-fin');
  if(nd) nd.value=debut; if(nf) nf.value=fin;
  if(modifyDiv) modifyDiv.style.display='none';
  if(info) info.style.display='block';
}

function toggleLoginModelModify(){
  const d=document.getElementById('ln-model-modify');
  if(d) d.style.display=d.style.display==='none'?'block':'none';
}

// ── LOGIN ──
async function doLogin() {
  const pilot = document.getElementById('ln-pilot').value;
  const poste = document.getElementById('ln-model').value;
  const pw = document.getElementById('ln-pw').value;
  document.getElementById('ln-err').textContent = '';
  if (!pilot) { document.getElementById('ln-err').textContent='Choisir un pilote'; return; }
  if (!poste) { document.getElementById('ln-err').textContent='Choisir un modèle horaire'; return; }
  const r = await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pilot,poste,pw})});
  if (!r) return;
  const d = await r.json();
  if (d.ok) {
    // Si l'utilisateur a modifié les horaires du jour, les sauvegarder avant d'entrer
    const modifyDiv=document.getElementById('ln-model-modify');
    if(modifyDiv&&modifyDiv.style.display!=='none'){
      const newDebut=document.getElementById('ln-new-debut').value;
      const newFin=document.getElementById('ln-new-fin').value;
      if(newDebut&&newFin){
        const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
        const todayKey=DAY_KEYS[new Date().getDay()];
        const mi=_cfgModels.findIndex(m=>m.nom===poste);
        if(mi>=0){
          if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
          if(!_cfgModels[mi].jours[todayKey])_cfgModels[mi].jours[todayKey]={};
          _cfgModels[mi].jours[todayKey].debut=newDebut;
          _cfgModels[mi].jours[todayKey].fin=newFin;
          await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:todayKey,debut:newDebut,fin:newFin})});
        }
        // Mettre à jour shift_debut_dt / shift_fin_dt et POSTES P/Q avec les horaires modifiés
        const today2=new Date();
        const [dH,dM]=newDebut.split(':').map(Number);
        const debDt=new Date(today2.getFullYear(),today2.getMonth(),today2.getDate(),dH,dM,0);
        const [fH,fM]=newFin.split(':').map(Number);
        const finDt2=new Date(today2.getFullYear(),today2.getMonth(),today2.getDate(),fH,fM,0);
        if(finDt2<=debDt) finDt2.setDate(finDt2.getDate()+1);
        await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:debDt.toISOString(),fin_iso:finDt2.toISOString()})});
      }
    }
    const s = await apiFetch('/api/state');
    document.getElementById('v-login').classList.remove('on');
    showApp(s||{pilot,poste});
  } else {
    document.getElementById('ln-err').textContent = d.error||'Erreur connexion';
  }
}

function doGuestLogin(){
  window._guestMode=true;
  document.getElementById('v-login').classList.remove('on');
  const app=document.getElementById('app');
  app.style.display='flex';app.classList.remove('hidden');
  setToday();
  loadCfg();
  _applyGuestMode();
  goTab('main');
  loadMainDecl();
}

function _applyGuestMode(){
  const actionBtns=document.getElementById('main-action-btns');
  if(actionBtns) actionBtns.style.display='none';
  const prodTab=document.getElementById('ht-prod');
  if(prodTab) prodTab.style.display='none';
  const guestBadge=document.getElementById('ht-guest-badge');
  if(guestBadge) guestBadge.style.display='';
  const pilotLbl=document.getElementById('hdr-pilot-lbl');
  if(pilotLbl) pilotLbl.textContent='👁 Invité';
}

function _getPobModelText(poste) {
  if(!poste||!_cfgModels) return poste||'—';
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  if(!m) return poste;
  const j=m.jours&&m.jours[dk];
  if(j&&j.debut&&j.fin) return `${poste} (${j.debut}→${j.fin})`;
  return poste;
}

function updatePobModel(poste){
  const el=document.getElementById('pob-model');
  if(el) el.textContent=_getPobModelText(poste||(_S_pilot_poste&&_S_pilot_poste.poste));
}

let _S_pilot_poste={poste:''};

function openPobModelEdit(){
  const poste=_S_pilot_poste&&_S_pilot_poste.poste;
  if(!poste){toast('Aucun modèle actif','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  const j=m&&m.jours&&m.jours[dk];
  document.getElementById('pobm-info').textContent=`Modèle : ${poste}`;
  document.getElementById('pobm-debut').value=(j&&j.debut)||'';
  document.getElementById('pobm-fin').value=(j&&j.fin)||'';
  openM('m-pobmodel');
}

async function savePobModelHours(){
  const poste=_S_pilot_poste&&_S_pilot_poste.poste;
  if(!poste) return;
  const debut=document.getElementById('pobm-debut').value;
  const fin=document.getElementById('pobm-fin').value;
  if(!debut||!fin){toast('Horaires incomplets','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(x=>x.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  updatePobModel(poste);
  closeM('m-pobmodel');
  toast('Horaires mis à jour','ok');
}

function showApp(s) {
  const app = document.getElementById('app');
  app.style.display = 'flex';
  app.classList.remove('hidden');
  if (s.pilot) { document.getElementById('f-pilote').value=s.pilot; document.getElementById('pob-pilot').textContent=s.pilot; }
  if (s.poste) {
    document.getElementById('f-poste').value=s.poste;
    document.getElementById('pob-poste').textContent=s.poste;
    _S_pilot_poste={poste:s.poste};
    updatePobModel(s.poste);
  }
  // Show prod tab button immediately if prod is active (don't wait for applyState)
  const tp=document.getElementById('ht-prod');
  if(tp) tp.classList.toggle('prod-visible',!!s.prod_active);
  // Also show the main banner immediately
  const mpb=document.getElementById('main-prod-banner');
  if(mpb) mpb.style.display=s.prod_active?'flex':'none';
  setToday();
  restoreFormFromStorage();
  pollState();
  pollEvts();
  loadCfg();
  goTab(s.prod_active ? 'prod' : 'main');
  const _today=new Date().toISOString().slice(0,10);
  const _hf=document.getElementById('hist-from'); if(_hf&&!_hf.value)_hf.value=_today;
  const _ht=document.getElementById('hist-to'); if(_ht&&!_ht.value)_ht.value=_today;
}

async function doLogout() {
  if (ST.prod_active) { toast('Terminer la production avant de déconnecter','err'); return; }
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
}

// ── Zoom texte ──
var _zoomOpen=false;
function applyZoom(pct){
  pct=Math.min(150,Math.max(70,pct));
  document.documentElement.style.removeProperty('zoom');
  document.documentElement.style.setProperty('--zf', pct/100);
  var sl=document.getElementById('zoom-slider');
  var vl=document.getElementById('zoom-val');
  if(sl) sl.value=pct;
  if(vl) vl.textContent=pct+'%';
  try{localStorage.setItem('kpi_zoom_pct',pct);}catch(e){}
}
function stepZoom(delta){
  var cur=+((document.getElementById('zoom-slider')||{value:100}).value);
  applyZoom(cur+delta);
}
function resetZoom(){applyZoom(135);}
function toggleZoomPop(){
  _zoomOpen=!_zoomOpen;
  var pop=document.getElementById('zoom-pop');
  if(pop) pop.style.display=_zoomOpen?'block':'none';
}
document.addEventListener('click',function(e){
  if(!_zoomOpen) return;
  var pop=document.getElementById('zoom-pop');
  var btn=document.getElementById('zoom-btn');
  if(pop&&btn&&!pop.contains(e.target)&&!btn.contains(e.target)){
    _zoomOpen=false;pop.style.display='none';
  }
});
(function(){
  try{var z=+localStorage.getItem('kpi_zoom_pct');applyZoom((z>=70&&z<=150)?z:135);}catch(e){applyZoom(135);}
})();

function resetToLogin() {
  ST={}; _curStopKey=null;
  window._guestMode=false;
  _missingDeclChecked=false;
  _ecartChecked=false;
  window._finPosteMode=false;
  document.getElementById('app').style.display='none';
  document.getElementById('ln-pw').value='';
  document.getElementById('ln-err').textContent='';
  document.getElementById('v-login').classList.add('on');
  _settingsUnlocked = false;
  // Restaurer les éléments cachés en mode invité
  const actionBtns=document.getElementById('main-action-btns');if(actionBtns) actionBtns.style.display='';
  const prodTab=document.getElementById('ht-prod');if(prodTab) prodTab.style.display='';
  const guestBadge=document.getElementById('ht-guest-badge');if(guestBadge) guestBadge.style.display='none';
  // Reset horaires login aux valeurs config
  const _lnM=document.getElementById('ln-model');
  if(_lnM) _lnM.value='';
  const _lnInfo=document.getElementById('ln-model-info');
  if(_lnInfo) _lnInfo.style.display='none';
  const _lnMod=document.getElementById('ln-model-modify');
  if(_lnMod) _lnMod.style.display='none';
  loadModelsForLogin();
}

function setToday() {
  const d=document.getElementById('f-date');
  if(d) d.value=new Date().toLocaleDateString('fr-CA');
}

// ── NAVIGATION ──
function goTab(tab) {
  if(window._guestMode && (tab==='prod'||tab==='finposte')){toast('Mode consultation — accès restreint','warn');return;}
  const _prevTab=_curTab;
  _curTab = tab;
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('on'));
  document.querySelectorAll('.htab').forEach(t=>t.classList.remove('on'));
  const vm={main:'v-main',prod:'v-prod',history:'v-history',settings:'v-settings',finposte:'v-finposte',kpi:'v-kpi',rapports:'v-rapports'};
  const el=document.getElementById(vm[tab]);
  if(el) el.classList.add('on');
  const nt={main:'ht-main',prod:'ht-prod',history:'ht-hist',settings:'ht-cfg',kpi:'ht-kpi',rapports:'ht-rapports'};
  const ntEl=document.getElementById(nt[tab]);
  if(ntEl) ntEl.classList.add('on');
  if(tab==='history') loadHist();
  else if(_prevTab==='history') _resetHistFilters();
  if(tab==='finposte') loadFPData();
  if(tab==='rapports') loadRapports();
  if(tab==='main') { loadMainDecl(); }
  if(tab==='kpi') loadKPI();
  if(tab==='settings') {
    _settingsUnlocked = false;
    document.getElementById('settings-lock').style.display = 'flex';
    document.getElementById('v-settings-content').style.display = 'none';
    document.getElementById('lock-pw').value='';
    document.getElementById('lock-err').textContent='';
  }
}

function switchMTab(t) { loadMainDecl(); } // kept for compatibility

// ── SETTINGS LOCK ──
async function unlockSettings() {
  const pw = document.getElementById('lock-pw').value;
  // Try with entered password against /api/settings
  const r = await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,_check_only:true})});
  const d = r?await r.json():{};
  if(d.ok) {
    _settingsUnlocked = true;
    _adminPw = pw;
    document.getElementById('settings-lock').style.display='none';
    document.getElementById('v-settings-content').style.display='block';
    loadCfg();
  } else {
    document.getElementById('lock-err').textContent = d.error||'Mot de passe incorrect';
  }
}

// ── POLLING ──
async function pollState() {
  const s = await apiFetch('/api/state');
  if(!s) return;
  ST=s;
  _ofElapAtPoll = s.of_elapsed_s||0;
  _stopWallAtPoll = s.stop_wall_s||0;
  _pauseTotalAtPoll = s.pause_total_s||0;
  _pauseBaseS = _pauseTotalAtPoll;
  _pauseStartMs = (s.is_paused && s.pause_start_iso) ? new Date(s.pause_start_iso).getTime() : 0;
  _pauseElapAtPoll = _pauseBaseS + (_pauseStartMs > 0 ? (Date.now()-_pauseStartMs)/1000 : 0);
  _lastPoll = Date.now();

  if(s.is_paused) {
    _curStopKey='_pause';
    _curStopElap = s.pause_total_s||0;
    if(s.pause_start_iso) _curStopElap += (Date.now()-new Date(s.pause_start_iso).getTime())/1000;
  } else if(s.active_stops&&s.active_stops.length>0) {
    const k=s.active_stops[0];
    _curStopKey=k;
    _curStopElap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
  } else {
    _curStopKey=null; _curStopElap=0;
  }

  applyState(s);
  if(_curTab==='main') loadMainDecl();
  if(_curTab==='kpi') loadKPI();
}

async function pollEvts() {
  const e = await apiFetch('/api/events_list');
  if(!Array.isArray(e)) return;
  // Filter to current pilot + today
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0');
  const todayPfx=dd+'/'+mm;
  const pilot=ST.pilot||'';
  gEvts=e.filter(ev=>{
    const dateOk=!ev.date||ev.date.startsWith(todayPfx);
    const pilotOk=!pilot||!ev.pilote||ev.pilote===pilot;
    return dateOk&&pilotOk;
  });
  // Don't override prod-view recap/timeline — applyState handles that from live tl_events
  if(!ST.prod_active){
    renderTL('tl-svg',gEvts);
    renderRecap(gEvts);
  }
}

function applyState(s) {
  // Header
  const hp=document.getElementById('hdr-pilot-lbl');
  if(hp&&s.pilot) hp.textContent=`${s.pilot} — ${s.poste||''}`;

  // Alert strip
  const al=document.getElementById('alert-strip');
  const stopOn=_curStopKey!==null;
  if(stopOn) {
    const lbl=_curStopKey==='_pause'?'PAUSE':getEvtLabel(_curStopKey);
    al.textContent=`⚠ ${lbl} EN COURS — cliquer pour terminer`;
    al.classList.add('on');
    al.onclick=()=>doEndStop();
    document.body.classList.add('stop-on');
  } else {
    al.classList.remove('on');
    al.onclick=null;
    document.body.classList.remove('stop-on');
  }

  // Prod tab visibility
  const tp=document.getElementById('ht-prod');
  if(tp) tp.classList.toggle('prod-visible',!!s.prod_active);

  // Main prod banner (shown when prod active but user is on main view)
  const mpb=document.getElementById('main-prod-banner');
  if(mpb) mpb.style.display=(s.prod_active&&_curTab==='main')?'flex':'none';
  if(s.prod_active&&s.form){
    const mpbOf=document.getElementById('mpb-of');
    if(mpbOf) mpbOf.textContent=s.form.of_num||'—';
  }

  // Excel busy bar
  const ebb=document.getElementById('excel-busy-bar');
  if(ebb) ebb.style.display=s.excel_busy?'block':'none';

  // Main btn-start
  const bs=document.getElementById('btn-start');
  if(bs) bs.disabled=s.prod_active;

  // POB (pilot/OF banner)
  if(s.prod_active&&s.form) {
    document.getElementById('pob-of').textContent=s.form.of_num||'—';
    // Don't override form while user is actively typing in it
    const af=document.activeElement;
    if(!af||!af.closest||!af.closest('.form-col')) fillFormFromState(s.form);
  }
  // OF start time in banner
  const pobOfStart=document.getElementById('pob-of-start');
  if(pobOfStart){
    if(s.of_start_iso){
      const os=new Date(s.of_start_iso);
      pobOfStart.textContent=String(os.getHours()).padStart(2,'0')+':'+String(os.getMinutes()).padStart(2,'0');
    } else {
      pobOfStart.textContent='—';
    }
  }

  // OF timing in pob banner
  const rcFin=document.getElementById('rc-fin');
  if(rcFin) rcFin.textContent=s.prod_active?'en cours':'—';

  // Render active stop chips (prod view + main view)
  renderStopChips(s);
  renderStopChipsMain(s);

  // Render live events for current prod (recap + timeline) from tl_events in state
  if(s.prod_active&&_curTab==='prod'){
    const le=tlEventsToDisplayFmt(s.tl_events||[]);
    // Add completed pauses from pause_periods
    (s.pause_periods||[]).forEach(([pStart,pEnd])=>{
      const s0=pStart?new Date(pStart):null;
      const e0=pEnd?new Date(pEnd):null;
      if(!s0) return;
      const pad=n=>String(n).padStart(2,'0');
      const toHMS=d=>pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
      const dur=e0?(e0.getTime()-s0.getTime())/1000:0;
      le.push({type:'Pause',cat:'_pause',debut:toHMS(s0),fin:e0?toHMS(e0):'',duree:dur>0?fmtDur(dur):'',comment:'',hors_trs:false,_live:!pEnd});
    });
    le.sort((a,b)=>a.debut.localeCompare(b.debut));
    renderTL('tl-svg',le);
    renderRecap(le);
  }

  // Pause button text
  const pbtn=document.getElementById('btn-pause');
  if(pbtn) pbtn.textContent=s.is_paused?'▶ Reprendre':'⏸ Pause';

  // TRS gauge
  updateGauge(s);
  updateAccModelInfo();
}

function getEvtLabel(key) {
  // Priorité : liste dynamique, puis EVENTS statique
  const dynEv=_evtsList.find(e=>e.key===key);
  if(dynEv) return (dynEv.cat==='ratt'?'Rattrapage: ':dynEv.cat==='pb'?'PB: ':'')+dynEv.label;
  const ev=EVENTS.find(e=>e[1]===key);
  if(ev) return (ev[2]==='ratt'?'Rattrapage: ':ev[2]==='pb'?'PB: ':'')+ev[0];
  if(key==='nettoyage') return 'Nettoyage';
  // Fallback par mot-clé (clés non normalisées ou issues d'une ancienne version)
  const kl=(key||'').toLowerCase();
  if(kl.includes('reunion')||kl.includes('union')||kl.includes('meeting')) return 'Réunion';
  if(kl.includes('pause')) return 'Pause';
  if(kl.includes('nettoyage')||kl.includes('nett')) return 'Nettoyage';
  // Recherche fuzzy dans _evtsList
  const fuzzy=_evtsList.find(e=>(e.label||'').toLowerCase().split(' ').some(w=>w.length>3&&kl.includes(w)));
  if(fuzzy) return fuzzy.label;
  return key||'Arrêt';
}

function tlEventsToDisplayFmt(tlEvts){
  return (tlEvts||[]).map(ev=>{
    if(!ev.key||ev.key.startsWith('_')) return null;
    const s=ev.start?new Date(ev.start):null;
    const e=ev.end?new Date(ev.end):null;
    if(!s) return null;
    const pad=n=>String(n).padStart(2,'0');
    const toHMS=d=>pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());
    const dur=e?(e.getTime()-s.getTime())/1000:0;
    let type;
    if(ev.cat==='nettoyage') type='Nettoyage'+(ev.nettoyage_type?' '+ev.nettoyage_type:'');
    else {
      const evDef=EVENTS.find(x=>x[1]===ev.key);
      if(evDef) type=(ev.cat==='ratt'?'Rattrapage: ':ev.cat==='pb'?'PB: ':'')+evDef[0];
      else type=ev.key;
    }
    return {type,cat:ev.cat||'autre',debut:toHMS(s),fin:e?toHMS(e):'',duree:dur>0?fmtDur(dur):'',comment:ev.comment||'',hors_trs:ev.hors_trs||false,_live:!ev.end};
  }).filter(Boolean);
}

let _lastMainChipKeys='';
function renderStopChipsMain(s) {
  const cont=document.getElementById('stop-chips-main');
  const sb=document.getElementById('stop-bottom-main');
  if(!cont||!sb) return;
  const stops=s.active_stops||[];
  const hasAny=stops.length>0||s.is_paused;
  if(!hasAny){sb.classList.remove('on');cont.innerHTML='';_lastMainChipKeys='';return;}
  sb.classList.add('on');
  const newKeys=stops.join(',')+(s.is_paused?'|pause':'');
  if(newKeys===_lastMainChipKeys){
    // Même set de stops : mettre à jour seulement les timers
    stops.forEach(k=>{
      const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
      const el=document.getElementById('main-chip-t-'+k);
      if(el) el.textContent=fmtDur2(elap);
    });
    if(s.is_paused){
      const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
      const el=document.getElementById('main-chip-t-_pause');
      if(el) el.textContent=fmtDur2(pe);
    }
    return;
  }
  _lastMainChipKeys=newKeys;
  let html='';
  stops.forEach(k=>{
    const lbl=getEvtLabel(k);
    const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
    html+=`<div class="stop-chip"><span class="chip-lbl">⛔ ${esc(lbl)}</span><span class="chip-tim" id="main-chip-t-${esc(k)}">${fmtDur2(elap)}</span><button class="btn-endstop" onclick="doEndStop('${esc(k)}')">✓ Terminer</button></div>`;
  });
  if(s.is_paused){
    const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
    html+=`<div class="stop-chip"><span class="chip-lbl">⏸ Pause</span><span class="chip-tim" id="main-chip-t-_pause">${fmtDur2(pe)}</span><button class="btn-endstop" onclick="doPause()">▶ Reprendre</button></div>`;
  }
  cont.innerHTML=html;
}

function renderStopChips(s) {
  const cont=document.getElementById('stop-chips');
  const sb=document.getElementById('stop-bottom');
  if(!cont||!sb) return;
  const stops=s.active_stops||[];
  const hasAny=stops.length>0||s.is_paused;
  if(!hasAny){sb.classList.remove('on');cont.innerHTML='';return;}
  sb.classList.add('on');
  let html='';
  stops.forEach(k=>{
    const lbl=getEvtLabel(k);
    const elap=s.timers&&s.timers[k]?s.timers[k].elapsed:0;
    html+=`<div class="stop-chip"><span class="chip-lbl">⛔ ${esc(lbl)}</span><span class="chip-tim" id="chip-t-${esc(k)}">${fmtDur2(elap)}</span><button class="btn-endstop" onclick="doEndStop('${esc(k)}')">✓ Terminer</button></div>`;
  });
  if(s.is_paused){
    const pe=(s.pause_total_s||0)+(s.pause_start_iso?(Date.now()-new Date(s.pause_start_iso).getTime())/1000:0);
    html+=`<div class="stop-chip"><span class="chip-lbl">⏸ Pause</span><span class="chip-tim" id="chip-t-_pause">${fmtDur2(pe)}</span><button class="btn-endstop" onclick="doPause()">▶ Reprendre</button></div>`;
  }
  cont.innerHTML=html;
}

// ── TICKER ──
function startTicker() {
  if(_ticker) clearInterval(_ticker);
  _ticker=setInterval(()=>{
    // Tick main-chip timers even without active prod
    {const dt0=(Date.now()-_lastPoll)/1000;
    (ST.active_stops||[]).forEach(k=>{
      const cel=document.getElementById('main-chip-t-'+k);
      if(cel&&ST.timers&&ST.timers[k]) cel.textContent=fmtDur2(ST.timers[k].elapsed+dt0);
    });
    if(ST.is_paused){const cel=document.getElementById('main-chip-t-_pause');
      if(cel) cel.textContent=fmtDur2(_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS);}
    }
    if(!ST.prod_active) return;
    const dt=(Date.now()-_lastPoll)/1000;
    // OF timer
    const ofEl=_ofElapAtPoll+(!ST.is_paused?dt:0);
    const t=document.getElementById('sc-of');
    if(t) t.textContent=fmtDur(ofEl);
    const rd=document.getElementById('rc-duree');
    if(rd) rd.textContent=fmtDur(ofEl);
    // Stops total
    const sw=_stopWallAtPoll+((_curStopKey&&_curStopKey!=='_pause')?dt:0);
    const ts=document.getElementById('sc-stops');
    if(ts) ts.textContent=fmtDur(sw);
    // Pause total
    const pauseNow=_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS;
    const tp=document.getElementById('sc-pause');
    if(tp) tp.textContent=fmtDur(pauseNow);
    // Pièces théoriques : avec déduction budget arrêts prévus
    const thEl=document.getElementById('sc-theo');
    if(thEl&&ST.prod_ref){
      const typeProd=document.getElementById('f-type_prod')?.value||ST.form?.type_prod||'';
      const coef=(window._equivCoefs&&window._equivCoefs[typeProd])||1;
      const _ofDedT=(ST.budget_state&&ST.budget_state.total_of_deductible_s)||0;
      const effOfElT=Math.max(1,_ofElapAtPoll+dt-_ofDedT);
      const theo=Math.round(ST.prod_ref*effOfElT/28800/coef);
      thEl.textContent=theo>0?theo+' pièces':'—';
    }
    // Mise à jour bannière accueil (durée OF et arrêts)
    if(ST.prod_active&&_curTab==='main'){
      const mpbDur=document.getElementById('mpb-dur');
      const mpbStops=document.getElementById('mpb-stops');
      const ofElBanner=_ofElapAtPoll+(!ST.is_paused?dt:0);
      const swBanner=_stopWallAtPoll+((_curStopKey&&_curStopKey!=='_pause')?dt:0);
      if(mpbDur) mpbDur.textContent=fmtDur(ofElBanner);
      if(mpbStops) mpbStops.textContent=fmtDur(swBanner);
    }
    // Update stop chips timers
    if(ST.active_stops){
      ST.active_stops.forEach(k=>{
        const cel=document.getElementById('chip-t-'+k);
        if(cel&&ST.timers&&ST.timers[k]) cel.textContent=fmtDur2(ST.timers[k].elapsed+dt);
      });
    }
    if(ST.is_paused){
      const cel=document.getElementById('chip-t-_pause');
      if(cel) cel.textContent=fmtDur2(_pauseStartMs>0?_pauseBaseS+(Date.now()-_pauseStartMs)/1000:_pauseBaseS);
    }
  },1000);
}

// ── MAIN VIEW ──
async function loadMainDecl() {
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0'),yyyy=now.getFullYear();
  const todayPfx=dd+'/'+mm;
  const [declData,evtData]=await Promise.all([apiFetch('/api/history'),apiFetch('/api/events_list')]);
  const _isProdRow=r=>['production','prod',''].includes((r.type||'').trim().toLowerCase());
  const decls=(Array.isArray(declData)?declData:(declData&&declData.rows?declData.rows:[])).filter(r=>
    (!r.date||r.date.startsWith(todayPfx))&&_isProdRow(r)
  );
  const evts=(Array.isArray(evtData)?evtData:[]).filter(r=>!r.date||r.date.startsWith(todayPfx));
  window._mainEvtsAll=evts;
  const allRows=[];
  decls.forEach(r=>allRows.push({...r,_rowType:'prod'}));
  evts.forEach(r=>allRows.push({...r,_rowType:'evt'}));
  allRows.sort((a,b)=>(b.debut||'').localeCompare(a.debut||''));
  const bd=document.getElementById('main-body');
  if(!bd) return;
  // Accumuler équivalences et arrêts du poste pour les jauges
  const curPilotD=ST.pilot||'';
  const pilotDecls=decls.filter(r=>!curPilotD||!r.pilote||r.pilote===curPilotD);
  const _hms2s=s=>{if(!s)return 0;const p=String(s).split(':');return p.length>=3?+p[0]*3600+ +p[1]*60+ +p[2]:p.length===2?+p[0]*60+ +p[1]:0;};
  // Filter to model horaire window only — exclude pre-shift prods (changement de série before model debut)
  const _dkLD=['dim','lun','mar','mer','jeu','ven','sam'][new Date().getDay()];
  const _mLD=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const _jLD=_mLD&&_mLD.jours&&_mLD.jours[_dkLD];
  const _mDebS=_jLD&&_jLD.debut?_hms2s(_jLD.debut):0;
  // Set global _shiftRefDt so loadMainKPI and loadKPI use the same reference
  if(_jLD&&_jLD.debut){const[_hS,_mS]=_jLD.debut.split(':').map(Number);_shiftRefDt=new Date();_shiftRefDt.setHours(_hS,_mS,0,0);}
  else if(ST.shift_start_iso){_shiftRefDt=new Date(ST.shift_start_iso);}
  else{_shiftRefDt=null;}
  const inShiftDecls=_mDebS>0?pilotDecls.filter(r=>_hms2s(r.fin||'')>=_mDebS||_hms2s(r.debut||'')>=_mDebS):pilotDecls;
  _todayEquivAccum=inShiftDecls.reduce((a,r)=>a+parseFloat(r.equiv||0),0);
  const _pilotEvts=evts.filter(r=>!curPilotD||!r.pilote||r.pilote===curPilotD);
  const inShiftEvts=_mDebS>0?_pilotEvts.filter(r=>_hms2s(r.fin||'')>=_mDebS||_hms2s(r.debut||'')>=_mDebS):_pilotEvts;
  window._lastMainRows=inShiftEvts; // pour checkMissingDecls()
  _todayStopAccum=inShiftEvts.reduce((a,r)=>a+_hms2s(r.duree||''),0);
  // Heure de la dernière déclaration prod enregistrée (dans la fenêtre du poste)
  if(inShiftDecls.length){
    const lastFin=inShiftDecls.map(r=>r.fin||'').filter(Boolean).sort().pop();
    if(lastFin){const[h,m,s]=(lastFin+'::').split(':').map(Number);const d=new Date();d.setHours(h,m,s||0,0);_lastProdDeclTime=d;}
  }
  if(!allRows.length){bd.innerHTML='<tr><td colspan="11" style="text-align:center;color:var(--gray);padding:16px">Aucune déclaration aujourd\'hui</td></tr>';loadMainKPI();updateGauge(ST);return;}
  window._rowMap={};
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(r.type&&r.type.toLowerCase().includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>');
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:'—';
    const cmt=esc(r.comment||'');
    const fibre=r.fibre||'';const fibreShort=esc(fibre.slice(0,9));
    return `<tr class="${isProd?'row-prod':'row-evt'}">
      <td>${tag}</td><td style="font-weight:700;color:${isProd?'#1e3a8a':'#dc2626'};text-decoration:underline;cursor:pointer" onclick="showMainRowDetail('${esc(String(key))}')" title="Voir détail">${esc(r.of||r.type||'—')}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:#6366f1;font-weight:600;cursor:${fibre?'pointer':''}" title="${esc(fibre)}" onclick="${fibre?'showFibre(\''+esc(fibre)+'\')':''}">${fibreShort}${fibre.length>9?'…':''}</td>
      <td style="font-size:calc(10px*var(--zf,1))">${esc(r.date||'')}</td><td style="font-size:calc(10px*var(--zf,1))">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:calc(11px*var(--zf,1))">${details}</td><td style="font-size:calc(11px*var(--zf,1))">${qty}</td><td>${info}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${cmt}">${cmt}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:calc(15px*var(--zf,1));cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
    </tr>`;
  }).join('');
  loadMainKPI();
  updateGauge(ST);
}

function saveMainModelHours(){
  const poste=ST.poste||'';if(!poste) return;
  const debut=document.getElementById('main-model-debut').value;
  const fin=document.getElementById('main-model-fin').value;
  if(!debut||!fin){toast('Horaires incomplets','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(m=>m.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  const el=document.getElementById('main-model-times');
  if(el) el.textContent=debut+' → '+fin;
  // Refresh Prod view model display and all TRS gauges/KPIs
  updatePobModel(poste);
  updateGauge(ST);
  loadMainKPI();
  loadKPI();
  toast('Horaires mis à jour','ok');
}

function updateMainModelDisplay(){
  const poste=ST.poste||'';
  const debutEl=document.getElementById('main-model-debut');
  const finEl=document.getElementById('main-model-fin');
  const timesEl=document.getElementById('main-model-times');
  const refCalcEl=document.getElementById('main-ref-calc');
  if(!poste) return;
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const m=_cfgModels.find(x=>x.nom===poste);
  const j=m&&m.jours&&m.jours[dk];
  const debut=(j&&j.debut)||'';const fin=(j&&j.fin)||'';
  if(timesEl) timesEl.textContent=debut&&fin?debut+' → '+fin:'Non défini';
  if(debutEl&&debut&&document.activeElement!==debutEl) debutEl.value=debut;
  if(finEl&&fin&&document.activeElement!==finEl) finEl.value=fin;
  // Show prod ref calculation
  if(refCalcEl&&ST.prod_ref>0){
    const ppm=(ST.prod_ref/480).toFixed(1);
    refCalcEl.textContent=`Réf: ${ST.prod_ref} pièces/8h = ${ppm} pcs/min`;
  }
}

async function loadMainKPI() {
  // Update model display
  updateMainModelDisplay();

  // Current shift
  const d=await apiFetch('/api/history_today');
  const evAll=await apiFetch('/api/events_list');
  const evArr=Array.isArray(evAll)?evAll:[];
  const now=new Date();
  const dd=String(now.getDate()).padStart(2,'0'),mm=String(now.getMonth()+1).padStart(2,'0');
  const todayPfx=dd+'/'+mm;
  const curPilot=ST.pilot||'';

  // Current shift stop time (today + current pilot)
  const curEvts=evArr.filter(e=>e.date&&e.date.startsWith(todayPfx)&&(!curPilot||!e.pilote||e.pilote===curPilot));
  const curStopS=curEvts.reduce((a,e)=>{try{const p=s=>s.split(':').reduce((acc,v,i)=>acc+(i===0?+v*3600:i===1?+v*60:+v),0);return a+Math.max(0,p(e.fin||'0:0:0')-p(e.debut||'0:0:0'));}catch(x){return a;}},0);
  // TRS à heure actuelle
  const heure=String(now.getHours()).padStart(2,'0')+'h'+String(now.getMinutes()).padStart(2,'0');
  const elHeure=document.getElementById('kpi0-heure');
  if(elHeure) elHeure.textContent=heure;
  if(d){
    // TRS Actuel : début de plage → fin de la dernière déclaration de prod, avec déduction budget
    let trs=-1;
    if(_todayEquivAccum>0&&_shiftRefDt&&ST.prod_ref>0){
      const refTime=_lastProdDeclTime||new Date();
      const shiftElap=(refTime.getTime()-_shiftRefDt.getTime())/1000;
      const _shDed=(ST.budget_state&&ST.budget_state.total_shift_deductible_s)||0;
      const effElap=Math.max(1,shiftElap-_shDed);
      if(effElap>0) trs=Math.round(_todayEquivAccum/(ST.prod_ref*effElap/28800)*100*10)/10;
    }
    const el0t=document.getElementById('kpi0-trs'),el0s=document.getElementById('kpi0-sub'),el0d=document.getElementById('kpi0-date');
    if(el0t) el0t.textContent=fmtTRSv(trs);
    if(el0s) el0s.textContent=(d.rows?d.rows.length:0)+' OF | Arrêts '+Math.round(curStopS/60)+' min';
    if(el0d){
      const today=now.toLocaleDateString('fr-FR',{day:'2-digit',month:'2-digit',year:'numeric'});
      const shiftIso=ST.shift_start_iso;
      const shiftTime=shiftIso?new Date(shiftIso).toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'}):null;
      el0d.textContent=today+(shiftTime?' — depuis '+shiftTime:'');
    }
  }

  // Stats pause/nettoyage/réunion/prod du poste actuel
  const _hms2sFn=s=>{if(!s)return 0;const p=String(s).split(':');return p.length>=3?+p[0]*3600+ +p[1]*60+ +p[2]:p.length===2?+p[0]*60+ +p[1]:0;};
  let pauseS=0,nettS=0,reunionS=0,arretS=0,prodS=0;
  curEvts.forEach(e=>{
    const dur=_hms2sFn(e.duree||'');
    const t=(e.type||'').toLowerCase();
    if(t.includes('pause')) pauseS+=dur;
    else if(t.includes('nett')) nettS+=dur;
    else if(t.includes('réunion')||t.includes('reunion')) reunionS+=dur;
    arretS+=dur;
  });
  // Prod = sum of durations from decls (today+pilot)
  if(d&&d.rows) d.rows.forEach(r=>{ prodS+=_hms2sFn(r.duree||''); });
  const _fm=s=>Math.round(s/60)+' min';
  const spEl=document.getElementById('main-stat-pause'),snEl=document.getElementById('main-stat-nett');
  const srEl=document.getElementById('main-stat-reunion'),saEl=document.getElementById('main-stat-arrets'),sprodEl=document.getElementById('main-stat-prod');
  if(spEl) spEl.textContent=_fm(pauseS);
  if(snEl) snEl.textContent=_fm(nettS);
  if(srEl) srEl.textContent=_fm(reunionS);
  if(saEl) saEl.textContent=_fm(arretS);
  if(sprodEl) sprodEl.textContent=_fm(prodS);
  // Nb OF in new accueil block
  const nbOfEl=document.getElementById('acc-nb-of');
  if(nbOfEl) nbOfEl.textContent=d&&d.rows?d.rows.length:0;

  // Previous sessions from history (group by pilot+date, exclude current session)
  const allRows=await apiFetch('/api/history');
  const rows=Array.isArray(allRows)?allRows:[];
  // Build sessions: group by pilot+date, excluding today+curPilot
  const sessions={};
  rows.forEach(r=>{
    const key=(r.pilote||'?')+'|'+(r.date||'?');
    if(r.date&&r.date.startsWith(todayPfx)&&r.pilote===curPilot) return; // skip current session
    if(!sessions[key]) sessions[key]={pilot:r.pilote||'?',poste:r.poste||'?',date:r.date||'?',rows:[],eq:0,trs_sum:0,trs_cnt:0};
    sessions[key].rows.push(r);
    sessions[key].eq+=parseFloat(r.equiv||0);
    const t=parseFloat(r.trs||0);
    if(t>0){sessions[key].trs_sum+=t;sessions[key].trs_cnt++;}
  });
  const sessArr=Object.values(sessions).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,2);
  for(let i=0;i<2;i++){
    const lbl=document.getElementById('kpi'+(i+1)+'-lbl'),tv=document.getElementById('kpi'+(i+1)+'-trs'),sv=document.getElementById('kpi'+(i+1)+'-sub');
    if(!lbl||!tv||!sv) continue;
    if(sessArr[i]){
      const s=sessArr[i];
      const avgT=s.trs_cnt>0?s.trs_sum/s.trs_cnt:-1;
      // Count stops for this session
      const sEvts=evArr.filter(e=>e.date===s.date&&e.pilote===s.pilot);
      const sStopS=sEvts.reduce((a,e)=>{try{const p=t=>t.split(':').reduce((acc,v,ii)=>acc+(ii===0?+v*3600:ii===1?+v*60:+v),0);return a+Math.max(0,p(e.fin||'0:0:0')-p(e.debut||'0:0:0'));}catch(x){return a;}},0);
      const dt=document.getElementById('kpi'+(i+1)+'-date');
      lbl.textContent=s.pilot+' — '+s.poste;
      tv.textContent=fmtTRSv(avgT);
      sv.textContent=s.rows.length+' OF | Arrêts '+Math.round(sStopS/60)+' min';
      if(dt){
        // Get first debut and last fin from this session's rows
        const debs=s.rows.map(r=>r.debut_of||r.debut||'').filter(Boolean).sort();
        const fins=s.rows.map(r=>r.fin_of||r.fin||'').filter(Boolean).sort().reverse();
        const timeRange=(debs.length&&fins.length)?` ${debs[0].slice(0,5)}→${fins[0].slice(0,5)}`:'';
        dt.textContent=s.date+timeRange;
      }
    } else {
      lbl.textContent=i===0?'Poste précédent':'Avant-dernier';
      tv.textContent='--%'; sv.textContent='—';
      const dt=document.getElementById('kpi'+(i+1)+'-date');
      if(dt) dt.textContent='';
    }
  }
}

// ── Remplit le groupe "Arrêts prévus" dans la popup interposte ──
function _fillIpArretsPrevus(){
  const g=document.getElementById('ip-arrprev-group');
  if(!g) return;
  g.innerHTML='';
  const ap=_cfgArretsPrevus||{};
  const ARRETS=[
    {key:'pause_min',lbl:'Pause',color:'#94a3b8'},
    {key:'clean_short_min',lbl:'Nettoyage court',color:'#f59e0b'},
    {key:'clean_long_min',lbl:'Nettoyage long',color:'#d97706'},
    {key:'clean_grand_min',lbl:'Nettoyage très long',color:'#b45309'},
    {key:'meeting_tol_min',lbl:'Réunion',color:'#3b82f6'},
  ];
  const visible=ARRETS.filter(a=>(ap[a.key]||0)>0);
  if(!visible.length){g.style.display='none';return;}
  g.style.display='block';
  const title=document.createElement('div');
  title.style.cssText='font-size:calc(10px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:6px;letter-spacing:.04em';
  title.textContent='Arrêts prévus';
  g.appendChild(title);
  const row=document.createElement('div');
  row.style.cssText='display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px;padding-bottom:10px;border-bottom:1px solid var(--border)';
  visible.forEach(a=>{
    const b=document.createElement('button');
    b.className='btn btn-ghost';
    b._origColor=a.color;
    b.style.cssText=`font-size:calc(12px*var(--zf,1));padding:5px 10px;border:2px solid ${a.color};color:${a.color};background:transparent;transition:all .15s`;
    b.textContent=a.lbl+' ('+ap[a.key]+'min max autorisé pendant ce poste)';
    b.onclick=()=>{
      document.getElementById('ip-custom').value=a.lbl;
      document.querySelectorAll('#ip-btns .btn, #ip-arrprev-group .btn').forEach(x=>{
        x.style.background='';x.style.color=x._origColor||'';x.style.borderColor=x._origColor||'var(--border)';x.style.transform='';
      });
      b.style.background=a.color;b.style.color='#fff';b.style.borderColor=a.color;
      b.style.transform='scale(0.93)';setTimeout(()=>{b.style.transform='';},150);
    };
    row.appendChild(b);
  });
  g.appendChild(row);
}

// ── START PROD ──
let _pendingGapS=0;
let _interposteLbls=["Changement de série","Réglage / Setup machine","Attente matière première","Réunion / Formation","Nettoyage interposte","Pause"];
let _interposteLblsEditing=[];

async function loadInterposteCfg(){
  const d=await apiFetch('/api/interposte_cfg');
  if(d&&d.labels&&d.labels.length) _interposteLbls=d.labels;
  _interposteLblsEditing=[..._interposteLbls];
  _renderInterposteLblsHTML();
}
function _renderInterposteLblsHTML(){
  const c=document.getElementById('interposte-list-ui');if(!c) return;
  c.innerHTML=_interposteLblsEditing.map((lbl,i)=>`
    <div style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1))">
      <span style="flex:1;font-weight:600">${esc(lbl)}</span>
      <button class="btn btn-ghost" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="editInterposteLbl(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmInterposteLbl(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun label</div>';
}
function addInterposteLbl(){const lbl=document.getElementById('ip-new-label').value.trim();if(!lbl){toast('Nom requis','err');return;}_interposteLblsEditing.push(lbl);document.getElementById('ip-new-label').value='';_renderInterposteLblsHTML();}
function rmInterposteLbl(i){_interposteLblsEditing.splice(i,1);_renderInterposteLblsHTML();}
function editInterposteLbl(i){const lbl=prompt('Label :',_interposteLblsEditing[i]);if(lbl&&lbl.trim()){_interposteLblsEditing[i]=lbl.trim();_renderInterposteLblsHTML();}}
async function saveInterposteCfg(){
  const r=await fetch('/api/interposte_cfg',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,labels:_interposteLblsEditing})});
  const d=r?await r.json():{};
  if(d&&d.ok){_interposteLbls=[..._interposteLblsEditing];toast('Labels interposte enregistrés','ok');}
  else toast(d?.error||'Erreur','err');
}

function _fmtMin(s){const m=Math.round(s/60),h=Math.floor(m/60),mi=m%60;return h?`${h}h ${mi}min`:`${mi} min`;}

// ── Gestion des gaps multiples (trous non déclarés avant chaque OF) ──
let _pendingGaps=[];
let _pendingGapIdx=0;
let _isFirstOfGaps=false;

function _showNextGap(){
  if(_pendingGapIdx>=_pendingGaps.length){goTab('prod');return;}
  const g=_pendingGaps[_pendingGapIdx];
  const total=_pendingGaps.length;
  const idx=_pendingGapIdx+1;
  const dur=_fmtMin(g.duree_s||0);
  // Titre + compteur
  const titleEl=document.getElementById('ps-title');
  if(titleEl) titleEl.textContent='⚠ Période non déclarée';
  const cntEl=document.getElementById('ps-counter');
  if(cntEl) cntEl.textContent=total>1?`Trou ${idx} / ${total}`:'';
  // Plage et durée
  const debut=g.debut||'';const fin=g.fin||'';
  document.getElementById('ps-text').textContent=`${dur} non déclarées : ${debut.replace(':','h')} → ${fin.replace(':','h')}`;
  document.getElementById('ps-gap-s').value=g.duree_s||0;
  document.getElementById('ps-start-iso').value='';
  document.getElementById('ps-custom').value='';
  // Bouton rétrodatage : seulement sur 1er OF + 1er gap
  const bdRow=document.getElementById('ps-backdate-row');
  const bt=document.getElementById('ps-backdate-time');
  const showBd=_isFirstOfGaps&&_pendingGapIdx===0;
  if(bdRow) bdRow.style.display=showBd?'':'none';
  if(bt) bt.textContent=debut.replace(':','h');
  psFillStopBtns();
  openM('m-preshift');
}

async function doStartProd() {
  saveFormToStorage();
  const r=await fetch('/api/start_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  if(!r) return;
  const d=await r.json();
  if(!d.ok){toast(d.error||'Erreur','err');return;}
  setToday();
  restoreFormFromStorage();
  await pollState();
  await pollEvts();
  _pendingGapS=d.gap_s||0;
  _pendingGaps=d.gaps||[];
  _pendingGapIdx=0;
  _isFirstOfGaps=!!(d.shift_model_start);
  _showNextGap();
}

// ── Ignorer ce trou sans déclarer, passer au suivant ──
async function psIgnorer(){
  closeM('m-preshift');
  _pendingGapIdx++;
  _showNextGap();
}

// ── Choix pré-poste ──
function psFillStopBtns(){
  const bc=document.getElementById('ps-stop-btns');
  if(!bc) return;
  bc.innerHTML='';

  const _makeBtn=(lbl,accent)=>{
    const b=document.createElement('button');
    b.className='btn btn-ghost';
    b.style.cssText=`font-size:calc(12px*var(--zf,1));transition:all .15s;border:2px solid ${accent||'var(--border)'};color:${accent||'var(--text)'};margin-bottom:3px`;
    b.textContent=lbl;
    b.onclick=()=>{
      document.getElementById('ps-custom').value=lbl;
      bc.querySelectorAll('.btn').forEach(x=>{x.style.background='';x.style.borderColor='var(--border)';x.style.color='var(--text)';x.style.transform='';});
      b.style.background='var(--navy)';b.style.color='#fff';b.style.borderColor='var(--navy)';
      b.style.transform='scale(0.93)';
      setTimeout(()=>confirmPsAsStop(),180);
    };
    return b;
  };

  const _makeSection=(title,labels,accent)=>{
    if(!labels.length) return;
    const hdr=document.createElement('div');
    hdr.style.cssText='font-size:calc(10px*var(--zf,1));font-weight:800;text-transform:uppercase;color:'+accent+';letter-spacing:.4px;margin:6px 0 4px;border-bottom:1px solid #e5e7eb;padding-bottom:2px';
    hdr.textContent=title;
    bc.appendChild(hdr);
    const row=document.createElement('div');
    row.style.cssText='display:flex;flex-wrap:wrap;gap:5px';
    labels.forEach(lbl=>row.appendChild(_makeBtn(lbl,accent)));
    bc.appendChild(row);
  };

  // Section "Arrêts prévus" depuis _cfgArretsPrevus
  const ap=_cfgArretsPrevus||{};
  const PREVUS_CFG=[
    {key:'pause_min',lbl:'Pause'},
    {key:'meeting_tol_min',lbl:'Réunion'},
    {key:'clean_short_min',lbl:'Nettoyage court'},
    {key:'clean_long_min',lbl:'Nettoyage long'},
    {key:'clean_grand_min',lbl:'Nettoyage très long'},
  ];
  const prevusList=PREVUS_CFG.filter(a=>(ap[a.key]||0)>0).map(a=>a.lbl);
  if(!prevusList.length) prevusList.push(...['Pause','Réunion','Nettoyage court','Nettoyage long']);
  _makeSection('⏱ Arrêts prévus',prevusList,'#d97706');

  // Sections par catégorie depuis _evtsList (paramètres)
  const evts=_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}));
  const prevusSet=new Set(prevusList.map(l=>l.toLowerCase()));
  const cats={pb:[],ratt:[],nettoyage:[],organisation:[],autre:[]};
  const seen=new Set(prevusList.map(l=>l.toLowerCase()));
  evts.forEach(e=>{
    const lbl=e.label||'';if(!lbl||seen.has(lbl.toLowerCase())) return;
    seen.add(lbl.toLowerCase());
    const c=e.cat||'autre';
    if(cats[c]!==undefined) cats[c].push(lbl);
    else cats.autre.push(lbl);
  });
  if(cats.pb.length||cats.ratt.length) _makeSection('🔴 Pannes / Rattrapages',[...cats.pb,...cats.ratt],'#dc2626');
  if(cats.nettoyage.length) _makeSection('🧹 Nettoyage',cats.nettoyage,'#f59e0b');
  if(cats.organisation.length) _makeSection('📋 Organisation',cats.organisation,'#3b82f6');
  if(cats.autre.length) _makeSection('⚫ Autre',cats.autre,'#64748b');

  document.getElementById('ps-custom').value='';
}

async function confirmPsAsStop(){
  const lbl=(document.getElementById('ps-custom').value||'').trim()||'Interposte';
  const gapS=parseFloat(document.getElementById('ps-gap-s').value)||0;
  const g=_pendingGaps[_pendingGapIdx]||{};
  closeM('m-preshift');
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({inter_of_s:gapS,label:lbl,comment:'',debut_hms:g.debut||'',fin_hms:g.fin||''})});
  await pollState();
  await pollEvts();
  loadMainDecl();
  _pendingGapIdx++;
  _showNextGap();
}

async function psChooseBackdate(){
  const startIso=document.getElementById('ps-start-iso').value;
  closeM('m-preshift');
  if(startIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:startIso})});
    await pollState();
    const bt=document.getElementById('ps-backdate-time');
    const tStr=bt?bt.textContent:'';
    toast('OF rétro-daté à '+(tStr||'l\'heure indiquée'),'ok');
  }
  // Rétrodatage = couvre toute la période → skip tous les gaps restants
  _pendingGapIdx=_pendingGaps.length;
  _showNextGap();
}

function ipShowModifyModel(){
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const jour=model&&model.jours&&model.jours[dk];
  const toHM=iso=>{if(!iso)return'';const d=new Date(iso);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  const deb=(jour&&jour.debut)||(model&&model.debut)||toHM(ST.shift_debut_iso)||'';
  const fin=(jour&&jour.fin)||(model&&model.fin)||toHM(ST.shift_fin_iso)||'';
  document.getElementById('ip-model-debut').value=deb;
  document.getElementById('ip-model-fin').value=fin;
  document.getElementById('ip-model-form').style.display='block';
}

async function ipConfirmModifyModel(){
  const newDebut=document.getElementById('ip-model-debut').value;
  const newFin=document.getElementById('ip-model-fin').value;
  if(!newDebut||!newFin){toast('Remplissez le début et la fin','warn');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mi=_cfgModels.findIndex(m=>m.nom===(ST.poste||''));
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=newDebut;
    _cfgModels[mi].jours[dk].fin=newFin;
    await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:ST.poste||'',day_key:dk,debut:newDebut,fin:newFin})});
  }
  document.getElementById('ip-model-form').style.display='none';
  // Recalculer le gap selon le contexte (fin de poste ou inter-OF)
  if(window._finPosteMode){
    // Gap = fin nouveau modèle - dernière déclaration
    if(_lastProdDeclTime){
      const[fh,fm]=newFin.split(':').map(Number);
      const modelFin=new Date();modelFin.setHours(fh,fm,0,0);
      if(modelFin<_lastProdDeclTime) modelFin.setDate(modelFin.getDate()+1); // poste de nuit
      const newGapS=(modelFin.getTime()-_lastProdDeclTime.getTime())/1000;
      if(newGapS>60){
        _pendingGapS=newGapS;
        document.getElementById('ip-duration').textContent=`Durée non déclarée : ${_fmtMin(newGapS)} (fin OF → fin modèle)`;
        const dh=String(_lastProdDeclTime.getHours()).padStart(2,'0'),dm=String(_lastProdDeclTime.getMinutes()).padStart(2,'0');
        document.getElementById('ip-debut').value=dh+':'+dm;
        document.getElementById('ip-fin').value=newFin;
        toast('Plage mise à jour','ok');
        return; // garder popup ouverte
      }
    }
    // Pas d'écart → aller directement à fin de poste
    window._finPosteMode=false;
    closeM('m-interposte');
    goTab('finposte');
  } else {
    // Contexte interposte normal : recalculer depuis model debut vers maintenant
    const[dh2,dm2]=newDebut.split(':').map(Number);
    const modelDebut=new Date();modelDebut.setHours(dh2,dm2,0,0);
    const now2=new Date();
    const newGapS=(now2.getTime()-modelDebut.getTime())/1000;
    if(newGapS>60){
      _pendingGapS=newGapS;
      document.getElementById('ip-duration').textContent=`Durée : ${_fmtMin(newGapS)} (début de poste → maintenant)`;
      document.getElementById('ip-debut').value=newDebut;
      const nh=String(now2.getHours()).padStart(2,'0'),nm=String(now2.getMinutes()).padStart(2,'0');
      document.getElementById('ip-fin').value=nh+':'+nm;
      toast('Plage mise à jour','ok');
      return; // garder popup ouverte
    }
    // Pas d'écart
    closeM('m-interposte');
    goTab('prod');
  }
}

async function confirmInterposte(){
  const lbl=document.getElementById('ip-custom').value.trim()||'Interposte';
  const cmt=document.getElementById('ip-comment').value.trim();
  const debutVal=document.getElementById('ip-debut').value;
  const finVal=document.getElementById('ip-fin').value;
  let interS=_pendingGapS;
  if(debutVal&&finVal){
    const[dh,dm]=debutVal.split(':').map(Number);const[fh,fm]=finVal.split(':').map(Number);
    let ds=dh*3600+dm*60,fs=fh*3600+fm*60;
    if(fs<ds) fs+=86400;
    interS=Math.max(0,fs-ds);
  }
  closeM('m-interposte');
  await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:interS,label:lbl,comment:cmt,debut_hms:debutVal,fin_hms:finVal})});
  // Si on vient d'un popup pré-poste, rétrodater le shift_start aussi
  if(window._psStartIso){
    await fetch('/api/set_of_start',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({iso:window._psStartIso})});
    window._psStartIso=null;
  }
  await pollState();
  await pollEvts();
  loadMainDecl(); // interposte row must appear in accueil without waiting
  if(window._finPosteMode){window._finPosteMode=false;await _doGoFinPoste();}else{goTab('prod');}
}

async function skipInterposte(){
  window._psStartIso=null;
  closeM('m-interposte');
  if(_pendingGapS>30){
    await fetch('/api/inter_of_confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({inter_of_s:_pendingGapS,label:'Interposte',comment:''})});
  }
  if(window._finPosteMode){window._finPosteMode=false;await _doGoFinPoste();}else{goTab('prod');}
}

async function doCancelProd(){
  await fetch('/api/force_reset_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
  goTab('main');
  toast('Production annulée','ok');
}

// ── STOP/PAUSE ──
let _evtsList=[]; // dynamic events list from server

async function loadEvtsList(){
  const d=await apiFetch('/api/events_cfg');
  if(!d||!d.events) return;
  _evtsList=d.events;
  // Sync with JS EVENTS array for getEvtLabel compatibility
  // EVENTS stays as fallback, but we prefer _evtsList
  rebuildStopGrids();
  renderEvtListUI();
}

function rebuildStopGrids(){
  const evts=_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}));
  const GRIDS={ratt:'sgrid-ratt',pb:'sgrid-pb',nettoyage:'sgrid-nettoyage',organisation:'sgrid-organisation'};
  const LABELS={nettoyage:'sgrid-nett-lbl',organisation:'sgrid-org-lbl'};
  Object.values(GRIDS).forEach(id=>{const g=document.getElementById(id);if(g)g.innerHTML='';});
  evts.forEach(e=>{
    const gid=GRIDS[e.cat]; if(!gid) return;
    const g=document.getElementById(gid); if(!g) return;
    const b=document.createElement('button');
    const col=STOP_COL[e.cat]||'#64748b';
    b.style.cssText=`background:${col};color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:calc(12px*var(--zf,1));font-weight:600;cursor:pointer`;
    b.textContent=e.label;
    b.onclick=()=>{closeM('m-stop');doStartStop(e.key,e.cat);};
    g.appendChild(b);
    g.style.display='';
    if(LABELS[e.cat]){const l=document.getElementById(LABELS[e.cat]);if(l)l.style.display='';}
  });
}

function buildStopGrids(){rebuildStopGrids();}

function getEvtLabelDynamic(key){
  const ev=_evtsList.find(e=>e.key===key)||EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]})).find(e=>e.key===key);
  if(!ev) return key||'Arrêt';
  return (ev.cat==='ratt'?'Rattrapage: ':ev.cat==='pb'?'PB: ':'')+ev.label;
}

// ── EVENTS LIST UI (Settings) ──
let _evtsEditing=[];
function renderEvtListUI(){
  _evtsEditing=(_evtsList.length?_evtsList:EVENTS.map(e=>({label:e[0],key:e[1],cat:e[2]}))).map(e=>({...e}));
  _renderEvtListHTML();
}

function _renderEvtListHTML(){
  const c=document.getElementById('events-list-ui');if(!c) return;
  const catLbl={pb:'🔴 Panne',ratt:'🟠 Rattrapage',nettoyage:'🟡 Nettoyage',organisation:'🔵 Organisation',autre:'⚫ Autre'};
  c.innerHTML=_evtsEditing.map((e,i)=>`
    <div style="display:flex;align-items:center;gap:6px;padding:5px 6px;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1))">
      <span style="flex:1;font-weight:600">${esc(e.label)}</span>
      <span style="font-size:calc(10px*var(--zf,1));color:var(--gray)">${catLbl[e.cat]||e.cat}</span>
      <button class="btn btn-ghost" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="editEvtItem(${i})">✏</button>
      <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmEvtItem(${i})">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:4px">Aucun arrêt configuré</div>';
}

function addEvtItem(){
  const lbl=document.getElementById('ev-new-label').value.trim();
  const cat=document.getElementById('ev-new-cat').value;
  if(!lbl){toast('Nom requis','err');return;}
  const key=lbl.toLowerCase().replace(/[^a-z0-9]/g,'_').slice(0,32)+'_'+Date.now().toString(36);
  _evtsEditing.push({label:lbl,key,cat});
  document.getElementById('ev-new-label').value='';
  _renderEvtListHTML();
}

function rmEvtItem(i){_evtsEditing.splice(i,1);_renderEvtListHTML();}

function editEvtItem(i){
  const e=_evtsEditing[i];
  const lbl=prompt('Nom de l\'arrêt :',e.label);
  if(lbl&&lbl.trim()){_evtsEditing[i].label=lbl.trim();_renderEvtListHTML();}
}

async function saveEvtList(){
  const r=await fetch('/api/events_cfg',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,events:_evtsEditing})});
  const d=r?await r.json():{};
  if(d&&d.ok){
    _evtsList=_evtsEditing.map(e=>({...e}));
    rebuildStopGrids();
    toast('Liste des arrêts enregistrée','ok');
  } else toast(d?.error||'Erreur','err');
}

function openStopModal(){openM('m-stop');}

async function doPause(){
  try{await fetch('/api/toggle_pause',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});}
  catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
  if(_curTab==='main') loadMainDecl();
}

async function doReunion(){
  // Trouver la clé arrêt de type réunion depuis _evtsList
  const reunionEvt=_evtsList.find(e=>/réunion|reunion|meeting/i.test(e.label||''));
  if(!reunionEvt){toast('Type Réunion non trouvé dans paramètres','err');return;}
  const key=reunionEvt.key, cat=reunionEvt.cat||'ratt';
  // Si déjà actif : terminer, sinon démarrer
  const isActive=ST.active_stops&&ST.active_stops.includes(key);
  if(isActive){
    await fetch('/api/end_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,comment:''})});
    toast('Réunion terminée','ok');
  } else {
    await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat,comment:''})});
    toast('Réunion commencée','ok');
  }
  await pollState();
  const btn=document.getElementById('btn-reunion');
  if(btn) btn.textContent=(!isActive)?'✓ Fin réunion':'👥 Réunion';
}

function doNettoyage(){
  // Populate budget labels from config
  const ap=_cfgArretsPrevus||{};
  const fmtMin=m=>m>0?`Budget autorisé : ${m} min`:'Non limité';
  const sc=document.getElementById('nett-lbl-court'),sl=document.getElementById('nett-lbl-long'),sg=document.getElementById('nett-lbl-grand');
  if(sc) sc.textContent=fmtMin(ap.clean_short_min||0);
  if(sl) sl.textContent=fmtMin(ap.clean_long_min||0);
  if(sg) sg.textContent=fmtMin(ap.clean_grand_min||0);
  openM('m-nett-type');
}

async function doStartNettoyage(ntype){
  closeM('m-nett-type');
  try{await fetch('/api/start_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ntype})});}
  catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
}

async function doStartStop(key,cat){
  try{
    const r=await fetch('/api/start_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key,cat})});
    if(!r||!r.ok) toast('Erreur déclaration arrêt','err');
  }catch(e){toast('Erreur connexion serveur','err');return;}
  await pollState();
  await pollEvts();
}

function doEndStop(key) {
  const k=key||_curStopKey;
  if(!k||k==='_pause'){doPause();return;}
  if(k==='nettoyage'){
    fetch('/api/end_nettoyage',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).then(()=>{pollState();pollEvts();if(_curTab==='main')loadMainDecl();});
    return;
  }
  // Show comment modal
  document.getElementById('cmt-stop-key').value=k;
  document.getElementById('cmt-stop-lbl').textContent='Arrêt : '+getEvtLabel(k);
  document.getElementById('cmt-stop-text').value='';
  openM('m-stopcmt');
  setTimeout(()=>document.getElementById('cmt-stop-text').focus(),100);
}

async function confirmEndStop() {
  const k=document.getElementById('cmt-stop-key').value;
  const cmt=document.getElementById('cmt-stop-text').value.trim();
  closeM('m-stopcmt');
  await fetch('/api/end_stop',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:k,comment:cmt})});
  await pollState();
  if(_curTab==='main') loadMainDecl();
  await pollEvts();
}

async function declareCustomStop(){
  const v=document.getElementById('custom-stop-input').value.trim();
  if(!v) return;
  document.getElementById('custom-stop-input').value='';
  closeM('m-stop');
  await doStartStop(v,'autre');
}

// ── FORM ──
function collectForm(){
  const f={};
  FORM_FIELDS.forEach(k=>{
    const el=document.getElementById('f-'+k);
    if(!el) return;
    f[k]=el.type==='number'?(parseFloat(el.value)||0):el.value;
  });
  f.pilote=document.getElementById('f-pilote')?.value||ST.pilot||'';
  f.poste=document.getElementById('f-poste')?.value||ST.poste||'';
  return f;
}

function fillFormFromState(form){
  if(!form) return;
  FORM_FIELDS.forEach(k=>{
    const el=document.getElementById('f-'+k);
    if(!el) return;
    const v=form[k];
    if(v!==undefined&&v!==null&&v!=='') el.value=v;
  });
  // Ensure nb_pers defaults to 10 if not set or 0
  const npEl=document.getElementById('f-nb_pers');
  if(npEl&&(!npEl.value||npEl.value==='0')) npEl.value='10';
  const ofEl=document.getElementById('pob-of');
  if(ofEl) ofEl.textContent=form.of_num||'—';
}

// Form persistence in localStorage (persist across restarts until new prod)
function saveFormToStorage(){
  const f=collectForm();
  try{localStorage.setItem('kpiorc_form',JSON.stringify(f));}catch(e){}
}
function restoreFormFromStorage(){
  try{
    const raw=localStorage.getItem('kpiorc_form');
    if(!raw) return;
    const f=JSON.parse(raw);
    // Only restore if prod_active and form has of_num, or if not active (post-prod)
    FORM_FIELDS.forEach(k=>{
      const el=document.getElementById('f-'+k);
      if(!el||!(k in f)) return;
      const v=f[k];
      if(v!==undefined&&v!==null&&v!=='') el.value=v;
    });
    const npEl=document.getElementById('f-nb_pers');
    if(npEl&&(!npEl.value||npEl.value==='0')) npEl.value='10';
  }catch(e){}
}

function scheduleAutoSave(){
  clearTimeout(_autoSaveTimer);
  _autoSaveTimer=setTimeout(()=>{
    const f=collectForm();
    fetch('/api/save_form',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(f)});
    saveFormToStorage();
  },1500);
}

// ── END PROD ──
async function doEndProdPreview(){
  const f=collectForm();
  const r=await fetch('/api/preview_end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if(!r||!r.ok){confirmEndProd();return;}
  const d=await r.json();
  renderEPModal(d,f);
  openM('m-endprod');
}

function renderEPModal(d,f){
  const ofNum=f.of_num||d.of_num||'';
  document.getElementById('ep-title').textContent=`⏹ Fin d'OF/prod${ofNum?' — '+ofNum:''}`;
  // Graphs
  drawPie('ep-pie',[
    {label:'Prod',value:d.prod_s||0,color:'#16a34a'},
    {label:'Arrêts',value:d.stop_s||0,color:'#dc2626'},
  ]);
  drawGauge('ep-gauge-arc','ep-gauge-pct',d.trs>=0?d.trs:0);
  const now=new Date();
  const dateStr=String(now.getDate()).padStart(2,'0')+'/'+String(now.getMonth()+1).padStart(2,'0')+'/'+now.getFullYear();
  document.getElementById('ep-stats').innerHTML=`
    <div class="ep-stat"><span class="lbl">N° OF</span><span class="val">${esc(ofNum||'—')}</span></div>
    <div class="ep-stat"><span class="lbl">Type produit</span><span class="val">${esc(f.type_prod||d.type_prod||'—')}</span></div>
    <div class="ep-stat"><span class="lbl">Qté fabriquée</span><span class="val">${esc(String(f.qte_fab||d.qte_fab||0))}</span></div>
    <div class="ep-stat"><span class="lbl">Qté emballée</span><span class="val">${esc(String(f.qte_emb||d.qte_emb||0))}</span></div>
    <div class="ep-stat"><span class="lbl">Date</span><span class="val">${dateStr}</span></div>
    <div class="ep-stat"><span class="lbl">TRS OF</span><span class="val">${fmtTRS(d.trs)}</span></div>
    <div class="ep-stat"><span class="lbl">Équivalence</span><span class="val">${(d.equiv||0).toFixed(1)}</span></div>
    <div class="ep-stat"><span class="lbl">Durée prod</span><span class="val">${fmtD2(d.prod_s||0)}</span></div>
    <div class="ep-stat"><span class="lbl">Total arrêts</span><span class="val">${fmtD2(d.stop_s||0)}</span></div>
  `;
  const evts=d.tl_events||[];
  const stopMap={};
  const totalS=d.stop_s||1;
  evts.forEach(e=>{
    if(!e.key||e.key==='prod'||e.key.startsWith('_')) return;
    const dur=(e.start&&e.end)?(new Date(e.end).getTime()-new Date(e.start).getTime())/1000:(e.dur_s||0);
    stopMap[e.key]=(stopMap[e.key]||0)+dur;
  });
  const tbody=document.getElementById('ep-stops');
  tbody.innerHTML=Object.entries(stopMap).map(([k,s])=>`
    <tr><td>${getEvtLabel(k)}</td><td>${fmtD2(s)}</td><td>${Math.round(s/totalS*100)}%</td></tr>
  `).join('')||'<tr><td colspan="3" style="color:var(--gray)">Aucun arrêt</td></tr>';
  const epDisplayEvts=tlEventsToDisplayFmt(evts);
  const _epS=parseHMStoT(d.debut,null)||(Date.now()-3600000);
  const _epE=parseHMStoT(d.now_str,null)||Date.now();
  drawTLFromISO('ep-tl',epDisplayEvts,new Date(_epS).toISOString(),new Date(_epE).toISOString());
}

async function confirmEndProd(){
  const f=collectForm();
  closeM('m-endprod');
  const r=await fetch('/api/end_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({form:f})});
  if(!r) return;
  const d=await r.json();
  if(d.ok){
    // Clear all form fields for next prod
    FORM_FIELDS.forEach(k=>{
      const el=document.getElementById('f-'+k);
      if(!el) return;
      el.value='';
    });
    document.getElementById('f-nb_pers').value='10';
    try{localStorage.removeItem('kpiorc_form');}catch(e){}
    await pollState();
    await pollEvts();
    goTab('main');
    toast('Production enregistrée','ok');
    // Auto-generate dashboard
    try{
      await fetch('/api/generate_dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    }catch(e){}
  } else toast(d.error||'Erreur','err');
}

async function forceResetProd(){
  await fetch('/api/force_reset_prod',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await pollState();
  toast('Production annulée','warn');
}

// ── PIE & GAUGE CHARTS ──
function drawPie(svgId, segments, opts) {
  const svg=document.getElementById(svgId);if(!svg) return;
  const fCenter=(opts&&opts.fCenter)||13;
  const fSub=(opts&&opts.fSub)||7;
  const fLeg=(opts&&opts.fLeg)||7;
  const total=segments.reduce((a,s)=>a+s.value,0);
  if(total<=0){svg.innerHTML=`<text x="65" y="60" text-anchor="middle" font-size="${fCenter}" fill="#94a3b8">Pas de données</text>`;return;}
  const cx=65,cy=57,r=44,ir=24;let html='',startAngle=-Math.PI/2;
  segments.forEach(seg=>{
    if(seg.value<=0) return;
    const angle=(seg.value/total)*2*Math.PI;if(angle<0.001) return;
    const endAngle=startAngle+angle,large=angle>Math.PI?1:0;
    const x1=(cx+r*Math.cos(startAngle)).toFixed(2),y1=(cy+r*Math.sin(startAngle)).toFixed(2);
    const x2=(cx+r*Math.cos(endAngle)).toFixed(2),y2=(cy+r*Math.sin(endAngle)).toFixed(2);
    const ix1=(cx+ir*Math.cos(startAngle)).toFixed(2),iy1=(cy+ir*Math.sin(startAngle)).toFixed(2);
    const ix2=(cx+ir*Math.cos(endAngle)).toFixed(2),iy2=(cy+ir*Math.sin(endAngle)).toFixed(2);
    html+=`<path d="M${x1},${y1} A${r},${r} 0 ${large},1 ${x2},${y2} L${ix2},${iy2} A${ir},${ir} 0 ${large},0 ${ix1},${iy1} Z" fill="${seg.color}"/>`;
    startAngle=endAngle;
  });
  const m=segments[0],mp=total>0?Math.round(m.value/total*100):0;
  html+=`<text x="${cx}" y="${cy+5}" text-anchor="middle" font-size="${fCenter}" font-weight="800" fill="#1a1f5e">${mp}%</text>`;
  html+=`<text x="${cx}" y="${cy+15}" text-anchor="middle" font-size="${fSub}" fill="#64748b">${esc(m.label)}</text>`;
  const legBase=(opts&&opts.legY)||108;
  let lx=0;segments.filter(s=>s.value>0).forEach(s=>{
    const p=Math.round(s.value/total*100);
    html+=`<rect x="${lx}" y="${legBase}" width="8" height="8" fill="${s.color}" rx="1"/>`;
    html+=`<text x="${lx+11}" y="${legBase+8}" font-size="${fLeg}" fill="#475569" font-weight="600">${esc(s.label)} ${p}%</text>`;lx+=65;
  });
  svg.innerHTML=html;
}

function drawGauge(arcId,pctId,trs) {
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct) return;
  const pArc=132,v=Math.max(0,Math.min(100,trs||0)),dash=(v/100)*pArc;
  const col=v>=90?'#16a34a':v>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash.toFixed(1)},${pArc}`);arc.setAttribute('stroke',col);
  pct.textContent=trs>=0?fmtTRS(trs):'--%';pct.setAttribute('fill',col);
}

// ── GAUGE (prod view) ──
function renderBudgetBars(containerId,bs){
  const el=document.getElementById(containerId);
  if(!el||!bs||!bs.per_type) return;
  const TYPES=['Pause','Réunion','Nettoyage court','Nettoyage long','Nettoyage très long'];
  const ap=_cfgArretsPrevus||{};
  let html='';
  let anyBar=false;
  TYPES.forEach(lbl=>{
    const d=bs.per_type[lbl]||{};
    const budget=d.budget_s||0;
    if(budget<=0) return;
    anyBar=true;
    const consumed=d.consumed_s||0;
    const pct=budget>0?Math.min(100,consumed/budget*100):0;
    const over=Math.max(0,consumed-budget);
    const color=pct>=100?'#dc2626':pct>=70?'#d97706':'#16a34a';
    const valStr=over>0?`<b>+${fmtDurShort(over)}</b>`:`${fmtDurShort(consumed)} / ${fmtDurShort(budget)}`;
    html+=`<div style="margin-bottom:5px">
      <div style="display:flex;justify-content:space-between;font-size:calc(10px*var(--zf,1));font-weight:700;margin-bottom:2px">
        <span style="color:#374151">${esc(lbl)}</span>
        <span style="color:${color}">${valStr}</span></div>
      <div style="background:#e5e7eb;border-radius:4px;height:6px">
        <div style="background:${color};width:${Math.min(100,pct).toFixed(0)}%;height:6px;border-radius:4px"></div>
      </div></div>`;
  });
  el.innerHTML=anyBar?html:'<div style="color:#92400e;font-size:calc(10px*var(--zf,1));opacity:.7">Aucun budget configuré</div>';
}

function updateGauge(s){
  const arc=document.getElementById('gauge-arc');
  const pct=document.getElementById('gauge-pct');
  const pobTrs=document.getElementById('pob-trs');
  if(!arc||!pct) return;
  // Budget arrêts prévus — déduction pour TRS OF et TRS Poste
  const bs=s.budget_state||{};
  const ofDed=(bs.total_of_deductible_s)||0;
  const shiftDed=(bs.total_shift_deductible_s)||0;
  // Estimate live TRS using coefficient
  const prodRef=s.prod_ref||200;
  const ofS=s.of_elapsed_s||0;
  const effOfS=Math.max(1,ofS-ofDed);
  const qFab=s.form?parseFloat(s.form.qte_fab||0):0;
  const typeProd=s.form?s.form.type_prod||'':'';
  const coef=(window._equivCoefs&&typeProd&&window._equivCoefs[typeProd])||1;
  const equiv=qFab*coef;
  let trs=-1;
  if(ofS>0&&prodRef>0&&equiv>0){
    trs=Math.round(equiv/(prodRef*effOfS/28800)*100*10)/10;
  }
  // Mise à jour barres budget accueil + prod en cours
  renderBudgetBars('budget-bars-acc',bs);
  renderBudgetBars('budget-bars-prod',bs);
  const trsStr=trs>=0?fmtTRS(trs):'—';
  pct.textContent=trsStr;
  if(pobTrs) pobTrs.textContent=trsStr;
  const pArc=132;
  const dash=trs>=0?Math.min(1,trs/100)*pArc:0;
  const col=trs>=90?'#16a34a':trs>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash},${pArc}`);
  arc.setAttribute('stroke',col);

  // TRS du poste (shift_start → dernière déclaration enregistrée)
  const arcPoste=document.getElementById('gauge-poste-arc');
  const pctPoste=document.getElementById('gauge-poste-pct');
  const lblPoste=document.getElementById('gauge-poste-lbl');
  // Même calcul pour l'accueil
  const arcPosteAcc=document.getElementById('gauge-poste-acc-arc');
  const pctPosteAcc=document.getElementById('gauge-poste-acc-pct');
  const lblPosteAcc=document.getElementById('gauge-poste-acc-lbl');
  // TRS Poste: prefer shift_debut_iso (set at login from model), fall back to model config
  if(s.shift_debut_iso){_shiftRefDt=new Date(s.shift_debut_iso);}
  else{
    const _dayKeysG=['dim','lun','mar','mer','jeu','ven','sam'];
    const _dkG=_dayKeysG[new Date().getDay()];
    const _modelG=_cfgModels&&_cfgModels.find(m=>m.nom===(s.poste||ST.poste||''));
    const _jourG=_modelG&&_modelG.jours&&_modelG.jours[_dkG];
    if(_jourG&&_jourG.debut){const[_hG,_mG]=_jourG.debut.split(':').map(Number);_shiftRefDt=new Date();_shiftRefDt.setHours(_hG,_mG,0,0);}
    else if(s.shift_start_iso){_shiftRefDt=new Date(s.shift_start_iso);}
    else{_shiftRefDt=null;}
  }
  // TRS Accueil : début de plage → fin de la dernière déclaration de prod, avec déduction budget
  if(_shiftRefDt&&s.prod_ref>0){
    const calcRef=_lastProdDeclTime||new Date();
    const shiftElap=(calcRef.getTime()-_shiftRefDt.getTime())/1000;
    const effShiftElap=Math.max(1,shiftElap-shiftDed);
    const todayEquiv=_todayEquivAccum||0;
    const trsPoste=effShiftElap>0&&todayEquiv>0?Math.round(todayEquiv/(s.prod_ref*effShiftElap/28800)*100*10)/10:-1;
    // Label : "Entre Xh et Yh" (Y = heure fin de la dernière déclaration, pas l'heure actuelle)
    let lbl;
    if(_lastProdDeclTime&&_shiftRefDt){
      const dh=String(_shiftRefDt.getHours()).padStart(2,'0'),dm=String(_shiftRefDt.getMinutes()).padStart(2,'0');
      const fh=String(_lastProdDeclTime.getHours()).padStart(2,'0'),fm=String(_lastProdDeclTime.getMinutes()).padStart(2,'0');
      lbl='Entre '+dh+'h'+dm+' et '+fh+'h'+fm;
    } else {
      lbl='TRS Poste';
    }
    const dashP=trsPoste>=0?Math.min(1,trsPoste/100)*pArc:0;
    const colP=trsPoste>=90?'#16a34a':trsPoste>=75?'#d97706':'#dc2626';
    [arcPoste,arcPosteAcc].forEach(el=>{if(el){el.setAttribute('stroke-dasharray',`${dashP},${pArc}`);el.setAttribute('stroke',colP);}});
    [pctPoste,pctPosteAcc].forEach(el=>{if(el){el.textContent=trsPoste>=0?fmtTRS(trsPoste):'—';el.setAttribute('fill',colP);}});
    [lblPoste,lblPosteAcc].forEach(el=>{if(el) el.textContent=lbl;});
  }

  // Pie charts
  const stopS=s.stop_wall_s||0;
  const ofDur=s.of_elapsed_s||0;
  const prodSof=Math.max(0,ofDur-stopS);
  drawPie('pie-of',[{label:'Prod',value:prodSof,color:'#16a34a'},{label:'Arrêts',value:stopS,color:'#dc2626'}],{fCenter:16,fSub:9,fLeg:9});
  // For poste pie — compute from shift start
  const shiftTotal=s.shift_start_iso?(Date.now()-new Date(s.shift_start_iso).getTime())/1000:0;
  const shiftStop=_todayStopAccum||0;
  const shiftProd=Math.max(0,shiftTotal-shiftStop);
  const postePieData=[{label:'Prod',value:shiftProd,color:'#16a34a'},{label:'Arrêts',value:shiftStop,color:'#dc2626'}];
  drawPie('pie-poste',postePieData,{fCenter:16,fSub:9,fLeg:9});
  drawPie('pie-poste-acc',postePieData,{fCenter:16,fSub:9,fLeg:9,legY:118});
}
// Accumulateurs poste (mis à jour à chaque loadMainDecl)
let _todayEquivAccum=0, _todayStopAccum=0, _lastProdDeclTime=null, _shiftRefDt=null;

// ── EDIT ROW (accueil) ──
function openEditRow(key) {
  const row=window._rowMap[String(key)];if(!row) return;
  const isProd=row._rowType==='prod';
  document.getElementById('er-rownum').value=row.row_num||'';
  document.getElementById('er-rowtype').value=row._rowType||'';
  document.getElementById('er-title').textContent=isProd?'✏ Modifier déclaration':'✏ Modifier événement';
  document.getElementById('er-prod-fields').style.display=isProd?'':'none';
  document.getElementById('er-evt-fields').style.display=isProd?'none':'';
  if(isProd){
    document.getElementById('er-of').value=row.of||'';
    document.getElementById('er-deb').value=(row.debut||'').slice(0,5);
    document.getElementById('er-fin').value=(row.fin||'').slice(0,5);
    document.getElementById('er-poste').value=row.poste||'';
    document.getElementById('er-pilote').value=row.pilote||'';
    document.getElementById('er-copilote').value=row.copilote||'';
    document.getElementById('er-nbpers').value=row.nb_pers||'';
    document.getElementById('er-taille').value=row.taille||'';
    document.getElementById('er-codeprod').value=row.code_prod||'';
    document.getElementById('er-typeprod').value=row.type_prod||'';
    document.getElementById('er-qtefab').value=row.qte_fab||'';
    document.getElementById('er-qteemb').value=row.qte_emb||'';
    document.getElementById('er-poids').value=row.poids||'';
    document.getElementById('er-fibre').value=row.fibre||'';
    document.getElementById('er-oftaie').value=row.of_taie||'';
    document.getElementById('er-traca').value=row.traca||'';
    document.getElementById('er-reftaie').value=row.ref_taie||'';
    document.getElementById('er-kit').value=row.kit||'';
    document.getElementById('er-qteinit').value=row.qte_init_taie||'';
    document.getElementById('er-nbtaie2').value=row.nb_taie2_choix||'';
    document.getElementById('er-nbdef').value=row.nb_def_cout||'';
    document.getElementById('er-mqtaie').value=row.mq_taie||'';
    document.getElementById('er-mqhousse').value=row.mq_housse_encart||'';
    document.getElementById('er-nbpp').value=row.nb_pp_cousue||'';
    document.getElementById('er-dureemq').value=row.duree_mq_mp||'';
    document.getElementById('er-manqpers').value=row.manquant_pers||'';
    document.getElementById('er-comment-prod').value=row.comment||'';
  } else {
    document.getElementById('er-evttype').value=row.type||'';
    document.getElementById('er-evtof').value=row.of||'';
    document.getElementById('er-evtdeb').value=(row.debut||'').slice(0,5);
    document.getElementById('er-evtfin').value=(row.fin||'').slice(0,5);
    document.getElementById('er-evtcomment').value=row.comment||'';
    document.getElementById('er-horstrs').checked=row.hors_trs||false;
  }
  openM('m-editrow');
}

let _pwaCallback=null;
function _showPwAction(title,cb){
  document.getElementById('pwa-title').textContent=title;
  document.getElementById('pwa-pw').value='';
  _pwaCallback=cb;
  openM('m-pw-action');
  setTimeout(()=>document.getElementById('pwa-pw').focus(),80);
}
function _pwaConfirm(){
  const pw=document.getElementById('pwa-pw').value;
  closeM('m-pw-action');
  if(_pwaCallback){_pwaCallback(pw);_pwaCallback=null;}
}

function saveEditRow() {
  _showPwAction('🔒 Confirmer modification',async function(pw){
  const rowNum=parseInt(document.getElementById('er-rownum').value);
  const rowType=document.getElementById('er-rowtype').value;
  if(!rowNum){toast('Ligne invalide','err');return;}
  if(!pw){toast('Mot de passe requis','err');return;}
  const v=id=>document.getElementById(id)?.value||'';
  const n=id=>parseFloat(document.getElementById(id)?.value)||0;
  let updates={};
  if(rowType==='prod'){
    // Col indices per DECL_HEADERS (1-based):
    // 2=OF,4=Poste,5=Pilote,6=Co-Pilote,7=Nb Pers,8=Taille,9=Code Prod,10=Type Prod
    // 11=Poids,12=Fibre,13=OF Taie,14=Traca,15=Ref Taie,16=Kit
    // 17=Heure Debut,18=Heure Fin,20=Qte Fab,21=Qte Emb
    // 26=Qte Init Taie,27=Nb Taie 2nd,28=Nb Def Cout,29=Mq Taie,30=Mq Housse,31=Nb PP
    // 33=Manquant MP (duree),34=Manquant Pers,36=Commentaire
    updates={'2':v('er-of'),'4':v('er-poste'),'5':v('er-pilote'),'6':v('er-copilote'),
      '7':n('er-nbpers'),'8':v('er-taille'),'9':v('er-codeprod'),'10':v('er-typeprod'),
      '11':n('er-poids'),'12':v('er-fibre'),'13':v('er-oftaie'),'14':v('er-traca'),
      '15':v('er-reftaie'),'16':v('er-kit'),
      '17':v('er-deb'),'18':v('er-fin'),'20':n('er-qtefab'),'21':n('er-qteemb'),
      '26':n('er-qteinit'),'27':n('er-nbtaie2'),'28':n('er-nbdef'),
      '29':n('er-mqtaie'),'30':n('er-mqhousse'),'31':n('er-nbpp'),
      '33':n('er-dureemq'),'34':n('er-manqpers'),'36':v('er-comment-prod')};
  } else {
    updates={'1':v('er-evttype'),'2':v('er-evtof'),'17':v('er-evtdeb'),'18':v('er-evtfin'),
      '36':v('er-evtcomment'),'37':document.getElementById('er-horstrs').checked?'OUI':''};
  }
  const r=await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rowNum,updates})});
  const d=r?await r.json():{};
  if(d&&d.ok){closeM('m-editrow');await loadMainDecl();if(_curTab==='history')await loadHist();loadKPI();toast('Ligne modifiée','ok');}
  else toast(d?.error||'Erreur modification','err');
  });
}

function deleteRow(key,rowNumId) {
  const rn=rowNumId?parseInt(document.getElementById(rowNumId)?.value):parseInt(window._rowMap[String(key)]?.row_num);
  if(!rn) return;
  _showPwAction('🗑 Confirmer suppression',async function(pw){
    if(!pw){toast('Mot de passe requis','err');return;}
    const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw,row_num:rn})});
    const d=r?await r.json():{};
    if(d&&d.ok){closeM('m-editrow');await loadMainDecl();if(_curTab==='history')await loadHist();loadKPI();toast('Supprimé','ok');}
    else toast(d?.error||'Erreur suppression','err');
  });
}

// ── EDIT STOP (recap prod view) ──
function buildEditStopOpts(){
  const s=document.getElementById('es-type');
  const s2=document.getElementById('er-evttype');
  const allOpts=[['Pause','_pause'],['Nettoyage','nettoyage'],...EVENTS.map(e=>[e[0],e[0]]),['Arrêt libre','Arrêt libre']];
  [s,s2].forEach(sel=>{
    if(!sel) return;
    allOpts.forEach(([lbl,val])=>{const o=document.createElement('option');o.value=val;o.textContent=lbl;sel.appendChild(o);});
  });
  // es-type uses key values (not display)
  if(s){s.innerHTML='';[['Pause','_pause'],['Nettoyage','nettoyage'],...EVENTS,['Arrêt libre','autre']].forEach(([lbl,key])=>{const o=document.createElement('option');o.value=key;o.textContent=lbl;s.appendChild(o);});}
}

function openEditStop(key){
  const ev=window._evMap[key];
  if(!ev) return;
  document.getElementById('es-key').value=key;
  document.getElementById('es-type').value=ev.type||key||'';
  const d=ev.debut||'',f2=ev.fin||'';
  document.getElementById('es-deb').value=d.length>=5?d.slice(0,5):d;
  document.getElementById('es-fin').value=f2.length>=5?f2.slice(0,5):f2;
  document.getElementById('es-cmt').value=ev.comment||'';
  openM('m-editstop');
}

async function saveEditStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev) return;
  const data={row_num:ev.row_num,type:document.getElementById('es-type').value,heure_debut:document.getElementById('es-deb').value,heure_fin:document.getElementById('es-fin').value,comment:document.getElementById('es-cmt').value};
  const r=await fetch('/api/edit_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Modifié','ok');}
  else toast('Erreur','err');
}

async function deleteStop(){
  const key=document.getElementById('es-key').value;
  const ev=window._evMap[key];
  if(!ev) return;
  const r=await fetch('/api/delete_row',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({row_num:ev.row_num})});
  if(r&&r.ok){closeM('m-editstop');await pollEvts();await loadMainDecl();loadKPI();toast('Supprimé','ok');}
}

// ── TIMELINE ──
function renderTL(svgId,evts){
  const now=new Date(), s4h=new Date(now-4*3600*1000);
  drawTLFromISO(svgId,evts,s4h.toISOString(),now.toISOString());
}

function drawTL(svgId,tlEvts,debutHMS,finHMS){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=52;
  let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  if(!debutHMS||!finHMS){svg.innerHTML=html;return;}
  const base=new Date();base.setHours(0,0,0,0);
  const pHMS=s=>{const[h,m,sec]=(s||'').split(':');return base.getTime()+(+h||0)*3600000+(+m||0)*60000+(+sec||0)*1000;};
  const tS=pHMS(debutHMS),tE=pHMS(finHMS);
  const span=tE-tS;if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  html+=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  let cur=tS;
  (tlEvts||[]).forEach(e=>{
    if(!e.key||e.key==='prod') return;
    const cat=e.cat||'autre';
    const x1=toX(cur),x2=toX(cur+(e.dur_s||0)*1000);
    html+=`<rect x="${x1}" y="${Y}" width="${Math.max(1,x2-x1)}" height="${H2}" fill="${STOP_COL[cat]||'#94a3b8'}" rx="2" opacity=".9"/>`;
    cur+=(e.dur_s||0)*1000;
  });
  html+=`<text x="2" y="${H-2}" font-size="10" fill="#374151" font-weight="600">${debutHMS.slice(0,5)}</text>`;
  html+=`<text x="${W-36}" y="${H-2}" font-size="10" fill="#374151" font-weight="600">${finHMS.slice(0,5)}</text>`;
  svg.innerHTML=html;
}

function drawTLFromISO(svgId,evts,startIso,endIso,prodOfList){
  const svg=document.getElementById(svgId);
  if(!svg) return;
  const W=800,Y=4,H2=28,H=52;
  let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
  const tS=new Date(startIso).getTime(),tE=new Date(endIso).getTime();
  const span=tE-tS;if(span<=0){svg.innerHTML=html;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  // Draw historical prod blocks from of_list (fin de poste view)
  (prodOfList||[]).forEach(of=>{
    const t1=parseHMStoT(of.debut,of.date||null);
    const t2=of.fin?parseHMStoT(of.fin,of.date||null):null;
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  });
  // Prod background — show for current active OF
  if(ST.of_start_iso&&ST.prod_active){
    const ps=new Date(ST.of_start_iso).getTime();
    const x1=toX(ps),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="4"/>`;
  }
  // Events
  (evts||[]).forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[String(key)]=ev;
    const t1=parseHMStoT(ev.debut,ev.date),t2=parseHMStoT(ev.fin,ev.date);
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||tE);
    if(x2<=x1) return;
    const cat=ev.cat||'autre';
    html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${STOP_COL[cat]||'#94a3b8'}" rx="2" opacity=".85"/>`;
  });
  // Current live stop
  if(_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${STOP_COL.pb||'#b91c1c'}" rx="2" opacity=".9"/>`;
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  // Hourly tick marks
  let tickT=Math.ceil(tS/3600000)*3600000;
  while(tickT<tE){
    const tx=toX(tickT);
    const hr=new Date(tickT).getHours();
    html+=`<line x1="${tx}" y1="${Y}" x2="${tx}" y2="${Y+H2}" stroke="rgba(0,0,0,.2)" stroke-width="1"/>`;
    html+=`<text x="${tx+2}" y="${Y+H2+14}" font-size="9" fill="#374151" font-weight="600">${String(hr).padStart(2,'0')}h</text>`;
    tickT+=3600000;
  }
  html+=`<text x="2" y="${Y+H2+14}" font-size="10" fill="#374151" font-weight="600">${fT(tS)}</text>`;
  html+=`<text x="${W-36}" y="${Y+H2+14}" font-size="10" fill="#374151" font-weight="600">${fT(tE)}</text>`;
  svg.innerHTML=html;
}

function parseHMStoT(hms,dateStr){
  if(!hms) return null;
  try{
    let base;
    if(dateStr){
      // Support both ISO (yyyy-mm-dd) and French (dd/mm/yyyy) formats
      const parts=dateStr.split('/');
      if(parts.length===3&&parts[2].length===4){
        base=new Date(parts[2]+'-'+parts[1]+'-'+parts[0]).getTime();
      } else {
        base=new Date(dateStr.slice(0,10)).getTime();
      }
    } else {
      const n=new Date();n.setHours(0,0,0,0);base=n.getTime();
    }
    const[h,m,s]=(hms||'').split(':').map(Number);return base+h*3600000+m*60000+(s||0)*1000;
  }catch(e){return null;}
}

// ── RECAP ──
function renderRecap(evts){
  const c=document.getElementById('recap-list');
  if(!c) return;
  window._evMap={};
  const stops=(evts||[]).filter(e=>e.type);
  if(!stops.length){c.innerHTML='<div style="color:var(--gray);font-size:calc(10px*var(--zf,1));padding:4px">Aucun arrêt</div>';return;}
  let html='';
  stops.forEach((ev,i)=>{
    const key=ev.debut||i;
    window._evMap[String(key)]=ev;
    const cat=ev.cat||'autre';
    const col=STOP_COL[cat]||'#94a3b8';
    html+=`<div class="si">
      <div class="sdot" style="background:${col}"></div>
      <div class="si-nm">${ev.type||'?'}</div>
      <div class="si-dur">${ev.duree||calcDur(ev.debut,ev.fin)||'—'}</div>
      <button class="btn-edit" onclick="openEditStop('${esc(String(key))}')">✏</button>
    </div>`;
  });
  if(_curStopKey){
    const lbl=_curStopKey==='_pause'?'Pause':getEvtLabel(_curStopKey);
    const col=_curStopKey==='_pause'?STOP_COL['_pause']:STOP_COL.pb;
    html+=`<div class="si" style="animation:blink .85s step-start infinite">
      <div class="sdot" style="background:${col}"></div>
      <div class="si-nm" style="font-size:calc(9px*var(--zf,1))">${lbl}</div>
      <div class="si-dur" style="font-size:calc(9px*var(--zf,1))">…</div>
    </div>`;
  }
  c.innerHTML=html;
}

function calcDur(d,f){
  if(!d||!f) return '';
  try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
  const diff=p(f)-p(d);return diff>0?fmtDur(diff):'';}catch(e){return '';}
}

// ── CODE INPUT OVERLAY ──
let _codeInputTarget = null;
let _csDigits = ['','','','','','','','',''];
let _csFormat = '6_3'; // '9' for 9-digit straight, '6_3' for 6+separator+3
function _csRender(){
  let allFilled=true;
  for(let i=0;i<9;i++){
    const el=document.getElementById('cs-'+i);if(!el) continue;
    const d=_csDigits[i]||'';
    el.textContent=d||'X';
    el.className='cs'+(d?' cs-filled':'');
    if(!d) allFilled=false;
  }
  if(allFilled){for(let i=0;i<9;i++){const e=document.getElementById('cs-'+i);if(e)e.classList.add('cs-done');}}
  else{
    const ai=_csDigits.findIndex(d=>!d);
    const activeEl=document.getElementById('cs-'+ai);
    if(activeEl) activeEl.classList.add('cs-active');
  }
}
function _csKeydown(e){
  if(/^[0-9]$/.test(e.key)){
    e.preventDefault();
    const pos=_csDigits.findIndex(d=>!d);
    if(pos!==-1){_csDigits[pos]=e.key;_csRender();}
  } else if(e.key==='Backspace'){
    e.preventDefault();
    let pos=_csDigits.findIndex(d=>!d);
    if(pos===-1) pos=9;
    if(pos>0){_csDigits[pos-1]='';_csRender();}
  } else if(e.key==='Enter'){
    e.preventDefault();_codeInputConfirm();
  } else if(e.key==='Escape'){
    e.preventDefault();closeM('m-code-input');
  }
}
function openCodeInput(fieldId, label, fmt) {
  _codeInputTarget = fieldId;
  _csFormat = fmt || '6_3';
  document.getElementById('code-input-lbl').textContent = label || 'Code';
  const raw = ((document.getElementById('f-' + fieldId) || {}).value || '').replace(/_/g,'').replace(/[^0-9]/g,'');
  _csDigits = Array(9).fill('').map((_,i)=>raw[i]||'');
  const sep=document.getElementById('cs-sep');
  const hint=document.getElementById('cs-hint');
  if(sep) sep.style.display=_csFormat==='9'?'none':'';
  if(hint) hint.textContent=_csFormat==='9'?'Tapez les 9 chiffres du N° OF':'Tapez les 6 premiers puis les 3 derniers chiffres';
  _csRender();
  openM('m-code-input');
  setTimeout(()=>{const s=document.getElementById('code-slots');if(s)s.focus();},80);
}
function _codeInputConfirm() {
  const all=_csDigits.join('');
  if(_csFormat==='9'){
    if(!/^[0-9]{9}$/.test(all)){toast('Format requis : 9 chiffres (ex : 123456789)','err');return;}
    const field=document.getElementById('f-'+_codeInputTarget);
    if(field){field.value=all;scheduleAutoSave();}
  } else {
    const d6=_csDigits.slice(0,6).join('');
    const d3=_csDigits.slice(6,9).join('');
    const v=d6+'_'+d3;
    if(!/^[0-9]{6}_[0-9]{3}$/.test(v)){toast('Format requis : 6 chiffres_3 chiffres (ex : 123456_789)','err');return;}
    const field=document.getElementById('f-'+_codeInputTarget);
    if(field){field.value=v;scheduleAutoSave();}
  }
  closeM('m-code-input');
}

// ── OF DETAIL REPORT ──
window._rptProdRows = [];
window._rptEvtRows = [];

// Shared OF/arrêt detail modal renderer
function _renderAndOpenOfDetail(r, ofEvts) {
  const isProd = r._rowType !== 'evt';
  const tc = r.trs>=90?'#16a34a':r.trs>=70?'#f59e0b':r.trs>=0?'#dc2626':'#94a3b8';
  const kitStr = (r.kit||'').toLowerCase();
  const infoRows = isProd ? [
    ['Heure début', r.debut||'—', '#374151'],
    ['Heure fin', r.fin||'—', '#374151'],
    ['Durée', r.duree||'—', '#059669'],
    ['Taille', r.taille||'—', '#374151'],
    ['Type produit', r.type_prod||'—', '#374151'],
    ['Lots de 2', kitStr==='oui'?'✓ Oui':'Non', kitStr==='oui'?'#16a34a':'#94a3b8'],
    ['Qté fabriquée', r.qte_fab||'—', '#1e3a8a'],
    ['Qté emballée', r.qte_emb||'—', '#374151'],
    ['Équivalence', r.equiv||'—', '#0891b2'],
    ...(r.date?[['Date', r.date, '#374151']]:[]),
    ...(r.poste?[['Poste', r.poste, '#374151']]:[]),
    ...(r.pilote?[['Pilote', r.pilote, '#374151']]:[]),
    ...(r.fibre?[['Fibre', r.fibre, '#6366f1']]:[]),
    ...(r.code_prod?[['Code Produit', r.code_prod, '#374151']]:[]),
    ...(r.ref_taie?[['Réf Taie', r.ref_taie, '#374151']]:[]),
    ...(r.nb_pers?[['Nb Personnes', r.nb_pers, '#374151']]:[]),
    ...(r.copilote?[['Co-Pilote', r.copilote, '#374151']]:[]),
    ...(r.traca?[['Traca', r.traca, '#374151']]:[]),
    ...(r.poids?[['Poids (g)', r.poids, '#374151']]:[]),
  ] : [
    ['Type arrêt', r.type||'—', '#dc2626'],
    ['Heure début', r.debut||'—', '#374151'],
    ['Heure fin', r.fin||'—', '#374151'],
    ['Durée', r.duree||'—', '#059669'],
    ...(r.date?[['Date', r.date, '#374151']]:[]),
    ...(r.poste?[['Poste', r.poste, '#374151']]:[]),
    ...(r.pilote?[['Pilote', r.pilote, '#374151']]:[]),
  ];
  const infoHtml = infoRows.map(([lbl,val,col])=>`
    <div style="display:flex;justify-content:space-between;align-items:baseline;padding:5px 0;border-bottom:1px solid #f1f5f9">
      <span style="font-size:calc(10px*var(--zf,1));font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.04em">${lbl}</span>
      <span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:${col};text-align:right;max-width:55%">${esc(String(val))}</span>
    </div>`).join('');
  const commentHtml = r.comment?`<div style="background:#fffbeb;border-left:3px solid #fbbf24;padding:6px 10px;margin:8px 0;font-size:calc(11px*var(--zf,1));color:#92400e;border-radius:0 6px 6px 0">💬 ${esc(r.comment)}</div>`:'';
  const evtsTableHtml = `<table style="width:100%;border-collapse:collapse">
    <thead><tr style="background:#f8fafc">
      <th style="padding:5px 8px;text-align:left;font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;border-bottom:2px solid #e2e8f0">Type</th>
      <th style="padding:5px 8px;text-align:left;font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;border-bottom:2px solid #e2e8f0">Plage</th>
      <th style="padding:5px 8px;text-align:left;font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#64748b;border-bottom:2px solid #e2e8f0">Durée</th>
    </tr></thead>
    <tbody>${ofEvts.length?ofEvts.map(e=>`<tr style="border-bottom:1px solid #f1f5f9">
      <td style="padding:5px 8px;font-weight:600;color:#374151">${esc(e.type||'')}</td>
      <td style="padding:5px 8px;color:#64748b;white-space:nowrap">${esc(e.debut||'')} → ${esc(e.fin||'')}</td>
      <td style="padding:5px 8px;font-weight:700;color:#dc2626">${esc(e.duree||'')}</td>
    </tr>`).join(''):'<tr><td colspan="3" style="padding:12px 8px;text-align:center;color:#94a3b8">Aucun arrêt</td></tr>'}
    </tbody></table>`;
  const headerTitle = isProd ? (r.of||'—') : (r.type||'—');
  const headerSub = isProd ? 'Détail OF' : 'Détail Arrêt';
  const rightHtml = isProd
    ? `<div style="text-align:right">
        <div style="font-size:calc(26px*var(--zf,1));font-weight:900;color:${tc};line-height:1">${r.trs>=0?r.trs.toFixed(1)+'%':'—'}</div>
        <div style="font-size:calc(9px*var(--zf,1));opacity:.55;text-transform:uppercase;letter-spacing:.08em">TRS</div>
       </div>`
    : `<div style="text-align:right">
        <div style="font-size:calc(18px*var(--zf,1));font-weight:900;color:#dc2626;line-height:1">${esc(r.duree||'—')}</div>
        <div style="font-size:calc(9px*var(--zf,1));opacity:.55;text-transform:uppercase;letter-spacing:.08em">Durée</div>
       </div>`;
  document.getElementById('of-detail-content').innerHTML = `
    <div style="background:var(--navy);color:#fff;padding:14px 20px;border-radius:14px 14px 0 0;display:flex;align-items:center;justify-content:space-between;flex-shrink:0">
      <div>
        <div style="font-size:calc(9px*var(--zf,1));text-transform:uppercase;letter-spacing:.1em;opacity:.55;margin-bottom:2px">${headerSub}</div>
        <div style="font-size:calc(20px*var(--zf,1));font-weight:900;font-family:monospace;letter-spacing:.05em">${esc(headerTitle)}</div>
      </div>
      <div style="display:flex;align-items:center;gap:16px">
        ${rightHtml}
        <button onclick="closeM('m-of-detail')" style="background:rgba(255,255,255,.15);border:none;color:#fff;border-radius:8px;width:34px;height:34px;font-size:calc(16px*var(--zf,1));cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0">✕</button>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;min-height:0;flex:1;overflow:hidden">
      <div style="padding:14px 16px;border-right:1px solid #e2e8f0;overflow-y:auto">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#94a3b8;letter-spacing:.08em;margin-bottom:6px">Informations</div>
        ${infoHtml}
        ${commentHtml}
      </div>
      <div style="padding:14px 16px;overflow-y:auto">
        <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#94a3b8;letter-spacing:.08em;margin-bottom:6px">${isProd?'Arrêts pendant cet OF'+(ofEvts.length?' ('+ofEvts.length+')':''):'—'}</div>
        ${isProd?evtsTableHtml:'<div style="padding:20px 0;text-align:center;color:#94a3b8;font-size:calc(11px*var(--zf,1))">Aucun arrêt associé</div>'}
      </div>
    </div>`;
  openM('m-of-detail');
}

function showOfDetail(ri) {
  const r = window._rptProdRows[ri];
  const evtRows = window._rptEvtRows || [];
  if(!r) return;
  function _hm2s(hm) { if(!hm) return 0; const p=(hm+':0:0').split(':').map(Number); return(p[0]||0)*3600+(p[1]||0)*60+(p[2]||0); }
  const debS = _hm2s(r.debut), finS = _hm2s(r.fin)||86400;
  const ofEvts = evtRows.filter(e => { const t = _hm2s(e.debut); return t >= debS && t <= finS; });
  _renderAndOpenOfDetail(r, ofEvts);
}

function showMainRowDetail(key) {
  const r = window._rowMap && window._rowMap[String(key)];
  if(!r) return;
  const isProd = r._rowType === 'prod';
  let ofEvts = [];
  if(isProd) {
    const _hm2s = hm=>{if(!hm)return 0;const[h,m]=(hm+':00').split(':').map(Number);return(h||0)*3600+(m||0)*60;};
    const debS=_hm2s(r.debut), finS=_hm2s(r.fin)||86400;
    ofEvts=(window._mainEvtsAll||[]).filter(e=>{const t=_hm2s(e.debut);return t>=debS&&t<=finS;});
  }
  _renderAndOpenOfDetail(r, ofEvts);
}


// ── ARRÊTS MANQUANTS ──
let _missingDeclChecked = false;
let _ecartChecked = false;

function _hasDeclaredType(keywords){
  // Cherche dans les déclarations du poste actuel (accueil)
  const rows=window._lastMainRows||[];
  return rows.some(r=>{
    const t=String(r.type||r._type||r[0]||'').toLowerCase();
    return keywords.some(k=>t.includes(k));
  });
}

async function checkMissingDecls(){
  // Refresh des déclarations avant check
  await loadMainDecl();
  const ap=_cfgArretsPrevus||{};
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const model=_cfgModels&&_cfgModels.find(m=>m.nom===(ST.poste||''));
  const jour=model&&model.jours&&model.jours[dk];
  const missing=[];
  if((ap.pause_min||0)>0 && !_hasDeclaredType(['pause'])){
    missing.push({label:'Pause',budget:ap.pause_min||0,keywords:['pause'],id:'md-pause'});
  }
  if((ap.clean_short_min||0)>0 && !_hasDeclaredType(['nettoyage','nett'])){
    missing.push({label:'Nettoyage court',budget:ap.clean_short_min||0,keywords:['nettoyage','nett'],id:'md-nett'});
  }
  if((ap.meeting_tol_min||0)>0 && !_hasDeclaredType(['réunion','reunion','meeting'])){
    missing.push({label:'Réunion',budget:ap.meeting_tol_min||0,keywords:['réunion','reunion','meeting'],id:'md-meet'});
  }
  if(!missing.length) return false;
  // Construire les lignes du modal
  const lastFin=_lastProdDeclTime;
  const container=document.getElementById('md-rows');
  container.innerHTML='';
  missing.forEach(item=>{
    // Heure de début suggérée = fin dernière prod ou fin modèle - budget
    let sugStart='',sugEnd='';
    if(jour&&jour.fin){
      const[fh,fm]=jour.fin.split(':').map(Number);
      const budMin=item.budget;
      const endM=fh*60+fm;
      const startM=endM-budMin;
      const sh=Math.floor(((startM%1440)+1440)%1440/60),sm=((startM%1440)+1440)%1440%60;
      sugStart=String(sh).padStart(2,'0')+':'+String(sm).padStart(2,'0');
      sugEnd=jour.fin;
    } else if(lastFin){
      sugStart=String(lastFin.getHours()).padStart(2,'0')+':'+String(lastFin.getMinutes()).padStart(2,'0');
      const e=new Date(lastFin.getTime()+item.budget*60000);
      sugEnd=String(e.getHours()).padStart(2,'0')+':'+String(e.getMinutes()).padStart(2,'0');
    }
    const div=document.createElement('div');
    div.style.cssText='background:#fffbeb;border:1.5px solid #fde68a;border-radius:8px;padding:10px 12px';
    div.innerHTML=`<div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:#92400e;margin-bottom:6px">${esc(item.label)} <span style="font-weight:400;color:#a16207">(prévu ${item.budget} min)</span></div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <label style="font-size:calc(11px*var(--zf,1));font-weight:600">De <input type="time" id="${item.id}-debut" value="${sugStart}" style="padding:4px 6px;border:1.5px solid #fde68a;border-radius:5px;font-size:calc(12px*var(--zf,1));font-weight:700;margin-left:4px"></label>
        <label style="font-size:calc(11px*var(--zf,1));font-weight:600">à <input type="time" id="${item.id}-fin" value="${sugEnd}" style="padding:4px 6px;border:1.5px solid #fde68a;border-radius:5px;font-size:calc(12px*var(--zf,1));font-weight:700;margin-left:4px"></label>
        <button onclick="addMissingDecl('${item.id}','${esc(item.label)}')" style="background:#f59e0b;color:#fff;border:none;border-radius:6px;padding:5px 12px;font-size:calc(11px*var(--zf,1));font-weight:700;cursor:pointer">+ Ajouter</button>
        <span id="${item.id}-ok" style="display:none;color:#16a34a;font-weight:700;font-size:calc(12px*var(--zf,1))">✓ Ajouté</span>
      </div>`;
    container.appendChild(div);
  });
  openM('m-missing-decl');
  return true;
}

async function addMissingDecl(rowId, label){
  const debut=document.getElementById(rowId+'-debut')?.value;
  const fin=document.getElementById(rowId+'-fin')?.value;
  if(!debut||!fin){toast('Remplissez les heures','warn');return;}
  const r=await fetch('/api/add_stop_decl',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({type:label,debut_hms:debut,fin_hms:fin,comment:'Déclaré rétroactivement'})});
  const d=r?await r.json():{};
  if(d.ok){
    const ok=document.getElementById(rowId+'-ok');if(ok) ok.style.display='';
    // Griser la ligne
    const btn=document.querySelector(`button[onclick="addMissingDecl('${rowId}','${label.replace(/'/g,"\\'")}')"]`);
    if(btn){btn.disabled=true;btn.style.opacity='.4';}
    toast(label+' ajouté','ok');
  } else {
    toast('Erreur : '+(d.error||'?'),'err');
  }
}

async function skipMissingDecl(){
  closeM('m-missing-decl');
  _missingDeclChecked=true;
  await doFinPoste();
}

// ── FIN DE POSTE ──
async function doFinPoste(){
  if(ST.prod_active){toast('Terminer la production en cours avant de finir le poste','err');return;}
  const fpd=await apiFetch('/api/fin_poste_data');
  if(fpd){
    window._ecartFpData=fpd;
    const hasIssue=(fpd.gap_intervals&&fpd.gap_intervals.length>0)||((fpd.overflow_s||0)>60)||((fpd.ecart_s||0)>60);
    if(hasIssue){_showEcartModal(fpd);}else{goTab('finposte');}
  } else goTab('finposte');
}

async function _doGoFinPoste(){
  if(!_ecartChecked){
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd&&(fpd.ecart_s||0)>60){
      window._ecartFpData=fpd;
      _showEcartModal(fpd);
      return;
    }
  }
  _ecartChecked=false;
  goTab('finposte');
}

function _showEcartModal(fpd){
  const ecart_min=Math.round((fpd.ecart_s||0)/60);
  const model_min=Math.round((fpd.model_dur_s||0)/60);
  const prod_min=Math.round((fpd.tot_s||0)/60);
  const stop_min=Math.max(0,model_min-prod_min-ecart_min);
  const modelDebut=fpd.model_debut||'';
  const modelFin=fpd.model_fin||'';
  const overflow_min=Math.round(fpd.overflow_min||0);
  const gaps=fpd.gap_intervals||[];
  const hasGaps=gaps.length>0;
  // Réinitialiser le panneau plage horaire au style normal avant d'appliquer la surcharge
  const _pP=document.getElementById('ecart-plage-panel');
  if(_pP){
    _pP.style.cssText='margin-bottom:14px';
    const _pD=_pP.querySelectorAll('div');
    if(_pD[0]) _pD[0].style.color='#0369a1';
    if(_pD[1]){_pD[1].style.background='#f0f9ff';_pD[1].style.borderColor='#bae6fd';}
  }
  // Guide header — basé sur les gaps réels, pas sur ecart_min
  const guidEl=document.getElementById('ecart-guide');
  if(guidEl){
    if(overflow_min>0){
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#fef2f2;border:1px solid #fca5a5;line-height:1.5';
      guidEl.innerHTML='<span style="color:#dc2626;font-weight:800;font-size:calc(13px*var(--zf,1))">⚠ Dépassement de plage : +'+overflow_min+' min au-delà de '+esc(modelFin)+'</span><br>'+
        '<span style="color:#7f1d1d">Un ou plusieurs OFs se terminent après la fin de la plage modèle. Modifiez la plage ci-dessous.</span>';
      // Mettre le panneau de modification de plage en rouge pour signaler l'urgence
      const plagePanel=document.getElementById('ecart-plage-panel');
      if(plagePanel){
        plagePanel.style.cssText='margin-bottom:14px;border-radius:10px;border:2px solid #dc2626;padding:10px;background:#fef2f2';
        const divs=plagePanel.querySelectorAll('div');
        if(divs[0]) divs[0].style.color='#dc2626'; // label titre
        if(divs[1]){ divs[1].style.background='#fef2f2'; divs[1].style.borderColor='#fca5a5'; }
      }
    } else if(!hasGaps){
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#f0fdf4;border:1px solid #bbf7d0;line-height:1.5';
      guidEl.innerHTML='<span style="color:#16a34a;font-weight:800;font-size:calc(13px*var(--zf,1))">✓ Toute la plage '+esc(modelDebut)+' → '+esc(modelFin)+' est couverte !</span>';
    } else {
      const gapTotal=gaps.reduce((a,g)=>a+(g.duree_min||0),0);
      guidEl.style.cssText='font-size:calc(12px*var(--zf,1));margin-bottom:10px;padding:8px 12px;border-radius:6px;background:#fff7ed;border:1px solid #fed7aa;line-height:1.5';
      guidEl.innerHTML='Objectif : couvrir <b>'+esc(modelDebut)+' → '+esc(modelFin)+'</b> ('+model_min+' min).<br>'+
        '<span style="color:#dc2626;font-weight:700">'+gapTotal+' min de plages non couvertes ('+gaps.length+' écart'+(gaps.length>1?'s':'')+').</span> '+
        '<span style="color:var(--gray)">Déclarez les arrêts manquants ou corrigez les horaires des OFs.</span>';
    }
  }
  // Stats
  document.getElementById('ecart-info').innerHTML=
    '<div style="display:grid;grid-template-columns:auto 1fr auto 1fr;gap:3px 16px;font-size:calc(12px*var(--zf,1))">'+
    '<span style="color:var(--gray)">Durée modèle :</span><b>'+model_min+' min</b>'+
    '<span style="color:var(--gray)">Prod déclarée :</span><b>'+prod_min+' min</b>'+
    '<span style="color:var(--gray)">Arrêts déclarés :</span><b>'+stop_min+' min</b>'+
    '<span style="color:#dc2626;font-weight:700">Écart :</span><b style="color:#dc2626;font-weight:900">'+ecart_min+' min</b>'+
    '</div>';
  // Prefill plage horaire du poste
  const pdebut=document.getElementById('ecart-plage-debut');
  const pfin=document.getElementById('ecart-plage-fin');
  if(pdebut)pdebut.value=fpd.model_debut||'';
  if(pfin)pfin.value=fpd.model_fin||'';
  // Gap intervals
  const gapsEl=document.getElementById('ecart-gaps');
  if(gapsEl){
    if(!hasGaps){
      // Ne pas afficher "Aucune plage non couverte" quand il y a un dépassement
      gapsEl.innerHTML=overflow_min>0?''
        :'<div style="color:#16a34a;font-size:calc(12px*var(--zf,1));font-weight:700;padding:4px 0">✓ Aucune plage non couverte</div>';
    } else {
      gapsEl.innerHTML='';
      gaps.forEach((g,gi)=>{
        const div=document.createElement('div');
        div.style.cssText='background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:8px 10px;margin-bottom:6px';
        div.innerHTML=
          '<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:6px">'+
            '<span style="font-size:calc(12px*var(--zf,1));font-weight:700;color:#dc2626">⚠ '+esc(g.debut)+' → '+esc(g.fin)+
              ' <span style="font-weight:400;color:#9f1239">('+g.duree_min+' min)</span></span>'+
            '<button class="btn btn-ghost" style="font-size:calc(11px*var(--zf,1));padding:3px 10px;color:#dc2626;border-color:#fca5a5" onclick="ecartToggleGapForm('+gi+')">+ Déclarer un arrêt</button>'+
          '</div>'+
          '<div id="ecart-gap-form-'+gi+'" style="display:none;margin-top:8px;border-top:1px solid #fca5a5;padding-top:8px">'+
            '<div style="display:flex;gap:6px;align-items:flex-end;flex-wrap:wrap">'+
              '<div><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Début</label>'+
                '<input type="time" id="ecart-gap-debut-'+gi+'" value="'+esc(g.debut)+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px"></div>'+
              '<div><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Fin</label>'+
                '<input type="time" id="ecart-gap-fin-'+gi+'" value="'+esc(g.fin)+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px"></div>'+
              '<div style="flex:1;min-width:160px"><label style="font-size:calc(10px*var(--zf,1));color:var(--gray);display:block;margin-bottom:2px">Type d\'arrêt</label>'+
                _buildEcartStopSelect(gi)+'</div>'+
              '<button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:5px 12px" onclick="saveEcartGapStop('+gi+')">✓ Ajouter</button>'+
            '</div>'+
          '</div>';
        gapsEl.appendChild(div);
      });
    }
  }
  // OFs list
  const list=document.getElementById('ecart-of-list');
  list.innerHTML='';
  (fpd.of_list||[]).forEach(of=>{
    const div=document.createElement('div');
    div.style.cssText='display:flex;align-items:center;gap:8px;padding:8px;background:#f8fafc;border-radius:6px;border:1px solid var(--border);flex-wrap:wrap';
    div.innerHTML='<span style="flex:1;font-size:calc(12px*var(--zf,1));font-weight:700;min-width:80px">'+(of.of||'OF')+
      ' <span style="font-weight:400;color:var(--gray);font-size:calc(11px*var(--zf,1))">'+esc(of.taille||'')+'</span> '+
      '<span style="font-size:calc(10px*var(--zf,1));color:#64748b">'+esc(of.debut||'')+'→'+esc(of.fin||'')+'</span></span>'+
      '<div style="display:flex;align-items:center;gap:4px">'+
      '<input type="time" class="ecart-debut" value="'+(of.debut||'')+'" data-oldebut="'+(of.debut||'')+'" data-ofnum="'+(of.of||'')+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px">'+
      '<span style="color:var(--gray)">→</span>'+
      '<input type="time" class="ecart-fin" value="'+(of.fin||'')+'" style="padding:4px 6px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));width:90px">'+
      '<button class="btn btn-prim" style="font-size:calc(11px*var(--zf,1));padding:3px 10px" onclick="saveEcartOf(this)">✓</button>'+
      '</div>';
    list.appendChild(div);
  });
  // Griser "Valider et terminer" tant que tout n'est pas résolu
  const valBtn=document.getElementById('ecart-valider-btn');
  if(valBtn){
    const locked=hasGaps||overflow_min>0;
    valBtn.disabled=locked;
    valBtn.style.opacity=locked?'0.4':'1';
    valBtn.style.cursor=locked?'not-allowed':'pointer';
    valBtn.title=locked?(hasGaps?'Justifiez toutes les plages non couvertes avant de terminer':'Résolvez le dépassement de plage avant de terminer'):'';
  }
  openM('m-ecart-poste');
}

function ecartToggleGapForm(gi){
  const f=document.getElementById('ecart-gap-form-'+gi);
  if(f) f.style.display=f.style.display==='none'?'block':'none';
}

async function saveEcartGapStop(gi){
  const debut=(document.getElementById('ecart-gap-debut-'+gi)||{}).value||'';
  const fin=(document.getElementById('ecart-gap-fin-'+gi)||{}).value||'';
  const type=((document.getElementById('ecart-gap-type-'+gi)||{}).value||'').trim();
  if(!debut||!fin||!type){toast('Renseigner début, fin et type d\'arrêt','err');return;}
  const r=await fetch('/api/add_stop_decl',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({type,debut_hms:debut,fin_hms:fin,comment:'Déclaré depuis réconciliation fin de poste'})});
  const d=r?await r.json():{};
  if(d.ok){
    toast(type+' ajouté','ok');
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
  } else {
    toast('Erreur : '+(d.error||'?'),'err');
  }
}

function _buildEcartStopSelect(gi){
  // Rassemble tous les types d'arrêt depuis paramètres
  const pauses=new Set(['Pause']);
  const nettoyage=new Set(['Nettoyage court','Nettoyage long','Nettoyage très long']);
  const pannes=new Set();
  const organisation=new Set();
  (_evtsList||[]).forEach(e=>{
    const lbl=e.label||'';if(!lbl)return;
    if(e.cat==='pb'||e.cat==='ratt')pannes.add(lbl);
    else if(e.cat==='nettoyage')nettoyage.add(lbl);
    else if(e.cat==='organisation')organisation.add(lbl);
    else if(e.cat==='autre')organisation.add(lbl);
  });
  const mkGrp=(name,items)=>{
    const arr=[...items];if(!arr.length)return'';
    return`<optgroup label="${esc(name)}">`+arr.map(l=>`<option value="${esc(l)}">${esc(l)}</option>`).join('')+'</optgroup>';
  };
  const html='<option value="">-- Choisir un type d\'arrêt --</option>'+
    mkGrp('⏸ Pauses',pauses)+
    mkGrp('🧹 Nettoyage',nettoyage)+
    mkGrp('🔴 Pannes / Rattrapages',pannes)+
    mkGrp('🔵 Organisation',organisation)+
    '<optgroup label="⚫ Autre"><option value="Autre (à justifier)">Autre (à justifier)</option></optgroup>';
  return`<select id="ecart-gap-type-${gi}" onchange="ecartCheckAutreType(this,${gi})" style="width:100%;padding:4px 8px;border:1.5px solid var(--border);border-radius:5px;font-size:calc(12px*var(--zf,1));background:#fff">${html}</select>`;
}

let _ecartGapAutreIdx=-1;
function ecartCheckAutreType(sel,gi){
  if(sel.value==='Autre (à justifier)'){
    _ecartGapAutreIdx=gi;
    const inp=document.getElementById('ecart-autre-label');if(inp)inp.value='';
    sel.value='';
    openM('m-ecart-autre');
  }
}

function confirmEcartAutre(){
  const inp=document.getElementById('ecart-autre-label');
  const lbl=(inp&&inp.value||'').trim();
  if(!lbl){toast('Saisir un libellé','err');return;}
  const gi=_ecartGapAutreIdx;
  const sel=document.getElementById('ecart-gap-type-'+gi);
  if(sel){
    let opt=[...sel.options].find(o=>o.value===lbl);
    if(!opt){opt=document.createElement('option');opt.value=lbl;opt.textContent=lbl;sel.appendChild(opt);}
    sel.value=lbl;
  }
  closeM('m-ecart-autre');
}

async function saveEcartPlageHoraire(){
  const debut=(document.getElementById('ecart-plage-debut')||{}).value||'';
  const fin=(document.getElementById('ecart-plage-fin')||{}).value||'';
  if(!debut||!fin){toast('Renseigner début et fin','err');return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const poste=ST.poste||'';
  if(!poste){toast('Poste non défini','err');return;}
  const mi=_cfgModels.findIndex(m=>m.nom===poste);
  if(mi>=0){
    if(!_cfgModels[mi].jours)_cfgModels[mi].jours={};
    if(!_cfgModels[mi].jours[dk])_cfgModels[mi].jours[dk]={};
    _cfgModels[mi].jours[dk].debut=debut;
    _cfgModels[mi].jours[dk].fin=fin;
  }
  await fetch('/api/update_model_today',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nom:poste,day_key:dk,debut,fin})});
  // Mettre à jour shift_debut_dt / shift_fin_dt → écrit aussi dans Excel POSTES P/Q
  const _today=new Date();
  const [_dh,_dm]=(debut||'00:00').split(':').map(Number);
  const [_fh,_fm]=(fin||'00:00').split(':').map(Number);
  const _debDt=new Date(_today.getFullYear(),_today.getMonth(),_today.getDate(),_dh,_dm,0);
  let _finDt=new Date(_today.getFullYear(),_today.getMonth(),_today.getDate(),_fh,_fm,0);
  if(_finDt<=_debDt) _finDt=new Date(_finDt.getTime()+86400000);
  await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:_debDt.toISOString(),fin_iso:_finDt.toISOString()})});
  toast('Plage mise à jour','ok');
  const fpd=await apiFetch('/api/fin_poste_data');
  if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
}

async function skipEcartPoste(){
  closeM('m-ecart-poste');
  _ecartChecked=true;
  await _doGoFinPoste();
}

function ecartOpenOfPanel(){
  // OF panel is always visible now - kept for backward compat
}

async function saveEcartOf(btn){
  const row=btn.closest('div');
  const debutIn=row.querySelector('.ecart-debut');
  const finIn=row.querySelector('.ecart-fin');
  const ofNum=debutIn?debutIn.dataset.ofnum:'';
  const oldDebut=debutIn?debutIn.dataset.oldebut:'';
  const newDebut=debutIn?debutIn.value:'';
  const newFin=finIn?finIn.value:'';
  if(!newDebut||!newFin){toast('Renseigner début et fin','err');return;}
  const r=await apiFetch('/api/update_of_time',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({of_num:ofNum,old_debut:oldDebut,new_debut:newDebut,new_fin:newFin})});
  if(r&&r.ok){
    toast('OF mis à jour','ok');
    if(debutIn) debutIn.dataset.oldebut=newDebut;
    // Recalculer l'écart
    const fpd=await apiFetch('/api/fin_poste_data');
    if(fpd){window._ecartFpData=fpd;_showEcartModal(fpd);}
  } else {
    toast('Erreur mise à jour','err');
  }
}

async function loadFPData(){
  const d=await apiFetch('/api/fin_poste_data');
  if(!d) return;
  document.getElementById('fp-trs').textContent=fmtTRS(d.trs_shift!==undefined?d.trs_shift:d.trs);
  document.getElementById('fp-trs-of').textContent=fmtTRS(d.trs);
  document.getElementById('fp-eq').textContent=Math.round(d.tot_equiv||0);
  document.getElementById('fp-nof').textContent=d.nb_of||0;
  // fp-prod-t sera mis à jour après le calcul de stopTotal ci-dessous
  const fpOuvEl=document.getElementById('fp-ouverture');
  if(fpOuvEl) fpOuvEl.textContent=(d.model_debut&&d.model_fin)?(d.model_debut+'→'+d.model_fin):(Math.round((d.model_dur_s||0)/60)+' min');
  const fpDedEl=document.getElementById('fp-ded');
  if(fpDedEl) fpDedEl.textContent=Math.round((d.planned_ded_s||0)/60)+' min';
  document.getElementById('fp-who').textContent=`${d.pilot||ST.pilot||''} — ${ST.poste||''}`;
  const now2=new Date();
  const fpDateEl=document.getElementById('fp-date');
  if(fpDateEl) fpDateEl.textContent=String(now2.getDate()).padStart(2,'0')+'/'+String(now2.getMonth()+1).padStart(2,'0')+'/'+now2.getFullYear();
  // Populate model selector
  const fpSel=document.getElementById('fp-model-sel');
  if(fpSel&&_cfgModels&&_cfgModels.length){
    fpSel.innerHTML='<option value="">— Modèle —</option>';
    _cfgModels.forEach(m=>{const o=document.createElement('option');o.value=m.nom||'';o.textContent=m.nom||'';if(m.nom===ST.poste)o.selected=true;fpSel.appendChild(o);});
  }
  window._fpData=d;

  // Count stops
  const stops=gEvts.filter(e=>e.type);
  const _hmsS=hms=>{if(!hms)return 0;const p=(hms+':0:0').split(':').map(Number);return(p[0]||0)*3600+(p[1]||0)*60+(p[2]||0);};
  const stopTotal=stops.reduce((a,e)=>a+Math.max(0,_hmsS(e.fin)-_hmsS(e.debut)),0);
  document.getElementById('fp-stop-t').textContent=fmtDurMS(stopTotal);
  // Net prod = sum of each OF's time minus overlapping stops within that period
  const _ofPs=(d.of_list||[]).map(p=>({s:_hmsS(p.debut),e:_hmsS(p.fin)})).filter(p=>p.e>p.s);
  const _stPs=stops.map(e=>({s:_hmsS(e.debut),e:_hmsS(e.fin)})).filter(e=>e.e>e.s);
  let _netProdS=0;
  _ofPs.forEach(pp=>{
    const ov=_stPs.map(sv=>({s:Math.max(sv.s,pp.s),e:Math.min(sv.e,pp.e)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    _netProdS+=Math.max(0,pp.e-pp.s-mg.reduce((a,o)=>a+(o.e-o.s),0));
  });
  document.getElementById('fp-prod-t').textContent=Math.round(_netProdS/60)+' min';

  // New metrics: Cadence/h, Nb pièces, Chg. fibre
  const ofList=d.of_list||[];
  const totQteFabFP=ofList.reduce((s,r)=>s+parseFloat(r.qte_fab||0),0);
  const elapsedEffSFP=Math.max(1,(d.model_dur_s||0)-(d.planned_ded_s||0));
  const cadenceHFP=elapsedEffSFP>0?Math.round(totQteFabFP/elapsedEffSFP*3600):0;
  const sortedProdFFP=ofList.filter(r=>r.fibre).sort((a,b)=>(a.debut||'').localeCompare(b.debut||''));
  let nbChangFibreFP=0;for(let i=1;i<sortedProdFFP.length;i++){if(sortedProdFFP[i].fibre!==sortedProdFFP[i-1].fibre)nbChangFibreFP++;}
  document.getElementById('fp-cadence').textContent=cadenceHFP;
  document.getElementById('fp-pieces').textContent=Math.round(totQteFabFP);
  document.getElementById('fp-chg-fibre').textContent=nbChangFibreFP;

  // Timeline — refresh events first, then combine historical + live
  await pollEvts();
  const shiftStart=d.shift_start_iso||new Date(Date.now()-8*3600*1000).toISOString();
  // Use gEvts (historical) + convert live tl_events to display format
  const liveEvts=tlEventsToDisplayFmt(ST.tl_events||[]);
  const allEvtsForTL=[...gEvts,...liveEvts];
  // Pre-fill datetime-local inputs — use model horaire theoretical times when available
  const toLocalDT=dt=>{const y=dt.getFullYear(),mo=String(dt.getMonth()+1).padStart(2,'0'),dy=String(dt.getDate()).padStart(2,'0'),h=String(dt.getHours()).padStart(2,'0'),mi=String(dt.getMinutes()).padStart(2,'0');return `${y}-${mo}-${dy}T${h}:${mi}`;};
  const dayKeys2=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk2=dayKeys2[new Date().getDay()];
  const modelFP=_cfgModels&&_cfgModels.find(m=>m.nom===ST.poste);
  const jourFP=modelFP&&modelFP.jours&&modelFP.jours[dk2];
  // Debut: prefer model debut, fallback to shift_start_iso
  let debDt=new Date(shiftStart);
  if(jourFP&&jourFP.debut){const[hh,mm]=jourFP.debut.split(':').map(Number);debDt=new Date();debDt.setHours(hh,mm,0,0);}
  // Fin: prefer model fin, fallback to now
  let finDt=new Date();
  if(jourFP&&jourFP.fin){const[hh,mm]=jourFP.fin.split(':').map(Number);finDt=new Date();finDt.setHours(hh,mm,0,0);}
  const fpDeb=document.getElementById('fp-debut-dt');
  const fpFin=document.getElementById('fp-fin-dt');
  if(fpDeb) fpDeb.value=toLocalDT(debDt);
  if(fpFin) fpFin.value=toLocalDT(finDt);
  // Show model horaire info (in header and gauge footer)
  const fpShiftInfo=document.getElementById('fp-shift-info');
  const fpShiftH=document.getElementById('fp-shift-hours');
  const fpHoraireDisp=document.getElementById('fp-horaire-display');
  const hStr2=jourFP&&jourFP.debut&&jourFP.fin?`${jourFP.debut} → ${jourFP.fin}`:'—';
  if(fpShiftInfo) fpShiftInfo.textContent=hStr2;
  if(fpShiftH) fpShiftH.textContent=hStr2;
  if(fpHoraireDisp) fpHoraireDisp.textContent=hStr2 !== '—' ? `Poste : ${hStr2}` : '—';
  drawTLFromISO('fp-tl',allEvtsForTL,debDt.toISOString(),finDt.toISOString(),d.of_list||[]);

  // Productions
  const fpb=document.getElementById('fp-prods');
  if(fpb&&d.of_list){
    fpb.innerHTML=d.of_list.map(p=>`
      <tr>
        <td style="font-weight:600">${esc(p.of||'')}</td>
        <td>${esc(p.debut||'')}</td>
        <td>${esc(p.fin||'')}</td>
        <td>${esc(p.taille||'')}</td>
        <td>${esc(String(p.qte_fab||0))}</td>
        <td>${esc(String(p.equiv||''))}</td>
        <td>${esc(p.duree||'')}</td>
        <td class="${(p.trs||0)>=90?'tg':(p.trs||0)>=75?'tm':'tb'}">${fmtTRS(p.trs||0)}</td>
      </tr>`).join('')||'<tr><td colspan="8" style="color:var(--gray)">Aucune production</td></tr>';
  }

  // Stops list
  const fsl=document.getElementById('fp-stops-list');
  if(fsl){
    const stopsByType={};
    gEvts.forEach(e=>{if(!e.type)return;if(!stopsByType[e.type])stopsByType[e.type]=0;
    try{const p=s=>s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0);
    stopsByType[e.type]+=Math.max(0,p(e.fin||'00:00:00')-p(e.debut||'00:00:00'));}catch(ex){}});
    fsl.innerHTML=Object.entries(stopsByType).map(([t,s])=>`
      <div class="flex" style="padding:4px 0;border-bottom:1px solid var(--border);font-size:calc(12px*var(--zf,1))">
        <span style="flex:1;font-weight:600">${esc(t)}</span>
        <span style="color:var(--gray)">${fmtDurMS(s)}</span>
      </div>`).join('')||'<span style="color:var(--gray);font-size:calc(11px*var(--zf,1))">Aucun arrêt</span>';
  }

  // Graphs
  const prodS=d.tot_s||0;
  const stopS=stopTotal;
  drawPie('fp-pie',[
    {label:'Prod',value:prodS,color:'#16a34a'},
    {label:'Arrêts',value:stopS,color:'#dc2626'},
  ],{fLeg:9,legY:118});
  const trsS=d.trs_shift!==undefined?d.trs_shift:d.trs;
  drawGauge('fp-gauge-arc','fp-gauge-pct',trsS>=0?trsS:0);
}

async function applyFPHoraires(){
  const debStr=document.getElementById('fp-debut-dt').value;
  const finStr=document.getElementById('fp-fin-dt').value;
  if(!debStr||!finStr){toast('Renseigner début et fin','err');return;}
  const deb=new Date(debStr),fin=new Date(finStr);
  if(fin<=deb){toast('Fin doit être après début','err');return;}
  const shiftS=(fin.getTime()-deb.getTime())/1000;
  const shiftInfo=document.getElementById('fp-shift-info');
  const pad=n=>String(n).padStart(2,'0');
  const fmt=d=>pad(d.getHours())+':'+pad(d.getMinutes());
  if(shiftInfo) shiftInfo.textContent=fmt(deb)+'→'+fmt(fin);
  await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:deb.toISOString(),fin_iso:fin.toISOString()})});
  // Recalc TRS
  const d=window._fpData;
  if(d){
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const dedS=d.planned_ded_s||0;
    const netS=Math.max(1,shiftS-dedS);
    const trs=netS>0&&prodRef>0?Math.round(totEquiv/(prodRef*netS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
  // Redraw timeline with custom range
  drawTLFromISO('fp-tl',gEvts,deb.toISOString(),fin.toISOString(),(window._fpData&&window._fpData.of_list)||[]);
  closeM('m-fp-horaires');
  toast('Horaires appliqués','ok');
}

async function recalcFPTRS(){
  const sel=document.getElementById('fp-model-sel');
  if(!sel||!window._fpData) return;
  const modelNom=sel.value;
  if(!modelNom) return;
  const model=_cfgModels.find(m=>m.nom===modelNom);
  if(!model) return;
  // Show model shift info
  const now=new Date();
  const dayKeys=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=dayKeys[now.getDay()];
  const jours=model.jours||{};
  const jour=jours[dk]||{};
  const shiftInfo=document.getElementById('fp-shift-info');
  const shiftHours=document.getElementById('fp-shift-hours');
  const fpHD=document.getElementById('fp-horaire-display');
  const hStr=jour.debut&&jour.fin?`${jour.debut} → ${jour.fin}`:'—';
  if(shiftInfo) shiftInfo.textContent=hStr;
  if(shiftHours) shiftHours.textContent=hStr;
  if(fpHD) fpHD.textContent=hStr !== '—' ? `Poste : ${hStr}` : '—';
  // Recalculate TRS using this model's shift duration
  if(jour.debut&&jour.fin){
    const p=s=>{const[h,m]=s.split(':').map(Number);return h*3600+m*60;};
    let shiftS=p(jour.fin)-p(jour.debut);
    if(shiftS<0) shiftS+=86400;
    // Build debut datetime and call server
    const[dH,dM]=jour.debut.split(':').map(Number);
    const debDt=new Date(now);debDt.setHours(dH,dM,0,0);
    if(debDt>now) debDt.setDate(debDt.getDate()-1);
    const finDt=new Date(debDt.getTime()+shiftS*1000);
    await fetch('/api/update_shift_horaires',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({debut_iso:debDt.toISOString(),fin_iso:finDt.toISOString()})});
    const d=window._fpData;
    const totEquiv=d.tot_equiv||0;
    const prodRef=d.prod_ref||ST.prod_ref||200;
    const dedS=d.planned_ded_s||0;
    const netS=Math.max(1,shiftS-dedS);
    const trs=netS>0&&prodRef>0?Math.round(totEquiv/(prodRef*netS/28800)*1000)/10:0;
    document.getElementById('fp-trs').textContent=fmtTRS(trs);
  }
}

async function confirmFinPoste(){
  const stops=gEvts.filter(e=>e.type);
  const pSec=s=>s?s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0):0;
  let arret_s=0,pause_s=0,nett_s=0;
  stops.forEach(e=>{
    const dur=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));
    const t=(e.type||'').toLowerCase();
    if(t.includes('pause')) pause_s+=dur;
    else if(t.includes('nett')) nett_s+=dur;
    else arret_s+=dur;
  });
  const fpData=await apiFetch('/api/fin_poste_data');
  const ofList=fpData&&fpData.of_list||[];
  const prod_total=ofList.reduce((s,o)=>s+parseInt(o.qte_fab||0),0);
  const dur_prod_total_s=ofList.reduce((s,o)=>s+pSec(o.duree||''),0);
  // Net prod = each OF's (fin-debut) minus overlapping stops within that period
  const _fpOfPs=ofList.map(o=>({s:pSec(o.debut||'0:0:0'),e:pSec(o.fin||'0:0:0')})).filter(p=>p.e>p.s);
  const _fpStPs=stops.map(e=>({s:pSec(e.debut||'0:0:0'),e:pSec(e.fin||'0:0:0')})).filter(e=>e.e>e.s);
  let _fpNet=0;
  _fpOfPs.forEach(pp=>{
    const ov=_fpStPs.map(sv=>({s:Math.max(sv.s,pp.s),e:Math.min(sv.e,pp.e)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    _fpNet+=Math.max(0,pp.e-pp.s-mg.reduce((a,o)=>a+(o.e-o.s),0));
  });
  const dur_prod_sans_arret_s=_fpNet;
  const posteRow={
    date:new Date().toLocaleDateString('fr-FR'),
    pilot:ST.pilot||'',
    copilote:ST.form&&ST.form.copilote||'',
    poste:ST.poste||'',
    nb_of:fpData?fpData.nb_of||0:0,
    prod_total,
    tot_equiv:fpData?fpData.tot_equiv||0:0,
    trs_shift:fpData?fpData.trs_shift||0:0,
    arret_min:Math.round(arret_s/60),
    pause_min:Math.round(pause_s/60),
    nett_min:Math.round(nett_s/60),
    dur_prod_total_min:Math.round(dur_prod_total_s/60),
    dur_prod_sans_arret_min:Math.round(dur_prod_sans_arret_s/60),
    dur_poste_theorique_min:fpData&&fpData.model_dur_s?Math.round(fpData.model_dur_s/60):0,
    comment:''
  };
  await fetch('/api/save_poste',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(posteRow)});
  await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  resetToLogin();
  toast('Bonne fin de poste !','ok');
}

// ── KPI TAB ──
function _kpiGauge(arcId,pctId,trs,pArc){
  // pArc=157 for bigger viewBox (radius 50), 132 for radius 42
  const P=pArc||157;
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct)return;
  const v=Math.max(0,Math.min(100,trs||0)),dash=(v/100)*P;
  const col=v>=90?'#16a34a':v>=75?'#d97706':'#dc2626';
  arc.setAttribute('stroke-dasharray',`${dash.toFixed(1)},${P}`);arc.setAttribute('stroke',col);
  pct.textContent=trs>=0?fmtTRS(trs):'--%';pct.setAttribute('fill',col);
}
function _kpiNoData(arcId,pctId,pArc){
  const arc=document.getElementById(arcId),pct=document.getElementById(pctId);if(!arc||!pct)return;
  arc.setAttribute('stroke-dasharray',`0,${pArc||157}`);arc.setAttribute('stroke','#334155');
  pct.textContent='--%';pct.setAttribute('fill','#475569');
}
function _drawKpiTL(svgId,evts,startISO,endISO,isCurrent){
  const svg=document.getElementById(svgId);if(!svg)return;
  const H=isCurrent?50:36,W=800;
  const tS=new Date(startISO).getTime(),tE=new Date(endISO).getTime();
  const span=tE-tS;
  if(span<=0){svg.innerHTML=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e293b" rx="4"/>`;return;}
  const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
  let html=`<rect x="0" y="0" width="${W}" height="${H}" fill="#1e2d48" rx="4"/>`;
  // Prod background (green)
  html+=`<rect x="0" y="0" width="${W}" height="${H}" fill="#166534" rx="4" opacity=".65"/>`;
  // Events (stops)
  (evts||[]).forEach(ev=>{
    const t1=parseHMStoT(ev.debut,ev.date),t2=parseHMStoT(ev.fin,ev.date);
    if(!t1) return;
    const x1=toX(t1),x2=toX(t2||(t1+300000));if(x2<=x1)return;
    const cat=ev.cat||'autre';const col=STOP_COL[cat]||'#94a3b8';
    html+=`<rect x="${x1}" y="0" width="${Math.max(2,x2-x1)}" height="${H}" fill="${col}" rx="2" opacity=".9"/>`;
  });
  // Current live stop for active session
  if(isCurrent&&_curStopKey&&_curStopKey!=='_pause'&&ST.prod_active){
    const se=_curStopElap+(Date.now()-_lastPoll)/1000;
    const sT=tE-se*1000;
    const x1=toX(sT),x2=toX(tE);
    if(x2>x1) html+=`<rect x="${x1}" y="0" width="${x2-x1}" height="${H}" fill="#dc2626" rx="2" opacity=".9"/>`;
  }
  const fT=t=>{const d=new Date(t);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
  const fs=isCurrent?11:9;
  html+=`<text x="4" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tS)}</text>`;
  html+=`<text x="${W-40}" y="${H-4}" font-size="${fs}" fill="#fff">${fT(tE)}</text>`;
  if(span>3600000){const mid=(tS+tE)/2;html+=`<line x1="${W/2}" y1="0" x2="${W/2}" y2="${H}" stroke="#475569" stroke-width="1" stroke-dasharray="3,3"/><text x="${W/2-18}" y="${H-4}" font-size="${fs}" fill="#e2e8f0">${fT(mid)}</text>`;}
  svg.innerHTML=html;
}

// ── KPI helpers ──
function _kpiParseFR(fr){if(!fr)return null;const p=fr.split('/');if(p.length!==3)return null;return new Date(parseInt(p[2]),parseInt(p[1])-1,parseInt(p[0])).getTime();}
function _kpiTrsColor(t){return t>=90?'#16a34a':t>=70?'#f59e0b':t>=0?'#dc2626':'#94a3b8';}

function _kpiDrawDonut(svgId,legendId,segments){
  const svg=document.getElementById(svgId);const leg=document.getElementById(legendId);
  if(!svg)return;
  const total=segments.reduce((a,s)=>a+s.v,0);
  if(total<=0){svg.innerHTML='<text x="50" y="55" text-anchor="middle" font-size="9" fill="#94a3b8">Aucune donnée</text>';if(leg)leg.innerHTML='';return;}
  const cx=50,cy=50,R=42,r=22;let html='',a=-Math.PI/2;
  segments.forEach(s=>{
    if(s.v<=0)return;
    const ang=(s.v/total)*2*Math.PI;if(ang<0.002)return;
    const ea=a+ang,lg=ang>Math.PI?1:0;
    const x1=(cx+R*Math.cos(a)).toFixed(1),y1=(cy+R*Math.sin(a)).toFixed(1);
    const x2=(cx+R*Math.cos(ea)).toFixed(1),y2=(cy+R*Math.sin(ea)).toFixed(1);
    const ix1=(cx+r*Math.cos(a)).toFixed(1),iy1=(cy+r*Math.sin(a)).toFixed(1);
    const ix2=(cx+r*Math.cos(ea)).toFixed(1),iy2=(cy+r*Math.sin(ea)).toFixed(1);
    html+=`<path d="M${x1},${y1} A${R},${R} 0 ${lg},1 ${x2},${y2} L${ix2},${iy2} A${r},${r} 0 ${lg},0 ${ix1},${iy1} Z" fill="${s.col}"/>`;
    a=ea;
  });
  const top=segments.reduce((a,s)=>s.v>a.v?s:a,segments[0]);
  const topPct=Math.round(top.v/total*100);
  html+=`<text x="${cx}" y="${cy+4}" text-anchor="middle" font-size="14" font-weight="800" fill="#1e293b">${topPct}%</text>`;
  html+=`<text x="${cx}" y="${cy+13}" text-anchor="middle" font-size="8" fill="#64748b">${esc(top.label)}</text>`;
  svg.innerHTML=html;
  if(leg) leg.innerHTML=segments.map(s=>{const p=Math.round(s.v/total*100);return `<div style="display:flex;align-items:center;gap:4px"><div style="width:8px;height:8px;border-radius:2px;background:${s.col};flex-shrink:0"></div><span style="color:#475569">${esc(s.label)}</span><b style="margin-left:3px;color:#1e293b">${p}%</b></div>`;}).join('');
}

function _kpiLineChart(containerId,items,valueKey,colorFn,unit,yMin,yMax){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const vals=items.map(it=>it[valueKey]||0);
  const minV=yMin!==undefined?yMin:Math.max(0,Math.min(...vals)-5);
  const maxV=yMax!==undefined?yMax:Math.max(...vals,1)+2;
  const padL=30,padR=8,padT=10,padB=62;
  const gW=W-padL-padR,gH=H-padT-padB;
  const toX=i=>padL+i/(Math.max(items.length-1,1))*gW;
  const toY=v=>padT+gH*(1-(v-minV)/(maxV-minV||1));
  // grid
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  const nTicks=4;
  for(let t=0;t<=nTicks;t++){
    const v=minV+(maxV-minV)*t/nTicks;
    const y=toY(v);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="8" fill="#94a3b8">${v%1?v.toFixed(0):v}${unit||''}</text>`;
  }
  // area fill
  let areaD=`M${toX(0).toFixed(1)},${(H-padB).toFixed(1)}`;
  items.forEach((it,i)=>{areaD+=` L${toX(i).toFixed(1)},${toY(it[valueKey]||0).toFixed(1)}`;});
  areaD+=` L${toX(items.length-1).toFixed(1)},${(H-padB).toFixed(1)} Z`;
  const areaCol=colorFn?colorFn(items[0]):'#6366f1';
  svg+=`<path d="${areaD}" fill="${areaCol}" opacity=".08"/>`;
  // line
  let lineD='';
  items.forEach((it,i)=>{lineD+=(i===0?'M':'L')+toX(i).toFixed(1)+','+toY(it[valueKey]||0).toFixed(1)+' ';});
  svg+=`<path d="${lineD}" fill="none" stroke="${areaCol}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
  // dots + value labels
  items.forEach((it,i)=>{
    const x=toX(i),y=toY(it[valueKey]||0);
    const col=colorFn?colorFn(it):areaCol;
    svg+=`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="3" fill="${col}" stroke="#fff" stroke-width="1.5"/>`;
    const v=it[valueKey]||0;
    svg+=`<text x="${x.toFixed(1)}" y="${(y-6).toFixed(1)}" text-anchor="middle" font-size="8" font-weight="700" fill="${col}">${v%1?v.toFixed(1):v}</text>`;
    // X labels: date (line1) + poste (line2) — bigger, black
    const lblParts=(it._xLabel||'').split('\n');
    svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+3).toFixed(1)}) rotate(35)" font-size="9" font-weight="700" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[0]||'')}</text>`;
    if(lblParts[1]) svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+17).toFixed(1)}) rotate(35)" font-size="8" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[1]||'')}</text>`;
  });
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiBarChart(containerId,items,valueKey,colorFn,unit){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const vals=items.map(it=>it[valueKey]||0);
  const maxV=Math.max(...vals,1);
  const padL=30,padR=8,padT=8,padB=46;
  const gW=W-padL-padR,gH=H-padT-padB;
  const n=items.length;
  const barW=Math.max(3,Math.min(30,gW/n-4));
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  [0,0.5,1].forEach(t=>{
    const v=maxV*t;const y=padT+gH*(1-t);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="8" fill="#94a3b8">${Math.round(v)}${t>0&&unit?unit:''}</text>`;
  });
  items.forEach((it,i)=>{
    const v=it[valueKey]||0;
    const x=padL+(i/n)*gW+(gW/n-barW)/2;
    const bH=v>0?(v/maxV)*gH:0;
    const y=padT+gH-bH;
    const col=colorFn?colorFn(it):'#f59e0b';
    svg+=`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${Math.max(0,bH).toFixed(1)}" fill="${col}" rx="2" opacity=".85"/>`;
    if(v>0)svg+=`<text x="${(x+barW/2).toFixed(1)}" y="${(y-2).toFixed(1)}" text-anchor="middle" font-size="8" font-weight="700" fill="${col}">${v%1?v.toFixed(1):v}</text>`;
    const lbl=(it._xLabel||'').split('\n')[0];
    svg+=`<text transform="translate(${(x+barW/2).toFixed(1)},${(H-padB+4).toFixed(1)}) rotate(35)" font-size="7" fill="#64748b" dominant-baseline="hanging">${esc(lbl)}</text>`;
  });
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiDualLineChart(containerId,items,series){
  const el=document.getElementById(containerId);if(!el)return;
  const rect=el.getBoundingClientRect();
  const W=Math.max(rect.width||400,200);
  const H=Math.max(rect.height||100,60);
  if(!items.length){el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%"><text x="${W/2}" y="${H/2}" text-anchor="middle" font-size="10" fill="#94a3b8">Aucune donnée</text></svg>`;return;}
  const allVals=series.flatMap(s=>items.map(it=>parseFloat(it[s.key]||0)));
  const minV=Math.max(0,Math.min(...allVals)-2);
  const maxV=Math.max(...allVals,1)+2;
  const padL=32,padR=8,padT=8,padB=62;
  const gW=W-padL-padR,gH=H-padT-padB;
  const toX=i=>padL+i/(Math.max(items.length-1,1))*gW;
  const toY=v=>padT+gH*(1-(v-minV)/(maxV-minV||1));
  let svg=`<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:100%">`;
  [0,0.25,0.5,0.75,1].forEach(t=>{
    const v=minV+(maxV-minV)*t;const y=toY(v);
    svg+=`<line x1="${padL}" y1="${y.toFixed(1)}" x2="${W-padR}" y2="${y.toFixed(1)}" stroke="#f1f5f9" stroke-width="1"/>`;
    svg+=`<text x="${padL-3}" y="${(y+4).toFixed(1)}" text-anchor="end" font-size="8" fill="#94a3b8">${v.toFixed(0)}</text>`;
  });
  series.forEach(ser=>{
    let lineD='';
    items.forEach((it,i)=>{lineD+=(i===0?'M':'L')+toX(i).toFixed(1)+','+toY(parseFloat(it[ser.key]||0)).toFixed(1)+' ';});
    svg+=`<path d="${lineD}" fill="none" stroke="${ser.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" ${ser.dash?`stroke-dasharray="${ser.dash}"`:''}/>`;
    items.forEach((it,i)=>{
      const x=toX(i),y=toY(parseFloat(it[ser.key]||0));
      const v=parseFloat(it[ser.key]||0);
      svg+=`<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="2.5" fill="${ser.color}" stroke="#fff" stroke-width="1.2"/>`;
      svg+=`<text x="${x.toFixed(1)}" y="${(y-5).toFixed(1)}" text-anchor="middle" font-size="7" fill="${ser.color}" font-weight="700">${v%1?v.toFixed(1):v}</text>`;
    });
  });
  items.forEach((it,i)=>{
    const x=toX(i);
    const lblParts=(it._xLabel||'').split('\n');
    svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+3).toFixed(1)}) rotate(35)" font-size="9" font-weight="700" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[0]||'')}</text>`;
    if(lblParts[1]) svg+=`<text transform="translate(${x.toFixed(1)},${(H-padB+17).toFixed(1)}) rotate(35)" font-size="8" fill="#1e293b" dominant-baseline="hanging">${esc(lblParts[1]||'')}</text>`;
  });
  let lx=padL;
  series.forEach(ser=>{
    svg+=`<line x1="${lx}" y1="${H-7}" x2="${lx+14}" y2="${H-7}" stroke="${ser.color}" stroke-width="2" ${ser.dash?`stroke-dasharray="${ser.dash}"`:''}/>`;
    svg+=`<text x="${lx+16}" y="${H-3}" font-size="8" fill="#475569">${esc(ser.label||ser.key)}</text>`;
    lx+=90;
  });
  svg+='</svg>';
  el.innerHTML=svg;
}

function _kpiInitDates(){
  const fi=document.getElementById('kpi-from'),ti=document.getElementById('kpi-to');
  if(!fi||!ti)return;
  if(!fi.value){const d=new Date();d.setDate(d.getDate()-13);fi.value=d.toISOString().slice(0,10);}
  if(!ti.value){ti.value=new Date().toISOString().slice(0,10);}
}

async function loadKPI(){
  _kpiInitDates();
  const fi=document.getElementById('kpi-from'),ti=document.getElementById('kpi-to');
  const fromMs=fi&&fi.value?new Date(fi.value).getTime():0;
  const toMs=ti&&ti.value?new Date(ti.value).getTime()+86399000:Date.now();
  const pSec=s=>s?s.split(':').reduce((a,v,i)=>a+(i===0?+v*3600:i===1?+v*60:+v),0):0;

  const [histData,evtData]=await Promise.all([apiFetch('/api/history'),apiFetch('/api/events_list')]);
  const allRows=Array.isArray(histData)?histData:[];
  const allEvts=Array.isArray(evtData)?evtData:[];

  const inRange=dateStr=>{const ms=_kpiParseFR(dateStr);return ms&&ms>=fromMs&&ms<=toMs;};
  const rows=allRows.filter(r=>inRange(r.date)&&(r.type||'').trim().toLowerCase()!=='');
  const prodRows=allRows.filter(r=>inRange(r.date)&&['production','prod',''].includes((r.type||'').trim().toLowerCase()));
  const evts=allEvts.filter(e=>inRange(e.date));

  // Build sessions (one per pilot+date+poste)
  const sessMap={};
  const sessRowsMap={};
  prodRows.forEach(r=>{
    const key=(r.pilote||'')+'||'+(r.date||'')+'||'+(r.poste||'');
    if(!sessMap[key]) sessMap[key]={pilot:r.pilote||'',date:r.date||'',poste:r.poste||'',nb_of:0,tot_equiv:0,tot_s:0,tot_qte:0,trs_vals:[]};
    if(!sessRowsMap[key]) sessRowsMap[key]=[];
    const s=sessMap[key];
    s.nb_of++;
    const eq=parseFloat(r.equiv||0)||0; s.tot_equiv+=eq;
    const qte=parseFloat(r.qte_fab||0)||0; s.tot_qte+=qte;
    const ds=pSec(r.debut||'0:0:0'),fs=pSec(r.fin||'0:0:0'),dur=Math.max(0,fs-ds); s.tot_s+=dur;
    const t=parseFloat(r.trs||0);if(t>0)s.trs_vals.push(t);
    sessRowsMap[key].push(r);
  });
  // Order chronologically
  const sessArr=Object.values(sessMap).sort((a,b)=>{const da=_kpiParseFR(a.date)||0,db=_kpiParseFR(b.date)||0;return da-db;});
  sessArr.forEach(s=>{
    s.trs=s.trs_vals.length?Math.round(s.trs_vals.reduce((a,v)=>a+v,0)/s.trs_vals.length*10)/10:-1;
    s.cad=s.tot_s>60?Math.round(s.tot_equiv/(s.tot_s/3600)*10)/10:0;
    s.cad_qte=s.tot_s>60&&s.tot_qte>0?Math.round(s.tot_qte/(s.tot_s/3600)*10)/10:0;
    const evtKey=s.pilot+'||'+s.date+'||'+s.poste;
    s.stop_min=Math.round((evts.filter(e=>(e.pilote||'')+'||'+(e.date||'')+'||'+(e.poste||'')==evtKey).reduce((a,e)=>a+Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0')),0))/60);
    // Fibre changes: count transitions between different fibre values
    const sRows=(sessRowsMap[evtKey]||[]).sort((a,b)=>pSec(a.debut||'0:0')-pSec(b.debut||'0:0'));
    let fibChg=0;
    for(let i=1;i<sRows.length;i++){const fa=(sRows[i-1].fibre||'').trim(),fb=(sRows[i].fibre||'').trim();if(fa&&fb&&fa!==fb)fibChg++;}
    s.nb_fibre_chg=fibChg;
    // Prod/arrêt ratio
    const tot_dur=s.tot_s+s.stop_min*60;
    s.prod_pct=tot_dur>0?Math.round(s.tot_s/tot_dur*100):0;
    s.stop_pct=100-s.prod_pct;
    // Label X axis
    const dParts=(s.date||'').split('/');
    const dateShort=dParts.length===3?dParts[0]+'/'+dParts[1]:s.date;
    s._xLabel=dateShort+'\n'+s.poste;
  });

  const nbSess=sessArr.length;
  const totalEquiv=sessArr.reduce((a,s)=>a+s.tot_equiv,0);
  const totalOF=sessArr.reduce((a,s)=>a+s.nb_of,0);
  const totalQte=sessArr.reduce((a,s)=>a+s.tot_qte,0);
  const totalProdS=sessArr.reduce((a,s)=>a+s.tot_s,0);
  const totalStopMin=sessArr.reduce((a,s)=>a+s.stop_min,0);
  const trsVals=sessArr.filter(s=>s.trs>=0).map(s=>s.trs);
  const avgTRS=trsVals.length?Math.round(trsVals.reduce((a,v)=>a+v,0)/trsVals.length*10)/10:-1;
  const avgCadH=totalProdS>0?Math.round(totalEquiv/(totalProdS/3600)*10)/10:0;
  const avgOFperSess=nbSess?Math.round(totalOF/nbSess*10)/10:0;
  const avgQtePerSess=nbSess?Math.round(totalQte/nbSess*10)/10:0;
  const avgEquivPerSess=nbSess?Math.round(totalEquiv/nbSess*10)/10:0;

  // ── KPI Cards ──
  const cards=document.getElementById('kpi-cards');
  if(cards){
    const avgStopMin=nbSess?Math.round(totalStopMin/nbSess):0;
    const mkCard=(lbl,val,col,sub)=>`<div style="padding:10px 14px;border-right:1px solid #f1f5f9;text-align:center;min-width:100px;flex-shrink:0;display:flex;flex-direction:column;justify-content:center">
      <div style="font-size:calc(22px*var(--zf,1));font-weight:900;color:${col};line-height:1;letter-spacing:-.5px">${val}</div>
      <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#94a3b8;letter-spacing:.4px;margin-top:4px;line-height:1.3">${lbl}</div>
      ${sub?`<div style="font-size:calc(9px*var(--zf,1));color:#64748b;margin-top:2px">${sub}</div>`:''}
    </div>`;
    cards.innerHTML=
      mkCard('TRS moyen',avgTRS>=0?avgTRS.toFixed(1)+'%':'—',_kpiTrsColor(avgTRS),nbSess?nbSess+' postes':'')+
      mkCard('Nb OF total',totalOF,'#6366f1','moy '+avgOFperSess+'/poste')+
      mkCard('Nb postes',nbSess,'#0891b2','')+
      mkCard('Pièces fabriquées',Math.round(totalQte),'#7c3aed',avgQtePerSess+'/poste')+
      mkCard('Équiv. totale',totalEquiv.toFixed(1),'#16a34a',avgEquivPerSess.toFixed(1)+'/poste')+
      mkCard('Cadence moy.',avgCadH>0?avgCadH+'/h':'—','#f59e0b','')+
      mkCard('Arrêts totaux',totalStopMin+'min','#dc2626',avgStopMin+'min/poste')+
      mkCard('Prod totale',Math.round(totalProdS/60)+'min','#0891b2','');
  }

  // ── 3 courbes (compact, côte à côte) ──
  const trsItems=sessArr.filter(s=>s.trs>=0);
  _kpiLineChart('kpi-trs-chart',trsItems,'trs',it=>_kpiTrsColor(it.trs),'%',0,100);
  _kpiLineChart('kpi-arr-chart',sessArr,'stop_min',()=>'#dc2626','min');
  _kpiLineChart('kpi-of-chart',sessArr,'nb_of',()=>'#6366f1','');

  // ── Cadence dual-line chart (éq./h + pièces/h) ──
  _kpiDualLineChart('kpi-cad-chart',sessArr,[
    {key:'cad',color:'#f59e0b',label:'Éq./h'}
  ]);

  // ── Évolution changements de fibre par poste ──
  _kpiLineChart('kpi-fibre-chart',sessArr,'nb_fibre_chg',()=>'#8b5cf6','');

  // ── Évolution pièces / équivalence ──
  _kpiDualLineChart('kpi-qte-chart',sessArr,[
    {key:'tot_qte',color:'#7c3aed',label:'Pièces'},
    {key:'tot_equiv',color:'#16a34a',label:'Équivalence',dash:'4,3'}
  ]);

  // ── Donut Prod/Arrêts ──
  _kpiDrawDonut('kpi-donut','kpi-donut-legend',[
    {label:'Prod',v:totalProdS,col:'#16a34a'},
    {label:'Arrêts',v:totalStopMin*60,col:'#dc2626'},
  ]);

  // ── Pareto ──
  const stopMap={};const stopCat={};
  evts.forEach(e=>{
    if(!e.type)return;
    const dur=Math.max(0,pSec(e.fin||'0:0:0')-pSec(e.debut||'0:0:0'));
    stopMap[e.type]=(stopMap[e.type]||0)+dur;
    if(!stopCat[e.type])stopCat[e.type]=e.cat||'autre';
  });
  const paretoArr=Object.entries(stopMap).sort((a,b)=>b[1]-a[1]);
  const paretoTotal=paretoArr.reduce((a,[,v])=>a+v,0);
  const parEl=document.getElementById('kpi-pareto-new');
  if(parEl){
    if(!paretoArr.length){parEl.innerHTML='<div style="color:#94a3b8;font-size:calc(11px*var(--zf,1))">Aucun arrêt</div>';}
    else{
      const maxP=paretoArr[0][1];
      let cumul=0;
      parEl.innerHTML=paretoArr.slice(0,12).map(([type,s])=>{
        const pct=Math.round(s/maxP*100);
        const min=Math.round(s/60);
        const pctTot=paretoTotal>0?Math.round(s/paretoTotal*100):0;
        cumul+=pctTot;
        const col=STOP_COL[stopCat[type]]||'#94a3b8';
        return `<div>
          <div style="display:flex;justify-content:space-between;font-size:calc(10px*var(--zf,1));color:#374151;margin-bottom:2px">
            <div style="display:flex;align-items:center;gap:3px;min-width:0"><div style="width:7px;height:7px;border-radius:1px;background:${col};flex-shrink:0"></div><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:110px">${esc(type)}</span></div>
            <span style="white-space:nowrap;color:#6b7280;flex-shrink:0">${min}m <b style="color:#1e293b">${pctTot}%</b></span>
          </div>
          <div style="background:#f1f5f9;border-radius:3px;height:9px;overflow:hidden">
            <div style="width:${pct}%;background:${col};height:100%;border-radius:3px"></div>
          </div>
        </div>`;
      }).join('');
    }
  }
}

// ── HISTORY ROW DETAIL ──
function showHistRowDetail(key){
  const r=window._rowMap&&window._rowMap[String(key)];
  if(!r) return;
  const isProd=r._rowType==='prod';
  let ofEvts=[];
  if(isProd){
    const _hm2s=hm=>{if(!hm)return 0;const[h,m]=(hm+':00').split(':').map(Number);return(h||0)*3600+(m||0)*60;};
    const debS=_hm2s(r.debut), finS=_hm2s(r.fin)||86400;
    ofEvts=(window._histEvtsAll||[]).filter(e=>{
      if(!e.date||e.date!==r.date) return false;
      if(e.pilote&&e.pilote!==r.pilote) return false;
      const t=_hm2s(e.debut); return t>=debS&&t<=finS;
    });
  }
  _renderAndOpenOfDetail(r, ofEvts);
}

// ── HISTORY ──
let _histFilters=new Set(); // filtres actifs

function toggleHistFilter(btn){
  const key=btn.dataset.hf;
  if(_histFilters.has(key)){
    _histFilters.delete(key);
    btn.style.background='none';
    btn.style.color=btn.style.borderColor; // remettre couleur texte
    btn.style.opacity='1';
  } else {
    _histFilters.add(key);
    btn.style.background=btn.style.borderColor;
    btn.style.color='#fff';
  }
  _applyHistFilter();
}

function _histMatchFilter(r){
  if(!_histFilters.size) return true;
  const t=String(r.type||r._rowType||'').toLowerCase();
  const isProd=r._rowType==='prod'||t==='production'||t==='prod'||t==='';
  if(_histFilters.has('production')&&isProd) return true;
  if(_histFilters.has('nettoyage')&&(t.includes('nettoyage')||t.includes('nett'))) return true;
  if(_histFilters.has('pause')&&t.includes('pause')) return true;
  if(_histFilters.has('reunion')&&(t.includes('réunion')||t.includes('reunion')||t.includes('meeting'))) return true;
  if(_histFilters.has('arret')&&!isProd&&!t.includes('nettoyage')&&!t.includes('pause')&&!(t.includes('réunion')||t.includes('reunion')||t.includes('meeting'))) return true;
  return false;
}

function filterHistBySearch(){
  const q=(document.getElementById('hist-search')?.value||'').toLowerCase().trim();
  document.querySelectorAll('#hist-bd tr[data-hftype]').forEach(tr=>{
    if(!q){tr.dataset.searchHidden='';return;}
    const ofTxt=(tr.cells[1]?.textContent||'').toLowerCase();
    tr.dataset.searchHidden=ofTxt.includes(q)?'':'1';
  });
  _applyHistFilter();
}

function _applyHistFilter(){
  const rows=document.querySelectorAll('#hist-bd tr[data-hftype]');
  rows.forEach(tr=>{
    if(tr.dataset.searchHidden==='1'){tr.style.display='none';return;}
    const t=tr.dataset.hftype||'';
    const isProd=t==='prod';
    const isNett=t.includes('nettoyage')||t.includes('nett');
    const isPause=t.includes('pause');
    const isReunion=t.includes('réunion')||t.includes('reunion')||t.includes('meeting');
    const isArret=!isProd&&!isNett&&!isPause&&!isReunion;
    let show=!_histFilters.size;
    if(!show){
      if(_histFilters.has('production')&&isProd) show=true;
      if(_histFilters.has('nettoyage')&&isNett) show=true;
      if(_histFilters.has('pause')&&isPause) show=true;
      if(_histFilters.has('reunion')&&isReunion) show=true;
      if(_histFilters.has('arret')&&isArret) show=true;
    }
    tr.style.display=show?'':'none';
  });
}

function _resetHistFilters(){
  _histFilters.clear();
  document.querySelectorAll('.hf-btn').forEach(b=>{
    b.style.background='none';
    b.style.color=b.style.borderColor;
    b.style.opacity='1';
  });
}

async function loadHist(){
  const today=new Date().toISOString().slice(0,10);
  const from=document.getElementById('hist-from').value||today;
  const to=document.getElementById('hist-to').value||today;
  const [declData,evtData]=await Promise.all([
    apiFetch(`/api/history?from=${from}&to=${to}`),
    apiFetch('/api/events_list')
  ]);
  const decls=Array.isArray(declData)?declData:(declData&&declData.rows?declData.rows:[]);
  const evts=Array.isArray(evtData)?evtData:[];
  // Filter events to the selected date range
  const fromD=new Date(from+'T00:00:00'),toD=new Date(to+'T23:59:59');
  const fmtFR=d=>{const[y,m,dy]=d.split('-');return `${dy}/${m}/${y}`;};
  const fmtFRfrom=fmtFR(from),fmtFRto=fmtFR(to);
  const evtsFiltered=evts.filter(e=>{
    if(!e.date) return false;
    // Support both dd/mm/yyyy and yyyy-mm-dd
    let dt;
    const parts=e.date.split('/');
    if(parts.length===3&&parts[2].length===4) dt=new Date(parts[2]+'-'+parts[1]+'-'+parts[0]);
    else dt=new Date(e.date);
    return dt>=fromD&&dt<=toD;
  });
  const hd=document.getElementById('hist-hd'),bd=document.getElementById('hist-bd');
  if(!hd||!bd) return;
  // Build unified row list same as loadMainDecl
  const allRows=[];
  decls.forEach(r=>allRows.push({...r,_rowType:'prod'}));
  evtsFiltered.forEach(r=>allRows.push({...r,_rowType:'evt'}));
  allRows.sort((a,b)=>{
    const da=a.date||'',db=b.date||'';
    if(da!==db) return db.localeCompare(da);
    return (b.debut||'').localeCompare(a.debut||'');
  });
  hd.innerHTML='<th>Type</th><th>OF</th><th>Fibre</th><th>Date</th><th>Poste</th><th>Pilote</th><th>Début</th><th>Fin</th><th>Détails</th><th>Qté/Durée</th><th>TRS/Info</th><th>Commentaire</th><th>Actions</th>';
  if(!allRows.length){bd.innerHTML='<tr><td colspan="12" style="text-align:center;color:var(--gray);padding:16px">Aucune donnée sur cette période</td></tr>';return;}
  window._rowMap=window._rowMap||{};
  window._histEvtsAll=evtsFiltered; // pour showHistRowDetail
  bd.innerHTML=allRows.map(r=>{
    const key=r.row_num||r.debut;
    window._rowMap[String(key)]=r;
    const isProd=r._rowType==='prod';
    const rt=String(r.type||'').toLowerCase();
    const hftype=isProd?'prod':rt||'arret';
    const t=parseFloat(r.trs||0);
    const tag=isProd?'<span class="row-tag tag-p">🏭 Prod</span>':(rt.includes('nett')?'<span class="row-tag tag-n">🧹 Nett.</span>':rt.includes('pause')?'<span class="row-tag tag-n" style="border-color:#f59e0b;color:#f59e0b">⏸ Pause</span>':(rt.includes('réunion')||rt.includes('reunion'))?'<span class="row-tag tag-n" style="border-color:#8b5cf6;color:#8b5cf6">👥 Réunion</span>':'<span class="row-tag tag-e">⛔ Arrêt</span>');
    const details=isProd?esc(r.taille||''):esc(r.type||'');
    const qty=isProd?esc(String(r.qte_fab||'')):esc(r.duree||'');
    const info=isProd&&t>0?`<span class="${t>=90?'tg':t>=75?'tm':'tb'}">${fmtTRS(t)}</span>`:'—';
    const cmt=esc(r.comment||'');
    const fbrH=r.fibre||'';const fbrShH=esc(fbrH.slice(0,9));
    return `<tr class="${isProd?'row-prod':'row-evt'}" data-hftype="${esc(hftype)}">
      <td>${tag}</td><td style="font-weight:700;color:#1e3a8a;text-decoration:underline;cursor:pointer" onclick="showHistRowDetail('${esc(String(key))}')" title="Voir détail">${esc(r.of||r.type||'—')}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:#6366f1;font-weight:600;cursor:${fbrH?'pointer':''}" title="${esc(fbrH)}" onclick="${fbrH?'showFibre(\''+esc(fbrH)+'\')':''}">${fbrShH}${fbrH.length>9?'…':''}</td>
      <td style="font-size:calc(10px*var(--zf,1))">${esc(r.date||'')}</td><td style="font-size:calc(10px*var(--zf,1))">${esc(r.poste||'')}</td>
      <td>${esc(r.pilote||'')}</td><td>${esc(r.debut||'')}</td><td>${esc(r.fin||'')}</td>
      <td style="font-size:calc(11px*var(--zf,1))">${details}</td><td style="font-size:calc(11px*var(--zf,1))">${qty}</td><td>${info}</td>
      <td style="font-size:calc(10px*var(--zf,1));color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${cmt}">${cmt}</td>
      <td><button onclick="openEditRow('${esc(String(key))}')" style="background:#6366f1;color:#fff;border:none;border-radius:5px;padding:4px 10px;font-size:calc(15px*var(--zf,1));cursor:pointer;font-weight:700" title="Modifier">✏</button></td>
    </tr>`;
  }).join('');
  _applyHistFilter();
}

// ── RAPPORTS DES POSTES ──
async function reloadAndLoadRapports(){
  toast('Rechargement depuis Excel…','ok');
  await apiFetch('/api/reload_excel',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  await loadRapports();
}
async function loadRapports(){
  const listEl=document.getElementById('rpt-list');
  if(!listEl) return;
  listEl.innerHTML='<div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Chargement…</div>';
  const sessions=await apiFetch('/api/past_sessions');
  if(!sessions||!sessions.length){
    listEl.innerHTML='<div style="padding:20px;text-align:center;color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun poste disponible</div>';
    return;
  }
  listEl.innerHTML=sessions.map((s,i)=>{
    const trsStr=s.trs>=0?s.trs.toFixed(1)+'%':'—';
    const trsCol=s.trs>=90?'#16a34a':s.trs>=70?'#f59e0b':s.trs>=0?'#dc2626':'#94a3b8';
    return `<div class="rpt-item" id="rpt-item-${i}" onclick="loadSessionReport('${esc(s.date)}','${esc(s.pilot)}','${esc(s.poste)}','rpt-item-${i}')"
      style="padding:10px 14px;border-bottom:1px solid var(--border);cursor:pointer;transition:background .15s">
      <div style="font-size:calc(12px*var(--zf,1));font-weight:800;color:var(--navy)">${esc(s.date)} — ${esc(s.poste)}</div>
      <div style="font-size:calc(11px*var(--zf,1));color:var(--gray);margin-top:2px">${esc(s.pilot||'?')} | ${s.nb_of} OF | Éq. ${s.tot_equiv}</div>
      <div style="font-size:calc(16px*var(--zf,1));font-weight:900;color:${trsCol};margin-top:2px">${trsStr}</div>
    </div>`;
  }).join('');
}

async function loadSessionReport(date,pilot,poste,itemId){
  document.querySelectorAll('.rpt-item').forEach(el=>el.style.background='');
  const sel=document.getElementById(itemId);if(sel) sel.style.background='#eff6ff';
  const detailEl=document.getElementById('rpt-detail');
  if(!detailEl) return;
  detailEl.innerHTML='<div style="padding:40px;text-align:center;color:var(--gray)">Chargement…</div>';
  const d=await apiFetch(`/api/session_report?date=${encodeURIComponent(date)}&pilot=${encodeURIComponent(pilot)}&poste=${encodeURIComponent(poste)}`);
  if(!d){detailEl.innerHTML='<div style="padding:40px;text-align:center;color:#dc2626">Erreur chargement</div>';return;}
  const trsS=d.trs_shift>=0?d.trs_shift:d.trs;
  const trsCol=trsS>=90?'#16a34a':trsS>=70?'#f59e0b':trsS>=0?'#dc2626':'#94a3b8';
  const stopMin=Math.round((d.stop_s||0)/60);
  const prodMin=Math.round(Math.max(0,(d.tot_s||0)-(d.stop_s||0))/60);
  const totalMin=Math.round((d.tot_s||0)/60)+stopMin;
  // Pareto des arrêts
  const stopMap={};
  (d.evt_rows||[]).forEach(r=>{
    const k=r.type||'Inconnu';
    if(!stopMap[k]) stopMap[k]=0;
    const p=r.duree?r.duree.split(':'):[0,0,0];
    stopMap[k]+=(parseInt(p[0]||0)*3600+parseInt(p[1]||0)*60+parseInt(p[2]||0))/60;
  });
  const stopArr=Object.entries(stopMap).sort((a,b)=>b[1]-a[1]);
  const maxStopMin=stopArr.length?stopArr[0][1]:1;
  const paretoHtml=stopArr.length?stopArr.map(([k,v])=>`
    <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">
      <div style="font-size:calc(10px*var(--zf,1));width:100px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--text)">${esc(k)}</div>
      <div style="flex:1;background:#f1f5f9;border-radius:4px;height:14px;overflow:hidden">
        <div style="height:100%;background:#dc2626;border-radius:4px;width:${Math.round(v/maxStopMin*100)}%;opacity:.8"></div>
      </div>
      <div style="font-size:calc(10px*var(--zf,1));font-weight:700;color:#dc2626;width:36px;text-align:right;flex-shrink:0">${Math.round(v)}mn</div>
    </div>`).join(''):'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>';
  window._rptProdRows = d.prod_rows || [];
  window._rptEvtRows = d.evt_rows || [];
  // Helper: compute net prod and stop overlap for each OF against evt_rows
  const _rptHmsMs=hm=>{if(!hm)return 0;const[h,m,s]=(hm+':0:0').split(':').map(Number);return(h||0)*3600000+(m||0)*60000+(s||0)*1000;};
  const _rptStEvts=(d.evt_rows||[]).map(e=>({s:_rptHmsMs(e.debut),e:_rptHmsMs(e.fin)})).filter(e=>e.e>e.s);
  function _rptNetProd(debHm,finHm){
    const dMs=_rptHmsMs(debHm),fMs=_rptHmsMs(finHm);
    if(fMs<=dMs) return {netMin:0,stopMin:0};
    const ov=_rptStEvts.map(sv=>({s:Math.max(sv.s,dMs),e:Math.min(sv.e,fMs)})).filter(o=>o.e>o.s);
    ov.sort((a,b)=>a.s-b.s);
    const mg=[];ov.forEach(o=>{if(mg.length&&o.s<=mg[mg.length-1].e)mg[mg.length-1].e=Math.max(mg[mg.length-1].e,o.e);else mg.push({...o});});
    const blocked=mg.reduce((a,o)=>a+(o.e-o.s),0);
    return {netMin:Math.round(Math.max(0,fMs-dMs-blocked)/60000),stopMin:Math.round(blocked/60000)};
  }
  const prodsHtml=(d.prod_rows||[]).map((r,ri)=>{
    const tc=r.trs>=90?'#16a34a':r.trs>=70?'#f59e0b':r.trs>=0?'#dc2626':'#94a3b8';
    const kitStr=(r.kit||'').toLowerCase();
    const kitDisp=kitStr==='oui'?'<span style="color:#16a34a;font-weight:800">✓</span>':'';
    const {netMin,stopMin}=_rptNetProd(r.debut,r.fin);
    const td='padding:4px 6px;text-align:center;font-size:calc(11px*var(--zf,1))';
    return `<tr style="border-bottom:1px solid var(--border);cursor:pointer" onclick="showOfDetail(${ri})" title="Voir détail OF">
      <td style="${td};font-weight:700;color:#1e3a8a;text-decoration:underline">${esc(r.of||'')}</td>
      <td style="${td}">${esc(r.taille||'')} ${esc(r.type_prod||'')}</td>
      <td style="${td}">${kitDisp}</td>
      <td style="${td}">${esc(r.qte_fab||'')}</td>
      <td style="${td};color:#0891b2;font-weight:700">${esc(r.equiv||'')}</td>
      <td style="${td};white-space:nowrap">${esc(r.debut||'')} → ${esc(r.fin||'')}</td>
      <td style="${td};color:#16a34a;font-weight:700">${netMin} min</td>
      <td style="${td};color:#dc2626;font-weight:700">${stopMin} min</td>
      <td style="${td};font-weight:800;color:${tc}">${r.trs>=0?r.trs.toFixed(1)+'%':'—'}</td>
      <td style="padding:4px 6px;font-size:calc(10px*var(--zf,1));color:var(--gray)">${esc(r.comment||'')}</td>
    </tr>`;
  }).join('');
  // Timeline inline builder (ne dépend pas de ST)
  function buildTL(prodRows,evtRows,dateStr,modelDebut,modelFin){
    const W=800,Y=4,H2=28,H=40;
    // Convertir dd/mm/yyyy → base ms
    const parts=dateStr.split('/');
    const baseMs=parts.length===3?new Date(parts[2]+'-'+parts[1]+'-'+parts[0]).getTime():Date.now();
    const hm2ms=hm=>{if(!hm)return null;const[h,m]=(hm+':00').split(':').map(Number);return baseMs+h*3600000+m*60000;};
    const mdMs=hm2ms(modelDebut),mfMs=hm2ms(modelFin);
    // Compute range
    let allMs=[];
    prodRows.forEach(r=>{if(r.debut)allMs.push(hm2ms(r.debut));if(r.fin)allMs.push(hm2ms(r.fin));});
    evtRows.forEach(r=>{if(r.debut)allMs.push(hm2ms(r.debut));if(r.fin)allMs.push(hm2ms(r.fin));});
    const tS=mdMs||Math.min(...allMs.filter(Boolean));
    const tE=mfMs||Math.max(...allMs.filter(Boolean));
    if(!tS||!tE||tE<=tS) return `<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
    const span=tE-tS;
    const toX=t=>Math.max(0,Math.min(W,(t-tS)/span*W));
    let html=`<rect x="0" y="${Y}" width="${W}" height="${H2}" fill="#e2e8f0" rx="4"/>`;
    prodRows.forEach(r=>{
      const t1=hm2ms(r.debut),t2=hm2ms(r.fin);
      if(!t1) return;
      const x1=toX(t1),x2=toX(t2||tE);
      if(x2>x1) html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="#bbf7d0" rx="3"/>`;
    });
    evtRows.forEach(r=>{
      const t1=hm2ms(r.debut),t2=hm2ms(r.fin);
      if(!t1) return;
      const x1=toX(t1),x2=toX(t2||t1+1800000);
      if(x2<=x1) return;
      const tl=(r.type||'').toLowerCase();
      const col=tl.includes('nett')?'#38bdf8':tl.includes('pause')?'#94a3b8':tl.includes('ratt')?'#f59e0b':'#dc2626';
      html+=`<rect x="${x1}" y="${Y}" width="${x2-x1}" height="${H2}" fill="${col}" rx="2" opacity=".75"/>`;
    });
    const fmt=ms=>{const d=new Date(ms);return d.getHours().toString().padStart(2,'0')+':'+d.getMinutes().toString().padStart(2,'0');};
    // Hourly tick marks
    let tickT=Math.ceil(tS/3600000)*3600000;
    while(tickT<tE){
      const tx=toX(tickT);
      const hr=new Date(tickT).getHours();
      html+=`<line x1="${tx}" y1="${Y}" x2="${tx}" y2="${Y+H2}" stroke="rgba(0,0,0,.2)" stroke-width="1"/>`;
      html+=`<text x="${tx+2}" y="${Y+H2+9}" font-size="7" fill="#374151">${String(hr).padStart(2,'0')}h</text>`;
      tickT+=3600000;
    }
    html+=`<text x="2" y="${Y+H2+9}" font-size="8" fill="#374151">${fmt(tS)}</text>`;
    html+=`<text x="${W-30}" y="${Y+H2+9}" font-size="8" fill="#374151">${fmt(tE)}</text>`;
    return html;
  }
  // Métriques supplémentaires
  const totQteFab=(d.prod_rows||[]).reduce((s,r)=>s+parseFloat(r.qte_fab||0),0);
  const elapsedEffS=Math.max(1,(d.model_dur_s||0)-(d.planned_ded_s||0));
  const cadenceH=elapsedEffS>0?Math.round(totQteFab/elapsedEffS*3600):0;
  const sortedProdF=(d.prod_rows||[]).filter(r=>r.fibre).sort((a,b)=>(a.debut||'').localeCompare(b.debut||''));
  let nbChangFibre=0;for(let i=1;i<sortedProdF.length;i++){if(sortedProdF[i].fibre!==sortedProdF[i-1].fibre)nbChangFibre++;}
  const prodRef=d.prod_ref||200;
  const cadencePcsMin=prodRef/480;
  const budgetData=d.budget_data||{};
  const arretsPrevu=Object.values(budgetData).reduce((a,b)=>a+(b.budget_min||0),0);
  const tempsUtile=Math.round(totalMin)-Math.round(arretsPrevu);
  const perteCadenceMin=cadencePcsMin>0?Math.round((prodMin*cadencePcsMin-(d.tot_equiv||0))/cadencePcsMin):0;
  const tlDebut=d.actual_debut||d.model_debut;
  const tlFin=d.actual_fin||d.model_fin;
  const tlContent=buildTL(d.prod_rows||[],d.evt_rows||[],date,tlDebut,tlFin);
  const plageStr=(tlDebut&&tlFin)?(' · '+esc(tlDebut)+' → '+esc(tlFin)):'';
  // ── Panneau gauche : passe en mode KPI ──
  const leftList=document.getElementById('rpt-left-list');
  const leftKpi=document.getElementById('rpt-left-kpi');
  const leftExe=document.getElementById('rpt-left-exe');
  if(leftList&&leftKpi&&leftExe){
    leftList.style.display='none';
    leftExe.style.width='246px';
    leftKpi.style.cssText='display:flex;flex-direction:row;flex:1;overflow:hidden';
    leftKpi.innerHTML=`
      <!-- Fine bande retour (gauche) -->
      <div onclick="rptBackToList()" title="Retour à la liste"
        style="width:44px;background:#4f46e5;border-right:2px solid #3730a3;cursor:pointer;flex-shrink:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;transition:background .15s;box-shadow:2px 0 8px rgba(79,70,229,.25);animation:blink .85s step-start infinite"
        onmouseenter="this.style.background='#4338ca';this.style.animation='none'" onmouseleave="this.style.background='#4f46e5';this.style.animation='blink .85s step-start infinite'">
        <span style="font-size:20px;color:#fff;user-select:none;line-height:1">‹</span>
        <span style="font-size:12px;color:#fff;user-select:none;writing-mode:vertical-lr;transform:rotate(180deg);letter-spacing:.08em;font-weight:800">Afficher les autres</span>
      </div>
      <!-- Contenu KPI -->
      <div style="flex:1;overflow-y:auto;display:flex;flex-direction:column">
      <div style="background:var(--navy);color:#fff;padding:10px 12px;flex-shrink:0">
        <div style="font-size:calc(12px*var(--zf,1));font-weight:800;opacity:.9">${esc(poste)}${d.model_debut&&d.model_fin?' — '+esc(d.model_debut)+' → '+esc(d.model_fin):''}</div>
        <div style="font-size:calc(10px*var(--zf,1));opacity:.75;margin-top:2px">${esc(pilot)} · ${esc(date)}</div>
        <div style="font-size:calc(10px*var(--zf,1));opacity:.65;margin-top:6px;font-weight:600;text-transform:uppercase;letter-spacing:.05em">Taux de rendement :</div>
        <div style="font-size:calc(30px*var(--zf,1));font-weight:900;color:${trsCol};line-height:1.1">${trsS>=0?trsS.toFixed(1)+'%':'—'}</div>
      </div>
      <div style="padding:8px 10px;display:flex;flex-direction:column;gap:6px">
        <div style="text-align:center">
          <svg id="rpt-gauge" viewBox="0 0 100 58" style="width:190px;display:block;margin:0 auto">
            <path d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#dde4ef" stroke-width="12" stroke-linecap="round"/>
            <path id="rpt-gauge-arc" d="M8,50 A42,42 0 0,1 92,50" fill="none" stroke="#16a34a" stroke-width="12" stroke-linecap="round" stroke-dasharray="0,1000"/>
            <text x="50" y="46" text-anchor="middle" font-size="14" font-weight="800" fill="#1a1f5e" id="rpt-gauge-pct">--%</text>
          </svg>
          <div style="font-size:calc(9px*var(--zf,1));color:var(--gray);margin-top:2px">TRS Poste</div>
        </div>
        <div style="text-align:center">
          <svg id="rpt-pie" viewBox="0 0 130 130" style="width:170px;height:170px;display:block;margin:0 auto"></svg>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px">
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(18px*var(--zf,1));color:#059669;font-weight:900">${Math.round(totQteFab)}</div><div class="fp-lbl" style="font-size:calc(10px*var(--zf,1))">Pièces</div></div>
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(18px*var(--zf,1));color:#0891b2;font-weight:900">${Math.round(d.tot_equiv||0)}</div><div class="fp-lbl" style="font-size:calc(10px*var(--zf,1))">Équiv.</div></div>
          <div class="fp-card" style="padding:6px;display:flex;align-items:center;justify-content:space-between;gap:4px">
            <div><div class="fp-big" style="font-size:calc(18px*var(--zf,1));color:#0369a1;font-weight:900">${cadenceH}</div><div class="fp-lbl" style="font-size:calc(10px*var(--zf,1))">Cad./h</div></div>
            <div style="text-align:right"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#64748b">${cadencePcsMin.toFixed(2)}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">pcs/min</div></div>
          </div>
          <div class="fp-card" style="padding:6px;display:flex;align-items:center;justify-content:space-between;gap:4px">
            <div><div class="fp-big" style="font-size:calc(18px*var(--zf,1));color:#7c3aed;font-weight:900">${d.nb_of||0}</div><div class="fp-lbl" style="font-size:calc(10px*var(--zf,1))">Nb OF</div></div>
            <div style="text-align:right"><div style="font-size:calc(11px*var(--zf,1));font-weight:700;color:#8b5cf6">${nbChangFibre}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Chg. fibre</div></div>
          </div>
        </div>
        <div style="display:flex;flex-direction:column;gap:4px">
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(12px*var(--zf,1))">${(d.model_debut&&d.model_fin)?(d.model_debut+'→'+d.model_fin):(Math.round((d.model_dur_s||0)/60)+' min')}</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Temps d’ouverture</div></div>
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(12px*var(--zf,1));color:#059669">${tempsUtile} min</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Temps utile</div></div>
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(12px*var(--zf,1));color:#16a34a">${prodMin} min</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Temps de fonctionnement</div></div>
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(12px*var(--zf,1));color:#dc2626">${stopMin} min</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Temps en arrêt</div></div>
          <div class="fp-card" style="padding:6px"><div class="fp-big" style="font-size:calc(12px*var(--zf,1));color:#f59e0b">${perteCadenceMin} min</div><div class="fp-lbl" style="font-size:calc(9px*var(--zf,1))">Perte cadence</div></div>
        </div>
      </div>
      </div>`;
  }
  // ── Panneau droit : timeline + tables (pleine largeur) ──
  detailEl.innerHTML=`
    <!-- Timeline -->
    <div style="padding:5px 12px;background:var(--card);border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(9px*var(--zf,1));font-weight:700;text-transform:uppercase;color:var(--gray);margin-bottom:3px">Timeline${plageStr}</div>
      <svg viewBox="0 0 800 52" preserveAspectRatio="none" style="width:100%;height:52px;display:block">${tlContent}</svg>
      <div class="tl-legend"><span><i style="background:#dc2626"></i>Arrêt</span><span><i style="background:#f59e0b"></i>Rattrapage</span><span><i style="background:#38bdf8"></i>Nettoyage</span><span><i style="background:#94a3b8"></i>Pause</span><span><i style="background:#bbf7d0;border:1px solid #86efac"></i>Prod</span></div>
    </div>
    <!-- Productions (pleine largeur) -->
    <div style="padding:8px 12px;border-bottom:1px solid var(--border);flex-shrink:0">
      <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:6px">Productions</div>
      <table style="width:100%;border-collapse:collapse;font-size:calc(11px*var(--zf,1))">
        <thead><tr style="background:#f8fafc;border-bottom:1px solid var(--border)">
          <th style="padding:4px 6px;text-align:center">OF</th><th style="padding:4px 6px;text-align:center">Taille</th>
          <th style="padding:4px 6px;text-align:center">Lots×2</th><th style="padding:4px 6px;text-align:center">Qté</th><th style="padding:4px 6px;text-align:center">Éq.</th>
          <th style="padding:4px 6px;text-align:center">Heures</th><th style="padding:4px 6px;text-align:center;color:#16a34a">Durée prod</th><th style="padding:4px 6px;text-align:center;color:#dc2626">Durée arrêts</th><th style="padding:4px 6px;text-align:center">TRS</th><th style="padding:4px 6px;text-align:left">Comm.</th>
        </tr></thead>
        <tbody>${prodsHtml||'<tr><td colspan="10" style="padding:8px;text-align:center;color:var(--gray)">Aucune production</td></tr>'}</tbody>
      </table>
    </div>
    <!-- Pareto + Arrêts côte à côte -->
    <div style="flex:1;overflow-y:auto;padding:8px 12px;display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div style="display:flex;flex-direction:column;gap:8px">
        <div class="card" style="padding:10px">
          <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:8px">Pareto des arrêts</div>
          ${paretoHtml||'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'}
        </div>
        <div class="card" style="padding:10px">
          <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:#92400e;margin-bottom:8px">⏱ Arrêts prévus</div>
          ${(()=>{
            const bd=d.budget_data||{};
            const bKeys=['clean_short_min','clean_long_min','clean_grand_min','meeting_tol_min','pause_min'];
            const rows=bKeys.map(bk=>{
              const b=bd[bk];if(!b||b.budget_min<=0)return '';
              const pct=Math.min(100,Math.round(b.used_min/b.budget_min*100));
              const col=b.used_min>b.budget_min?'#dc2626':b.used_min/b.budget_min>=0.8?'#d97706':'#16a34a';
              return '<div style="margin-bottom:7px">'
                +'<div style="display:flex;justify-content:space-between;font-size:calc(11px*var(--zf,1));margin-bottom:3px">'
                +'<span style="color:var(--text)">'+esc(b.label)+'</span>'
                +'<span style="font-weight:700;color:'+col+'">'+Math.round(b.used_min)+'/'+Math.round(b.budget_min)+' min</span>'
                +'</div>'
                +'<div style="background:#f1f5f9;border-radius:4px;height:12px;overflow:hidden">'
                +'<div style="height:100%;background:'+col+';border-radius:4px;width:'+pct+'%;opacity:.85"></div>'
                +'</div></div>';
            }).filter(Boolean).join('');
            return rows||'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun budget configuré</div>';
          })()}
        </div>
      </div>
      <div class="card" style="padding:10px;display:flex;flex-direction:column;gap:6px">
        <div style="font-size:calc(11px*var(--zf,1));font-weight:800;text-transform:uppercase;color:var(--gray);margin-bottom:6px">Détail arrêts</div>
        ${(d.evt_rows||[]).length?`<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:calc(10px*var(--zf,1))">
          <thead><tr style="background:#f8fafc;border-bottom:1px solid var(--border)">
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray);white-space:nowrap">Arrêt</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">OF</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Type</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray);white-space:nowrap">Plage</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Durée</th>
            <th style="padding:3px 5px;text-align:left;font-weight:700;color:var(--gray)">Commentaire</th>
          </tr></thead>
          <tbody>${(d.evt_rows||[]).map(r=>`<tr style="border-bottom:1px solid var(--border)">
            <td style="padding:4px 5px;font-weight:600;white-space:nowrap;max-width:90px;overflow:hidden;text-overflow:ellipsis">${esc(r.type||'')}</td>
            <td style="padding:4px 5px;color:#0369a1;font-weight:700">${esc(r.of||'—')}</td>
            <td style="padding:4px 5px;color:var(--text)">${esc(r.type_prod||'—')}</td>
            <td style="padding:4px 5px;white-space:nowrap;color:var(--gray)">${esc(r.debut||'')} → ${esc(r.fin||'')}</td>
            <td style="padding:4px 5px;font-weight:700;white-space:nowrap">${esc(r.duree||'')}</td>
            <td style="padding:4px 5px;color:var(--gray);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(r.comment||'')}">${esc(r.comment||'—')}</td>
          </tr>`).join('')}</tbody>
        </table></div>`:'<div style="color:var(--gray);font-size:calc(12px*var(--zf,1))">Aucun arrêt</div>'}
      </div>
    </div>`;
  // Dessiner gauge et pie (éléments maintenant dans le DOM)
  drawGauge('rpt-gauge-arc','rpt-gauge-pct',trsS>=0?trsS:0);
  drawPie('rpt-pie',[
    {label:'Prod',value:prodMin,color:'#16a34a'},
    {label:'Arrêts',value:stopMin,color:'#dc2626'},
    {label:'Autre',value:Math.max(0,totalMin-prodMin-stopMin),color:'#94a3b8'}
  ],{fCenter:16,fSub:10,fLeg:10,legY:118});
}

function rptBackToList(){
  const leftExe=document.getElementById('rpt-left-exe');
  const leftList=document.getElementById('rpt-left-list');
  const leftKpi=document.getElementById('rpt-left-kpi');
  const detailEl=document.getElementById('rpt-detail');
  if(leftExe) leftExe.style.width='280px';
  if(leftList) leftList.style.display='flex';
  if(leftKpi){leftKpi.style.display='none';leftKpi.innerHTML='';}
  if(detailEl) detailEl.innerHTML='<div style="padding:24px;color:var(--gray);text-align:center">Sélectionnez un poste</div>';
}

// ── SETTINGS ──
async function setDbPath(){
  const path=document.getElementById('cfg-db-path').value.trim();
  if(!path){toast('Chemin requis','err');return;}
  const st=document.getElementById('cfg-db-status');
  if(st){st.textContent='Enregistrement…';st.style.color='var(--amber)';}
  const r=await fetch('/api/set_db',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,path})});
  const d=r?await r.json():{};
  if(d&&d.ok){
    if(st){st.textContent='✓ Fichier Excel configuré';st.style.color='var(--green)';}
    toast('Fichier Excel enregistré','ok');
    loadCfg();
  } else {
    if(st){st.textContent='✗ '+(d&&d.error||'Erreur — vérifiez le chemin');st.style.color='var(--red)';}
    toast(d&&d.error||'Chemin invalide','err');
  }
}

let _cfgArretsPrevus = {};

async function loadCfg(){
  const d=await apiFetch('/api/config');
  if(!d) return;
  _cfgPwds=d.pilot_passwords||{};
  _cfgModels=d.modeles_horaires||[];
  _cfgModelsBase=d.modeles_horaires_base||JSON.parse(JSON.stringify(_cfgModels));
  _cfgArretsPrevus=d.arrets_prevus||{};
  const prEl=document.getElementById('cfg-pr');
  if(prEl) prEl.value=d.prod_ref||200;
  // Show current db path
  const dbEl=document.getElementById('cfg-db-path');
  const dbSt=document.getElementById('cfg-db-status');
  if(dbEl&&d.db_path) dbEl.value=d.db_path;
  if(dbSt&&d.db_name){dbSt.textContent='Fichier actuel : '+d.db_name;dbSt.style.color='var(--green)';}
  // Arrêts prévus
  const apMap={'ap-clean-short':'clean_short_min','ap-clean-long':'clean_long_min','ap-clean-grand':'clean_grand_min','ap-meeting':'meeting_tol_min','ap-pause':'pause_min'};
  Object.entries(apMap).forEach(([elId,key])=>{const el=document.getElementById(elId);if(el)el.value=(_cfgArretsPrevus[key]||0);});
  renderPwdList();
  renderModelList();
  await loadEvtsList();
  await loadInterposteCfg();
  // Refresh login model dropdown
  const sel=document.getElementById('ln-model');
  if(sel){
    while(sel.options.length>1) sel.remove(1);
    _cfgModels.forEach(m=>{const o=document.createElement('option');o.value=m.nom||'';o.textContent=m.nom||'';sel.appendChild(o);});
  }
  updateAccModelInfo();
}

function updateAccModelInfo(){
  const el=document.getElementById('acc-model-info');
  if(!el) return;
  const poste=ST&&ST.poste;
  if(!poste){el.style.display='none';return;}
  const DAY_KEYS=['dim','lun','mar','mer','jeu','ven','sam'];
  const dk=DAY_KEYS[new Date().getDay()];
  const mEff=_cfgModels.find(m=>m.nom===poste);
  const mBase=_cfgModelsBase.find(m=>m.nom===poste);
  const jEff=mEff&&mEff.jours&&mEff.jours[dk]||{};
  const jBase=mBase&&mBase.jours&&mBase.jours[dk]||{};
  if(!jEff.debut&&!jEff.fin){el.style.display='none';return;}
  const isOverridden=(jEff.debut!==jBase.debut)||(jEff.fin!==jBase.fin);
  el.style.display='block';
  if(isOverridden){
    el.innerHTML=`⏰ <b>${esc(poste)}</b> : <b style="color:#d97706">${esc(jEff.debut)} → ${esc(jEff.fin)}</b> <span style="background:#fef3c7;color:#92400e;font-size:calc(10px*var(--zf,1));padding:1px 5px;border-radius:4px;font-weight:700">⚠ temporaire</span>`;
  } else {
    el.innerHTML=`⏰ <b>${esc(poste)}</b> : ${esc(jEff.debut)} → ${esc(jEff.fin)}`;
  }
}

async function saveArretsPrevus(){
  const clean_short=parseFloat(document.getElementById('ap-clean-short')?.value||0)||0;
  const clean_long=parseFloat(document.getElementById('ap-clean-long')?.value||0)||0;
  const clean_grand=parseFloat(document.getElementById('ap-clean-grand')?.value||0)||0;
  const meeting=parseFloat(document.getElementById('ap-meeting')?.value||0)||0;
  const pause_m=parseFloat(document.getElementById('ap-pause')?.value||0)||0;
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,clean_short_min:clean_short,clean_long_min:clean_long,clean_grand_min:clean_grand,meeting_tol_min:meeting,pause_min:pause_m})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Arrêts prévus enregistrés','ok');await loadCfg();}
  else toast(d?.error||'Erreur','err');
}

function renderPwdList(){
  const c=document.getElementById('pwd-list');
  if(!c) return;
  c.innerHTML=Object.entries(_cfgPwds).map(([nm,pw])=>`
    <div class="pr">
      <div class="pn">${esc(nm)}</div>
      <input type="password" id="pwi-${esc(nm)}" value="${esc(String(pw))}" data-n="${esc(nm)}">
      <button class="btn-eye" onclick="toggleEye('pwi-${esc(nm)}')">👁</button>
      <button class="btn btn-sec" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="rmPilot('${esc(nm)}')">✕</button>
    </div>`).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1));padding:3px">Aucun pilote configuré</div>';
}

function toggleEye(id){const i=document.getElementById(id);if(i)i.type=i.type==='password'?'text':'password';}
function addPilot(){
  const n=document.getElementById('np-name').value.trim(),pw=document.getElementById('np-pw').value;
  if(!n){toast('Nom requis','err');return;}
  _cfgPwds[n]=pw;
  document.getElementById('np-name').value='';document.getElementById('np-pw').value='';
  renderPwdList();
}
function rmPilot(n){delete _cfgPwds[n];renderPwdList();}
async function savePwds(){
  document.querySelectorAll('#pwd-list input[data-n]').forEach(i=>_cfgPwds[i.dataset.n]=i.value);
  const r=await fetch('/api/pilot_passwords_excel',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,pilot_passwords:_cfgPwds})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('MDP enregistrés et sauvegardés dans Excel','ok');loadCfg();}else toast(d&&d.error||'Erreur','err');
}

const DAYS=[{k:'lun',l:'Lun'},{k:'mar',l:'Mar'},{k:'mer',l:'Mer'},{k:'jeu',l:'Jeu'},{k:'ven',l:'Ven'},{k:'sam',l:'Sam'},{k:'dim',l:'Dim'}];
function renderModelList(){
  const c=document.getElementById('models-list');
  if(!c) return;
  c.innerHTML=_cfgModelsBase.map((m,mi)=>{
    const j=m.jours||{};
    return `<div class="model-card">
      <div class="mch">
        <input value="${esc(m.nom||'Poste '+(mi+1))}" onchange="_cfgModelsBase[${mi}].nom=this.value" placeholder="Nom du poste">
        <button class="btn btn-danger" style="font-size:calc(10px*var(--zf,1));padding:2px 6px" onclick="_cfgModelsBase.splice(${mi},1);renderModelList()">✕</button>
      </div>
      <div class="day-grid">${DAYS.map(d=>{const dc=j[d.k]||{};
        return `<div class="day-box"><div class="day-lbl">${d.l}</div>
          <input type="time" onchange="setDay(${mi},'${d.k}','debut',this.value)" value="${dc.debut||'05:00'}" style="margin-bottom:2px">
          <input type="time" onchange="setDay(${mi},'${d.k}','fin',this.value)" value="${dc.fin||'13:00'}">
        </div>`;}).join('')}
      </div>
    </div>`;
  }).join('')||'<div style="color:var(--gray);font-size:calc(11px*var(--zf,1))">Aucun modèle horaire</div>';
}

function setDay(mi,day,field,val){if(!_cfgModelsBase[mi])return;if(!_cfgModelsBase[mi].jours)_cfgModelsBase[mi].jours={};if(!_cfgModelsBase[mi].jours[day])_cfgModelsBase[mi].jours[day]={};_cfgModelsBase[mi].jours[day][field]=val;}
function addModel(){_cfgModelsBase.push({nom:'Nouveau poste',jours:{lun:{debut:'05:00',fin:'13:00'},mar:{debut:'05:00',fin:'13:00'},mer:{debut:'05:00',fin:'13:00'},jeu:{debut:'05:00',fin:'13:00'},ven:{debut:'05:00',fin:'13:00'},sam:{debut:'',fin:''},dim:{debut:'',fin:''}}});renderModelList();}
async function saveModels(){
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,modeles_horaires:_cfgModelsBase})});
  const d=r?await r.json():{};
  if(d&&d.ok){toast('Modèles enregistrés','ok');loadCfg();}else toast(d&&d.error||'Erreur','err');
}
async function saveProdRef(){
  const v=parseFloat(document.getElementById('cfg-pr').value)||200;
  const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pw:_adminPw,prod_ref:v})});
  const d=r?await r.json():{};
  d&&d.ok?toast('Enregistré','ok'):toast(d&&d.error||'Erreur','err');
}

async function generateDashboard(){
  const st=document.getElementById('dash-status');
  if(st){st.textContent='Génération en cours…';st.style.color='var(--amber)';}
  let d={};
  try{
    const r=await fetch('/api/generate_dashboard',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
    if(r&&r.ok) d=await r.json();
  }catch(e){}
  if(d&&d.ok){
    if(st){st.textContent='✓ Dashboard prêt — cliquez pour ouvrir';st.style.color='var(--green)';}
    toast('Dashboard généré !','ok');
    // Ouvrir via Flask (même origine, pas de CORS)
    window.open('http://127.0.0.1:5001/dashboard','_blank');
  } else {
    const errMsg=d&&d.error?d.error:'Erreur inconnue — vérifiez que le fichier Excel est configuré dans les paramètres';
    if(st){st.textContent='✗ '+errMsg;st.style.color='var(--red)';}
    toast('Erreur dashboard: '+errMsg,'err');
  }
}

// ── MODALS ──
function openM(id){const m=document.getElementById(id);if(m){m.classList.add('on');m.style.display='flex';}}
function closeM(id){const m=document.getElementById(id);if(m){m.classList.remove('on');m.style.display='';}}
document.addEventListener('click',e=>{if(e.target.classList.contains('overlay'))closeM(e.target.id);});

// ── UTILS ──
async function apiFetch(url,retries=2){
  try{
    const r=await fetch(url);
    if(r.status===429){
      if(retries>0){await new Promise(res=>setTimeout(res,1200));return apiFetch(url,retries-1);}
      return null;
    }
    return r.ok?await r.json():null;
  }catch(e){return null;}
}
function showFibre(name){
  if(!name)return;
  let m=document.getElementById('m-fibre-info');
  if(!m){
    m=document.createElement('div');
    m.id='m-fibre-info';
    m.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:900;display:flex;align-items:center;justify-content:center';
    m.innerHTML='<div style="background:#fff;border-radius:14px;padding:28px 32px;min-width:260px;max-width:420px;box-shadow:0 20px 60px rgba(0,0,0,.3);text-align:center">'+
      '<div style="font-size:calc(11px*var(--zf,1));font-weight:700;text-transform:uppercase;color:#6366f1;letter-spacing:.6px;margin-bottom:8px">Fibre</div>'+
      '<div id="m-fibre-name" style="font-size:calc(18px*var(--zf,1));font-weight:800;color:#1e293b;word-break:break-all;margin-bottom:18px"></div>'+
      '<button onclick="document.getElementById(\'m-fibre-info\').style.display=\'none\'" style="background:#6366f1;color:#fff;border:none;border-radius:8px;padding:8px 24px;font-size:calc(13px*var(--zf,1));font-weight:700;cursor:pointer">Fermer</button>'+
      '</div>';
    m.onclick=e=>{if(e.target===m)m.style.display='none';};
    document.body.appendChild(m);
  }
  document.getElementById('m-fibre-name').textContent=name;
  m.style.display='flex';
}
function fmtDur2(s){if(!s||s<0)return'0:00';const m=Math.floor(s/60),sec=Math.floor(s%60);return m+':'+String(sec).padStart(2,'0');}
function fmtTRSv(v){return(v===null||v===undefined||isNaN(v)||v<0)?'--%':parseFloat(v).toFixed(1)+'%';}
function fmtDur(s){if(!s||s<0)return'00:00:00';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),sec=Math.floor(s%60);return[h,m,sec].map(x=>String(x).padStart(2,'0')).join(':');}
function fmtD2(s){if(!s||s<0)return'0 min';const h=Math.floor(s/3600),m=Math.floor((s%3600)/60);return h?h+'h'+String(m).padStart(2,'0'):m+' min';}
function fmtDurMS(s){s=Math.round(s||0);const m=Math.floor(s/60),sec=s%60;return m>0?m+' min'+(sec?' '+sec+'s':''):sec+'s';}
function fmtDurShort(s){s=Math.round(s||0);const m=Math.floor(s/60),sec=s%60;return m>0?m+'m'+(sec?' '+String(sec).padStart(2,'0')+'s':''):sec+'s';}
function fmtTRS(v){return(v===null||v===undefined||isNaN(v))?'--%':parseFloat(v).toFixed(1)+'%';}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function toast(msg,type){
  let t=document.getElementById('_toast');
  if(!t){t=document.createElement('div');t.id='_toast';t.style.cssText='position:fixed;bottom:16px;right:16px;padding:8px 14px;border-radius:7px;font-size:calc(13px*var(--zf,1));font-weight:600;z-index:999;transition:opacity .3s;box-shadow:0 4px 12px rgba(0,0,0,.18)';document.body.appendChild(t);}
  t.textContent=msg;t.style.background=type==='ok'?'#16a34a':'#dc2626';t.style.color='#fff';t.style.opacity='1';
  clearTimeout(t._to);t._to=setTimeout(()=>t.style.opacity='0',3000);
}
</script>
</body>
</html>"""
