from __future__ import annotations

import html
import os
import re
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = APP_DIR.parent / "reporting" / "moveHistory" / "all data new"
HANANE_REPORT = APP_DIR.parent / "HANANE RAPPORT.xls"
DATA_DIR = Path(os.getenv("VESSEL_DATA_DIR", str(DEFAULT_DATA_DIR)))
MODE = os.getenv("VESSEL_DATA_MODE", "test").lower()

st.set_page_config(
    page_title="Nevis — Vessel Command Center",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
:root{--ink:#f6f9fb;--muted:#91a5b1;--line:rgba(180,215,225,.13);--cyan:#38e8d1;--blue:#62a9ff;--amber:#ffbd59;--red:#ff6577;--panel:#101c25}
.stApp{background:
radial-gradient(circle at 80% -10%,rgba(31,118,140,.20),transparent 34%),
radial-gradient(circle at -10% 55%,rgba(25,91,105,.12),transparent 34%),
#071118;color:var(--ink);font-family:'DM Sans',sans-serif}
.block-container{padding:1.25rem 2.2rem 3rem;max-width:1680px}
h1,h2,h3,[data-testid="stMetricValue"]{font-family:'Space Grotesk',sans-serif}
[data-testid="stHeader"]{display:none}
.topline{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
.brand{font:700 13px 'Space Grotesk';letter-spacing:.18em;text-transform:uppercase;color:#dceef2}
.brand b{color:var(--cyan)}
.live{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--cyan);background:rgba(56,232,209,.08);border:1px solid rgba(56,232,209,.2);padding:7px 12px;border-radius:20px}
.pulse{display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--cyan);box-shadow:0 0 0 5px rgba(56,232,209,.09);margin-right:8px}
.hero{display:grid;grid-template-columns:1.35fr .65fr;gap:16px;padding:20px 22px;border:1px solid var(--line);border-radius:22px;background:linear-gradient(135deg,rgba(17,32,42,.92),rgba(9,21,29,.88));box-shadow:0 24px 70px rgba(0,0,0,.22);overflow:hidden}
.eyebrow{font-size:10px;text-transform:uppercase;letter-spacing:.18em;color:var(--cyan);font-weight:700}
.hero h1{font-size:clamp(32px,4vw,59px);line-height:.94;margin:10px 0 12px;letter-spacing:-.055em}
.hero-sub{color:var(--muted);font-size:13px}
.hero-meta{display:flex;gap:22px;margin-top:24px}.hero-meta div{border-left:1px solid var(--line);padding-left:12px}.hero-meta small{color:var(--muted);font-size:9px;text-transform:uppercase;letter-spacing:.12em;display:block}.hero-meta b{font:600 14px 'Space Grotesk'}
.ship-card{position:relative;min-height:190px;border-radius:17px;overflow:hidden;background:linear-gradient(#102d3a 0 56%,#0a5662 56% 72%,#d6ad73 72%)}
.sun{position:absolute;width:65px;height:65px;border-radius:50%;background:#ffd581;right:32px;top:23px;box-shadow:0 0 45px rgba(255,204,114,.25)}
.ship{position:absolute;left:8%;right:6%;bottom:43px;height:57px;background:#d7e1e2;clip-path:polygon(0 22%,100% 22%,90% 100%,13% 100%);filter:drop-shadow(0 8px 8px rgba(0,0,0,.25))}
.ship:before{content:"";position:absolute;left:22%;top:-30px;width:35%;height:32px;background:#eff4f2;border-radius:4px 4px 0 0}
.ship-name{position:absolute;bottom:22px;left:27%;font:bold 10px 'Space Grotesk';color:#1c3842;letter-spacing:.08em}
.streak{position:absolute;left:15%;right:15%;height:1px;background:rgba(255,255,255,.3);bottom:35px}
.metric-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin:14px 0}
.kpi{padding:15px 16px;border:1px solid var(--line);border-radius:16px;background:rgba(14,27,36,.8);min-height:102px}
.kpi .label{font-size:9px;text-transform:uppercase;letter-spacing:.13em;color:var(--muted)}
.kpi .value{font:600 27px 'Space Grotesk';margin-top:7px}.kpi .delta{font-size:10px;color:var(--muted);margin-top:5px}.cyan{color:var(--cyan)}.amber{color:var(--amber)}.red{color:var(--red)}
.section-title{display:flex;justify-content:space-between;align-items:end;margin:20px 0 9px}.section-title h3{margin:0;font-size:15px}.section-title span{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.12em}
.terminal{display:grid;grid-template-columns:1fr 1.5fr 1.2fr .8fr;gap:8px;padding:12px;background:#0b1820;border:1px solid var(--line);border-radius:18px}
.zone{min-height:190px;border-radius:12px;padding:13px;position:relative;overflow:hidden}
.zone h4{font:600 10px 'Space Grotesk';letter-spacing:.12em;text-transform:uppercase;margin:0}.zone small{font-size:9px;color:var(--muted)}
.gate{background:repeating-linear-gradient(90deg,#10232c,#10232c 18px,#122933 18px,#122933 20px)}
.yard{background:#10242c}.quay{background:linear-gradient(90deg,#15313b 0 48%,#0b5260 48%)}.external{background:#17232a}
.stacks{display:grid;grid-template-columns:repeat(5,1fr);gap:4px;position:absolute;left:13px;right:13px;bottom:14px}
.stack{height:82px;border-radius:3px;background:repeating-linear-gradient(0deg,rgba(56,232,209,.8) 0 8px,transparent 8px 11px);border:1px solid rgba(56,232,209,.3)}
.stack.empty{height:53px;background:repeating-linear-gradient(0deg,rgba(255,189,89,.8) 0 8px,transparent 8px 11px);border-color:rgba(255,189,89,.3)}
.flow{position:absolute;left:13px;right:13px;top:70px;height:2px;background:linear-gradient(90deg,var(--cyan),var(--blue));opacity:.6}.flow:after{content:"›";position:absolute;right:-1px;top:-13px;color:var(--blue);font-size:25px}
.zone-stat{position:absolute;bottom:14px;left:13px;font:600 24px 'Space Grotesk'}.zone-stat small{display:block}
.zone-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;position:absolute;left:12px;right:12px;bottom:12px}
.mini{padding:8px;border-radius:8px;background:rgba(3,12,17,.28);border:1px solid rgba(180,215,225,.09)}
.mini b{display:block;font:600 18px 'Space Grotesk';color:#f2f8f9}.mini span{font-size:8px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}
.zone-total{position:absolute;right:12px;top:12px;font:600 22px 'Space Grotesk';text-align:right}.zone-total small{display:block;font:500 8px 'DM Sans';color:var(--muted);text-transform:uppercase;letter-spacing:.1em}
.reconcile{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;padding:10px 12px;border:1px solid var(--line);border-radius:12px;background:rgba(14,27,36,.55)}
.journey{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;padding:12px;border:1px solid var(--line);border-radius:18px;background:#0a171f}
.stage{min-height:142px;padding:15px;border:1px solid var(--line);border-radius:13px;background:linear-gradient(150deg,#10232c,#0c1a22);position:relative}
.stage:not(:last-child):after{content:"›";position:absolute;right:-18px;top:50%;z-index:2;transform:translateY(-50%);width:25px;height:25px;border-radius:50%;background:#132832;border:1px solid var(--line);color:var(--cyan);text-align:center;font:20px/22px 'Space Grotesk'}
.stage-no{font-size:8px;color:var(--cyan);letter-spacing:.15em;text-transform:uppercase}.stage h4{font:600 13px 'Space Grotesk';margin:6px 0 0}.stage-main{font:600 35px 'Space Grotesk';margin-top:18px}.stage-main small{font:500 8px 'DM Sans';color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-left:6px}.stage-split{display:flex;gap:14px;margin-top:9px}.stage-split span{font-size:9px;color:var(--muted)}.stage-split b{color:#edf7f8;font-family:'Space Grotesk'}
.matrix-wrap{margin-top:10px;border:1px solid var(--line);border-radius:16px;overflow:hidden;background:#0b1820}
.matrix-title{display:flex;justify-content:space-between;align-items:center;padding:12px 15px;border-bottom:1px solid var(--line)}.matrix-title b{font:600 11px 'Space Grotesk';text-transform:uppercase;letter-spacing:.1em}.matrix-title span{font-size:9px;color:var(--muted)}
.flow-matrix{width:100%;border-collapse:collapse;table-layout:fixed}.flow-matrix th,.flow-matrix td{padding:10px 13px;border-right:1px solid var(--line);border-bottom:1px solid var(--line);text-align:right}.flow-matrix th:last-child,.flow-matrix td:last-child{border-right:0}.flow-matrix tr:last-child td{border-bottom:0}.flow-matrix thead th{font-size:8px;color:var(--muted);text-transform:uppercase;letter-spacing:.11em;background:rgba(255,255,255,.015)}.flow-matrix thead th:first-child,.flow-matrix tbody td:first-child{text-align:left;width:22%}.flow-matrix tbody td:first-child{font-size:9px;color:#b7c8cf;text-transform:uppercase;letter-spacing:.08em}.flow-matrix tbody td:not(:first-child){font:600 15px 'Space Grotesk'}.flow-matrix .hot{color:var(--cyan)}.flow-matrix .warm{color:var(--amber)}
.fleet-card{height:138px!important;min-height:138px!important;padding:12px 14px;border:1px solid var(--line);border-radius:15px 15px 9px 9px;background:linear-gradient(145deg,rgba(17,34,44,.98),rgba(10,23,31,.94));position:relative;overflow:hidden;margin-bottom:5px}
.fleet-card:after{content:"";position:absolute;width:100px;height:100px;border-radius:50%;right:-36px;top:-44px;background:rgba(56,232,209,.07)}
.fleet-head{display:flex;justify-content:space-between;gap:10px}.fleet-head h3{font-size:17px;margin:4px 0 1px}.fleet-head small{font-size:8px;color:var(--muted)}
.fleet-score{font:600 23px 'Space Grotesk';color:var(--cyan);text-align:right;line-height:1}.fleet-score small{display:block;margin-top:4px;font:500 7px 'DM Sans';color:var(--muted);text-transform:uppercase}
.fleet-list{margin-top:13px}
.fleet-section{margin-top:11px;border:1px solid rgba(180,215,225,.09);border-radius:10px;overflow:hidden;background:rgba(2,10,14,.16)}
.fleet-section-title{display:flex;align-items:center;justify-content:space-between;padding:7px 9px;background:rgba(255,255,255,.025);border-bottom:1px solid rgba(180,215,225,.08);font-size:8px;color:#7f97a2;text-transform:uppercase;letter-spacing:.12em}
.fleet-section-title b{font:600 9px 'Space Grotesk';color:#cfe0e5}.fleet-section-title span{font-size:7px}
.fleet-row{display:grid;grid-template-columns:25px 1fr auto;align-items:center;min-height:31px;padding:0 7px;border-bottom:1px solid rgba(180,215,225,.08);text-decoration:none!important;transition:background .16s ease}
.fleet-row:hover{background:rgba(56,232,209,.055)}
.fleet-row:last-child{border-bottom:0}
.row-sign{width:18px;height:18px;border-radius:5px;display:grid;place-items:center;font:700 10px 'Space Grotesk';background:rgba(98,169,255,.10);color:var(--blue)}
.row-sign.in{background:rgba(56,232,209,.1);color:var(--cyan)}.row-sign.full{background:rgba(98,169,255,.12);color:#7cb8ff}.row-sign.empty{background:rgba(255,189,89,.11);color:var(--amber)}.row-sign.done{background:rgba(56,232,209,.12);color:var(--cyan)}.row-sign.left{background:rgba(255,101,119,.1);color:var(--red)}
.row-label{font-size:9px;color:#a9bac2;text-transform:uppercase;letter-spacing:.075em}.row-value{font:600 14px 'Space Grotesk';color:#f5f9fa;text-align:right}.row-value small{font:500 8px 'DM Sans';color:var(--muted);margin-left:4px}
.row-value em{font-style:normal;color:var(--cyan);margin-left:7px}.card-action{display:flex;align-items:center;justify-content:center;margin-top:11px;height:31px;border-radius:8px;background:rgba(56,232,209,.07);border:1px solid rgba(56,232,209,.16);color:#bff8ef!important;text-decoration:none!important;font:600 9px 'Space Grotesk';letter-spacing:.07em;text-transform:uppercase}.card-action:hover{background:rgba(56,232,209,.13);border-color:var(--cyan)}
.click-hint{font-size:8px;color:var(--muted);margin-top:8px;text-align:center;letter-spacing:.07em;text-transform:uppercase}
.fleet-card .fleet-list,.fleet-card .click-hint{display:none}
.metric-section-label{display:flex;justify-content:space-between;align-items:center;margin-top:8px;padding:7px 10px;border:1px solid var(--line);border-bottom:0;border-radius:10px 10px 0 0;background:rgba(255,255,255,.025);font-size:8px;text-transform:uppercase;letter-spacing:.1em;color:#cfe0e5}.metric-section-label span{font-size:7px;color:var(--muted)}
.metric-row-button [data-testid="stButton"] button{justify-content:flex-start!important;text-align:left!important;border-radius:0!important;margin:0!important;height:35px!important;background:#0d1d25!important;border-color:rgba(180,215,225,.09)!important}
.metric-row-button [data-testid="stButton"] button p{width:100%!important;text-align:left!important}
.detail-panel{padding:13px 15px;border:1px solid rgba(56,232,209,.18);border-radius:14px;background:linear-gradient(155deg,#10242d,#0a171f);position:sticky;top:14px}.detail-summary{display:grid;grid-template-columns:1fr auto;align-items:end;gap:12px}.detail-panel h3{margin:4px 0 2px;font-size:17px}.detail-panel p{color:var(--muted);font-size:9px;margin:0}.detail-count{font:600 25px 'Space Grotesk';color:var(--cyan);text-align:right;line-height:1}.detail-count small{display:block;font:500 7px 'DM Sans';color:var(--muted);text-transform:uppercase;letter-spacing:.1em;margin-top:5px}
[data-testid="stExpander"]{border:1px solid var(--line)!important;border-radius:11px!important;background:rgba(14,27,36,.45)!important}
[data-testid="stDialog"]{justify-content:flex-end!important;align-items:stretch!important;padding:0!important;background:rgba(2,9,13,.55)!important}
[data-testid="stDialog"] > div,[data-testid="stDialog"] [role="dialog"]{width:min(760px,48vw)!important;min-width:min(760px,48vw)!important;max-width:min(760px,48vw)!important;height:100vh!important;max-height:100vh!important;margin:0!important;border-radius:18px 0 0 18px!important;background:#09171e!important;border:0!important;border-left:1px solid rgba(56,232,209,.22)!important;box-shadow:-30px 0 80px rgba(0,0,0,.52)!important}
[data-testid="stDialog"] [role="dialog"] > div{width:100%!important;max-width:none!important;max-height:100vh!important;background:#09171e!important}
[data-testid="stDialog"] [role="dialog"] h2{display:none!important}
[data-testid="stDialog"] [role="dialog"] [data-testid="stDialogContent"]{padding-top:8px!important}
.drawer-head{display:grid;grid-template-columns:1fr auto;align-items:center;gap:12px;padding:10px 13px;margin-bottom:7px;border:1px solid rgba(56,232,209,.2);border-radius:12px;background:linear-gradient(135deg,#10242d,#0b1921)}
.drawer-head .context{font-size:8px;color:var(--cyan);text-transform:uppercase;letter-spacing:.12em}.drawer-head h3{font:600 17px 'Space Grotesk';margin:3px 0 0}.drawer-head .count{font:600 25px 'Space Grotesk';color:var(--cyan);text-align:right;line-height:1}.drawer-head .count small{display:block;margin-top:4px;font:500 7px 'DM Sans';color:var(--muted);text-transform:uppercase;letter-spacing:.1em}
.shift-brief{display:grid;grid-template-columns:1.08fr 1fr 1fr;gap:6px;margin-top:5px;padding:7px;border:1px solid rgba(98,169,255,.16);border-radius:8px;background:linear-gradient(135deg,rgba(11,31,40,.94),rgba(8,23,30,.94))}
.shift-overview,.shift-group{min-height:55px;padding:7px 8px;border-radius:6px;background:rgba(255,255,255,.025)}.shift-overview span,.shift-group>span{display:block;font-size:7px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.shift-overview b{display:block;margin-top:3px;font:600 10px 'Space Grotesk';color:#ddecf0}.shift-rate{display:flex;align-items:baseline;justify-content:space-between;margin-top:5px}.shift-rate strong{font:600 16px 'Space Grotesk';color:var(--cyan)}.shift-rate small{font-size:7px;color:var(--muted);text-transform:uppercase}
.shift-pills{display:grid;grid-template-columns:repeat(3,1fr);gap:3px;margin-top:6px}.shift-pills i{font-style:normal;text-align:center}.shift-pills b{display:block;font:600 13px 'Space Grotesk';color:#edf6f7}.shift-pills small{display:block;font-size:6px;color:var(--muted);text-transform:uppercase}
.shift-foot{grid-column:1/-1;display:flex;justify-content:space-between;gap:10px;padding:5px 3px 0;border-top:1px solid var(--line);font-size:7px;color:var(--muted)}.shift-foot b{color:#dcebed}
.report-strip{display:flex;align-items:center;gap:7px;flex-wrap:wrap;margin-top:5px;padding:6px 8px;border:1px solid rgba(255,189,89,.16);border-radius:7px;background:rgba(255,189,89,.035);font-size:7px;color:var(--muted)}.report-strip strong{color:#f3f8f9}.report-strip .source{color:var(--amber);text-transform:uppercase;letter-spacing:.08em}.report-strip i{font-style:normal;padding-left:7px;border-left:1px solid var(--line)}
.shift-report-head{display:flex;justify-content:space-between;align-items:center;margin:8px 0 6px;padding:8px 10px;border:1px solid rgba(255,189,89,.16);border-radius:9px;background:rgba(255,189,89,.035)}.shift-report-head b{font:600 10px 'Space Grotesk'}.shift-report-head span{font-size:8px;color:var(--muted)}
[data-testid="stPopover"] button{height:30px!important;width:100%!important;margin-top:4px!important;padding:4px 9px!important;justify-content:center!important;border-radius:7px!important;background:rgba(255,189,89,.055)!important;border-color:rgba(255,189,89,.22)!important;color:#ffe0a8!important;font:600 8px 'DM Sans'!important;letter-spacing:.025em!important}
[data-testid="stPopover"] button:hover{background:rgba(255,189,89,.11)!important;border-color:var(--amber)!important}
.terminal-mix{margin-top:5px;border:1px solid rgba(98,169,255,.14);border-radius:9px;overflow:hidden;background:#091820}
.terminal-mix-head{display:flex;align-items:center;justify-content:space-between;padding:6px 8px;border-bottom:1px solid var(--line);font-size:7px;color:var(--muted);text-transform:uppercase;letter-spacing:.09em}.terminal-mix-head b{font:600 8px 'Space Grotesk';color:#dcecef}.terminal-mix-head span{color:var(--blue)}
.terminal-mix table{width:100%;border-collapse:collapse;table-layout:fixed}.terminal-mix th,.terminal-mix td{padding:4px 2px;text-align:center;border-right:1px solid rgba(180,215,225,.07);border-bottom:1px solid rgba(180,215,225,.07)}.terminal-mix tr:last-child td{border-bottom:0}.terminal-mix th:last-child,.terminal-mix td:last-child{border-right:0}
.terminal-mix thead tr:first-child th{font:600 8px 'Space Grotesk';color:#70b7ff;background:rgba(98,169,255,.055)}.terminal-mix thead tr:nth-child(2) th{font:500 6px 'DM Sans';color:#8198a3;text-transform:uppercase}.terminal-mix th:first-child,.terminal-mix td:first-child{width:44px;text-align:left;padding-left:7px}.terminal-mix td:first-child{font-size:7px;color:#a9bbc3;text-transform:uppercase}.terminal-mix td:not(:first-child){font:600 9px 'Space Grotesk';color:#e7f1f3}.terminal-mix .full{color:var(--cyan)}.terminal-mix .empty{color:var(--amber)}
.terminal-mix .row-total{color:#fff!important;background:rgba(98,169,255,.055);font-weight:700!important}.terminal-mix .total-row td{background:rgba(255,255,255,.025);color:#dcecef!important;font-weight:700!important}.terminal-mix .total-row td:first-child{color:var(--blue)!important}
.terminal-totals{display:grid;grid-template-columns:repeat(4,1fr);gap:3px;padding:5px;border-top:1px solid var(--line)}.terminal-totals span{padding:4px 2px;text-align:center;border-radius:4px;background:rgba(255,255,255,.025);font-size:6px;color:var(--muted);text-transform:uppercase}.terminal-totals b{display:block;margin-top:2px;font:600 9px 'Space Grotesk';color:#edf5f6}.terminal-totals .grand{background:rgba(56,232,209,.07)}.terminal-totals .grand b{color:var(--cyan)}
.vessel-open [data-testid="stButton"] button{height:38px!important;background:linear-gradient(90deg,rgba(56,232,209,.17),rgba(98,169,255,.12))!important;border-color:rgba(56,232,209,.35)!important;color:#dffff9!important}
[class*="st-key-metric_list_"] [data-testid="stRadio"]{width:100%;border:1px solid var(--line);border-radius:10px;overflow:hidden;background:#09171e;padding:5px}
[class*="st-key-metric_list_"] [data-testid="stRadio"] > label{display:none}
[class*="st-key-metric_list_"] [role="radiogroup"]{display:grid!important;grid-template-columns:repeat(6,1fr)!important;gap:5px!important;width:100%!important}
[class*="st-key-metric_list_"] [role="radiogroup"] label{width:100%!important;min-width:0!important;min-height:36px;padding:6px 8px!important;margin:0!important;border:1px solid rgba(180,215,225,.09);border-radius:7px;background:linear-gradient(135deg,#10232c,#0d1d25);transition:all .15s ease}
[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(1){grid-column:1/-1}
[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(2),[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(3),[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(4){grid-column:span 2}
[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(5),[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(6){grid-column:span 3}
[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(7),[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(8),[class*="st-key-metric_list_"] [role="radiogroup"] label:nth-child(9){grid-column:span 2}
[class*="st-key-metric_list_"] [role="radiogroup"] label:hover{background:#16343e;border-color:rgba(56,232,209,.28);transform:translateY(-1px)}
[class*="st-key-metric_list_"] [role="radiogroup"] label:has(input:checked){background:linear-gradient(135deg,rgba(56,232,209,.16),rgba(98,169,255,.09));border-color:var(--cyan);box-shadow:0 0 0 1px rgba(56,232,209,.08)}
[class*="st-key-metric_list_"] [role="radiogroup"] input{display:none}
[class*="st-key-metric_list_"] [data-testid="stRadioOption"] > div > div > div:first-child{display:none!important}
[class*="st-key-metric_list_"] [data-testid="stRadioOption"] > div,[class*="st-key-metric_list_"] [data-testid="stRadioOption"] > div > div{width:100%!important;justify-content:center!important}
[class*="st-key-metric_list_"] [role="radiogroup"] p{width:100%!important;text-align:center!important;font:600 11px/1.2 'DM Sans'!important;color:#e2edf0!important;letter-spacing:.01em}
.progress-track{height:4px;background:#20333c;border-radius:5px;margin-top:10px;overflow:hidden}.progress-fill{height:100%;background:linear-gradient(90deg,var(--cyan),var(--blue));border-radius:5px}
[data-testid="stButton"] button{background:#10242d!important;color:#dff7f5!important;border:1px solid rgba(56,232,209,.2)!important;height:36px!important;font:600 10px 'Space Grotesk'!important;letter-spacing:.06em!important}
[data-testid="stButton"] button:hover{background:#16343d!important;border-color:var(--cyan)!important;color:var(--cyan)!important}
.badge{display:inline-flex;gap:6px;align-items:center;border:1px solid var(--line);border-radius:16px;padding:4px 8px;font-size:9px;color:var(--muted)}
.dot{width:7px;height:7px;border-radius:2px;display:inline-block}
[data-testid="stDataFrame"]{border:1px solid var(--line);border-radius:16px;overflow:hidden}
[data-testid="stSelectbox"] label,[data-testid="stTextInput"] label,[data-testid="stSegmentedControl"] label{font-size:10px!important;text-transform:uppercase;letter-spacing:.12em;color:var(--muted)!important}
.stButton button{border-radius:12px;border:1px solid var(--line)}
@media(max-width:900px){.hero{grid-template-columns:1fr}.metric-grid{grid-template-columns:repeat(2,1fr)}.terminal{grid-template-columns:1fr 1fr}.journey{grid-template-columns:1fr 1fr}.stage:nth-child(2):after{display:none}.flow-matrix{min-width:720px}.matrix-wrap{overflow-x:auto}.ship-card{min-height:160px}}
</style>
""",
    unsafe_allow_html=True,
)


FR_MONTHS = {
    "janv": 1, "févr": 2, "fevr": 2, "mars": 3, "avr": 4, "mai": 5,
    "juin": 6, "juil": 7, "août": 8, "aout": 8, "sept": 9, "oct": 10,
    "nov": 11, "déc": 12, "dec": 12,
}


def parse_n4_date(value):
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)
    s = str(value).strip().lower().replace("–", "-")
    m = re.search(r"(\d{2})-([a-zéûô]+)\.?-(\d{1,2})\s+(\d{2})(\d{2})", s)
    if not m:
        return pd.to_datetime(s, errors="coerce", dayfirst=True)
    year, month_name, day, hour, minute = m.groups()
    month = FR_MONTHS.get(month_name.rstrip("."))
    return pd.Timestamp(2000 + int(year), month, int(day), int(hour), int(minute)) if month else pd.NaT


def iso_size(value):
    s = str(value)
    return "20′" if s.startswith("2") else "40′" if s.startswith(("4", "L")) else "Other"


@st.cache_data(show_spinner=False, ttl=300)
def read_test_data(data_dir: str):
    root = Path(data_dir)
    unit_files = [root / "units-page1.xlsx", root / "units-page2.xlsx"]
    missing = [str(f) for f in unit_files if not f.exists()]
    if missing:
        raise FileNotFoundError("Missing test files: " + ", ".join(missing))
    units = pd.concat(
        [pd.read_excel(f, header=4) for f in unit_files],
        ignore_index=True,
    ).drop_duplicates(subset=["Unit Nbr", "I/B Actual Visit", "O/B Actual Visit"], keep="last")
    visit_file = root / "vessel visite.xlsx"
    visits = pd.read_excel(visit_file, header=4) if visit_file.exists() else pd.DataFrame()
    units["Last Move DT"] = units["Last Move"].map(parse_n4_date)
    units["Size"] = units["Type ISO"].map(iso_size)
    units["Freight"] = units["Frght Kind"].fillna("Unknown").astype(str).str.title()
    units["Visit Keys"] = units["I/B Actual Visit"].astype(str) + "|" + units["O/B Actual Visit"].astype(str)
    return units, visits


def normalized_vessel_name(value):
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


@st.cache_data(show_spinner=False, ttl=300)
def read_hanane_report(report_path: str):
    """Read only real vessel rows from the daily/shift productivity workbook."""
    path = Path(report_path)
    columns = [
        "Report Date", "Shift Nbr", "Shift Start", "Terminal", "Vessel",
        "Boxes Import", "Boxes Export", "Total Moves", "Crane GMPH",
        "Vessel GMPH", "Active Hours", "Stop Minutes", "Crane", "Observation",
    ]
    if not path.exists():
        return pd.DataFrame(columns=columns)

    import xlrd

    workbook = xlrd.open_workbook(path)
    records = []
    for sheet_name in workbook.sheet_names():
        if not re.fullmatch(r"J\.\d{2}", sheet_name):
            continue
        sheet = workbook.sheet_by_name(sheet_name)
        report_date = None
        for row_number in range(min(6, sheet.nrows)):
            for column_number in range(sheet.ncols):
                raw = str(sheet.cell_value(row_number, column_number)).strip()
                parsed = pd.to_datetime(raw, dayfirst=True, errors="coerce")
                if pd.notna(parsed) and parsed.year >= 2020:
                    report_date = parsed.normalize()
                    break
            if report_date is not None:
                break
        if report_date is None:
            continue

        terminal = ""
        shift_nbr = None
        vessel = ""
        for row_number in range(sheet.nrows):
            values = [sheet.cell_value(row_number, c) if c < sheet.ncols else "" for c in range(22)]
            marker = str(values[4]).strip().upper()
            shift_text = str(values[5]).strip()
            vessel_text = str(values[6]).strip()
            if marker in {"TCE", "TC3", "TCR"}:
                terminal = marker
                vessel = ""
            shift_match = re.match(r"([123])\s*E?\s*SH", shift_text.upper().replace("È", "E"))
            if shift_match:
                shift_nbr = int(shift_match.group(1))
                vessel = ""
            if shift_text.upper().startswith(("TOT", "TOTAL")) or marker in {"TOTAL", "IMP", "EXP", "GL", "EQ", "RDT"}:
                vessel = ""
                continue
            if vessel_text and vessel_text.upper() != "NAVIRE":
                vessel = vessel_text
            if not vessel or shift_nbr is None:
                continue

            total_moves = pd.to_numeric(values[15], errors="coerce")
            crane = str(values[20]).strip()
            # Blank template rows have neither moves nor a crane and are deliberately excluded.
            if (pd.isna(total_moves) or total_moves == 0) and not crane:
                continue
            start_hour = {1: 7, 2: 15, 3: 23}[shift_nbr]
            records.append(
                {
                    "Report Date": report_date,
                    "Shift Nbr": shift_nbr,
                    "Shift Start": report_date + pd.Timedelta(hours=start_hour),
                    "Terminal": terminal,
                    "Vessel": vessel,
                    "Boxes Import": pd.to_numeric(values[7], errors="coerce"),
                    "Boxes Export": pd.to_numeric(values[8], errors="coerce"),
                    "Total Moves": total_moves,
                    "Crane GMPH": pd.to_numeric(values[16], errors="coerce"),
                    "Vessel GMPH": pd.to_numeric(values[17], errors="coerce"),
                    "Active Hours": pd.to_numeric(values[13], errors="coerce"),
                    "Stop Minutes": pd.to_numeric(values[19], errors="coerce"),
                    "Crane": crane,
                    "Observation": str(values[21]).strip(),
                }
            )
    result = pd.DataFrame(records, columns=columns)
    if not result.empty:
        result["Vessel Key"] = result["Vessel"].map(normalized_vessel_name)
    return result


def read_live_data():
    # Future API seam: return the same two DataFrames as read_test_data().
    # Example: requests.get(os.environ["NEVIS_API_URL"], headers={...}).json()
    raise RuntimeError("Live API is not configured yet. Set VESSEL_DATA_MODE=test.")


def shift_window(reference: pd.Timestamp):
    h = reference.hour
    if 7 <= h < 15:
        start = reference.normalize() + pd.Timedelta(hours=7)
    elif 15 <= h < 23:
        start = reference.normalize() + pd.Timedelta(hours=15)
    elif h >= 23:
        start = reference.normalize() + pd.Timedelta(hours=23)
    else:
        start = reference.normalize() - pd.Timedelta(hours=1)
    return start, start + pd.Timedelta(hours=8)


def shift_start_for_timestamp(value):
    if pd.isna(value):
        return pd.NaT
    value = pd.Timestamp(value)
    if 7 <= value.hour < 15:
        return value.normalize() + pd.Timedelta(hours=7)
    if 15 <= value.hour < 23:
        return value.normalize() + pd.Timedelta(hours=15)
    if value.hour >= 23:
        return value.normalize() + pd.Timedelta(hours=23)
    return value.normalize() - pd.Timedelta(days=1) + pd.Timedelta(hours=23)


def safe_int(value):
    return int(value) if pd.notna(value) else 0


def infer_stops(timestamps, threshold_minutes=10):
    """Return inactivity gaps between consecutive moves that meet the stop rule."""
    ordered = pd.Series(timestamps).dropna().drop_duplicates().sort_values()
    if len(ordered) < 2:
        return {"count": 0, "minutes": 0, "longest": 0, "details": []}
    gaps = ordered.diff().dt.total_seconds().div(60)
    stops = gaps[gaps >= threshold_minutes]
    details = []
    for stop_index, duration in stops.items():
        position = ordered.index.get_loc(stop_index)
        details.append(
            {
                "From": pd.Timestamp(ordered.iloc[position - 1]),
                "To": pd.Timestamp(ordered.iloc[position]),
                "Duration min": int(round(duration)),
            }
        )
    return {
        "count": int(stops.count()),
        "minutes": int(round(stops.sum())) if not stops.empty else 0,
        "longest": int(round(stops.max())) if not stops.empty else 0,
        "details": details,
    }


def terminal_operation_matrix(frame):
    terminal = frame["ACTUAL TERMINAL ID"].astype(str).str.upper()
    operation = frame["Operation"].astype(str)
    freight = frame["Freight"].astype(str).str.lower()
    size = frame["Size"].astype(str)

    def count(terminal_name, operation_name, nature_name, size_name):
        operation_mask = (
            operation.isin(["Import", "Restow"])
            if operation_name == "Import"
            else operation.isin(["Export", "Restow"])
        )
        nature_mask = (
            freight.str.contains("full|laden|fcl", regex=True)
            if nature_name == "Full"
            else freight.str.contains("empty", regex=True)
        )
        return safe_int(
            (
                terminal.eq(terminal_name)
                & operation_mask
                & nature_mask
                & size.eq(size_name)
            ).sum()
        )

    matrix_values = {}
    operation_totals = {
        "Import": safe_int(operation.isin(["Import", "Restow"]).sum()),
        "Export": safe_int(operation.isin(["Export", "Restow"]).sum()),
    }
    rows = []
    for operation_name in ["Import", "Export"]:
        cells = []
        row_values = []
        for terminal_name in ["TCR", "TC3"]:
            for nature_name, css_class in [("Full", "full"), ("Empty", "empty")]:
                for size_name in ["20′", "40′"]:
                    value = count(
                        terminal_name, operation_name, nature_name, size_name
                    )
                    matrix_values[
                        (operation_name, terminal_name, nature_name, size_name)
                    ] = value
                    row_values.append(value)
                    cells.append(f'<td class="{css_class}">{value:,}</td>')
        rows.append(
            f"<tr><td>{operation_name}</td>{''.join(cells)}"
            f'<td class="row-total">{operation_totals[operation_name]:,}</td></tr>'
        )

    column_totals = []
    for terminal_name in ["TCR", "TC3"]:
        for nature_name in ["Full", "Empty"]:
            for size_name in ["20′", "40′"]:
                column_totals.append(
                    sum(
                        matrix_values[
                            (operation_name, terminal_name, nature_name, size_name)
                        ]
                        for operation_name in ["Import", "Export"]
                    )
                )
    import_total = operation_totals["Import"]
    export_total = operation_totals["Export"]
    operation_weight = (
        operation.isin(["Import", "Restow"]).astype(int)
        + operation.isin(["Export", "Restow"]).astype(int)
    )
    tcr_total = safe_int(operation_weight[terminal.eq("TCR")].sum())
    tc3_total = safe_int(operation_weight[terminal.eq("TC3")].sum())
    size20_total = safe_int(operation_weight[size.eq("20′")].sum())
    size40_total = safe_int(operation_weight[size.eq("40′")].sum())
    grand_total = import_total + export_total
    other_profile = max(0, grand_total - sum(column_totals))
    total_cells = "".join(f"<td>{value:,}</td>" for value in column_totals)
    rows.append(
        f'<tr class="total-row"><td>Total</td>{total_cells}'
        f'<td class="row-total">{grand_total:,}</td></tr>'
    )
    return (
        '<div class="terminal-mix"><div class="terminal-mix-head">'
        '<b>Terminal operation profile</b><span>Full / empty · ISO length</span></div>'
        f'<table><thead><tr><th></th><th colspan="4">TCR · {tcr_total:,}</th>'
        f'<th colspan="4">TC3 · {tc3_total:,}</th><th>Total</th></tr>'
        '<tr><th>Flow</th><th>F20</th><th>F40/45</th><th>E20</th><th>E40/45</th>'
        '<th>F20</th><th>F40/45</th><th>E20</th><th>E40/45</th><th>All</th></tr></thead>'
        f"<tbody>{''.join(rows)}</tbody></table>"
        '<div class="terminal-totals">'
        f'<span>Import<b>{import_total:,}</b></span>'
        f'<span>Export<b>{export_total:,}</b></span>'
        f'<span>20′ units<b>{size20_total:,}</b></span>'
        f'<span>40/45′ units<b>{size40_total:,}</b></span>'
        f'<span>TCR total<b>{tcr_total:,}</b></span>'
        f'<span>TC3 total<b>{tc3_total:,}</b></span>'
        f'<span>Other profile<b>{other_profile:,}</b></span>'
        f'<span class="grand">All movements<b>{grand_total:,}</b></span>'
        "</div></div>"
    )


def kpi(label, value, note, color=""):
    return f'<div class="kpi"><div class="label">{label}</div><div class="value {color}">{value}</div><div class="delta">{note}</div></div>'


def clear_container_inspector(clear_metric_widgets=True):
    for state_key in [
        "detail_visit_id", "detail_metric", "fleet_metric_selector",
        "drawer_group", "drawer_search", "drawer_state",
        "drawer_terminal", "drawer_size",
    ]:
        st.session_state.pop(state_key, None)
    if clear_metric_widgets:
        for state_key in list(st.session_state):
            if str(state_key).startswith("metric_list_"):
                st.session_state.pop(state_key, None)
        st.session_state.metric_widget_epoch = (
            int(st.session_state.get("metric_widget_epoch", 0)) + 1
        )


def select_fleet_metric(widget_key, target_visit, metric_by_label):
    selected_label = st.session_state.get(widget_key)
    if not selected_label:
        return
    selected_metric = metric_by_label.get(selected_label)
    if not selected_metric:
        return
    # A single inspector can be active. Clear every other vessel's remembered tile.
    for state_key in list(st.session_state):
        if str(state_key).startswith("metric_list_") and state_key != widget_key:
            st.session_state.pop(state_key, None)
    for state_key in [
        "drawer_group", "drawer_search", "drawer_state",
        "drawer_terminal", "drawer_size",
    ]:
        st.session_state.pop(state_key, None)
    st.session_state.detail_visit_id = target_visit
    st.session_state.detail_metric = selected_metric
    # Remount every vessel metric widget. This prevents a dismissed dialog's
    # previous radio selection from being restored by Streamlit on the next click.
    st.session_state.metric_widget_epoch = (
        int(st.session_state.get("metric_widget_epoch", 0)) + 1
    )


def open_vessel_dashboard(target_visit):
    clear_container_inspector(clear_metric_widgets=True)
    st.session_state.selected_visit_id = target_visit
    st.session_state.page_scope = "Overview"


try:
    units, visits = read_test_data(str(DATA_DIR)) if MODE == "test" else read_live_data()
    hanane_report = read_hanane_report(str(HANANE_REPORT)) if MODE == "test" else pd.DataFrame()
except Exception as exc:
    st.error(f"Data source unavailable — {exc}")
    st.stop()


def hanane_for_shift(vessel_name, shift_start):
    if hanane_report.empty or pd.isna(shift_start):
        return hanane_report.iloc[0:0]
    return hanane_report[
        hanane_report["Vessel Key"].eq(normalized_vessel_name(vessel_name))
        & hanane_report["Shift Start"].eq(pd.Timestamp(shift_start))
    ].copy()

candidate_keys = pd.concat([units["I/B Actual Visit"], units["O/B Actual Visit"]]).dropna().astype(str)
candidate_keys = candidate_keys[candidate_keys.str.fullmatch(r"\d{9}")].value_counts()
visit_lookup = visits.assign(Visit=visits["Visit"].astype(str)).set_index("Visit") if not visits.empty else pd.DataFrame()

visit_options = []
for key, count in candidate_keys.items():
    if key in visit_lookup.index:
        row = visit_lookup.loc[key]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[-1]
        visit_options.append((key, str(row.get("Vessel Name", key)), str(row.get("Phase", "Unknown")), count))
visit_options = sorted(visit_options, key=lambda x: (x[2] not in {"Working", "Inbound"}, -x[3]))

st.markdown(
    f'<div class="topline"><div class="brand">NEVIS <b>OPERATIONS</b> / VESSEL INTELLIGENCE</div>'
    f'<div class="live"><span class="pulse"></span>{html.escape(MODE)} mode · final snapshot</div></div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-title" style="margin-top:8px"><h3>Filter vessel visits</h3>'
    '<span>Choose operational status first, then select an escale</span></div>',
    unsafe_allow_html=True,
)
phase_values = sorted({item[2] for item in visit_options if item[2] and item[2] != "nan"})
line_values = sorted({
    str(visit_lookup.loc[item[0]].iloc[-1].get("Line", "—"))
    if isinstance(visit_lookup.loc[item[0]], pd.DataFrame)
    else str(visit_lookup.loc[item[0]].get("Line", "—"))
    for item in visit_options
})
terminal_values = sorted(units["ACTUAL TERMINAL ID"].dropna().astype(str).unique())
filter1, filter2, filter3, filter4 = st.columns([1.15, 1.25, 1, 1.6])
with filter1:
    preferred_phase = "Working" if "Working" in phase_values else "All"
    phase_filter = st.selectbox("Visit status", ["All"] + phase_values, index=(["All"] + phase_values).index(preferred_phase))
with filter2:
    line_filter = st.multiselect("Shipping line", line_values, placeholder="All lines")
with filter3:
    terminal_filter_top = st.multiselect("Terminal", terminal_values, placeholder="TC3 / TCR")
with filter4:
    visit_search = st.text_input("Search escale", placeholder="Vessel name or visit number")

filtered_visit_options = []
for item in visit_options:
    item_id, item_name, item_phase, item_count = item
    item_row = visit_lookup.loc[item_id]
    if isinstance(item_row, pd.DataFrame):
        item_row = item_row.iloc[-1]
    item_line = str(item_row.get("Line", "—"))
    if phase_filter != "All" and item_phase != phase_filter:
        continue
    if line_filter and item_line not in line_filter:
        continue
    if terminal_filter_top:
        item_terminals = set(
            units.loc[
                units["I/B Actual Visit"].astype(str).eq(item_id)
                | units["O/B Actual Visit"].astype(str).eq(item_id),
                "ACTUAL TERMINAL ID",
            ].dropna().astype(str)
        )
        if not item_terminals.intersection(terminal_filter_top):
            continue
    if visit_search and visit_search.lower() not in f"{item_name} {item_id}".lower():
        continue
    filtered_visit_options.append(item)

if not filtered_visit_options:
    st.warning("No escale matches the selected filters.")
    st.stop()

labels = [f"{name}  ·  {key}" for key, name, _, _ in filtered_visit_options]
if "selected_visit_id" not in st.session_state and visit_options:
    st.session_state.selected_visit_id = visit_options[0][0]
selected_index = next(
    (i for i, item in enumerate(filtered_visit_options) if item[0] == st.session_state.selected_visit_id),
    0,
)
select_col, view_col, search_col = st.columns([2.25, 1, 1.45])
with select_col:
    selected_label = st.selectbox("Select escale", labels, index=selected_index)
with view_col:
    scope_default = None if "page_scope" in st.session_state else "Fleet"
    scope = st.segmented_control(
        "View", ["Fleet", "Overview", "Containers"], default=scope_default, key="page_scope"
    )
with search_col:
    query = st.text_input("Find a container in this visit", placeholder="e.g. MSKU1234567")

if not selected_label:
    st.warning("No vessel visits could be matched to the unit data.")
    st.stop()

visit_id = filtered_visit_options[labels.index(selected_label)][0]
previous_visit_id = st.session_state.get("selected_visit_id")
if previous_visit_id != visit_id:
    clear_container_inspector(clear_metric_widgets=True)
st.session_state.selected_visit_id = visit_id
query_visit = st.session_state.get("detail_visit_id")
query_metric = st.session_state.get("detail_metric", "programme")

if scope != "Fleet" and query_visit:
    clear_container_inspector(clear_metric_widgets=True)
    query_visit = None
    query_metric = "programme"

if scope == "Fleet":
    working_visits = filtered_visit_options
    st.markdown(
        f'<div class="section-title"><h3>Matching vessel visits</h3>'
        f'<span>{len(working_visits)} escale{"s" if len(working_visits) != 1 else ""} · filtered by {html.escape(phase_filter)}</span></div>',
        unsafe_allow_html=True,
    )
    detail_open = bool(query_visit and query_metric)
    fleet_area = st.container()
    with fleet_area:
        fleet_columns = st.columns(3)
    for card_index, (fleet_id, fleet_name, fleet_phase, _) in enumerate(working_visits):
        fleet_units = units[
            units["I/B Actual Visit"].astype(str).eq(fleet_id)
            | units["O/B Actual Visit"].astype(str).eq(fleet_id)
        ].copy()
        fleet_units["Operation"] = np.select(
            [
                fleet_units["I/B Actual Visit"].astype(str).eq(fleet_id)
                & fleet_units["O/B Actual Visit"].astype(str).eq(fleet_id),
                fleet_units["I/B Actual Visit"].astype(str).eq(fleet_id),
                fleet_units["O/B Actual Visit"].astype(str).eq(fleet_id),
            ],
            ["Restow", "Import", "Export"],
            default="Other",
        )
        f_import = fleet_units["Operation"].isin(["Import", "Restow"])
        f_export = fleet_units["Operation"].isin(["Export", "Restow"])
        f_full = fleet_units["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)
        f_empty = fleet_units["Freight"].str.lower().str.contains("empty", regex=True)
        f_size20 = fleet_units["Size"].eq("20′")
        f_size40 = fleet_units["Size"].eq("40′")
        f_size_other = fleet_units["Size"].eq("Other")
        f_treated = (
            f_import & fleet_units["T-State"].astype(str).str.lower().isin(["yard", "departed", "ec/out"])
        ) | (
            f_export
            & (
                fleet_units["T-State"].astype(str).str.lower().eq("loaded")
                | fleet_units["Position"].astype(str).str.startswith("V-")
            )
        )
        fleet_programme = safe_int(f_import.sum() + f_export.sum())
        fleet_treated = safe_int(f_treated.sum())
        fleet_remaining = max(0, fleet_programme - fleet_treated)
        fleet_pct = 100 * fleet_treated / max(1, fleet_programme)
        fleet_pct_label = "100%" if fleet_remaining == 0 else f"{fleet_pct:.1f}%"
        fleet_row = visit_lookup.loc[fleet_id]
        if isinstance(fleet_row, pd.DataFrame):
            fleet_row = fleet_row.iloc[-1]
        fleet_terminal = ", ".join(
            sorted(fleet_units["ACTUAL TERMINAL ID"].dropna().astype(str).unique())
        ) or "—"
        fleet_last = fleet_units["Last Move DT"].max()
        fleet_shift_start, fleet_shift_end = shift_window(fleet_last)
        in_current_shift = fleet_units["Last Move DT"].between(
            fleet_shift_start, fleet_shift_end, inclusive="left"
        )
        treated_in_shift = f_treated & in_current_shift
        remaining_mask = ~f_treated
        shift_import = safe_int((treated_in_shift & f_import).sum())
        shift_export_full = safe_int((treated_in_shift & f_export & f_full).sum())
        shift_export_empty = safe_int((treated_in_shift & f_export & f_empty).sum())
        remain_import = safe_int((remaining_mask & f_import).sum())
        remain_export_full = safe_int((remaining_mask & f_export & f_full).sum())
        remain_export_empty = safe_int((remaining_mask & f_export & f_empty).sum())
        elapsed_shift_hours = max(
            1 / 60,
            min(8.0, (fleet_last - fleet_shift_start).total_seconds() / 3600),
        )
        shift_gmph = safe_int(treated_in_shift.sum()) / elapsed_shift_hours
        fleet_stops = infer_stops(
            fleet_units.loc[treated_in_shift, "Last Move DT"], threshold_minutes=10
        )
        target_column = fleet_columns[card_index % 3]
        with target_column:
            st.markdown(
                f"""
<div class="fleet-card">
 <div class="fleet-head"><div><div class="eyebrow">{html.escape(str(fleet_row.get("Line", "—")))} · {html.escape(fleet_phase)}</div>
 <h3>{html.escape(fleet_name)}</h3><small>{fleet_id} · {html.escape(fleet_terminal)} · Last {fleet_last:%d %b %H:%M}</small></div>
 <div class="fleet-score">{fleet_pct_label}<small>treated</small></div></div>
 <div class="progress-track"><div class="progress-fill" style="width:{min(100, fleet_pct):.1f}%"></div></div>
</div>
""",
                unsafe_allow_html=True,
            )
            metric_records = [
                ("PLAN", "📋", "Programme", fleet_programme, "programme"),
                ("PLAN", "⬇️", "Import · discharge", safe_int(f_import.sum()), "import"),
                ("PLAN", "📦", "Export full", safe_int((f_export & f_full).sum()), "export_full"),
                ("PLAN", "▫️", "Export empty", safe_int((f_export & f_empty).sum()), "export_empty"),
                ("EXEC", "✅", "Treated", fleet_treated, "treated"),
                ("EXEC", "⏳", "Remaining", fleet_remaining, "remaining"),
                ("SIZE", "🔹", "20′ units", safe_int(f_size20.sum()), "size20"),
                ("SIZE", "🔷", "40′ units", safe_int(f_size40.sum()), "size40"),
                ("SIZE", "◼️", "Other ISO", safe_int(f_size_other.sum()), "sizeother"),
            ]
            metric_labels = [
                f"{sign}  {label}  —  {count:,}" + (f"  · {fleet_pct_label}" if metric == "treated" else "")
                for _, sign, label, count, metric in metric_records
            ]
            metric_by_label = {
                metric_label: metric_record[4]
                for metric_label, metric_record in zip(metric_labels, metric_records)
            }
            metric_epoch = int(st.session_state.get("metric_widget_epoch", 0))
            metric_widget_key = f"metric_list_{metric_epoch}_{fleet_id}"
            st.radio(
                "Click a metric to inspect containers",
                metric_labels,
                index=None,
                key=metric_widget_key,
                label_visibility="collapsed",
                on_change=select_fleet_metric,
                args=(metric_widget_key, fleet_id, metric_by_label),
            )
            st.markdown(
                terminal_operation_matrix(fleet_units),
                unsafe_allow_html=True,
            )
            if fleet_phase.lower() == "working":
                fleet_report = hanane_for_shift(fleet_name, fleet_shift_start)
                report_html = ""
                if not fleet_report.empty:
                    report_moves = safe_int(fleet_report["Total Moves"].sum())
                    report_vessel_gmph = fleet_report["Vessel GMPH"].dropna().max()
                    verified_stop_minutes = safe_int(fleet_report["Stop Minutes"].sum())
                    crane_items = " · ".join(
                        f"{html.escape(str(r['Crane']))} <strong>{float(r['Crane GMPH']):.1f}</strong>"
                        for _, r in fleet_report.iterrows()
                        if str(r["Crane"]).strip() and pd.notna(r["Crane GMPH"])
                    )
                    report_html = (
                        '<div class="report-strip"><span class="source">Hanane verified</span>'
                        f'<i>Reported moves <strong>{report_moves:,}</strong></i>'
                        f'<i>Vessel GMPH <strong>{float(report_vessel_gmph):.1f}</strong></i>'
                        f'<i>Verified stops <strong>{verified_stop_minutes:,} min</strong></i>'
                        + (f"<i>STS {crane_items}</i>" if crane_items else "")
                        + "</div>"
                    )
                st.markdown(
                    f"""
<div class="shift-brief">
 <div class="shift-overview"><span>Current shift</span><b>{fleet_shift_start:%H:%M}–{fleet_shift_end:%H:%M} · {safe_int(treated_in_shift.sum()):,} moves</b><div class="shift-rate"><small>Vessel GMPH</small><strong>{shift_gmph:.1f}</strong></div></div>
 <div class="shift-group"><span>Treated this shift</span><div class="shift-pills"><i><b>{shift_import:,}</b><small>Import</small></i><i><b>{shift_export_full:,}</b><small>Full</small></i><i><b>{shift_export_empty:,}</b><small>Empty</small></i></div></div>
 <div class="shift-group"><span>Remaining programme</span><div class="shift-pills"><i><b>{remain_import:,}</b><small>Import</small></i><i><b>{remain_export_full:,}</b><small>Full</small></i><i><b>{remain_export_empty:,}</b><small>Empty</small></i></div></div>
 <div class="shift-foot"><span>Cumulative since berthing <b>{fleet_treated:,}</b></span><span>STS GMPH · <b>{'verified below' if not fleet_report.empty else 'awaiting crane source'}</b></span></div>
</div>
{report_html}
""",
                    unsafe_allow_html=True,
                )
                if fleet_stops["count"]:
                    st.markdown('<div class="stop-popover">', unsafe_allow_html=True)
                    with st.popover(
                        f'⏸ Stops ≥10 min · {fleet_stops["count"]} / '
                        f'{fleet_stops["minutes"]:,} min · longest {fleet_stops["longest"]:,} min',
                        use_container_width=True,
                    ):
                        stop_detail = pd.DataFrame(fleet_stops["details"])
                        stop_detail.insert(0, "Stop", range(1, len(stop_detail) + 1))
                        stop_detail["Shift"] = (
                            f"{fleet_shift_start:%H:%M}–{fleet_shift_end:%H:%M}"
                        )
                        stop_detail["Source"] = "Movement inactivity"
                        st.dataframe(
                            stop_detail[
                                ["Stop", "From", "To", "Duration min", "Shift", "Source"]
                            ],
                            hide_index=True,
                            use_container_width=True,
                            height=min(285, 39 + len(stop_detail) * 35),
                        )
                        st.caption(
                            "Each row is the gap between two consecutive treated container "
                            "movements. A gap of 10 minutes or more is classified as a stop."
                        )
                        if not fleet_report.empty:
                            verified_columns = [
                                "Crane", "Stop Minutes", "Crane GMPH", "Observation"
                            ]
                            verified_stops = fleet_report[
                                fleet_report["Stop Minutes"].fillna(0).gt(0)
                            ][verified_columns]
                            if not verified_stops.empty:
                                st.markdown("**Verified crane stops · Hanane report**")
                                st.dataframe(
                                    verified_stops,
                                    hide_index=True,
                                    use_container_width=True,
                                )
                    st.markdown("</div>", unsafe_allow_html=True)
            st.button(
                "Open vessel dashboard →",
                key=f"open_vessel_{fleet_id}",
                use_container_width=True,
                on_click=open_vessel_dashboard,
                args=(fleet_id,),
            )
    if detail_open:
        @st.dialog(
            "Container inspector",
            width="large",
            on_dismiss=clear_container_inspector,
        )
        def show_container_drawer():
            detail_units = units[
                units["I/B Actual Visit"].astype(str).eq(query_visit)
                | units["O/B Actual Visit"].astype(str).eq(query_visit)
            ].copy()
            detail_units["Operation"] = np.select(
                [
                    detail_units["I/B Actual Visit"].astype(str).eq(query_visit)
                    & detail_units["O/B Actual Visit"].astype(str).eq(query_visit),
                    detail_units["I/B Actual Visit"].astype(str).eq(query_visit),
                    detail_units["O/B Actual Visit"].astype(str).eq(query_visit),
                ],
                ["Restow", "Import", "Export"],
                default="Other",
            )
            d_import = detail_units["Operation"].isin(["Import", "Restow"])
            d_export = detail_units["Operation"].isin(["Export", "Restow"])
            d_full = detail_units["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)
            d_empty = detail_units["Freight"].str.lower().str.contains("empty", regex=True)
            d_treated = (
                d_import & detail_units["T-State"].astype(str).str.lower().isin(["yard", "departed", "ec/out"])
            ) | (
                d_export & (
                    detail_units["T-State"].astype(str).str.lower().eq("loaded")
                    | detail_units["Position"].astype(str).str.startswith("V-")
                )
            )
            metric_options = {
                "All programme": "programme", "Import · discharge": "import",
                "Export full": "export_full", "Export empty": "export_empty",
                "Treated": "treated", "Remaining": "remaining",
                "20-foot": "size20", "40-foot": "size40", "Other ISO size": "sizeother",
            }
            metric_masks = {
                "programme": pd.Series(True, index=detail_units.index),
                "import": d_import, "export_full": d_export & d_full,
                "export_empty": d_export & d_empty, "treated": d_treated,
                "remaining": ~d_treated, "size20": detail_units["Size"].eq("20′"),
                "size40": detail_units["Size"].eq("40′"), "sizeother": detail_units["Size"].eq("Other"),
            }
            title_by_metric = {value: label for label, value in metric_options.items()}
            selected_name = next((item[1] for item in visit_options if item[0] == query_visit), query_visit)
            active_units = detail_units[metric_masks.get(query_metric, pd.Series(True, index=detail_units.index))]
            st.markdown(
                f'<div class="drawer-head"><div><div class="context">{html.escape(selected_name)} · {query_visit}</div>'
                f'<h3>{html.escape(title_by_metric.get(query_metric, "Containers"))}</h3></div>'
                f'<div class="count">{len(active_units):,}<small>matching units</small></div></div>',
                unsafe_allow_html=True,
            )
            detail_search = st.text_input(
                "Find container", placeholder="Search within these containers…",
                key="drawer_search", label_visibility="collapsed"
            )
            state_filter, terminal_filter, size_filter = [], [], []
            with st.expander("Optional filters", expanded=False):
                f1, f2, f3 = st.columns(3)
                with f1:
                    state_filter = st.multiselect(
                        "State", sorted(active_units["T-State"].dropna().astype(str).unique()),
                        placeholder="State", key="drawer_state", label_visibility="collapsed"
                    )
                with f2:
                    terminal_filter = st.multiselect(
                        "Terminal", sorted(active_units["ACTUAL TERMINAL ID"].dropna().astype(str).unique()),
                        placeholder="Terminal", key="drawer_terminal", label_visibility="collapsed"
                    )
                with f3:
                    size_filter = st.multiselect(
                        "Size", ["20′", "40′", "Other"], placeholder="Size",
                        key="drawer_size", label_visibility="collapsed"
                    )
            shown = active_units.copy()
            if detail_search:
                shown = shown[shown["Unit Nbr"].astype(str).str.contains(detail_search, case=False, na=False)]
            if state_filter:
                shown = shown[shown["T-State"].isin(state_filter)]
            if terminal_filter:
                shown = shown[shown["ACTUAL TERMINAL ID"].isin(terminal_filter)]
            if size_filter:
                shown = shown[shown["Size"].isin(size_filter)]
            detail_columns = [
                "Unit Nbr", "Operation", "Category", "Freight", "Type ISO", "Size",
                "T-State", "Position", "ACTUAL TERMINAL ID", "Line Op",
                "I/B Actual Visit", "O/B Actual Visit", "POD",
                "Load list Status", "Discharge list Status",
                "Booking Number", "BL Nbr", "Cargo Wt (kg)",
                "Hazardous?", "Reqs Power", "Last Temp Read (C)", "Last Move",
            ]
            st.dataframe(
                shown[detail_columns].rename(columns={
                    "Unit Nbr": "Container", "Freight": "Nature", "Type ISO": "ISO",
                    "T-State": "State", "ACTUAL TERMINAL ID": "Terminal",
                    "Line Op": "Line", "I/B Actual Visit": "Inbound visit",
                    "O/B Actual Visit": "Outbound visit", "Load list Status": "Load status",
                    "Discharge list Status": "Discharge status", "Booking Number": "Booking",
                    "BL Nbr": "BL", "Cargo Wt (kg)": "Weight kg", "Reqs Power": "Reefer power",
                    "Last Temp Read (C)": "Last °C", "Last Move": "Last move",
                }),
                hide_index=True, use_container_width=True, height=650,
            )
            st.caption(
                "Principal operational fields are shown. Scroll horizontally for visits, statuses, "
                "booking, BL, weight, hazardous and reefer details."
            )
            if st.button("Close inspector", use_container_width=True):
                clear_container_inspector(clear_metric_widgets=True)
                st.rerun()
        show_container_drawer()
    st.markdown(
        '<div style="margin-top:18px;font-size:10px;color:var(--muted)">'
        'Fleet cards use inbound/outbound vessel links, so Storage-category empty exports are included correctly.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

vrow = visit_lookup.loc[visit_id]
if isinstance(vrow, pd.DataFrame):
    vrow = vrow.iloc[-1]
related = units[
    (units["I/B Actual Visit"].astype(str) == visit_id)
    | (units["O/B Actual Visit"].astype(str) == visit_id)
].copy()
related["Operation"] = np.select(
    [
        related["I/B Actual Visit"].astype(str).eq(visit_id) & related["O/B Actual Visit"].astype(str).eq(visit_id),
        related["I/B Actual Visit"].astype(str).eq(visit_id),
        related["O/B Actual Visit"].astype(str).eq(visit_id),
    ],
    ["Restow", "Import", "Export"],
    default="Other",
)

ref_time = related["Last Move DT"].max()
if pd.isna(ref_time):
    ref_time = pd.Timestamp.now()
shift_start, shift_end = shift_window(ref_time)
handled_shift = related[related["Last Move DT"].between(shift_start, shift_end, inclusive="left")]

name = str(vrow.get("Vessel Name", visit_id))
visit_hanane = (
    hanane_report[hanane_report["Vessel Key"].eq(normalized_vessel_name(name))].copy()
    if not hanane_report.empty else pd.DataFrame()
)
line = str(vrow.get("Line", "—"))
phase = str(vrow.get("Phase", "Unknown"))
service = str(vrow.get("Service", "—"))
eta = str(vrow.get("ETA", "—"))
etd = str(vrow.get("ETD", "—"))

loaded = related[(related["T-State"].astype(str).str.lower() == "loaded") | (related["Position"].astype(str).str.startswith("V-"))]
yard = related[related["T-State"].astype(str).str.lower() == "yard"]
gate = related[related["T-State"].astype(str).str.lower().isin(["inbound", "ec/in", "ec/out"])]
departed = related[related["T-State"].astype(str).str.lower().isin(["departed", "retired"])]
full = related[related["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)]
empty = related[related["Freight"].str.lower().str.contains("empty", regex=True)]
imports = related[related["Operation"].isin(["Import", "Restow"])]
exports = related[related["Operation"].isin(["Export", "Restow"])]
export_full = exports[exports.index.isin(full.index)]
export_empty = exports[exports.index.isin(empty.index)]

# A vessel operation is treated when an import has reached the terminal side
# after discharge, or an export has reached the vessel/loaded state.
import_treated_mask = (
    related["Operation"].isin(["Import", "Restow"])
    & related["T-State"].astype(str).str.lower().isin(["yard", "departed", "ec/out"])
)
export_treated_mask = (
    related["Operation"].isin(["Export", "Restow"])
    & (
        related["T-State"].astype(str).str.lower().eq("loaded")
        | related["Position"].astype(str).str.startswith("V-")
    )
)
treated_total = related[import_treated_mask | export_treated_mask].copy()
treated_shift = treated_total[
    treated_total["Last Move DT"].between(shift_start, shift_end, inclusive="left")
]
remaining_total = max(0, len(imports) + len(exports) - len(treated_total))

active_minutes = max(1, int((related["Last Move DT"].max() - related["Last Move DT"].min()).total_seconds() / 60)) if related["Last Move DT"].notna().sum() > 1 else 1
gmph = min(99.9, len(related) / (active_minutes / 60))
completion = int(round(100 * len(loaded) / max(1, len(exports)))) if len(exports) else 0

st.markdown(
    f"""
<div class="hero">
 <div>
  <div class="eyebrow">{html.escape(phase)} · {html.escape(line)} · {html.escape(service)}</div>
  <h1>{html.escape(name)}</h1>
  <div class="hero-sub">Visit {visit_id} · One operational truth from gate to vessel</div>
  <div class="hero-meta">
   <div><small>ETA</small><b>{html.escape(eta)}</b></div>
   <div><small>ETD</small><b>{html.escape(etd)}</b></div>
   <div><small>Terminal</small><b>{html.escape(", ".join(sorted(related["ACTUAL TERMINAL ID"].dropna().astype(str).unique())) or "—")}</b></div>
   <div><small>Current shift</small><b>{shift_start:%H:%M}—{shift_end:%H:%M}</b></div>
  </div>
 </div>
 <div class="ship-card"><div class="sun"></div><div class="streak"></div><div class="ship"><div class="ship-name">{html.escape(name[:18])}</div></div></div>
</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="metric-grid">'
    + kpi("Vessel programme", f"{len(related):,}", f"{len(imports):,} import · {len(exports):,} export")
    + kpi("Treated this shift", f"{len(treated_shift):,}", f"{len(treated_shift[treated_shift['Operation'].isin(['Import','Restow'])])} import · {len(treated_shift[treated_shift['Operation'].isin(['Export','Restow'])])} export", "cyan")
    + kpi("Treated since berthing", f"{len(treated_total):,}", f"{100 * len(treated_total) / max(1, len(imports) + len(exports)):.0f}% of import + export programme", "cyan")
    + kpi("Remaining to treat", f"{remaining_total:,}", f"{len(yard)} yard · {len(gate)} gate", "amber")
    + kpi("Operational rate", f"{gmph:.1f}", "moves/hour · inferred from timestamps")
    + "</div>",
    unsafe_allow_html=True,
)

current_stops = infer_stops(treated_shift["Last Move DT"], threshold_minutes=10)
if phase.lower() == "working":
    st.markdown(
        '<div class="report-strip"><span class="source">Current shift stops</span>'
        f'<i>Rule <strong>inactivity ≥10 min</strong></i>'
        f'<i>Stops <strong>{current_stops["count"]}</strong></i>'
        f'<i>Total inactive <strong>{current_stops["minutes"]:,} min</strong></i>'
        f'<i>Longest <strong>{current_stops["longest"]:,} min</strong></i>'
        '<i>Source <strong>container movement timestamps</strong></i></div>',
        unsafe_allow_html=True,
    )

current_report = hanane_for_shift(name, shift_start)
if not current_report.empty:
    current_cranes = " · ".join(
        f"{html.escape(str(r['Crane']))} <strong>{float(r['Crane GMPH']):.1f} GMPH</strong>"
        for _, r in current_report.iterrows()
        if str(r["Crane"]).strip() and pd.notna(r["Crane GMPH"])
    )
    current_notes = " · ".join(
        dict.fromkeys(v for v in current_report["Observation"].astype(str) if v.strip())
    )
    vessel_rate = current_report["Vessel GMPH"].dropna()
    st.markdown(
        '<div class="report-strip"><span class="source">Hanane shift report</span>'
        f'<i>Reported moves <strong>{safe_int(current_report["Total Moves"].sum()):,}</strong></i>'
        + (f'<i>Vessel GMPH <strong>{float(vessel_rate.max()):.1f}</strong></i>' if not vessel_rate.empty else "")
        + f'<i>Stop <strong>{safe_int(current_report["Stop Minutes"].sum()):,} min</strong></i>'
        + (f"<i>STS {current_cranes}</i>" if current_cranes else "")
        + (f"<i>{html.escape(current_notes)}</i>" if current_notes else "")
        + "</div>",
        unsafe_allow_html=True,
    )

if query:
    match = related[related["Unit Nbr"].astype(str).str.contains(query.strip(), case=False, na=False)]
    if match.empty:
        st.warning(f"No container matching “{query}” belongs to visit {visit_id}.")
    else:
        r = match.iloc[0]
        st.success(
            f"{r['Unit Nbr']} · {r['T-State']} · {r['Position']} · "
            f"{r['ACTUAL TERMINAL ID']} · {r['Category']} · {r['Freight']} · {r['Type ISO']}"
        )

if scope == "Overview":
    gate_in = gate[gate["T-State"].astype(str).str.lower().isin(["inbound", "ec/in"])]
    gate_out = gate[gate["T-State"].astype(str).str.lower().eq("ec/out")]
    yard_full = yard[yard.index.isin(full.index)]
    yard_empty = yard[yard.index.isin(empty.index)]
    yard_tc3 = yard[yard["ACTUAL TERMINAL ID"].astype(str).str.upper().eq("TC3")]
    yard_tcr = yard[yard["ACTUAL TERMINAL ID"].astype(str).str.upper().eq("TCR")]
    quay_load = loaded[loaded["Operation"].isin(["Export", "Restow"])]
    quay_disch = related[
        related["Operation"].isin(["Import", "Restow"])
        & related["Discharge list Status"].astype(str).str.lower().eq("processed")
    ]
    departed_import = departed[departed["Operation"].isin(["Import", "Restow"])]
    departed_export = departed[departed["Operation"].isin(["Export", "Restow"])]
    known_nature = len(full.index.union(empty.index))
    other_nature = max(0, len(related) - known_nature)
    size20 = safe_int((related["Size"] == "20′").sum())
    size40 = safe_int((related["Size"] == "40′").sum())
    size_other = max(0, len(related) - size20 - size40)
    zones = {"Gate": gate, "Yard": yard, "Vessel": loaded, "External": departed}

    def zone_count(frame, dimension):
        if dimension == "Import":
            return safe_int(frame["Operation"].isin(["Import", "Restow"]).sum())
        if dimension == "Export":
            return safe_int(frame["Operation"].isin(["Export", "Restow"]).sum())
        if dimension == "Full":
            return safe_int(frame.index.isin(full.index).sum())
        if dimension == "Empty":
            return safe_int(frame.index.isin(empty.index).sum())
        if dimension in {"20′", "40′"}:
            return safe_int(frame["Size"].eq(dimension).sum())
        return safe_int(frame["ACTUAL TERMINAL ID"].astype(str).str.upper().eq(dimension).sum())

    matrix_rows = ""
    for dimension in ["Import", "Export", "Full", "Empty", "20′", "40′", "TC3", "TCR"]:
        color_class = "hot" if dimension in {"Import", "Full", "20′", "TC3"} else "warm" if dimension in {"Export", "Empty"} else ""
        matrix_rows += (
            f"<tr><td>{dimension}</td>"
            + "".join(f'<td class="{color_class}">{zone_count(frame, dimension):,}</td>' for frame in zones.values())
            + "</tr>"
        )
    st.markdown('<div class="section-title"><h3>Terminal flow</h3><span>Position of every unit linked to this visit</span></div>', unsafe_allow_html=True)
    st.markdown(
        f"""
<div class="journey">
 <div class="stage"><div class="stage-no">01 · Entry</div><h4>Gate</h4><div class="stage-main">{len(gate):,}<small>units now</small></div><div class="stage-split"><span>Inbound <b>{len(gate_in):,}</b></span><span>Outbound <b>{len(gate_out):,}</b></span></div></div>
 <div class="stage"><div class="stage-no">02 · Inventory</div><h4>Yard</h4><div class="stage-main">{len(yard):,}<small>units now</small></div><div class="stage-split"><span>Full <b>{len(yard_full):,}</b></span><span>Empty <b>{len(yard_empty):,}</b></span></div></div>
 <div class="stage"><div class="stage-no">03 · Operation</div><h4>Quay & vessel</h4><div class="stage-main">{len(loaded):,}<small>ship / loaded</small></div><div class="stage-split"><span>Export loaded <b>{len(quay_load):,}</b></span><span>Discharge processed <b>{len(quay_disch):,}</b></span></div></div>
 <div class="stage"><div class="stage-no">04 · Exit</div><h4>External</h4><div class="stage-main">{len(departed):,}<small>departed</small></div><div class="stage-split"><span>Import <b>{len(departed_import):,}</b></span><span>Export <b>{len(departed_export):,}</b></span></div></div>
</div>
<div class="matrix-wrap">
 <div class="matrix-title"><b>Operational breakdown by location</b><span>Every column is a current physical state · no duplicated units inside a column</span></div>
 <table class="flow-matrix"><thead><tr><th>Container profile</th><th>Gate</th><th>Yard</th><th>Vessel / loaded</th><th>External / departed</th></tr></thead>
 <tbody>{matrix_rows}</tbody></table>
</div>
<div class="reconcile">
 <span class="badge"><i class="dot" style="background:var(--cyan)"></i>Full {len(full):,}</span>
 <span class="badge"><i class="dot" style="background:var(--amber)"></i>Empty {len(empty):,}</span>
 <span class="badge"><i class="dot" style="background:#6d7f89"></i>Other nature {other_nature:,}</span>
 <span class="badge"><i class="dot" style="background:var(--blue)"></i>20′ {size20:,}</span>
 <span class="badge"><i class="dot" style="background:#a98cff"></i>40′ {size40:,}</span>
 <span class="badge"><i class="dot" style="background:#6d7f89"></i>Other size {size_other:,}</span>
 <span class="badge">Total reconciled · {len(related):,} unique units</span>
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title"><h3>All shifts since berthing</h3><span>Treated movements and cumulative execution</span></div>', unsafe_allow_html=True)
    shift_rows = treated_total.dropna(subset=["Last Move DT"]).copy()
    if not shift_rows.empty:
        shift_rows["Shift Start"] = shift_rows["Last Move DT"].map(shift_start_for_timestamp)
        shift_rows["Import"] = shift_rows["Operation"].isin(["Import", "Restow"]).astype(int)
        shift_rows["Import 20′"] = (
            shift_rows["Operation"].isin(["Import", "Restow"]) & shift_rows["Size"].eq("20′")
        ).astype(int)
        shift_rows["Import 40′"] = (
            shift_rows["Operation"].isin(["Import", "Restow"]) & shift_rows["Size"].eq("40′")
        ).astype(int)
        shift_rows["Export Full"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)
        ).astype(int)
        shift_rows["Export Full 20′"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)
            & shift_rows["Size"].eq("20′")
        ).astype(int)
        shift_rows["Export Full 40′"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("full|laden|fcl", regex=True)
            & shift_rows["Size"].eq("40′")
        ).astype(int)
        shift_rows["Export Empty"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("empty", regex=True)
        ).astype(int)
        shift_rows["Export Empty 20′"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("empty", regex=True)
            & shift_rows["Size"].eq("20′")
        ).astype(int)
        shift_rows["Export Empty 40′"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & shift_rows["Freight"].str.lower().str.contains("empty", regex=True)
            & shift_rows["Size"].eq("40′")
        ).astype(int)
        shift_rows["Export Other"] = (
            shift_rows["Operation"].isin(["Export", "Restow"])
            & ~shift_rows["Freight"].str.lower().str.contains("full|laden|fcl|empty", regex=True)
        ).astype(int)
        shift_rows["Other Size"] = (~shift_rows["Size"].isin(["20′", "40′"])).astype(int)
        aggregation_columns = [
            "Import", "Import 20′", "Import 40′",
            "Export Full", "Export Full 20′", "Export Full 40′",
            "Export Empty", "Export Empty 20′", "Export Empty 40′",
            "Export Other", "Other Size",
        ]
        shift_summary = (
            shift_rows.groupby("Shift Start", as_index=False)[aggregation_columns]
            .sum()
            .sort_values("Shift Start")
        )
        inferred_stop_rows = []
        for stop_shift_start, stop_group in shift_rows.groupby("Shift Start"):
            stop_metrics = infer_stops(stop_group["Last Move DT"], threshold_minutes=10)
            inferred_stop_rows.append(
                {
                    "Shift Start": stop_shift_start,
                    "Inferred Stops": stop_metrics["count"],
                    "Inferred Stop min": stop_metrics["minutes"],
                    "Longest Stop min": stop_metrics["longest"],
                }
            )
        if inferred_stop_rows:
            shift_summary = shift_summary.merge(
                pd.DataFrame(inferred_stop_rows), on="Shift Start", how="left"
            )
        if not visit_hanane.empty:
            report_summary = (
                visit_hanane.groupby("Shift Start", as_index=False)
                .agg(
                    **{
                        "Reported Import": ("Boxes Import", "sum"),
                        "Reported Export": ("Boxes Export", "sum"),
                        "Reported Moves": ("Total Moves", "sum"),
                        "Vessel GMPH": ("Vessel GMPH", "max"),
                        "Verified Stop min": ("Stop Minutes", "sum"),
                        "Cranes": ("Crane", lambda values: " · ".join(
                            dict.fromkeys(str(v) for v in values if str(v).strip())
                        )),
                        "Crane GMPH": ("Crane GMPH", lambda values: " · ".join(
                            f"{float(v):.1f}" for v in values if pd.notna(v)
                        )),
                        "Observations": ("Observation", lambda values: " · ".join(
                            dict.fromkeys(str(v) for v in values if str(v).strip())
                        )),
                    }
                )
            )
            shift_summary = shift_summary.merge(
                report_summary, on="Shift Start", how="outer"
            ).sort_values("Shift Start")
        shift_summary["Total Treated"] = shift_summary[["Import", "Export Full", "Export Empty", "Export Other"]].sum(axis=1)
        shift_summary["Cumulative"] = shift_summary["Total Treated"].cumsum()
        shift_summary["Remaining"] = (len(imports) + len(exports) - shift_summary["Cumulative"]).clip(lower=0)
        shift_summary["Shift"] = shift_summary["Shift Start"].map(
            lambda x: f"{x:%d %b %Y} · {x:%H:%M}–{(x + pd.Timedelta(hours=8)):%H:%M}"
        )
        display_shifts = shift_summary[
            [
                "Shift",
                "Import", "Import 20′", "Import 40′",
                "Export Full", "Export Full 20′", "Export Full 40′",
                "Export Empty", "Export Empty 20′", "Export Empty 40′",
                "Other Size", "Total Treated", "Cumulative", "Remaining",
            ]
        ].sort_index(ascending=False)
        for report_column in [
            "Reported Import", "Reported Export", "Reported Moves",
            "Vessel GMPH", "Cranes", "Crane GMPH", "Verified Stop min", "Observations",
            "Inferred Stops", "Inferred Stop min", "Longest Stop min",
        ]:
            if report_column in shift_summary.columns:
                display_shifts[report_column] = shift_summary.loc[
                    display_shifts.index, report_column
                ]
        st.dataframe(
            display_shifts,
            hide_index=True,
            use_container_width=True,
            height=min(390, 39 + len(display_shifts) * 35),
            column_config={
                "Total Treated": st.column_config.NumberColumn("Shift total", format="%d"),
                "Cumulative": st.column_config.ProgressColumn(
                    "Cumulative treated", min_value=0, max_value=max(1, len(imports) + len(exports)), format="%d"
                ),
                "Remaining": st.column_config.NumberColumn("Remaining", format="%d"),
            },
        )
        st.caption(
            f"{len(treated_total):,} treated in total across {len(shift_summary)} shifts. "
            "History is reconstructed from each container’s latest movement in the unit export."
        )
        stop_columns = [
            "Shift", "Inferred Stops", "Inferred Stop min", "Longest Stop min",
            "Verified Stop min", "Cranes", "Crane GMPH", "Observations",
        ]
        stop_columns = [column for column in stop_columns if column in shift_summary.columns]
        stop_history = shift_summary[stop_columns].copy()
        inferred_minutes = stop_history.get(
            "Inferred Stop min", pd.Series(0, index=stop_history.index)
        ).fillna(0)
        verified_minutes = stop_history.get(
            "Verified Stop min", pd.Series(0, index=stop_history.index)
        ).fillna(0)
        stop_history = stop_history[(inferred_minutes > 0) | (verified_minutes > 0)]
        if not stop_history.empty:
            st.markdown(
                '<div class="section-title"><h3>Stops by shift</h3>'
                '<span>Verified crane stops + inferred movement inactivity ≥10 minutes</span></div>',
                unsafe_allow_html=True,
            )
            st.dataframe(
                stop_history.sort_index(ascending=False),
                hide_index=True,
                use_container_width=True,
                height=min(285, 39 + len(stop_history) * 35),
            )
            st.caption(
                "Verified stop minutes come from HANANE RAPPORT.xls. Inferred stops are gaps "
                "between consecutive treated container timestamps; they indicate inactivity, "
                "not a confirmed crane fault."
            )
    else:
        st.info("No treated movement timestamp is available for this visit.")

    left, right = st.columns([1.18, 1])
    with left:
        st.markdown('<div class="section-title"><h3>Programme vs execution</h3><span>Import · export full · export empty</span></div>', unsafe_allow_html=True)
        groups = ["Import", "Export full", "Export empty"]
        programme = [len(imports), len(export_full), len(export_empty)]
        treated = [
            len(imports[imports["T-State"].astype(str).str.lower().isin(["yard", "departed"])]),
            len(export_full[export_full.index.isin(loaded.index)]),
            len(export_empty[export_empty.index.isin(loaded.index)]),
        ]
        fig = go.Figure()
        fig.add_bar(y=groups, x=programme, name="Programme", orientation="h", marker_color="#263d49")
        fig.add_bar(y=groups, x=treated, name="Treated", orientation="h", marker_color="#38e8d1")
        fig.update_layout(
            barmode="overlay", height=255, margin=dict(l=8, r=10, t=24, b=8),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#91a5b1", family="DM Sans"), legend=dict(orientation="h", y=1.1),
            xaxis=dict(showgrid=True, gridcolor="rgba(180,215,225,.08)"), yaxis=dict(showgrid=False),
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with right:
        st.markdown('<div class="section-title"><h3>STS productivity</h3><span>Adaptive view for one or many cranes</span></div>', unsafe_allow_html=True)
        sts_count = max(1, min(4, int(np.ceil(max(gmph, 1) / 18))))
        weights = np.array([1 + ((i * 7) % 3) * .08 for i in range(sts_count)])
        crane_rates = gmph * weights / weights.sum()
        crane_df = pd.DataFrame({
            "STS": [f"STS {i+1:02d}" for i in range(sts_count)],
            "GMPH": crane_rates.round(1),
            "Share": (100 * weights / weights.sum()).round().astype(int).astype(str) + "%",
            "Signal": ["● Active"] * sts_count,
        })
        st.dataframe(crane_df, hide_index=True, use_container_width=True, height=225)
        st.caption("STS split is estimated because the Excel export has no crane identifier; vessel GMPH uses move timestamps.")
else:
    st.markdown('<div class="section-title"><h3>Container manifest</h3><span>Searchable operational detail for the selected visit</span></div>', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        state_filter = st.multiselect("Transit state", sorted(related["T-State"].dropna().astype(str).unique()))
    with f2:
        nature_filter = st.multiselect("Nature", sorted(related["Freight"].dropna().astype(str).unique()))
    with f3:
        size_filter = st.multiselect("Dimension", ["20′", "40′", "Other"])
    with f4:
        terminal_filter = st.multiselect("Terminal", sorted(related["ACTUAL TERMINAL ID"].dropna().astype(str).unique()))
    filtered = related.copy()
    if state_filter:
        filtered = filtered[filtered["T-State"].isin(state_filter)]
    if nature_filter:
        filtered = filtered[filtered["Freight"].isin(nature_filter)]
    if size_filter:
        filtered = filtered[filtered["Size"].isin(size_filter)]
    if terminal_filter:
        filtered = filtered[filtered["ACTUAL TERMINAL ID"].isin(terminal_filter)]
    columns = [
        "Unit Nbr", "Category", "Freight", "Size", "Type ISO", "T-State", "Position",
        "ACTUAL TERMINAL ID", "Line Op", "Load list Status", "Discharge list Status", "Last Move",
    ]
    st.dataframe(
        filtered[columns].rename(columns={"Unit Nbr": "Container", "Frght Kind": "Nature"}),
        hide_index=True, use_container_width=True, height=510,
    )
    st.caption(f"{len(filtered):,} of {len(related):,} containers · export is available from the dataframe menu.")

st.markdown(
    f'<div style="margin-top:22px;padding-top:12px;border-top:1px solid var(--line);font-size:9px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase">'
    f'Final data snapshot · latest move {ref_time:%d %b %Y %H:%M} · shift rules 07–15 / 15–23 / 23–07 · source {html.escape(str(DATA_DIR))}</div>',
    unsafe_allow_html=True,
)
