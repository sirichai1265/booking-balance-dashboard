#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
build_dashboard.py
==================
สร้าง Dashboard (ไฟล์ HTML เดียว เปิดในเบราว์เซอร์ได้เลย) สำหรับ "ค้นหาบุ๊คที่ค้างรับ"
จากไฟล์ Booking Balance Summary (ชีต Data)

การใช้งาน:
    python build_dashboard.py "9-12-PD - Copy.xls"                       # จากไฟล์ต้นฉบับโดยตรง (แนะนำ)
    python build_dashboard.py "output/Booking Balance Summary - Copy.xlsx"  # หรือจากไฟล์สรุปที่เปิด/Save ด้วย Excel แล้ว
    python build_dashboard.py            # ไม่ใส่ path -> หาไฟล์ให้เอง

ต้องมี lib:  pip install xlrd openpyxl
- ถ้าใส่ไฟล์ .xls ต้นฉบับ: คำนวณยอดคงเหลือเอง (เรียกใช้ build_booking_balance_report.py) ไม่ต้องพึ่ง Excel เลย
- ถ้าใส่ไฟล์ .xlsx (ชีต Data): อ่านค่าที่ Excel คำนวณไว้แล้ว (data_only) — ต้องเปิดไฟล์ใน Excel แล้ว Save
  หนึ่งครั้งก่อน ไม่งั้นค่าจะว่าง
ผลลัพธ์:
  - Booking Balance Dashboard.html  (ในโฟลเดอร์เดียวกับไฟล์ต้นทาง)
  - Daily Booking <YYYY-MM-DD>.xlsx (รูปแบบ "daily booking" 1 ชีต — บันทึกลงโฟลเดอร์โปรเจกต์นี้เสมอ
    ไม่ว่าไฟล์ต้นทางจะอยู่ที่ไหน; แทนปุ่มดาวน์โหลด Excel บนเว็บที่เอาออกแล้ว)
"""

import glob
import html
import json
import os
import sys

import openpyxl

TYPE_COLS = ["GP22", "GP42", "GP45", "RE22", "RE45", "UT22", "UT42", "PC22", "PC42"]


def num(v):
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 2)
    except (TypeError, ValueError):
        return 0


def fmt_dt(v):
    s = str(v or "").strip()
    if s.isdigit() and len(s) == 12:
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]} {s[8:10]}:{s[10:12]}"
    if s.isdigit() and len(s) == 8:
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    return s


def load_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in rows[0]]

    # "Pickup" มี 2 คอลัมน์: อันแรก = รหัส depot, อันสุดท้าย = จำนวนที่รับแล้ว
    pu_positions = [i for i, h in enumerate(header) if h == "Pickup"]
    i_pucode = pu_positions[0]
    i_puqty = pu_positions[-1]
    first_idx = {}
    for i, h in enumerate(header):
        first_idx.setdefault(h, i)

    def g(row, name, default=""):
        i = first_idx.get(name)
        return row[i] if i is not None and i < len(row) else default

    out = []
    for row in rows[1:]:
        bk = g(row, "BK No")
        if bk is None or str(bk).strip() == "":
            continue
        rem = {t: num(g(row, f"{t} Remaining")) for t in TYPE_COLS}
        booked = {t: num(g(row, t)) for t in TYPE_COLS}
        rec = {
            "bk": str(bk).strip(),
            "vsl": str(g(row, "VSL") or "").strip(),
            "voy": str(g(row, "VOY") or "").strip(),
            "etd": fmt_dt(g(row, "ETD")),
            "por": str(g(row, "POR") or "").strip(),
            "lod": str(g(row, "LOD") or "").strip(),
            "dis": str(g(row, "DIS") or "").strip(),
            "tpsz": str(g(row, "TPSZ") or "").strip(),
            "pucode": str(row[i_pucode] or "").strip(),
            "puname": str(g(row, "Pickup Name") or "").strip(),
            "trandt": fmt_dt(g(row, "TRAN DT")),
            "cust": str(g(row, "ORG CUST") or "").strip(),
            "commodity": str(g(row, "COMMODITY") or "").strip(),
            "traffic": str(g(row, "TRAFFIC ORDER") or "").strip(),
            "group": str(g(row, "Group") or "").strip(),
            "booked_qty": sum(booked.values()),
            "pickup_qty": num(row[i_puqty]),
            "balance": num(g(row, "Balance (Booked - Pickup)")),
            "rem": rem,
        }
        # balance เผื่อไม่มีค่า cache
        if not rec["balance"]:
            rec["balance"] = sum(rem.values())
        out.append(rec)
    return out


def load_rows_from_xls(path):
    """อ่านไฟล์บุ๊คต้นฉบับ (.xls) โดยตรง แล้วคำนวณยอดคงเหลือเอง (ใช้ตรรกะเดียวกับ
    build_booking_balance_report.py) — ไม่ต้องพึ่งค่าที่ cache ไว้ใน .xlsx / ไม่ต้องเปิด Excel ก่อน
    """
    import build_booking_balance_report as bbr

    headers, rows = bbr.read_xls(path)
    recs, _ = bbr.build_records(headers, rows)

    def raw_at(r, name):
        i = headers.index(name)
        return r["raw"][i]

    out = []
    for r in recs:
        rem = dict(zip(TYPE_COLS, r["remaining"]))
        out.append({
            "bk": r["bk"],
            "vsl": r["vsl"],
            "voy": r["voy"],
            "etd": fmt_dt(raw_at(r, "ETD")),
            "por": str(raw_at(r, "POR") or "").strip(),
            "lod": str(raw_at(r, "LOD") or "").strip(),
            "dis": str(raw_at(r, "DIS") or "").strip(),
            "tpsz": r["tpsz"],
            "pucode": r["code"],
            "puname": r["name"],
            "trandt": fmt_dt(raw_at(r, "TRAN DT")),
            "cust": str(raw_at(r, "ORG CUST") or "").strip(),
            "commodity": str(raw_at(r, "COMMODITY") or "").strip(),
            "traffic": str(raw_at(r, "TRAFFIC ORDER") or "").strip(),
            "group": bbr.group_of(r["code"]),
            "booked_qty": num(r["booked"]),
            "pickup_qty": num(r["pickup_qty"]),
            "balance": num(r["balance"]),
            "rem": {t: num(v) for t, v in rem.items()},
        })
    return out


HTML_TEMPLATE = r"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Booking Pending Pickup</title>
<style>
  :root{
    --bg:#f4f6f9; --card:#ffffff; --ink:#1f2937; --muted:#6b7280;
    --line:#e5e7eb; --accent:#1F4E78; --accent2:#2E75B6;
    --bkk:#2E75B6; --lch:#2f9e6b; --amber:#b7791f; --chip:#eef2f7;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);
       font-family:"Segoe UI",Roboto,"Helvetica Neue",Arial,"Noto Sans Thai",sans-serif;font-size:14px}
  header{background:var(--accent);color:#fff;padding:16px 22px;
         display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap}
  header h1{margin:0;font-size:28px;font-weight:700}
  header .sub{opacity:.8;font-size:12.5px;margin-top:3px}
  header .header-right{display:flex;align-items:center;gap:16px;flex:none}
  header .logo{flex:none;background:#fff;padding:5px 10px;border-radius:6px;display:flex;align-items:center}
  header .logo img{height:64px;display:block}
  .wrap{padding:18px 22px 60px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px}
  .kpi .v{font-size:24px;font-weight:700;line-height:1.1}
  .kpi .l{color:var(--muted);font-size:12px;margin-top:4px}
  .controls{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 14px;
            display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:14px}
  .controls input[type=text],.controls select{
     padding:8px 10px;border:1px solid var(--line);border-radius:8px;font-size:13.5px;background:#fff;color:var(--ink)}
  .controls input[type=text]{flex:1;min-width:240px}
  .controls button{padding:8px 12px;border:1px solid var(--accent2);background:var(--accent2);color:#fff;
     border-radius:8px;cursor:pointer;font-size:13px}
  .controls button.ghost{background:#fff;color:var(--accent2)}
  .charts{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
  @media(max-width:900px){.charts{grid-template-columns:1fr}}
  .layout{display:grid;grid-template-columns:1fr;gap:14px;align-items:start}
  .panel{background:var(--card);border:1px solid var(--line);border-radius:10px;overflow:hidden}
  .panel h3{margin:0;padding:10px 14px;font-size:13px;border-bottom:1px solid var(--line);background:#fafbfc}
  .tablewrap{overflow-x:auto;max-height:70vh;overflow-y:auto}
  table{border-collapse:collapse;width:100%;font-size:12.7px}
  th,td{padding:7px 9px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
  th{position:sticky;top:0;background:#eef2f7;cursor:pointer;user-select:none;font-weight:600}
  th.sortasc::after{content:" \25B2";color:var(--muted)}
  th.sortdesc::after{content:" \25BC";color:var(--muted)}
  td.num,th.num{text-align:right}
  th.rem{background:#e3ecf5;font-size:11px;padding:7px 6px}
  td.cust{max-width:160px;overflow:hidden;text-overflow:ellipsis}
  th.cust{max-width:160px}
  td.rem{color:#c9ced6;font-variant-numeric:tabular-nums}
  td.rem.has{color:var(--accent);font-weight:700;background:#eef5ff}
  .bal{color:var(--accent)}
  tbody tr:hover{background:#f6f9fd}
  tr.detail td{background:#fbfcfe;white-space:normal;font-size:12px;color:var(--muted)}
  .pill{display:inline-block;padding:1px 7px;border-radius:999px;font-size:11px;font-weight:600;color:#fff}
  .pill.BKK{background:var(--bkk)} .pill.LCH{background:var(--lch)} .pill.OTHER{background:#888}
  .bal{font-weight:700}
  .bar-row{display:flex;align-items:center;gap:8px;padding:6px 14px;font-size:12px}
  .bar-row .name{width:96px;flex:none;color:var(--muted);overflow:hidden;text-overflow:ellipsis}
  .bar-track{flex:1;background:#eef2f7;border-radius:5px;height:16px;display:flex;overflow:hidden}
  .bar-track .seg{height:100%}
  .bar-row .val{width:30px;flex:none;text-align:right;font-weight:600}
  .legend{display:flex;flex-wrap:wrap;gap:10px;padding:8px 14px 2px;font-size:11px;color:var(--muted)}
  .legend-item{display:inline-flex;align-items:center;gap:4px}
  .legend-item i{width:10px;height:10px;border-radius:2px;display:inline-block}
  .zone-head{display:flex;align-items:center;gap:6px;font-size:11px;font-weight:700;color:var(--ink);
             padding:6px 14px;background:#f3f6fa;border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
  .zone-head .dot{width:8px;height:8px;border-radius:50%;display:inline-block}
  .bar-chips{padding:0 14px 8px 118px;display:flex;flex-wrap:wrap;gap:4px 8px;font-size:10.5px;color:var(--muted)}
  .bar-chips span{display:inline-flex;align-items:center;gap:3px}
  .bar-chips i{width:7px;height:7px;border-radius:2px;display:inline-block}
  .muted{color:var(--muted)}
  .expander{cursor:pointer;color:var(--accent2);font-weight:700}
  .typechips span{display:inline-block;background:var(--chip);border-radius:6px;padding:1px 6px;margin:1px 3px 1px 0;font-size:11px}
</style>
</head>
<body>
<header>
  <div>
    <h1>Booking Pending Pickup</h1>
    <div class="sub">ที่มา: __SRC__ &nbsp;|&nbsp; อัพเดทเมื่อ __GEN__ &nbsp;|&nbsp; __NREC__ BK No ค้างรับ</div>
  </div>
  <div class="header-right">
    <div class="logo"><img src="logo.png" alt="Heung-A Line"></div>
  </div>
</header>
<div class="wrap">

  <div class="kpis" id="kpis"></div>

  <div class="charts">
    <div class="panel">
      <h3>ตู้ค้างรับ แยกตาม Pickup depot</h3>
      <div id="chart" style="padding:6px 0 12px"></div>
    </div>
    <div class="panel">
      <h3>ตู้ค้างรับ แยกตามชนิด</h3>
      <div id="chartType" style="padding:6px 0 12px"></div>
    </div>
  </div>

  <div class="controls">
    <input type="text" id="q" placeholder="ค้นหา: BK No / ลูกค้า / depot / commodity / VSL / VOY ...">
    <select id="fGroup"><option value="">โซนทั้งหมด</option><option>BKK</option><option>LCH</option></select>
    <select id="fPickup"><option value="">Pickup ทั้งหมด</option></select>
    <select id="fType">
      <option value="">ทุกชนิดตู้</option>
      <option value="GP">GP (ตู้แห้ง)</option>
      <option value="RE">RE (ตู้เย็น)</option>
      <option value="UT">UT (เปิดข้าง/หลังคา)</option>
      <option value="PC">PC (แฟลตแร็ค)</option>
    </select>
    <button class="ghost" id="clear">ล้างตัวกรอง</button>
  </div>

  <div class="layout">
    <div class="panel full">
      <h3 id="resultHead"></h3>
      <div class="tablewrap">
        <table id="tbl">
          <thead><tr>
            <th data-k="bk">BK No</th>
            <th data-k="tpsz">TPSZ</th>
            <th data-k="pucode">Pickup</th>
            <th data-k="cust" class="cust">ลูกค้า (ORG CUST)</th>
            <th data-k="booked_qty" class="num">Booked</th>
            <th data-k="pickup_qty" class="num">รับแล้ว</th>
            <th data-k="balance" class="num">คงเหลือ</th>
            <th data-k="rem_GP22" class="num rem">GP22</th>
            <th data-k="rem_GP42" class="num rem">GP42</th>
            <th data-k="rem_GP45" class="num rem">GP45</th>
            <th data-k="rem_RE22" class="num rem">RE22</th>
            <th data-k="rem_RE45" class="num rem">RE45</th>
            <th data-k="rem_UT22" class="num rem">UT22</th>
            <th data-k="rem_UT42" class="num rem">UT42</th>
            <th data-k="rem_PC22" class="num rem">PC22</th>
            <th data-k="rem_PC42" class="num rem">PC42</th>
          </tr></thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
    </div>
  </div>
</div>

<script>
const DATA = __DATA__;
const TYPE_COLS = __TYPES__;

const $ = s => document.querySelector(s);
const q = $('#q'), fGroup = $('#fGroup'), fPickup = $('#fPickup'), fType = $('#fType');
let sortKey = 'balance', sortDir = -1;

// เติม dropdown Pickup Name
[...new Set(DATA.map(d => d.puname).filter(Boolean))].sort()
  .forEach(n => { const o = document.createElement('option'); o.value = n; o.textContent = n; fPickup.appendChild(o); });

function filtered(){
  const s = q.value.trim().toLowerCase();
  const g = fGroup.value, p = fPickup.value, t = fType.value;
  return DATA.filter(d => {
    if (g && d.group !== g) return false;
    if (p && d.puname !== p) return false;
    if (t){
      const sum = TYPE_COLS.filter(c => c.startsWith(t)).reduce((a,c)=>a+(d.rem[c]||0),0);
      if (sum <= 0) return false;
    }
    if (s){
      const hay = [d.bk,d.vsl,d.voy,d.pucode,d.puname,d.cust,d.commodity,d.traffic,d.lod,d.dis,d.tpsz,d.etd]
        .join(' ').toLowerCase();
      if (!hay.includes(s)) return false;
    }
    return true;
  });
}

function keyVal(d, k){ return k.startsWith('rem_') ? (d.rem[k.slice(4)] || 0) : d[k]; }

function sortRows(rows){
  return rows.slice().sort((a,b) => {
    let x = keyVal(a, sortKey), y = keyVal(b, sortKey);
    if (typeof x === 'number' && typeof y === 'number') return (x-y)*sortDir;
    return String(x).localeCompare(String(y),'th')*sortDir;
  });
}

function esc(s){ return String(s??'').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

function render(){
  const rows = sortRows(filtered());
  const tbody = $('#tbody');
  tbody.innerHTML = '';
  const totBal = rows.reduce((a,d)=>a+d.balance,0);
  const totBk  = rows.length;
  $('#resultHead').innerHTML = `แสดง <b>${totBk}</b> BK No &nbsp;|&nbsp; ตู้ค้างรับรวม <b>${totBal}</b>`;

  for (const d of rows){
    const tr = document.createElement('tr');
    tr.innerHTML =
      `<td><span class="expander" data-bk="${esc(d.bk)}">${esc(d.bk)} &#9662;</span></td>`+
      `<td>${esc(d.tpsz)}</td>`+
      `<td>${esc(d.pucode)}</td>`+
      `<td class="cust" title="${esc(d.cust)}">${esc(d.cust)}</td>`+
      `<td class="num">${d.booked_qty}</td><td class="num">${d.pickup_qty}</td>`+
      `<td class="num bal">${d.balance}</td>`+
      TYPE_COLS.map(c => `<td class="num rem${d.rem[c]>0?' has':''}">${d.rem[c]>0?d.rem[c]:'·'}</td>`).join('');
    tbody.appendChild(tr);
  }
  drawChart(rows);
  drawTypeChart(rows);
}

const TYPE_COLORS = {
  GP22:'#2E75B6', GP42:'#5B9BD5', GP45:'#9DC3E6',
  RE22:'#2f9e6b', RE45:'#8fcdae',
  UT22:'#b7791f', UT42:'#e0ac4c',
  PC22:'#8e44ad', PC42:'#c39bd3'
};

function legendHtml(entries){
  return `<div class="legend">${entries.map(([label,color]) =>
    `<span class="legend-item"><i style="background:${color}"></i>${label}</span>`).join('')}</div>`;
}

// ตู้ค้างรับ แยกตาม Pickup depot — แต่ละแท่งแบ่งสัดส่วนตามชนิดตู้ (GP22..PC42)
function drawChart(rows){
  const by = {};
  rows.forEach(d => {
    const k = d.puname || '(ไม่ระบุ)';
    if (!by[k]) by[k] = { total: 0, types: {}, group: d.group || 'OTHER' };
    by[k].total += d.balance;
    TYPE_COLS.forEach(c => { by[k].types[c] = (by[k].types[c]||0) + (d.rem[c]||0); });
  });
  const items = Object.entries(by).sort((a,b)=>b[1].total-a[1].total);
  const max = Math.max(1, ...items.map(i=>i[1].total));
  const usedTypes = TYPE_COLS.filter(c => items.some(([,v]) => v.types[c] > 0));

  const renderBar = ([n,v]) => {
    const segs = TYPE_COLS.filter(c => v.types[c] > 0).map(c =>
      `<div class="seg" style="width:${v.types[c]/max*100}%;background:${TYPE_COLORS[c]}" title="${n} — ${c}: ${v.types[c]}"></div>`
    ).join('');
    const chips = TYPE_COLS.filter(c => v.types[c] > 0).map(c =>
      `<span><i style="background:${TYPE_COLORS[c]}"></i>${c}: ${v.types[c]}</span>`
    ).join('');
    return `<div class="bar-row"><div class="name" title="${esc(n)}">${esc(n.split(' ')[0])}</div>`+
      `<div class="bar-track">${segs}</div><div class="val">${v.total}</div></div>`+
      `<div class="bar-chips">${chips}</div>`;
  };

  let html = usedTypes.length ? legendHtml(usedTypes.map(c => [c, TYPE_COLORS[c]])) : '';
  let any = false;
  [['BKK','var(--bkk)'],['LCH','var(--lch)'],['OTHER','#888']].forEach(([zone,color]) => {
    const zItems = items.filter(([,v]) => v.group === zone);
    if (!zItems.length) return;
    any = true;
    html += `<div class="zone-head"><span class="dot" style="background:${color}"></span>${zone === 'OTHER' ? 'อื่นๆ' : zone}</div>`
      + zItems.map(renderBar).join('');
  });
  $('#chart').innerHTML = any ? html : '<div class="bar-row muted">ไม่มีข้อมูล</div>';
}

// ตู้ค้างรับ แยกตามชนิด — แต่ละแท่งแบ่งสัดส่วนตามโซน BKK / LCH
// ตู้ค้างรับ แยกตามชนิด — จัดกลุ่ม BKK ก่อน แล้ว LCH ตามด้วยชนิดตู้เรียงตามลำดับคงที่
function drawTypeChart(rows){
  const by = { BKK: {}, LCH: {}, OTHER: {} };
  rows.forEach(d => {
    const g = d.group === 'BKK' ? 'BKK' : d.group === 'LCH' ? 'LCH' : 'OTHER';
    TYPE_COLS.forEach(c => {
      const v = d.rem[c] || 0;
      if (!v) return;
      by[g][c] = (by[g][c]||0) + v;
    });
  });
  const max = Math.max(1, ...Object.values(by).flatMap(o => Object.values(o)));
  let html = '';
  let any = false;
  [['BKK','var(--bkk)'],['LCH','var(--lch)'],['OTHER','#888']].forEach(([zone,color]) => {
    const typesInZone = TYPE_COLS.filter(c => by[zone][c] > 0);
    if (!typesInZone.length) return;
    any = true;
    html += `<div class="zone-head"><span class="dot" style="background:${color}"></span>${zone === 'OTHER' ? 'อื่นๆ' : zone}</div>`;
    html += typesInZone.map(c => {
      const v = by[zone][c];
      return `<div class="bar-row"><div class="name">${c}</div>`+
        `<div class="bar-track"><div class="seg" style="width:${v/max*100}%;background:${color}" title="${zone} — ${c}: ${v}"></div></div>`+
        `<div class="val">${v}</div></div>`;
    }).join('');
  });
  $('#chartType').innerHTML = any ? html : '<div class="bar-row muted">ไม่มีข้อมูล</div>';
}

function drawKpis(){
  const bk = DATA.length;
  const bal = DATA.reduce((a,d)=>a+d.balance,0);
  const dep = new Set(DATA.map(d=>d.puname).filter(Boolean)).size;
  const bkk = DATA.filter(d=>d.group==='BKK').reduce((a,d)=>a+d.balance,0);
  const lch = DATA.filter(d=>d.group==='LCH').reduce((a,d)=>a+d.balance,0);
  $('#kpis').innerHTML = [
    ['BK No ค้างรับ', bk],
    ['ตู้ค้างรับรวม', bal],
    ['จำนวน Pickup depot', dep],
    ['ค้างรับ โซน BKK', bkk],
    ['ค้างรับ โซน LCH', lch],
  ].map(([l,v]) => `<div class="kpi"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
}

// expand รายละเอียด
$('#tbody').addEventListener('click', e => {
  const ex = e.target.closest('.expander'); if (!ex) return;
  const tr = ex.closest('tr');
  if (tr.nextElementSibling && tr.nextElementSibling.classList.contains('detail')){
    tr.nextElementSibling.remove(); return;
  }
  const d = DATA.find(x => x.bk === ex.dataset.bk); if (!d) return;
  const chips = TYPE_COLS.filter(c => d.rem[c] > 0).map(c => `<span>${c}: ${d.rem[c]}</span>`).join('') || '<span>-</span>';
  const dr = document.createElement('tr');
  dr.className = 'detail';
  dr.innerHTML = `<td colspan="16">
     <b>VSL:</b> ${esc(d.vsl)||'-'} &nbsp; <b>VOY:</b> ${esc(d.voy)||'-'} &nbsp; <b>DIS:</b> ${esc(d.dis)||'-'}<br>
     <b>ETD:</b> ${esc(d.etd)||'-'} &nbsp; <b>POR:</b> ${esc(d.por)||'-'} &nbsp; <b>LOD:</b> ${esc(d.lod)||'-'} &nbsp;
     <b>TRAN DT:</b> ${esc(d.trandt)||'-'}<br>
     <b>Pickup Name:</b> ${esc(d.puname)||'-'}<br>
     <b>Commodity:</b> ${esc(d.commodity)||'-'}<br>
     <b>TRAFFIC ORDER:</b> ${esc(d.traffic)||'-'}<br>
     <b>คงเหลือแยกชนิด:</b> <span class="typechips">${chips}</span></td>`;
  tr.after(dr);
});

// sort
document.querySelectorAll('#tbl th[data-k]').forEach(th => {
  th.addEventListener('click', () => {
    const k = th.dataset.k;
    if (sortKey === k) sortDir *= -1; else { sortKey = k; sortDir = (k==='balance'||k==='booked_qty'||k==='pickup_qty'||k.startsWith('rem_')) ? -1 : 1; }
    document.querySelectorAll('#tbl th').forEach(x => x.classList.remove('sortasc','sortdesc'));
    th.classList.add(sortDir === 1 ? 'sortasc' : 'sortdesc');
    render();
  });
});

[q,fGroup,fPickup,fType].forEach(el => el.addEventListener('input', render));
$('#clear').addEventListener('click', () => { q.value=''; fGroup.value=''; fPickup.value=''; fType.value=''; render(); });

drawKpis();
document.querySelector('#tbl th[data-k="balance"]').classList.add('sortdesc');
render();
</script>
</body>
</html>
"""


def write_daily_booking_excel(recs, out_dir, src_name):
    """สร้างไฟล์ Excel รูปแบบ 'Daily Booking' (บุ๊คค้างรับทั้งหมด 1 ชีต) — บันทึกลงโฟลเดอร์โปรเจกต์
    (PENDING BKG) เสมอ ไม่ว่าไฟล์ต้นทางจะอยู่โฟลเดอร์ไหน เรียกอัตโนมัติทุกครั้งที่รัน build_dashboard.py
    แทนปุ่มดาวน์โหลด Excel บนเว็บที่เอาออกไปแล้ว
    """
    import datetime

    import build_booking_balance_report as bbr
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter

    cols = (["BK No", "VSL", "VOY", "ETD", "POR", "LOD", "DIS", "TPSZ", "Pickup", "Pickup Name",
             "TRAN DT", "ORG CUST", "Commodity", "TRAFFIC ORDER", "Group", "Booked", "PickedUp", "Balance"]
            + [f"{t} Remaining" for t in TYPE_COLS])
    ncol = len(cols)

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Booking"
    bbr.put_title(ws, f"DAILY BOOKING — Outstanding pickups   |   source: {src_name}", ncol)
    for c, h in enumerate(cols, start=1):
        ws.cell(row=2, column=c, value=h)
    bbr.style_header_row(ws, 2, ncol)

    rows_sorted = sorted(recs, key=lambda d: (d["group"] or "ZZZ", d["puname"], d["bk"]))
    r = 3
    for d in rows_sorted:
        vals = [d["bk"], d["vsl"], d["voy"], d["etd"], d["por"], d["lod"], d["dis"], d["tpsz"],
                d["pucode"], d["puname"], d["trandt"], d["cust"], d["commodity"], d["traffic"],
                d["group"], d["booked_qty"], d["pickup_qty"], d["balance"]]
        vals += [d["rem"].get(t, 0) for t in TYPE_COLS]
        for c, v in enumerate(vals, start=1):
            bbr.write_cell(ws, r, c, v)
        r += 1

    tot = r
    bbr.write_cell(ws, tot, 1, "TOTAL", bold=True, fill=bbr.FILL_TOTAL)
    for c in range(2, 16):
        bbr.write_cell(ws, tot, c, None, bold=True, fill=bbr.FILL_TOTAL)
    for c in range(16, ncol + 1):
        L = get_column_letter(c)
        bbr.write_cell(ws, tot, c, f"=SUM({L}3:{L}{tot - 1})", bold=True, fill=bbr.FILL_TOTAL)

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}{tot - 1}"
    bbr.set_widths(ws, [18, 8, 8, 15, 8, 8, 8, 10, 9, 32, 10, 28, 24, 26, 7, 8, 9, 9] + [10] * len(TYPE_COLS))

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    path = os.path.join(out_dir, f"Daily Booking {today}.xlsx")
    wb.save(path)
    return path


def main():
    if len(sys.argv) > 1:
        src = sys.argv[1]
    else:
        cands = (glob.glob("*Balance Summary*Copy*.xlsx") or glob.glob("*Balance Summary*.xlsx")
                 or glob.glob("output/*Balance Summary*.xlsx") or glob.glob("*PD*.xls"))
        if not cands:
            raise SystemExit("ระบุ path ไฟล์ .xls ต้นฉบับ หรือ .xlsx (ชีต Data) เป็น argument")
        src = cands[0]
    src = os.path.abspath(src)
    print(f"[src] {src}")

    if src.lower().endswith(".xls"):
        recs = load_rows_from_xls(src)  # คำนวณเองจากไฟล์ต้นฉบับ ไม่ต้องพึ่ง Excel
    else:
        recs = load_rows(src)
    if not recs:
        raise SystemExit("อ่านข้อมูลไม่ได้ — เปิดไฟล์ .xlsx ใน Excel แล้ว Save 1 ครั้งก่อน (ต้องมีค่าที่คำนวณแล้ว)")
    print(f"[data] {len(recs)} รายการค้างรับ, ตู้รวม {sum(r['balance'] for r in recs)}")

    import datetime
    html_out = (HTML_TEMPLATE
                .replace("__DATA__", json.dumps(recs, ensure_ascii=False))
                .replace("__TYPES__", json.dumps(TYPE_COLS))
                .replace("__SRC__", html.escape(os.path.basename(src)))
                .replace("__GEN__", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
                .replace("__NREC__", str(len(recs))))

    out = os.path.join(os.path.dirname(src), "Booking Balance Dashboard.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"[out] {out}")

    here = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = write_daily_booking_excel(recs, here, os.path.basename(src))
    print(f"[out] {xlsx_path}")


if __name__ == "__main__":
    main()
