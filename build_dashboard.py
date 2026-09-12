#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
build_dashboard.py
==================
สร้าง Dashboard (ไฟล์ HTML เดียว เปิดในเบราว์เซอร์ได้เลย) สำหรับ "ค้นหาบุ๊คที่ค้างรับ"
จากไฟล์ Booking Balance Summary (ชีต Data)

การใช้งาน:
    python build_dashboard.py "output/Booking Balance Summary - Copy.xlsx"
    python build_dashboard.py            # ไม่ใส่ path -> หา *Balance Summary*.xlsx ให้เอง

ต้องมี lib:  pip install openpyxl
- อ่านค่าที่ Excel คำนวณไว้แล้ว (data_only) ; ถ้ายังไม่เคยเปิดไฟล์ใน Excel ค่าจะยังว่าง
  ให้เปิดไฟล์ .xlsx ใน Excel แล้ว Save หนึ่งครั้งก่อน
ผลลัพธ์:  Booking Balance Dashboard.html  (ในโฟลเดอร์เดียวกับไฟล์ต้นทาง)
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


HTML_TEMPLATE = r"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Booking Balance Dashboard</title>
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
  header h1{margin:0;font-size:19px;font-weight:700}
  header .sub{opacity:.8;font-size:12.5px;margin-top:3px}
  header .header-right{display:flex;align-items:center;gap:16px;flex:none}
  header .clock{text-align:right;flex:none}
  header .clock .time{font-size:22px;font-weight:700;font-variant-numeric:tabular-nums;line-height:1.1}
  header .clock .date{opacity:.8;font-size:12.5px;margin-top:3px}
  header .logo{flex:none;background:#fff;padding:5px 10px;border-radius:6px;display:flex;align-items:center}
  header .logo img{height:32px;display:block}
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
  .bar-track{flex:1;background:#eef2f7;border-radius:5px;height:16px;position:relative}
  .bar-fill{background:var(--accent2);height:100%;border-radius:5px}
  .bar-row .val{width:30px;flex:none;text-align:right;font-weight:600}
  .muted{color:var(--muted)}
  .expander{cursor:pointer;color:var(--accent2);font-weight:700}
  .typechips span{display:inline-block;background:var(--chip);border-radius:6px;padding:1px 6px;margin:1px 3px 1px 0;font-size:11px}
</style>
</head>
<body>
<header>
  <div>
    <h1>Booking Balance Dashboard &mdash; บุ๊คที่ค้างรับ</h1>
    <div class="sub">ที่มา: __SRC__ &nbsp;|&nbsp; สร้างเมื่อ __GEN__ &nbsp;|&nbsp; __NREC__ BK No ค้างรับ</div>
  </div>
  <div class="header-right">
    <div class="clock">
      <div class="time" id="clockTime"></div>
      <div class="date" id="clockDate"></div>
    </div>
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
            <th data-k="group">โซน</th>
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
      `<td><span class="pill ${d.group||'OTHER'}">${esc(d.group||'-')}</span></td>`+
      `<td class="num">${d.booked_qty}</td><td class="num">${d.pickup_qty}</td>`+
      `<td class="num bal">${d.balance}</td>`+
      TYPE_COLS.map(c => `<td class="num rem${d.rem[c]>0?' has':''}">${d.rem[c]>0?d.rem[c]:'·'}</td>`).join('');
    tbody.appendChild(tr);
  }
  drawChart(rows);
  drawTypeChart(rows);
}

function drawChart(rows){
  const by = {};
  rows.forEach(d => { const k = d.puname || '(ไม่ระบุ)'; by[k] = (by[k]||0) + d.balance; });
  const items = Object.entries(by).sort((a,b)=>b[1]-a[1]);
  const max = Math.max(1, ...items.map(i=>i[1]));
  $('#chart').innerHTML = items.map(([n,v]) =>
    `<div class="bar-row"><div class="name" title="${esc(n)}">${esc(n.split(' ')[0])}</div>`+
    `<div class="bar-track"><div class="bar-fill" style="width:${v/max*100}%"></div></div>`+
    `<div class="val">${v}</div></div>`).join('') || '<div class="bar-row muted">ไม่มีข้อมูล</div>';
}

function drawTypeChart(rows){
  const by = {};
  TYPE_COLS.forEach(c => by[c] = 0);
  rows.forEach(d => TYPE_COLS.forEach(c => by[c] += (d.rem[c]||0)));
  const items = Object.entries(by).filter(i=>i[1]>0).sort((a,b)=>b[1]-a[1]);
  const max = Math.max(1, ...items.map(i=>i[1]));
  $('#chartType').innerHTML = items.map(([n,v]) =>
    `<div class="bar-row"><div class="name">${n}</div>`+
    `<div class="bar-track"><div class="bar-fill" style="width:${v/max*100}%"></div></div>`+
    `<div class="val">${v}</div></div>`).join('') || '<div class="bar-row muted">ไม่มีข้อมูล</div>';
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
  dr.innerHTML = `<td colspan="17">
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

function tickClock(){
  const now = new Date();
  $('#clockTime').textContent = now.toLocaleTimeString('en-US', {hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false});
  $('#clockDate').textContent = now.toLocaleDateString('en-US', {weekday:'long',year:'numeric',month:'long',day:'numeric'});
}
tickClock();
setInterval(tickClock, 1000);

drawKpis();
document.querySelector('#tbl th[data-k="balance"]').classList.add('sortdesc');
render();
</script>
</body>
</html>
"""


def main():
    if len(sys.argv) > 1:
        src = sys.argv[1]
    else:
        cands = (glob.glob("*Balance Summary*Copy*.xlsx") or glob.glob("*Balance Summary*.xlsx")
                 or glob.glob("output/*Balance Summary*.xlsx"))
        if not cands:
            raise SystemExit("ระบุ path ไฟล์ .xlsx (ชีต Data) เป็น argument")
        src = cands[0]
    src = os.path.abspath(src)
    print(f"[src] {src}")

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


if __name__ == "__main__":
    main()
