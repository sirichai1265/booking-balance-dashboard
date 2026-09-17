#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
build_booking_balance_report.py
================================
สร้าง "สรุปบุ๊คคงเหลือ (Booking Balance Summary)" ตาม skill booking-balance-summary-report

ทำงาน 2 ขั้น:
  STAGE 1 (ไม่ใช้ AI) : อ่านไฟล์ .xls ด้วย xlrd -> ดึง text/ข้อมูลดิบออกมาเป็น
                        CSV + JSON (โฟลเดอร์ _extracted/) เพื่อให้ตรวจ/ให้ AI วิเคราะห์ได้
  STAGE 2            : คำนวณยอดคงเหลือต่อ BK No, กรองเฉพาะที่ยังค้าง (Balance > 0),
                        ทำความสะอาดข้อมูล, แล้วสร้างไฟล์ Excel แบบผูกสูตร ลงโฟลเดอร์ที่ชื่อตรงกับ
                        ไฟล์ต้นฉบับ (เช่น ไฟล์ "9-17-PD.xls" -> โฟลเดอร์ "9-17-PD/"):
                          - 1 ไฟล์รวม 3 ชีต : Data / Balance Summary / Summary
                          - ไฟล์แยกตาม Pickup Name : "bkg pending - <ชื่อ>.xlsx" — Calibri 11 ทั้งไฟล์,
                            row height 13 ทุกแถว, ไฮไลต์ทั้งแถวเหลือง/เขียวตาม TPSZ (20RF,R40H / 20OT,20FR,
                            40OT,40FR), TRAFFIC ORDER ที่มีคำว่า precool เป็นตัวแดงหนา, ตัดคอลัมน์ DOC CUST
                            เสมอ, ตัด COMMON REMARK ยกเว้นไฟล์กลุ่ม BKK01/BKK02/BKK04/LCH55, ปิดท้ายด้วย
                            สรุปยอดรวม + แยกตาม Pickup Name (กรณีไฟล์รวมหลาย code เช่น BKK01+BKK04)

การใช้งาน:
    python build_booking_balance_report.py "9-9-PD - Copy.xls"
    python build_booking_balance_report.py "9-14-PD.xls" "9-15-BKGWK.xls"   # รวม/เพิ่มข้อมูลจากหลายไฟล์
                                                                             # (ไฟล์หลังทับ BK No ซ้ำของไฟล์ก่อน)
    python build_booking_balance_report.py               # จะหาไฟล์ *PD*.xls ในโฟลเดอร์นี้เอง

ต้องมี lib:  pip install xlrd openpyxl
(openpyxl เปิดไฟล์ .xls รุ่นเก่าไม่ได้ ต้องใช้ xlrd อ่าน; ไฟล์ผลลัพธ์เป็น .xlsx)
Excel จะ recalc สูตรให้อัตโนมัติตอนเปิดไฟล์ (ถ้ามี LibreOffice ก็ใช้ recalc.py ได้)
"""

import csv
import glob
import json
import os
import re
import sys

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# ค่าคงที่ / สไตล์
# ----------------------------------------------------------------------------
TYPE_COLS = ["GP22", "GP42", "GP45", "RE22", "RE45", "UT22", "UT42", "PC22", "PC42"]

FONT_NAME = "Calibri"
C_TITLE_BG = "1F4E78"
C_HEADER_BG = "2E75B6"
C_TOTAL_BG = "D9E1F2"
C_GROUPSUM_BG = "FCE4D6"
C_FLAG_BG = "FFF2A8"
C_WHITE = "FFFFFF"

THIN = Side(style="thin", color="BDD7EE")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

F_TITLE = Font(name=FONT_NAME, bold=True, color=C_WHITE, size=13)
F_HEADER = Font(name=FONT_NAME, bold=True, color=C_WHITE)
F_CELL = Font(name=FONT_NAME)
F_BOLD = Font(name=FONT_NAME, bold=True)

FILL_TITLE = PatternFill("solid", fgColor=C_TITLE_BG)
FILL_HEADER = PatternFill("solid", fgColor=C_HEADER_BG)
FILL_TOTAL = PatternFill("solid", fgColor=C_TOTAL_BG)
FILL_GROUPSUM = PatternFill("solid", fgColor=C_GROUPSUM_BG)
FILL_FLAG = PatternFill("solid", fgColor=C_FLAG_BG)

# ชื่อ Pickup Name รวม สำหรับ BKK01 + BKK04 (รวมเป็นแถว/ไฟล์เดียวเสมอ)
BKK0104_DISPLAY = "PAT TERMINAL 1 & 2 (PORT AUTHORITY OF THAILAND)"

# ---- กติกาเฉพาะไฟล์แยกตาม Pickup depot (bkg pending - <ชื่อ>.xlsx) ----
PRECOOL_RE = re.compile(r"pre-?cool", re.IGNORECASE)
COMMON_REMARK_KEEP_CODES = {"BKK01", "BKK02", "BKK04", "LCH55"}
C_HL_YELLOW = "FFFF00"
C_HL_GREEN = "92D050"


def tpsz_highlight_color(tpsz):
    """คืนสี fill ถ้า TPSZ เข้าเงื่อนไขไฮไลต์ทั้งแถว ไม่งั้นคืน None
    เหลือง: มี 20RF หรือ R40H  ;  เขียว: มี 20OT, 20FR, 40OT หรือ 40FR
    """
    s = str(tpsz or "").upper()
    if "20RF" in s or "R40H" in s:
        return C_HL_YELLOW
    if any(tok in s for tok in ("20OT", "20FR", "40OT", "40FR")):
        return C_HL_GREEN
    return None


# ----------------------------------------------------------------------------
# STAGE 1 : อ่าน .xls -> ดึงข้อมูลดิบ (ไม่ใช้ AI)
# ----------------------------------------------------------------------------
def read_xls(path):
    """คืน (headers, rows) โดย rows เป็น list ของ list ตามคอลัมน์ต้นฉบับ
    - ตรวจหาแถว header อัตโนมัติ (แถวที่มีคำว่า 'BK No')
    - ตัดแถวสุดท้าย (footer รวมยอด) ทิ้ง
    """
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)

    header_row = None
    for r in range(min(sheet.nrows, 10)):
        rowvals = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        if "BK No" in rowvals:
            header_row = r
            break
    if header_row is None:
        raise SystemExit("หา header row (คอลัมน์ 'BK No') ไม่พบ")

    headers = [str(sheet.cell_value(header_row, c)).strip() for c in range(sheet.ncols)]

    rows = []
    for r in range(header_row + 1, sheet.nrows):
        vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        rows.append(vals)

    # แถวสุดท้ายเป็น footer รวมยอด (BK No ว่าง) -> ตัดทิ้ง
    while rows and str(rows[-1][0]).strip() == "":
        rows.pop()

    return headers, rows


def dump_extracted(headers, rows, outdir):
    """เขียนข้อมูลดิบเป็น CSV + JSON เพื่อให้ตรวจสอบ / ให้ AI วิเคราะห์"""
    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "9-9-PD_raw.csv")
    json_path = os.path.join(outdir, "9-9-PD_raw.json")

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for row in rows:
            w.writerow([_clean_text(v) for v in row])

    records = [
        {headers[i]: _json_val(v) for i, v in enumerate(row)}
        for row in rows
    ]
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=1)

    return csv_path, json_path


def _clean_text(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).replace("\n", " \\n ").strip()


def _json_val(v):
    if isinstance(v, float):
        return int(v) if v.is_integer() else v
    return str(v).strip()


# ----------------------------------------------------------------------------
# STAGE 2 : คำนวณ / กรอง / ทำความสะอาด
# ----------------------------------------------------------------------------
def col_index(headers, name, last=False):
    idxs = [i for i, h in enumerate(headers) if h == name]
    if not idxs:
        raise SystemExit(f"ไม่พบคอลัมน์ {name!r}")
    return idxs[-1] if last else idxs[0]


def build_records(headers, rows):
    """คำนวณ Balance, กรอง > 0, ทำความสะอาด, คืน list ของ dict"""
    i_bk = col_index(headers, "BK No")
    i_vsl = col_index(headers, "VSL")
    i_voy = col_index(headers, "VOY")
    i_tpsz = col_index(headers, "TPSZ")
    i_pucode = col_index(headers, "Pickup", last=False)     # คอลัมน์รหัส depot
    i_puname = col_index(headers, "Pickup Name")
    i_doc = col_index(headers, "DOC CUST")
    i_org = col_index(headers, "ORG CUST")
    i_traffic = col_index(headers, "TRAFFIC ORDER")
    i_type = [col_index(headers, t) for t in TYPE_COLS]
    i_puqty = col_index(headers, "Pickup", last=True)       # คอลัมน์จำนวนที่รับแล้ว

    recs = []
    for vals in rows:
        if str(vals[i_bk]).strip() == "":
            continue
        booked_types = [float(vals[i] or 0) for i in i_type]
        booked = sum(booked_types)
        pickup_qty = float(vals[i_puqty] or 0)
        balance = booked - pickup_qty
        if balance <= 0:
            continue

        org = str(vals[i_org]).strip()
        doc = str(vals[i_doc]).strip()
        if org == "":
            org = doc
        traffic = str(vals[i_traffic]).strip()
        if traffic == "":
            traffic = "GOOD AND CLEAN CONTAINER"

        # แถวต้นฉบับ (คงคอลัมน์เดิมทั้งหมด) + เขียนทับเฉพาะช่องว่างที่ต้อง clean
        raw = list(vals)
        raw[i_org] = org
        raw[i_traffic] = traffic

        code = str(vals[i_pucode]).strip()
        name = str(vals[i_puname]).strip()

        recs.append({
            "raw": raw,
            "bk": str(vals[i_bk]).strip(),
            "vsl": str(vals[i_vsl]).strip(),
            "voy": str(vals[i_voy]).strip(),
            "tpsz": str(vals[i_tpsz]).strip(),
            "code": code,
            "name": name,
            "types": booked_types,
            "booked": booked,
            "pickup_qty": pickup_qty,
            "balance": balance,
            "remaining": waterfall_remaining(booked_types, pickup_qty),
            "no_name": name == "",
        })
    return recs, {
        "i_type": i_type, "i_puqty": i_puqty, "i_pucode": i_pucode,
        "i_puname": i_puname,
    }


def waterfall_remaining(types, pickup_qty):
    """หักจำนวนที่รับแล้วแบบน้ำตก ตามลำดับ GP22->...->PC42"""
    out = []
    cum = 0.0
    for t in types:
        deduct = min(t, max(0.0, pickup_qty - cum))
        out.append(t - deduct)
        cum += t
    return out


def group_of(code):
    if code == "":
        return ""
    if code[:3] == "BKK" or code == "LCH55":
        return "BKK"
    if code[:3] == "LCH":
        return "LCH"
    return "OTHER"


def sanitize_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', "-", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(".").strip()
    return name


# ----------------------------------------------------------------------------
# ตัวช่วยเขียนชีต
# ----------------------------------------------------------------------------
def style_header_row(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HEADER
        cell.fill = FILL_HEADER
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def put_title(ws, text, ncol, row=1):
    ws.cell(row=row, column=1, value=text).font = F_TITLE
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    for c in range(1, ncol + 1):
        ws.cell(row=row, column=c).fill = FILL_TITLE
    ws.row_dimensions[row].height = 20


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_cell(ws, r, c, value, *, bold=False, fill=None, num_fmt=None, border=True):
    cell = ws.cell(row=r, column=c, value=value)
    cell.font = F_BOLD if bold else F_CELL
    if fill:
        cell.fill = fill
    if border:
        cell.border = BORDER
    if num_fmt:
        cell.number_format = num_fmt
    return cell


# ----------------------------------------------------------------------------
# OUTPUT 1 : ไฟล์รวม 3 ชีต
# ----------------------------------------------------------------------------
def _force_recalc(wb):
    # ให้ Excel คำนวณสูตรใหม่ทั้งหมดตอนเปิดไฟล์ (เพราะเราไม่ได้ cache ค่าไว้)
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass


def build_combined(recs, headers, name_by_code, out_path, src_label):
    wb = Workbook()

    # ---------- ชีต 1 : Data ----------
    ws = wb.active
    ws.title = "Data"
    n_orig = len(headers)

    # ตำแหน่งคอลัมน์บนชีต Data
    i_type = [headers.index(t) + 1 for t in TYPE_COLS]        # 1-based
    L_type = [get_column_letter(x) for x in i_type]
    i_puqty = len(headers) - 1                                # คอลัมน์ Pickup(qty) = ก่อน Return
    L_puqty = get_column_letter(i_puqty)
    i_pucode = headers.index("Pickup") + 1
    L_pucode = get_column_letter(i_pucode)

    col_balance = n_orig + 1
    col_group = n_orig + 2
    col_rem0 = n_orig + 3
    L_balance = get_column_letter(col_balance)
    L_group = get_column_letter(col_group)
    L_rem = [get_column_letter(col_rem0 + k) for k in range(9)]

    full_headers = list(headers) + ["Balance (Booked - Pickup)", "Group"] + [f"{t} Remaining" for t in TYPE_COLS]
    ncol = len(full_headers)

    put_title(ws, f"DATA — Outstanding bookings (Balance > 0)   |   source: {src_label}", ncol)
    for c, h in enumerate(full_headers, start=1):
        ws.cell(row=2, column=c, value=h)
    style_header_row(ws, 2, ncol)

    recs_sorted = sorted(recs, key=lambda x: (x["name"], x["bk"]))
    for idx, rec in enumerate(recs_sorted):
        r = idx + 3
        flag = FILL_FLAG if rec["no_name"] else None
        for c, v in enumerate(rec["raw"], start=1):
            write_cell(ws, r, c, _num(v), fill=flag)
        # Balance
        write_cell(ws, r, col_balance,
                   f"=SUM({L_type[0]}{r}:{L_type[-1]}{r})-{L_puqty}{r}", fill=flag)
        # Group
        write_cell(ws, r, col_group,
                   f'=IF({L_pucode}{r}="","",IF(OR(LEFT({L_pucode}{r},3)="BKK",{L_pucode}{r}="LCH55"),"BKK",'
                   f'IF(LEFT({L_pucode}{r},3)="LCH","LCH","OTHER")))', fill=flag)
        # Remaining (waterfall)
        for k in range(9):
            if k == 0:
                f = f"=MAX(0,{L_type[0]}{r}-MAX(0,{L_puqty}{r}))"
            else:
                f = (f"=MAX(0,{L_type[k]}{r}-MAX(0,{L_puqty}{r}-"
                     f"SUM({L_type[0]}{r}:{L_type[k-1]}{r})))")
            write_cell(ws, r, col_rem0 + k, f, fill=flag)

    last_row = len(recs_sorted) + 2
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(ncol)}{last_row}"
    set_widths(ws, [18, 8, 8, 15, 6, 8, 8, 14, 9, 34, 12, 30, 30, 26, 30, 30]
               + [7] * 11 + [22, 8] + [15] * 9)

    _flag_note(ws, recs_sorted, last_row + 2, 1)

    # เก็บ mapping row บน Data เพื่ออ้างอิงในชีตอื่น
    data_row_of = {rec["bk"] + "|" + rec["code"]: idx + 3 for idx, rec in enumerate(recs_sorted)}

    # ---------- ชีต 2 : Balance Summary (ต่อ BK No) ----------
    ws2 = wb.create_sheet("Balance Summary")
    cols2 = (["BK No", "VSL", "VOY", "Pickup", "Pickup Name", "TPSZ",
              "Booked Qty", "Pickup Qty", "Balance (Return)"]
             + [f"{t} Remaining" for t in TYPE_COLS])
    ncol2 = len(cols2)
    put_title(ws2, "BALANCE SUMMARY — remaining containers per BK No", ncol2)
    ws2.cell(row=2, column=1,
             value=("Balance = SUM(GP22..PC42) − Pickup qty.  Only BK No with Balance > 0 are listed. "
                    "Rows with no Pickup Name are highlighted amber and excluded from the grouped Summary."))
    ws2.cell(row=2, column=1).font = Font(name=FONT_NAME, italic=True)
    ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncol2)

    for c, h in enumerate(cols2, start=1):
        ws2.cell(row=3, column=c, value=h)
    style_header_row(ws2, 3, ncol2)

    bs_rows = sorted(recs_sorted, key=lambda x: (x["name"], x["bk"]))
    r = 4
    for rec in bs_rows:
        dr = data_row_of[rec["bk"] + "|" + rec["code"]]
        flag = FILL_FLAG if rec["no_name"] else None
        write_cell(ws2, r, 1, f"=Data!A{dr}", fill=flag)
        write_cell(ws2, r, 2, f"=Data!B{dr}", fill=flag)
        write_cell(ws2, r, 3, f"=Data!C{dr}", fill=flag)
        write_cell(ws2, r, 4, f"=Data!{L_pucode}{dr}", fill=flag)
        write_cell(ws2, r, 5, f"=Data!J{dr}", fill=flag)
        write_cell(ws2, r, 6, f"=Data!H{dr}", fill=flag)
        write_cell(ws2, r, 7, f"=SUM(Data!{L_type[0]}{dr}:Data!{L_type[-1]}{dr})", fill=flag)
        write_cell(ws2, r, 8, f"=Data!{L_puqty}{dr}", fill=flag)
        write_cell(ws2, r, 9, f"=Data!{L_balance}{dr}", fill=flag)
        for k in range(9):
            write_cell(ws2, r, 10 + k, f"=Data!{L_rem[k]}{dr}", fill=flag)
        r += 1

    tot = r
    write_cell(ws2, tot, 1, "TOTAL", bold=True, fill=FILL_TOTAL)
    for c in range(2, 7):
        write_cell(ws2, tot, c, None, bold=True, fill=FILL_TOTAL)
    for c in (7, 8, 9):
        L = get_column_letter(c)
        write_cell(ws2, tot, c, f"=SUM({L}4:{L}{r-1})", bold=True, fill=FILL_TOTAL)
    for k in range(9):
        L = get_column_letter(10 + k)
        write_cell(ws2, tot, 10 + k, f"=SUM({L}4:{L}{r-1})", bold=True, fill=FILL_TOTAL)

    ws2.freeze_panes = "A4"
    ws2.auto_filter.ref = f"A3:{get_column_letter(ncol2)}{r-1}"
    set_widths(ws2, [18, 8, 8, 9, 34, 14, 11, 10, 13] + [15] * 9)
    _flag_note(ws2, bs_rows, tot + 2, 1)

    # ---------- ชีต 3 : Summary (จัดกลุ่มตาม Pickup) ----------
    ws3 = wb.create_sheet("Summary")
    cols3 = ["Group", "Pickup", "Pickup Name", "Number of Records"] + TYPE_COLS + ["Total Containers"]
    ncol3 = len(cols3)
    put_title(ws3, "SUMMARY — outstanding containers grouped by Pickup depot", ncol3)
    for c, h in enumerate(cols3, start=1):
        ws3.cell(row=2, column=c, value=h)
    style_header_row(ws3, 2, ncol3)

    codes_present = sorted({rec["code"] for rec in recs_sorted if rec["code"] != ""})
    entries = []
    used = set()
    for code in codes_present:
        if code in ("BKK01", "BKK04"):
            continue
        entries.append({"group": group_of(code), "codes": [code],
                        "name": name_by_code.get(code, "")})
    if "BKK01" in codes_present or "BKK04" in codes_present:
        entries.append({"group": "BKK", "codes": ["BKK01", "BKK04"], "name": BKK0104_DISPLAY})
    entries.sort(key=lambda e: (0 if e["group"] == "BKK" else 1 if e["group"] == "LCH" else 2, e["codes"][0]))

    rem_ranges = [f"Data!${L_rem[k]}:${L_rem[k]}" for k in range(9)]
    r = 3
    first_data_r = r
    for e in entries:
        codes = e["codes"]
        write_cell(ws3, r, 1, e["group"])
        write_cell(ws3, r, 2, "+".join(codes))
        write_cell(ws3, r, 3, e["name"])
        write_cell(ws3, r, 4, "=" + "+".join(f'COUNTIF(Data!$I:$I,"{c}")' for c in codes), num_fmt="0")
        for k in range(9):
            f = "=" + "+".join(f'SUMIF(Data!$I:$I,"{c}",{rem_ranges[k]})' for c in codes)
            write_cell(ws3, r, 5 + k, f, num_fmt="0")
        write_cell(ws3, r, 14, f"=SUM(E{r}:M{r})", num_fmt="0")
        r += 1

    tot = r
    write_cell(ws3, tot, 1, "TOTAL", bold=True, fill=FILL_TOTAL)
    write_cell(ws3, tot, 2, None, bold=True, fill=FILL_TOTAL)
    write_cell(ws3, tot, 3, None, bold=True, fill=FILL_TOTAL)
    for c in range(4, 15):
        L = get_column_letter(c)
        write_cell(ws3, tot, c, f"=SUM({L}{first_data_r}:{L}{r-1})", bold=True, fill=FILL_TOTAL, num_fmt="0")

    # ----- Group Summary (เริ่มคอลัมน์ C, เว้น 2 แถวหลัง TOTAL) -----
    gs = tot + 3
    gs_cols = ["Group", "Number of Records"] + TYPE_COLS + ["Total Containers"]
    for i, h in enumerate(gs_cols):
        write_cell(ws3, gs, 3 + i, h, bold=True, fill=FILL_GROUPSUM)
    for gi, grp in enumerate(["BKK", "LCH"]):
        rr = gs + 1 + gi
        write_cell(ws3, rr, 3, grp, bold=True, fill=FILL_GROUPSUM)
        write_cell(ws3, rr, 4, f'=COUNTIF(Data!${L_group}:${L_group},"{grp}")',
                   bold=True, fill=FILL_GROUPSUM, num_fmt="0")
        for k in range(9):
            write_cell(ws3, rr, 5 + k,
                       f'=SUMIF(Data!${L_group}:${L_group},"{grp}",{rem_ranges[k]})',
                       bold=True, fill=FILL_GROUPSUM, num_fmt="0")
        write_cell(ws3, rr, 14, f"=SUM(E{rr}:M{rr})", bold=True, fill=FILL_GROUPSUM, num_fmt="0")
    gt = gs + 3
    write_cell(ws3, gt, 3, "GRAND TOTAL", bold=True, fill=FILL_TOTAL)
    for c in range(4, 15):
        L = get_column_letter(c)
        write_cell(ws3, gt, c, f"=SUM({L}{gs+1}:{L}{gs+2})", bold=True, fill=FILL_TOTAL, num_fmt="0")

    ws3.freeze_panes = "A3"
    set_widths(ws3, [8, 14, 40, 18] + [7] * 9 + [16])

    _force_recalc(wb)
    wb.save(out_path)
    return out_path, entries, name_by_code


def _num(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _flag_note(ws, rows, r, c):
    bad = [x["bk"] for x in rows if x["no_name"]]
    if not bad:
        return
    ws.cell(row=r, column=c,
            value=(f"NOTE: {len(bad)} row(s) highlighted above have NO Pickup Name "
                   f"(Pickup code also blank) — verify the pickup depot for BK No: "
                   + ", ".join(bad))).font = F_BOLD


def full_name_map(headers, rows):
    """สร้าง mapping รหัส Pickup -> ชื่อ Pickup จาก (headers, rows) ที่โหลดไว้แล้ว
    (เผื่อบางรหัสไม่มีแถวค้างรอบนี้ เช่น BKK04 ที่ถูก merge เข้ากับ BKK01)
    """
    i_code = headers.index("Pickup")
    i_name = headers.index("Pickup Name")
    out = {}
    for v in rows:
        code = str(v[i_code]).strip()
        name = str(v[i_name]).strip()
        if code and name and code not in out:
            out[code] = name
    return out


# ----------------------------------------------------------------------------
# OUTPUT 2 : ไฟล์แยกตาม Pickup Name
# ----------------------------------------------------------------------------
PP_WIDE_COLS = {"ORG CUST": 26, "COMMODITY": 24, "TRAFFIC ORDER": 28}
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autofit_width(header_text, raw_values, floor=1, pad=1):
    max_len = len(str(header_text))
    for v in raw_values:
        s = str(v)
        if len(s) > max_len:
            max_len = len(s)
    return max(max_len + pad, floor)


def build_per_pickup(recs, headers, entries, name_by_code, outdir):
    # ฟอนต์ Tahoma 8 เฉพาะไฟล์ชุดนี้ (ไม่แตะ F_CELL/F_HEADER/F_BOLD ที่ไฟล์รวมใช้ร่วมกัน)
    pp_font = Font(name="Tahoma", size=8)
    pp_font_bold = Font(name="Tahoma", size=8, bold=True)
    pp_font_header = Font(name="Tahoma", size=8, bold=True, color=C_WHITE)
    pp_font_precool = Font(name="Tahoma", size=8, bold=True, color="FF0000")
    ROW_H = 13

    i_pucode = headers.index("Pickup")
    i_puname = headers.index("Pickup Name")
    i_doc = headers.index("DOC CUST")
    i_common = headers.index("COMMON REMARK")
    i_tpsz = headers.index("TPSZ")
    i_traffic = headers.index("TRAFFIC ORDER")
    i_types = [headers.index(t) for t in TYPE_COLS]
    pu_positions = [i for i, h in enumerate(headers) if h == "Pickup"]
    i_puqty = pu_positions[-1]                     # คอลัมน์จำนวนที่รับแล้ว (ไม่ใช่ i_pucode)
    i_return = headers.index("Return")

    os.makedirs(outdir, exist_ok=True)
    made = []
    grand = 0.0
    for e in entries:
        codes = set(e["codes"])
        sub = [rec for rec in recs if rec["code"] in codes and not rec["no_name"]]
        if not sub:
            continue
        sub.sort(key=lambda x: x["bk"])

        # ตัด Pickup code/name/DOC CUST/ตู้ที่ booked/รับแล้ว/Return/Balance เสมอ (เหลือแต่ยอดคงเหลือสุทธิ
        # ต่อชนิดตู้) ; ตัด COMMON REMARK ยกเว้นกลุ่ม BKK01/BKK02/BKK04/LCH55
        drop = {i_pucode, i_puname, i_doc, i_puqty, i_return} | set(i_types)
        if not (codes & COMMON_REMARK_KEEP_CODES):
            drop.add(i_common)
        kept_idx = [i for i in range(len(headers)) if i not in drop]
        kept_headers = [headers[i] for i in kept_idx]

        col_rem0 = len(kept_headers) + 1
        ncol = col_rem0 + 8
        pos_traffic = kept_headers.index("TRAFFIC ORDER") + 1

        out_headers = kept_headers + list(TYPE_COLS)

        wb = Workbook()
        ws = wb.active
        ws.title = "Pending"
        for c, h in enumerate(out_headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = pp_font_header
            cell.fill = FILL_HEADER
            cell.border = BORDER
            cell.alignment = ALIGN_LEFT

        for idx, rec in enumerate(sub):
            r = idx + 2
            row_fill_color = tpsz_highlight_color(rec["raw"][i_tpsz])
            row_fill = PatternFill("solid", fgColor=row_fill_color) if row_fill_color else None
            kept_vals = [rec["raw"][i] for i in kept_idx]
            for c, v in enumerate(kept_vals, start=1):
                cell = ws.cell(row=r, column=c, value=_num(v))
                cell.font = pp_font
                cell.border = BORDER
                cell.alignment = ALIGN_LEFT
                if row_fill:
                    cell.fill = row_fill
            if PRECOOL_RE.search(str(rec["raw"][i_traffic] or "")):
                ws.cell(row=r, column=pos_traffic).font = pp_font_precool

            for k in range(9):
                rc = ws.cell(row=r, column=col_rem0 + k, value=_num(rec["remaining"][k]))
                rc.font = pp_font
                rc.border = BORDER
                rc.alignment = ALIGN_LEFT
                if row_fill:
                    rc.fill = row_fill
            ws.row_dimensions[r].height = ROW_H

        tr = len(sub) + 2
        for c in range(1, ncol + 1):
            L = get_column_letter(c)
            cell = ws.cell(row=tr, column=c)
            cell.font = pp_font_bold
            cell.fill = FILL_TOTAL
            cell.border = BORDER
            cell.alignment = ALIGN_LEFT
            if c == 1:
                cell.value = "TOTAL"
            elif c >= col_rem0:
                cell.value = f"=SUM({L}2:{L}{tr - 1})"
        ws.row_dimensions[tr].height = ROW_H

        # ---- สรุปยอดบุ๊คทั้งหมด + แยกตาม pickup name (เผื่อไฟล์รวมหลาย code เช่น BKK01+BKK04) ----
        by_name = {}
        for rec in sub:
            n, b = by_name.setdefault(rec["name"], [0, 0.0])
            by_name[rec["name"]][0] = n + 1
            by_name[rec["name"]][1] = b + rec["balance"]

        sr = tr + 3
        for c, h in enumerate(["Pickup Name", "Number of Records", "Total Balance"], start=1):
            cell = ws.cell(row=sr, column=c, value=h)
            cell.font = pp_font_bold
            cell.fill = FILL_GROUPSUM
            cell.border = BORDER
            cell.alignment = ALIGN_LEFT
        ws.row_dimensions[sr].height = ROW_H

        rr = sr + 1
        for name in sorted(by_name):
            n, b = by_name[name]
            for c, v in enumerate([name, n, b], start=1):
                cell = ws.cell(row=rr, column=c, value=v)
                cell.font = pp_font
                cell.border = BORDER
                cell.alignment = ALIGN_LEFT
            ws.row_dimensions[rr].height = ROW_H
            rr += 1

        for c, v in enumerate(["สรุปยอดบุ๊คทั้งหมด", len(sub), sum(rec["balance"] for rec in sub)], start=1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.font = pp_font_bold
            cell.fill = FILL_TOTAL
            cell.border = BORDER
            cell.alignment = ALIGN_LEFT
        ws.row_dimensions[rr].height = ROW_H

        # ไม่ freeze panes ; ความกว้างคอลัมน์ชิดตามจำนวนตัวอักษรจริง (ขั้นต่ำ 10) ยกเว้น
        # ORG CUST / COMMODITY / TRAFFIC ORDER ที่คงความกว้างคงที่ไว้ให้อ่านง่าย
        ws.freeze_panes = None
        ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}{tr - 1}"
        widths = []
        for oi, h in zip(kept_idx, kept_headers):
            if h in PP_WIDE_COLS:
                widths.append(PP_WIDE_COLS[h])
            else:
                vals = [_num(rec["raw"][oi]) for rec in sub]
                widths.append(_autofit_width(h, vals))
        for k, t in enumerate(TYPE_COLS):
            rem_vals = [_num(rec["remaining"][k]) for rec in sub]
            widths.append(_autofit_width(t, rem_vals))
        set_widths(ws, widths)

        ws.row_dimensions[1].height = ROW_H

        _force_recalc(wb)
        fname = f"bkg pending - {sanitize_filename(e['name'])}.xlsx"
        path = os.path.join(outdir, fname)
        wb.save(path)
        made.append(path)
        grand += sum(rec["balance"] for rec in sub)

    return made, grand


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------
def merge_sources(paths):
    """อ่านไฟล์ .xls หนึ่งไฟล์ขึ้นไปแล้วรวมแถวเข้าด้วยกันโดย key = BK No
    ไฟล์หลัง ๆ ใน list (ลำดับตามที่ระบุใน argument) จะ override แถวที่ BK No ซ้ำกับไฟล์ก่อนหน้า
    — ใช้เมื่อมีไฟล์เสริม เช่น ไฟล์รายสัปดาห์ ที่ต้องเอามา "เพิ่มข้อมูล" ทับไฟล์รายวันเดิม
    คืน (headers, merged_rows, info) โดย info คือ list ของ (ชื่อไฟล์, จำนวนแถว, BK No ใหม่, BK No ที่ถูกทับ)
    """
    headers = None
    merged = {}
    info = []
    for p in paths:
        h, rows = read_xls(p)
        if headers is None:
            headers = h
        elif h != headers:
            raise SystemExit(f"โครงสร้างคอลัมน์ของ {p} ไม่ตรงกับไฟล์แรก — รวมไฟล์นี้ไม่ได้")
        i_bk = h.index("BK No")
        new_bk = updated_bk = 0
        for row in rows:
            bk = str(row[i_bk]).strip()
            if not bk:
                continue
            if bk in merged:
                updated_bk += 1
            else:
                new_bk += 1
            merged[bk] = row
        info.append((os.path.basename(p), len(rows), new_bk, updated_bk))
    return headers, list(merged.values()), info


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) > 1:
        srcs = sys.argv[1:]
    else:
        cands = glob.glob(os.path.join(here, "*PD*.xls")) + glob.glob(os.path.join(here, "*PD*.xls*"))
        if not cands:
            raise SystemExit("ระบุไฟล์ .xls เป็น argument (ใส่ได้หลายไฟล์เพื่อรวม/เพิ่มข้อมูล) หรือวางไฟล์ *PD*.xls ไว้ในโฟลเดอร์นี้")
        srcs = [cands[0]]
    srcs = [os.path.abspath(s) for s in srcs]
    src_label = " + ".join(os.path.basename(s) for s in srcs)
    print(f"[src] {src_label}")

    # ---- STAGE 1 : ดึงข้อมูลดิบ (ไม่ใช้ AI) — รวมหลายไฟล์ถ้ามี ----
    headers, rows, merge_info = merge_sources(srcs)
    for name, n, new_bk, upd_bk in merge_info:
        print(f"[merge] {name}: {n} แถว (BK No ใหม่ {new_bk}, ทับของเดิม {upd_bk})")
    if len(srcs) > 1:
        print(f"[merge] รวมทั้งหมด {len(rows)} BK No (unique)")
    ex_dir = os.path.join(here, "_extracted")
    csv_p, json_p = dump_extracted(headers, rows, ex_dir)
    print(f"[stage1] header {len(headers)} คอลัมน์, ข้อมูล {len(rows)} แถว")
    print(f"[stage1] -> {csv_p}")
    print(f"[stage1] -> {json_p}")

    # ---- STAGE 2 : คำนวณ + สร้างรายงาน ----
    recs, _meta = build_records(headers, rows)
    print(f"[stage2] BK No ที่ยังค้าง (Balance > 0): {len(recs)} รายการ")
    by_group = {}
    for rec in recs:
        by_group.setdefault(group_of(rec["code"]) or "(blank)", [0, 0.0])
        by_group[group_of(rec["code"]) or "(blank)"][0] += 1
        by_group[group_of(rec["code"]) or "(blank)"][1] += rec["balance"]
    for g, (n, b) in sorted(by_group.items()):
        print(f"           {g:8} records={n:3}  balance={b:.0f}")
    total_balance = sum(rec["balance"] for rec in recs)
    print(f"[stage2] GRAND TOTAL balance = {total_balance:.0f}")

    # ชื่อโฟลเดอร์ผลลัพธ์ = ชื่อไฟล์ต้นฉบับ (ไฟล์แรกถ้าใส่หลายไฟล์) ไม่ใช้ "output" คงที่อีกต่อไป
    out_dir = os.path.join(here, os.path.splitext(os.path.basename(srcs[0]))[0])
    os.makedirs(out_dir, exist_ok=True)
    name_by_code = full_name_map(headers, rows)
    combined = os.path.join(out_dir, "Booking Balance Summary.xlsx")
    combined, entries, name_by_code = build_combined(recs, headers, name_by_code, combined, src_label)
    print(f"[out1] {combined}")

    made, grand = build_per_pickup(recs, headers, entries, name_by_code, out_dir)
    for p in made:
        print(f"[out2] {p}")

    # ---- consistency check ----
    excl = sum(rec["balance"] for rec in recs if rec["no_name"])
    ok = abs((total_balance - excl) - grand) < 1e-6
    for rec in recs:
        assert abs(sum(rec["remaining"]) - rec["balance"]) < 1e-6, rec["bk"]
    print(f"[check] sum(9 remaining) == balance ต่อแถว : OK")
    print(f"[check] per-file TOTAL balance รวม = {grand:.0f} , "
          f"combined GRAND TOTAL (ไม่รวม no-name) = {total_balance - excl:.0f} : "
          f"{'OK' if ok else 'MISMATCH'}")
    print("\nเสร็จ. เปิดไฟล์ใน Excel เพื่อให้สูตรคำนวณค่าอัตโนมัติ")


if __name__ == "__main__":
    main()
