# Booking Balance Dashboard

เครื่องมือแปลงไฟล์บุ๊ค (`.xls` แบบ 98PD/BKGPD) เป็นรายงานยอดคงเหลือต่อ BK No และ Dashboard
สำหรับค้นหา/กรองบุ๊คที่ยังค้างรับ

## ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|---|---|
| `build_booking_balance_report.py` | อ่าน `.xls` ต้นฉบับ (xlrd) → ดึงข้อมูลดิบ → คำนวณยอดคงเหลือต่อ BK No → สร้าง Excel ผูกสูตร (`output/Booking Balance Summary.xlsx`) + ไฟล์แยกตาม Pickup depot |
| `build_dashboard.py` | อ่านชีต `Data` จากไฟล์ Excel สรุปยอดคงเหลือ (openpyxl, `data_only`) → สร้าง `output/Booking Balance Dashboard.html` ซึ่งเป็น dashboard ค้นหา/กรอง/sort แบบ self-contained |
| `output/` | ไฟล์ผลลัพธ์ (Excel รายงาน + Dashboard HTML) |

## วิธีใช้

```bash
pip install xlrd openpyxl
python build_booking_balance_report.py "9-9-PD - Copy.xls"
python build_dashboard.py "output/Booking Balance Summary - Copy.xlsx"
```

เปิด `output/Booking Balance Dashboard.html` ในเบราว์เซอร์เพื่อค้นหา/กรองบุ๊คที่ค้างรับ
(ค้นหาข้อความ, กรองตามโซน BKK/LCH, Pickup depot, ชนิดตู้, sort ทุกคอลัมน์, ดาวน์โหลด CSV)

## หมายเหตุ

Repo นี้เป็น **private** เพราะข้อมูลตัวอย่างมีชื่อลูกค้า/รายละเอียดบุ๊คจริง — ไฟล์ `.xls` ต้นฉบับ
และข้อมูลดิบที่ดึงออกมา (`_extracted/`) ถูก `.gitignore` ไว้ไม่ให้เข้า repo
