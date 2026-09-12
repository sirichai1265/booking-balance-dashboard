# Booking Balance Dashboard

เครื่องมือแปลงไฟล์บุ๊ค (`.xls` แบบ 98PD/BKGPD) เป็นรายงานยอดคงเหลือต่อ BK No และ Dashboard
สำหรับค้นหา/กรองบุ๊คที่ยังค้างรับ

🔗 Dashboard (live): https://sirichai1265.github.io/booking-balance-dashboard/

## ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|---|---|
| `build_booking_balance_report.py` | อ่าน `.xls` ต้นฉบับ (xlrd) → ดึงข้อมูลดิบ → คำนวณยอดคงเหลือต่อ BK No → สร้าง Excel ผูกสูตร (`output/Booking Balance Summary.xlsx`) + ไฟล์แยกตาม Pickup depot |
| `build_dashboard.py` | สร้าง `output/Booking Balance Dashboard.html` (dashboard ค้นหา/กรอง/sort แบบ self-contained) — อ่านได้ทั้งจากไฟล์ `.xls` ต้นฉบับโดยตรง (คำนวณเอง ไม่ต้องพึ่ง Excel) หรือจากไฟล์ `.xlsx` สรุปที่เปิด/Save ด้วย Excel แล้ว |
| `index.html` | สำเนาของ dashboard ที่ใช้ deploy บน GitHub Pages (root ของเว็บ) |
| `output/` | ไฟล์ผลลัพธ์ (Excel รายงาน + Dashboard HTML) |

## วิธีใช้

```bash
pip install xlrd openpyxl
python build_booking_balance_report.py "9-12-PD - Copy.xls"   # สร้างรายงาน Excel
python build_dashboard.py "9-12-PD - Copy.xls"                 # สร้าง dashboard จากไฟล์ต้นฉบับโดยตรง
cp "output/Booking Balance Dashboard.html" index.html          # sync ขึ้น GitHub Pages
```

เปิด `output/Booking Balance Dashboard.html` ในเบราว์เซอร์เพื่อค้นหา/กรองบุ๊คที่ค้างรับ
(ค้นหาข้อความ, กรองตามโซน BKK/LCH, Pickup depot, ชนิดตู้, sort ทุกคอลัมน์, คลิก BK No เพื่อดูรายละเอียด)

## หมายเหตุ

Repo นี้เป็น **public** (จำเป็นสำหรับ GitHub Pages บนแพลน Free) — ข้อมูลตัวอย่างใน dashboard/รายงาน
มีชื่อลูกค้า/รายละเอียดบุ๊คจริงเปิดเผยต่อสาธารณะ ไฟล์ `.xls` ต้นฉบับ (บุ๊คทั้งหมด ไม่ใช่แค่ที่ค้าง)
และข้อมูลดิบที่ดึงออกมา (`_extracted/`) ยังคง `.gitignore` ไว้ไม่ให้เข้า repo
