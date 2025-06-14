# 📊 Stock Support Calculator

โปรแกรมวิเคราะห์หุ้นบนเดสก์ท็อปด้วย **Python + Tkinter** ใช้งานง่าย ดึงข้อมูลจาก **Yahoo Finance** พร้อมกราฟสวยงามและฟีเจอร์ครบครัน

---

## 🚀 ฟีเจอร์หลัก

- 🔍 **ค้นหาหุ้น** จาก Yahoo Finance
- 📉 **คำนวณแนวรับ/แนวต้าน** อัตโนมัติ 3 ระดับ
- 🏢 **ดูข้อมูลบริษัท**: ภาคส่วน, อุตสาหกรรม, Market Cap, P/E, ปันผล
- 📈 **วิเคราะห์ทางเทคนิค**:
  - เส้นค่าเฉลี่ยเคลื่อนที่ MA 50/200 วัน
  - RSI (14 วัน) พร้อมระดับ Overbought/Oversold
  - ปริมาณ Volume
- 📊 **กราฟแบบโต้ตอบ** พร้อมเครื่องมือซูม/บันทึก
- 📝 **บันทึกกราฟเป็น PDF**
- ⭐ **Watchlist** จัดการหุ้นโปรด
- 🕘 **ประวัติการค้นหา** ล่าสุด 10 รายการ

---

## 💻 วิธีติดตั้ง

1. ติดตั้ง [Python 3.x](https://www.python.org/)
2. ติดตั้งไลบรารี:
   ```bash
   pip install tkinter yfinance pandas matplotlib
   ```
3. รันโปรแกรม:
   ```bash
   python stock_support_calculator.py
   ```

---

## 🧭 วิธีใช้งาน

1. กรอกรหัสหุ้น (เช่น `AAPL`, `MSFT`)
2. เลือกช่วงเวลา → คลิก “คำนวณ”
3. คลิก “ดูกราฟ” เพื่อดูภาพรวมและตัวชี้วัด
4. คลิก “บันทึก PDF” หากต้องการเก็บรายงาน
5. จัดการ **Watchlist** เพิ่ม/ลบหุ้นโปรดได้

---

## 🛠 ปัญหาที่พบบ่อย

| ปัญหา                      | วิธีแก้                           |
|---------------------------|------------------------------------|
| ❌ ไม่มีอินเทอร์เน็ต     | ตรวจสอบการเชื่อมต่อ              |
| ⚠️ รหัสหุ้นไม่ถูกต้อง    | ตรวจสอบว่ามีใน Yahoo Finance     |
| 📉 กราฟไม่แสดงผล         | ตรวจสอบ `matplotlib` ติดตั้งครบ  |

---

> พัฒนาเพื่อช่วยนักลงทุน วิเคราะห์ง่าย ใช้งานสะดวก  
> 💡 โดยใช้ Python และ Tkinter เป็นหลัก
