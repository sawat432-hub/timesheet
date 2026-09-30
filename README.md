# Monthly Time Sheet (เว็บ + Google Login + Firebase)

เว็บบันทึกเวลาทำงานรายเดือนของ Dr.Virach & Associates  
- พนักงานล็อกอินด้วย **Google** · ข้อมูลเก็บใน **Firebase Firestore** · โฮสต์บน **GitHub Pages**
- ปุ่ม “ดาวน์โหลด Excel” สร้างไฟล์จาก `template.xlsx` (ฟอร์ม TIMESHEET เดิม) — หน้าตา สูตร วันหยุด เหมือนเดิมทุกอย่าง
- แท็บ “รายงานวิเคราะห์”: รายจ็อบ / รายปี / รายพนักงาน / รายเดือน / Period / วันลา / O.T. (ส่งออก Excel ได้)
- แอดมินอนุมัติรายงาน ดูและดาวน์โหลดของทุกคน (ZIP) และตั้งค่ารายชื่อลูกค้า/วันหยุดกลาง

## 1) ตั้งค่า Firebase (ครั้งเดียว ~10 นาที)
1. ไปที่ https://console.firebase.google.com → **Add project**
2. **Build → Authentication → Get started → Sign-in method → Google → Enable** (ใส่อีเมลสนับสนุนโปรเจกต์)
3. **Build → Firestore Database → Create database** (เลือก production mode, region `asia-southeast1` สิงคโปร์)
4. แท็บ **Rules** → วางเนื้อหาไฟล์ `firestore.rules` ทั้งหมด → **Publish**  
   (แก้อีเมลเจ้าของระบบในฟังก์ชัน `isOwner()` ให้ตรงกับของคุณ)
5. **Project settings (⚙) → Your apps → Web (</>)** → Register app → คัดลอก `firebaseConfig`
6. เปิดไฟล์ `firebase-config.js` วางค่าแทน `null` และแก้ `ADMIN_EMAILS` (อีเมลแอดมินคนแรก)

## 2) ขึ้นเว็บด้วย GitHub Pages
1. สร้าง repository ใหม่บน GitHub (private/public ก็ได้) แล้วอัปโหลดไฟล์ทั้งโฟลเดอร์นี้ (branch `main`)
2. **Settings → Pages → Build and deployment → Source: GitHub Actions**  
   (ไฟล์ `.github/workflows/pages.yml` จะ deploy ให้อัตโนมัติทุกครั้งที่ push)
3. รอ Actions เสร็จ จะได้ลิงก์ `https://<ชื่อผู้ใช้>.github.io/<ชื่อ repo>/`
4. **กลับไป Firebase → Authentication → Settings → Authorized domains → Add domain**  
   ใส่ `<ชื่อผู้ใช้>.github.io` (ถ้าไม่ใส่ ปุ่ม Google Login จะขึ้น `auth/unauthorized-domain`)

## 3) การใช้งาน
- คนแรกที่ล็อกอินด้วยอีเมลใน `ADMIN_EMAILS` จะเป็น **แอดมิน** โดยอัตโนมัติ
- พนักงานคนอื่นล็อกอินครั้งแรกจะเป็น **พนักงาน** — แอดมินเปลี่ยนสิทธิ์ได้ที่แท็บ “พนักงาน / อนุมัติ”
- แท็บ “ตั้งค่า / วันหยุด” (แอดมิน): รายชื่อลูกค้า/จ็อบ, Period, วันหยุด, ชั่วโมงมาตรฐาน — มีผลกับทุกคนและไฟล์ Excel
- ขั้นตอนสิ้นเดือน: พนักงานกด “ส่งรายงาน” → แอดมินตรวจและกด “อนุมัติ” (ล็อกไม่ให้แก้) → ดาวน์โหลด Excel
- ยังไม่ใส่ `firebase-config.js` = **โหมดทดลอง** (ข้อมูลอยู่ในเบราว์เซอร์เครื่องนั้นเท่านั้น) ใช้ลองระบบได้

## ทดสอบในเครื่อง
ต้องเปิดผ่านเซิร์ฟเวอร์ ไม่ใช่ดับเบิลคลิกไฟล์ (เพราะใช้ ES modules):  
`python3 -m http.server 8000` แล้วเปิด http://localhost:8000

## ข้อจำกัดของฟอร์ม Excel เดิม
- รองรับ **32 แถว/เดือน** (สูตร SUMIFS ใน TIME SHEET อ่านช่วงแถว 6–37) ระบบจะตัดแถววันว่างออกก่อน ถ้ายังเกินจะเตือน
- ไฟล์ที่ได้ให้ Excel คำนวณเองเมื่อเปิด — ถ้าเปิดใน Protected View ให้กด **Enable Editing** ตัวเลขจึงขึ้น
- แผ่น “TIME REPORT O.T.” เป็นสูตรเดิมที่คิด O.T. ทีละแถว ส่วนหน้าเว็บ/สรุป O.T. คิดสะสมทั้งวัน (ตรงกับแถว OVER TIME ใน TIME SHEET)
