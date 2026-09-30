/* ===================== ตั้งค่า Firebase =====================
   1) สร้างโปรเจกต์ที่ https://console.firebase.google.com  → เพิ่ม Web app (</>)
   2) คัดลอกค่า firebaseConfig มาวางแทนที่ null ด้านล่าง
   3) เปิด Authentication → Google  และสร้าง Firestore Database
   4) นำเนื้อหาไฟล์ firestore.rules ไปวางที่ Firestore → Rules → Publish
   (ดูขั้นตอนละเอียดใน README.md)

   ตราบใดที่ firebaseConfig ยังเป็น null ระบบจะทำงานใน "โหมดทดลอง" (เก็บข้อมูลในเครื่องเท่านั้น)
   หมายเหตุ: ค่า apiKey ของ Firebase ไม่ใช่ความลับ ความปลอดภัยอยู่ที่ Security Rules */
/* ตัวอย่าง:
export const firebaseConfig = {
  apiKey: "AIzaSyBAO50f0i4xPnrcmF6A-Xumk05tH9-xlEo",
  authDomain: "virach-timesheet.firebaseapp.com",
  projectId: "virach-timesheet",
  storageBucket: "virach-timesheet.firebasestorage.app",
  messagingSenderId: "867034215000",
  appId: "1:867034215000:web:22d719be828446394c102d",
  measurementId: "G-826YP7LE94"
};
*/

/* อีเมลเจ้าของระบบ (แอดมินคนแรก) — ต้องตรงกับที่ระบุใน firestore.rules ด้วย */
export const ADMIN_EMAILS = ['sawat432@gmail.com'];

/* จำกัดให้ล็อกอินได้เฉพาะโดเมนอีเมลบริษัท เช่น ''  (เว้นว่าง '' = อนุญาตทุกบัญชี Google)
   ⚠ เป็นการตรวจฝั่งหน้าเว็บ ถ้าต้องการบังคับจริงให้เปิดบรรทัด domain ใน firestore.rules ด้วย */
export const ALLOWED_EMAIL_DOMAIN = '';
