# บันทึก ME1 IPD ด้วยเสียง — วิธีติดตั้ง

ไฟล์: `Code.gs` (หลังบ้าน) · `app.html`, `manifest-app.json`, `icon-192.png`, `icon-512.png` (หน้าเว็บ)

## ขั้น 0 — ทดสอบการพูดก่อน (ยังไม่ต้องติดตั้งหลังบ้าน)
- อัปโหลด `app.html`, `manifest-app.json` ขึ้น repo `Mic-test` → เปิด `https://volgzangief134.github.io/Mic-test/app.html`
- กด **ลองแบบตัวอย่าง** → ใช้ฟอร์มจริงได้ทุกช่อง (รายชื่อยาครบ 456 คำ, รายชื่อ จนท. เป็นตัวอย่าง) แต่ไม่บันทึกจริง
- ถ้าต้องการวัด % ความแม่นยำ ใช้หน้าทดสอบ `index.html` v4 (ต้องอัปโหลดทับไฟล์เดิม ตอนนี้ใน repo ยังเป็น v2)

## ขั้น 1 — ให้ Apps Script สร้าง Google Sheet ให้
1. เปิด https://script.google.com → **New project** → ตั้งชื่อ เช่น `ME1 IPD Voice`
2. ลบโค้ดเดิม → วางโค้ดทั้งหมดจาก `Code.gs` → กด 💾 Save
3. ⚙️ **Project Settings → Script Properties → Add script property**
   `TEAM_CODE` = รหัสทีม (ตั้งเอง เช่น 6 หลัก) → Save script properties
4. กลับหน้า Editor → ช่องเลือกฟังก์ชันด้านบน เลือก **setupSpreadsheet** → **Run**
   - ครั้งแรกจะขอสิทธิ์ → Review permissions → เลือกบัญชี → Allow
   - ดู **Execution log**: จะขึ้น “สร้างชีตแล้ว: https://docs.google.com/…” → เปิดลิงก์นั้น
   - ชีตอยู่ใน Google Drive ของคุณ ชื่อ “แบบบันทึก ME1 IPD (แอปเสียง)”
5. (เช็ก) เลือกฟังก์ชัน **testInit** → Run → log ต้องมีรายการตัวเลือกออกมา

ชีตที่สร้าง:
| แท็บ | มาจาก |
|---|---|
| OK ME1 IPD | หัวตารางแถว 1–2 เหมือนไฟล์ต้นแบบ + dropdown คอลัมน์ B, D–M, O, P + สูตรช่วย S–AA |
| OK Admin | รายการตัวเลือกทุกคอลัมน์จากไฟล์ต้นแบบ (แก้/เพิ่มรายชื่อที่นี่) |
| App_Drugs | รายชื่อยาตั้งต้น 456 คำ จากบันทึก ก.ย. 69 |
| App_Glossary | คำศัพท์แก้คำ (เพิ่มคำที่ระบบได้ยินผิดได้เอง) |
| App_Log | ประวัติการบันทึกจากแอป + ข้อความจากเสียงก่อนแก้ |

ยังไม่ได้สร้างให้: แท็บสรุป (OK SumProcess, SumTE, SumDE) — ใช้สูตรจากไฟล์เดิมคัดลอกมาได้ หรือให้ผมทำให้ภายหลัง

## ขั้น 2 — เปิดใช้เป็น Web App
1. **Deploy → New deployment** → ⚙️ **Web app**
   - Execute as: **Me** · Who has access: **Anyone** (ป้องกันด้วยรหัสทีม)
2. **Deploy** → คัดลอก **Web app URL** (`https://script.google.com/macros/s/…/exec`)
3. เปิด `app.html` แก้บรรทัด `var API_URL = 'PASTE_YOUR_WEB_APP_URL_HERE';` → วางลิงก์ → อัปโหลดทับบน GitHub
4. มือถือ: เปิด `…/Mic-test/app.html` → ใส่รหัสทีม → เริ่มบันทึก

> แก้ Code.gs ภายหลัง: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy** (ลิงก์เดิมใช้ต่อได้)

## หมายเหตุ
- หน้าเว็บได้รับเฉพาะ “รหัส + ชื่อเล่น” (เช่น 03 แมว) · ชีตเก็บชื่อเต็มตามรูปแบบเดิม
- ถ้าสูตรคอลัมน์ S–AA ขึ้น error ให้แคปหน้าจอส่งมา (อาจเกี่ยวกับการตั้งค่าภาษาของชีต)
- เปลี่ยนรหัสทีม: แก้ `TEAM_CODE` ใน Script Properties แล้วแจ้งทีม
