# CS WFH Check-in v2 — คู่มือติดตั้ง

ระบบใหม่ใช้แค่ Microsoft 365 ที่บริษัทมีอยู่แล้ว ไม่ต้องใช้ Supabase หรือบริการภายนอกอื่น

```
หน้าเช็คอิน (GitHub Pages)
   │  ส่ง ชื่อ/เวลา/รูป (base64)
   ▼
Workflow ①  "รับเช็คอิน"  (สร้างใหม่ด้วยบัญชีที่ยังใช้งาน)
   ├─ เซฟรูปลง SharePoint  →  CSTeam / Documents / Checkin-Photos
   ├─ เพิ่มแถวใน CS_Checkin_Log.xlsx (ImageURL = ลิงก์ SharePoint)
   └─ โพสต์การ์ดใน Teams (เหมือนเดิม)

Flow ②  "เตือนเช็คอิน"  (ใหม่ — รันทุก 10 นาที)
   ├─ Run script "CheckinReminder" ใน Excel → หาคนที่ต้องเตือน
   ├─ ส่ง DM ส่วนตัวผ่าน Flow bot (ก่อนเข้ากะ / สาย / พักเกิน / ลืมเลิกงาน)
   └─ โพสต์รายชื่อคนที่ยังไม่เช็คอินในแชนแนล

Flow ③  "สรุปรายวัน"  (ใหม่ — รันวันละครั้ง)
   └─ โพสต์สรุป เข้า/ออก/พัก/สาย ของเมื่อวานในแชนแนล
```

ทุก action ที่ใช้ (Excel Online Business, SharePoint, Microsoft Teams, Recurrence) เป็น **standard connector** ใช้ได้ด้วย license M365 ปกติ ไม่ต้องซื้อ Premium

---

## ขั้นที่ 0 — ใช้บัญชีที่ไม่ผูกกับคนใดคนหนึ่ง

ระบบเดิมผูกกับอีเมลพนักงานที่ถูกปิดไปแล้ว จึงแก้ทั้ง GitHub และ Workflow ไม่ได้ และ Workflow เดิมอาจหยุดทำงานเมื่อไรก็ได้ (connection ของเจ้าของที่ถูกปิดจะหมดอายุ) ครั้งนี้จึงสร้างใหม่ทั้งหมดโดย:

- สร้าง Flow ด้วยบัญชีของคุณเอง แล้ว **เพิ่ม co-owner** อย่างน้อย 1 คน (Flow → Share → เพิ่มหัวหน้าทีมหรือ M365 Group "CS Team") ถ้าใครลาออก อีกคนยังแก้ต่อได้
- GitHub ให้สมัครด้วยอีเมลกลางของทีม (เช่น shared mailbox ของ CS หรือ IT) ไม่ใช่อีเมลส่วนตัวพนักงาน
- ไฟล์ Excel อยู่ใน SharePoint CSTeam อยู่แล้ว ใช้ต่อได้เลย

## ขั้นที่ 1 — สร้าง Workflow ① "รับเช็คอิน" ใหม่

make.powerautomate.com → **Create → Instant cloud flow → Skip** แล้วเพิ่ม action ตามนี้

1. **Trigger: Microsoft Teams › When a Teams webhook request is received**
   - Who can trigger the flow: **Anyone** (หน้าเช็คอินไม่มีการล็อกอิน)
2. **Initialize variable** · Name `PhotoUrl` · Type String · Value ว่าง
3. **Condition**: `empty(triggerBody()?['imageBase64'])` **is equal to** `false`
   - **True** ใส่ 3 action:
     1. **SharePoint › Create file**
        - Site Address: `https://megaclinicthailand1.sharepoint.com/sites/CSTeam`
        - Folder Path: `/Shared Documents/Checkin-Photos` (สร้างโฟลเดอร์นี้ใน Documents ก่อน)
        - File Name: `triggerBody()?['fileName']`
        - File Content: `base64ToBinary(triggerBody()?['imageBase64'])`
     2. **SharePoint › Create sharing link for a file or folder**
        - Site Address เหมือนข้างบน · Library Name: Documents
        - Item Id: `outputs('Create_file')?['body/ItemId']`
        - Link Type: **View** · Link Scope: **Organization**
     3. **Set variable** `PhotoUrl` = `outputs('Create_sharing_link_for_a_file_or_folder')?['body/link/webUrl']`
4. **Excel Online (Business) › Add a row into a table**
   - Location: CSTeam · Library: Documents · File: `CS_Checkin_Log.xlsx` · Table: ตารางในชีต CheckinLog
     (ถ้าชีตยังไม่เป็น Table ให้เลือกช่วง A1:F → Insert → Table → ✔ My table has headers ก่อน)
   - Date `triggerBody()?['date']` · Time `triggerBody()?['time']` · Name `triggerBody()?['name']`
   - Action `triggerBody()?['action']` · ActionLabel `triggerBody()?['actionLabel']` · ImageURL `variables('PhotoUrl')`
5. **Microsoft Teams › Post card in a chat or channel**
   - Post as: Flow bot · Post in: Channel · Team: CS Team · Channel: แชนแนลเช็คอิน
   - Adaptive Card:

```json
{
  "type": "AdaptiveCard",
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.4",
  "body": [
    { "type": "TextBlock", "size": "Medium", "weight": "Bolder",
      "text": "@{triggerBody()?['actionLabel']} — @{triggerBody()?['name']}" },
    { "type": "FactSet", "facts": [
      { "title": "วันที่", "value": "@{triggerBody()?['date']}" },
      { "title": "เวลา", "value": "@{triggerBody()?['time']} น." }
    ]}
  ],
  "actions": [
    { "type": "Action.OpenUrl", "title": "📷 ดูรูปยืนยัน", "url": "@{if(empty(variables('PhotoUrl')), 'https://megaclinicthailand1.sharepoint.com/sites/CSTeam', variables('PhotoUrl'))}" }
  ]
}
```

6. **Save** → กลับไปเปิด trigger แล้ว **copy HTTP URL** ที่ได้ไว้ใช้ในขั้นที่ 2

> การ์ดแสดงรูปจาก SharePoint ตรง ๆ ไม่ได้ (ต้องล็อกอิน) จึงใช้ปุ่ม "ดูรูปยืนยัน" แทน

## ขั้นที่ 2 — เปิดหน้าเช็คอินใหม่บน GitHub Pages

1. สมัคร/ล็อกอิน GitHub ด้วยอีเมลกลางของทีม → **New repository** ชื่อ `mega-cs-checkin` · Public
2. เปิด `index.html` ไฟล์ใหม่ แก้บรรทัด `const WEBHOOK_URL = 'PASTE_YOUR_POWER_AUTOMATE_URL_HERE'` เป็น URL จากขั้นที่ 1
3. **Add file → Upload files** → อัปโหลด `index.html` → Commit
4. **Settings → Pages** → Source: Deploy from a branch · Branch: `main` / root → Save
5. รอ 1–2 นาที จะได้ลิงก์ `https://<ชื่อบัญชี>.github.io/mega-cs-checkin/` → แจ้งทีมให้ใช้ลิงก์ใหม่ (และปักหมุดเป็นแท็บใน Teams ได้)
6. ทดสอบเช็คอิน 1 ครั้ง: การ์ดเข้า Teams · แถวเข้า Excel · รูปเข้าโฟลเดอร์ Checkin-Photos

> หน้าเดิม (piyaorn-beep.github.io) ยังเปิดได้และยังยิงไป Workflow เดิม — แจ้งทีมให้เลิกใช้ลิงก์เก่า เพื่อไม่ให้ข้อมูลแยกสองที่

## ขั้นที่ 3 — ติดตั้ง Office Script + กรอกรายชื่อ/กะ

1. เปิด `CS_Checkin_Log.xlsx` บนเว็บ → แท็บ **Automate** → **New Script**
2. ลบโค้ดตัวอย่าง วางโค้ดจาก `office-scripts/CheckinReminder.ts` → ตั้งชื่อ `CheckinReminder` → Save
3. กด **Run** 1 ครั้ง สคริปต์จะสร้างชีต `Staff`, `Schedule`, `ReminderLog` ให้อัตโนมัติ
4. กรอกชีต **Staff** (1 แถวต่อ 1 คน — Name ต้องตรงกับชื่อในหน้าเช็คอิน)

| Name | FullName | Email | DefaultStart | DefaultEnd | WorkDays | Active |
|---|---|---|---|---|---|---|
| มะมาย | ศิริลักษณ์ | xxx@megaclinicthailand.com | 06:00 | 15:00 | จ,อ,พ,พฤ,ศ | |
| ปูเป้ | ณัฐธิดา | xxx@megaclinicthailand.com | 15:00 | 00:00 | | |
| ไนซ์ | ณัฐนิชา | xxx@megaclinicthailand.com | | | | |

   - DefaultStart/End = กะประจำ (ถ้ามี) · กะข้ามเที่ยงคืนใส่ได้เลย เช่น 17:00–02:00
   - WorkDays ว่าง = ทุกวัน · ใส่ `อา จ อ พ พฤ ศ ส` คั่นด้วย , ได้
   - Active ใส่ `no` = ปิดการเตือนคนนี้ (เช่นลาออก/ลาพัก)

5. ชีต **Schedule** (ไม่บังคับ — สำหรับคนที่กะหมุนเวียน) ใส่ 1 แถว/คน/วัน — ถ้ามีแถวในวันนั้นจะใช้แทน Default

| Date | Name | Start | End |
|---|---|---|---|
| 14/10/2026 | ไนซ์ | 10:00 | 19:00 |
| 14/10/2026 | ปูเป้ | OFF | |

6. ทดสอบ: Automate → CheckinReminder → Run แล้วดูผลใน Output (ระหว่างทดสอบใน Power Automate ใช้ mode = `test` จะไม่บันทึก ReminderLog)

**การเตือนที่สคริปต์ทำ** (ปรับตัวเลขได้ที่หัวไฟล์สคริปต์)

| ประเภท | เงื่อนไข | ส่งถึง |
|---|---|---|
| ⏰ pre | 10 นาทีก่อนเข้ากะ ยังไม่เช็คอิน | DM |
| 🔔 late | เลยเวลาเข้ากะ 10 นาที ยังไม่เช็คอิน | DM + แชนแนล |
| ☕ break | กดพักแล้วเกิน 65 นาที ยังไม่กดกลับ | DM |
| 🌙 out | เลยเวลาเลิกกะ 30 นาที ยังไม่กดเลิกงาน (ไม่มีกะ: เข้างานเกิน 10 ชม.) | DM |

แต่ละเรื่องเตือน **ครั้งเดียว** ต่อกะ (จำไว้ในชีต ReminderLog)

## ขั้นที่ 4 — สร้าง Flow ② "เตือนเช็คอิน"

make.powerautomate.com → **Create → Scheduled cloud flow**

1. **Recurrence**: Interval `10` Minute · Time zone `(UTC+07:00) Bangkok, Hanoi, Jakarta`
2. **Excel Online (Business) › Run script**
   - Location: SharePoint Site – CSTeam · Document Library: Documents · File: `CS_Checkin_Log.xlsx`
   - Script: `CheckinReminder` · mode: `remind`
3. **Apply to each** → Select output: `outputs('Run_script')?['body/result/reminders']`
   - ข้างใน: **Microsoft Teams › Post message in a chat or channel**
     - Post as: **Flow bot** · Post in: **Chat with Flow bot**
     - Recipient: `items('Apply_to_each')?['email']`
     - Message: `items('Apply_to_each')?['message']`
4. **Condition**: `outputs('Run_script')?['body/result/channelMessage']` **is not equal to** (ว่าง)
   - **True**: **Microsoft Teams › Post message in a chat or channel**
     - Post as: Flow bot · Post in: Channel · Team: CS Team · Channel: (แชนแนลเช็คอิน)
     - Message: `outputs('Run_script')?['body/result/channelMessage']`
5. Save → กด **Test → Manually** ดูผล

> ครั้งแรกที่ Flow bot ทักใครใน Teams คนนั้นจะเห็นแชต "Workflows" / "Flow" เด้งขึ้นมา — แจ้งทีมไว้ก่อนว่าเป็นบอทเตือนเช็คอิน

## ขั้นที่ 5 — Flow ③ "สรุปรายวัน" (แนะนำ)

Scheduled cloud flow → Recurrence ทุก 1 Day เวลา 09:00 (Bangkok)
→ **Run script** (ไฟล์เดิม, mode = `summary`)
→ **Post message in a chat or channel** (Flow bot → Channel) · Message = `outputs('Run_script')?['body/result/channelMessage']`

ตัวอย่างข้อความ:

```
📋 สรุปเช็คอิน CS WFH — วันจันทร์ที่ 12 ต.ค. 2569
• มะมาย — เข้า 05:57 · ออก 15:01 · พัก 60 น. · ทำงาน 8 ชม. 4 น. ✅
• ปูเป้ — เข้า 15:12 · ออก - · พัก 75 น. ⚠️ สาย 12 นาที, พักเกิน 15 นาที, ไม่ได้กดเลิกงาน
• ไนซ์ — ❌ ไม่มีการเช็คอิน (กะ 10:00–19:00)
```

(วันทำงานนับ 05:00–05:00 ของวันถัดไป เหมือนหน้าเช็คอิน กะดึกจึงนับรวมวันเดียวกัน)

---

## แก้ปัญหา

| อาการ | สาเหตุ / วิธีแก้ |
|---|---|
| ขึ้น "บันทึกแล้ว (ยังไม่มี webhook)" | ยังไม่ได้วาง URL ใน `WEBHOOK_URL` |
| ขึ้น `Teams error: HTTP 401/403` | Who can trigger ไม่ใช่ Anyone |
| ขึ้น `Teams error: Failed to fetch` | URL ผิด หรือ Flow ถูกปิด (Turn on) |
| ส่งได้แต่รูปไม่เข้า SharePoint | ดู run history ของ Workflow ① ที่ action Create file — มักเป็น Folder Path ผิด |
| ไม่มีใครได้ DM | Email ในชีต Staff ไม่ตรงบัญชี Teams หรือชื่อ (Name) ไม่ตรงกับในหน้าเช็คอิน |
| เตือนผิดเวลา | ตรวจ Time zone ใน Recurrence และค่า Start/End ในชีต (ต้องเป็นเวลา ไม่ใช่ข้อความแปลก ๆ) |
| อยากเตือนซ้ำใหม่เพื่อทดสอบ | ลบแถวในชีต ReminderLog |

## ความปลอดภัย (ควรทำ)

- URL ของ Workflow ① จะอยู่ในโค้ดสาธารณะบน GitHub ใครได้ลิงก์ก็ยิงข้อมูลเข้า Teams ได้ — ถ้าเริ่มมีสแปม ให้สร้าง URL ใหม่ใน Workflow แล้วอัปเดตใน `index.html`
- รูปเก่าบน Supabase ตอนนี้ยังเปิดดูได้ผ่านลิงก์ใน Excel แต่โปรเจกต์อาจถูกปิดในอนาคต ถ้าต้องเก็บไว้เป็นหลักฐาน ควรดาวน์โหลดเก็บลง SharePoint
