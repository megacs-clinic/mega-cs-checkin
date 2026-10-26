/**
 * MEGA CLINIC — CS WFH Check-in Reminder (Office Script)
 * ------------------------------------------------------------
 * วางสคริปต์นี้ใน Excel ไฟล์ CS_Checkin_Log.xlsx (Automate > New Script)
 * แล้วให้ Power Automate เรียกด้วย action "Run script" ทุก 10 นาที
 *
 * mode = "remind"  (ค่าเริ่มต้น) → คืนรายการคนที่ต้องเตือน
 * mode = "summary"               → คืนสรุปการทำงานของ "เมื่อวาน" สำหรับโพสต์ในแชนแนล
 * mode = "test"                  → เหมือน remind แต่ไม่บันทึก ReminderLog (ไว้ทดสอบ)
 *
 * ชีตที่ใช้ (ถ้ายังไม่มี สคริปต์จะสร้างให้อัตโนมัติพร้อมหัวตาราง):
 *  - CheckinLog  : Date | Time | Name | Action | ActionLabel | ImageURL   (มีอยู่แล้ว)
 *  - Staff       : Name | FullName | Email | DefaultStart | DefaultEnd | WorkDays | Active
 *  - Schedule    : Date | Name | Start | End      (ตารางกะรายวัน ใช้แทน Default ได้ / ใส่ Start = OFF คือวันหยุด)
 *  - ReminderLog : SentAt | Key | Name | Type      (กันเตือนซ้ำ — ไม่ต้องแก้)
 */

// ───────────── ตั้งค่า ─────────────
const TZ_OFFSET_MIN = 7 * 60;          // เวลาไทย UTC+7
const PRE_SHIFT_MIN = 10;              // เตือนก่อนเข้ากะกี่นาที
const LATE_AFTER_MIN = 10;             // เลยเวลาเข้ากะกี่นาทีแล้วยังไม่เช็คอิน → เตือน + แจ้งแชนแนล
const BREAK_LIMIT_MIN = 60;            // พักได้กี่นาที
const BREAK_GRACE_MIN = 5;             // ผ่อนผันหลังครบเวลาพัก
const OUT_AFTER_MIN = 30;              // เลยเวลาเลิกกะกี่นาทีแล้วยังไม่กดเลิกงาน → เตือน
const NO_SCHEDULE_FORGOT_OUT_HR = 10;  // ไม่มีตารางกะ: เข้างานแล้วเกินกี่ชม.ยังไม่กดเลิก → เตือน
const CHECKIN_PAGE = "https://megacs-clinic.github.io/mega-cs-checkin/";

const THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
const THAI_DAYS = ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"];
const THAI_DAYS_FULL = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
const MIN = 60000;
const DAY = 86400000;

interface Reminder {
  name: string;
  fullName: string;
  email: string;
  type: string;      // pre | late | break | out
  key: string;       // ใช้กันเตือนซ้ำ
  message: string;   // HTML สำหรับส่ง DM
}

interface Result {
  mode: string;
  nowText: string;
  count: number;
  reminders: Reminder[];
  channelMessage: string;  // HTML สำหรับโพสต์ในแชนแนล ("" = ไม่ต้องโพสต์)
}

interface Ev { name: string; action: string; t: number; }
interface Staff { name: string; fullName: string; email: string; defStart: string; defEnd: string; workDays: string; active: boolean; }
interface Shift { name: string; start: number; end: number; }

function main(workbook: ExcelScript.Workbook, mode?: string): Result {
  const m = (mode || "remind").toLowerCase();
  // "เวลาไทย" แทนด้วย ms แบบ UTC wall-clock (ใช้ getUTC* อ่านค่า)
  const now = Date.now() + TZ_OFFSET_MIN * MIN;

  ensureSheet(workbook, "Staff", ["Name", "FullName", "Email", "DefaultStart", "DefaultEnd", "WorkDays", "Active"]);
  ensureSheet(workbook, "Schedule", ["Date", "Name", "Start", "End"]);
  const logSheet = ensureSheet(workbook, "ReminderLog", ["SentAt", "Key", "Name", "Type"]);

  const staff = readStaff(workbook);
  const events = readEvents(workbook);

  if (m === "summary") {
    return buildSummary(workbook, staff, events, now);
  }

  const sentKeys = readSentKeys(logSheet);
  const out: Reminder[] = [];
  const lateNames: string[] = [];
  const handled: { [k: string]: boolean } = {};

  const add = (st: Staff, type: string, key: string, msg: string) => {
    if (sentKeys[key]) return;
    sentKeys[key] = true;
    out.push({ name: st.name, fullName: st.fullName, email: st.email, type: type, key: key, message: msg });
  };

  // ── 1) เตือนตามตารางกะ (เมื่อวาน + วันนี้ เพื่อรองรับกะข้ามเที่ยงคืน)
  const today0 = Math.floor(now / DAY) * DAY;
  for (const day0 of [today0 - DAY, today0]) {
    for (const sh of shiftsForDay(workbook, staff, day0)) {
      const st = findStaff(staff, sh.name);
      if (!st || !st.email) continue;
      handled[sh.name] = true;
      const evs = events.filter(e => e.name === sh.name && e.t >= sh.start - 3 * 60 * MIN && e.t <= sh.end + 6 * 60 * MIN);
      const hasIn = evs.some(e => e.action === "in");
      const last = evs.length ? evs[evs.length - 1] : null;
      const sk = ymd(sh.start) + "@" + hm(sh.start);

      if (!hasIn && now >= sh.start - PRE_SHIFT_MIN * MIN && now < sh.start) {
        add(st, "pre", `pre|${st.name}|${sk}`,
          `⏰ สวัสดีค่ะคุณ${st.name} อีก ${Math.ceil((sh.start - now) / MIN)} นาทีจะถึงเวลาเข้ากะ <b>${hm(sh.start)}–${hm(sh.end)} น.</b><br>อย่าลืมกด <b>เข้างาน WFH</b> นะคะ 👉 <a href="${CHECKIN_PAGE}">เปิดหน้าเช็คอิน</a>`);
      }
      if (!hasIn && now >= sh.start + LATE_AFTER_MIN * MIN && now < sh.end) {
        lateNames.push(`${st.name} (กะ ${hm(sh.start)})`);
        add(st, "late", `late|${st.name}|${sk}`,
          `🔔 คุณ${st.name} ยังไม่ได้เช็คอินเข้างานค่ะ (กะเริ่ม ${hm(sh.start)} น. เลยมา ${Math.floor((now - sh.start) / MIN)} นาที)<br>ถ้าเริ่มงานแล้วรบกวนกด <b>เข้างาน WFH</b> 👉 <a href="${CHECKIN_PAGE}">เปิดหน้าเช็คอิน</a>`);
      }
      if (hasIn && last && last.action !== "out" && now >= sh.end + OUT_AFTER_MIN * MIN && now < sh.end + 6 * 60 * MIN) {
        add(st, "out", `out|${st.name}|${sk}`,
          `🌙 คุณ${st.name} เลยเวลาเลิกกะ (${hm(sh.end)} น.) แล้ว แต่ยังไม่ได้กด <b>เลิกงาน</b> ค่ะ<br>ถ้าเลิกงานแล้วอย่าลืมกดนะคะ 👉 <a href="${CHECKIN_PAGE}">เปิดหน้าเช็คอิน</a>`);
      }
    }
  }

  // ── 2) เตือนพักเกินเวลา + ลืมเลิกงาน (ไม่ต้องมีตารางกะ)
  for (const st of staff) {
    if (!st.active || !st.email) continue;
    const evs = events.filter(e => e.name === st.name && e.t <= now && e.t >= now - 16 * 60 * MIN);
    if (!evs.length) continue;
    const last = evs[evs.length - 1];
    if (last.action === "break" && now - last.t >= (BREAK_LIMIT_MIN + BREAK_GRACE_MIN) * MIN && now - last.t < 4 * 60 * MIN) {
      add(st, "break", `break|${st.name}|${ymd(last.t)}@${hm(last.t)}`,
        `☕ คุณ${st.name} พักมาแล้ว ${Math.floor((now - last.t) / MIN)} นาที (เริ่มพัก ${hm(last.t)} น.)<br>ถ้ากลับมาทำงานแล้ว อย่าลืมกด <b>กลับจากพัก</b> นะคะ 👉 <a href="${CHECKIN_PAGE}">เปิดหน้าเช็คอิน</a>`);
    }
    if (!handled[st.name] && (last.action === "in" || last.action === "back")) {
      const ins = evs.filter(e => e.action === "in");
      const firstIn = ins.length ? ins[ins.length - 1] : null;
      if (firstIn && now - firstIn.t >= NO_SCHEDULE_FORGOT_OUT_HR * 60 * MIN) {
        add(st, "out", `out|${st.name}|${ymd(firstIn.t)}@${hm(firstIn.t)}`,
          `🌙 คุณ${st.name} เช็คอินตั้งแต่ ${hm(firstIn.t)} น. และยังไม่ได้กด <b>เลิกงาน</b> ค่ะ<br>ถ้าเลิกงานแล้วอย่าลืมกดนะคะ 👉 <a href="${CHECKIN_PAGE}">เปิดหน้าเช็คอิน</a>`);
      }
    }
  }

  // ── 3) ข้อความรวมในแชนแนล (เฉพาะคนที่สายและยังไม่ได้แจ้งแชนแนล)
  const newLate = out.filter(r => r.type === "late").map(r => r.name);
  const lateLines = lateNames.filter(n => newLate.indexOf(n.split(" (")[0]) >= 0);
  const channelMessage = lateLines.length
    ? `🔔 <b>ยังไม่ได้เช็คอินเข้างาน</b> (${hm(now)} น.)<br>${lateLines.map(n => "• " + n).join("<br>")}<br>ส่งข้อความเตือนส่วนตัวให้แล้วค่ะ`
    : "";

  if (m !== "test" && out.length) {
    const rows = out.map(r => [fmt(now), r.key, r.name, r.type]);
    appendRows(logSheet, rows);
  }

  return { mode: m, nowText: fmt(now), count: out.length, reminders: out, channelMessage: channelMessage };
}

// ───────────── สรุปรายวัน ─────────────
function buildSummary(workbook: ExcelScript.Workbook, staff: Staff[], events: Ev[], now: number): Result {
  // วันทำงาน = 05:00 ถึง 05:00 ของวันถัดไป (ตรงกับหน้าเช็คอิน)
  const day0 = Math.floor((now - 5 * 60 * MIN) / DAY) * DAY - DAY;
  const from = day0 + 5 * 60 * MIN, to = from + DAY;
  const shifts = shiftsForDay(workbook, staff, day0);
  const lines: string[] = [];
  const names: string[] = [];
  staff.filter(s => s.active).forEach(s => names.push(s.name));
  events.filter(e => e.t >= from && e.t < to).forEach(e => { if (names.indexOf(e.name) < 0) names.push(e.name); });

  for (const n of names) {
    const evs = events.filter(e => e.name === n && e.t >= from && e.t < to);
    const sh = shifts.filter(s => s.name === n)[0];
    if (!evs.length) {
      if (sh) lines.push(`• <b>${n}</b> — ❌ ไม่มีการเช็คอิน (กะ ${hm(sh.start)}–${hm(sh.end)})`);
      continue;
    }
    const inE = evs.filter(e => e.action === "in")[0];
    const outs = evs.filter(e => e.action === "out");
    const outE = outs.length ? outs[outs.length - 1] : null;
    let breakMin = 0;
    let open = 0;
    for (const e of evs) {
      if (e.action === "break") open = e.t;
      if (e.action === "back" && open) { breakMin += (e.t - open) / MIN; open = 0; }
    }
    const notes: string[] = [];
    if (sh && inE && inE.t > sh.start + 5 * MIN) notes.push(`สาย ${Math.round((inE.t - sh.start) / MIN)} นาที`);
    if (breakMin > BREAK_LIMIT_MIN) notes.push(`พักเกิน ${Math.round(breakMin - BREAK_LIMIT_MIN)} นาที`);
    if (open) notes.push("ไม่ได้กดกลับจากพัก");
    if (!outE) notes.push("ไม่ได้กดเลิกงาน");
    const work = inE && outE ? ` · ทำงาน ${hours((outE.t - inE.t) / MIN - breakMin)}` : "";
    lines.push(`• <b>${n}</b> — เข้า ${inE ? hm(inE.t) : "-"} · ออก ${outE ? hm(outE.t) : "-"} · พัก ${Math.round(breakMin)} น.${work}` +
      (notes.length ? ` ⚠️ ${notes.join(", ")}` : " ✅"));
  }
  const d = new Date(day0);
  const title = `📋 <b>สรุปเช็คอิน CS WFH — วัน${THAI_DAYS_FULL[d.getUTCDay()]}ที่ ${d.getUTCDate()} ${THAI_MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear() + 543}</b>`;
  return {
    mode: "summary", nowText: fmt(now), count: lines.length, reminders: [],
    channelMessage: lines.length ? `${title}<br>${lines.join("<br>")}` : `${title}<br>ไม่มีรายการ`
  };
}

// ───────────── อ่านข้อมูล ─────────────
function readEvents(workbook: ExcelScript.Workbook): Ev[] {
  const ws = workbook.getWorksheet("CheckinLog");
  if (!ws) return [];
  const rng = ws.getUsedRange();
  if (!rng) return [];
  const v = rng.getValues();
  const h = v[0].map(x => String(x).trim().toLowerCase());
  const iD = h.indexOf("date"), iT = h.indexOf("time"), iN = h.indexOf("name"), iA = h.indexOf("action");
  const out: Ev[] = [];
  for (let r = 1; r < v.length; r++) {
    const d = parseDate(v[r][iD]);
    const t = parseTime(v[r][iT]);
    const n = String(v[r][iN]).trim();
    const a = String(v[r][iA]).trim().toLowerCase();
    if (d === null || t === null || !n || !a) continue;
    out.push({ name: n, action: a, t: d + t });
  }
  out.sort((a, b) => a.t - b.t);
  return out;
}

function readStaff(workbook: ExcelScript.Workbook): Staff[] {
  const v = workbook.getWorksheet("Staff").getUsedRange().getValues();
  const out: Staff[] = [];
  for (let r = 1; r < v.length; r++) {
    const name = String(v[r][0]).trim();
    if (!name) continue;
    const act = String(v[r][6]).trim().toLowerCase();
    out.push({
      name: name, fullName: String(v[r][1]).trim(), email: String(v[r][2]).trim(),
      defStart: cellToTimeText(v[r][3]), defEnd: cellToTimeText(v[r][4]),
      workDays: String(v[r][5]).trim(), active: !(act === "no" || act === "n" || act === "false" || act === "0" || act === "ไม่")
    });
  }
  return out;
}

function shiftsForDay(workbook: ExcelScript.Workbook, staff: Staff[], day0: number): Shift[] {
  const v = workbook.getWorksheet("Schedule").getUsedRange().getValues();
  const out: Shift[] = [];
  const fromSchedule: { [n: string]: boolean } = {};
  for (let r = 1; r < v.length; r++) {
    const d = parseDate(v[r][0]);
    const n = String(v[r][1]).trim();
    if (d !== day0 || !n) continue;
    fromSchedule[n] = true;  // มีบรรทัดในตาราง = ใช้ตารางแทน Default (รวมถึง OFF)
    const s = parseTime(v[r][2]), e = parseTime(v[r][3]);
    if (s === null || e === null) continue; // OFF / ว่าง = วันหยุด
    out.push({ name: n, start: d + s, end: d + e + (e <= s ? DAY : 0) });
  }
  const dow = THAI_DAYS[new Date(day0).getUTCDay()];
  for (const st of staff) {
    if (!st.active || fromSchedule[st.name] || !st.defStart || !st.defEnd) continue;
    const days = st.workDays.split(/[,\s]+/).filter(x => x);
    if (days.length && days.indexOf(dow) < 0) continue;
    const s = parseTime(st.defStart), e = parseTime(st.defEnd);
    if (s === null || e === null) continue;
    out.push({ name: st.name, start: day0 + s, end: day0 + e + (e <= s ? DAY : 0) });
  }
  return out;
}

function readSentKeys(ws: ExcelScript.Worksheet): { [k: string]: boolean } {
  const v = ws.getUsedRange().getValues();
  const keys: { [k: string]: boolean } = {};
  for (let r = 1; r < v.length; r++) keys[String(v[r][1])] = true;
  return keys;
}

// ───────────── helpers ─────────────
function ensureSheet(workbook: ExcelScript.Workbook, name: string, headers: string[]): ExcelScript.Worksheet {
  let ws = workbook.getWorksheet(name);
  if (!ws) {
    ws = workbook.addWorksheet(name);
    const hr = ws.getRangeByIndexes(0, 0, 1, headers.length);
    hr.setValues([headers]);
    hr.getFormat().getFont().setBold(true);
    hr.getFormat().getFill().setColor("#1B2A4A");
    hr.getFormat().getFont().setColor("#FFFFFF");
  }
  return ws;
}

function appendRows(ws: ExcelScript.Worksheet, rows: (string | number)[][]) {
  const used = ws.getUsedRange();
  const start = used ? used.getRowIndex() + used.getRowCount() : 1;
  ws.getRangeByIndexes(start, 0, rows.length, rows[0].length).setValues(rows);
}

function findStaff(staff: Staff[], name: string): Staff | null {
  const s = staff.filter(x => x.name === name && x.active);
  return s.length ? s[0] : null;
}

/** แปลงวันที่ → ms (เที่ยงคืนของวันนั้น, wall-clock) รองรับ "13 ต.ค. 2569", serial ของ Excel, "2026-10-13", "13/10/2569" */
function parseDate(x: string | number | boolean): number | null {
  if (typeof x === "number") return x > 20000 ? Math.round((Math.floor(x) - 25569) * DAY) : null;
  const s = String(x).trim();
  if (!s) return null;
  const parts = s.split(/\s+/);
  if (parts.length === 3) {
    const mi = THAI_MONTHS.indexOf(parts[1]);
    if (mi >= 0) return Date.UTC(fixYear(Number(parts[2])), mi, Number(parts[0]));
  }
  let mm = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (mm) return Date.UTC(fixYear(Number(mm[1])), Number(mm[2]) - 1, Number(mm[3]));
  mm = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mm) return Date.UTC(fixYear(Number(mm[3])), Number(mm[2]) - 1, Number(mm[1]));
  return null;
}
function fixYear(y: number): number { return y > 2400 ? y - 543 : y; }

/** แปลงเวลา → ms นับจากเที่ยงคืน รองรับเลขเศษส่วนของ Excel และข้อความ "14:57" / "9.30" */
function parseTime(x: string | number | boolean): number | null {
  if (typeof x === "number") return Math.round((x % 1) * 1440) * MIN;
  const mm = String(x).trim().match(/^(\d{1,2})[:.](\d{2})/);
  if (!mm) return null;
  return (Number(mm[1]) * 60 + Number(mm[2])) * MIN;
}
function cellToTimeText(x: string | number | boolean): string {
  const t = parseTime(x);
  return t === null ? "" : hm(t);
}

function p2(n: number): string { return (n < 10 ? "0" : "") + n; }
function hm(t: number): string { const d = new Date(t); return p2(d.getUTCHours()) + ":" + p2(d.getUTCMinutes()); }
function ymd(t: number): string { const d = new Date(t); return d.getUTCFullYear() + "-" + p2(d.getUTCMonth() + 1) + "-" + p2(d.getUTCDate()); }
function fmt(t: number): string { return ymd(t) + " " + hm(t); }
function hours(min: number): string { return Math.floor(min / 60) + " ชม. " + Math.round(min % 60) + " น."; }
