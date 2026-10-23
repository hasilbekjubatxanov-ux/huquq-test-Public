// Huquq testlari: natijalarni qabul qilish (doPost) va admin reytingi (doGet).
// Har bir test o'z kaliti bilan hisoblanadi va o'z varag'iga yoziladi.

// Admin sahifasiga kirish paroli — o'zingiz o'zgartiring.
const ADMIN_PAROL = "BU_YERGA_PAROL_YOZING";

const TESTLAR = {
  asosiy: { varaq: "Asosiy test", kalit: "DABCADBCADBCADBACDBCBDCADBCADB" },
  pro:    { varaq: "PRO test",    kalit: "BDCADBADCBDABDCADBCDCABDABDCBA" }
};
const BALL = [0.9, 1.5, 2.6]; // 1–10, 11–20, 21–30-savollar
const USTUNLAR = ["Sana", "Ism", "Ball", "To‘g‘ri", "Oson", "O‘rtacha", "Qiyin", "Xato savollar",
  "Javoblar", "Vaqt", "Soniya", "Vaqt tugadi", "Chiqishlar", "Telegram", "Telegram ID", "Qurilma"];

function testId_(id) {
  return Object.prototype.hasOwnProperty.call(TESTLAR, id) ? id : "asosiy";
}

function varaq_(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TESTLAR[id].varaq);
  if (!sh) {
    sh = ss.insertSheet(TESTLAR[id].varaq);
    sh.appendRow(USTUNLAR);
    sh.setFrozenRows(1);
  }
  return sh;
}

function hisobla_(id, javoblar) {
  const k = TESTLAR[id].kalit, b = [0, 0, 0], xato = [];
  for (let i = 0; i < k.length; i++) {
    const a = javoblar[i] || "-";
    if (a === k[i]) b[Math.floor(i / 10)]++;
    else xato.push((i + 1) + " (" + a + "→" + k[i] + ")");
  }
  const ball = Math.round((b[0] * BALL[0] + b[1] * BALL[1] + b[2] * BALL[2]) * 10) / 10;
  return { ball: ball, togri: b[0] + b[1] + b[2], oson: b[0], orta: b[1], qiyin: b[2], xato: xato };
}

// Jadvalda formula sifatida o'qilmasligi uchun
function matn_(v, max) {
  const s = String(v == null ? "" : v).slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const id = testId_(d.test);
  const javoblar = String(d.answers || "").toUpperCase().replace(/[^ABCD-]/g, "-").slice(0, 30);
  const h = hisobla_(id, javoblar);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    varaq_(id).appendRow([new Date(), matn_(d.name, 100), h.ball, h.togri, h.oson, h.orta, h.qiyin,
      h.xato.join(", ") || "yo‘q", javoblar, matn_(d.time, 10), Number(d.seconds) || 0, d.timeUp ? "ha" : "",
      Number(d.leaves) || 0, matn_(d.tg, 100), matn_(d.tgId, 30), matn_(d.device, 40)]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== ADMIN_PAROL) return json_({ error: "parol" });
  const id = testId_(p.test);
  const tz = Session.getScriptTimeZone();
  const qiymatlar = varaq_(id).getDataRange().getValues().slice(1);
  const korilgan = {};
  const rows = qiymatlar.map(function (r) {
    // Takroriy urinish: bir xil Telegram ID, qurilma yoki ism
    const kim = [String(r[14] || ""), String(r[15] || ""), String(r[1] || "").trim().toLowerCase()];
    const takror = kim.some(function (x) { return x && korilgan[x]; });
    kim.forEach(function (x) { if (x) korilgan[x] = true; });
    return {
      ism: String(r[1]).replace(/^'/, ""), ball: Number(r[2]) || 0, togri: r[3], oson: r[4], orta: r[5], qiyin: r[6],
      xato: String(r[7]), javoblar: String(r[8]), soniya: Number(r[10]) || 0, tugadi: r[11] === "ha", chiqish: r[12],
      sana: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "dd.MM.yyyy HH:mm") : String(r[0]), takror: takror
    };
  });
  rows.sort(function (a, b) { return b.ball - a.ball || a.soniya - b.soniya; });
  return json_({ rows: rows });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
