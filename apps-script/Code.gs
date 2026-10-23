// Huquq testlari: natijalarni qabul qilish, admin reytingi va pullik test kodlari.
// Har bir test o'z kaliti bilan hisoblanadi va o'z varag'iga yoziladi.
//
// Pullik (PRO) test savollari saytda emas, shu skriptda saqlanadi: o'quvchi to'g'ri
// bir martalik kod kiritgandagina savollarni oladi, tahlil esa test tugagach beriladi.
// Repodagi bu faylda savollar yo'q (PRO_DATA = null); savollar qo'shilgan to'liq fayl
// faqat Google skriptga qo'yiladi va GitHub'ga yuklanmaydi.

// Admin sahifasiga kirish paroli — o'zingiz o'zgartiring.
const ADMIN_PAROL = "BU_YERGA_PAROL_YOZING";

const PRO_DATA = null; // @@PRO_DATA@@

const TESTLAR = {
  asosiy: { varaq: "Asosiy test", kalit: "DABCADBCADBCADBACDBCBDCADBCADB" },
  pro:    { varaq: "PRO test", kalit: PRO_DATA ? PRO_DATA.kalit : "", pullik: true }
};
const BALL = [0.9, 1.5, 2.6]; // 1–10, 11–20, 21–30-savollar
const LIMIT = 60 * 60;        // test vaqti, soniya
const USTUNLAR = ["Sana", "Ism", "Ball", "To‘g‘ri", "Oson", "O‘rtacha", "Qiyin", "Xato savollar",
  "Javoblar", "Vaqt", "Soniya", "Vaqt tugadi", "Chiqishlar", "Telegram", "Telegram ID", "Qurilma", "Kod"];
const KOD_VARAQ = "Kodlar";
const KOD_USTUNLAR = ["Kod", "Holat", "Ism", "Telegram", "Telegram ID", "Qurilma", "Boshlangan", "Tugagan", "Ball", "Javoblar"];
const KOD_HARFLAR = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // 0/O va 1/I adashmasligi uchun yo'q

function testId_(id) {
  return Object.prototype.hasOwnProperty.call(TESTLAR, id) ? id : "asosiy";
}

function varaq_(nom, ustunlar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nom);
  if (!sh) {
    sh = ss.insertSheet(nom);
    sh.appendRow(ustunlar);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < ustunlar.length) {
    sh.getRange(1, 1, 1, ustunlar.length).setValues([ustunlar]);
  }
  return sh;
}
function natijaVaraq_(id) { return varaq_(TESTLAR[id].varaq, USTUNLAR); }
function kodVaraq_() { return varaq_(KOD_VARAQ, KOD_USTUNLAR); }

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
function javoblar_(v) {
  return String(v || "").toUpperCase().replace(/[^ABCD-]/g, "-").slice(0, 30);
}
function kodNorm_(v) {
  return String(v || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}
function kodKorinish_(k) {
  return k.slice(0, 4) + "-" + k.slice(4);
}
function mmss_(s) {
  return String(Math.floor(s / 60)).padStart(2, "0") + ":" + String(s % 60).padStart(2, "0");
}

function natijaYoz_(id, d, javoblar, soniya, kod) {
  const h = hisobla_(id, javoblar);
  natijaVaraq_(id).appendRow([new Date(), matn_(d.name, 100), h.ball, h.togri, h.oson, h.orta, h.qiyin,
    h.xato.join(", ") || "yo‘q", javoblar, mmss_(soniya), soniya, d.timeUp === true || d.timeUp === "1" ? "ha" : "",
    Number(d.leaves) || 0, matn_(d.tg, 100), matn_(d.tgId, 30), matn_(d.device, 40), kod ? kodKorinish_(kod) : ""]);
  return h;
}

// Bepul test natijasi (sayt o'zi yuboradi)
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const id = testId_(d.test);
  if (TESTLAR[id].pullik) return ContentService.createTextOutput("pullik");
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    natijaYoz_(id, d, javoblar_(d.answers), Math.min(LIMIT, Number(d.seconds) || 0), "");
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "start") return json_(qulf_(function () { return boshlash_(p); }));
  if (p.action === "finish") return json_(qulf_(function () { return tugatish_(p); }));
  if (p.key !== ADMIN_PAROL) return json_({ error: "parol" });
  if (p.action === "codes") return json_(kodlar_());
  if (p.action === "gen") return json_(qulf_(function () { return kodYarat_(Number(p.n) || 10); }));
  return json_(reyting_(testId_(p.test)));
}

function qulf_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Kod qatorini topish: [qator raqami (1 dan), qiymatlar] yoki null
function kodTop_(kod) {
  const sh = kodVaraq_(), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (kodNorm_(v[i][0]) === kod) return { sh: sh, row: i + 1, v: v[i] };
  return null;
}

// O'quvchi kod bilan testni boshlaydi (yoki shu qurilmada davom ettiradi)
function boshlash_(p) {
  if (!PRO_DATA) return { error: "yoq" };
  const kod = kodNorm_(p.code), device = String(p.device || "");
  if (kod.length < 6) return { error: "kod" };
  const k = kodTop_(kod);
  if (!k) return { error: "kod" };
  const holat = String(k.v[1]);
  if (holat === "tugagan") return { error: "tugagan" };
  let boshlangan = k.v[6] instanceof Date ? k.v[6] : null;
  if (holat === "boshlangan") {
    if (!device || String(k.v[5]) !== device) return { error: "band" };
  } else {
    boshlangan = new Date();
    k.sh.getRange(k.row, 2, 1, 6).setValues([["boshlangan", matn_(p.name, 100), matn_(p.tg, 100), matn_(p.tgId, 30), matn_(device, 40), boshlangan]]);
  }
  const otgan = Math.max(0, Math.floor((Date.now() - boshlangan.getTime()) / 1000));
  return { ok: true, elapsed: otgan, q: PRO_DATA.savollar };
}

// Test tugadi: ball hisoblanadi, natija yoziladi, kalit va tahlil qaytariladi
function tugatish_(p) {
  if (!PRO_DATA) return { error: "yoq" };
  const kod = kodNorm_(p.code), device = String(p.device || "");
  const k = kodTop_(kod);
  if (!k || !device || String(k.v[5]) !== device) return { error: "kod" };
  let javoblar, h;
  if (String(k.v[1]) === "tugagan") {
    javoblar = javoblar_(k.v[9]);
    h = hisobla_("pro", javoblar);
  } else if (String(k.v[1]) === "boshlangan") {
    javoblar = javoblar_(p.answers);
    const boshlangan = k.v[6] instanceof Date ? k.v[6].getTime() : Date.now();
    const soniya = Math.min(LIMIT, Math.max(0, Math.floor((Date.now() - boshlangan) / 1000)));
    h = natijaYoz_("pro", { name: k.v[2], tg: k.v[3], tgId: k.v[4], device: device, leaves: p.leaves,
      timeUp: p.timeUp === "1" || soniya >= LIMIT }, javoblar, soniya, kod);
    k.sh.getRange(k.row, 2).setValue("tugagan");
    k.sh.getRange(k.row, 8, 1, 3).setValues([[new Date(), h.ball, javoblar]]);
  } else {
    return { error: "kod" };
  }
  return { ok: true, key: TESTLAR.pro.kalit, tahlil: PRO_DATA.tahlil, ball: h.ball };
}

function kodYarat_(n) {
  n = Math.max(1, Math.min(200, Math.floor(n)));
  const sh = kodVaraq_(), bor = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { bor[kodNorm_(r[0])] = true; });
  const yangi = [];
  while (yangi.length < n) {
    let k = "";
    for (let i = 0; i < 8; i++) k += KOD_HARFLAR[Math.floor(Math.random() * KOD_HARFLAR.length)];
    if (bor[k]) continue;
    bor[k] = true;
    yangi.push([kodKorinish_(k), "bo‘sh", "", "", "", "", "", "", "", ""]);
  }
  sh.getRange(sh.getLastRow() + 1, 1, yangi.length, KOD_USTUNLAR.length).setValues(yangi);
  return { ok: true, codes: yangi.map(function (r) { return r[0]; }) };
}

function kodlar_() {
  const tz = Session.getScriptTimeZone();
  const v = kodVaraq_().getDataRange().getValues().slice(1);
  return { codes: v.filter(function (r) { return r[0]; }).map(function (r) {
    return { kod: String(r[0]), holat: String(r[1]), ism: String(r[2]).replace(/^'/, ""), ball: r[8] === "" ? null : Number(r[8]),
      sana: r[6] instanceof Date ? Utilities.formatDate(r[6], tz, "dd.MM.yyyy HH:mm") : "" };
  }) };
}

function reyting_(id) {
  const tz = Session.getScriptTimeZone();
  const qiymatlar = natijaVaraq_(id).getDataRange().getValues().slice(1);
  const korilgan = {};
  const rows = qiymatlar.map(function (r) {
    // Takroriy urinish: bir xil Telegram ID, qurilma yoki ism
    const kim = [String(r[14] || ""), String(r[15] || ""), String(r[1] || "").trim().toLowerCase()];
    const takror = kim.some(function (x) { return x && korilgan[x]; });
    kim.forEach(function (x) { if (x) korilgan[x] = true; });
    return {
      ism: String(r[1]).replace(/^'/, ""), ball: Number(r[2]) || 0, togri: r[3], oson: r[4], orta: r[5], qiyin: r[6],
      xato: String(r[7]), javoblar: String(r[8]), soniya: Number(r[10]) || 0, tugadi: r[11] === "ha", chiqish: r[12],
      sana: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "dd.MM.yyyy HH:mm") : String(r[0]), takror: takror,
      kod: String(r[16] || "")
    };
  });
  rows.sort(function (a, b) { return b.ball - a.ball || a.soniya - b.soniya; });
  return { rows: rows };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
