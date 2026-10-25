// Huquq testlari: natijalarni qabul qilish, admin reytingi va pullik test kodlari.
// Har bir test o'z kaliti bilan hisoblanadi va o'z varag'iga yoziladi.
//
// Yashirin testlar (PRO, 8–10) savollari saytda emas, Google skriptdagi alohida fayllarda turadi:
//   ProData.gs   — const PRO_DATA = {...};
//   Test8910.gs  — const T8910_DATA = {...};
// O'quvchi testni boshlagandagina savollarni oladi, kalit va tahlil esa test tugagach beriladi.
// Bu fayllar faqat Google skriptga qo'yiladi va GitHub'ga yuklanmaydi.

// Admin sahifasiga kirish paroli — o'zingiz o'zgartiring.
const ADMIN_PAROL = "BU_YERGA_PAROL_YOZING";

// Yashirin testlar: data — savollar turgan o'zgaruvchi nomi, kodVaraq — o'z kodlar varag'i,
// bepul: true — kodsiz, hamma uchun ochiq (har bir qurilma/Telegram akkauntga yashirin kod avtomatik beriladi),
// yopiq: true — yangi urinish boshlanmaydi, avval boshlaganlar natijasini ko'raveradi.
const TESTLAR = {
  asosiy: { varaq: "Asosiy test", kalit: "DABCADBCADBCADBACDBCBDCADBCADB" },
  pro:    { varaq: "PRO test", data: "PRO_DATA", kodVaraq: "Kodlar", bepul: true, yopiq: true },
  t8910:  { varaq: "8–10-sinf test", data: "T8910_DATA", kodVaraq: "Kodlar 8–10", bepul: true, yopiq: true }
};
function data_(id) {
  const n = TESTLAR[id].data;
  if (n === "PRO_DATA") return typeof PRO_DATA !== "undefined" ? PRO_DATA : null;
  if (n === "T8910_DATA") return typeof T8910_DATA !== "undefined" ? T8910_DATA : null;
  return null;
}
function pullik_(id) { return !!TESTLAR[id].data; }
function kalit_(id) {
  if (TESTLAR[id].kalit) return TESTLAR[id].kalit;
  const d = data_(id);
  return d ? d.kalit : "";
}
const BALL = [0.9, 1.5, 2.6]; // 1–10, 11–20, 21–30-savollar
const LIMIT = 60 * 60;        // test vaqti, soniya
const USTUNLAR = ["Sana", "Ism", "Ball", "To‘g‘ri", "Oson", "O‘rtacha", "Qiyin", "Xato savollar",
  "Javoblar", "Vaqt", "Soniya", "Vaqt tugadi", "Chiqishlar", "Telegram", "Telegram ID", "Qurilma", "Kod", "Urinish"];
// Ball va Javoblar — birinchi urinishniki (reyting va sertifikat shu bo'yicha);
// keyingi urinishlar soni "Urinishlar"da, davom etayotgan qayta urinish boshlangan vaqti "Joriy urinish"da.
const KOD_USTUNLAR = ["Kod", "Holat", "Ism", "Telegram", "Telegram ID", "Qurilma", "Boshlangan", "Tugagan", "Ball", "Javoblar", "Urinishlar", "Joriy urinish"];
const FIKR_VARAQ = "Fikrlar";
const FIKR_USTUNLAR = ["Sana", "Test", "Ism", "Ball", "Fikr", "Kod", "Telegram", "Qurilma"];
const KOD_HARFLAR = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // 0/O va 1/I adashmasligi uchun yo'q

function testId_(id) {
  return Object.prototype.hasOwnProperty.call(TESTLAR, id) ? id : "asosiy";
}
// Kodlar bo'limi uchun: yashirin test bo'lmasa — PRO
function kodTestId_(id) {
  id = testId_(id);
  return pullik_(id) ? id : "pro";
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
function kodVaraq_(id) { return varaq_(TESTLAR[id].kodVaraq, KOD_USTUNLAR); }
function fikrVaraq_() { return varaq_(FIKR_VARAQ, FIKR_USTUNLAR); }

// O'quvchining test haqidagi fikri (natijadan oldin majburiy)
function fikrYoz_(id, ism, ball, fikr, kod, tg, device) {
  const t = String(fikr || "").trim();
  if (!t) return;
  fikrVaraq_().appendRow([new Date(), id, matn_(ism, 100), ball, matn_(t, 1500), kod ? kodKorinish_(kod) : "", matn_(tg, 100), matn_(device, 40)]);
}

function hisobla_(id, javoblar) {
  const k = kalit_(id), b = [0, 0, 0], xato = [];
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

function natijaYoz_(id, d, javoblar, soniya, kod, urinish) {
  const h = hisobla_(id, javoblar);
  natijaVaraq_(id).appendRow([new Date(), matn_(d.name, 100), h.ball, h.togri, h.oson, h.orta, h.qiyin,
    h.xato.join(", ") || "yo‘q", javoblar, mmss_(soniya), soniya, d.timeUp === true || d.timeUp === "1" ? "ha" : "",
    Number(d.leaves) || 0, matn_(d.tg, 100), matn_(d.tgId, 30), matn_(d.device, 40), kod ? kodKorinish_(kod) : "", urinish || 1]);
  return h;
}

// Bepul test natijasi (sayt o'zi yuboradi)
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const id = testId_(d.test);
  if (pullik_(id)) return ContentService.createTextOutput("pullik");
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const h = natijaYoz_(id, d, javoblar_(d.answers), Math.min(LIMIT, Number(d.seconds) || 0), "");
    fikrYoz_(id, d.name, h.ball, d.feedback, "", d.tg, d.device);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "start") return json_(qulf_(function () { return boshlash_(kodTestId_(p.test), p); }));
  if (p.action === "finish") return json_(qulf_(function () { return tugatish_(kodTestId_(p.test), p); }));
  if (p.key !== ADMIN_PAROL) return json_({ error: "parol" });
  if (p.action === "codes") return json_(kodlar_(kodTestId_(p.test)));
  if (p.action === "feedback") return json_(fikrlar_(testId_(p.test)));
  if (p.action === "gen") return json_(qulf_(function () { return kodYarat_(kodTestId_(p.test), Number(p.n) || 10); }));
  return json_(reyting_(testId_(p.test)));
}

function qulf_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// Kod qatorini topish: [qator raqami (1 dan), qiymatlar] yoki null
function kodTop_(id, kod) {
  const sh = kodVaraq_(id), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (kodNorm_(v[i][0]) === kod) return { sh: sh, row: i + 1, v: v[i] };
  return null;
}

// Kod egasi: o'sha qurilma yoki o'sha Telegram akkaunt
function egasi_(k, device, tgId) {
  return (device && String(k.v[5]) === device) || (tgId && String(k.v[4]) === String(tgId));
}

// Bepul rejim: shu qurilma yoki Telegram akkauntga avval berilgan kod (bo'lmasa — null)
function egaTop_(id, device, tgId) {
  const v = kodVaraq_(id).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if ((device && String(v[i][5]) === device) || (tgId && String(v[i][4]) === tgId)) return kodNorm_(v[i][0]);
  }
  return null;
}

// O'quvchi kod bilan testni boshlaydi, davom ettiradi yoki tugatgach qayta yechadi
function boshlash_(id, p) {
  const D = data_(id), T = TESTLAR[id];
  if (!D) return { error: "yoq" };
  const device = String(p.device || ""), tgId = String(p.tgId || "");
  let kod = kodNorm_(p.code);
  if (!kod && T.bepul) {
    if (!device && !tgId) return { error: "kod" };
    const sinov = p.sinov && p.sinov === ADMIN_PAROL; // yopiq testni admin sinab ko'rishi
    kod = egaTop_(id, device, tgId) || (T.yopiq && !sinov ? "" : kodNorm_(kodYarat_(id, 1).codes[0]));
    if (!kod) return { error: "yopiq" };
  }
  if (kod.length < 6) return { error: "kod" };
  const k = kodTop_(id, kod);
  if (!k) return { error: "kod" };
  const holat = String(k.v[1]);
  if (T.yopiq && holat !== "boshlangan" && holat !== "tugagan" && !(p.sinov && p.sinov === ADMIN_PAROL)) return { error: "yopiq" };
  if (holat === "tugagan") {
    // Qayta urinish: cheksiz, reytingga va sertifikatga ta'sir qilmaydi
    if (!egasi_(k, device, tgId)) return { error: "band" };
    let joriy = k.v[11] instanceof Date ? k.v[11] : null;
    if (!joriy || p.yangi === "1") { joriy = new Date(); k.sh.getRange(k.row, 12).setValue(joriy); }
    return { ok: true, retake: true, code: kodKorinish_(kod), attempt: (Number(k.v[10]) || 1) + 1,
      first: { answers: String(k.v[9]), endAt: k.v[7] instanceof Date ? k.v[7].getTime() : 0 },
      elapsed: Math.max(0, Math.floor((Date.now() - joriy.getTime()) / 1000)), q: D.savollar };
  }
  let boshlangan = k.v[6] instanceof Date ? k.v[6] : null;
  if (holat === "boshlangan") {
    if (!egasi_(k, device, tgId)) return { error: "band" };
  } else {
    boshlangan = new Date();
    k.sh.getRange(k.row, 2, 1, 6).setValues([["boshlangan", matn_(p.name, 100), matn_(p.tg, 100), matn_(p.tgId, 30), matn_(device, 40), boshlangan]]);
  }
  const otgan = Math.max(0, Math.floor((Date.now() - boshlangan.getTime()) / 1000));
  return { ok: true, code: kodKorinish_(kod), elapsed: otgan, q: D.savollar };
}

// Test tugadi: ball hisoblanadi, natija yoziladi, kalit va tahlil qaytariladi
function tugatish_(id, p) {
  const D = data_(id);
  if (!D) return { error: "yoq" };
  const kod = kodNorm_(p.code), device = String(p.device || ""), tgId = String(p.tgId || "");
  const k = kodTop_(id, kod);
  if (!k || !egasi_(k, device, tgId)) return { error: "kod" };
  let javoblar, h;
  if (String(k.v[1]) === "tugagan" && k.v[11] instanceof Date) {
    // Qayta urinish: faqat o'quvchiga ko'rsatiladi, jadvalga va reytingga yozilmaydi
    h = hisobla_(id, javoblar_(p.answers));
    k.sh.getRange(k.row, 11, 1, 2).setValues([[(Number(k.v[10]) || 1) + 1, ""]]);
    return { ok: true, retake: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball };
  } else if (String(k.v[1]) === "tugagan") {
    javoblar = javoblar_(k.v[9]);
    h = hisobla_(id, javoblar);
  } else if (String(k.v[1]) === "boshlangan") {
    javoblar = javoblar_(p.answers);
    const boshlangan = k.v[6] instanceof Date ? k.v[6].getTime() : Date.now();
    const soniya = Math.min(LIMIT, Math.max(0, Math.floor((Date.now() - boshlangan) / 1000)));
    h = natijaYoz_(id, { name: k.v[2], tg: k.v[3], tgId: k.v[4], device: device, leaves: p.leaves,
      timeUp: p.timeUp === "1" || soniya >= LIMIT }, javoblar, soniya, kod);
    fikrYoz_(id, k.v[2], h.ball, p.fb, kod, k.v[3], device);
    k.sh.getRange(k.row, 2).setValue("tugagan");
    k.sh.getRange(k.row, 8, 1, 5).setValues([[new Date(), h.ball, javoblar, 1, ""]]);
  } else {
    return { error: "kod" };
  }
  return { ok: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball };
}

function kodYarat_(id, n) {
  n = Math.max(1, Math.min(200, Math.floor(n)));
  const sh = kodVaraq_(id), bor = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) { bor[kodNorm_(r[0])] = true; });
  const yangi = [];
  while (yangi.length < n) {
    let k = "";
    for (let i = 0; i < 8; i++) k += KOD_HARFLAR[Math.floor(Math.random() * KOD_HARFLAR.length)];
    if (bor[k]) continue;
    bor[k] = true;
    const qator = KOD_USTUNLAR.map(function () { return ""; });
    qator[0] = kodKorinish_(k); qator[1] = "bo‘sh";
    yangi.push(qator);
  }
  sh.getRange(sh.getLastRow() + 1, 1, yangi.length, KOD_USTUNLAR.length).setValues(yangi);
  return { ok: true, codes: yangi.map(function (r) { return r[0]; }) };
}

function kodlar_(id) {
  const tz = Session.getScriptTimeZone();
  const v = kodVaraq_(id).getDataRange().getValues().slice(1);
  return { codes: v.filter(function (r) { return r[0]; }).map(function (r) {
    return { kod: String(r[0]), holat: String(r[1]), ism: String(r[2]).replace(/^'/, ""), ball: r[8] === "" ? null : Number(r[8]),
      urinishlar: Number(r[10]) || 0,
      sana: r[6] instanceof Date ? Utilities.formatDate(r[6], tz, "dd.MM.yyyy HH:mm") : "" };
  }) };
}

function fikrlar_(id) {
  const tz = Session.getScriptTimeZone();
  const v = fikrVaraq_().getDataRange().getValues().slice(1).filter(function (r) { return String(r[1]) === id; });
  return { items: v.reverse().map(function (r) {
    return { ism: String(r[2]).replace(/^'/, ""), ball: r[3], fikr: String(r[4]).replace(/^'/, ""),
      sana: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "dd.MM.yyyy HH:mm") : String(r[0]) };
  }) };
}

function reyting_(id) {
  const tz = Session.getScriptTimeZone();
  const qiymatlar = natijaVaraq_(id).getDataRange().getValues().slice(1);
  const korilgan = {};
  const rows = qiymatlar.map(function (r) {
    // Takroriy urinish: bir xil Telegram ID, qurilma yoki ism
    const kim = [String(r[14] || ""), String(r[15] || ""), String(r[1] || "").trim().toLowerCase()];
    const takror = Number(r[17]) > 1 || kim.some(function (x) { return x && korilgan[x]; });
    kim.forEach(function (x) { if (x) korilgan[x] = true; });
    return {
      ism: String(r[1]).replace(/^'/, ""), ball: Number(r[2]) || 0, togri: r[3], oson: r[4], orta: r[5], qiyin: r[6],
      xato: String(r[7]), javoblar: String(r[8]), soniya: Number(r[10]) || 0, tugadi: r[11] === "ha", chiqish: r[12],
      sana: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "dd.MM.yyyy HH:mm") : String(r[0]), takror: takror,
      kod: String(r[16] || ""), urinish: Number(r[17]) || 1
    };
  });
  rows.sort(function (a, b) { return b.ball - a.ball || a.soniya - b.soniya; });
  return { rows: rows };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
