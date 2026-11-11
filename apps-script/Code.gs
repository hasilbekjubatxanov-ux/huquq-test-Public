// Huquq testlari: natijalarni qabul qilish, admin reytingi va pullik test kodlari.
// Har bir test o'z kaliti bilan hisoblanadi va o'z varag'iga yoziladi.
//
// Yashirin testlar (PRO, 8–10) savollari saytda emas, Google skriptdagi alohida fayllarda turadi:
//   ProData.gs   — const PRO_DATA = {...};
//   Test8910.gs  — const T8910_DATA = {...};
//   Test811.gs   — const T811_DATA = {...};
//   NhhData.gs   — const NHH_DATA = {...};
//   KonstData.gs — const KONST_DATA = {...};
// O'quvchi testni boshlagandagina savollarni oladi, kalit va tahlil esa test tugagach beriladi.
// Bu fayllar faqat Google skriptga qo'yiladi va GitHub'ga yuklanmaydi.

// Admin sahifasiga kirish paroli — o'zingiz o'zgartiring.
const ADMIN_PAROL = "BU_YERGA_PAROL_YOZING";
// Parolli testlar: test id → kirish paroli (GitHub'ga haqiqiy parol yozilmaydi, faqat Google skriptda).
const TEST_PAROLLARI = { konst: "BU_YERGA_TEST_PAROLINI_YOZING" };
function testParol_(id) { return Object.prototype.hasOwnProperty.call(TEST_PAROLLARI, id) ? String(TEST_PAROLLARI[id]) : ""; }

// Yashirin testlar: data — savollar turgan o'zgaruvchi nomi, kodVaraq — o'z kodlar varag'i,
// bepul: true — kodsiz, hamma uchun ochiq (har bir qurilma/Telegram akkauntga yashirin kod avtomatik beriladi),
// yopiq: true — yangi urinish boshlanmaydi, avval boshlaganlar natijasini ko'raveradi.
// sanash: true — natija to'g'ri javoblar soni (masalan, 45 dan 40), yozma javobli savollar bilan.
// rasch: true — natija Rash modeli bo'yicha: o'quvchi test tugagach ballini ko'rmaydi, admin "Rash hisoblash"ni
// bosgach hamma natija bir vaqtda e'lon qilinadi va test yangi urinishlar uchun yopiladi.
const TESTLAR = {
  asosiy: { varaq: "Asosiy test", kalit: "DABCADBCADBCADBACDBCBDCADBCADB" },
  pro:    { varaq: "PRO test", data: "PRO_DATA", kodVaraq: "Kodlar", bepul: true, yopiq: true },
  t8910:  { varaq: "8–10-sinf test", data: "T8910_DATA", kodVaraq: "Kodlar 8–10", bepul: true },
  t811:   { varaq: "11-sinf PRO test", data: "T811_DATA", kodVaraq: "Kodlar 11 PRO", bepul: true },
  nhh:    { varaq: "NHH PRO test", data: "NHH_DATA", kodVaraq: "Kodlar NHH", bepul: true, yopiq: true, sanash: true, daqiqa: 90 },
  konst:  { varaq: "Konstitutsiya test", data: "KONST_DATA", kodVaraq: "Kodlar Konst", bepul: true, sanash: true, rasch: true }
};
function data_(id) {
  const n = TESTLAR[id].data;
  if (n === "PRO_DATA") return typeof PRO_DATA !== "undefined" ? PRO_DATA : null;
  if (n === "T8910_DATA") return typeof T8910_DATA !== "undefined" ? T8910_DATA : null;
  if (n === "T811_DATA") return typeof T811_DATA !== "undefined" ? T811_DATA : null;
  if (n === "NHH_DATA") return typeof NHH_DATA !== "undefined" ? NHH_DATA : null;
  if (n === "KONST_DATA") return typeof KONST_DATA !== "undefined" ? KONST_DATA : null;
  return null;
}
function pullik_(id) { return !!TESTLAR[id].data; }
function kalit_(id) {
  if (TESTLAR[id].kalit) return TESTLAR[id].kalit;
  const d = data_(id);
  return d ? d.kalit : "";
}
const BALL = [0.9, 1.5, 2.6]; // 1–10, 11–20, 21–30-savollar
const LIMIT = 60 * 60;        // test vaqti, soniya (testda "daqiqa" berilgan bo'lsa, o'sha)
function limit_(id) { return (TESTLAR[id] && TESTLAR[id].daqiqa ? TESTLAR[id].daqiqa : 60) * 60; }
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

// Yozma javob: katta-kichik harf, apostrof, bo'sh joy va belgilar hisobga olinmaydi; son bo'lsa faqat raqamlar
function normJ_(v) {
  return String(v == null ? "" : v).toLowerCase().replace(/[ʻʼ‘’'`´]/g, "").replace(/[^\p{L}\p{N}]/gu, "");
}
function yozmaTogri_(a, k) {
  if (!a || a === "-") return false;
  return String(k).split("|").some(function (x) {
    if (!/^\d+$/.test(x)) return normJ_(a) === normJ_(x);
    // Son: "100000", "100 000", "100 ming", "yuz ming", "bir yil", "60 kun" — hammasi qabul qilinadi
    return String(a).replace(/\D/g, "") === x || sonOqi_(a) === Number(x);
  });
}
// Javobdagi sonni o'qiydi: raqamlar va so'z bilan yozilgan sonlar (lotin va kirill), "ming", "yuz"
const KIRILL_ = { "а":"a","б":"b","в":"v","г":"g","д":"d","е":"e","ё":"yo","ж":"j","з":"z","и":"i","й":"y","к":"k","л":"l","м":"m","н":"n","о":"o","п":"p","р":"r","с":"s","т":"t","у":"u","ф":"f","х":"x","ц":"ts","ч":"ch","ш":"sh","ъ":"","ь":"","э":"e","ю":"yu","я":"ya","ў":"o","қ":"q","ғ":"g","ҳ":"h" };
const SONLAR_ = { bir: 1, ikki: 2, uch: 3, tort: 4, besh: 5, olti: 6, yetti: 7, sakkiz: 8, toqqiz: 9, on: 10,
  yigirma: 20, ottiz: 30, qirq: 40, ellik: 50, oltmish: 60, yetmish: 70, sakson: 80, toqson: 90 };
function sonOqi_(v) {
  let s = String(v == null ? "" : v).toLowerCase().replace(/[а-яёўқғҳ]/g, function (c) { return KIRILL_[c] != null ? KIRILL_[c] : c; });
  s = s.replace(/[ʻʼ‘’'`´]/g, "").replace(/(\d)[\s.,](?=\d{3}(\D|$))/g, "$1");
  const t = s.match(/\d+|[a-z]+/g) || [];
  let jami = 0, joriy = 0, bor = false;
  t.forEach(function (w) {
    if (/^\d+$/.test(w)) { joriy += Number(w); bor = true; }
    else if (SONLAR_[w] != null) { joriy += SONLAR_[w]; bor = true; }
    else if (w === "yuz") { joriy = (joriy || 1) * 100; bor = true; }
    else if (w === "ming") { jami += (joriy || 1) * 1000; joriy = 0; bor = true; }
    else if (w === "million" || w === "mln") { jami = (jami + (joriy || 1)) * 1000000; joriy = 0; bor = true; }
  });
  return bor ? jami + joriy : null;
}
// To'g'ri javoblar soni bo'yicha test (sanash: true): javoblar "|" bilan ajratilgan
function hisoblaSanash_(id, javoblar) {
  const k = kalit_(id), a = String(javoblar || "").split("|"), xato = [];
  let togri = 0;
  for (let i = 0; i < k.length; i++) {
    const j = a[i] || "-", ok = k[i].length === 1 && /[ABCD]/.test(k[i]) ? j === k[i] : yozmaTogri_(j, k[i]);
    if (ok) togri++;
    else xato.push((i + 1) + " (" + j + "→" + String(k[i]).split("|")[0] + ")");
  }
  return { ball: togri, togri: togri, oson: "", orta: "", qiyin: "", xato: xato };
}

function hisobla_(id, javoblar) {
  if (TESTLAR[id].sanash) return hisoblaSanash_(id, javoblar);
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
function javoblar_(v, id) {
  if (id && TESTLAR[id] && TESTLAR[id].sanash) {
    const n = kalit_(id).length, k = kalit_(id), a = String(v || "").split("|");
    const out = [];
    for (let i = 0; i < n; i++) {
      const j = String(a[i] == null ? "-" : a[i]).trim();
      if (/^[ABCD]$/.test(k[i])) out.push(/^[ABCD]$/.test(j.toUpperCase()) ? j.toUpperCase() : "-");
      else out.push(j ? matn_(j.replace(/\|/g, " "), 60).replace(/^'/, "") : "-");
    }
    return out.join("|");
  }
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
    const h = natijaYoz_(id, d, javoblar_(d.answers, id), Math.min(limit_(id), Number(d.seconds) || 0), "");
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
  if (p.action === "rasch") return json_(qulf_(function () { return raschHisobla_(testId_(p.test)); }));
  if (p.action === "raschoff") return json_(qulf_(function () { return raschBekor_(testId_(p.test)); }));
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
  // Rash testi: natijalar e'lon qilingach yangi urinish boshlanmaydi
  const elon = T.rasch ? raschElon_(id) : null, yopiq = T.yopiq || !!elon;
  // Sinov rejimi faqat admin paroli bilan: noto'g'ri parol bilan test boshlanmaydi
  if (p.sinov && p.sinov !== ADMIN_PAROL) return { error: "yopiq" };
  const sinovOk = !!p.sinov && p.sinov === ADMIN_PAROL;
  // Parolli test: yangi o'quvchi faqat to'g'ri parol bilan boshlaydi (avval boshlaganlar parolsiz davom etadi)
  const tp = testParol_(id), parolOk = !tp || sinovOk || String(p.parol || "") === tp;
  let kod = kodNorm_(p.code);
  if (!kod && T.bepul) {
    if (!device && !tgId) return { error: "kod" };
    const sinov = p.sinov && p.sinov === ADMIN_PAROL; // yopiq testni admin sinab ko'rishi
    const eski = egaTop_(id, device, tgId);
    if (!eski && !parolOk) return { error: "testparol" };
    kod = eski || (yopiq && !sinov ? "" : kodNorm_(kodYarat_(id, 1).codes[0]));
    if (!kod) return { error: "yopiq" };
  }
  if (kod.length < 6) return { error: "kod" };
  const k = kodTop_(id, kod);
  if (!k) return { error: "kod" };
  const holat = String(k.v[1]);
  if (yopiq && holat !== "boshlangan" && holat !== "tugagan" && !(p.sinov && p.sinov === ADMIN_PAROL)) return { error: "yopiq" };
  // Rash testi: natija e'lon qilinmaguncha qayta yechib bo'lmaydi (kalit va tahlil ochilib qolmasligi uchun)
  if (sinovOk) sinovBelgila_(id, kod);
  if (holat === "tugagan" && T.rasch && !elon && !sinovKod_(id, kod)) return { error: egasi_(k, device, tgId) ? "kutish" : "band" };
  if (holat === "tugagan") {
    // Qayta urinish: cheksiz, reytingga va sertifikatga ta'sir qilmaydi
    if (!egasi_(k, device, tgId)) return { error: "band" };
    let joriy = k.v[11] instanceof Date ? k.v[11] : null;
    if (!joriy || p.yangi === "1") { joriy = new Date(); k.sh.getRange(k.row, 12).setValue(joriy); }
    return { ok: true, retake: true, code: kodKorinish_(kod), attempt: (Number(k.v[10]) || 1) + 1,
      first: { answers: String(k.v[9]), endAt: k.v[7] instanceof Date ? k.v[7].getTime() : 0 },
      elapsed: Math.max(0, Math.floor((Date.now() - joriy.getTime()) / 1000)), q: D.savollar, sinov: sinovOk };
  }
  let boshlangan = k.v[6] instanceof Date ? k.v[6] : null;
  if (holat === "boshlangan") {
    if (!egasi_(k, device, tgId)) return { error: "band" };
  } else {
    boshlangan = new Date();
    k.sh.getRange(k.row, 2, 1, 6).setValues([["boshlangan", matn_(p.name, 100), matn_(p.tg, 100), matn_(p.tgId, 30), matn_(device, 40), boshlangan]]);
  }
  const otgan = Math.max(0, Math.floor((Date.now() - boshlangan.getTime()) / 1000));
  return { ok: true, code: kodKorinish_(kod), elapsed: otgan, q: D.savollar, sinov: sinovOk };
}
// Admin sinovi: bu kod natijasi darhol ko'rsatiladi va Rash hisobiga kirmaydi
function sinovBelgila_(id, kod) { PropertiesService.getScriptProperties().setProperty("sinov_" + id + "_" + kod, "1"); }
function sinovKod_(id, kod) { return !!kod && PropertiesService.getScriptProperties().getProperty("sinov_" + id + "_" + kodNorm_(kod)) === "1"; }

// Test tugadi: ball hisoblanadi, natija yoziladi, kalit va tahlil qaytariladi
function tugatish_(id, p) {
  const D = data_(id);
  if (!D) return { error: "yoq" };
  const kod = kodNorm_(p.code), device = String(p.device || ""), tgId = String(p.tgId || "");
  const k = kodTop_(id, kod);
  if (!k || !egasi_(k, device, tgId)) return { error: "kod" };
  const rasch = !!TESTLAR[id].rasch;
  let javoblar, h;
  if (String(k.v[1]) === "tugagan" && k.v[11] instanceof Date) {
    // Qayta urinish: faqat o'quvchiga ko'rsatiladi, jadvalga va reytingga yozilmaydi
    h = hisobla_(id, javoblar_(p.answers, id));
    k.sh.getRange(k.row, 11, 1, 2).setValues([[(Number(k.v[10]) || 1) + 1, ""]]);
    return { ok: true, retake: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball };
  } else if (String(k.v[1]) === "tugagan") {
    javoblar = javoblar_(k.v[9], id);
    h = hisobla_(id, javoblar);
  } else if (String(k.v[1]) === "boshlangan") {
    javoblar = javoblar_(p.answers, id);
    const boshlangan = k.v[6] instanceof Date ? k.v[6].getTime() : Date.now();
    const soniya = Math.min(limit_(id), Math.max(0, Math.floor((Date.now() - boshlangan) / 1000)));
    h = natijaYoz_(id, { name: k.v[2], tg: k.v[3], tgId: k.v[4], device: device, leaves: p.leaves,
      timeUp: p.timeUp === "1" || soniya >= limit_(id) }, javoblar, soniya, kod);
    fikrYoz_(id, k.v[2], h.ball, p.fb, kod, k.v[3], device);
    k.sh.getRange(k.row, 2).setValue("tugagan");
    k.sh.getRange(k.row, 8, 1, 5).setValues([[new Date(), h.ball, javoblar, 1, ""]]);
  } else {
    return { error: "kod" };
  }
  if (rasch) {
    // Natijalar e'lon qilinmaguncha o'quvchiga ball, kalit va tahlil berilmaydi
    if (sinovKod_(id, kod)) return { ok: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball, sinov: true };
    if (!raschElon_(id) && ertaRuxsat_(id, kod)) return { ok: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball, erta: true };
    if (!raschElon_(id)) return { ok: true, kutish: true };
    return { ok: true, key: kalit_(id), tahlil: D.tahlil, ball: h.ball, rasch: raschTop_(id, kod) };
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
  if (TESTLAR[id].rasch) {
    const elon = raschElon_(id), xarita = elon ? raschXarita_(id) : {};
    rows.forEach(function (r) { const x = xarita[kodNorm_(r.kod)]; if (x && !r.takror) { r.rasch = x.ball; r.daraja = x.daraja; r.theta = x.theta; } });
    return { rows: rows, rasch: elon };
  }
  return { rows: rows };
}

// ───────── Rash modeli ─────────
// Savol qiyinligi va o'quvchi qobiliyati (teta, logit) javoblar jadvalidan birgalikda baholanadi (JMLE).
// Ball = 75 × (teta − teta(0 to'g'ri)) / (teta(hammasi to'g'ri) − teta(0 to'g'ri)): 0 dan 75 gacha.
// Darajalar milliy sertifikatdagi chegaralar bo'yicha.
const RASCH_MAX = 75;
const RASCH_DARAJA = [[70, "A+"], [65, "A"], [60, "B+"], [55, "B"], [50, "C+"], [46, "C"]];
const RASCH_USTUNLAR = ["O‘rin", "Ism", "To‘g‘ri", "Teta", "Ball", "Daraja", "Kod", "Soniya", "Chiqishlar"];
const RASCH_SAVOL_USTUNLAR = ["Savol", "To‘g‘ri topganlar", "Foiz", "Qiyinlik (logit)", "Izoh"];
function raschVaraqNom_(id) { return "Rash " + TESTLAR[id].varaq; }
function raschSavolNom_(id) { return "Rash savollar " + TESTLAR[id].varaq; }
function daraja_(ball) {
  for (let i = 0; i < RASCH_DARAJA.length; i++) if (ball >= RASCH_DARAJA[i][0]) return RASCH_DARAJA[i][1];
  return "";
}
// E'lon qilingan bo'lsa — { sana, n }, aks holda null
// Erta natija: "Erta natija" varag'ida kodi yozilgan o'quvchi Rash e'lonini kutmasdan
// to'g'ri javoblari va tahlilni ko'radi (Rash bali e'londan keyin chiqadi). Qator qo'shish kifoya.
const ERTA_VARAQ = "Erta natija";
function ertaRuxsat_(id, kod) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERTA_VARAQ);
  if (!sh || !kod) return false;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === id && kodNorm_(v[i][1]) === kodNorm_(kod)) return true;
  }
  return false;
}
function raschElon_(id) {
  const v = PropertiesService.getScriptProperties().getProperty("rasch_" + id);
  if (!v) return null;
  try { return JSON.parse(v); } catch (e) { return null; }
}
function raschBekor_(id) {
  PropertiesService.getScriptProperties().deleteProperty("rasch_" + id);
  return { ok: true };
}
// Javoblar satridan 0/1 vektor
function togriVektor_(id, javoblar) {
  const k = kalit_(id), sanash = !!TESTLAR[id].sanash;
  const a = sanash ? String(javoblar || "").split("|") : String(javoblar || "").split("");
  const out = [];
  for (let i = 0; i < k.length; i++) {
    const j = a[i] || "-";
    out.push((String(k[i]).length === 1 && /[ABCD]/.test(k[i]) ? j === k[i] : yozmaTogri_(j, k[i])) ? 1 : 0);
  }
  return out;
}
function logistik_(x) { return 1 / (1 + Math.exp(-x)); }
// r ta to'g'ri javobga mos teta (b — savollar qiyinligi). 0 va to'liq natija 0,3 ga suriladi.
function tetaTop_(r, b) {
  const L = b.length;
  const rr = r <= 0 ? 0.3 : r >= L ? L - 0.3 : r;
  let t = Math.log(rr / (L - rr));
  for (let it = 0; it < 100; it++) {
    let e = 0, v = 0;
    for (let i = 0; i < L; i++) { const p = logistik_(t - b[i]); e += p; v += p * (1 - p); }
    const d = Math.max(-1, Math.min(1, (rr - e) / Math.max(v, 1e-9)));
    t += d;
    if (Math.abs(d) < 1e-6) break;
  }
  return t;
}
// X — o'quvchilar × savollar (0/1). Qaytaradi: { b: [qiyinlik yoki null], faol: [savol indekslari], teta: [...] }
function rasch_(X) {
  const P = X.length, L = P ? X[0].length : 0;
  let savol = [], odam = [];
  for (let i = 0; i < L; i++) savol.push(i);
  for (let p = 0; p < P; p++) odam.push(p);
  // Hamma topgan / hech kim topmagan savollar va 0 / to'liq natijali o'quvchilar baholashdan chiqariladi
  for (let ozgardi = true; ozgardi;) {
    ozgardi = false;
    const o2 = odam.filter(function (p) { let r = 0; savol.forEach(function (i) { r += X[p][i]; }); return r > 0 && r < savol.length; });
    if (o2.length !== odam.length) { odam = o2; ozgardi = true; }
    const s2 = savol.filter(function (i) { let s = 0; odam.forEach(function (p) { s += X[p][i]; }); return s > 0 && s < odam.length; });
    if (s2.length !== savol.length) { savol = s2; ozgardi = true; }
  }
  const Ls = savol.length, Np = odam.length;
  if (Ls < 2 || Np < 2) return null;
  const s = savol.map(function (i) { let c = 0; odam.forEach(function (p) { c += X[p][i]; }); return c; });
  const nr = []; for (let r = 0; r <= Ls; r++) nr.push(0);
  odam.forEach(function (p) { let r = 0; savol.forEach(function (i) { r += X[p][i]; }); nr[r]++; });
  let b = s.map(function (c) { return Math.log((Np - c) / c); });
  const markaz = function () { const m = b.reduce(function (a, x) { return a + x; }, 0) / Ls; b = b.map(function (x) { return x - m; }); };
  markaz();
  const th = []; for (let r = 0; r <= Ls; r++) th.push(r > 0 && r < Ls ? Math.log(r / (Ls - r)) : 0);
  for (let it = 0; it < 300; it++) {
    let maxd = 0;
    for (let i = 0; i < Ls; i++) {
      let e = 0, v = 0;
      for (let r = 1; r < Ls; r++) if (nr[r]) { const p = logistik_(th[r] - b[i]); e += nr[r] * p; v += nr[r] * p * (1 - p); }
      const d = Math.max(-1, Math.min(1, (e - s[i]) / Math.max(v, 1e-9)));
      b[i] += d; maxd = Math.max(maxd, Math.abs(d));
    }
    markaz();
    for (let r = 1; r < Ls; r++) if (nr[r]) {
      let e = 0, v = 0;
      for (let i = 0; i < Ls; i++) { const p = logistik_(th[r] - b[i]); e += p; v += p * (1 - p); }
      const d = Math.max(-1, Math.min(1, (r - e) / Math.max(v, 1e-9)));
      th[r] += d; maxd = Math.max(maxd, Math.abs(d));
    }
    if (maxd < 1e-5) break;
  }
  // JMLE siljishini tuzatish
  b = b.map(function (x) { return x * (Ls - 1) / Ls; });
  const tetaR = []; for (let r = 0; r <= Ls; r++) tetaR.push(tetaTop_(r, b));
  const teta = X.map(function (row) { let r = 0; savol.forEach(function (i) { r += row[i]; }); return tetaR[r]; });
  const bAll = []; for (let i = 0; i < L; i++) bAll.push(null);
  savol.forEach(function (i, k) { bAll[i] = b[k]; });
  return { b: bAll, faol: savol, teta: teta, tetaMin: tetaR[0], tetaMax: tetaR[Ls] };
}
// Admin: Rash bo'yicha hisoblash, varaqlarga yozish va e'lon qilish
function raschHisobla_(id) {
  if (!TESTLAR[id].rasch) return { error: "rasch" };
  const qiymatlar = natijaVaraq_(id).getDataRange().getValues().slice(1);
  const korilgan = {}, qatnashchi = [];
  qiymatlar.forEach(function (r) {
    const kim = [String(r[14] || ""), String(r[15] || ""), String(r[1] || "").trim().toLowerCase()];
    const takror = Number(r[17]) > 1 || kim.some(function (x) { return x && korilgan[x]; });
    kim.forEach(function (x) { if (x) korilgan[x] = true; });
    if (takror || sinovKod_(id, String(r[16] || ""))) return; // admin sinovi hisobga kirmaydi
    qatnashchi.push({ ism: String(r[1]).replace(/^'/, ""), kod: String(r[16] || ""), soniya: Number(r[10]) || 0,
      chiqish: Number(r[12]) || 0, x: togriVektor_(id, javoblar_(r[8], id)) });
  });
  if (qatnashchi.length < 2) return { error: "kam", n: qatnashchi.length };
  const R = rasch_(qatnashchi.map(function (q) { return q.x; }));
  if (!R) return { error: "kam", n: qatnashchi.length };
  const n = qatnashchi.length, oraliq = R.tetaMax - R.tetaMin;
  qatnashchi.forEach(function (q, i) {
    q.togri = q.x.reduce(function (a, x) { return a + x; }, 0);
    q.teta = Math.round(R.teta[i] * 100) / 100;
    q.ball = Math.round(Math.max(0, Math.min(RASCH_MAX, oraliq > 0 ? RASCH_MAX * (R.teta[i] - R.tetaMin) / oraliq : 0)) * 10) / 10;
    q.daraja = daraja_(q.ball);
  });
  qatnashchi.sort(function (a, b) { return b.ball - a.ball || a.chiqish - b.chiqish || a.soniya - b.soniya; });
  const sh = varaq_(raschVaraqNom_(id), RASCH_USTUNLAR);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, RASCH_USTUNLAR.length).clearContent();
  sh.getRange(2, 1, n, RASCH_USTUNLAR.length).setValues(qatnashchi.map(function (q, i) {
    return [i + 1, matn_(q.ism, 100), q.togri, q.teta, q.ball, q.daraja, q.kod, q.soniya, q.chiqish];
  }));
  const L = R.b.length, savollar = [];
  for (let i = 0; i < L; i++) {
    let c = 0; qatnashchi.forEach(function (q) { c += q.x[i]; });
    savollar.push({ n: i + 1, togri: c, foiz: Math.round(c / n * 1000) / 10, b: R.b[i] == null ? null : Math.round(R.b[i] * 100) / 100 });
  }
  const sh2 = varaq_(raschSavolNom_(id), RASCH_SAVOL_USTUNLAR);
  if (sh2.getLastRow() > 1) sh2.getRange(2, 1, sh2.getLastRow() - 1, RASCH_SAVOL_USTUNLAR.length).clearContent();
  sh2.getRange(2, 1, L, RASCH_SAVOL_USTUNLAR.length).setValues(savollar.map(function (q) {
    return [q.n, q.togri, q.foiz, q.b == null ? "" : q.b, q.b == null ? (q.togri === 0 ? "hech kim topmagan — hisobga kirmadi" : "hamma topgan — hisobga kirmadi") : ""];
  }));
  const elon = { sana: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm"), n: n };
  PropertiesService.getScriptProperties().setProperty("rasch_" + id, JSON.stringify(elon));
  return { ok: true, n: n, sana: elon.sana, savollar: savollar };
}
// Kod → { ball, daraja, teta, orin, jami }
function raschXarita_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(raschVaraqNom_(id)), out = {};
  if (!sh) return out;
  const v = sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== ""; });
  v.forEach(function (r) {
    const k = kodNorm_(r[6]);
    if (k) out[k] = { orin: Number(r[0]), togri: Number(r[2]), teta: Number(r[3]), ball: Number(r[4]), daraja: String(r[5]), jami: v.length };
  });
  return out;
}
function raschTop_(id, kod) { return raschXarita_(id)[kodNorm_(kod)] || null; }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
