/**************************************************************************
 * ApneScan — Worldwide Stats  (Google Apps Script backend)
 * ------------------------------------------------------------------------
 * Ye ek FREE server hai (koi hosting nahi chahiye). Ye do cheezein rakhta hai:
 *   - kul ab tak kitne document scan hue (total)
 *   - aaj kitne scan hue (per-day)
 *   - abhi kitne log online hai (pichhle 5 min me active app)
 *
 * PRIVACY: yahan sirf GINTI aati hai. Kabhi bhi document ya patient ki
 * jankari nahi bheji jaati.
 *
 * ====== SETUP (ek baar, 5 minute) ======
 * 1. https://sheets.google.com par ek nayi blank Google Sheet banao.
 * 2. Us sheet me: Extensions -> Apps Script.
 * 3. Saara default code hata kar YE poora code paste karo. Save (disk icon).
 * 4. Upar "Deploy" -> "New deployment" -> gear icon -> "Web app".
 *      - Description: ApneScan Stats
 *      - Execute as:  Me
 *      - Who has access:  Anyone
 *    "Deploy" dabao, permissions allow karo.
 * 5. Ek "Web app URL" milega (https://script.google.com/macros/s/..../exec).
 *    USE COPY karke mujhe bhej dena — main app me daal dunga.
 *
 * (Sheet apne aap 3 tab bana lega: totals, daily, online.)
 **************************************************************************/

// Optional: ek secret taaki koi aur aapki ginti na badha sake. App me bhi
// yahi daalna hoga. Khaali chhod sakte ho (tab koi secret nahi lagega).
var SECRET = "";

// IMPORTANT: "Aaj (today)" ki ginti ISI timezone ke hisaab se hoti hai. Pehle
// ye script ke apne timezone par thi — agar wo India (IST) nahi hota to subah
// IST me kiye scan server ke hisaab se "kal" me chale
// jaate the aur "Aaj (duniya)" 0 dikhta tha. Ab poora din/hafta/ghanta India
// (Asia/Kolkata = UTC+5:30) me gina jaata hai. Doosre desh ke liye yahi badal
// dena (jaise "America/New_York").
var TZ = "Asia/Kolkata";

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _sheet(name, headers) {
  var ss = _ss();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.appendRow(headers);
  }
  return sh;
}

function _today() {
  // Server date, YYYY-MM-DD
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd");
}

function _getTotal() {
  var sh = _sheet("totals", ["key", "value"]);
  var v = sh.getRange("B1").getValue();
  return Number(v) || 0;
}

function _setTotal(n) {
  var sh = _sheet("totals", ["key", "value"]);
  sh.getRange("A1").setValue("total_scans");
  sh.getRange("B1").setValue(n);
}

// Sheet me date likhte hi Google use TEXT se DATE bana deta hai, isliye seedha
// String(cell) === "2026-07-20" kabhi match nahi hota tha — har scan par nayi
// row banti thi aur "Aaj (today)" hamesha 0 dikhta tha. Ye helper dono ko
// yyyy-MM-dd me laakar sahi compare karta hai.
function _cellDay(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  }
  return String(v);
}

function _getToday() {
  var sh = _sheet("daily", ["date", "count"]);
  var day = _today();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (_cellDay(data[i][0]) === day) return { row: i + 1, count: Number(data[i][1]) || 0 };
  }
  // Nayi row: date ko TEXT ke roop me likho. Warna Google "2026-07-23" ko
  // apne aap DATE bana deta tha aur read-back par timezone ke hisaab se din
  // shift ho jaata tha (isliye "Aaj" 0 dikhta tha). "@"=text format pehle
  // laga kar likhne se din kabhi shift nahi hota.
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1).setNumberFormat("@");
  sh.getRange(r, 1).setValue(day);
  sh.getRange(r, 2).setValue(0);
  return { row: r, count: 0 };
}

// ====== AAJ ki ginti — ab 'daily' sheet ke date-cell par NAHI ======
// Purani daily-sheet wali ginti timezone/date-convert ki wajah se 0 aati thi.
// Import/Print ki ginti (jo bilkul theek chalti hai) 'totals' sheet me ek
// TEXT key ("imports"/"prints") se hoti hai — exact string match, koi date
// nahi. Isliye "aaj" ki ginti bhi ab usi pakke tareeke se: har din ki apni
// key  day_YYYY-MM-DD  (jaise day_2026-07-23). Ye kabhi date me convert nahi
// hoti, isliye "aaj" hamesha sahi ginega.
function _dayKey() {
  return "day_" + _today();
}

function _todayCount() {
  return _getCounter(_dayKey()).val;
}

function _addToday(inc) {
  return _bumpCounter(_dayKey(), inc);
}

// Online = alag-alag client jinhone pichhle 5 min me ping kiya
function _touchOnline(clientId) {
  if (!clientId) return;
  var sh = _sheet("online", ["client", "last_seen"]);
  var data = sh.getDataRange().getValues();
  var now = new Date().getTime();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(clientId)) {
      sh.getRange(i + 1, 2).setValue(now);
      return;
    }
  }
  sh.appendRow([clientId, now]);
}

function _onlineCount() {
  var sh = _sheet("online", ["client", "last_seen"]);
  var data = sh.getDataRange().getValues();
  var now = new Date().getTime();
  var WINDOW = 5 * 60 * 1000; // 5 minutes
  var live = 0;
  var keepRows = [["client", "last_seen"]];
  for (var i = 1; i < data.length; i++) {
    var seen = Number(data[i][1]) || 0;
    if (now - seen <= WINDOW) {
      live++;
      keepRows.push([data[i][0], data[i][1]]);
    }
  }
  // Purani entries kabhi-kabhi saaf kar do (sheet chhoti rahe)
  if (data.length - 1 > live + 200) {
    sh.clear();
    sh.getRange(1, 1, keepRows.length, 2).setValues(keepRows);
  }
  return live;
}

// ---- v22 extras: users/version/desh (sirf GINTI — kabhi koi document nahi) ----
// col 6 = us client ke kul scans (rank ke liye).
function _touchClient(clientId, ver, country, n) {
  if (!clientId) return;
  var sh = _sheet("clients", ["client", "first_seen", "last_seen", "version", "country", "scans"]);
  var data = sh.getDataRange().getValues();
  var now = new Date().getTime();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(clientId)) {
      sh.getRange(i + 1, 3).setValue(now);
      if (ver) sh.getRange(i + 1, 4).setValue(ver);
      if (country) sh.getRange(i + 1, 5).setValue(country);
      if (n) sh.getRange(i + 1, 6).setValue((Number(data[i][5]) || 0) + n);
      return;
    }
  }
  sh.appendRow([clientId, now, now, ver || "", country || "", n || 0]);
}

// Kisi client ka rank (kul scans ke hisaab se, 1 = sabse zyada) + uske scans
function _rankOf(clientId) {
  if (!clientId) return { rank: 0, myscans: 0 };
  var sh = _sheet("clients", ["client", "first_seen", "last_seen", "version", "country", "scans"]);
  var data = sh.getDataRange().getValues();
  var mine = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(clientId)) { mine = Number(data[i][5]) || 0; break; }
  }
  var rank = 1;
  for (var j = 1; j < data.length; j++) {
    if ((Number(data[j][5]) || 0) > mine) rank++;
  }
  return { rank: rank, myscans: mine };
}

// Aaj ke 24 ghanton ka worldwide array (0..23) — hourly sheet se
function _todayHours() {
  var sh = _sheet("hourly", ["hour", "count"]);
  var data = sh.getDataRange().getValues();
  var day = _today();
  var out = [];
  for (var h = 0; h < 24; h++) out.push(0);
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0] || "");
    if (key.indexOf(day + "-") === 0) {
      var hh = parseInt(key.slice(day.length + 1), 10);
      if (hh >= 0 && hh < 24) out[hh] = Number(data[i][1]) || 0;
    }
  }
  return out;
}

function _usersCount() {
  var sh = _sheet("clients", ["client", "first_seen", "last_seen", "version", "country"]);
  return Math.max(0, sh.getLastRow() - 1);
}

function _breakdown(col) {
  // col: 4 = version, 5 = country
  var sh = _sheet("clients", ["client", "first_seen", "last_seen", "version", "country"]);
  var data = sh.getDataRange().getValues();
  var out = {};
  for (var i = 1; i < data.length; i++) {
    var k = String(data[i][col - 1] || "").trim();
    if (k) out[k] = (out[k] || 0) + 1;
  }
  return out;
}

function _week() {
  // pichhle 7 din: [[date, count], ...] purane-se-naye — wahi pakki
  // day_YYYY-MM-DD key se (jaise 'aaj').
  var out = [];
  for (var j = 6; j >= 0; j--) {
    var dt = new Date(new Date().getTime() - j * 86400000);
    var key = Utilities.formatDate(dt, TZ, "yyyy-MM-dd");
    out.push([key, _getCounter("day_" + key).val]);
  }
  return out;
}

function _bumpHour(n) {
  var sh = _sheet("hourly", ["hour", "count"]);
  var key = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd-HH");
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue((Number(data[i][1]) || 0) + n);
      return;
    }
  }
  sh.appendRow([key, n]);
  if (data.length > 200) sh.deleteRows(2, data.length - 100);   // purana saaf
}

function _hourCount() {
  var sh = _sheet("hourly", ["hour", "count"]);
  var key = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd-HH");
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return Number(data[i][1]) || 0;
  }
  return 0;
}

function _bumpPeak(online) {
  // aaj ka sabse bada online — pakki  peak_YYYY-MM-DD  key se (max rakhta hai)
  var c = _getCounter("peak_" + _today());
  if (online > c.val) {
    _sheet("totals", ["key", "value"]).getRange(c.row, 2).setValue(online);
    return online;
  }
  return c.val;
}

// Generic ginti-counter (import/print jaise) — totals sheet me key-value row.
function _getCounter(key) {
  var sh = _sheet("totals", ["key", "value"]);
  var data = sh.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === key) return { row: i + 1, val: Number(data[i][1]) || 0 };
  }
  sh.appendRow([key, 0]);
  return { row: sh.getLastRow(), val: 0 };
}

function _bumpCounter(key, n) {
  var c = _getCounter(key);
  _sheet("totals", ["key", "value"]).getRange(c.row, 2).setValue(c.val + n);
  return c.val + n;
}

function _stats(client) {
  var online = _onlineCount();
  var rk = _rankOf(client || "");
  return {
    ok: true,
    srv: 4,                 // server code version — redeploy check ke liye
    today_key: _dayKey(),   // konsi key gini ja rahi hai (diagnosis)
    total: _getTotal(),
    today: _todayCount(),
    online: online,
    users: _usersCount(),
    week: _week(),
    todayHours: _todayHours(),
    peak: _bumpPeak(online),
    hour: _hourCount(),
    imports: _getCounter("imports").val,
    prints: _getCounter("prints").val,
    versions: _breakdown(4),
    countries: _breakdown(5),
    rank: rk.rank,
    myscans: rk.myscans
  };
}

// ---- HTTP entry points ----
// GET  ?action=stats            -> current numbers
// GET  ?action=ping&client=ID   -> mark online + return numbers
// GET  ?action=scan&client=ID&n=1 -> add scans + return numbers
// GET  ?action=event&client=ID&imp=1  -> +1 import  (prt=1 -> +1 print)
// (POST bhi same params ke saath chalega; imp/prt kisi bhi action ke saath)
function doGet(e) {
  return _handle(e);
}
function doPost(e) {
  return _handle(e);
}

function _handle(e) {
  var out = { ok: false };
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (SECRET && p.secret !== SECRET && (p.action === "scan")) {
      out = { ok: false, error: "bad secret" };
      return _json(out);
    }
    var action = p.action || "stats";
    var client = p.client || "";
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (action === "scan") {
        var n = Math.max(0, Math.min(100, parseInt(p.n || "1", 10) || 1));
        _setTotal(_getTotal() + n);
        _addToday(n);
        _bumpHour(n);
        _touchOnline(client);
        _touchClient(client, p.v || "", p.c || "", n);
      } else if (action === "ping") {
        _touchOnline(client);
        _touchClient(client, p.v || "", p.c || "", 0);
      }
      // import/print ki ginti — kisi bhi action ke saath aa sakti hai
      var imp = Math.max(0, Math.min(500, parseInt(p.imp || "0", 10) || 0));
      var prt = Math.max(0, Math.min(500, parseInt(p.prt || "0", 10) || 0));
      if (imp) _bumpCounter("imports", imp);
      if (prt) _bumpCounter("prints", prt);
      out = _stats(client);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return _json(out);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test button (Apps Script me "Run" -> testStats)
function testStats() {
  Logger.log(_stats());
}
