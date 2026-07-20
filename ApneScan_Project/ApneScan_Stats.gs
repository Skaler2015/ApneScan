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
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
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
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
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
  sh.appendRow([day, 0]);
  return { row: sh.getLastRow(), count: 0 };
}

// Aaj ka TOTAL — purane bug ki wajah se ek hi din ki kai rows ban gayi hain,
// isliye sab matching rows ko jod kar dikhao (data waapas sahi dikhega).
function _todayCount() {
  var sh = _sheet("daily", ["date", "count"]);
  var day = _today();
  var data = sh.getDataRange().getValues();
  var sum = 0;
  for (var i = 1; i < data.length; i++) {
    if (_cellDay(data[i][0]) === day) sum += Number(data[i][1]) || 0;
  }
  return sum;
}

function _addToday(inc) {
  var sh = _sheet("daily", ["date", "count"]);
  var t = _getToday();
  var nv = t.count + inc;
  sh.getRange(t.row, 2).setValue(nv);
  return nv;
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

function _stats() {
  return {
    ok: true,
    total: _getTotal(),
    today: _todayCount(),
    online: _onlineCount()
  };
}

// ---- HTTP entry points ----
// GET  ?action=stats            -> current numbers
// GET  ?action=ping&client=ID   -> mark online + return numbers
// GET  ?action=scan&client=ID&n=1 -> add scans + return numbers
// (POST bhi same params ke saath chalega)
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
        _touchOnline(client);
      } else if (action === "ping") {
        _touchOnline(client);
      }
      out = _stats();
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
