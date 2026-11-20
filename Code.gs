// CAFÉ ONE — ربط الأجهزة وحفظ نسخة Excel على Google Sheets
// برمجة د. عبدالله صالح · جميع الحقوق محفوظة
// الصق هذا الكود كما هو. لا حاجة لتعديل أي شيء.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var b = JSON.parse(e.postData.contents);
    var P = PropertiesService.getScriptProperties();
    var key = P.getProperty("K");
    if (!key) {
      key = String(b.k || "");
      if (!key) return out({ err: "key" });
      P.setProperty("K", key);
    }
    if (String(b.k) !== key) return out({ err: "key" });

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (b.dump) {
      b.dump.forEach(function (t) {
        var name = "DB " + t[0];
        var sh = ss.getSheetByName(name) || ss.insertSheet(name);
        sh.clear();
        var rows = t[1];
        var w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
        var data = rows.map(function (r) {
          var x = r.map(function (v) {
            return (typeof v === "string" && /^[=+\-@]/.test(v)) ? "'" + v : v;
          });
          while (x.length < w) x.push("");
          return x;
        });
        if (data.length) {
          sh.getRange(1, 1, data.length, w).setValues(data);
          sh.getRange(1, 1, 1, w).setFontWeight("bold");
          sh.setFrozenRows(1);
        }
        sh.setRightToLeft(true);
      });
      return out({ ok: 1, url: ss.getUrl() });
    }

    var sh2 = ss.getSheetByName("sync");
    if (!sh2) {
      sh2 = ss.insertSheet("sync");
      sh2.appendRow(["id", "su", "del", "json"]);
    }
    var data2 = sh2.getDataRange().getValues();
    var idx = {};
    for (var i = 1; i < data2.length; i++) idx[data2[i][0]] = i + 1;
    var now = Date.now();
    (b.items || []).forEach(function (it) {
      var row = [it.id, now, it.del ? 1 : 0, it.del ? "" : it.j];
      if (idx[it.id]) sh2.getRange(idx[it.id], 1, 1, 4).setValues([row]);
      else { sh2.appendRow(row); idx[it.id] = sh2.getLastRow(); }
    });
    var since = +b.since || 0;
    var rows2 = [];
    var d3 = sh2.getDataRange().getValues();
    for (var k = 1; k < d3.length; k++) {
      if (+d3[k][1] > since) rows2.push({ id: d3[k][0], su: +d3[k][1], del: !!d3[k][2], j: d3[k][3] });
    }
    return out({ ok: 1, now: now, rows: rows2, url: ss.getUrl() });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out({ ok: 1, msg: "CAFE ONE ready" }); }

// إذا أردت ربط الجدول بكلمة جديدة شغّل هذه الدالة مرة واحدة من المحرر.
function resetKey() { PropertiesService.getScriptProperties().deleteProperty("K"); }

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
