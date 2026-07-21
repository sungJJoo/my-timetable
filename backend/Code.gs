// 시간표 앱 백엔드: 이 스프레드시트를 DB로 사용 (Google Apps Script)
var KEY = "여기에-비밀키를-입력"; // 원하는 문자열로 바꾸고, 앱 동기화 설정에 같은 값을 입력하세요.

var EV_H = ["id", "day", "start", "end", "title", "place", "color", "authors", "until"];
var NT_H = ["id", "text", "ts", "authors"];

function doGet(e) {
  if (((e && e.parameter.key) || "") !== KEY) return json({ error: "bad key" });
  return json(readAll());
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  if ((body.key || "") !== KEY) return json({ error: "bad key" });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeAll(body.data);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

function sheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function readAll() {
  var ev = sheet("events", EV_H).getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] !== ""; })
    .map(function (r) {
      return { id: String(r[0]), day: +r[1], start: +r[2], end: +r[3], title: String(r[4]), place: String(r[5]), color: +r[6], authors: parseAuthors(r[7]), until: toDateStr(r[8]) };
    });
  var nt = sheet("notes", NT_H).getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] !== ""; })
    .map(function (r) {
      return { id: String(r[0]), text: String(r[1]), ts: +r[2], authors: parseAuthors(r[3]) };
    });
  return { events: ev, notes: nt };
}

function writeAll(data) {
  var es = sheet("events", EV_H);
  es.clearContents();
  es.appendRow(EV_H);
  if (data.events.length) {
    // "until"(YYYY-MM-DD) 열이 날짜로 자동 변환되지 않도록 서식을 텍스트로 고정
    es.getRange(2, EV_H.indexOf("until") + 1, data.events.length, 1).setNumberFormat("@");
    es.getRange(2, 1, data.events.length, EV_H.length).setValues(data.events.map(function (e) {
      return [e.id, e.day, e.start, e.end, e.title, e.place || "", e.color, (e.authors || []).join(","), e.until || ""];
    }));
  }
  var ns = sheet("notes", NT_H);
  ns.clearContents();
  ns.appendRow(NT_H);
  if (data.notes.length) {
    ns.getRange(2, 1, data.notes.length, NT_H.length).setValues(data.notes.map(function (n) {
      return [n.id, n.text, n.ts, (n.authors || []).join(",")];
    }));
  }
}

function parseAuthors(v) {
  return String(v || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean);
}

function toDateStr(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return v.getFullYear() + "-" + String(v.getMonth() + 1).padStart(2, "0") + "-" + String(v.getDate()).padStart(2, "0");
  }
  return String(v);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
