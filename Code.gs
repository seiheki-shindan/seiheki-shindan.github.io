// 性癖タイプ診断:回答を受け取ってスプレッドシートに保存し、管理画面に返す Google Apps Script
// 設置手順は 設置手順.md を見てください。

const SHEET_NAME = "results";
const FETISH_IDS = ["voice", "scent", "kiss", "clothes", "bondage", "pain", "roleplay", "toys", "thrill", "multi"];
const HEADER = ["日時", "匿名ID", "タイプ", "LF", "WS", "PR", "TI"].concat(FETISH_IDS).concat(["回答(50問)"]);

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADER); sh.setFrozenRows(1); }
  return sh;
}

// 最初に1回だけ実行する。シートを用意し、管理者キーを自動で作って「管理者キー」シートに書き出す。
function setup() {
  sheet_();
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty("ADMIN_KEY");
  if (!key) { key = Utilities.getUuid().replace(/-/g, ""); props.setProperty("ADMIN_KEY", key); }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("管理者キー") || ss.insertSheet("管理者キー");
  sh.getRange("A1:B1").setValues([["管理者キー(admin.html で入力)", key]]);
}

// 診断ページから回答を受け取る。同じ匿名IDからの再回答は上書きする。
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!/^[A-Z]{4}$/.test(d.code || "")) return out_({ error: "bad request" });
    const num = (v) => (typeof v === "number" && isFinite(v) ? Math.max(0, Math.min(100, Math.round(v))) : "");
    const anon = String(d.anon || "").slice(0, 40).replace(/[^a-z0-9]/gi, "");
    const row = [new Date(), anon, d.code, num(d.pct && d.pct.LF), num(d.pct && d.pct.WS), num(d.pct && d.pct.PR), num(d.pct && d.pct.TI)]
      .concat(FETISH_IDS.map((k) => num(d.fet && d.fet[k])))
      .concat([JSON.stringify((d.answers || []).slice(0, 50).map(num))]);
    const sh = sheet_();
    if (anon) {
      const ids = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (ids[i][0] === anon) { sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return out_({ ok: true }); }
      }
    }
    sh.appendRow(row);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// 管理画面に結果を返す。スクリプトプロパティ ADMIN_KEY と一致するキーが必要。
function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  if (!key || !e.parameter.key || e.parameter.key !== key) return out_({ error: "unauthorized" });
  const values = sheet_().getDataRange().getValues().slice(1);
  const rows = values.map((r) => ({
    code: r[2],
    pct: { LF: r[3], WS: r[4], PR: r[5], TI: r[6] },
    fet: Object.fromEntries(FETISH_IDS.map((k, i) => [k, r[7 + i]])),
  }));
  return out_({ rows });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
