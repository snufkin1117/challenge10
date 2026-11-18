// チャレンジ10 β版：結果をスプレッドシートに記録するスクリプト
// スプレッドシートの「拡張機能」→「Apps Script」に、この内容をすべて貼り付けて使います。

const SHEET_NAME = "結果";
const CATS = ["勉強", "お手伝い", "友達", "運動", "好きなこと", "生活"];
const HEADER = ["受信日時", "端末ID", "学年", "記録回数", "開始日", "終了日"]
  .concat(CATS, ["話なし", "アンケート"])
  .concat(Array.from({ length: 10 }, (_, i) => (i + 1) + "日目"), ["周目"]);
const ROUND_COL = HEADER.length; // 「周目」の列（いちばん右）

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, ROUND_COL).getValue() !== "周目") {
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]); // 古い見出しに「周目」を足す
  }
  const round = Number(d.round) || 1;

  const row = [new Date(), d.uid, d.grade, d.count, d.start, d.end]
    .concat(CATS.map(c => d.cats[c] || 0), [d.none, d.survey || ""])
    .concat(Array.from({ length: 10 }, (_, i) => d.days[i] || ""), [round]);

  // 同じ端末・同じ周目から2回送られた場合は、新しい内容で上書きする
  // （おかわりで次の周に進んだら、別の行として追加する）
  const n = sh.getLastRow() - 1;
  const keys = n > 0 ? sh.getRange(2, 1, n, ROUND_COL).getValues().map(r => r[1] + "#" + (Number(r[ROUND_COL - 1]) || 1)) : [];
  const idx = keys.indexOf(d.uid + "#" + round);
  if (idx >= 0) {
    sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  CacheService.getScriptCache().remove("total");
  return ContentService.createTextOutput("ok");
}

// アプリのホーム画面に出す「みんなのチャレンジ合計回数」を返す
// （「記録回数」の列をすべて足した数。1分間はキャッシュして負荷を減らす）
function doGet() {
  const cache = CacheService.getScriptCache();
  let total = cache.get("total");
  if (total === null) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const n = sh && sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0;
    total = n ? sh.getRange(2, 4, n, 1).getValues().flat().reduce((a, v) => a + (Number(v) || 0), 0) : 0;
    cache.put("total", String(total), 60);
  }
  return ContentService.createTextOutput(JSON.stringify({ total: Number(total) }))
    .setMimeType(ContentService.MimeType.JSON);
}
