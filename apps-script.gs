// チャレンジ10 β版：結果をスプレッドシートに記録するスクリプト
// スプレッドシートの「拡張機能」→「Apps Script」に、この内容をすべて貼り付けて使います。

const SHEET_NAME = "結果";
const CATS = ["勉強", "お手伝い", "友達", "運動", "好きなこと", "生活"];
const HEADER = ["受信日時", "端末ID", "学年", "記録回数", "開始日", "終了日"]
  .concat(CATS, ["話なし", "アンケート"])
  .concat(Array.from({ length: 10 }, (_, i) => (i + 1) + "日目"));

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }

  const row = [new Date(), d.uid, d.grade, d.count, d.start, d.end]
    .concat(CATS.map(c => d.cats[c] || 0), [d.none, d.survey || ""])
    .concat(Array.from({ length: 10 }, (_, i) => d.days[i] || ""));

  // 同じ端末から2回送られた場合は、新しい内容で上書きする
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat() : [];
  const idx = ids.indexOf(d.uid);
  if (idx >= 0) {
    sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return ContentService.createTextOutput("ok");
}
