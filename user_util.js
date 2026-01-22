// =======================================
// user_util.gs
// =======================================

function normalizeEmail_(email) {
  return email ? String(email).trim().toLowerCase() : '';
}

function getActiveUserEmail_() {
  return normalizeEmail_(Session.getActiveUser().getEmail());
}

/**
 * スプレッドシートを取得する関数
 * IDではなく「URL」で開くように変更
 */
function getAppSpreadsheet() {
  // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
  // ↓↓↓ ここにスプレッドシートの「URL全体」を貼り付けてください ↓↓↓
  var url = 'https://docs.google.com/spreadsheets/d/1vQAZTBpBBnZtUpCSwg1lXlOvP7wqIwQFwrbT4A1xMTk/edit'; 
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  // URLを使って開く
  try {
    return SpreadsheetApp.openByUrl(url);
  } catch (e) {
    // URLで失敗した場合の予備策：このスクリプトがシートに紐付いているならこれで開く
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch (e2) {
      throw new Error('スプレッドシートを開けませんでした。URLが正しいか確認してください。\n詳細: ' + e);
    }
  }
}