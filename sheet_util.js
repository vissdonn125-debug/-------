// =======================================
// sheet_util.gs — シート取得の共通ヘルパー
// =======================================

/**
 * アプリ用スプレッドシートから、指定したシート名のシートを取得する。
 * 見つからなければエラーにする。
 *
 * 使用例:
 *   var headerSheet = getSheet_(SHEET_NAMES.HEADER);
 *   var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
 *
 * @param {string} sheetName
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet_(sheetName) {
  var ss = getAppSpreadsheet();  // user_util.gs で定義済み
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }
  return sheet;
}

/**
 * 申請ヘッダシート取得用のショートカット（任意）
 */
function getHeaderSheet_() {
  return getSheet_(SHEET_NAMES.HEADER);
}

/**
 * 明細シート取得用のショートカット（任意）
 */
function getDetailSheet_() {
  return getSheet_(SHEET_NAMES.DETAIL);
}

/**
 * ユーザーマスタシート取得用のショートカット（任意）
 */
function getUserMasterSheet_() {
  return getSheet_(SHEET_NAMES.USER_MASTER);
}
