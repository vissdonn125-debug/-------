// =======================================
// status_util.gs — ステータス周りの共通ヘルパー
// =======================================
//
// 前提：app_constants.gs で STATUS が定義されていること。
//   var STATUS = {
//     DRAFT:    '下書き',
//     APPLYING: '申請中',
//     APPROVED: '承認済',
//     REJECTED: '却下',
//     FIXED:    '経理確定'
//   };
//
// シート上には日本語ラベル（"申請中" など）が保存されている想定で、
// ロジックでは "APPLYING" などのキーで判定したいので、
// その橋渡しをするユーティリティです。
// =======================================

/**
 * 日本語ラベルからステータスキーを取得する
 *
 * 例:
 *   getStatusKeyByLabel('申請中')   -> 'APPLYING'
 *   getStatusKeyByLabel('承認済')   -> 'APPROVED'
 *   存在しないラベルの場合は null を返す
 *
 * @param {string} label  シートに入っている日本語ラベル
 * @return {string|null}  STATUS のキー (DRAFT / APPLYING / APPROVED / REJECTED / FIXED) または null
 */
function getStatusKeyByLabel(label) {
  if (!label) return null;
  var target = String(label).trim();
  for (var key in STATUS) {
    if (!STATUS.hasOwnProperty(key)) continue;
    if (STATUS[key] === target) {
      return key;
    }
  }
  return null;
}

/**
 * ステータスキーから日本語ラベルを取得する（おまけ）
 *
 * 例:
 *   getStatusLabelByKey('APPLYING') -> '申請中'
 *
 * @param {string} key
 * @return {string|null}
 */
function getStatusLabelByKey(key) {
  if (!key) return null;
  if (STATUS.hasOwnProperty(key)) {
    return STATUS[key];
  }
  return null;
}
