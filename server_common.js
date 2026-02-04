// =======================================
// server_common.gs — 共通ロジック
// =======================================

/**
 * 科目マスタ（税率情報付き）を取得する共通関数
 * 戻り値: { name: string, taxRate: number, keywords: string }[]
 */
function getSubjectMasterWithTax_() {
    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) return [];

    // A列:Name, B列:Tax, C列:Keywords
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    return values.map(function (r) {
        return {
            name: r[0],
            taxRate: Number(r[1]) || 10,
            keywords: r[2] || ''
        };
    }).filter(function (item) { return item.name; });
}

/**
 * ユーザー情報の正規化や共通ユーティリティもここに移動可能ですが、
 * まずは緊急度の高いマスタ取得のみ共通化します。
 */
