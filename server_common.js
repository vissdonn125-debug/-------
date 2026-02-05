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

    // A:Name, B:Tax, C:Keywords, D:Branch
    // 安全のため getDataRange を使用してからマップする
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    if (numRows <= 1) return [];

    var values = sheet.getRange(2, 1, numRows - 1, sheet.getLastColumn()).getValues();

    return values.map(function (r) {
        return {
            name: r[0],
            taxRate: Number(r[1]) || 10,
            keywords: r[2] || '',
            branch: (r.length > 3 ? r[3] : '') || '' // ★追加: 拠点
        };
    }).filter(function (item) { return item.name; });
}

/**
 * ユーザー情報の正規化や共通ユーティリティ
 */

/**
 * ユーザーのEmailと拠点のマッピングを取得
 * Returns: { "email@example.com": "BranchName" }
 */
function getUserBranchMap_() {
    var userBranchMap = {};
    var sheetUser = getSheet_(SHEET_NAMES.USER_MASTER);
    if (!sheetUser) return userBranchMap;

    var lastRow = sheetUser.getLastRow();
    if (lastRow <= 1) return userBranchMap;

    // A:ID, B:Email, C:Name, D:Role, E:Manager, F:Branch
    var uVals = sheetUser.getRange(2, 2, lastRow - 1, 5).getValues(); // Email(B) 〜 Branch(F)
    uVals.forEach(function (r) {
        // r[0]=Email, r[4]=Branch (F列はB列から見て+4)
        var email = normalizeEmail_(r[0]);
        var branch = r[4] || '';
        if (email) userBranchMap[email] = branch;
    });
    return userBranchMap;
}

/**
 * 排他制御付きで関数を実行するヘルパー
 * @param {Function} callback - ロック取得後に実行する処理
 * @param {number} timeoutMs - ロック待機時間(ms) Default: 10000
 */
function runWithLock_(callback, timeoutMs) {
    var lock = LockService.getScriptLock();
    if (lock.tryLock(timeoutMs || 10000)) {
        try {
            return callback();
        } finally {
            lock.releaseLock();
        }
    } else {
        throw new Error('サーバーが混み合っています。もう一度やり直してください。');
    }
}
