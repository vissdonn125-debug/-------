
/**
 * スナップショットシートの初期化・取得
 */
function ensureSnapshotSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.MONTHLY_SNAPSHOT);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.MONTHLY_SNAPSHOT);
        sheet.hideSheet(); // ユーザーからは隠す
        // Header: Month, Branch (Email/Name), Balance, UpdatedAt
        sheet.appendRow(['Month', 'BranchEmail', 'EndingBalance', 'UpdatedAt']);
    }
    return sheet;
}

/**
 * 指定日以降のスナップショットを無効化（削除）
 * データ更新時に呼び出すこと
 * @param {Date|string} dateObj 変更があった日付
 */
function invalidateSnapshots_(dateObj) {
    var sheet = ensureSnapshotSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    var changedDate = new Date(dateObj);
    var changedMonthStr = Utilities.formatDate(changedDate, TIMEZONE, 'yyyy-MM');

    // 行削除は後ろから
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A列: Month
    var rowsToDelete = [];

    for (var i = 0; i < data.length; i++) {
        var rowMonth = data[i][0];
        if (rowMonth >= changedMonthStr) {
            // 行番号は i + 2
            rowsToDelete.push(i + 2);
        }
    }

    // まとめて削除はできないので、後ろから順に削除
    rowsToDelete.reverse().forEach(function (r) {
        sheet.deleteRow(r);
    });
}

/**
 * 指定月のスナップショットを取得・なければ計算して保存
 * @param {string} targetMonth 'yyyy-MM'
 * @param {string} branchFilter (Optional)
 * @return {number} 前月までの繰越残高
 */
function getCarryOverWithSnapshot_(targetMonth, branchFilter) {
    var sheet = ensureSnapshotSheet_();

    // yyyy-MM の '前月' を計算
    var targetDate = new Date(targetMonth + '-01');
    var prevDate = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
    var prevMonthStr = Utilities.formatDate(prevDate, TIMEZONE, 'yyyy-MM');

    // 1. スナップショットの検索 (Read)
    var branchKey = branchFilter || 'ALL';
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
        if (data[i][0] === prevMonthStr && data[i][1] === branchKey) {
            return Number(data[i][2]);
        }
    }

    // 2. なければ計算 (Calculate)
    var calculatedBalance = calculateBalanceUpTo_(prevMonthStr, branchFilter);

    // 3. 保存 (Write with Lock)
    var lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) {
        try {
            // ダブルチェック: ロック取得中に誰かが書いたかもしれない
            var checkData = sheet.getDataRange().getValues();
            for (var j = 1; j < checkData.length; j++) {
                if (checkData[j][0] === prevMonthStr && checkData[j][1] === branchKey) {
                    return Number(checkData[j][2]);
                }
            }
            sheet.appendRow([prevMonthStr, branchKey, calculatedBalance, new Date()]);
        } finally {
            lock.releaseLock();
        }
    }

    return calculatedBalance;
}

/**
 * (旧ロジック改良) 指定月 *まで* の累積残高を計算
 * @param {string} endMonthStr 'yyyy-MM' (この月を含む)
 * @param {string} branchFilter
 */
function calculateBalanceUpTo_(endMonthStr, branchFilter) {
    var coBranchMap = {};
    if (branchFilter && branchFilter !== 'ALL') {
        var sheetUser = getSheet_(SHEET_NAMES.USER_MASTER);
        if (sheetUser) {
            var lastRowU = sheetUser.getLastRow();
            if (lastRowU > 1) {
                var uVals = sheetUser.getRange(2, 2, lastRowU - 1, 5).getValues(); // B:Email, F:Branch
                uVals.forEach(function (r) {
                    coBranchMap[normalizeEmail_(r[0])] = r[4] || '';
                });
            }
        }
    }

    var d = new Date(endMonthStr + '-01');
    var nextD = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    var nextMonthStr = Utilities.formatDate(nextD, TIMEZONE, 'yyyy-MM');

    var result = getReportDataBefore_(nextMonthStr);

    var income = 0;
    var expense = 0;

    result.details.forEach(function (det) {
        if (branchFilter && branchFilter !== 'ALL') {
            var email = normalizeEmail_(det.applicantEmail);
            if (coBranchMap[email] !== branchFilter) return;
        }

        var amt = Number(det.amount);
        // 定数利用による安全性向上
        if (det.subject === TRANSACTION_TYPE.DEPOSIT || det.subject === TRANSACTION_TYPE.PROVISIONAL) {
            income += amt;
        } else {
            expense += amt;
        }
    });

    return income - expense;
}
