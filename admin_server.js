// =======================================
// admin_server.gs
// =======================================

/**
 * 管理者ダッシュボード用データ取得
 */
function getAdminDashboardData() {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') {
        throw new Error('管理者権限がありません。');
    }

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
    var subjectSheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);

    // 1. 科目マスタ取得 (共通関数使用)
    var subjectList = getSubjectMasterWithTax_();

    // 2. 拠点リストを取得 (拠点マスタから)
    var branchList = [];
    var sheetBranch = getSheet_(SHEET_NAMES.BRANCH_MASTER);
    if (sheetBranch) {
        var lastRowB = sheetBranch.getLastRow();
        if (lastRowB > 1) {
            // B列(2列目)が拠点名
            var bVals = sheetBranch.getRange(2, BRANCH_MASTER_COL.NAME, lastRowB - 1, 1).getValues().flat();
            branchList = bVals.filter(function (b) { return b; }).sort();
        }
    }

    // 3. 申請中(APPLYING)のヘッダを取得
    var pendingList = [];
    var lastRowHead = headerSheet.getLastRow();
    if (lastRowHead <= 1) return { pendingList: [], subjectList: subjectList, branches: branchList };

    var headValues = headerSheet.getRange(2, 1, lastRowHead - 1, HEADER_COL.COL_COUNT).getValues();
    var appIds = [];
    var appMap = {};

    for (var i = 0; i < headValues.length; i++) {
        var row = headValues[i];
        var status = row[HEADER_COL.STATUS - 1];

        // 申請中のみ対象
        if (status === STATUS.APPLYING) {
            var appId = row[HEADER_COL.APPLICATION_ID - 1];
            var appDate = row[HEADER_COL.APPLICATION_DATE - 1];

            var appObj = {
                appId: appId,
                applicant: row[HEADER_COL.APPLICANT_NAME - 1],
                email: row[HEADER_COL.APPLICANT_EMAIL - 1],
                dept: row[HEADER_COL.APPLICANT_DEPT - 1],
                totalAmount: row[HEADER_COL.TOTAL_AMOUNT - 1],
                date: Utilities.formatDate(new Date(appDate), TIMEZONE, 'yyyy/MM/dd'),
                details: []
            };

            pendingList.push(appObj);
            appIds.push(appId);
            appMap[appId] = appObj;
        }
    }

    if (appIds.length === 0) {
        return { pendingList: [], subjectList: subjectList, branches: branchList };
    }

    // 3. 明細を取得して紐付け (共通関数利用)
    getDetailsForApps_(appIds, appMap);

    return {
        pendingList: pendingList,
        subjectList: subjectList
    };

    return {
        pendingList: pendingList,
        subjectList: subjectList
    };
}

/**
 * 申請の承認・却下実行 (明細編集含む)
 */
function processApplicationWithEdit(appId, action, details) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    // LockService: 排他制御
    runWithLock_(function () {
        // 1. ヘッダ更新
        var lastRowHead = headerSheet.getLastRow();
        var headData = headerSheet.getRange(2, 1, lastRowHead - 1, 1).getValues(); // ID列のみ取得
        var rowIndex = -1;

        for (var i = 0; i < headData.length; i++) {
            if (String(headData[i][0]) === String(appId)) {
                rowIndex = i + 2;
                break;
            }
        }

        if (rowIndex === -1) throw new Error('申請が見つかりません');

        // 支払済チェック
        var currentPayStatus = headerSheet.getRange(rowIndex, HEADER_COL.PAYMENT_STATUS).getValue();
        if (currentPayStatus === '支払済') {
            throw new Error('支払済みの申請は編集・承認・却下できません。支払管理画面で「未払い」に戻してから操作してください。');
        }

        // ステータス更新
        var newStatus = (action === 'approve') ? STATUS.APPROVED : STATUS.REJECTED;
        var timestamp = new Date();

        // 個別にセット
        headerSheet.getRange(rowIndex, HEADER_COL.STATUS).setValue(newStatus);
        headerSheet.getRange(rowIndex, HEADER_COL.APPROVER_EMAIL).setValue(user.email);

        if (action === 'approve') {
            headerSheet.getRange(rowIndex, HEADER_COL.APPROVED_AT).setValue(timestamp);
        } else {
            headerSheet.getRange(rowIndex, HEADER_COL.REJECTED_AT).setValue(timestamp);

            // 却下メール送信
            sendRejectionEmail_(rowIndex, appId, user.name);
            // チャット通知
            sendChatNotification_('【却下】申請ID: ' + appId + ' が却下されました。');
        }

        // 2. 明細更新 (details配列の内容でDetailシートを上書き)
        if (details && details.length > 0) {
            var dLastRow = detailSheet.getLastRow();
            if (dLastRow > 1) {
                var dFullRange = detailSheet.getRange(2, 1, dLastRow - 1, DETAIL_COL.COL_COUNT);
                var dValues = dFullRange.getValues();

                // ID列インデックス (0-based)
                var idIdx = DETAIL_COL.DETAIL_ID - 1;

                // 行特定用インデックスマップ (ID -> 0-based row index in dValues)
                var dMap = {};
                for (var k = 0; k < dValues.length; k++) {
                    dMap[String(dValues[k][idIdx])] = k;
                }

                var totalAmount = 0;
                var modifiedRows = [];

                details.forEach(function (d) {
                    var idx = dMap[String(d.detailId)];
                    if (idx !== undefined) {
                        var amt = Number(d.amount);
                        totalAmount += amt;

                        // メモリ上の配列を更新
                        dValues[idx][DETAIL_COL.USAGE_DATE - 1] = new Date(d.usageDate);
                        dValues[idx][DETAIL_COL.AMOUNT - 1] = amt;
                        dValues[idx][DETAIL_COL.TAX_RATE - 1] = d.taxRate;
                        dValues[idx][DETAIL_COL.VENDOR - 1] = d.vendor;
                        dValues[idx][DETAIL_COL.SUBJECT - 1] = d.subject;
                        dValues[idx][DETAIL_COL.PURPOSE - 1] = d.purpose;
                        dValues[idx][DETAIL_COL.INVOICE_REG - 1] = d.invoiceReg;
                    }
                });

                // シート全体を一括更新 (パフォーマンス向上)
                dFullRange.setValues(dValues);

                // ヘッダの合計金額も再計算して更新
                headerSheet.getRange(rowIndex, HEADER_COL.TOTAL_AMOUNT).setValue(totalAmount);
            }
        }

        // ★キャッシュ無効化
        if (rowIndex > 0) {
            var appDate = headerSheet.getRange(rowIndex, HEADER_COL.APPLICATION_DATE).getValue();
            invalidateSnapshots_(appDate);
        }
    });
}

/**
 * 差し戻し実行
 */
function processReturn(appId, comment, details) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    // processApplicationWithEdit とほぼ同様だがステータスが違う
    // 共通化できるが、シンプルに分けて記述

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    runWithLock_(function () {
        // ヘッダ検索
        var lastRowHead = headerSheet.getLastRow();
        var headData = headerSheet.getRange(2, 1, lastRowHead - 1, 1).getValues();
        var rowIndex = -1;
        for (var i = 0; i < headData.length; i++) {
            if (String(headData[i][0]) === String(appId)) {
                rowIndex = i + 2;
                break;
            }
        }
        if (rowIndex === -1) throw new Error('申請が見つかりません');

        // 支払済チェック
        var currentPayStatus = headerSheet.getRange(rowIndex, HEADER_COL.PAYMENT_STATUS).getValue();
        if (currentPayStatus === '支払済') {
            throw new Error('支払済みの申請は差し戻しできません。支払管理画面で「未払い」に戻してから操作してください。');
        }

        // 更新
        var timestamp = new Date();
        headerSheet.getRange(rowIndex, HEADER_COL.STATUS).setValue(STATUS.RETURNED);
        headerSheet.getRange(rowIndex, HEADER_COL.APPROVER_EMAIL).setValue(user.email);
        headerSheet.getRange(rowIndex, HEADER_COL.RETURNED_AT).setValue(timestamp);
        headerSheet.getRange(rowIndex, HEADER_COL.RETURN_COMMENT).setValue(comment);

        // 明細更新 (一括更新版)
        if (details && details.length > 0) {
            var dLastRow = detailSheet.getLastRow();
            if (dLastRow > 1) {
                var dFullRange = detailSheet.getRange(2, 1, dLastRow - 1, DETAIL_COL.COL_COUNT);
                var dValues = dFullRange.getValues();
                var idIdx = DETAIL_COL.DETAIL_ID - 1;
                var dMap = {};
                for (var k = 0; k < dValues.length; k++) dMap[String(dValues[k][idIdx])] = k;

                var totalAmount = 0;
                details.forEach(function (d) {
                    var idx = dMap[String(d.detailId)];
                    if (idx !== undefined) {
                        var amt = Number(d.amount);
                        totalAmount += amt;
                        dValues[idx][DETAIL_COL.USAGE_DATE - 1] = new Date(d.usageDate);
                        dValues[idx][DETAIL_COL.AMOUNT - 1] = amt;
                        dValues[idx][DETAIL_COL.TAX_RATE - 1] = d.taxRate;
                        dValues[idx][DETAIL_COL.VENDOR - 1] = d.vendor;
                        dValues[idx][DETAIL_COL.SUBJECT - 1] = d.subject;
                        dValues[idx][DETAIL_COL.PURPOSE - 1] = d.purpose;
                        dValues[idx][DETAIL_COL.INVOICE_REG - 1] = d.invoiceReg;
                    }
                });
                dFullRange.setValues(dValues);
                headerSheet.getRange(rowIndex, HEADER_COL.TOTAL_AMOUNT).setValue(totalAmount);
            }
        }

        // ★キャッシュ無効化
        if (rowIndex > 0) {
            var appDate = headerSheet.getRange(rowIndex, HEADER_COL.APPLICATION_DATE).getValue();
            invalidateSnapshots_(appDate);
        }
    });

    // 差し戻しメール送信
    sendReturnEmail_(rowIndex, appId, comment);
    // チャット通知
    sendChatNotification_('【差し戻し】申請ID: ' + appId + ' が差し戻されました。\n理由: ' + comment);
}

/**
 * 却下メール送信
 */
function sendRejectionEmail_(rowIndex, appId, adminName) {
    try {
        var sheet = getSheet_(SHEET_NAMES.HEADER);
        var applicantEmail = sheet.getRange(rowIndex, HEADER_COL.APPLICANT_EMAIL).getValue();

        if (applicantEmail) {
            MailApp.sendEmail({
                to: applicantEmail,
                subject: '【経費申請】申請が却下されました',
                body: adminName + ' 様により、以下の経費申請が却下されました。\n\n' +
                    '申請ID: ' + appId + '\n\n' +
                    'アプリを確認してください。'
            });
        }
    } catch (e) {
        Logger.log('Mail Fail: ' + e);
    }
}

/**
 * 差し戻しメール送信
 */
function sendReturnEmail_(rowIndex, appId, comment) {
    try {
        var sheet = getSheet_(SHEET_NAMES.HEADER);
        var applicantEmail = sheet.getRange(rowIndex, HEADER_COL.APPLICANT_EMAIL).getValue();

        if (applicantEmail) {
            MailApp.sendEmail({
                to: applicantEmail,
                subject: '【経費申請】申請が差し戻されました',
                body: '以下の経費申請が差し戻されました。\n\n' +
                    '申請ID: ' + appId + '\n' +
                    '理由: ' + comment + '\n\n' +
                    'アプリを確認し、修正して再申請してください。'
            });
        }
    } catch (e) {
        Logger.log('Mail Fail: ' + e);
    }
}

/**
 * チャットワーク/Google Chat等への通知 (Webhook)
 */
function sendChatNotification_(message) {
    try {
        // スクリプトプロパティ 'CHAT_WEBHOOK_URL' からURL取得
        var url = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
        if (!url) return;

        var payload = {
            text: message
        };

        UrlFetchApp.fetch(url, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload)
        });
    } catch (e) {
        Logger.log('Chat Notification Failed: ' + e);
    }
}

/**
 * 画像データ取得 (管理者用)
 */
function api_getReceiptImage(fileId) {
    if (!fileId) return null;
    try {
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var b64 = Utilities.base64Encode(blob.getBytes());
        var mime = blob.getContentType();
        return 'data:' + mime + ';base64,' + b64;
    } catch (e) {
        Logger.log('Image Fetch Error: ' + e);
        return null;
    }
}

/**
 * 科目マスタ追加
 */
function api_addSubject(name, taxRate, keywords, branch) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);

    runWithLock_(function () {
        // 重複チェック
        var list = sheet.getRange(2, 1, sheet.getLastRow() || 1, 1).getValues().flat().filter(String);
        if (list.includes(name)) throw new Error('既に存在する科目名です');

        sheet.appendRow([name, taxRate || 10, keywords || '', branch || '']);
    });

    // リストを返す(Object Array)
    return getSubjectListObject_(sheet);
}

/**
 * 科目マスタ更新
 */
function api_updateSubject(oldName, name, taxRate, keywords, branch) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);

    runWithLock_(function () {
        var data = sheet.getDataRange().getValues(); // 全データ取得
        // ヘッダ飛ばして検索
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === oldName) {
                // 更新
                sheet.getRange(i + 1, 1).setValue(name);
                sheet.getRange(i + 1, 2).setValue(taxRate);
                sheet.getRange(i + 1, 3).setValue(keywords);
                sheet.getRange(i + 1, 4).setValue(branch || '');
                break;
            }
        }
    });

    return getSubjectListObject_(sheet);
}

/**
 * 科目マスタ削除
 */
function api_deleteSubject(name) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);

    runWithLock_(function () {
        // getLastRow() check to prevent error if empty
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
            // 後ろから削除（行ズレ防止）
            for (var i = values.length - 1; i >= 0; i--) {
                if (values[i] === name) {
                    sheet.deleteRow(i + 2);
                }
            }
        }
    });

    return getSubjectListObject_(sheet);
}

function getSubjectListObject_(sheet) {
    var last = sheet.getLastRow();
    if (last <= 1) return [];
    var vals = sheet.getRange(2, 1, last - 1, 4).getValues(); // Name, Tax, Keywords, Branch
    return vals.map(function (r) {
        return { name: r[0], taxRate: r[1], keywords: r[2], branch: r[3] || '' };
    }).filter(function (o) { return !!o.name; });
}

/**
 * 月次レポート取得
 * targetMonth: 'yyyy-MM'
 */
/**
 * 月次レポート取得 (改修版)
 * targetMonth: 'yyyy-MM'
 * targetBranch: string (Optional)
 */
function api_getMonthlyReport(targetMonth, targetBranch) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    // 0. 前月繰越計算 (スナップショット利用)
    var carryOverBalance = getCarryOverWithSnapshot_(targetMonth, targetBranch);

    var data = getReportData_(targetMonth);

    // ユーザーマスタから拠点情報を取得してマップ化
    var userBranchMap = getUserBranchMap_();

    // 1. 集計用変数
    var totalIncome = 0;
    var totalExpense = 0;

    // インボイス集計: 全体 + 科目別
    // 構造: { total: { with: {10:0, 8:0}, ... }, bySubject: { "交通費": { with:..., without:... } } }
    var invoiceStats = {
        total: {
            withInvoice: { "10": 0, "8": 0, "other": 0 },
            withoutInvoice: { "10": 0, "8": 0, "other": 0 },
            unknown: { "10": 0, "8": 0, "other": 0 }
        },
        bySubject: {}
    };

    // 日別集計用
    var daysMap = {};
    var [y, m] = targetMonth.split('-');
    var lastDay = new Date(y, m, 0).getDate();

    // 初期化
    for (var d = 1; d <= lastDay; d++) {
        var dayStr = targetMonth + '-' + ('0' + d).slice(-2);
        daysMap[dayStr] = {
            income: 0,
            expense: 0,
            subjects: {}
        };
    }

    var subjects = new Set();

    data.details.forEach(function (det) {
        // 拠点フィルタ
        if (targetBranch) {
            var email = normalizeEmail_(det.applicantEmail);
            var uBranch = userBranchMap[email] || '';

            // 入金・資金移動系は、申請ヘッダの「部署」列(APPLICANT_DEPT)に拠点名を保存しているためそれを使う
            if (det.subject === '入金' || det.subject === '出金' || det.subject === '調整') {
                uBranch = det.applicantDept || '';
            }

            // マスタに無い場合などは空文字扱い。完全一致で判定
            if (uBranch !== targetBranch) return;
        }

        var amt = Number(det.amount);
        var sub = det.subject || '未分類';
        subjects.add(sub);
        var date = det.usageDate; // yyyy-MM-dd

        // 日別集計
        if (!daysMap[date]) daysMap[date] = { income: 0, expense: 0, subjects: {} }; // 念のため
        if (!daysMap[date].subjects[sub]) daysMap[date].subjects[sub] = 0;
        daysMap[date].subjects[sub] += amt;

        // 入出金判定
        if (sub === '入金' || sub === '仮払受入') {
            totalIncome += amt;
            daysMap[date].income += amt;
        } else {
            totalExpense += amt;
            daysMap[date].expense += amt;

            // インボイス集計 (支出のみ)
            var rate = String(det.taxRate || 10);
            if (rate !== '10' && rate !== '8') rate = 'other';

            var reg = det.invoiceReg || '不明';
            var cat = 'unknown';
            if (reg === 'あり') cat = 'withInvoice';
            else if (reg === 'なし') cat = 'withoutInvoice';

            // 全体集計
            invoiceStats.total[cat][rate] = (invoiceStats.total[cat][rate] || 0) + amt;

            // 科目別集計
            if (!invoiceStats.bySubject[sub]) {
                invoiceStats.bySubject[sub] = {
                    withInvoice: { "10": 0, "8": 0, "other": 0 },
                    withoutInvoice: { "10": 0, "8": 0, "other": 0 },
                    unknown: { "10": 0, "8": 0, "other": 0 }
                };
            }
            invoiceStats.bySubject[sub][cat][rate] += amt;
        }
    });

    // 日次残高計算 (running balance)
    var currentBalance = carryOverBalance; // 繰越からスタート
    var dayKeys = Object.keys(daysMap).sort();

    dayKeys.forEach(function (day) {
        var dObj = daysMap[day];
        var prev = currentBalance;
        currentBalance = prev + dObj.income - dObj.expense;
        dObj.balance = currentBalance;
        dObj.prevBalance = prev;
    });

    return {
        targetMonth: targetMonth,
        carryOver: carryOverBalance, // 繰越額
        totalBalance: currentBalance, // 最終残高
        scorecard: {
            income: totalIncome,
            expense: totalExpense,
            balance: totalIncome - totalExpense
        },
        grid: {
            days: dayKeys,
            subjects: Array.from(subjects).sort(),
            data: daysMap
        },
        invoice: invoiceStats
    };
}

/**
 * 日別(人別)レポート取得
 * targetDate: 'yyyy-MM' for monthly, 'yyyy-MM-dd' for specific date
 * targetBranch: string (Optional)
 * ※ 申請日(APPLICATION_DATE)でフィルタリング
 */
function api_getDailyReport(targetDate, targetBranch) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    // ヘッダ取得
    var hLast = headerSheet.getLastRow();
    if (hLast <= 1) return { list: [] };
    var hVals = headerSheet.getRange(2, 1, hLast - 1, HEADER_COL.COL_COUNT).getValues();

    var isMonthly = (targetDate.length === 7); // yyyy-MM

    // 拠点フィルタ用のマップ作成
    var userBranchMap = {};
    if (targetBranch) {
        userBranchMap = getUserBranchMap_();
    }

    // 対象の申請IDを収集 (申請日でフィルタ)
    var targetAppIds = {};
    hVals.forEach(function (row) {
        var status = row[HEADER_COL.STATUS - 1];
        if (status !== STATUS.APPROVED && status !== STATUS.FIXED) return;

        var appDateRaw = row[HEADER_COL.APPLICATION_DATE - 1];
        if (!appDateRaw) return;
        var appDateObj = new Date(appDateRaw);
        var appDaily = Utilities.formatDate(appDateObj, TIMEZONE, 'yyyy-MM-dd');
        var appMonthly = appDaily.substring(0, 7);

        var isMatch = false;
        if (isMonthly) {
            if (appMonthly === targetDate) isMatch = true;
        } else {
            if (appDaily === targetDate) isMatch = true;
        }

        if (isMatch) {
            targetAppIds[row[HEADER_COL.APPLICATION_ID - 1]] = {
                name: row[HEADER_COL.APPLICANT_NAME - 1],
                email: row[HEADER_COL.APPLICANT_EMAIL - 1]
            };
        }
    });

    // 明細取得
    var dLast = detailSheet.getLastRow();
    if (dLast <= 1) return { list: [] };
    var dVals = detailSheet.getRange(2, 1, dLast - 1, DETAIL_COL.COL_COUNT).getValues();

    var userMap = {};

    dVals.forEach(function (row) {
        var appId = row[DETAIL_COL.APPLICATION_ID - 1];
        if (!targetAppIds[appId]) return;

        var subject = row[DETAIL_COL.SUBJECT - 1];
        // 入金は除外して経費のみ集計
        if (subject === '入金' || subject === '仮払受入') return;

        var email = normalizeEmail_(targetAppIds[appId].email);
        // 拠点フィルタ
        if (targetBranch && userBranchMap[email] !== targetBranch) return;

        var name = targetAppIds[appId].name || '不明';
        if (!userMap[name]) userMap[name] = 0;
        userMap[name] += Number(row[DETAIL_COL.AMOUNT - 1]);
    });

    var list = Object.keys(userMap).map(function (name) {
        return { name: name, total: userMap[name] };
    }).sort(function (a, b) { return b.total - a.total; });

    return { list: list };
}

/**
 * 共通: レポート用データ収集 (指定月の承認済み・確定済みデータを収集)
 */
/**
 * 共通: レポート用データ収集 (指定月の承認済み・確定済みデータを収集)
 * targetDate: 'yyyy-MM' per month OR 'yyyy-MM-dd' per day
 */
function getReportData_(targetDate) {
    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    // ヘッダ取得
    var hLast = headerSheet.getLastRow();
    if (hLast <= 1) return { details: [] };
    var hVals = headerSheet.getRange(2, 1, hLast - 1, HEADER_COL.COL_COUNT).getValues();

    var targetAppIds = {}; // appId -> { applicantName }

    // 日付指定フォーマット判定
    var isMonthly = (targetDate.length === 7); // yyyy-MM

    hVals.forEach(function (row) {
        var status = row[HEADER_COL.STATUS - 1];
        // セキュリティ/精度強化: レポート対象を「承認済」または「経理確定」に限定する
        // (以前は APPLYING なども含めてしまっていたリスクを解消)
        if (status !== STATUS.APPROVED && status !== STATUS.FIXED) return;

        // ヘッダレベルでは対象候補としてIDをプールするだけにする（詳細は明細の日付で見る）
        targetAppIds[row[HEADER_COL.APPLICATION_ID - 1]] = {
            name: row[HEADER_COL.APPLICANT_NAME - 1],
            email: row[HEADER_COL.APPLICANT_EMAIL - 1],
            dept: row[HEADER_COL.APPLICANT_DEPT - 1]
        };
    });

    // 明細取得
    var dLast = detailSheet.getLastRow();
    if (dLast <= 1) return { details: [] };
    var dVals = detailSheet.getRange(2, 1, dLast - 1, DETAIL_COL.COL_COUNT).getValues();

    var validDetails = [];

    dVals.forEach(function (row) {
        var appId = row[DETAIL_COL.APPLICATION_ID - 1];
        if (!targetAppIds[appId]) return; // 対象外ヘッダ

        var uDate = row[DETAIL_COL.USAGE_DATE - 1];
        var uObj = new Date(uDate);
        var uDaily = Utilities.formatDate(uObj, TIMEZONE, 'yyyy-MM-dd');
        var uMonthly = uDaily.substring(0, 7);

        var isMatch = false;
        if (isMonthly) {
            if (uMonthly === targetDate) isMatch = true;
        } else {
            if (uDaily === targetDate) isMatch = true;
        }

        if (isMatch) {
            validDetails.push({
                usageDate: uDaily,
                amount: row[DETAIL_COL.AMOUNT - 1],
                subject: row[DETAIL_COL.SUBJECT - 1],
                taxRate: row[DETAIL_COL.TAX_RATE - 1],
                invoiceReg: row[DETAIL_COL.INVOICE_REG - 1],
                applicantName: targetAppIds[appId].name,
                applicantEmail: targetAppIds[appId].email,
                applicantDept: targetAppIds[appId].dept
            });
        }
    });

    return { details: validDetails };
}

/**
 * 入金登録 (管理者のみ)
 * date: 'yyyy-MM-dd'
 * amount: number
 * memo: string
 * type: string ('入金', '出金', '調整')
 * branch: string (Optional)
 */
function api_registerDeposit(date, amount, memo, type, branch) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
    var subject = type || '入金';

    // ID生成
    var dStr = date.replace(/-/g, '');
    var rand = Math.floor(Math.random() * 10000).toString();
    var appId = 'DEP-' + dStr + '-' + rand;

    var lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
        try {
            // ヘッダ登録 (HEADER_COL の定義に従う)
            var now = new Date();
            var hRow = [];
            hRow[HEADER_COL.APPLICATION_ID - 1] = appId;
            hRow[HEADER_COL.APPLICATION_DATE - 1] = date;
            hRow[HEADER_COL.APPLICANT_EMAIL - 1] = user.email;
            hRow[HEADER_COL.APPLICANT_NAME - 1] = '管理者(' + subject + ')';
            hRow[HEADER_COL.APPLICANT_DEPT - 1] = branch || '管理部'; // 拠点情報を部署列に流用
            hRow[HEADER_COL.TOTAL_AMOUNT - 1] = Number(amount);
            hRow[HEADER_COL.STATUS - 1] = STATUS.APPROVED;
            hRow[HEADER_COL.APPROVED_AT - 1] = Utilities.formatDate(now, TIMEZONE, 'yyyy/MM/dd HH:mm:ss');

            headerSheet.appendRow(hRow);

            // 明細登録 (DETAIL_COL の定義に従う)
            var detailId = 'D-' + Math.floor(Math.random() * 100000000);
            var dRow = [];
            dRow[DETAIL_COL.DETAIL_ID - 1] = detailId;
            dRow[DETAIL_COL.APPLICATION_ID - 1] = appId;
            dRow[DETAIL_COL.USAGE_DATE - 1] = date;
            dRow[DETAIL_COL.AMOUNT - 1] = Number(amount);
            dRow[DETAIL_COL.VENDOR - 1] = 'ー';
            dRow[DETAIL_COL.SUBJECT - 1] = subject;
            dRow[DETAIL_COL.PURPOSE - 1] = memo || (subject + '登録');
            dRow[DETAIL_COL.TAX_RATE - 1] = 0;
            dRow[DETAIL_COL.TAX_AMOUNT - 1] = 0;

            detailSheet.appendRow(dRow);

            // ★キャッシュ無効化
            invalidateSnapshots_(date);

        } finally {
            lock.releaseLock();
        }
    } else {
        throw new Error('サーバーが混み合っています。再試行してください。');
    }

    return { success: true, appId: appId };
}

/**
 * 共通: 特定月より前の全ての承認済みデータを取得 (繰越用)
 */
function getReportDataBefore_(targetMonth) {
    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    var hLast = headerSheet.getLastRow();
    if (hLast <= 1) return { details: [] };
    var hVals = headerSheet.getRange(2, 1, hLast - 1, HEADER_COL.COL_COUNT).getValues();

    var targetAppIds = {};
    hVals.forEach(function (row) {
        var status = row[HEADER_COL.STATUS - 1];
        if (status !== STATUS.APPROVED && status !== STATUS.FIXED) return;
        targetAppIds[row[HEADER_COL.APPLICATION_ID - 1]] = {
            name: row[HEADER_COL.APPLICANT_NAME - 1],
            email: row[HEADER_COL.APPLICANT_EMAIL - 1]
        };
    });

    var dLast = detailSheet.getLastRow();
    if (dLast <= 1) return { details: [] };
    var dVals = detailSheet.getRange(2, 1, dLast - 1, DETAIL_COL.COL_COUNT).getValues();

    var validDetails = [];
    dVals.forEach(function (row) {
        var appId = row[DETAIL_COL.APPLICATION_ID - 1];
        if (!targetAppIds[appId]) return;

        var uDate = row[DETAIL_COL.USAGE_DATE - 1];
        var uDaily = Utilities.formatDate(new Date(uDate), TIMEZONE, 'yyyy-MM-dd');
        var uMonthly = uDaily.substring(0, 7);

        // 対象月より「前」か判定
        if (uMonthly < targetMonth) {
            validDetails.push({
                amount: row[DETAIL_COL.AMOUNT - 1],
                subject: row[DETAIL_COL.SUBJECT - 1],
                applicantEmail: targetAppIds[appId].email
            });
        }
    });

    return { details: validDetails };
}

/**
 * 支払い確認画面用データ取得
 * 対象: 承認済み(APPROVED) または 却下(REJECTED)
 * startDate: 'yyyy-MM-dd' (Optional)
 * endDate: 'yyyy-MM-dd' (Optional)
 */
function api_getPaymentList(startDate, endDate) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var lastRow = headerSheet.getLastRow();
    if (lastRow <= 1) return [];

    var values = headerSheet.getDataRange().getValues();
    values.shift();

    var list = [];

    // Parse filtering dates
    var startObj = startDate ? new Date(startDate) : null;
    var endObj = endDate ? new Date(endDate) : null;
    if (startObj) startObj.setHours(0, 0, 0, 0);
    if (endObj) endObj.setHours(23, 59, 59, 999);

    for (var i = 0; i < values.length; i++) {
        var row = values[i];
        var status = String(row[HEADER_COL.STATUS - 1] || '').trim();

        if (status === STATUS.APPLYING || status === STATUS.DRAFT || status === '') continue;

        var dateObj = row[HEADER_COL.APPROVED_AT - 1] || row[HEADER_COL.REJECTED_AT - 1] || row[HEADER_COL.RETURNED_AT - 1] || row[HEADER_COL.APPLICATION_DATE - 1];

        if (!dateObj) continue;

        var d = (dateObj instanceof Date) ? dateObj : new Date(dateObj);
        if (isNaN(d.getTime())) continue;

        // Date range filtering
        if (startObj && d < startObj) continue;
        if (endObj && d > endObj) continue;

        var payStatus = (row.length >= HEADER_COL.PAYMENT_STATUS) ? row[HEADER_COL.PAYMENT_STATUS - 1] : '';
        if (!payStatus) payStatus = '未払い';

        if (status === STATUS.REJECTED) {
            // 却下の場合は支払対象外だが履歴には出す
            payStatus = '-';
        }

        list.push({
            appId: row[HEADER_COL.APPLICATION_ID - 1],
            applicant: row[HEADER_COL.APPLICANT_NAME - 1],
            totalAmount: row[HEADER_COL.TOTAL_AMOUNT - 1],
            status: status,
            approvedAt: Utilities.formatDate(new Date(dateObj), TIMEZONE, 'yyyy/MM/dd'),
            paymentStatus: payStatus,
            paymentStatus: payStatus,
            dept: row[HEADER_COL.APPLICANT_DEPT - 1] || '', // ★追加: 部署(拠点)
            details: [] // 必要なら詳細取得ロジックを追加(HeavyになるのでOnDemand推奨)
        });
    }


    var appMap = {};
    for (var j = 0; j < list.length; j++) {
        appMap[list[j].appId] = list[j];
    }
    var appIds = list.map(function (a) { return a.appId; });

    // 詳細を取得・紐付け
    getDetailsForApps_(appIds, appMap);

    // 新しい順
    return list.reverse();
}

/**
 * アプリケーションIDのリストに対して詳細データを取得・紐付けする共通関数
 */
function getDetailsForApps_(appIds, appMap) {
    if (!appIds || appIds.length === 0) return;

    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
    var lastRowDet = detailSheet.getLastRow();
    if (lastRowDet <= 1) return;

    // データ量が多い場合はフィルタリングすべきだが、現状は全取得してメモリ上でマッチング
    // (GASの制限内で動作する前提)
    var detValues = detailSheet.getRange(2, 1, lastRowDet - 1, DETAIL_COL.COL_COUNT).getValues();

    for (var j = 0; j < detValues.length; j++) {
        var dRow = detValues[j];
        var dAppId = dRow[DETAIL_COL.APPLICATION_ID - 1];

        if (appMap[dAppId]) { // 対象の申請のみ処理
            var rUrl = dRow[DETAIL_COL.RECEIPT_URL - 1];
            var fileId = '';
            if (rUrl) {
                // URLからID抽出 (互換性向上)
                var idMatch = rUrl.match(/id=([a-zA-Z0-9_-]+)/) || rUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
                if (idMatch) fileId = idMatch[1];

                // 複数ある場合は最初の1つ目のIDを取得（画像表示用）
                if (!fileId && rUrl.indexOf('\n') !== -1) {
                    var firstUrl = rUrl.split('\n')[0];
                    var m2 = firstUrl.match(/id=([a-zA-Z0-9_-]+)/) || firstUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
                    if (m2) fileId = m2[1];
                }
            }

            // details配列が無ければ初期化 (念のため)
            if (!appMap[dAppId].details) appMap[dAppId].details = [];

            var usageDateStr = '';
            try {
                var dVal = dRow[DETAIL_COL.USAGE_DATE - 1];
                if (dVal) {
                    var dObj = (dVal instanceof Date) ? dVal : new Date(dVal);
                    if (!isNaN(dObj.getTime())) {
                        usageDateStr = Utilities.formatDate(dObj, TIMEZONE, 'yyyy-MM-dd');
                    }
                }
            } catch (e) {
                // date parse error, ignore
            }

            appMap[dAppId].details.push({
                detailId: dRow[DETAIL_COL.DETAIL_ID - 1],
                usageDate: usageDateStr, // 失敗時は空文字
                amount: dRow[DETAIL_COL.AMOUNT - 1],
                taxRate: dRow[DETAIL_COL.TAX_RATE - 1],
                vendor: dRow[DETAIL_COL.VENDOR - 1],
                subject: dRow[DETAIL_COL.SUBJECT - 1],
                purpose: dRow[DETAIL_COL.PURPOSE - 1],
                invoiceReg: dRow[DETAIL_COL.INVOICE_REG - 1] || '不明',
                hasImage: !!fileId,
                fileId: fileId
            });
        }
    }
}

/**
 * 支払いステータス更新
 * appIds: string[]
 * status: '支払済' | '未払い'
 */
function api_updatePaymentStatus(appIds, newStatus) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.HEADER);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { count: 0 };

    var lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
        try {
            // A列(ID)取得
            var ids = sheet.getRange(2, HEADER_COL.APPLICATION_ID, lastRow - 1, 1).getValues().flat().map(String);
            var targetSet = {};
            appIds.forEach(function (id) { targetSet[id] = true; });
            var count = 0; // Initialize count
            var minDate = null; // キャッシュ無効化用の最小日付

            for (var i = 0; i < ids.length; i++) {
                if (targetSet[ids[i]]) {
                    var r = i + 2; // 行番号
                    // ステータス書き込み
                    // PAYMENT_STATUS カラムが存在するか確認が必要だが、前回のFixでsetup.jsに追加済み
                    // 万が一列が足りない場合、getHeaderSheet側で自動追加はしていないので、
                    // ここでsetValueするとエラーになる可能性があるが、setupScriptProperties等で整備前提。
                    sheet.getRange(r, HEADER_COL.PAYMENT_STATUS).setValue(newStatus);
                    count++;

                    // 日付を取得してキャッシュ無効化対象を特定
                    var appDate = sheet.getRange(r, HEADER_COL.APPLICATION_DATE).getValue();
                    if (minDate === null || appDate < minDate) minDate = appDate;
                }
            }

            // まとめて無効化
            if (minDate) invalidateSnapshots_(Utilities.formatDate(new Date(minDate), TIMEZONE, 'yyyy-MM-dd'));

        } finally {
            lock.releaseLock();
        }
    } else {
        throw new Error('サーバーが混み合っています。');
    }
    return { count: count };
}

/**
 * 申請の詳細(明細)を取得するAPI (History用)
 * Totalクリックで内訳表示用
 */
function api_getAppDetailsForModal(appId) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var result = api_getApplication(appId); // app_server.gsの既存関数を流用
    return result.items || [];
}

/**
 * 画像データ一括取得
 * fileIds: string[]
 * return: { id: base64, ... }
 */
function api_getImagesBatch(fileIds) {
    if (!fileIds || !Array.isArray(fileIds) || fileIds.length === 0) return {};

    var result = {};
    // ユニーク化
    var uniqueIds = fileIds.filter(function (x, i, self) {
        return self.indexOf(x) === i && x;
    });

    uniqueIds.forEach(function (id) {
        try {
            var file = DriveApp.getFileById(id);
            var blob = file.getBlob();
            var b64 = Utilities.base64Encode(blob.getBytes());
            var mime = blob.getContentType();
            result[id] = 'data:' + mime + ';base64,' + b64;
        } catch (e) {
            Logger.log('Image Batch Error (' + id + '): ' + e);
            result[id] = null; // エラー時はnullか空文字
        }
    });
    return result;
}


