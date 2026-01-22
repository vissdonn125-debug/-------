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

    // 1. 科目マスタ取得
    var subjectList = [];
    var lastRowSub = subjectSheet.getLastRow();
    if (lastRowSub > 1) {
        subjectList = subjectSheet.getRange(2, 1, lastRowSub - 1, 1).getValues().flat().filter(String);
    }

    // 2. 申請中(APPLYING)のヘッダを取得
    var pendingList = [];
    var lastRowHead = headerSheet.getLastRow();
    if (lastRowHead <= 1) return { pendingList: [], subjectList: subjectList };

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
        return { pendingList: [], subjectList: subjectList };
    }

    // 3. 明細を取得して紐付け
    var lastRowDet = detailSheet.getLastRow();
    if (lastRowDet > 1) {
        var detValues = detailSheet.getRange(2, 1, lastRowDet - 1, DETAIL_COL.COL_COUNT).getValues();

        for (var j = 0; j < detValues.length; j++) {
            var dRow = detValues[j];
            var dAppId = dRow[DETAIL_COL.APPLICATION_ID - 1];

            if (appMap[dAppId]) {
                var rUrl = dRow[DETAIL_COL.RECEIPT_URL - 1];
                var fileId = '';
                if (rUrl) {
                    // URLからID抽出 (簡易的)
                    var match = rUrl.match(/id=([a-zA-Z0-9_-]+)/);
                    if (match) fileId = match[1];
                    // 複数ある場合は最初の1つ目のIDを取得（画像表示用）
                    if (!fileId && rUrl.indexOf('\n') !== -1) {
                        var firstUrl = rUrl.split('\n')[0];
                        var m2 = firstUrl.match(/id=([a-zA-Z0-9_-]+)/);
                        if (m2) fileId = m2[1];
                    }
                }

                appMap[dAppId].details.push({
                    detailId: dRow[DETAIL_COL.DETAIL_ID - 1],
                    usageDate: Utilities.formatDate(new Date(dRow[DETAIL_COL.USAGE_DATE - 1]), TIMEZONE, 'yyyy-MM-dd'),
                    amount: dRow[DETAIL_COL.AMOUNT - 1],
                    taxRate: dRow[DETAIL_COL.TAX_RATE - 1],
                    vendor: dRow[DETAIL_COL.VENDOR - 1],
                    subject: dRow[DETAIL_COL.SUBJECT - 1],
                    purpose: dRow[DETAIL_COL.PURPOSE - 1],
                    invoiceReg: dRow[DETAIL_COL.INVOICE_REG - 1] || '不明', // インボイス登録有無
                    hasImage: !!fileId,
                    fileId: fileId
                });
            }
        }
    }

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

    // ステータス更新
    var newStatus = (action === 'approve') ? STATUS.APPROVED : STATUS.REJECTED;
    var timestamp = new Date();

    // 更新するセル: [ステータス, 承認者Mail, 承認日時, 却下日時, 差し戻し日時, 差し戻しコメント]
    // HEADER_COL.STATUS(8) から HEADER_COL.RETURN_COMMENT(13) あたりを更新

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
        var lastRowDet = detailSheet.getLastRow();
        var detIds = detailSheet.getRange(2, DETAIL_COL.DETAIL_ID, lastRowDet - 1, 1).getValues().flat(); // ID列

        // 行特定用マップ
        var detRowMap = {};
        for (var k = 0; k < detIds.length; k++) {
            detRowMap[detIds[k]] = k + 2;
        }

        var totalAmount = 0;

        details.forEach(function (d) {
            var rIdx = detRowMap[d.detailId];
            if (rIdx) {
                var amt = Number(d.amount);
                totalAmount += amt;

                // 更新実行
                detailSheet.getRange(rIdx, DETAIL_COL.USAGE_DATE).setValue(new Date(d.usageDate));
                detailSheet.getRange(rIdx, DETAIL_COL.AMOUNT).setValue(amt);
                detailSheet.getRange(rIdx, DETAIL_COL.TAX_RATE).setValue(d.taxRate);
                detailSheet.getRange(rIdx, DETAIL_COL.VENDOR).setValue(d.vendor);
                detailSheet.getRange(rIdx, DETAIL_COL.SUBJECT).setValue(d.subject);
                detailSheet.getRange(rIdx, DETAIL_COL.PURPOSE).setValue(d.purpose);
                detailSheet.getRange(rIdx, DETAIL_COL.INVOICE_REG).setValue(d.invoiceReg);
            }
        });

        // ヘッダの合計金額も再計算して更新
        headerSheet.getRange(rowIndex, HEADER_COL.TOTAL_AMOUNT).setValue(totalAmount);
    }
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

    // 更新
    var timestamp = new Date();
    headerSheet.getRange(rowIndex, HEADER_COL.STATUS).setValue(STATUS.RETURNED);
    headerSheet.getRange(rowIndex, HEADER_COL.APPROVER_EMAIL).setValue(user.email);
    headerSheet.getRange(rowIndex, HEADER_COL.RETURNED_AT).setValue(timestamp);
    headerSheet.getRange(rowIndex, HEADER_COL.RETURN_COMMENT).setValue(comment);

    // 明細更新 (同様)
    if (details && details.length > 0) {
        var lastRowDet = detailSheet.getLastRow();
        var detIds = detailSheet.getRange(2, DETAIL_COL.DETAIL_ID, lastRowDet - 1, 1).getValues().flat();
        var detRowMap = {};
        for (var k = 0; k < detIds.length; k++) detRowMap[detIds[k]] = k + 2;
        var totalAmount = 0;
        details.forEach(function (d) {
            var rIdx = detRowMap[d.detailId];
            if (rIdx) {
                var amt = Number(d.amount);
                totalAmount += amt;
                detailSheet.getRange(rIdx, DETAIL_COL.USAGE_DATE).setValue(new Date(d.usageDate));
                detailSheet.getRange(rIdx, DETAIL_COL.AMOUNT).setValue(amt);
                detailSheet.getRange(rIdx, DETAIL_COL.TAX_RATE).setValue(d.taxRate);
                detailSheet.getRange(rIdx, DETAIL_COL.VENDOR).setValue(d.vendor);
                detailSheet.getRange(rIdx, DETAIL_COL.SUBJECT).setValue(d.subject);
                detailSheet.getRange(rIdx, DETAIL_COL.PURPOSE).setValue(d.purpose);
                detailSheet.getRange(rIdx, DETAIL_COL.INVOICE_REG).setValue(d.invoiceReg);
            }
        });
        headerSheet.getRange(rowIndex, HEADER_COL.TOTAL_AMOUNT).setValue(totalAmount);
    }

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
function api_addSubject(name) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);
    sheet.appendRow([name, 10, '']); // デフォルト税率10%

    // リストを返す
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(String);
}

/**
 * 科目マスタ削除
 */
function api_deleteSubject(name) {
    var user = getCurrentUserInfo();
    if (user.role !== 'ADMIN') throw new Error('権限がありません');

    var sheet = getSheet_(SHEET_NAMES.SUBJECT_MASTER);
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();

    // 後ろから削除（行ズレ防止）
    for (var i = values.length - 1; i >= 0; i--) {
        if (values[i] === name) {
            sheet.deleteRow(i + 2);
        }
    }

    var last = sheet.getLastRow();
    if (last <= 1) return [];
    return sheet.getRange(2, 1, last - 1, 1).getValues().flat().filter(String);
}
