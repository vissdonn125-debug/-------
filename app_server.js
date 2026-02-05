// =======================================
// app_server.gs — カート対応版
// =======================================

function api_listMyApplications(targetMonth) {
  var sheet = getSheet_(SHEET_NAMES.HEADER);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, HEADER_COL.COL_COUNT).getValues();
  var myEmail = getActiveUserEmail_();

  var list = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[HEADER_COL.APPLICATION_ID - 1]) continue;

    var rowEmail = normalizeEmail_(row[HEADER_COL.APPLICANT_EMAIL - 1]);
    if (rowEmail !== myEmail) continue;

    var appDate = row[HEADER_COL.APPLICATION_DATE - 1];
    var appDateObj = (appDate instanceof Date) ? appDate : new Date(appDate);
    var appDateIso = Utilities.formatDate(appDateObj, TIMEZONE, 'yyyy-MM-dd');
    var appMonth = Utilities.formatDate(appDateObj, TIMEZONE, 'yyyy-MM');

    if (targetMonth && targetMonth !== '' && appMonth !== targetMonth) continue;

    list.push({
      applicationId: row[HEADER_COL.APPLICATION_ID - 1],
      applicationDate: appDateIso,
      displayDate: Utilities.formatDate(appDateObj, TIMEZONE, 'MM/dd'),
      applicantName: row[HEADER_COL.APPLICANT_NAME - 1],
      totalAmount: row[HEADER_COL.TOTAL_AMOUNT - 1],
      statusLabel: row[HEADER_COL.STATUS - 1]
    });
  }
  return list.reverse();
}

/**
 * 自分の申請履歴取得 (月指定 & 画像つき)
 */
function api_getMyHistory(targetMonth) {
  var user = getCurrentUserInfo();
  var headerSheet = getSheet_(SHEET_NAMES.HEADER);
  var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
  var lastRow = headerSheet.getLastRow();

  if (lastRow <= 1) return [];

  // デフォルト今月
  if (!targetMonth) {
    targetMonth = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM');
  }

  var values = headerSheet.getRange(2, 1, lastRow - 1, HEADER_COL.COL_COUNT).getValues();
  var myApps = [];
  var appIdList = [];

  // 1. ヘッダ取得 & フィルタ
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    // メールアドレスで判定
    if (normalizeEmail_(row[HEADER_COL.APPLICANT_EMAIL - 1]) === normalizeEmail_(user.email)) {
      var dObj = new Date(row[HEADER_COL.APPLICATION_DATE - 1]);
      var dStr = Utilities.formatDate(dObj, TIMEZONE, 'yyyy-MM-dd');
      var mStr = dStr.substring(0, 7);

      if (targetMonth !== 'ALL' && mStr !== targetMonth) continue;

      var aid = row[HEADER_COL.APPLICATION_ID - 1];
      myApps.push({
        applicationId: aid,
        displayDate: dStr,
        totalAmount: row[HEADER_COL.TOTAL_AMOUNT - 1],
        statusLabel: row[HEADER_COL.STATUS - 1],
        dept: row[HEADER_COL.APPLICANT_DEPT - 1] || '',
        firstImageId: null // 後続処理で取得
      });
      appIdList.push(aid);
    }
  }

  if (myApps.length === 0) return [];

  // 2. 明細から画像取得 (該当アプリIDのみ対象)
  var dLast = detailSheet.getLastRow();
  if (dLast > 1) {
    var dVals = detailSheet.getRange(2, 1, dLast - 1, DETAIL_COL.COL_COUNT).getValues();
    var appMap = {};
    myApps.forEach(function (a) { appMap[a.applicationId] = a; });

    var aidsSet = {}; // Setの代わりにObjectを使用(GAS互換)
    appIdList.forEach(function (id) { aidsSet[id] = true; });

    for (var j = 0; j < dVals.length; j++) {
      var rAid = dVals[j][DETAIL_COL.APPLICATION_ID - 1];
      // まだ画像が見つかっていない、かつ対象の申請IDである場合
      if (aidsSet[rAid] && appMap[rAid] && !appMap[rAid].firstImageId) {
        var rUrl = dVals[j][DETAIL_COL.RECEIPT_URL - 1];
        if (rUrl) {
          var m = rUrl.match(/id=([a-zA-Z0-9_-]+)/) || rUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (m) appMap[rAid].firstImageId = m[1];
        }
      }
    }
  }

  return myApps.sort(function (a, b) {
    return new Date(b.displayDate) - new Date(a.displayDate); // 新しい順
  });
}

/**
 * カート申請作成 (複数明細) - API
 * data: { applicantId, items: [ { amount, vendor, subject, receiptUrl, ... } ] }
 */
function api_submitExpense(data) {
  // 1. ユーザー特定
  var currentUser = getCurrentUserInfo();
  var userInfo = null;

  // セキュリティ強化: 管理者でない場合、applicantId の指定を無視して自分自身として申請する
  if (currentUser.role === ROLES.ADMIN && data.applicantId) {
    userInfo = findUserById_(data.applicantId);
  }

  // 管理者以外、または該当ユーザーが見つからない場合は自分自身
  if (!userInfo) {
    userInfo = currentUser;
  }

  // LockService: 排他制御
  return runWithLock_(function () {
    // 2. 準備
    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);
    var now = new Date();
    var appDate = now;
    var ymd = Utilities.formatDate(appDate, TIMEZONE, 'yyyyMMdd');
    var appId = 'APP-' + ymd + '-' + now.getTime();

    // 3. 明細保存 & 合計計算
    var items = data.items || [];
    var totalAmount = 0;

    items.forEach(function (item) {
      var amount = Number(item.amount || 0);
      totalAmount += amount;

      // receiptUrlが配列の場合は改行区切りで結合して保存
      var receiptUrl = item.receiptUrl || '';
      if (Array.isArray(receiptUrl)) {
        receiptUrl = receiptUrl.join('\n');
      }

      var taxRate = Number(item.taxRate) || 10;

      // 利用日を取得（item.usageDateがあればそれを使用、なければappDate）
      var usageDate = appDate;
      if (item.usageDate) {
        try {
          usageDate = new Date(item.usageDate);
        } catch (e) {
          usageDate = appDate;
        }
      }

      var detailRow = [];
      detailRow[DETAIL_COL.DETAIL_ID - 1] = 'DET-' + now.getTime() + '-' + Math.floor(Math.random() * 10000);
      detailRow[DETAIL_COL.APPLICATION_ID - 1] = appId;
      detailRow[DETAIL_COL.USAGE_DATE - 1] = usageDate;
      detailRow[DETAIL_COL.AMOUNT - 1] = amount;
      detailRow[DETAIL_COL.TAX_RATE - 1] = taxRate;
      detailRow[DETAIL_COL.TAX_AMOUNT - 1] = 0;
      detailRow[DETAIL_COL.VENDOR - 1] = item.vendor || '';
      detailRow[DETAIL_COL.SUBJECT - 1] = item.subject || '';
      detailRow[DETAIL_COL.PAYMENT_METHOD - 1] = item.paymentMethod || '';
      detailRow[DETAIL_COL.PURPOSE - 1] = item.purpose || '';
      detailRow[DETAIL_COL.RECEIPT_URL - 1] = receiptUrl;
      detailRow[DETAIL_COL.OCR_SCORE - 1] = item.ocrScore || '';
      detailRow[DETAIL_COL.INVOICE_REG - 1] = item.invoiceReg || '不明';

      detailSheet.appendRow(detailRow);
    });

    // 4. ヘッダー保存
    var lastRow = headerSheet.getLastRow();
    var serial = lastRow; // 簡易

    var headerRow = [];
    headerRow[HEADER_COL.APPLICATION_ID - 1] = appId;
    headerRow[HEADER_COL.APPLICATION_DATE - 1] = appDate;
    headerRow[HEADER_COL.APPLICANT_EMAIL - 1] = userInfo.email;
    headerRow[HEADER_COL.APPLICANT_NAME - 1] = userInfo.name;
    headerRow[HEADER_COL.APPLICANT_DEPT - 1] = userInfo.dept || '';
    headerRow[HEADER_COL.SERIAL_NO - 1] = serial;
    headerRow[HEADER_COL.TOTAL_AMOUNT - 1] = totalAmount; // 自動計算した合計
    headerRow[HEADER_COL.STATUS - 1] = STATUS.APPLYING;
    headerRow[HEADER_COL.APPROVER_EMAIL - 1] = userInfo.approverEmail;
    headerRow[HEADER_COL.NEEDS_CHECK - 1] = false;

    headerSheet.appendRow(headerRow);

    return { applicationId: appId };
  });
}

/**
 * 申請詳細取得 (編集用)
 */
function api_getApplication(appId) {
  var headerSheet = getSheet_(SHEET_NAMES.HEADER);
  var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

  // ヘッダ情報取得
  var hValues = headerSheet.getDataRange().getValues();
  var header = null;
  for (var i = 1; i < hValues.length; i++) {
    if (String(hValues[i][HEADER_COL.APPLICATION_ID - 1]) === String(appId)) {
      var row = hValues[i];
      header = {
        appId: row[HEADER_COL.APPLICATION_ID - 1],
        status: row[HEADER_COL.STATUS - 1],
        email: row[HEADER_COL.APPLICANT_EMAIL - 1]
      };
      break;
    }
  }

  if (!header) throw new Error('申請が見つかりません');
  // 権限チェック (自身の申請か確認)
  if (normalizeEmail_(header.email) !== getActiveUserEmail_()) throw new Error('権限がありません');
  // ステータスチェック (承認済などは編集不可)
  if (header.status === STATUS.APPROVED || header.status === STATUS.FIXED) throw new Error('承認済みのため編集できません');

  // 明細取得
  var dValues = detailSheet.getDataRange().getValues();
  var items = [];
  for (var j = 1; j < dValues.length; j++) {
    var dRow = dValues[j];
    if (String(dRow[DETAIL_COL.APPLICATION_ID - 1]) === String(appId)) {
      var rUrl = dRow[DETAIL_COL.RECEIPT_URL - 1] || '';
      var imgData = '';
      // ここでは画像バイナリまでは返さず、URLのみ返す（Frontendで解析済みステータスとして扱うため、画像表示はURLベースか、再取得が必要）
      // ※簡易化のため、receiptUrlをそのまま返す。Frontendで `imgData` として扱うには工夫が必要だが、
      // 編集時は「解析完了」状態で復元し、画像はURLリンク等の扱いに留めるか、
      // あるいは api_getReceiptImage で各画像をBase64化して返すか。
      // Base64化は重いので、URLリストを返し、Frontendで「読込済」として表示するのみとする。

      var receiptUrls = rUrl.split('\n').filter(String);

      items.push({
        id: Date.now() + Math.random(), // Frontend用ID
        amount: dRow[DETAIL_COL.AMOUNT - 1],
        vendor: dRow[DETAIL_COL.VENDOR - 1],
        subject: dRow[DETAIL_COL.SUBJECT - 1],
        usageDate: Utilities.formatDate(new Date(dRow[DETAIL_COL.USAGE_DATE - 1]), TIMEZONE, 'yyyy-MM-dd'),
        taxRate: dRow[DETAIL_COL.TAX_RATE - 1],
        purpose: dRow[DETAIL_COL.PURPOSE - 1],
        invoiceReg: dRow[DETAIL_COL.INVOICE_REG - 1],
        status: 'done',
        receiptUrl: receiptUrls,
        imgData: [] // 画像データはここには含めない(重いため)。Frontendで「保存済画像」として表示対応する。
      });
    }
  }

  return { header: header, items: items };
}

/**
 * 申請更新 (既存申請の置換)
 */
function api_updateApplication(appId, data) {
  // LockService: 排他制御
  return runWithLock_(function () {
    var headerSheet = getSheet_(SHEET_NAMES.HEADER);
    var detailSheet = getSheet_(SHEET_NAMES.DETAIL);

    // 1. ヘッダ特定
    var hValues = headerSheet.getDataRange().getValues();
    var hRowIndex = -1;
    for (var i = 1; i < hValues.length; i++) {
      if (String(hValues[i][HEADER_COL.APPLICATION_ID - 1]) === String(appId)) {
        hRowIndex = i + 1;
        break;
      }
    }
    if (hRowIndex < 0) throw new Error('申請が見つかりません');

    // 2. 権限 & ステータスチェック
    var hData = headerSheet.getRange(hRowIndex, 1, 1, HEADER_COL.COL_COUNT).getValues()[0];
    var ownerEmail = hData[HEADER_COL.APPLICANT_EMAIL - 1];
    var currentStatus = hData[HEADER_COL.STATUS - 1];

    // 自身の申請か確認
    if (normalizeEmail_(ownerEmail) !== getActiveUserEmail_()) throw new Error('権限がありません');

    // 承認済みデータの改ざん防止
    if (currentStatus === STATUS.APPROVED || currentStatus === STATUS.FIXED) {
      throw new Error('承認済み・確定済みの申請は編集できません');
    }

    // 3. 明細削除
    var dLast = detailSheet.getLastRow();
    if (dLast > 1) {
      var dResult = detailSheet.getRange(2, 1, dLast - 1, DETAIL_COL.COL_COUNT).getValues();
      // 後ろから削除
      for (var j = dResult.length - 1; j >= 0; j--) {
        if (String(dResult[j][DETAIL_COL.APPLICATION_ID - 1]) === String(appId)) {
          detailSheet.deleteRow(j + 2);
        }
      }
    }

    // 4. 明細再登録 & 合計計算
    var now = new Date();
    var items = data.items || [];
    var totalAmount = 0;

    items.forEach(function (item) {
      var amount = Number(item.amount || 0);
      totalAmount += amount;

      var receiptUrl = item.receiptUrl || '';
      if (Array.isArray(receiptUrl)) receiptUrl = receiptUrl.join('\n');

      var usageDate = now;
      if (item.usageDate) usageDate = new Date(item.usageDate);

      var detailRow = [];
      detailRow[DETAIL_COL.DETAIL_ID - 1] = 'DET-' + now.getTime() + '-' + Math.floor(Math.random() * 10000);
      detailRow[DETAIL_COL.APPLICATION_ID - 1] = appId;
      detailRow[DETAIL_COL.USAGE_DATE - 1] = usageDate;
      detailRow[DETAIL_COL.AMOUNT - 1] = amount;
      detailRow[DETAIL_COL.TAX_RATE - 1] = Number(item.taxRate) || 10;
      detailRow[DETAIL_COL.TAX_AMOUNT - 1] = 0;
      detailRow[DETAIL_COL.VENDOR - 1] = item.vendor || '';
      detailRow[DETAIL_COL.SUBJECT - 1] = item.subject || '';
      detailRow[DETAIL_COL.PAYMENT_METHOD - 1] = item.paymentMethod || '';
      detailRow[DETAIL_COL.PURPOSE - 1] = item.purpose || '';
      detailRow[DETAIL_COL.RECEIPT_URL - 1] = receiptUrl;
      detailRow[DETAIL_COL.OCR_SCORE - 1] = item.ocrScore || '';
      detailRow[DETAIL_COL.INVOICE_REG - 1] = item.invoiceReg || '不明';

      detailSheet.appendRow(detailRow);
    });

    // 5. ヘッダ更新
    headerSheet.getRange(hRowIndex, HEADER_COL.TOTAL_AMOUNT).setValue(totalAmount);
    headerSheet.getRange(hRowIndex, HEADER_COL.STATUS).setValue(STATUS.APPLYING); // 申請中に戻す
    headerSheet.getRange(hRowIndex, HEADER_COL.APPLICATION_DATE).setValue(now); // 申請日更新

    return { applicationId: appId };
  });
}

// 承認ロジック等は admin_server.gs に移行