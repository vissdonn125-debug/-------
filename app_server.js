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
 * カート申請作成 (複数明細)
 * data: { applicantId, items: [ { amount, vendor, subject, receiptUrl, ... } ] }
 */
function createApplicationWithDetails(data) {
  // 1. ユーザー特定
  var userInfo = null;
  if (data.applicantId) userInfo = findUserById_(data.applicantId);
  if (!userInfo) userInfo = findUserByEmail_(getActiveUserEmail_());
  if (!userInfo) {
    userInfo = { email: getActiveUserEmail_(), name: 'Unknown', approverEmail: '', role: 'APPLICANT' };
  }

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

  items.forEach(function(item) {
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
    detailRow[DETAIL_COL.DETAIL_ID - 1]      = 'DET-' + now.getTime() + '-' + Math.floor(Math.random()*10000);
    detailRow[DETAIL_COL.APPLICATION_ID - 1] = appId;
    detailRow[DETAIL_COL.USAGE_DATE - 1]     = usageDate;
    detailRow[DETAIL_COL.AMOUNT - 1]         = amount;
    detailRow[DETAIL_COL.TAX_RATE - 1]       = taxRate;
    detailRow[DETAIL_COL.TAX_AMOUNT - 1]     = 0;
    detailRow[DETAIL_COL.VENDOR - 1]         = item.vendor || '';
    detailRow[DETAIL_COL.SUBJECT - 1]        = item.subject || '';
    detailRow[DETAIL_COL.PAYMENT_METHOD - 1] = item.paymentMethod || '';
    detailRow[DETAIL_COL.PURPOSE - 1]        = item.purpose || '';
    detailRow[DETAIL_COL.RECEIPT_URL - 1]    = receiptUrl;
    detailRow[DETAIL_COL.OCR_SCORE - 1]      = item.ocrScore || '';
    detailRow[DETAIL_COL.INVOICE_REG - 1]    = item.invoiceReg || '不明';

    detailSheet.appendRow(detailRow);
  });

  // 4. ヘッダー保存
  var lastRow = headerSheet.getLastRow();
  var serial = lastRow; // 簡易

  var headerRow = [];
  headerRow[HEADER_COL.APPLICATION_ID - 1]   = appId;
  headerRow[HEADER_COL.APPLICATION_DATE - 1] = appDate;
  headerRow[HEADER_COL.APPLICANT_EMAIL - 1]  = userInfo.email;
  headerRow[HEADER_COL.APPLICANT_NAME - 1]   = userInfo.name;
  headerRow[HEADER_COL.APPLICANT_DEPT - 1]   = userInfo.dept || '';
  headerRow[HEADER_COL.SERIAL_NO - 1]        = serial;
  headerRow[HEADER_COL.TOTAL_AMOUNT - 1]     = totalAmount; // 自動計算した合計
  headerRow[HEADER_COL.STATUS - 1]           = STATUS.APPLYING;
  headerRow[HEADER_COL.APPROVER_EMAIL - 1]   = userInfo.approverEmail;
  headerRow[HEADER_COL.NEEDS_CHECK - 1]      = false;
  
  headerSheet.appendRow(headerRow);

  return { applicationId: appId };
}

// 承認ロジック等は admin_server.gs に移行