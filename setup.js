// =======================================
// setup.gs — スプレッドシート初期セットアップ
// =======================================
//
// 役割：
// - 必要なシートを作成
// - 1行目にヘッダを設定
// - ユーザーマスタにログインユーザーの行を1件だけ自動追加（存在しない場合）
//
// 事前準備：
// - Script Properties に APP_SPREADSHEET_ID を設定
// - app_constants.gs に SHEET_NAMES / HEADER_COL / DETAIL_COL / ROLES / TIMEZONE を定義
// - user_util.gs に getAppSpreadsheet() / getActiveUserEmail_() / normalizeEmail_() を定義
//
// 使い方：
// - エディタ上部の関数一覧から setupAppSheets を選択して実行
// =======================================

/**
 * 初期セットアップのメイン関数
 * - 必要なシートを作成
 * - 1行目にヘッダを設定
 * - ユーザーマスタにログインユーザーの行を1件だけ自動追加（存在しない場合）
 */
function setupAppSheets() {
  var ss = getAppSpreadsheet();

  setupHeaderSheet_(ss);
  setupDetailSheet_(ss);
  setupSubjectSheet_(ss);
  setupUserSheet_(ss);

  Logger.log('初期セットアップが完了しました。');
}

/**
 * 申請ヘッダシートの作成＆ヘッダ行設定
 */
function setupHeaderSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.HEADER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.HEADER);
  }

  // 既に1行目に何か入っていればヘッダ設定はスキップ
  var firstRow = sheet.getRange(1, 1, 1, HEADER_COL.COL_COUNT).getValues()[0];
  var hasHeader = firstRow.some(function (v) { return v !== ''; });
  if (hasHeader) {
    Logger.log('申請ヘッダシートは既にヘッダが設定されています。');
    return;
  }

  var headers = [
    '申請ID',         // APPLICATION_ID
    '申請日',         // APPLICATION_DATE
    '申請者メール',   // APPLICANT_EMAIL
    '申請者氏名',     // APPLICANT_NAME
    '部署',           // APPLICANT_DEPT
    '申請連番',       // SERIAL_NO
    '合計金額',       // TOTAL_AMOUNT
    '状態',           // STATUS
    '承認者メール',   // APPROVER_EMAIL
    '承認日時',       // APPROVED_AT
    '却下日時',       // REJECTED_AT
    '差し戻し日時',   // RETURNED_AT
    '差し戻しコメント', // RETURN_COMMENT
    '経理確定日時',   // FIXED_AT
    '要確認フラグ',   // NEEDS_CHECK
    '備考',           // REMARKS
    '支払状況'        // PAYMENT_STATUS
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 簡単なフォーマット（任意）
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  Logger.log('申請ヘッダシートのヘッダを設定しました。');
}

/**
 * 明細シートの作成＆ヘッダ行設定
 */
function setupDetailSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.DETAIL);
  }

  var firstRow = sheet.getRange(1, 1, 1, DETAIL_COL.COL_COUNT).getValues()[0];
  var hasHeader = firstRow.some(function (v) { return v !== ''; });
  if (hasHeader) {
    Logger.log('明細シートは既にヘッダが設定されています。');
    return;
  }

  var headers = [
    '明細ID',          // DETAIL_ID
    '申請ID',          // APPLICATION_ID
    '利用日',          // USAGE_DATE
    '金額（税込）',    // AMOUNT
    '税率',            // TAX_RATE
    '税額',            // TAX_AMOUNT
    '支払先',          // VENDOR
    '科目',            // SUBJECT
    '支払方法',        // PAYMENT_METHOD
    '利用目的・メモ',  // PURPOSE
    '領収書URL',       // RECEIPT_URL
    'OCR信頼度',       // OCR_SCORE
    'インボイス登録'   // INVOICE_REG
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  Logger.log('明細シートのヘッダを設定しました。');
}

/**
 * 科目マスタシートの作成＆ヘッダ行設定
 */
function setupSubjectSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.SUBJECT_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SUBJECT_MASTER);
  }

  var firstRow = sheet.getRange(1, 1, 1, 3).getValues()[0];
  var hasHeader = firstRow.some(function (v) { return v !== ''; });
  if (hasHeader) {
    Logger.log('科目マスタシートは既にヘッダが設定されています。');
    return;
  }

  var headers = [
    '科目名',                 // 科目名
    'デフォルト税率',         // デフォルト税率
    'キーワード（カンマ区切り）' // キーワード
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  Logger.log('科目マスタシートのヘッダを設定しました。');
}

/**
 * ユーザーマスタシートの作成＆ヘッダ行設定＋ログインユーザー1件追加
 */
function setupUserSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.USER_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USER_MASTER);
  }

  // 1行目ヘッダ
  var firstRow = sheet.getRange(1, 1, 1, USER_MASTER_COL.COL_COUNT).getValues()[0];
  var firstColHeader = String(firstRow[0]).trim();

  // 構造チェック＆移行ロジック
  if (firstColHeader !== 'ユーザーID' && firstColHeader !== 'EMAIL' && firstColHeader !== 'メールアドレス') {
    // ヘッダが空、または全く未知の状態 -> 新規作成
    var headers = [
      'ユーザーID',          // ID
      'メールアドレス',      // EMAIL
      '氏名',                // NAME
      'ロール',              // ROLE
      '上長メールアドレス',  // MANAGER_MAIL
      '拠点'                 // BRANCH
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    Logger.log('ユーザーマスタシートのヘッダを新規設定しました。');
  } else if (firstColHeader === 'メールアドレス' || firstColHeader === 'EMAIL') {
    // 【重要】旧構造からの移行：1列目にID列を挿入し、最後に拠点列を追加
    Logger.log('旧構造のユーザーマスタを検知しました。移行を開始します...');
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue('ユーザーID');
    sheet.getRange(1, 6).setValue('拠点'); // 旧4列 + 1 = 5列。6列目が拠点。

    // 既存行へのIDと拠点の補填
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var range = sheet.getRange(2, 1, lastRow - 1, 6);
      var values = range.getValues();
      for (var i = 0; i < values.length; i++) {
        if (!values[i][0]) values[i][0] = 'U-' + Utilities.getUuid().substring(0, 8); // ID
        if (!values[i][5]) values[i][5] = '本社'; // 拠点
      }
      range.setValues(values);
    }
    Logger.log('既存データの構造移行が完了しました。');
  } else {
    Logger.log('ユーザーマスタの構造は最新です。');
  }

  // すでに2行目以降にデータがある場合は、管理者追加チェックへ
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // 既存データがある場合でも、自分のロールがADMINでない（または見つからない）場合は更新を勧める等の処理は別途だが、
    // ここでは構造維持に留める。
    Logger.log('ユーザーマスタにデータが存在します。');
    return;
  }

  // ログインユーザーを1件だけ追加
  var currentEmail = getActiveUserEmail_(); // 小文字化済み
  var name = currentEmail; // とりあえずメール＝氏名扱い。必要ならあとで編集
  var role = ROLES.ADMIN;   // セットアップ実行者は管理者とする
  var managerEmail = currentEmail;
  var branch = '本社';      // デフォルト

  var row = [];
  row[USER_MASTER_COL.ID - 1] = 'U-' + Utilities.getUuid().substring(0, 8);
  row[USER_MASTER_COL.EMAIL - 1] = currentEmail;
  row[USER_MASTER_COL.NAME - 1] = name;
  row[USER_MASTER_COL.ROLE - 1] = role;
  row[USER_MASTER_COL.MANAGER_EMAIL - 1] = managerEmail;
  row[USER_MASTER_COL.BRANCH - 1] = branch;

  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  Logger.log('ユーザーマスタにログインユーザー(ADMIN)の行を1件追加しました: ' + currentEmail);
}
/**
 * Script Properties を簡単に設定するヘルパー関数
 * エディタからこの関数を実行すると、プロパティが設定されます。
 */
function setupScriptProperties() {
  var props = PropertiesService.getScriptProperties();

  // ↓↓↓↓↓ ここにAPIキーを入力してください ↓↓↓↓↓
  var GEMINI_API_KEY = 'YOUR_API_KEY_HERE';
  // ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

  if (GEMINI_API_KEY !== 'YOUR_API_KEY_HERE') {
    props.setProperty('GEMINI_API_KEY', GEMINI_API_KEY);
    Logger.log('GEMINI_API_KEY を設定しました。');
  } else {
    Logger.log('【注意】GEMINI_API_KEY がデフォルトのままです。書き換えてから実行してください。');
  }

  // スプレッドシートIDも必要ならここで設定可能
  var AppSSId = '1vQAZTBpBBnZtUpcSwg1lXlOvP7wqIwQFwrbT4A1xMTk'; // 現在のID
  props.setProperty('APP_SPREADSHEET_ID', AppSSId);
  Logger.log('APP_SPREADSHEET_ID を設定しました: ' + AppSSId);
}
