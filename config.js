// =======================================
// config.gs (修正版)
// =======================================

/**
 * アプリで使用するスプレッドシートのID
 * ※ Script Properties に設定していない場合は、ここに直接IDを書いてください
 */
function getAppSpreadsheetId() {
  // 1. Script Properties から取得を試みる
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('APP_SPREADSHEET_ID');
  
  // 2. 設定がなければ、以下のID（ハードコード）を使用する
  if (!id) {
    // ↓↓↓↓↓ ここにスプレッドシートのIDを貼り付けてください ↓↓↓↓↓
    return '1vQAZTBpBBnZtUpcSwg1lXlOvP7wqIwQFwrbT4A1xMTk'; 
    // ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
  }
  
  return id;
}

/**
 * Gemini API の API キー取得
 */
function getGeminiApiKey() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  
  // もしプロパティになければエラー（またはここに直接 'AIza...' を書いてもOK）
  if (!key) {
    throw new Error('GEMINI_API_KEY が Script Properties に設定されていません。');
  }
  return key;
}

/**
 * ログインユーザー情報を返すヘルパー（OCR用）
 * スプレッドシート読み込みエラーを回避するため、
 * シート取得に失敗しても最低限の情報を返すように修正
 */
function getCurrentUserInfo_Internal() {
  try {
    // 既存のユーザー取得関数があればそれを使う
    if (typeof getCurrentUserInfo === 'function') {
      return getCurrentUserInfo();
    }
  } catch (e) {
    // スプレッドシートが開けないエラーなどはここで無視する
    Logger.log('ユーザー情報の取得に失敗（スプレッドシート接続エラーの可能性）: ' + e);
  }

  // フォールバック: セッションからメールアドレスだけ取得して返す
  var email = Session.getActiveUser().getEmail();
  return {
    name: email, // 名前が取れないのでメールアドレスで代用
    email: email
  };
}