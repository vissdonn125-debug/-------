// =======================================
// gemini_ocr.gs — Gemini 2.5 Flash (Base64版)
// =======================================

// ★ Gemini 設定
// ご希望の 2.5 Flash に戻します
var GEMINI_MODEL_ID = 'gemini-2.5-flash';
var GEMINI_API_VERSION = 'v1beta'; 
var GEMINI_API_URL = 'https://generativelanguage.googleapis.com/' + GEMINI_API_VERSION + '/models/' + GEMINI_MODEL_ID + ':generateContent';

// ★ 共有ドライブ設定
var IMAGES_FOLDER_ID = '1mkxln-CXGTcl7PRt1nJkmtwqpfMB-4A-'; 

/**
 * フロントエンドから Base64 文字列を受け取って処理する
 * @param {string} base64Data 画像の実データ
 * @param {string} mimeType 画像の種類 (image/jpeg など)
 */
function uploadReceiptForOcr(base64Data, mimeType) {
  try {
    if (!base64Data) {
      throw new Error('画像データが届いていません。');
    }

    // 文字列からBlobを復元
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, "receipt.jpg");

    // --- 1) 画像保存 ---
    var uploadInfo = saveReceiptToDrive_(blob);

    // --- 2) Gemini API で解析 ---
    var ocrResult = analyzeReceiptWithGemini_(base64Data, mimeType);

    return {
      receiptUrl: uploadInfo.fileUrl,
      ocr: ocrResult
    };

  } catch (e) {
    Logger.log('OCR Error: ' + e.toString());
    throw new Error(e.message || '予期せぬエラーが発生しました');
  }
}

/**
 * Gemini API 呼び出し
 */
function analyzeReceiptWithGemini_(base64Data, mimeType) {
  var apiKey = getGeminiApiKey();

  var promptText =
    'このレシート画像を解析し、以下の情報をJSON形式のみで出力してください。\n' +
    'Markdown記号は不要です。\n\n' +
    '税率の判定:\n' +
    '- 食品・飲料（外食除く）は8%\n' +
    '- それ以外は10%\n\n' +
    '科目の判定:\n' +
    '- 食品・飲料は「消耗品（食品）」\n' +
    '- 駐車場は「駐車場代」\n' +
    '- その他は内容に応じて判定\n\n' +
    'インボイス登録の判定:\n' +
    '- 「登録番号」「T+13桁の数字」等の記載があれば「あり」\n' +
    '- なければ「不明」\n\n' +
    '{\n' +
    '  "date": "YYYY-MM-DD",\n' +
    '  "vendor": "店名",\n' +
    '  "amount": "数値(円)",\n' +
    '  "taxRate": "10 or 8",\n' +
    '  "subject": "勘定科目",\n' +
    '  "invoiceReg": "あり or なし or 不明",\n' +
    '  "confidence": "0.0-1.0"\n' +
    '}';

  var payload = {
    contents: [{
      parts: [
        { text: promptText },
        { inlineData: { mimeType: mimeType, data: base64Data } } // ここでBase64を直接使う
      ]
    }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.0
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(GEMINI_API_URL + '?key=' + apiKey, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    Logger.log('Gemini Error: ' + body);
    if (code === 404) throw new Error('モデル(2.5-flash)が見つかりません。');
    if (code === 429) throw new Error('API制限(429)です。少し待ってください。');
    throw new Error('Gemini API エラー (Code: ' + code + ')');
  }

  // 解析
  var jsonResponse = JSON.parse(body);
  var text = '';
  try {
    text = jsonResponse.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('AIからの応答が空でした。');
  }

  // JSONパース
  var parsed;
  try {
    // Markdownの ```json などを除去
    var jsonStr = text.replace(/```json/g, '').replace(/```/g, '').trim();
    var s = jsonStr.indexOf('{');
    var e = jsonStr.lastIndexOf('}');
    if(s !== -1 && e !== -1) jsonStr = jsonStr.substring(s, e+1);
    
    parsed = JSON.parse(jsonStr);
  } catch (e) {
    parsed = {};
  }

  return {
    vendor: parsed.vendor || '',
    amount: Number(parsed.amount || 0),
    date: parsed.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
    taxRate: Number(parsed.taxRate) || 10,
    subject: parsed.subject || '消耗品費',
    invoiceReg: parsed.invoiceReg || '不明',
    confidence: parsed.confidence || 0.8
  };
}

/**
 * Google Drive 保存
 */
function saveReceiptToDrive_(blob) {
  var parentFolder;
  try {
    if (IMAGES_FOLDER_ID) parentFolder = DriveApp.getFolderById(IMAGES_FOLDER_ID);
  } catch (e) {}

  if (!parentFolder) {
    var it = DriveApp.getFoldersByName('経費申請AIOCRデモ');
    parentFolder = it.hasNext() ? it.next() : DriveApp.createFolder('経費申請AIOCRデモ');
  }

  var userInfo = { name: 'user' };
  try {
    if (typeof getCurrentUserInfo === 'function') userInfo = getCurrentUserInfo();
  } catch (e) {}

  var safeName = (userInfo.name || 'user').replace(/[\\\/:\*\?"<>\|]/g, '_');
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  var subFolderName = safeName + '_' + today;

  var subIt = parentFolder.getFoldersByName(subFolderName);
  var subFolder = subIt.hasNext() ? subIt.next() : parentFolder.createFolder(subFolderName);

  var fileName = 'receipt_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HHmmss') + '.jpg';
  var file = subFolder.createFile(blob.copyBlob().setName(fileName));

  return { fileId: file.getId(), fileUrl: file.getUrl() };
}