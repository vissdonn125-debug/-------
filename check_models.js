// =======================================
// check_models.gs
// =======================================

function listAvailableGeminiModels() {
  // 1. config.gs の関数を使って APIキーを取得
  // ※もしエラーになる場合は、ここに直接 'AIza...' を代入してください
  var apiKey = getGeminiApiKey(); 
  
  // モデル一覧取得用のエンドポイント (v1beta)
  var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;

  var options = {
    method: 'get',
    muteHttpExceptions: true
  };

  Logger.log('APIキーを確認中...');
  Logger.log('モデル一覧を取得しています...');

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      Logger.log('❌ エラーが発生しました (HTTP Code: ' + code + ')');
      Logger.log('詳細: ' + body);
      return;
    }

    var json = JSON.parse(body);
    if (!json.models) {
      Logger.log('モデル情報が見つかりませんでした。');
      return;
    }

    Logger.log('✅ 取得成功！以下のモデルが使用可能です:\n');
    Logger.log('========================================');
    
    // 見やすく整形して出力
    json.models.forEach(function(model) {
      // "generateContent" (チャットやOCRに使う機能) に対応しているものだけ表示
      if (model.supportedGenerationMethods && model.supportedGenerationMethods.includes('generateContent')) {
        // モデルIDから 'models/' を取り除いて表示
        var modelId = model.name.replace('models/', '');
        Logger.log('ID: ' + modelId);
        // Logger.log('名前: ' + model.displayName); // 必要ならコメントアウト解除
        // Logger.log('----------------------------------------');
      }
    });
    Logger.log('========================================');
    Logger.log('※ gemini_ocr.gs の GEMINI_MODEL_ID には、上記の「ID」を設定してください。');

  } catch (e) {
    Logger.log('例外エラーが発生しました: ' + e);
  }
}