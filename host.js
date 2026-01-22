// =======================================
// host.gs — Webアプリエントリ
// =======================================

/**
 * Webアプリエントリポイント
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  var html = template.evaluate()
    .setTitle('経費精算AIOCR')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 埋め込み用

  return html;
}

/**
 * HTMLテンプレート内から別ファイルを include したい場合用
 * （今回は index だけなので未使用でもOK）
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
