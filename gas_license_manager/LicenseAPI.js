/**
 * ライセンス管理API
 * 指定のシート（'Licenses'）にライセンスキーとルートIDを記録・参照し、
 * 有効なユーザーかどうかを判定する。
 */
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const licenseKey = params.licenseKey;
    const rootId = params.rootId;
    
    // パラメータチェック
    if (!licenseKey || !rootId) {
       return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Missing parameters'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    // スプレッドシート側の対象シートを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Licenses');
    if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Licenses sheet not found'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 2行目からループしてキーを探す（1行目はヘッダーの前提）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === licenseKey) {
        let status = data[i][1];
        let registeredRootId = data[i][2];
        
        // 【アクティベーション（初回登録）】
        if (action === 'activate') {
           if (status === 'unused') {
              // 未使用なら使用済みにし、RootId等を登録
              sheet.getRange(i + 1, 2).setValue('active');
              sheet.getRange(i + 1, 3).setValue(rootId);
              sheet.getRange(i + 1, 4).setValue(params.email || "");
              sheet.getRange(i + 1, 5).setValue(params.officeName || "");
              sheet.getRange(i + 1, 6).setValue(params.userName || "");
              return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Activated successfully'})).setMimeType(ContentService.MimeType.JSON);
           } else {
              return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Key is already used'})).setMimeType(ContentService.MimeType.JSON);
           }
        }
        
        // 【通常のライセンスチェック】
        if (action === 'verify') {
           if (status === 'active' && registeredRootId === rootId) {
              return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'License verified'})).setMimeType(ContentService.MimeType.JSON);
           } else {
              return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Invalid license or root ID mismatch'})).setMimeType(ContentService.MimeType.JSON);
           }
        }
      }
    }
    
    // 一致するライセンスキーがない場合
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'License key not found'})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * （テスト用）ブラウザから直接アクセスされた際のメッセージ
 */
function doGet(e) {
  return ContentService.createTextOutput("License API Web App is running.");
}
