/**
 * ライセンス管理API
 * 指定のシート（'Licenses'）にライセンスキーとルートIDを記録・参照し、
 * 有効なユーザーかどうかを判定する。
 * また「利用状況」シートに時系列の利用・認証ログを記録する。
 */

const USAGE_SHEET_NAME = '利用状況';
const USAGE_HEADERS = [
  '実行日時', 'ライセンスキー', '事務所名', '担当者名', 'メールアドレス',
  'ドライブRoot ID', 'スプレッドシートID', 'アクション', '会計ソフト',
  '処理枚数', '生成仕訳行数', '要確認・スキップ', '実行結果', 'エラー内容',
  '処理時間(秒)', 'バージョン'
];

/**
 * 「利用状況」シートを取得、存在しなければヘッダー付きで新規作成する
 */
function getOrCreateUsageSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(USAGE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USAGE_SHEET_NAME);
    sheet.appendRow(USAGE_HEADERS);
    
    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, USAGE_HEADERS.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // 列幅の視認性設定
    sheet.setColumnWidth(1, 160); // 実行日時
    sheet.setColumnWidth(2, 280); // ライセンスキー
    sheet.setColumnWidth(3, 160); // 事務所名
    sheet.setColumnWidth(4, 120); // 担当者名
    sheet.setColumnWidth(5, 180); // メールアドレス
    sheet.setColumnWidth(6, 200); // ドライブRoot ID
    sheet.setColumnWidth(7, 200); // スプレッドシートID
    sheet.setColumnWidth(8, 120); // アクション
    sheet.setColumnWidth(9, 100); // 会計ソフト
    sheet.setColumnWidth(10, 80); // 処理枚数
    sheet.setColumnWidth(11, 100); // 生成仕訳行数
    sheet.setColumnWidth(12, 160); // 要確認・スキップ
    sheet.setColumnWidth(13, 80);  // 実行結果
    sheet.setColumnWidth(14, 200); // エラー内容
    sheet.setColumnWidth(15, 100); // 処理時間(秒)
    sheet.setColumnWidth(16, 90);  // バージョン
  }
  return sheet;
}

/**
 * 「利用状況」シートにログ行を追記する
 */
function appendUsageLog(ss, logEntry) {
  try {
    const sheet = getOrCreateUsageSheet(ss);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    const row = [
      logEntry.timestamp || nowStr,
      logEntry.licenseKey || "",
      logEntry.officeName || "",
      logEntry.userName || "",
      logEntry.email || "",
      logEntry.rootId || "",
      logEntry.spreadsheetId || "",
      logEntry.action || "",
      logEntry.accountingSoftware || "",
      logEntry.processedCount !== undefined && logEntry.processedCount !== null ? logEntry.processedCount : "",
      logEntry.journalCount !== undefined && logEntry.journalCount !== null ? logEntry.journalCount : "",
      logEntry.warningCount || "",
      logEntry.status || "",
      logEntry.errorMessage || "",
      logEntry.processingTimeSec !== undefined && logEntry.processingTimeSec !== null ? logEntry.processingTimeSec : "",
      logEntry.version || ""
    ];
    sheet.appendRow(row);
  } catch (err) {
    console.error('利用状況ログの記録に失敗しました: ', err);
  }
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const licenseKey = params.licenseKey;
    const rootId = params.rootId;
    
    // パラメータチェック
    if (!licenseKey || !rootId) {
       console.error('Missing parameters: licenseKey or rootId', params);
       return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Missing parameters'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // スプレッドシート側の対象シートを取得
    const sheet = ss.getSheetByName('Licenses');
    if (!sheet) {
        console.error('Licenses sheet not found in the active spreadsheet.');
        return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Licenses sheet not found'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 2行目からループしてキーを探す（1行目はヘッダーの前提）
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === licenseKey) {
        let status = data[i][1];
        let registeredRootId = data[i][2];
        let registeredEmail = data[i][3] || "";
        let registeredOffice = data[i][4] || "";
        let registeredUser = data[i][5] || "";
        
        // 【利用ログの記録】
        if (action === 'log_usage') {
          appendUsageLog(ss, {
            licenseKey: licenseKey,
            officeName: params.officeName || registeredOffice,
            userName: params.userName || registeredUser,
            email: params.email || registeredEmail,
            rootId: rootId,
            spreadsheetId: params.spreadsheetId || "",
            action: params.usageAction || "機能利用",
            accountingSoftware: params.accountingSoftware || "",
            processedCount: params.processedCount,
            journalCount: params.journalCount,
            warningCount: params.warningCount || "",
            status: params.usageStatus || "成功",
            errorMessage: params.errorMessage || "",
            processingTimeSec: params.processingTimeSec,
            version: params.version || ""
          });
          return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Usage logged successfully'})).setMimeType(ContentService.MimeType.JSON);
        }

        // 【ステータス確認のみ（事前チェック用）】
        if (action === 'check_status') {
           return ContentService.createTextOutput(JSON.stringify({
             status: 'success', 
             keyStatus: status,
             registeredRootId: registeredRootId || ""
           })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // 【アクティベーション（初回登録）】
        if (action === 'activate') {
           if (status === 'unused') {
              // 未使用なら使用済みにし、RootId等を登録
              sheet.getRange(i + 1, 2).setValue('active');
              sheet.getRange(i + 1, 3).setValue(rootId);
              sheet.getRange(i + 1, 4).setValue(params.email || "");
              sheet.getRange(i + 1, 5).setValue(params.officeName || "");
              sheet.getRange(i + 1, 6).setValue(params.userName || "");
              console.info(`Activated successfully. Key: ${licenseKey}, RootID: ${rootId}`);

              // 利用状況シートにアクティベーション成功ログを記録
              appendUsageLog(ss, {
                licenseKey: licenseKey,
                officeName: params.officeName || "",
                userName: params.userName || "",
                email: params.email || "",
                rootId: rootId,
                spreadsheetId: params.spreadsheetId || "",
                action: "アクティベーション",
                status: "成功",
                version: params.version || ""
              });

              return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Activated successfully'})).setMimeType(ContentService.MimeType.JSON);
           } else {
              console.warn(`Activation failed: Key is already used. Key: ${licenseKey}`);
              
              // 利用状況シートに失敗ログを記録
              appendUsageLog(ss, {
                licenseKey: licenseKey,
                officeName: registeredOffice,
                userName: registeredUser,
                email: registeredEmail,
                rootId: rootId,
                spreadsheetId: params.spreadsheetId || "",
                action: "アクティベーション",
                status: "エラー",
                errorMessage: "Key is already used",
                version: params.version || ""
              });

              return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Key is already used'})).setMimeType(ContentService.MimeType.JSON);
           }
        }
        
        // 【通常のライセンスチェック】
        if (action === 'verify') {
           if (status === 'active' && registeredRootId === rootId) {
              // 利用状況シートに認証成功ログを記録
              appendUsageLog(ss, {
                licenseKey: licenseKey,
                officeName: registeredOffice,
                userName: registeredUser,
                email: registeredEmail,
                rootId: rootId,
                spreadsheetId: params.spreadsheetId || "",
                action: "ライセンス認証",
                status: "成功",
                version: params.version || ""
              });

              return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'License verified'})).setMimeType(ContentService.MimeType.JSON);
           } else {
              const errMsg = status !== 'active' ? `Status is ${status}` : 'Root ID mismatch';
              console.warn(`Verification failed: Invalid license or root ID mismatch. Key: ${licenseKey}, ExpectedRoot: ${registeredRootId}, ProvidedRoot: ${rootId}`);

              // 利用状況シートに認証失敗ログを記録
              appendUsageLog(ss, {
                licenseKey: licenseKey,
                officeName: registeredOffice,
                userName: registeredUser,
                email: registeredEmail,
                rootId: rootId,
                spreadsheetId: params.spreadsheetId || "",
                action: "ライセンス認証",
                status: "エラー",
                errorMessage: errMsg,
                version: params.version || ""
              });

              return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Invalid license or root ID mismatch'})).setMimeType(ContentService.MimeType.JSON);
           }
        }
      }
    }
    
    // 一致するライセンスキーがない場合
    console.warn(`License key not found: ${licenseKey}`);
    appendUsageLog(ss, {
      licenseKey: licenseKey,
      rootId: rootId,
      spreadsheetId: params.spreadsheetId || "",
      action: action === 'log_usage' ? (params.usageAction || "機能利用") : (action || "不明"),
      status: "エラー",
      errorMessage: "License key not found",
      version: params.version || ""
    });

    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'License key not found'})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('doPostエラー: ', error);
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * （テスト用）ブラウザから直接アクセスされた際のメッセージ
 */
function doGet(e) {
  return ContentService.createTextOutput("License API Web App is running.");
}
