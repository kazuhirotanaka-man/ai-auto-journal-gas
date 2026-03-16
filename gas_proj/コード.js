/**
 * メニュー表示・イベントハンドラ・メインフローを管理するエントリーポイント
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧾 仕訳作成AI')
      .addItem('▶ 未処理ファイルの解析を実行', 'processNewReceipts')
      .addSeparator()
      .addItem('🖼 プレビュー用サイドバーを開く', 'showSidebar')
      .addItem('🪟 プレビューをポップアップ（大画面）で開く', 'showDialog')
      .addSeparator()
      .addItem('💾 会計ソフト形式でエクスポート', 'exportCsvHandler')
      .addToUi();
}

function exportCsvHandler() {
  ExportService.exportToCsv();
}

/**
 * サイドバーを表示する
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('証憑プレビュー')
      .setWidth(400); // サイドバーは最大300pxまでしか幅が広がらない仕様です
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * モードレスダイアログ（ポップアップ）を表示する
 * ユーザーがサイズ変更や移動が可能で、より大きく表示できます
 */
function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('証憑プレビュー')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModelessDialog(html, '証憑プレビュー');
}

/**
 * サイドバーのHTMLから定期的に呼ばれ、現在選択されている行のファイルIDを返す
 */
function getActiveFileId() {
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) return null;
  
  const sheet = activeRange.getSheet();
  if (sheet.getName() !== '仕訳データ') return null;
  
  const row = activeRange.getRow();
  if (row < 2) return null; // ヘッダー行
  
  // ファイルIDはL列 (12列目) にあると仮定して取得
  const fileId = sheet.getRange(row, 12).getValue();
  return fileId ? String(fileId) : null;
}

/**
 * メインの処理フロー：差分抽出 -> Gemini連携 -> シート書き込み -> ログ記録
 */
function processNewReceipts() {
  const ui = SpreadsheetApp.getUi();
  
  // APIキーの事前チェック
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey || apiKey.trim() === "") {
    ui.alert("⚠️ 初期設定エラー", "Gemini APIキーが設定されていません。\n\nメニューの「拡張機能」＞「Apps Script」からエディタを開き、左側の「プロジェクトの設定（歯車マーク）」＞「スクリプトプロパティ」に GEMINI_API_KEY を追加してください。", ui.ButtonSet.OK);
    return;
  }
  
  const config = ConfigService.getAllConfig();
  
  if (!config.folderId) {
    ui.alert("設定エラー", "対象のDriveフォルダURLが設定されていません。設定シートを確認してください。", ui.ButtonSet.OK);
    return;
  }

  // 1. 未処理ファイルの取得
  ui.alert("情報", "フォルダをスキャンし、未処理の証憑を確認しています...", ui.ButtonSet.OK);
  let unprocessedFiles;
  try {
    unprocessedFiles = DriveServiceObj.getUnprocessedFiles(config.folderId);
  } catch(e) {
    ui.alert("エラー", "ファイル取得中にエラーが発生しました。\n" + e.message, ui.ButtonSet.OK);
    return;
  }

  if (unprocessedFiles.length === 0) {
    ui.alert("完了", "未処理の証憑ファイルはありませんでした。", ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert("処理開始", `未処理のファイルが ${unprocessedFiles.length} 件見つかりました。\nAIによる解析を開始しますか？\n（数分かかる場合があります）`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  // 2. ループ処理でGeminiに送信し、結果をシートに追記
  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < unprocessedFiles.length; i++) {
    const file = unprocessedFiles[i];
    
    try {
      // Geminiで解析
      const parsedData = GeminiService.analyzeReceipt(file, config);
      
      // シートに追記
      JournalService.appendJournalEntry(parsedData, file);
      
      // ログに記録
      JournalService.logProcessedFile(file);
      
      successCount++;
    } catch (e) {
      console.error(`ファイル [${file.getName()}] の処理に失敗しました: `, e);
      errorCount++;
      // エラーログ等に記録することも検討
    }
  }

  ui.alert("処理完了", `解析が完了しました。\n成功: ${successCount}件\n失敗: ${errorCount}件`, ui.ButtonSet.OK);
}
