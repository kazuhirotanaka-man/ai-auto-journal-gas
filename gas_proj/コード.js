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
  const html = HtmlService.createHtmlOutputFromFile('ExportDialog')
    .setTitle('エクスポート対象の選択')
    .setWidth(320)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'エクスポート対象の選択');
}

/**
 * ExportDialogからのエクスポート呼び出し
 */
function executeExportProcess(statuses) {
  return ExportService.exportToCsvWithStatuses(statuses);
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
  const html = HtmlService.createHtmlOutputFromFile('Popup')
    .setTitle('証憑プレビューと仕訳確認')
    .setWidth(1500)
    .setHeight(770);
  SpreadsheetApp.getUi().showModelessDialog(html, '証憑プレビューと仕訳確認');
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
  } catch (e) {
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

/**
 * ポップアップ用のデータを取得する
 * @param {number} rowIndex 取得する行番号（指定がない場合は現在のアクティブセル行）
 * @returns {object} 行データと設定情報のオブジェクト
 */
function getPopupData(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('仕訳データ');
  if (!sheet) return null;

  let targetRow = rowIndex;
  if (!targetRow) {
    const activeRange = SpreadsheetApp.getActiveRange();
    if (!activeRange || activeRange.getSheet().getName() !== '仕訳データ') {
      return null;
    }
    targetRow = activeRange.getRow();
  }

  if (targetRow < 2) return null; // ヘッダー行等は除外

  const maxRow = sheet.getLastRow();
  if (targetRow > maxRow) return null; // データがない行

  // 行のデータを取得 (1列目〜13列目)
  const rowData = sheet.getRange(targetRow, 1, 1, 13).getValues()[0];
  if (!rowData[0] && !rowData[11]) return null; // データが空（日付もIDもない）

  const config = ConfigService.getAllConfig();

  // 日付のフォーマット処理
  let dateVal = rowData[0];
  if (dateVal instanceof Date) {
    dateVal = Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
  } else if (dateVal !== "" && dateVal != null) {
    dateVal = String(dateVal).replace(/\//g, "-");
  }

  return {
    rowIndex: targetRow,
    fileId: rowData[11] ? String(rowData[11]) : null,
    data: {
      date: dateVal,
      debitAccount: rowData[1],
      debitSubAccount: rowData[2],
      debitTax: rowData[3],
      debitAmount: rowData[4],
      creditAccount: rowData[5],
      creditSubAccount: rowData[6],
      creditTax: rowData[7],
      creditAmount: rowData[8],
      description: rowData[9],
      warningText: rowData[10],
      status: rowData[12] || '未確認'
    },
    // ドロップダウン用のリスト
    accountsList: config.accountsList || [],
    subAccountsList: config.subAccountsList || [],
    taxCategoryList: config.taxCategoryList || []
  };
}

/**
 * ポップアップからの更新リクエストを処理し、次の行があればそのデータを返す
 * @param {number} rowIndex 更新対象の行
 * @param {object} updateData 更新するデータ
 * @param {string} action アクション ('save', 'exclude', 'back', 'justNext')
 * @returns {object} 次(または前)の行のデータ、なければnull
 */
function updateAndProcessNext(rowIndex, updateData, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('仕訳データ');

  if ((action === 'save' || action === 'exclude') && updateData) {
    // データの更新 (1〜10列目と13列目)
    const range = sheet.getRange(rowIndex, 1, 1, 13);
    const currentValues = range.getValues()[0];

    // updateDataの内容で上書き
    currentValues[0] = updateData.date;
    currentValues[1] = updateData.debitAccount;
    currentValues[2] = updateData.debitSubAccount;
    currentValues[3] = updateData.debitTax;
    currentValues[4] = updateData.debitAmount;
    currentValues[5] = updateData.creditAccount;
    currentValues[6] = updateData.creditSubAccount;
    currentValues[7] = updateData.creditTax;
    currentValues[8] = updateData.creditAmount;
    currentValues[9] = updateData.description;

    // ステータスと背景色の更新
    if (action === 'exclude') {
      currentValues[12] = '対象外';
      range.setBackground("#e0e0e0"); // 灰色
      currentValues[10] = '対象外';
    } else {
      // save の場合はステータスを確認済等に
      currentValues[12] = updateData.status || '確認済';
      if (currentValues[12] === '確認済') {
        range.setBackground(null); // 背景色クリア
        currentValues[10] = "確認済"; // 警告クリア
      }
    }

    range.setValues([currentValues]);
  }

  // 移動先の行を決定
  let nextRow = rowIndex;
  if (action === 'back') {
    nextRow = rowIndex - 1;
  } else {
    nextRow = rowIndex + 1;
  }

  const maxRow = sheet.getLastRow();
  if (nextRow < 2 || nextRow > maxRow) {
    return null; // 端に到達した
  }

  // シート側の選択行も移動させる
  sheet.getRange(nextRow, 1).activate();

  return getPopupData(nextRow);
}
