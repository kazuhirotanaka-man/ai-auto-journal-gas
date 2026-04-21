/**
 * メニュー表示・イベントハンドラ・メインフローを管理するエントリーポイント
 */

const FREEE_COL = { INCOME_EXPENSE: 0, ACCRUAL_DATE: 1, PARTNER: 2, REG_NUM: 3, PAY_STATUS: 4, PAY_DATE: 5, WALLET: 6, ACC_ITEM: 7, AMOUNT: 8, TAX: 9, ITEM: 10, DEPT: 11, TAG: 12, DESC: 13, GUESS_DOC: 14, GUESS_PAY: 15, WARNING: 16, FILE_ID: 17, STATUS: 18 };
const YAYOI_COL = { DATE: 0, DEBIT_ACC: 1, DEBIT_SUB: 2, DEBIT_TAX: 3, DEBIT_AMT: 4, CREDIT_ACC: 5, CREDIT_SUB: 6, CREDIT_TAX: 7, CREDIT_AMT: 8, DESC: 9, WARNING: 10, FILE_ID: 11, STATUS: 12 };

/**
 * 日付（文字列またはDate）を "yyyy-MM-dd" 形式の文字列にフォーマットする
 * @param {Date|string|number} rawDate - フォーマット対象の日付
 * @param {string} [tz] - タイムゾーン (例: "Asia/Tokyo")。省略時は"Asia/Tokyo"
 * @returns {string} フォーマットされた日付文字列
 */
function formatRawDate(rawDate, tz) {
  tz = tz || "Asia/Tokyo";
  if (rawDate instanceof Date) {
    return Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
  } else if (rawDate !== "" && rawDate != null) {
    const d = new Date(rawDate);
    if (!isNaN(d.getTime())) {
      // "YYYY/MM/DD" 等の文字列から構築されたDateを正しくフォーマットする
      return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    }
    return String(rawDate).replace(/\//g, "-");
  }
  return "";
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧾 仕訳作成AI')
    .addItem('▶ 未処理ファイルの解析を実行', 'processNewReceipts')
    .addSeparator()
    .addItem('🖼 プレビュー用サイドバーを開く', 'showSidebar')
    .addItem('🪟 プレビューをポップアップ（大画面）で開く', 'showDialog')
    .addSeparator()
    .addItem('💾 会計ソフト形式でエクスポート', 'exportCsvHandler')
    .addItem('☁️ freee会計に取引登録', 'exportFreeeHandler')
    .addToUi();

  // freee連携メニュー
  ui.createMenu('freee連携')
    .addItem('認証', 'showFreeeSidebar')
    .addItem('操作事業所の選択', 'showCompanySelectorSidebar')
    .addSeparator()
    .addItem('マスタ取得', 'fetchFreeeMasters')
    .addSeparator()
    .addItem('トークン自動更新設定', 'toggleAutoRefreshToken')
    .addSeparator()
    .addItem('【デモ】取引の登録〜削除', 'runDealLifecycleSample')
    .addSeparator()
    .addItem('連携を解除', 'reset')
    .addToUi();
}

function exportCsvHandler() {
  if (!LicenseService.requireLicense()) return;
  
  const html = HtmlService.createHtmlOutputFromFile('ExportDialog')
    .setTitle('エクスポート対象の選択')
    .setWidth(320)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'エクスポート対象の選択');
}

function exportFreeeHandler() {
  if (!LicenseService.requireLicense()) return;
  
  const html = HtmlService.createHtmlOutputFromFile('freee_ExportDialog')
    .setTitle('freee会計へ取引登録')
    .setWidth(320)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'freee会計へ取引登録');
}

/**
 * ExportDialogからのエクスポート呼び出し
 * @param {string[]} statuses
 */
function executeExportProcess(statuses) {
  return ExportService.exportToCsvWithStatuses(statuses);
}

/**
 * サイドバーを表示する
 */
function showSidebar() {
  if (!LicenseService.requireLicense()) return;
  
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
  if (!LicenseService.requireLicense()) return;
  
  const config = ConfigService.getAllConfig();
  if (config.accountingSoftware === "freee会計") {
    const html = HtmlService.createHtmlOutputFromFile('freee_Popup')
      .setTitle('証憑プレビューと仕訳確認 (freee会計)')
      .setWidth(1500)
      .setHeight(770);
    SpreadsheetApp.getUi().showModelessDialog(html, '証憑プレビューと仕訳確認 (freee会計)');
  } else {
    const html = HtmlService.createHtmlOutputFromFile('Popup')
      .setTitle('証憑プレビューと仕訳確認')
      .setWidth(1500)
      .setHeight(770);
    SpreadsheetApp.getUi().showModelessDialog(html, '証憑プレビューと仕訳確認');
  }
}

/**
 * サイドバーのHTMLから定期的に呼ばれ、現在選択されている行のファイルIDを返す
 * @returns {string|null} ファイルID。取得できない場合はnull
 */
function getActiveFileId() {
  const activeRange = SpreadsheetApp.getActiveRange();
  if (!activeRange) return null;

  const sheet = activeRange.getSheet();
  const sheetName = sheet.getName();
  if (sheetName !== '仕訳データ' && sheetName !== 'freee取引データ') return null;

  let row = activeRange.getRow();
  if (row < 2) return null; // ヘッダー行

  if (sheetName === 'freee取引データ') {
    // 収支(0列目)が空でない行まで遡る
    while (row >= 2) {
      const rowData = sheet.getRange(row, 1, 1, 19).getValues()[0];
      if (rowData[FREEE_COL.INCOME_EXPENSE] !== "") {
        return rowData[FREEE_COL.FILE_ID] ? String(rowData[FREEE_COL.FILE_ID]) : null;
      }
      row--;
    }
    return null;
  } else {
    // 弥生用：ファイルIDは対象列にある
    const fileId = sheet.getRange(row, YAYOI_COL.FILE_ID + 1).getValue();
    return fileId ? String(fileId) : null;
  }
}

/**
 * メインの処理フロー：差分抽出 -> Gemini連携 -> シート書き込み -> ログ記録
 */
function processNewReceipts() {
  if (!LicenseService.requireLicense()) return;

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
  
  // 証票フォルダのルートID同一性チェック
  if (!LicenseService.isEvidenceFolderValid(config.folderId)) {
    ui.alert("設定エラー", "指定された証票フォルダは、このツールに関連付けられたGoogleドライブ内にありません。\nセキュリティ保護のため、外部ドライブのフォルダは指定できません。", ui.ButtonSet.OK);
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
    const item = unprocessedFiles[i];
    const file = item.file;
    const relativePath = item.relativePath;

    try {
      // Geminiで解析
      const parsedData = GeminiService.analyzeReceipt(file, config, relativePath);

      // シートに追記
      JournalService.appendJournalEntry(parsedData, file, config);

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
 * @param {number} [rowIndex] 取得する行番号（指定がない場合は現在のアクティブセル行）
 * @returns {object|null} 行データと設定情報のオブジェクト、存在しない場合はnull
 */
function getPopupData(rowIndex) {
  const config = ConfigService.getAllConfig();

  const isFreee = config.accountingSoftware === "freee会計";
  const sheetName = isFreee ? 'freee取引データ' : '仕訳データ';
  const dataColumns = isFreee ? 19 : 13;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  let targetRow = rowIndex;
  if (!targetRow) {
    const activeRange = SpreadsheetApp.getActiveRange();
    if (!activeRange || activeRange.getSheet().getName() !== sheetName) {
      return null;
    }
    targetRow = activeRange.getRow();
  }

  if (targetRow < 2) return null; // ヘッダー行等は除外

  const maxRow = sheet.getLastRow();
  if (targetRow > maxRow) return null; // データがない行

  // 行のデータを取得
  const rowData = sheet.getRange(targetRow, 1, 1, dataColumns).getValues()[0];
  if (rowData.join("").length === 0) return null; // データが空

  const tz = ss.getSpreadsheetTimeZone() || "Asia/Tokyo";

  if (isFreee) {
    // 1列目（収支）が空白の場合、空白でなくなるまで上方向に遡って取引の開始行（Head）を見つける
    while (targetRow >= 2) {
      if (sheet.getRange(targetRow, 1).getValue() !== "") {
        break;
      }
      targetRow--;
    }
    // 万が一見つからなかった場合は元の行を採用
    if (targetRow < 2) targetRow = rowIndex;

    const headRowRange = sheet.getRange(targetRow, 1, 1, dataColumns);
    const headRowData = headRowRange.getValues()[0];
    const fileId = headRowData[FREEE_COL.FILE_ID] ? String(headRowData[FREEE_COL.FILE_ID]) : null;

    // Head行から下方向に、同じfileIdかつ収支が空白（＝同一取引の明細）である行を取得し続ける
    let detailsData = [];
    let currentRow = targetRow;

    while (currentRow <= maxRow) {
      const rowVars = sheet.getRange(currentRow, 1, 1, dataColumns).getValues()[0];

      // 最初（Head行）は無条件で追加。2行目以降は 収支が空白 なら同じ取引の明細とみなす。
      if (currentRow === targetRow || rowVars[FREEE_COL.INCOME_EXPENSE] === "") {
        detailsData.push(rowVars);
      } else {
        break; // 別の取引が始まったら終了
      }
      currentRow++;
    }

    const rowCount = detailsData.length;

    const retData = {
      startRowIndex: targetRow,
      rowCount: rowCount,
      fileId: fileId,
      data: {
        transaction: {
          incomeExpense: headRowData[FREEE_COL.INCOME_EXPENSE],
          date: formatRawDate(headRowData[FREEE_COL.ACCRUAL_DATE], tz),
          partner: headRowData[FREEE_COL.PARTNER],
          registrationNumber: headRowData[FREEE_COL.REG_NUM],
          paymentStatus: headRowData[FREEE_COL.PAY_STATUS],
          paymentDate: formatRawDate(headRowData[FREEE_COL.PAY_DATE], tz),
          wallet: headRowData[FREEE_COL.WALLET],
          status: headRowData[FREEE_COL.STATUS] || '未確認',
          warningText: headRowData[FREEE_COL.WARNING],
          guessedDocumentType: headRowData[FREEE_COL.GUESS_DOC],
          guessedPaymentMethod: headRowData[FREEE_COL.GUESS_PAY]
        },
        details: detailsData.map(row => ({
          accountItem: row[FREEE_COL.ACC_ITEM],
          amount: row[FREEE_COL.AMOUNT],
          taxCategory: row[FREEE_COL.TAX],
          item: row[FREEE_COL.ITEM],
          department: row[FREEE_COL.DEPT],
          memoTag: row[FREEE_COL.TAG],
          description: row[FREEE_COL.DESC]
        }))
      },
      // ドロップダウン用のリスト(freee用)
      freeeAccountsList: config.freeeAccountsList || [],
      freeeTaxCategoryList: config.freeeTaxCategoryList || [],
      freeeWalletsList: config.freeeWalletsList || [],
      freeePartnersList: config.freeePartnersList || [],
      freeeItemsList: config.freeeItemsList || [],
      freeeDepartmentsList: config.freeeDepartmentsList || [],
      freeeTagsList: config.freeeTagsList || []
    };
    return retData;
  } else {
    // 弥生用ロジック
    const dateVal = formatRawDate(rowData[YAYOI_COL.DATE], tz);

    const retData = {
      rowIndex: targetRow,
      fileId: rowData[YAYOI_COL.FILE_ID] ? String(rowData[YAYOI_COL.FILE_ID]) : null,
      data: {
        date: dateVal,
        debitAccount: rowData[YAYOI_COL.DEBIT_ACC],
        debitSubAccount: rowData[YAYOI_COL.DEBIT_SUB],
        debitTax: rowData[YAYOI_COL.DEBIT_TAX],
        debitAmount: rowData[YAYOI_COL.DEBIT_AMT],
        creditAccount: rowData[YAYOI_COL.CREDIT_ACC],
        creditSubAccount: rowData[YAYOI_COL.CREDIT_SUB],
        creditTax: rowData[YAYOI_COL.CREDIT_TAX],
        creditAmount: rowData[YAYOI_COL.CREDIT_AMT],
        description: rowData[YAYOI_COL.DESC],
        warningText: rowData[YAYOI_COL.WARNING],
        status: rowData[YAYOI_COL.STATUS] || '未確認'
      },
      // ドロップダウン用のリスト
      accountsList: config.accountsList || [],
      subAccountsList: config.subAccountsList || [],
      taxCategoryList: config.taxCategoryList || []
    };
    return retData;
  }
}

/**
 * ポップアップからの更新リクエストを処理し、次の行があればそのデータを返す
 * @param {number} rowIndex 更新対象の行
 * @param {object} updateData 更新するデータ
 * @param {string} action アクション ('save', 'exclude', 'back', 'justNext')
 * @returns {object|null} 次(または前)の行のデータ、なければnull
 */
function updateAndProcessNext(rowIndex, updateData, action) {
  const config = ConfigService.getAllConfig();

  const isFreee = config.accountingSoftware === "freee会計";
  const sheetName = isFreee ? 'freee取引データ' : '仕訳データ';
  const dataColumns = isFreee ? 19 : 13;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if ((action === 'save' || action === 'exclude') && updateData) {
    if (isFreee) {
      const rowCount = updateData.details.length;
      const range = sheet.getRange(rowIndex, 1, rowCount, dataColumns);
      const values = range.getValues();

      values.forEach((currentValues, i) => {
        if (i === 0) {
          // Head行
          currentValues[FREEE_COL.INCOME_EXPENSE] = updateData.transaction.incomeExpense;
          currentValues[FREEE_COL.ACCRUAL_DATE] = updateData.transaction.date;
          currentValues[FREEE_COL.PARTNER] = updateData.transaction.partner;
          currentValues[FREEE_COL.REG_NUM] = updateData.transaction.registrationNumber;
          currentValues[FREEE_COL.PAY_STATUS] = updateData.transaction.paymentStatus;
          currentValues[FREEE_COL.PAY_DATE] = updateData.transaction.paymentDate;
          currentValues[FREEE_COL.WALLET] = updateData.transaction.wallet;

          if (action === 'exclude') {
            currentValues[FREEE_COL.STATUS] = '対象外';
          } else {
            currentValues[FREEE_COL.STATUS] = updateData.status || '確認済';
          }
        } else {
          // 明細行は取引基本項目および推測系・状態管理項目を空白にする
          currentValues[FREEE_COL.INCOME_EXPENSE] = "";
          currentValues[FREEE_COL.ACCRUAL_DATE] = "";
          currentValues[FREEE_COL.PARTNER] = "";
          currentValues[FREEE_COL.REG_NUM] = "";
          currentValues[FREEE_COL.PAY_STATUS] = "";
          currentValues[FREEE_COL.PAY_DATE] = "";
          currentValues[FREEE_COL.WALLET] = "";

          currentValues[FREEE_COL.GUESS_DOC] = "";
          currentValues[FREEE_COL.GUESS_PAY] = "";
          currentValues[FREEE_COL.WARNING] = "";
          currentValues[FREEE_COL.FILE_ID] = "";
          currentValues[FREEE_COL.STATUS] = "";
        }

        const detail = updateData.details[i];
        currentValues[FREEE_COL.ACC_ITEM] = detail.accountItem;
        currentValues[FREEE_COL.AMOUNT] = detail.amount;
        currentValues[FREEE_COL.TAX] = detail.taxCategory;
        currentValues[FREEE_COL.ITEM] = detail.item;
        currentValues[FREEE_COL.DEPT] = detail.department;
        currentValues[FREEE_COL.TAG] = detail.memoTag;
        currentValues[FREEE_COL.DESC] = detail.description;
      });
      range.setValues(values);
      if (action === 'exclude') {
        range.setBackground("#e0e0e0");
      } else {
        if (updateData.status === '確認済' || !updateData.status) {
          range.setBackground(null);
        }
      }
    } else {
      const range = sheet.getRange(rowIndex, 1, 1, dataColumns);
      const currentValues = range.getValues()[0];

      // データの更新
      currentValues[YAYOI_COL.DATE] = updateData.date;
      currentValues[YAYOI_COL.DEBIT_ACC] = updateData.debitAccount;
      currentValues[YAYOI_COL.DEBIT_SUB] = updateData.debitSubAccount;
      currentValues[YAYOI_COL.DEBIT_TAX] = updateData.debitTax;
      currentValues[YAYOI_COL.DEBIT_AMT] = updateData.debitAmount;
      currentValues[YAYOI_COL.CREDIT_ACC] = updateData.creditAccount;
      currentValues[YAYOI_COL.CREDIT_SUB] = updateData.creditSubAccount;
      currentValues[YAYOI_COL.CREDIT_TAX] = updateData.creditTax;
      currentValues[YAYOI_COL.CREDIT_AMT] = updateData.creditAmount;
      currentValues[YAYOI_COL.DESC] = updateData.description;

      // ステータスと背景色の更新
      if (action === 'exclude') {
        currentValues[YAYOI_COL.STATUS] = '対象外';
        range.setBackground("#e0e0e0"); // 灰色
      } else {
        // save の場合はステータスを確認済等に
        currentValues[YAYOI_COL.STATUS] = updateData.status || '確認済';
        if (currentValues[YAYOI_COL.STATUS] === '確認済') {
          range.setBackground(null); // 背景色クリア
        }
      }

      range.setValues([currentValues]);
    }
  }

  // 移動先の行を決定
  let nextRow = rowIndex;
  if (action === 'back') {
    // 1つ前の行に移動（getPopupData がそこから上に遡って Head を特定します）
    nextRow = rowIndex - 1;
  } else {
    // 複数明細の場合は、更新した明細の行数分だけスキップする
    const step = (updateData && updateData.details) ? updateData.details.length : 1;
    nextRow = rowIndex + step;
  }

  const maxRow = sheet.getLastRow();
  if (nextRow < 2 || nextRow > maxRow) {
    return null; // 端に到達した
  }

  // シート側の選択行も移動させる
  sheet.getRange(nextRow, 1).activate();
  const nextData = getPopupData(nextRow);

  return nextData;
}

/**
 * freeeモードでの明細行プレビュー削除
 * @param {number} rowIndex 取引の先頭行
 * @param {number} detailIndex 削除する明細のインデックス(0から)
 * @returns {object|null} 再読み込みした行データ、データが空になった場合はnull
 */
function deleteFreeeDetailRow(rowIndex, detailIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('freee取引データ');
  const targetRow = rowIndex + detailIndex;

  if (detailIndex === 0) {
    const currentRowData = sheet.getRange(targetRow, 1, 1, 19).getValues()[0];

    const maxRow = sheet.getLastRow();
    if (targetRow < maxRow) {
      const nextRowRange = sheet.getRange(targetRow + 1, 1, 1, 19);
      const nextRowData = nextRowRange.getValues()[0];

      // 次の行が同じ取引の明細行である場合（収支が空の場合）
      if (nextRowData[FREEE_COL.INCOME_EXPENSE] === "") {
        // 先頭行の情報を引き継ぐ
        nextRowData[FREEE_COL.INCOME_EXPENSE] = currentRowData[FREEE_COL.INCOME_EXPENSE];
        nextRowData[FREEE_COL.ACCRUAL_DATE] = currentRowData[FREEE_COL.ACCRUAL_DATE];
        nextRowData[FREEE_COL.PARTNER] = currentRowData[FREEE_COL.PARTNER];
        nextRowData[FREEE_COL.REG_NUM] = currentRowData[FREEE_COL.REG_NUM];
        nextRowData[FREEE_COL.PAY_STATUS] = currentRowData[FREEE_COL.PAY_STATUS];
        nextRowData[FREEE_COL.PAY_DATE] = currentRowData[FREEE_COL.PAY_DATE];
        nextRowData[FREEE_COL.WALLET] = currentRowData[FREEE_COL.WALLET];
        nextRowData[FREEE_COL.GUESS_DOC] = currentRowData[FREEE_COL.GUESS_DOC];
        nextRowData[FREEE_COL.GUESS_PAY] = currentRowData[FREEE_COL.GUESS_PAY];
        nextRowData[FREEE_COL.WARNING] = currentRowData[FREEE_COL.WARNING];
        nextRowData[FREEE_COL.FILE_ID] = currentRowData[FREEE_COL.FILE_ID];
        nextRowData[FREEE_COL.STATUS] = currentRowData[FREEE_COL.STATUS];

        const bg = sheet.getRange(targetRow, 1, 1, 19).getBackgrounds()[0];
        nextRowRange.setValues([nextRowData]);
        nextRowRange.setBackgrounds([bg]);
      }
    }
  }

  sheet.deleteRow(targetRow);

  // 削除後に再読み込み
  // もし1行しかなくてそれが削除された場合、rowIndexには次の取引が来ているはずなのでそのまま読める
  const maxRowPostDelete = sheet.getLastRow();
  if (rowIndex > maxRowPostDelete || rowIndex < 2) {
    return null; // 全てなくなった場合
  }

  // rowIndexの行を選択しておく
  sheet.getRange(rowIndex, 1).activate();
  return getPopupData(rowIndex);
}

/**
 * freeeモードでの明細行プレビュー追加
 * @param {number} rowIndex 取引の先頭行
 * @param {number} detailIndex 追加元となる明細のインデックス(0から)
 * @returns {object|null} 再読み込みした行データ
 */
function insertFreeeDetailRow(rowIndex, detailIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('freee取引データ');
  const targetRow = rowIndex + detailIndex;

  sheet.insertRowAfter(targetRow);

  // 追加元の明細行の背景色を引き継ぐ (任意)
  try {
    const bg = sheet.getRange(targetRow, 1, 1, 19).getBackgrounds()[0];
    sheet.getRange(targetRow + 1, 1, 1, 19).setBackgrounds([bg]);
  } catch (e) {
    // 無視
  }

  // rowIndexの行を選択しておく
  sheet.getRange(rowIndex, 1).activate();
  return getPopupData(rowIndex);
}

/**
 * シートが手動編集されたときの自動トリガー
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();
  if (sheetName.startsWith('設定') || sheetName.startsWith('マスタ')) {
    ConfigService.clearCache();
  }
}
