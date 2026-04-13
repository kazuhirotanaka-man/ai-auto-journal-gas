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
 * @param {number} rowIndex 取得する行番号（指定がない場合は現在のアクティブセル行）
 * @returns {object} 行データと設定情報のオブジェクト
 */
function getPopupData(rowIndex) {
  console.time(`getPopupData-total`);
  console.time(`get-ConfigService`);
  const config = ConfigService.getAllConfig();
  console.timeEnd(`get-ConfigService`);
  
  const isFreee = config.accountingSoftware === "freee会計";
  const sheetName = isFreee ? 'freee取引データ' : '仕訳データ';
  const dataColumns = isFreee ? 19 : 13;
  const fileIdColIndex = isFreee ? 17 : 11; // 0-indexed

  console.time(`get-SheetAccess`);
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
  if (!rowData[0] && !rowData[fileIdColIndex]) return null; // データが空

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
    const fileId = headRowData[17] ? String(headRowData[17]) : null;
    console.timeEnd(`get-SheetAccess`);

    // Head行から下方向に、同じfileIdかつ収支が空白（＝同一取引の明細）である行を取得し続ける
    console.time(`get-loopDetails`);
    let detailsData = [];
    let currentRow = targetRow;
    
    while (currentRow <= maxRow) {
      const rowVars = sheet.getRange(currentRow, 1, 1, dataColumns).getValues()[0];
      const rFileId = rowVars[17] ? String(rowVars[17]) : null;
      
      // 最初（Head行）は無条件で追加。2行目以降は 収支(rowVars[0])が空白 なら同じ取引の明細とみなす。
      if (currentRow === targetRow || rowVars[0] === "") {
        detailsData.push(rowVars);
      } else {
        break; // 別の取引が始まったら終了
      }
      currentRow++;
    }
    console.timeEnd(`get-loopDetails`);

    const rowCount = detailsData.length;

    console.time(`get-FormatData`);
    const parseDateRaw = function(rawDate) {
      if (rawDate instanceof Date) {
        return Utilities.formatDate(rawDate, ss.getSpreadsheetTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
      } else if (rawDate !== "" && rawDate != null) {
        // AI出力が "2024/4/1" などゼロ埋めされていない場合への対応
        const d = new Date(rawDate);
        if (!isNaN(d.getTime())) {
          return Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
        }
        return String(rawDate).replace(/\//g, "-");
      }
      return "";
    };

    const retData = {
      startRowIndex: targetRow,
      rowCount: rowCount,
      fileId: fileId,
      data: {
        transaction: {
          incomeExpense: headRowData[0],
          date: parseDateRaw(headRowData[1]),
          partner: headRowData[2],
          registrationNumber: headRowData[3],
          paymentStatus: headRowData[4],
          paymentDate: parseDateRaw(headRowData[5]),
          wallet: headRowData[6],
          status: headRowData[18] || '未確認',
          warningText: headRowData[16],
          guessedDocumentType: headRowData[14],
          guessedPaymentMethod: headRowData[15]
        },
        details: detailsData.map(row => ({
          accountItem: row[7],
          amount: row[8],
          taxCategory: row[9],
          item: row[10],
          department: row[11],
          memoTag: row[12],
          description: row[13]
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
    console.timeEnd(`get-FormatData`);
    console.timeEnd(`getPopupData-total`);
    return retData;
  } else {
    // 弥生用ロジック
    console.timeEnd(`get-SheetAccess`);
    console.time(`get-FormatData`);
    // 日付のフォーマット処理
    let dateVal = rowData[0];
    if (dateVal instanceof Date) {
      dateVal = Utilities.formatDate(dateVal, ss.getSpreadsheetTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
    } else if (dateVal !== "" && dateVal != null) {
      const d = new Date(dateVal);
      if (!isNaN(d.getTime())) {
        dateVal = Utilities.formatDate(d, ss.getSpreadsheetTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        dateVal = String(dateVal).replace(/\//g, "-");
      }
    }

    const retData = {
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
    console.timeEnd(`get-FormatData`);
    console.timeEnd(`getPopupData-total`);
    return retData;
  }
}

/**
 * ポップアップからの更新リクエストを処理し、次の行があればそのデータを返す
 * @param {number} rowIndex 更新対象の行
 * @param {object} updateData 更新するデータ
 * @param {string} action アクション ('save', 'exclude', 'back', 'justNext')
 * @returns {object} 次(または前)の行のデータ、なければnull
 */
function updateAndProcessNext(rowIndex, updateData, action) {
  console.time(`updateAndProcessNext-total`);

  console.time(`update-ConfigService`);
  const config = ConfigService.getAllConfig();
  console.timeEnd(`update-ConfigService`);

  const isFreee = config.accountingSoftware === "freee会計";
  const sheetName = isFreee ? 'freee取引データ' : '仕訳データ';
  const dataColumns = isFreee ? 19 : 13;

  console.time(`update-SheetUpdate`);
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
          currentValues[0] = updateData.transaction.incomeExpense;
          currentValues[1] = updateData.transaction.date;
          currentValues[2] = updateData.transaction.partner;
          currentValues[3] = updateData.transaction.registrationNumber;
          currentValues[4] = updateData.transaction.paymentStatus;
          currentValues[5] = updateData.transaction.paymentDate;
          currentValues[6] = updateData.transaction.wallet;

          if (action === 'exclude') {
            currentValues[18] = '対象外';
          } else {
            currentValues[18] = updateData.status || '確認済';
          }
        } else {
          // 明細行は取引基本項目および推測系・状態管理項目を空白にする
          currentValues[0] = "";
          currentValues[1] = "";
          currentValues[2] = "";
          currentValues[3] = "";
          currentValues[4] = "";
          currentValues[5] = "";
          currentValues[6] = "";

          currentValues[14] = ""; // 推測証票種別
          currentValues[15] = ""; // 推測決済方法
          currentValues[16] = ""; // AI解析精度
          currentValues[17] = ""; // ファイルID
          currentValues[18] = ""; // ステータス
        }

        const detail = updateData.details[i];
        currentValues[7] = detail.accountItem;
        currentValues[8] = detail.amount;
        currentValues[9] = detail.taxCategory;
        currentValues[10] = detail.item;
        currentValues[11] = detail.department;
        currentValues[12] = detail.memoTag;
        currentValues[13] = detail.description;
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

      // データの更新 (1〜10列目と13列目)
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
      } else {
        // save の場合はステータスを確認済等に
        currentValues[12] = updateData.status || '確認済';
        if (currentValues[12] === '確認済') {
          range.setBackground(null); // 背景色クリア
        }
      }
      
      range.setValues([currentValues]);
    }
  }
  console.timeEnd(`update-SheetUpdate`);

  console.time(`update-findNextRow`);
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
    console.timeEnd(`update-findNextRow`);
    console.timeEnd(`updateAndProcessNext-total`);
    return null; // 端に到達した
  }

  // シート側の選択行も移動させる
  sheet.getRange(nextRow, 1).activate();
  console.timeEnd(`update-findNextRow`);

  console.time(`update-getPopupData_call`);
  const nextData = getPopupData(nextRow);
  console.timeEnd(`update-getPopupData_call`);
  
  console.timeEnd(`updateAndProcessNext-total`);
  return nextData;
}

/**
 * freeeモードでの明細行プレビュー削除
 * @param {number} rowIndex 取引の先頭行
 * @param {number} detailIndex 削除する明細のインデックス(0から)
 */
function deleteFreeeDetailRow(rowIndex, detailIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('freee取引データ');
  const targetRow = rowIndex + detailIndex;
  
  if (detailIndex === 0) {
    const currentRowData = sheet.getRange(targetRow, 1, 1, 19).getValues()[0];
    const fileId = currentRowData[17];
    
    const maxRow = sheet.getLastRow();
    if (targetRow < maxRow) {
      const nextRowRange = sheet.getRange(targetRow + 1, 1, 1, 19);
      const nextRowData = nextRowRange.getValues()[0];
      
      // 次の行が同じ取引の明細行である場合（収支が空の場合）
      if (nextRowData[0] === "") {
        // 先頭行の情報を引き継ぐ
        nextRowData[0] = currentRowData[0]; // 収支
        nextRowData[1] = currentRowData[1]; // 発生日
        nextRowData[2] = currentRowData[2]; // 取引先
        nextRowData[3] = currentRowData[3]; // 登録番号
        nextRowData[4] = currentRowData[4]; // 決済ステータス
        nextRowData[5] = currentRowData[5]; // 決済期日
        nextRowData[6] = currentRowData[6]; // 決済口座
        nextRowData[14] = currentRowData[14]; // 推測ドキュメント
        nextRowData[15] = currentRowData[15]; // 推測決済手段
        nextRowData[16] = currentRowData[16]; // 警告テキスト
        nextRowData[17] = currentRowData[17]; // fileId
        nextRowData[18] = currentRowData[18]; // ステータス
        
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
 * シートが手動編集されたときの自動トリガー
 */
function onEdit(e) {
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();
  if (sheetName.startsWith('設定') || sheetName.startsWith('マスタ')) {
    ConfigService.clearCache();
  }
}
