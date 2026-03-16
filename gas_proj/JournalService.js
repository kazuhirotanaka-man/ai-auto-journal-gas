/**
 * スプレッドシートへのデータ出力とログ記録を担当するサービス
 */
const JournalService = {
  
  /**
   * 解析された仕訳データを「仕訳データ」シートの最下行に追記する
   * @param {object} parsedData AIから返ってきたJSONオブジェクト
   * @param {GoogleAppsScript.Drive.File} file 対象のファイルオブジェクト
   */
  appendJournalEntry: function(parsedData, file) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("仕訳データ");
    
    // システム上の最終行を取得（12列目: ファイルID列 を基準に検索）
    const columnL = sheet.getRange("L:L").getValues();
    let lastRow = 1;
    for (let i = columnL.length - 1; i >= 0; i--) {
      if (columnL[i][0] !== "") {
        lastRow = i + 1;
        break;
      }
    }
    const targetRowIndex = lastRow + 1;

    let isTarget = parsedData.isTarget !== false;

    // もし既存の単一出力などで entries 配列がない場合は互換性のため配列化する
    let entries = parsedData.entries;
    if (!entries || !Array.isArray(entries)) {
      entries = [parsedData];
    }
    
    let rowDataArray = [];
    let bgColors = [];

    entries.forEach(entry => {
      let warningText = "";
      let bgColor = null;

      if (!isTarget) {
        warningText = "対象外";
        bgColor = "#e0e0e0"; // 灰色
        entry.description = "【対象外】" + (entry.description || "");
      } else {
        if (entry.confidence === "低" || entry.confidence === "中") {
          warningText = "要確認 (" + entry.confidence + ")";
          bgColor = "#ffebee"; // 薄い赤
        } else {
          warningText = "高"; // 高い場合も明記してユーザーを安心させる
        }
      }
      
      rowDataArray.push([
        entry.date || "",
        entry.debitAccount || "",
        entry.debitSubAccount || "",
        entry.debitTaxCategory || (!isTarget ? "" : "対象外"),
        entry.amount || 0,
        entry.creditAccount || (!isTarget ? "" : "現金"),
        entry.creditSubAccount || "",
        entry.creditTaxCategory || (!isTarget ? "" : "対象外"),
        entry.amount || 0,
        entry.description || "",
        warningText,
        file.getId() // ファイルID (12列目)
      ]);
      bgColors.push(bgColor);
    });
    
    // 複数行を一括で書き込む
    if (rowDataArray.length > 0) {
      sheet.getRange(targetRowIndex, 1, rowDataArray.length, 12).setValues(rowDataArray);
      
      // 背景色の適用（行ごとに設定）
      bgColors.forEach((color, index) => {
        if (color) {
          sheet.getRange(targetRowIndex + index, 1, 1, 12).setBackground(color);
        }
      });
    }
  },

  /**
   * 処理が完了したファイルIDを「システムログ」シートに記録する
   * @param {GoogleAppsScript.Drive.File} file 
   */
  logProcessedFile: function(file) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("システムログ");
    
    const now = new Date();
    // A列: ファイルID, B列: 日時, C列: ファイル名(備考)
    sheet.appendRow([file.getId(), now, file.getName()]);
  }
};
