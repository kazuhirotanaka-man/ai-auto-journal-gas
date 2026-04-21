/**
 * 仕訳データシートの内容を各種会計ソフトのフォーマットでエクスポートするサービス
 */
const ExportService = {
  
  /**
   * 選択されたステータスに応じてCSVを出力し、出力を完了したデータのステータスを「ダウンロード済み」にする
   */
  exportToCsvWithStatuses: function(statuses) {
    const config = ConfigService.getAllConfig();
    const software = config.accountingSoftware || "弥生会計";
    const isFreee = software === "freee会計";
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = isFreee ? "freee取引データ" : "仕訳データ";
    const sheet = ss.getSheetByName(sheetName);
    
    // データのある範囲を取得（ヘッダーを除く）
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { error: "エクスポート対象のデータがありません。" };
    }
    
    const dataColumns = isFreee ? 18 : 13;
    const statusColIndex = isFreee ? 17 : 12; // 0-indexed
    
    // データを取得
    const dataRange = sheet.getRange(2, 1, lastRow - 1, dataColumns);
    const dataList = dataRange.getValues();
    
    let targetRows = [];
    let rowIndicesToUpdate = [];
    
    dataList.forEach((row, index) => {
      const status = row[statusColIndex];
      if (statuses.includes(status)) {
        targetRows.push(row);
        rowIndicesToUpdate.push(index + 2); // データ行は2行目から
      }
    });

    if (targetRows.length === 0) {
      return { error: "選択されたステータスに該当するデータがありません。" };
    }

    let csvData = "";
    
    if (software === "弥生会計") {
      csvData = this.buildYayoiCsv(targetRows);
    } else if (software === "freee会計") {
      return { error: `freee会計の場合は、CSVエクスポート機能ではなく専用のAPIを利用した「取引登録」機能を後日利用する予定です。（現在はエクスポートを行いません）` };
    } else {
      return { error: `会計ソフト「${software}」の出力形式は現在未対応です。` };
    }
    
    // CSVファイルをDriveのルートに作成
    const fileName = `仕訳エクスポート_${software}_${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmmss")}.csv`;
    const blob = Utilities.newBlob("", "text/csv", fileName).setDataFromString(csvData, "Shift_JIS"); // 汎用的にShift-JIS
    const file = DriveApp.createFile(blob);
    
    // 出力対象となった行のステータスを一括で「ダウンロード済み」に変更
    rowIndicesToUpdate.forEach(rowIdx => {
      sheet.getRange(rowIdx, statusColIndex + 1).setValue("ダウンロード済み");
    });

    // ダイアログを表示してダウンロードリンクを提供
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; text-align: center; padding: 20px;">
        <p>Google Driveにエクスポートファイルを作成しました。<br>以下のボタンからダウンロードできます。</p>
        <br>
        <a href="${downloadUrl}" target="_blank" onclick="google.script.host.close()" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">
          ファイルをダウンロード
        </a>
      </div>
    `).setWidth(350).setHeight(270);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'エクスポート完了');
    return { success: true };
  },
  
  /**
   * 弥生会計のインポート形式（汎用形式）に変換
   * 識別フラグ, 伝票No, 決算, 取引日付, 借方勘定科目, 借方補助科目, 借方部門, 借方税区分, 借方金額, 借方税金額, 貸方勘定科目, 貸方補助科目, 貸方部門, 貸方税区分, 貸方金額, 貸方税金額, 摘要, ...
   * @param {Array[]} dataList 
   */
  buildYayoiCsv: function(dataList) {
    let lines = [];
    
    dataList.forEach(row => {
      // rowの構成: 0:日付, 1:借方科目, 2:借方補助, 3:借方税区分, 4:借方金額, 5:貸方科目, 6:貸方補助, 7:貸方税区分, 8:貸方金額, 9:摘要
      const dateRaw = row[0];
      let formattedDate = "";
      if (dateRaw instanceof Date) {
        formattedDate = Utilities.formatDate(dateRaw, "Asia/Tokyo", "yyyy/MM/dd");
      } else {
        formattedDate = dateRaw;
      }
      
      const csvRow = [
        "2000",          // 識別フラグ (2000: 仕訳データ)
        "",              // 伝票No
        "",              // 決算
        formattedDate,   // 取引日付
        row[1] || "",    // 借方勘定科目
        row[2] || "",    // 借方補助科目
        "",              // 借方部門
        row[3] || "対象外", // 借方税区分
        row[4] || "0",   // 借方金額
        "0",             // 借方税金額
        row[5] || "",    // 貸方勘定科目
        row[6] || "",    // 貸方補助科目
        "",              // 貸方部門
        row[7] || "対象外", // 貸方税区分
        row[8] || "0",   // 貸方金額
        "0",             // 貸方税金額
        row[9] || "",    // 摘要
        "0",             // 摘要コードなど以降は省略・デフォルト値
        "3",             // 帳簿区分
        "0",             // 本支店区分表示
        "", "", "", "", ""
      ];
      
      // 各フィールドをダブルクォートで囲み、カンマで結合
      const lineStr = csvRow.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",");
      lines.push(lineStr);
    });
    
    // ヘッダーなしでデータ行から出力するのが弥生形式の基本ですが、必要に応じて追加
    return lines.join("\r\n");
  }
};
