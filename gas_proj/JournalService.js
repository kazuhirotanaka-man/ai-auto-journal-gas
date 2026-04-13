/**
 * スプレッドシートへのデータ出力とログ記録を担当するサービス
 */
const JournalService = {
  
  /**
   * 解析された仕訳データを「仕訳データ」シートの最下行に追記する
   * @param {object} parsedData AIから返ってきたJSONオブジェクト
   * @param {GoogleAppsScript.Drive.File} file 対象のファイルオブジェクト
   */
  /**
   * 解析された仕訳データを対象のシートの最下行に追記する
   * @param {object} parsedData AIから返ってきたJSONオブジェクト
   * @param {GoogleAppsScript.Drive.File} file 対象のファイルオブジェクト
   * @param {object} config 設定オブジェクト
   */
  appendJournalEntry: function(parsedData, file, config) {
    if (!config) {
      config = ConfigService.getAllConfig();
    }
    
    if (config.accountingSoftware === "freee会計") {
      this.appendFreeeJournalEntry(parsedData, file, config);
    } else {
      this.appendYayoiJournalEntry(parsedData, file, config);
    }
  },

  /**
   * freee形式のデータを「freee取引データ」シートの最下行に追記する
   */
  appendFreeeJournalEntry: function(parsedData, file, config) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("freee取引データ");
    
    if (!sheet) {
      sheet = ss.insertSheet("freee取引データ");
      sheet.appendRow([
        "収支", "発生日", "取引先", "登録番号", "決済ステータス", "決済期日", "決済口座",
        "勘定科目", "金額", "税区分", "品目", "部門", "メモタグ", "備考",
        "推測証票種別", "推測決済方法", "AI解析精度", "ファイルID", "ステータス"
      ]);
    }
    
    // システム上の最終行を取得（17列目: ファイルID列 を基準に検索）
    const columnQ = sheet.getRange("Q:Q").getValues();
    let lastRow = 1;
    for (let i = columnQ.length - 1; i >= 0; i--) {
      if (columnQ[i][0] !== "") {
        lastRow = i + 1;
        break;
      }
    }
    const targetRowIndex = lastRow + 1;

    let isTarget = parsedData.isTarget !== false;

    let entries = parsedData.entries;
    if (!entries || !Array.isArray(entries)) {
      entries = [parsedData]; // 念のため
    }
    
    let rowDataArray = [];
    let bgColors = [];

    let prevGroupKey = null;

    // ヘルパー：リストに存在するかチェック。存在しない場合は空文字にする
    const checkList = (val, list) => {
      if (!val) return "";
      if (!Array.isArray(list)) return val;
      return list.includes(val) ? val : "";
    };

    entries.forEach(entry => {
      let warningText = "";
      let bgColor = null;

      if (!isTarget) {
        warningText = "対象外";
        bgColor = "#e0e0e0";
        entry.remarks = "【対象外】" + (parsedData.description || "");
      } else {
        if (entry.confidence === "低" || entry.confidence === "中") {
          warningText = "要確認 (" + entry.confidence + ")";
          bgColor = "#ffebee";
        } else {
          warningText = "高";
        }
      }
      
      const inOutList = ["収入", "支出"];
      const statusList = ["決済済", "未決済"];

      const incomeExpense = checkList(entry.incomeExpense || (!isTarget ? "" : "支出"), inOutList);
      const accrualDate = entry.accrualDate || "";
      const partner = checkList(entry.partner || "", config.freeePartnersList);
      const paymentStatus = checkList(entry.paymentStatus || (!isTarget ? "" : "決済済"), statusList);
      const paymentDate = entry.paymentDate || "";
      const wallet = paymentStatus === "未決済" ? "" : checkList(entry.wallet || (!isTarget ? "" : "現金"), config.freeeWalletsList);

      const currentGroupKey = `${incomeExpense}_${accrualDate}_${partner}_${paymentStatus}_${paymentDate}_${wallet}`;

      const accountItem = checkList(entry.accountItem || "", config.freeeAccountsList);
      const taxCategory = checkList(entry.taxCategory || "", config.freeeTaxCategoryList);
      const item = checkList(entry.item || "", config.freeeItemsList);
      const department = checkList(entry.department || "", config.freeeDepartmentsList);
      const memoTag = checkList(entry.memoTag || "", config.freeeTagsList);
      const registrationNumber = entry.registrationNumber || "";
      
      let displayRowIncomeExpense = incomeExpense;
      let displayRowAccrualDate = accrualDate;
      let displayRowPartner = partner;
      let displayRowPaymentStatus = paymentStatus;
      let displayRowPaymentDate = paymentDate;
      let displayRowWallet = wallet;
      let displayRowRegistrationNumber = registrationNumber;

      if (prevGroupKey === currentGroupKey) {
        // 同一取引の2行目以降は前項目の6項目を空表示にする
        displayRowIncomeExpense = "";
        displayRowAccrualDate = "";
        displayRowPartner = "";
        displayRowPaymentStatus = "";
        displayRowPaymentDate = "";
        displayRowWallet = "";
        displayRowRegistrationNumber = "";
      } else {
        prevGroupKey = currentGroupKey;
      }

      rowDataArray.push([
        displayRowIncomeExpense, // 1 (A)
        displayRowAccrualDate,   // 2 (B)
        displayRowPartner,       // 3 (C)
        displayRowRegistrationNumber, // 4 (D)
        displayRowPaymentStatus, // 5 (E)
        displayRowPaymentDate,   // 6 (F)
        displayRowWallet,        // 7 (G)
        accountItem,             // 8 (H)
        entry.amount || (!isTarget ? "" : 0), // 9 (I)
        taxCategory,             // 10 (J)
        item,                    // 11 (K)
        department,              // 12 (L)
        memoTag,                 // 13 (M)
        entry.remarks || "",     // 14 (N)
        entry.guessedDocumentType || (!isTarget ? "" : "領収書・レシート"), // 15
        entry.guessedPaymentMethod || (!isTarget ? "" : "現金"),          // 16
        warningText,             // 17 AI解析精度
        file.getId(),            // 18 ファイルID
        "未確認"                  // 19 ステータス
      ]);
      bgColors.push(bgColor);
    });
    
    if (rowDataArray.length > 0) {
      sheet.getRange(targetRowIndex, 1, rowDataArray.length, 19).setValues(rowDataArray);
      
      // 背景色の適用
      bgColors.forEach((color, index) => {
        if (color) {
          sheet.getRange(targetRowIndex + index, 1, 1, 19).setBackground(color);
        }
      });
    }
  },

  /**
   * 弥生形式・汎用形式のデータを「仕訳データ」シートの最下行に追記する
   */
  appendYayoiJournalEntry: function(parsedData, file, config) {
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

    // ヘルパー：リストに存在するかチェック。存在しない場合は空文字にする
    const checkList = (val, list) => {
      if (!val) return "";
      if (!Array.isArray(list)) return val;
      return list.includes(val) ? val : "";
    };

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

      const debitAccount = checkList(entry.debitAccount || "", config.accountsList);
      const debitSubAccount = checkList(entry.debitSubAccount || "", config.subAccountsList);
      const debitTaxCategory = checkList(entry.debitTaxCategory || (!isTarget ? "" : "対象外"), config.taxCategoryList);
      
      const creditAccount = checkList(entry.creditAccount || (!isTarget ? "" : "現金"), config.accountsList);
      const creditSubAccount = checkList(entry.creditSubAccount || "", config.subAccountsList);
      const creditTaxCategory = checkList(entry.creditTaxCategory || (!isTarget ? "" : "対象外"), config.taxCategoryList);
      
      rowDataArray.push([
        entry.date || "",
        debitAccount,
        debitSubAccount,
        debitTaxCategory,
        entry.amount || 0,
        creditAccount,
        creditSubAccount,
        creditTaxCategory,
        entry.amount || 0,
        entry.description || "",
        warningText,
        file.getId(), // ファイルID (12列目)
        "未確認"       // ステータス (13列目)
      ]);
      bgColors.push(bgColor);
    });
    
    // 複数行を一括で書き込む
    if (rowDataArray.length > 0) {
      sheet.getRange(targetRowIndex, 1, rowDataArray.length, 13).setValues(rowDataArray);
      
      // 背景色の適用（行ごとに設定）
      bgColors.forEach((color, index) => {
        if (color) {
          sheet.getRange(targetRowIndex + index, 1, 1, 13).setBackground(color);
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
