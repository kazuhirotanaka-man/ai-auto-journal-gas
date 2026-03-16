/**
 * 設定シート等で定義された「名前付き範囲」から設定値を取得するサービス
 */
const ConfigService = {
  /**
   * 名前付き範囲から単一のセルの値を取得する
   * 範囲が見つからない場合はエラーではなく空文字を返す（任意項目のため）
   * @param {string} rangeName 
   * @returns {string}
   */
  getValue: function(rangeName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName(rangeName);
    if (!range) {
      return ""; // 必須項目ではないものも多いためエラーにせず空文字を返す
    }
    return range.getValue();
  },

  /**
   * 名前付き範囲から1次元配列（リスト）を取得する（空文字やnullを除外）
   * 範囲が見つからない場合は空配列を返す
   * @param {string} rangeName 
   * @returns {string[]}
   */
  getList: function(rangeName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName(rangeName);
    if (!range) {
      return []; // リストも任意の場合があるため空配列を返す
    }
    const values = range.getValues();
    // 1次元配列平坦化し、空の項目を除外
    return values.flat().filter(v => v !== "" && v != null);
  },

  /**
   * 現在の全ての設定オブジェクトを取得する
   */
  getAllConfig: function() {
    return {
      companyName: this.getValue("設定_自社名"),
      industryType: this.getValue("設定_業種"),
      businessDetails: this.getValue("設定_事業内容"),
      folderId: this.extractFolderId(this.getValue("設定_フォルダURL")),
      accountingSoftware: this.getValue("設定_会計ソフト"),
      extraPrompt: this.getValue("設定_追加プロンプト"),
      accountsList: this.getList("設定_勘定科目リスト"),
      subAccountsList: this.getList("設定_補助科目リスト"),
      taxCategoryList: this.getList("設定_税区分リスト")
    };
  },

  /**
   * フォルダURLからIDを抽出する（既にIDの場合はそのまま返す）
   * @param {string} urlOrId
   */
  extractFolderId: function(urlOrId) {
    if (!urlOrId) return "";
    const match = urlOrId.match(/folders\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : urlOrId;
  }
};
