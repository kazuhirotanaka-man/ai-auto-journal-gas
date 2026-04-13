/**
 * 設定シート等で定義された「名前付き範囲」から設定値を取得するサービス
 */
const ConfigService = {
  _cache: null,
  /**
   * 名前付き範囲から単一のセルの値を取得する
   * 範囲が見つからない場合はエラーではなく空文字を返す（任意項目のため）
   * @param {string} rangeName 
   * @returns {string}
   */
  getValue: function (rangeName) {
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
  getList: function (rangeName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName(rangeName);
    if (!range) {
      return []; // リストも任意の場合があるため空配列を返す
    }
    const values = range.getValues();
    // 1次元配列平坦化し、空の項目を除外
    if (values.length === 0) {
      return [];
    } else if (values[0].length === 1) {
      return values.flat().filter(v => v !== "" && v != null);
    } else {
      return values.filter(v => v.some(vv => vv !== "" && vv != null));
    }
  },

  /**
   * freeeマスタシートからB列（2列目）のリストを取得する
   * @param {string} sheetName
   * @returns {string[]}
   */
  getFreeeMasterList: function (sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // B列 (列番号2) のデータを取得
    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    return values.flat().filter(v => v !== "" && String(v).trim() !== "");
  },

  /**
   * 現在の全ての設定オブジェクトを取得する
   */
  getAllConfig: function () {
    if (this._cache) return this._cache;

    const cache = CacheService.getScriptCache();
    const cachedConfig = cache.get('APP_CONFIG_CACHE');
    if (cachedConfig) {
      try {
        this._cache = JSON.parse(cachedConfig);
        return this._cache;
      } catch (e) {
        // パースエラーの場合は再取得へフォールバック
      }
    }

    const config = {
      companyName: this.getValue("設定_自社名"),
      industryType: this.getValue("設定_業種"),
      businessDetails: this.getValue("設定_事業内容"),
      folderId: this.extractFolderId(this.getValue("設定_フォルダURL")),
      accountingSoftware: this.getValue("設定_会計ソフト"),
      extraPrompt: this.getValue("設定_追加プロンプト"),
      accountsList: this.getList("設定_勘定科目リスト"),
      subAccountsList: this.getList("設定_補助科目リスト"),
      accountsAndSubAccountsList: this.getList("設定_科目・補助科目"),
      taxCategoryList: this.getList("設定_税区分リスト")
    };

    if (config.accountingSoftware === "freee会計") {
      config.freeeAccountsList = this.getFreeeMasterList("マスタfreee勘定科目");
      config.freeeTaxCategoryList = this.getFreeeMasterList("マスタfreee税区分");
      config.freeeWalletsList = this.getFreeeMasterList("マスタfreee口座");
      config.freeePartnersList = this.getFreeeMasterList("マスタfreee取引先");
      config.freeeItemsList = this.getFreeeMasterList("マスタfreee品目");
      config.freeeDepartmentsList = this.getFreeeMasterList("マスタfreee部門");
      config.freeeTagsList = this.getFreeeMasterList("マスタfreeeメモタグ");
    }

    this._cache = config;
    try {
      cache.put('APP_CONFIG_CACHE', JSON.stringify(config), 3600); // 1時間のキャッシュ
    } catch (e) {}

    return config;
  },

  /**
   * キャッシュをクリアする
   */
  clearCache: function () {
    this._cache = null;
    try {
      CacheService.getScriptCache().remove('APP_CONFIG_CACHE');
    } catch (e) {}
  },

  /**
   * フォルダURLからIDを抽出する（既にIDの場合はそのまま返す）
   * @param {string} urlOrId
   */
  extractFolderId: function (urlOrId) {
    if (!urlOrId) return "";
    const match = urlOrId.match(/folders\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : urlOrId;
  }
};
