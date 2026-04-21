/**
 * ライセンス認証に関連する機能を提供するサービス
 * （将来的にライブラリとして分離する際の主要コンポーネント）
 */
const LicenseService = {
  // ライセンス管理APIのWeb App URL
  API_ENDPOINT: 'https://script.google.com/macros/s/AKfycbzyAALAzum57v1BR05Gci0GL9YRyZTZqe-N332oFvB4COfXQuA-EHZjOBIomM_VE40g/exec',
  LICENSE_SHEET_NAME: 'License',

  /**
   * ライセンス保持用の非表示シートを取得または作成する
   */
  _getLicenseSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.LICENSE_SHEET_NAME);
    if (!sheet) {
      // 存在しなければ新規作成してすぐ非表示にする
      sheet = ss.insertSheet(this.LICENSE_SHEET_NAME);
      
      // B列以降（2列目から最後まで）を削除
      const maxCols = sheet.getMaxColumns();
      if (maxCols > 1) {
        sheet.deleteColumns(2, maxCols - 1);
      }
      
      // 2行目以降（2行目から最後まで）を削除
      const maxRows = sheet.getMaxRows();
      if (maxRows > 1) {
        sheet.deleteRows(2, maxRows - 1);
      }
      
      sheet.hideSheet();
    }
    return sheet;
  },

  /**
   * 保存されているライセンスキーを取得
   */
  _getSavedKey: function() {
    const sheet = this._getLicenseSheet();
    const val = sheet.getRange("A1").getValue();
    return val ? String(val).trim() : null;
  },

  /**
   * ライセンスキーを保存
   */
  _saveKey: function(key) {
    const sheet = this._getLicenseSheet();
    sheet.getRange("A1").setValue(key);
  },

  /**
   * ライセンスキーの保存をクリア
   */
  _clearKey: function() {
    const sheet = this._getLicenseSheet();
    sheet.getRange("A1").clearContent();
  },

  /**
   * 指定したファイルまたはフォルダが配置されている最上位ドライブのルートIDを取得する
   * @param {string} [targetId] 対象のファイル/フォルダID。省略時は現在のアクティブなスプレッドシートのID。
   * @returns {string} ルートフォルダのID
   */
  getDriveRootId: function(targetId) {
    try {
      const idToFetch = targetId || SpreadsheetApp.getActiveSpreadsheet().getId();
      let fileOrFolder;
      try {
        fileOrFolder = DriveApp.getFileById(idToFetch);
      } catch (e) {
        // 指定されたIDがフォルダである場合のフォールバック
        fileOrFolder = DriveApp.getFolderById(idToFetch);
      }
      
      let parents = fileOrFolder.getParents();
      let rootFolderId = null;
      
      while (parents.hasNext()) {
        let parentFolder = parents.next();
        rootFolderId = parentFolder.getId();
        parents = parentFolder.getParents();
      }
      
      // マイドライブまたは共有ドライブの最上位レベルのフォルダIDを返す
      if (!rootFolderId) {
         rootFolderId = DriveApp.getRootFolder().getId();
      }

      return rootFolderId;
    } catch (e) {
      console.error('ルートIDの取得に失敗しました: ', e.message);
      throw e;
    }
  },

  /**
   * 設定された証票格納フォルダのルートIDが、スプレッドシートのルートIDと一致するかチェックする
   * @param {string} folderId 証票格納フォルダのID
   * @returns {boolean} 一致していればtrue
   */
  isEvidenceFolderValid: function(folderId) {
     if (!folderId) return true; // 設定前はスキップ
     
     try {
       const ssRootId = this.getDriveRootId();
       const folderRootId = this.getDriveRootId(folderId);
       
       Logger.log(`[証票フォルダチェック] SS_Root: ${ssRootId}, Folder_Root: ${folderRootId}`);
       return ssRootId === folderRootId;
     } catch(e) {
       Logger.log('証票フォルダのルートID確認に失敗: ' + e.message);
       return false;
     }
  },

  /**
   * 実際の認証要求とUIを伴うフロー
   * ツールの起動時や実行時に呼び出される。
   * @returns {boolean} 最終的に認証されていれば true
   */
  requireLicense: function() {
    let licenseKey = this._getSavedKey();
    const ui = SpreadsheetApp.getUi();
    
    // 1. すでにキーが保存されていれば自動verify
    if (licenseKey) {
      const isVerified = this.verifyLicense(licenseKey, 'verify');
      if (isVerified) {
        return true;
      }
      // verify失敗（＝別のドライブに悪意をもってコピーされた等）の場合、キーを一旦クリアする
      this._clearKey();
    }

    // 2. キーがない（またはverify失敗）場合、UIから入力を求める
    const promptResponse = ui.prompt(
      '🌟 ライセンス認証',
      'このツールを使用するにはライセンスキーが必要です。\n購入時にお渡ししたキーを入力してください。',
      ui.ButtonSet.OK_CANCEL
    );

    if (promptResponse.getSelectedButton() == ui.Button.OK) {
      licenseKey = promptResponse.getResponseText().trim();
      
      if (!licenseKey) {
         ui.alert('エラー', 'キーが入力されていません。', ui.ButtonSet.OK);
         return false;
      }

      // 1. まず「すでにこの環境(ルートID)で有効なキーか」をverifyで確認する
      let isValid = this.verifyLicense(licenseKey, 'verify');
      
      // 2. もしverifyが通らなければ、未使用の新規キーである可能性にかけてactivateを試みる
      if (!isValid) {
         isValid = this.verifyLicense(licenseKey, 'activate');
      }

      if (isValid) {
         // 成功したら非表示シートに保存
         this._saveKey(licenseKey);
         ui.alert('認証完了', 'ライセンスの認証が完了しました！\nこのファイルや、ここからコピーしたファイルを開く際は、設定が引き継がれるため自動で認証されます。', ui.ButtonSet.OK);
         return true;
      } else {
         ui.alert('認証失敗', '無効なキー、または既に使用されている（別のドライブに紐付いている）キーです。', ui.ButtonSet.OK);
         return false;
      }
    }
    
    // 入力キャンセル時
    return false;
  },

  /**
   * API通信によるライセンスのアクティベーション／検証を行う
   * @param {string} licenseKey ユーザー入力または保存されたライセンスキー
   * @param {string} action 'activate' または 'verify'
   * @returns {boolean} 認証結果
   */
  verifyLicense: function(licenseKey, action = 'verify') {
    if (!licenseKey) {
      return false;
    }

    const rootId = this.getDriveRootId();
    
    const payload = {
      action: action,
      licenseKey: licenseKey,
      rootId: rootId
    };
    
    const options = {
      method: "post",
      payload: JSON.stringify(payload),
      contentType: "application/json",
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(this.API_ENDPOINT, options);
      const result = JSON.parse(response.getContentText());
      
      if (result.status === 'success') {
        return true;
      } else {
        Logger.log('認証エラー: ' + (result.message || '不明なエラー'));
        return false;
      }
    } catch (e) {
      Logger.log('通信エラー: ' + e.message);
      return false;
    }
  }
};
