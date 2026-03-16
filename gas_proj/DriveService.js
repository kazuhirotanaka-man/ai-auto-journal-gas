/**
 * Google Driveから画像ファイルを取得し、処理状態を管理するサービス
 */
const DriveServiceObj = {
  
  /**
   * 指定フォルダ内のすべての画像ファイル（サブフォルダ含む）を取得する
   * @param {string} folderId 
   * @returns {GoogleAppsScript.Drive.File[]} 画像ファイルの配列
   */
  getAllImageFiles: function(folderId) {
    if (!folderId) throw new Error("フォルダIDが指定されていません。設定シートを確認してください。");
    const folder = DriveApp.getFolderById(folderId);
    let allFiles = [];
    this._getImagesRecursively(folder, allFiles);
    return allFiles;
  },

  _getImagesRecursively: function(folder, filesArray) {
    // 現在のフォルダ内のすべてのファイルを取得
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      filesArray.push(file);
    }

    // サブフォルダを再帰的に検索
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      this._getImagesRecursively(subFolder, filesArray);
    }
  },

  /**
   * システムログ（名前付き範囲）から処理済みファイルIDのリストを取得する
   * @returns {string[]}
   */
  getProcessedFileIds: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName("ログ_処理済みファイルIDs");
    if (!range) return [];
    
    const values = range.getValues();
    return values.flat().filter(v => v !== "" && v != null).map(String);
  },

  /**
   * 未処理のファイルだけを抽出する
   * @param {string} folderId 
   * @returns {GoogleAppsScript.Drive.File[]} 未処理のファイル配列
   */
  getUnprocessedFiles: function(folderId) {
    const processedIds = this.getProcessedFileIds();
    const allFiles = this.getAllImageFiles(folderId);
    
    // システムログに存在しないファイルIDのみフィルタリング
    return allFiles.filter(file => !processedIds.includes(file.getId()));
  }
};
