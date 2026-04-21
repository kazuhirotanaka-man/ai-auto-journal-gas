/**
 * ライセンス管理用管理者メニュー
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔑 ライセンス管理')
    .addItem('新規ライセンス発行', 'generateNewLicense')
    .addSeparator()
    .addItem('選択中のライセンスをリセット（PC変更・未認証へ）', 'resetSelectedLicense')
    .addItem('選択中のライセンスを無効化（解約扱い）', 'deactivateSelectedLicense')
    .addToUi();
}

/**
 * 新規ライセンスの発行
 */
function generateNewLicense() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Licenses');
  if (!sheet) {
    ui.alert('エラー', 'Licenses シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // UUIDを生成 (ハイフンなしで少し短く読みやすくするかは好みですが一旦そのまま)
  const newKey = Utilities.getUuid();
  const issueDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  
  // A:Key, B:Status, C:RootID, D:Email, E:Office, F:Name, G:備考/日付
  sheet.appendRow([newKey, 'unused', '', '', '', '', '発行: ' + issueDate]);
  
  // 発行された行の背景色などをつけても親切かもしれません
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 7).setBackground('#e8f4f8');
  
  ui.alert('✨ 発行完了', '新しいライセンスキーを最下行に追加しました！\n\n' + newKey, ui.ButtonSet.OK);
}

/**
 * 選択行のライセンスを「未認証(unused)」にリセットする
 */
function resetSelectedLicense() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== 'Licenses') {
    ui.alert('エラー', 'Licenses シートのタブを開いた状態で実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    ui.alert('エラー', '登録されているライセンスの行（2行目以降）をどこかクリックして選択した状態で実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  const key = sheet.getRange(row, 1).getValue();
  if (!key) {
    ui.alert('エラー', '選択した行にライセンスキーがありません。', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🔄 リセット確認', 
    `以下のライセンスを「未認証」状態に戻し、登録情報をクリアしますか？\n（お客様のPC入れ替え時等に使用します）\n\n対象キー: ${key}`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    sheet.getRange(row, 2).setValue('unused'); // Status
    sheet.getRange(row, 3, 1, 4).clearContent(); // C(RootID) 〜 F(Name) を空にする
    sheet.getRange(row, 1, 1, 7).setBackground('#ffffff');
    ui.alert('リセット完了', 'ライセンスをリセットし、再度別のGoogleドライブで認証可能な状態にしました。', ui.ButtonSet.OK);
  }
}

/**
 * 選択行のライセンスを「無効(inactive)」にする
 */
function deactivateSelectedLicense() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== 'Licenses') {
    ui.alert('エラー', 'Licenses シートのタブを開いた状態で実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    ui.alert('エラー', '登録されているライセンスの行を選択して実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  const key = sheet.getRange(row, 1).getValue();
  if (!key) {
    ui.alert('エラー', '選択した行にライセンスキーがありません。', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🚫 無効化・解約確認', 
    `以下のライセンスを解約（無効）扱いにしますか？\n（現在このキーを使用中のツールは次回起動時から停止します）\n\n対象キー: ${key}`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    sheet.getRange(row, 2).setValue('inactive'); // Status
    sheet.getRange(row, 1, 1, 7).setBackground('#fce8e6'); // 無効化されたことがわかるように赤背景
    ui.alert('無効化完了', 'ライセンスを無効化（解約）しました。', ui.ButtonSet.OK);
  }
}
