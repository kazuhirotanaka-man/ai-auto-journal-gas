/**
 * freeeから各種マスタデータを取得し、指定されたシートに書き出す機能
 */
function fetchFreeeMasters() {
  const ui = getSafeUi();
  const companyId = PropertiesService.getUserProperties().getProperty('selectedCompanyId');

  if (!companyId) {
    ui.alert('エラー', '操作対象の事業所が選択されていません。\nメニュー「freee連携」＞「操作事業所の選択」から設定してください。', ui.ButtonSet.OK);
    return;
  }

  const confirmResult = ui.alert(
    '確認',
    'freeeから各種マスタ情報（勘定科目、税区分、口座、取引先、品目、部門、メモタグ）を取得し、シートを上書きします。\n（件数が多い場合は時間がかかることがあります）\n\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );

  if (confirmResult !== ui.Button.OK) {
    return;
  }

  // 取得中はUIがブロックされるため、Toastで通知（可能であれば）
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('マスタ取得を開始しました。完了するまでお待ちください...', '処理中', -1);
  } catch (e) {
    // UIがないなどで失敗した場合は無視
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let currentStepNum = 0;
    const totalSteps = 7;
    let currentItemName = '';

    const updateToast = () => {
      try {
        ss.toast(`現在: ${currentItemName} を取得しています...`, `マスタ取得状況 ( ${currentStepNum} / ${totalSteps} )`, 60);
        SpreadsheetApp.flush();
      } catch (e) {}
    };

    // 取得時の共通パラメータ (可能な限り多くの件数を一度に取得するため limit: 3000 を設定)
    const queryParams = { company_id: companyId, limit: 3000 };

    // 1. 勘定科目 (Account Items)
    currentStepNum = 1; currentItemName = '勘定科目'; updateToast();
    Logger.log('勘定科目を取得中...');
    const accountItemsRes = callFreeeApi('get', '/api/1/account_items', queryParams, null);
    if (accountItemsRes && accountItemsRes.account_items) {
      const activeAccountItems = accountItemsRes.account_items.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee勘定科目', activeAccountItems, [
        'id', 'name', 'shortcut', 'shortcut_num', 'tax_code', 'account_category', 'account_category_id', 'group_name'
      ]);
    }

    // 2. 税区分 (Taxes) - 税区分はエンドポイントがパスパラメータで、limitは通常不要
    currentStepNum = 2; currentItemName = '税区分'; updateToast();
    Logger.log('税区分を取得中...');
    const taxesRes = callFreeeApi('get', `/api/1/taxes/companies/${companyId}`, null, null);
    if (taxesRes && taxesRes.taxes) {
      const activeTaxes = taxesRes.taxes.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee税区分', activeTaxes, [
        'code', 'name_ja', 'name', 'display_name', 'available'
      ]);
    }

    // 3. 口座 (Walletables)
    currentStepNum = 3; currentItemName = '口座'; updateToast();
    Logger.log('口座を取得中...');
    const walletablesRes = callFreeeApi('get', '/api/1/walletables', { company_id: companyId }, null);
    if (walletablesRes && walletablesRes.walletables) {
      const activeWalletables = walletablesRes.walletables.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee口座', activeWalletables, [
        'id', 'name', 'type', 'bank_id'
      ]);
    }

    // 4. 取引先 (Partners)
    currentStepNum = 4; currentItemName = '取引先'; updateToast();
    Logger.log('取引先を取得中...');
    const partnersRes = callFreeeApi('get', '/api/1/partners', queryParams, null);
    if (partnersRes && partnersRes.partners) {
      const activePartners = partnersRes.partners.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee取引先', activePartners, [
        'id', 'name', 'code', 'name_kana', 'partner_doc_friendly_name'
      ]);
    }

    // 5. 品目 (Items)
    currentStepNum = 5; currentItemName = '品目'; updateToast();
    Logger.log('品目を取得中...');
    const itemsRes = callFreeeApi('get', '/api/1/items', queryParams, null);
    if (itemsRes && itemsRes.items) {
      const activeItems = itemsRes.items.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee品目', activeItems, [
        'id', 'name'
      ]);
    }

    // 6. 部門 (Departments / Sections in freee API)
    currentStepNum = 6; currentItemName = '部門'; updateToast();
    Logger.log('部門を取得中...');
    const sectionsRes = callFreeeApi('get', '/api/1/sections', queryParams, null);
    if (sectionsRes && sectionsRes.sections) {
      const activeSections = sectionsRes.sections.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreee部門', activeSections, [
        'id', 'name', 'long_name', 'company_id'
      ]);
    }

    // 7. メモタグ (Tags)
    currentStepNum = 7; currentItemName = 'メモタグ'; updateToast();
    Logger.log('メモタグを取得中...');
    const tagsRes = callFreeeApi('get', '/api/1/tags', queryParams, null);
    if (tagsRes && tagsRes.tags) {
      const activeTags = tagsRes.tags.filter(item => item.available !== false);
      writeToSheet(ss, 'マスタfreeeメモタグ', activeTags, [
        'id', 'name'
      ]);
    }

    // Toastを消す(別のメッセージで上書き)
    try {
      ss.toast('すべてのマスタ情報の取得・シート反映が完了しました。', '完了', 5);
    } catch (e) {}

    ui.alert('完了', 'すべてのマスタの取得とシートへの書き出しが完了しました。', ui.ButtonSet.OK);

  } catch (e) {
    Logger.log('マスタ取得エラー: ' + e.message);
    ui.alert('エラー', 'マスタ情報の取得中にエラーが発生しました。\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 取得したデータを指定のシートに書き出す（既存データはクリアされます）
 * @param {Spreadsheet} ss スプレッドシートオブジェクト
 * @param {string} sheetName 書き込み先のシート名
 * @param {Array<Object>} dataArr 書き込むオブジェクトの配列
 * @param {Array<string>} columns 出力するプロパティ名の配列（ヘッダーとしても使用）
 */
function writeToSheet(ss, sheetName, dataArr, columns) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // シートが存在しない場合は新規作成
    sheet = ss.insertSheet(sheetName);
  }

  // 既存のデータをクリア
  sheet.clear();

  if (!dataArr || dataArr.length === 0) {
    sheet.getRange(1, 1).setValue('データがありません');
    return;
  }

  // ヘッダー行をセット
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);

  // データ行を生成
  const rows = dataArr.map(item => {
    return columns.map(col => {
      const val = item[col];
      // null や undefined は空文字にする
      if (val === null || val === undefined) {
        return '';
      }
      // オブジェクトや配列はJSON文字列に変換
      if (typeof val === 'object') {
        return JSON.stringify(val);
      }
      // 文字列の前に ' をつけないとGASが日付や数値として誤認する場合があるが、
      // 基本的にはAPIの生データをそのまま書き出す
      return val;
    });
  });

  // データをシートに書き込む
  // rows は 2次元配列
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, columns.length).setValues(rows);
  }

  // 見やすくするためにヘッダーを太字＆背景色、列幅を自動調整等の処理を任意で追加
  const headerRange = sheet.getRange(1, 1, 1, columns.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  // フィルターを設定
  if (sheet.getFilter() != null) {
    sheet.getFilter().remove();
  }
  sheet.getDataRange().createFilter();
}
