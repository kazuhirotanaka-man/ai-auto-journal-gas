/**
 * 【サンプル】freee APIから事業所一覧を取得してシートに書き出す
 * 
 * この関数は、callFreeeApi ユーティリティを使用して事業所一覧を取得し、
 * 現在アクティブなシートにその内容を書き出すサンプルです。
 * 
 * この関数を参考に、新しい業務ロジックを作成してください。
 */
function getCompanies() {
  try {
    // callFreeeApi を使ってAPIを呼び出す
    const result = callFreeeApi(
      'get',              // HTTPメソッド
      '/api/1/companies', // APIのエンドポイントパス
      null,               // URLクエリパラメータ (不要な場合はnull)
      null                // リクエストボディ (不要な場合はnull)
    );

    // レスポンスのハンドリング
    if (!result || !result.companies || result.companies.length === 0) {
      getSafeUi().alert('事業所が見つかりませんでした。');
      return;
    }

    const companies = result.companies;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("事業所一覧") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("事業所一覧");
    sheet.clear(); 

    // シートへの書き出し
    const headers = ['事業所ID', '表示名', '事業所名', '事業所名カナ', '事業所番号', 'ロール'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const rows = companies.map(c => [c.id, c.display_name, c.name, c.name_kana, c.company_number, c.role]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    getSafeUi().alert('事業所一覧を取得し、シートに書き出しました。');

  } catch (e) {
    // エラーハンドリング
    getSafeUi().alert('エラーが発生しました: ' + e.message);
  }
}

/**
 * プロパティに保存されている事業所IDを取得する
 * @return {string} 選択されている事業所ID
 */
function getSelectedCompanyId() {
  const companyId = PropertiesService.getUserProperties().getProperty('selectedCompanyId');
  if (!companyId) {
    throw new Error('操作対象の事業所が選択されていません。\nメニューの「操作事業所の選択」から設定してください。');
  }
  return companyId;
}

/**
 * 【デモ】取引の登録 → 取得 → 削除 → 再取得 を一連の流れで実行する
 */
function runDealLifecycleSample() {
  const ui = getSafeUi();
  if (ui.alert('取引の登録・取得・削除のデモを実行します。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) {
    return;
  }

  try {
    const companyId = getSelectedCompanyId();
    Logger.log(`対象の事業者ID: ${companyId}`);

    // --- ① 取引の登録（POST）---
    Logger.log('--- ① 取引の登録（POST）---');
    
    // 取引登録に必要な情報をAPIから動的に取得
    Logger.log('取引登録に必要な情報をAPIから取得します...');
    const accountItemsRes = callFreeeApi('get', '/api/1/account_items', { company_id: companyId }, null);
    const salesAccountItem = accountItemsRes.account_items.find(item => item.name === '売上高');
    if (!salesAccountItem) throw new Error('勘定科目「売上高」が見つかりませんでした。');
    const accountItemId = salesAccountItem.id;
    Logger.log(`勘定科目「売上高」のID: ${accountItemId} を使用します。`);

    const taxesRes = callFreeeApi('get', `/api/1/taxes/companies/${companyId}`, null, null);
    const salesTax = taxesRes.taxes.find(tax => tax.name_ja === '課税売上10%');
    if (!salesTax) throw new Error('税区分「課税売上10%」が見つかりませんでした。');
    const taxCode = salesTax.code;
    Logger.log(`税区分「課税売上10%」のコード: ${taxCode} を使用します。`);

    const today = new Date().toISOString().slice(0, 10);
    const dealBody = {
      company_id: companyId,
      issue_date: today,
      type: 'income', // 収入
      details: [
        {
          account_item_id: accountItemId,
          amount: 1000,
          tax_code: taxCode
        }
      ]
    };

    const createdDeal = callFreeeApi('post', '/api/1/deals', null, dealBody);
    const dealId = createdDeal.deal.id;
    Logger.log(`取引を作成しました。取引ID: ${dealId}`);
    ui.alert(`取引を作成しました。\n取引ID: ${dealId}`);

    // ② 取引の取得（GET）
    Logger.log('\n--- ② 登録直後の取引一覧を取得（GET）---');
    const dealsBeforeDelete = callFreeeApi('get', '/api/1/deals', { company_id: companyId }, null);
    Logger.log('現在の取引一覧:', dealsBeforeDelete.deals);
    ui.alert('登録直後の取引一覧をログに出力しました。');

    // ③ 取引の削除（DELETE）
    Logger.log(`\n--- ③ 取引を削除（DELETE） ID: ${dealId} ---`);
    callFreeeApi('delete', `/api/1/deals/${dealId}`, { company_id: companyId }, null);
    Logger.log('取引を削除しました。');
    ui.alert(`取引ID: ${dealId} を削除しました。`);

    // ④ 再度、取引の取得（GET）
    Logger.log('\n--- ④ 削除後の取引一覧を取得（GET）---');
    const dealsAfterDelete = callFreeeApi('get', '/api/1/deals', { company_id: companyId }, null);
    Logger.log('現在の取引一覧:', dealsAfterDelete.deals);
    ui.alert('削除後の取引一覧をログに出力しました。\n実行ログをご確認ください。');

  } catch (e) {
    Logger.log('エラーが発生しました: ' + e.message);
    ui.alert('エラーが発生しました: ' + e.message);
  }
}
