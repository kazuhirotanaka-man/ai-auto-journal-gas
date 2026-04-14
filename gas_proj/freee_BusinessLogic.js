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

/**
 * マスタシートから名前で検索して値を取得する
 */
function getFreeeMasterValueByName(sheetName, name, returnColIndex = 0, searchColIndex = 1) {
  if (!name) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const maxCol = Math.max(returnColIndex, searchColIndex) + 1;
  const values = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][searchColIndex] == name) {
      return values[i][returnColIndex];
    }
  }
  return null;
}

/**
 * freee_ExportDialogからの呼び出し：取引の登録を実行
 * @param {string[]} statuses 処理対象のステータス
 */
function executeFreeeExportProcess(statuses) {
  const ui = getSafeUi();
  const companyId = getSelectedCompanyId();
  if (!companyId) return { error: "事業所が選択されていません。" };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('freee取引データ');
  if (!sheet) return { error: "「freee取引データ」シートが見つかりません。" };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { error: "出力対象のデータがありません。" };
  }

  // A列～T列まで取得する (T列のインデックスは19)
  const MAX_COL = 20;
  const RESULT_COL_INDEX = 19;

  const range = sheet.getRange(2, 1, lastRow - 1, MAX_COL);
  const values = range.getValues();

  // 取引ごとにまとめる処理
  let transactions = [];
  let currentTx = null;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const isHead = (row[FREEE_COL.INCOME_EXPENSE] !== "");

    if (isHead) {
      if (currentTx) {
        transactions.push(currentTx);
      }
      currentTx = {
        rowIndex: i, // 0-indexed values array matching
        head: row,
        details: [row], // First detail is the head row itself for detail cols
        detailRowIndices: [i]
      };
    } else {
      if (currentTx) {
        currentTx.details.push(row);
        currentTx.detailRowIndices.push(i);
      }
    }
  }
  if (currentTx) transactions.push(currentTx);

  let successCount = 0;
  let errorCount = 0;
  let attachmentErrorCount = 0;

  SpreadsheetApp.getActiveSpreadsheet().toast('対象の取引を登録しています...', '処理中', -1);

  for (const tx of transactions) {
    const head = tx.head;
    const status = head[FREEE_COL.STATUS];

    // 指定されたステータス以外は無視
    if (!statuses.includes(status)) continue;

    let isAttachmentFailed = false;
    let receiptId = null;

    // ファイルのアップロード (ファイルボックスへの登録)
    const fileId = head[FREEE_COL.FILE_ID] || null;
    if (fileId) {
      try {
        const driveFile = DriveApp.getFileById(fileId);
        const blob = driveFile.getBlob();
        
        // receipts APIの呼び出し (multipart/form-data送信)
        const receiptPayload = {
          company_id: companyId,
          receipt: blob,
          description: driveFile.getName()
        };
        
        const receiptRes = callFreeeApi('post', '/api/1/receipts', null, receiptPayload, true);
        if (receiptRes && receiptRes.receipt && receiptRes.receipt.id) {
          receiptId = receiptRes.receipt.id;
        }
      } catch (fileErr) {
        Logger.log(`ファイルアップロードエラー(FileID: ${fileId}): ${fileErr.message}`);
        attachmentErrorCount++;
        isAttachmentFailed = true;
      }
    }

    // ペイロード構築
    try {
      const typeStr = head[FREEE_COL.INCOME_EXPENSE];
      let dealType = "";
      if (typeStr === "収入") dealType = "income";
      else if (typeStr === "支出") dealType = "expense";
      else throw new Error("収支には「収入」または「支出」を指定してください。");

      const issueDate = formatRawDate(head[FREEE_COL.ACCRUAL_DATE]) || formatRawDate(new Date());

      const partnerName = head[FREEE_COL.PARTNER];
      let partnerId = undefined;
      if (partnerName) {
        partnerId = getFreeeMasterValueByName('マスタfreee取引先', partnerName, 0, 1);
        if (!partnerId) throw new Error(`取引先「${partnerName}」が見つかりませんでした。`);
      }
      const refNumber = head[FREEE_COL.REG_NUM] || undefined;

      let detailsPayload = [];
      for (const row of tx.details) {
        const taxName = row[FREEE_COL.TAX];
        const taxCode = taxName ? getFreeeMasterValueByName('マスタfreee税区分', taxName, 0, 1) : undefined;
        if (taxCode === null || taxCode === undefined) throw new Error(`税区分「${taxName}」が見つかりませんでした。`);

        const accountName = row[FREEE_COL.ACC_ITEM];
        let accountId = undefined;
        if (accountName) {
          accountId = getFreeeMasterValueByName('マスタfreee勘定科目', accountName, 0, 1);
          if (!accountId) throw new Error(`勘定科目「${accountName}」が見つかりませんでした。`);
        }

        const itemName = row[FREEE_COL.ITEM];
        let itemId = undefined;
        if (itemName) {
          itemId = getFreeeMasterValueByName('マスタfreee品目', itemName, 0, 1);
          if (!itemId) throw new Error(`品目「${itemName}」が見つかりませんでした。`);
        }

        const deptName = row[FREEE_COL.DEPT];
        let deptId = undefined;
        if (deptName) {
          deptId = getFreeeMasterValueByName('マスタfreee部門', deptName, 0, 1);
          if (!deptId) throw new Error(`部門「${deptName}」が見つかりませんでした。`);
        }

        const tagName = row[FREEE_COL.TAG];
        let tagId = undefined;
        if (tagName) {
          tagId = getFreeeMasterValueByName('マスタfreeeメモタグ', tagName, 0, 1);
          if (!tagId) throw new Error(`メモタグ「${tagName}」が見つかりませんでした。`);
        }

        let amount = row[FREEE_COL.AMOUNT];
        if (amount === "" || amount == null) amount = 0;

        const detail = {
          tax_code: taxCode,
          account_item_id: accountId,
          amount: parseInt(amount, 10)
        };
        if (itemId) detail.item_id = itemId;
        if (deptId) detail.section_id = deptId;
        if (tagId) detail.tag_ids = [tagId];


        const description = row[FREEE_COL.DESC];
        if (description) detail.description = description;

        detailsPayload.push(detail);
      }

      const dealBody = {
        company_id: companyId,
        issue_date: issueDate,
        type: dealType,
        details: detailsPayload
      };

      if (partnerId) {
        dealBody.partner_id = partnerId;
      }

      if (receiptId) {
        dealBody.receipt_ids = [receiptId];
      }

      if (refNumber) {
        dealBody.ref_number = String(refNumber);
      }

      const payStatus = head[FREEE_COL.PAY_STATUS];
      if (payStatus === "決済済") {
        const walletName = head[FREEE_COL.WALLET];
        let walletId = undefined;
        let walletType = undefined;
        if (walletName) {
          walletId = getFreeeMasterValueByName('マスタfreee口座', walletName, 0, 1);
          walletType = getFreeeMasterValueByName('マスタfreee口座', walletName, 2, 1);
          if (!walletId || !walletType) throw new Error(`口座「${walletName}」が見つかりませんでした。`);
        }

        const payDate = formatRawDate(head[FREEE_COL.PAY_DATE]) || issueDate;

        let totalAmount = 0;
        for (const pd of detailsPayload) totalAmount += pd.amount;

        if (walletId && walletType) {
          dealBody.payments = [{
            amount: totalAmount,
            date: payDate,
            from_walletable_id: walletId,
            from_walletable_type: walletType
          }];
        }
      }

      // APIによる登録
      const createdDeal = callFreeeApi('post', '/api/1/deals', null, dealBody);
      const dealId = createdDeal.deal.id;

      // 書き戻し
      for (const rIndex of tx.detailRowIndices) {
        let resultMsg = dealId;
        if (isAttachmentFailed) {
          resultMsg += " (添付失敗)";
        }
        values[rIndex][RESULT_COL_INDEX] = resultMsg; // freee登録結果
        values[rIndex][FREEE_COL.STATUS] = '登録済';
      }

      successCount++;
    } catch (err) {
      // 取引登録でエラーが発生した場合、アップロード済みの証憑をロールバック(削除)する
      if (receiptId) {
        try {
          callFreeeApi('delete', `/api/1/receipts/${receiptId}`, { company_id: companyId }, null);
          Logger.log(`ロールバック完了: ファイルボックスから証憑(${receiptId})を削除しました。`);
        } catch (rollbackErr) {
          Logger.log(`ロールバック失敗: 証憑(${receiptId})の削除に失敗しました: ${rollbackErr.message}`);
        }
      }

      for (const rIndex of tx.detailRowIndices) {
        let resultMsg = "エラー: " + err.message;
        if (isAttachmentFailed) {
          resultMsg += " (添付失敗含む)";
        }
        values[rIndex][RESULT_COL_INDEX] = resultMsg;
      }
      errorCount++;
    }
  }

  // シートの更新 (全体を上書きすると入力規則違反の列があった場合にエラーになるため、対象の2列のみ更新する)
  const updateValues = values.map(row => [row[FREEE_COL.STATUS], row[RESULT_COL_INDEX]]);
  sheet.getRange(2, FREEE_COL.STATUS + 1, updateValues.length, 2).setValues(updateValues);

  SpreadsheetApp.getActiveSpreadsheet().toast(`登録完了 成功: ${successCount} 件, 失敗: ${errorCount} 件`, '完了', 5);
  
  return {
    success: true,
    successCount: successCount,
    errorCount: errorCount,
    attachmentErrorCount: attachmentErrorCount
  };
}
