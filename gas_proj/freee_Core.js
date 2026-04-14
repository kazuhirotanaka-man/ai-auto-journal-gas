// freee APIのクライアントIDとクライアントシークレットを設定
const CLIENT_ID = '705718818614667';
const CLIENT_SECRET = 'YRCW1uGn_ReGOR3MmgzpoA_MEXbfbL63Yx4cU2YnNtNwN4C4QjJQ5U7KskoxLPHUDmcpvTA7X9YBl4q3-DhjgQ';

const FREEE_API_BASE_URL = 'https://api.freee.co.jp';

/**
 * UIを安全に取得する。UIが利用できないコンテキスト（デバッガなど）では、
 * Loggerにフォールバックするダミーオブジェクトを返す。
 * @return {Ui} SpreadsheetApp.getUi() またはダミーのUIオブジェクト
 */
function getSafeUi() {
  try {
    const ui = SpreadsheetApp.getUi();
    // UIオブジェクトが正常に機能するかを、alertの存在チェックで確認
    if (!ui.alert) {
      throw new Error('UI context not available.');
    }
    return ui;
  } catch (e) {
    // UIが利用できない場合、ダミーのUIオブジェクトを返す
    const DUMMY_BUTTON = { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' };
    return {
      Button: DUMMY_BUTTON,
      ButtonSet: { OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
      alert: function (title, message, buttons) {
        if (message === undefined) { Logger.log('ALERT: ' + title); }
        else { Logger.log('ALERT: ' + title + '\n' + message); }

        // 確認ダイアログのデバッグ実行時は、肯定的な応答を返す
        if (buttons === this.ButtonSet.YES_NO) {
          Logger.log('【デバッグ】確認ダイアログ(YES_NO)で「YES」が押されたと仮定');
          return this.Button.YES;
        }
        if (buttons === this.ButtonSet.OK_CANCEL) {
          Logger.log('【デバッグ】確認ダイアログ(OK_CANCEL)で「OK」が押されたと仮定');
          return this.Button.OK;
        }

        // 通常のalertの場合
        Logger.log('【デバッグ】OKボタンが押されたと仮定');
        return this.Button.OK;
      },
      prompt: function (title, promptText) {
        Logger.log(`PROMPT: ${title}\n${promptText}`);
        return { getSelectedButton: () => DUMMY_BUTTON.CANCEL, getResponseText: () => null };
      },
      showSidebar: (html) => Logger.log('Sidebar would be shown.'),
      createMenu: (caption) => ({
        addItem: function () { return this; },
        addSeparator: function () { return this; },
        addToUi: function () { }
      })
    };
  }
}

/**
 * OAuth2サービスを取得する
 * @return {OAuth2.Service} OAuth2サービス
 */
function getFreeeService() {
  return OAuth2.createService('freee')
    .setAuthorizationBaseUrl('https://accounts.secure.freee.co.jp/public_api/authorize')
    .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('read write');
}

/**
 * 認証用のサイドバーを表示する
 */
function showFreeeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('freee_Sidebar')
    .setTitle('freee 認証');
  getSafeUi().showSidebar(html);
}

/**
 * 事業所選択用のサイドバーを表示する
 */
function showCompanySelectorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('freee_CompanySelector')
    .setTitle('事業所選択');
  getSafeUi().showSidebar(html);
}

/**
 * [サイドバー用] 認可コードとアクセストークンを交換する
 * @param {string} code 認可コード
 * @return {boolean} 成功したかどうか
 */
function exchangeCode(code) {
  try {
    return getFreeeService().exchangeCodeForToken(code);
  } catch (e) {
    Logger.log('Token exchange failed: ' + e.message);
    return false;
  }
}

/**
 * [サイドバー用] 初期表示情報を返す
 * @return {object} 認証状態と認証URL
 */
function getSidebarState() {
  const service = getFreeeService();
  const isAuthorized = service.hasAccess();
  return {
    isAuthorized: isAuthorized,
    authUrl: isAuthorized ? null : service.getAuthorizationUrl()
  };
}

/**
 * 認証をリセット（連携を解除）する
 */
function reset() {
  const ui = getSafeUi();
  const result = ui.alert(
    '確認',
    'freeeとの連携を解除しますか？',
    ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    getFreeeService().reset();
    ui.alert('連携を解除しました。\n再度利用するには、メニューからもう一度「認証」を行ってください。');
  }
}

/**
 * freee APIを呼び出すための汎用関数
 * @param {string} method HTTPメソッド (get, post, put, delete)
 * @param {string} urlPath エンドポイントのパス (例: '/api/1/companies')
 * @param {object} [urlQuery] URLクエリパラメータのオブジェクト (例: { company_id: 123 })
 * @param {object} [body] リクエストボディのオブジェクト
 * @param {boolean} [multipart=false] マルチパートフォームデータとして送信するかどうか
 * @return {object|null} APIからのレスポンスをパースしたJSONオブジェクト
 */
function callFreeeApi(method, urlPath, urlQuery, body, multipart = false) {
  const freeeService = getFreeeService();
  if (!freeeService.hasAccess()) {
    throw new Error('認証が必要です。メニューから「認証」を実行してください。');
  }

  let fullUrl = FREEE_API_BASE_URL + (urlPath.startsWith('/') ? urlPath : '/' + urlPath);

  if (urlQuery) {
    const queryString = Object.keys(urlQuery)
      .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(urlQuery[key])}`)
      .join('&');
    if (queryString) {
      fullUrl += '?' + queryString;
    }
  }

  const options = {
    method: method.toLowerCase(),
    headers: {
      'Authorization': 'Bearer ' + freeeService.getAccessToken(),
      'X-Api-Version': '2020-06-15'
    },
    muteHttpExceptions: true
  };

  if (body) {
    if (multipart) {
      options.payload = body;
    } else {
      options.payload = JSON.stringify(body);
      options.headers['Content-Type'] = 'application/json';
    }
  }

  const response = UrlFetchApp.fetch(fullUrl, options);
  const responseBody = response.getContentText();
  const responseCode = response.getResponseCode();

  if (responseCode >= 400) {
    throw new Error(`APIエラー (${responseCode}): ${responseBody}`);
  }

  return responseBody ? JSON.parse(responseBody) : null;
}

// --- 事業所選択関連の関数 ---

/**
 * [サイドバー用] 事業所選択サイドバーの初期情報を取得する
 */
function getCompanySelectorInfo() {
  try {
    const result = callFreeeApi('get', '/api/1/companies', null, null);
    const selectedCompanyId = PropertiesService.getUserProperties().getProperty('selectedCompanyId');
    let selectedCompanyName = null;

    if (selectedCompanyId) {
      const selectedCompany = result.companies.find(c => c.id == selectedCompanyId);
      if (selectedCompany) {
        selectedCompanyName = selectedCompany.display_name;
      }
    }

    return {
      companies: result.companies,
      selectedCompanyId: selectedCompanyId,
      selectedCompanyName: selectedCompanyName
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * [サイドバー用] 選択された事業所IDを保存する
 * @param {string} companyId 保存する事業所ID
 */
function setSelectedCompany(companyId) {
  PropertiesService.getUserProperties().setProperty('selectedCompanyId', companyId);
}

/**
 * トークンをリフレッシュする（トリガー実行用）
 */
function refreshFreeeToken() {
  try {
    const service = getFreeeService();
    // トークンが存在するか確認
    // hasAccess()は期限切れの場合リフレッシュを試みるが、
    // ここでは期限に関わらずリフレッシュを実行するため、
    // まずトークンの存在確認として利用（内部でリフレッシュが走っても問題ない）
    if (service.hasAccess()) {
      service.refresh();
      Logger.log('Token refreshed successfully.');
    } else {
      Logger.log('Access token not available. Please re-authorize.');
    }
  } catch (e) {
    Logger.log('Failed to refresh token: ' + e.message);
  }
}

/**
 * トークン自動リフレッシュの有効化/無効化を切り替える
 */
function toggleAutoRefreshToken() {
  const ui = getSafeUi();
  const triggerName = 'refreshFreeeToken';
  const triggers = ScriptApp.getProjectTriggers();
  let existingTrigger = null;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === triggerName) {
      existingTrigger = trigger;
      break;
    }
  }

  if (existingTrigger) {
    const result = ui.alert(
      '自動更新設定',
      '現在、トークンの自動更新は「有効」です。\n無効にしますか？',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      ScriptApp.deleteTrigger(existingTrigger);
      ui.alert('自動更新を無効にしました。');
    }
  } else {
    const result = ui.alert(
      '自動更新設定',
      '現在、トークンの自動更新は「無効」です。\n毎月1日に自動更新するように設定しますか？',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      ScriptApp.newTrigger(triggerName)
        .timeBased()
        .onMonthDay(1)
        .atHour(9)
        .create();
      ui.alert('自動更新を有効にしました。（毎月1日 9時頃実行）');
    }
  }
}
