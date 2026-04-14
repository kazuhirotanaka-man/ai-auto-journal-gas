/**
 * freee形式のプロンプトを生成する
 * @param {object} config 
 * @param {string} fileNameInfo 
 * @param {string} companyNameInfo 
 * @param {string} businessContext 
 * @returns {string} プロンプト文字列
 */
function generateFreeePrompt(config, fileNameInfo, companyNameInfo, businessContext) {
  const accountsInfo = config.freeeAccountsList && config.freeeAccountsList.length > 0 ? `勘定科目の候補: [${config.freeeAccountsList.join(", ")}]` : "適切な勘定科目を推測してください。";
  const taxCategoryInfo = config.freeeTaxCategoryList && config.freeeTaxCategoryList.length > 0 ? `税区分の候補: [${config.freeeTaxCategoryList.join(", ")}]` : "税区分は適切なものを推測するか空欄にしてください。";
  const walletsInfo = config.freeeWalletsList && config.freeeWalletsList.length > 0 ? `決済口座の候補: [${config.freeeWalletsList.join(", ")}]` : "決済口座は適切なものを推測するか空欄にしてください。";
  const partnersInfo = config.freeePartnersList && config.freeePartnersList.length > 0 ? `取引先の候補: [${config.freeePartnersList.join(", ")}]` : "取引先は証憑から読み取ってください。";
  const itemsInfo = config.freeeItemsList && config.freeeItemsList.length > 0 ? `品目の候補: [${config.freeeItemsList.join(", ")}]` : "品目は空欄または証憑から読み取った一般的な名称にしてください。";
  const deptsInfo = config.freeeDepartmentsList && config.freeeDepartmentsList.length > 0 ? `部門の候補: [${config.freeeDepartmentsList.join(", ")}]` : "部門は空欄にしてください。";
  const tagsInfo = config.freeeTagsList && config.freeeTagsList.length > 0 ? `メモタグの候補: [${config.freeeTagsList.join(", ")}]` : "メモタグは空欄にしてください。";
  
  return `
あなたはプロの会計事務所職員です。添付された証憑画像（領収書、請求書など）を解析し、freee会計向けの仕訳データを作成してください。
以下のJSONフォーマットで回答してください。JSON以外のテキストは出力しないでください。

前提条件：
${fileNameInfo}
${companyNameInfo}
${businessContext}
${config.extraPrompt ? `追加の指針: ${config.extraPrompt}` : ""}

条件：
1. ${accountsInfo} ここから最も適切なものを選んでください。
2. ${taxCategoryInfo}
3. ${walletsInfo}
4. ${partnersInfo}
5. ${itemsInfo}
6. ${deptsInfo}
7. ${tagsInfo}
8. 収支は「収入」「支出」のいずれか。
9. 決済ステータスは「決済済」「未決済」のいずれか。未決済の場合は「決済口座」を空文字列("")にしてください。決済済みの場合のみは決済口座を設定してください。
10. 発生日と決済期日は YYYY/MM/DD 形式。読み取れない場合は空欄。決済期日は必ず発生日以降の日付にしてください。
11. 金額は数値（カンマなし）。
12. 推測証票種別は「受取請求書」「領収書・レシート」「発行請求書」「クレカ利用明細」「その他」から選択してください。
13. 推測決済方法は証票が「受取請求書」「領収書・レシート」の場合、「現金」「クレジットカード」「振込・引落し」から選択してください。
14. 適格請求書発行事業者登録番号（"T"とそれに続く13桁の数字。例: T1234567890123）が記載されている場合は、"registrationNumber"に抽出してください。記載がない場合は空文字列("")にしてください。
15. confidenceは、読み取り結果に対する自信度（"高", "中", "低"）を入れてください。
16. 1つの証憑から複数の異なる取引明細（異なる税率、異なる勘定科目など）が読み取れる場合は、配列 "entries" の中に複数の明細データを含めてください。
17. そもそも会計の取引記録として不要な画像（単なるメモ、他の証憑と重複、業務に無関係なもの等）であると判断した場合は、"isTarget": false とし、その理由を最初の "description" に入れてください。

出力フォーマット（JSON）:
{
  "isTarget": true,
  "description": "",
  "entries": [
    {
      "incomeExpense": "支出",
      "accrualDate": "YYYY/MM/DD",
      "partner": "Amazon",
      "paymentStatus": "決済済",
      "paymentDate": "YYYY/MM/DD",
      "wallet": "現金",
      "accountItem": "消耗品費",
      "amount": 1000,
      "taxCategory": "課税仕入",
      "item": "事務用品",
      "department": "営業部",
      "memoTag": "",
      "registrationNumber": "T1234567890123",
      "remarks": "ボールペン",
      "guessedDocumentType": "領収書・レシート",
      "guessedPaymentMethod": "現金",
      "confidence": "高"
    }
  ]
}
`;
}

/**
 * 弥生形式・汎用形式のプロンプトを生成する
 * @param {object} config 
 * @param {string} fileNameInfo 
 * @param {string} companyNameInfo 
 * @param {string} businessContext 
 * @returns {string} プロンプト文字列
 */
function generateYayoiPrompt(config, fileNameInfo, companyNameInfo, businessContext) {
  const accountsInfo = config.accountsList && config.accountsList.length > 0 ? `勘定科目の候補: [${config.accountsList.join(", ")}]` : "適切な勘定科目を推測してください。";
  const subAccountsInfo = config.subAccountsList && config.subAccountsList.length > 0 ? `補助科目の候補: [${config.subAccountsList.join(", ")}]` : "補助科目は（明確に指定がない限り）空欄にしてください。";
  const taxCategoryInfo = config.taxCategoryList && config.taxCategoryList.length > 0 ? `税区分の候補: [${config.taxCategoryList.join(", ")}]` : "税区分は適切なものを推測するか空欄にしてください。";
  const extraPromptInfo = config.extraPrompt ? `追加の指針: ${config.extraPrompt}` : "";

  return `
あなたはプロの会計事務所職員です。添付された証憑画像（領収書、請求書など）を解析し、仕訳データを作成してください。
以下のJSONフォーマットで回答してください。JSON以外のテキストは出力しないでください。

前提条件：
${fileNameInfo}
${companyNameInfo}
${businessContext}

条件：
1. ${accountsInfo} ここから最も適切なものを選んでください。
2. ${subAccountsInfo}
3. ${taxCategoryInfo} ここから借方・貸方のそれぞれの税区分を選んでください。
4. ${extraPromptInfo}
5. 日付は YYYY/MM/DD 形式。読み取れない場合は空欄。
6. 金額は数値（カンマなし）。読み取れない場合は 0。
7. 借方／貸方の判定は、通常の支払い（経費）であれば 借方: 経費科目 / 貸方: 現金など とします。
   （※ここでは簡略化のため、「支払った経費」として片側の科目（借方）と金額を特定することに注力してください。もう片方は固定フォーマットにならって空欄・あるいは「現金／未払金」等のデフォルトで構いません。ツール側の後処理で適宜補完します。）
   なお、1つの証憑から複数の異なる取引内容（異なる税率、異なる勘定科目など）が読み取れる場合は、配列 "entries" の中に複数の仕訳データを含めてください。
8. confidenceは、読み取り結果に対する自信度（"高", "中", "低"）を入れてください。手書きで読みづらかったり、科目判定に迷った場合は"低"や"中"にしてください。
9. そもそも会計の取引記録として不要な画像（単なるメモ、他の証憑と重複、業務に無関係なもの等）であると判断した場合は、"isTarget": false とし、その理由を最初の "description" に入れてください。仕訳が必要な証憑の場合は "isTarget": true にしてください。

出力フォーマット（JSON）:
{
  "isTarget": true,
  "entries": [
    {
      "date": "YYYY/MM/DD",
      "amount": 1000,
      "debitAccount": "消耗品費",
      "debitSubAccount": "",
      "debitTaxCategory": "対象外",
      "creditAccount": "現金",
      "creditSubAccount": "",
      "creditTaxCategory": "対象外",
      "description": "Amazon / オフィス用品",
      "confidence": "高"
    }
  ]
}
`;
}

/**
 * Gemini APIと通信し、画像から仕訳情報を抽出するサービス
 */
const GeminiService = {

  /**
   * 画像ファイルと設定リストを元にGeminiに解析をリクエストする
   * @param {GoogleAppsScript.Drive.File} file 対象のファイルオブジェクト
   * @param {object} config 設定オブジェクト
   * @returns {object} 解析結果 (JSON)
   */
  analyzeReceipt: function (file, config) {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      throw new Error("Gemini APIキーが設定されていません。Script Properties (GEMINI_API_KEY) を確認してください。");
    }

    const mimeType = file.getMimeType();

    // 対応していないファイル形式の場合は、API通信を行わずに「対象外」として返す
    const validMimes = ["image/jpeg", "image/png", "application/pdf", "image/webp", "image/heic", "image/heif"];
    if (!validMimes.includes(mimeType)) {
      return {
        isTarget: false,
        entries: [{
          description: "対象外（画像・PDF以外の形式は処理できません）",
          debitTaxCategory: "対象外",
          creditTaxCategory: "対象外"
        }]
      };
    }

    const bytes = file.getBlob().getBytes();
    const base64Data = Utilities.base64Encode(bytes);

    // ユーザー設定からプロンプトを構築
    const fileNameInfo = `ファイル名: ${file.getName()}`;
    const companyNameInfo = config.companyName ? `自社名（この会計データの主体）: ${config.companyName}\n※この自社名が発行元となっている請求書は「発行した請求書（売上など）」、宛先となっている場合は「受け取った請求書（経費など）」として区別してください。` : "";
    const industryInfo = config.industryType ? `自社の業種（大分類）: ${config.industryType}` : "";
    const businessInfo = config.businessDetails ? `具体的な事業内容: ${config.businessDetails}` : "";
    const businessContext = (industryInfo || businessInfo) ? `${industryInfo ? industryInfo + "\\n" : ""}${businessInfo ? businessInfo + "\\n" : ""}※この業種・事業内容特有の経費科目や取引の性質を考慮して仕訳を推測してください。` : "";

    let systemPrompt = config.accountingSoftware === "freee会計"
      ? generateFreeePrompt(config, fileNameInfo, companyNameInfo, businessContext)
      : generateYayoiPrompt(config, fileNameInfo, companyNameInfo, businessContext);

    // Gemini 2.0 Flash Lite (ご指定の Flash Lite を使用)
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=${apiKey}`;

    const payload = {
      contents: [
        {
          parts: [
            { text: systemPrompt },
            {
              inlineData: {
                mimeType: mimeType,
                data: base64Data
              }
            }
          ]
        }
      ]
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error(`Gemini API エラー: ${jsonResponse.error ? jsonResponse.error.message : response.getContentText()}`);
    }

    try {
      const textResult = jsonResponse.candidates[0].content.parts[0].text;
      // markdownのコードブロック ```json ... ``` を除去してパース
      const cleanJson = textResult.replace(/```json/g, '').replace(/```/g, '').trim();
      return JSON.parse(cleanJson);
    } catch (e) {
      throw new Error("AIのレスポンスをJSONとして解釈できませんでした。 レスポンス: " + JSON.stringify(jsonResponse));
    }
  }

};
