/**
 * Gemini APIと通信し、画像から仕訳情報を抽出するサービス
 */
const GeminiService = {

  /**
   * 画像ファイルと設定リストを元にGeminiに解析をリクエストする
   * @param {GoogleAppsScript.Drive.File} file 
   * @param {object} config 
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

    const accountsInfo = config.accountsList.length > 0 ? `勘定科目の候補: [${config.accountsList.join(", ")}]` : "適切な勘定科目を推測してください。";
    const subAccountsInfo = config.subAccountsList.length > 0 ? `補助科目の候補: [${config.subAccountsList.join(", ")}]` : "補助科目は（明確に指定がない限り）空欄にしてください。";
    const taxCategoryInfo = config.taxCategoryList && config.taxCategoryList.length > 0 ? `税区分の候補: [${config.taxCategoryList.join(", ")}]` : "税区分は適切なものを推測するか空欄にしてください。";
    const extraPromptInfo = config.extraPrompt ? `追加の指針: ${config.extraPrompt}` : "";

    const systemPrompt = `
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
