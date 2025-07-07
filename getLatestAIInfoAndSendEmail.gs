/**
 * 生成AIの最新情報をWeb検索とGeminiを活用して収集し、事実確認を行った上でメールで送信するスクリプト
 */
function getLatestAIInfoAndSendEmail() {
  // --- 設定項目 ---
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const CUSTOM_SEARCH_API_KEY = PropertiesService.getScriptProperties().getProperty('CUSTOM_SEARCH_API_KEY');
  const CUSTOM_SEARCH_ENGINE_ID = PropertiesService.getScriptProperties().getProperty('CUSTOM_SEARCH_ENGINE_ID');
  const RECIPIENT_EMAIL = Session.getActiveUser().getEmail();
  const DAYS_PRIOR = 7; // 何日前からの情報を取得するか (例: 7で過去1週間)
  const MAX_SEARCH_RESULTS_PER_QUERY = 5; // 各検索クエリからWebページ全文を取得する記事の最大数 (API制限と処理時間を考慮し少なめに設定)
  const MAX_CONTENT_LENGTH_FOR_GEMINI = 10000; // Geminiに渡すWebページコンテンツの最大文字数 (トークン制限を考慮)

  // --- フィルタリング設定 ---
  // 信頼できるドメインのホワイトリスト（小文字で定義）
  const TRUSTED_DOMAINS = [
    'openai.com',
    'blog.google',
    'deepmind.google',
    'microsoft.com',
    'techcrunch.com',
    'wired.com',
    'theverge.com',
    'nature.com',
    'science.org',
    'arxiv.org', // 研究論文サイト
    'github.com', // 主要プロジェクト
    'news.google.com', // Googleニュース（ただし、リンク先ドメインも確認必要）
    'forbesjapan.com',
    'nikkei.com',
    'bloomberg.com',
    'reuters.com',
    'itmedia.co.jp',
    'impress.co.jp',
    'ascii.jp',
    'zenn.dev', // 技術系ブログプラットフォーム
    'note.com', // 技術系ブログ（選別必要だが、有用なものもあるため一旦含める）
    // 必要に応じて追加・削除してください
  ];

  // 除外キーワードのブラックリスト（小文字で定義）
  const EXCLUDE_KEYWORDS = [
    '研修', 'セミナー', 'ウェビナー', '講座', 'イベント', '求人', '採用', '募集', '広告', 'キャンペーン', '割引',
    '無料体験', 'コンサルティング', 'ソリューション', '導入事例', '資料ダウンロード', 'フォーム', 'お問い合わせ',
    '申込み', '登録', '参加', '開催', '発表会', '展示会', 'レポート', 'ホワイトペーパー', '限定', 'セール',
    'ログイン', 'サインアップ', 'プライバシーポリシー', '利用規約', 'サイトマップ', '会社概要', '連絡先'
  ];

  // 除外するファイル拡張子
  const EXCLUDE_EXTENSIONS = ['.pdf', '.zip', '.docx', '.xlsx', '.pptx', '.jpg', '.png', '.gif', '.mp4', '.mp3'];


  // APIキーとIDが設定されているかチェック
  if (!GEMINI_API_KEY || !CUSTOM_SEARCH_API_KEY || !CUSTOM_SEARCH_ENGINE_ID) {
    const errorMessage = 'APIキーまたは検索エンジンIDがスクリプトプロパティに設定されていません。';
    Logger.log(errorMessage);
    MailApp.sendEmail(RECIPIENT_EMAIL, 'GASエラー: 設定不備', errorMessage + ' スクリプトプロパティに「GEMINI_API_KEY」「CUSTOM_SEARCH_API_KEY」「CUSTOM_SEARCH_ENGINE_ID」を設定してください。');
    return;
  }

  // --- プロンプトの組み立て ---
  const today = new Date();
  const targetDate = new Date(today.getTime() - DAYS_PRIOR * 24 * 60 * 60 * 1000);
  const targetDateString = targetDate.toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });

  const searchQueries = [
    `"生成AI 最新情報" ${targetDateString}以降`,
    `"LLM 技術革新" ${targetDateString}以降`,
    `"画像生成AI 進化" ${targetDateString}以降`,
    `"音声AI 動向" ${targetDateString}以降`,
    `"マルチモーダルAI 研究" ${targetDateString}以降`,
    `"GPT-4o 新機能" ${targetDateString}以降`,
    `"Gemini 1.5 Flash アップデート" ${targetDateString}以降`,
    `"Claude 3.5 Sonnet 特徴" ${targetDateString}以降`,
    `"Sora 最新情報" ${targetDateString}以降`,
    `"Stable Diffusion 3.0 発表" ${targetDateString}以降`,
    `"Llama 3.1 開発" ${targetDateString}以降`,
    `"生成AI トレンド" ${targetDateString}以降`,
    `"生成AI ビジネス応用事例" ${targetDateString}以降`,
    `"生成AI 規制 動向" ${targetDateString}以降`,
    `"AI 著作権 問題" ${targetDateString}以降`,
    `"生成AI ブレイクスルー" ${targetDateString}以降`,
    `"Google AI 最新情報" ${targetDateString}以降`,
    `"Google Gemini アップデート" ${targetDateString}以降`,
    `"Google DeepMind 新技術" ${targetDateString}以降`,
    `"Google Bard 後継" ${targetDateString}以降`,
    `"Google Vertex AI 新機能" ${targetDateString}以降`,
    `"Google AI in Search" ${targetDateString}以降`,
    `"Google AI Ethic" ${targetDateString}以降`
  ];

  // --- Gemini APIへのリクエスト関数 ---
  function callGeminiAPI(promptContent, apiKey) {
    const API_ENDPOINT = `https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    const requestBody = {
      contents: [{
        parts: [{
          text: promptContent
        }]
      }]
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };

    let response;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
    } catch (e) {
      throw new Error('Gemini API呼び出しエラー: ' + e.message);
    }

    const responseText = response.getContentText();
    try {
      const jsonResponse = JSON.parse(responseText);
      if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts && jsonResponse.candidates[0].content.parts[0]) {
        return jsonResponse.candidates[0].content.parts[0].text;
      } else {
        const errorDetails = jsonResponse.error ? JSON.stringify(jsonResponse.error) : '詳細不明';
        throw new Error('Geminiからのレスポンス形式が不正です。詳細: ' + errorDetails + ' レスポンス: ' + responseText);
      }
    } catch (e) {
      throw new Error('GeminiレスポンスのJSON解析エラー: ' + e.message + ' レスポンス: ' + responseText);
    }
  }

  // --- Web検索を実行する関数 ---
  function performCustomSearch(query, apiKey, cxId, numResults) {
    const searchUrl = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cxId}&q=${encodeURIComponent(query)}&num=${numResults}`;
    Logger.log(`検索URL: ${searchUrl}`);
    let response;
    try {
      response = UrlFetchApp.fetch(searchUrl);
    } catch (e) {
      throw new Error('Custom Search API呼び出しエラー: ' + e.message);
    }

    const responseText = response.getContentText();
    try {
      const jsonResponse = JSON.parse(responseText);
      if (jsonResponse.items) {
        return jsonResponse.items.map(item => ({
          title: item.title,
          link: item.link,
          snippet: item.snippet
        }));
      } else {
        Logger.log('Custom Search APIからのレスポンスにitemsがありません: ' + responseText);
        return [];
      }
    } catch (e) {
      throw new Error('Custom Search APIレスポンスのJSON解析エラー: ' + e.message + ' レスポンス: ' + responseText);
    }
  }

  /**
   * URLからホスト名（ドメイン）を抽出するヘルパー関数
   * @param {string} url
   * @returns {string|null} ホスト名、またはnull
   */
  function getHostnameFromUrl(url) {
    try {
      const match = url.match(/^https?:\/\/([^/]+)/i);
      if (match && match[1]) {
        let hostname = match[1];
        // 'www.' を除去
        if (hostname.startsWith('www.')) {
          hostname = hostname.substring(4);
        }
        return hostname.toLowerCase();
      }
    } catch (e) {
      Logger.log(`URLからホスト名抽出エラー: ${url}, ${e.message}`);
    }
    return null;
  }

  /**
   * 指定されたURLのHTMLから主要なテキストコンテンツを抽出する簡易スクレイピング関数
   *
   * @param {string} url 抽出対象のURL
   * @returns {string} 抽出されたテキストコンテンツ
   */
  function extractMainContent(url) {
    try {
      const pathnameMatch = url.match(/^(?:https?:\/\/[^/]+)?([^?#]+)/i); // パス名部分を取得
      const pathname = pathnameMatch ? pathnameMatch[1].toLowerCase() : '';

      // ファイル拡張子による除外
      if (EXCLUDE_EXTENSIONS.some(ext => pathname.endsWith(ext))) {
          Logger.log(`コンテンツ抽出スキップ: 除外対象のファイル形式 ${url}`);
          return '';
      }

      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const contentType = response.getHeaders()['Content-Type'];
      if (contentType && !contentType.includes('text/html')) {
          Logger.log(`コンテンツ抽出スキップ: HTML以外のコンテンツタイプ ${contentType} for ${url}`);
          return '';
      }
      
      const html = response.getContentText();

      // HTMLタグを除去し、余分な改行やスペースを整理する
      let text = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
      text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');

      // 特定の主要コンテンツブロック（body, main, article）の中身を優先的に抽出
      let mainContentMatch = text.match(/<main[^>]*>([\s\S]*?)<\/main>/i) ||
                             text.match(/<article[^>]*>([\s\S]*?)<\/article>/i) ||
                             text.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
      
      text = mainContentMatch ? mainContentMatch[1] : text;

      // HTMLタグを全て除去
      text = text.replace(/<[^>]*>/g, '');

      // 複数行にわたる改行やスペースを1つにまとめる
      text = text.replace(/\s+/g, ' ').trim();

      return text.substring(0, MAX_CONTENT_LENGTH_FOR_GEMINI); // 長すぎる場合は truncate
    } catch (e) {
      Logger.log(`Webコンテンツ抽出エラー from ${url}: ` + e.message);
      return '';
    }
  }

  // --- メイン処理 ---
  let collectedInfoHtml = '';
  let factCheckResult = "事実確認は行われませんでした。";
  let contentForGemini = []; // Geminiに渡すための整形されたコンテンツとメタデータ

  // 全ての検索クエリでWeb検索を行い、結果からWebページ全文を取得
  for (const query of searchQueries) {
    Logger.log(`Web検索開始: ${query}`);
    let searchResults;
    try {
      searchResults = performCustomSearch(query, CUSTOM_SEARCH_API_KEY, CUSTOM_SEARCH_ENGINE_ID, MAX_SEARCH_RESULTS_PER_QUERY);
      Logger.log(`Web検索結果 (${query}): ${searchResults.length}件取得`);
    } catch (e) {
      Logger.log(`Web検索エラー (${query}): ${e.message}`);
      collectedInfoHtml += `<p><strong>Web検索エラー: ${query}</strong><br>${e.message}</p>`;
      continue;
    }

    if (searchResults.length === 0) {
      Logger.log(`Web検索結果なし: ${query}`);
      continue;
    }

    // 各検索結果をフィルタリングし、合格したものだけコンテンツ取得とGemini用に整形
    for (const item of searchResults) {
      const urlLower = item.link.toLowerCase();
      const titleLower = item.title.toLowerCase();
      const snippetLower = item.snippet.toLowerCase();

      // --- ドメインフィルタリング ---
      const domain = getHostnameFromUrl(item.link);
      if (!domain || !TRUSTED_DOMAINS.includes(domain)) {
        Logger.log(`ドメインフィルタリングによりスキップ: ${item.link} (ドメイン: ${domain || '不明'})`);
        continue;
      }

      // --- キーワード除外フィルタリング ---
      const combinedText = titleLower + ' ' + snippetLower;
      if (EXCLUDE_KEYWORDS.some(keyword => combinedText.includes(keyword))) {
        Logger.log(`キーワードフィルタリングによりスキップ: ${item.link} (キーワード検出)`);
        continue;
      }

      // フィルタリングを通過した場合のみコンテンツを抽出
      Logger.log(`コンテンツ抽出開始: ${item.link}`);
      const pageContent = extractMainContent(item.link);

      if (pageContent.length > 50) { // ある程度の長さがあるコンテンツのみ採用
        contentForGemini.push({
          title: item.title,
          url: item.link,
          snippet: item.snippet,
          full_content: pageContent
        });
        Logger.log(`コンテンツ抽出完了 (${item.link}): ${pageContent.length}文字`);
      } else {
         Logger.log(`コンテンツ抽出スキップ (短すぎるか抽出不可): ${item.link}`);
      }
    }
  }

  // 全ての収集したコンテンツをGeminiに渡して情報を生成
  if (contentForGemini.length === 0) {
    collectedInfoHtml = '<p>指定された期間のWeb検索結果から有用な情報が見つかりませんでした。</p>';
    Logger.log('有用なWeb検索結果が見つからなかったため、情報生成をスキップしました。');
  } else {
    // Geminiに渡すコンテンツをプロンプトに組み込む
    let formattedSourcesForGemini = contentForGemini.map((item, index) => {
      return `--- 参照元記事 ${index + 1} ---\nタイトル: ${item.title}\nURL: ${item.url}\n概要: ${item.snippet}\n\nコンテンツ抜粋:\n${item.full_content}\n`;
    }).join('\n\n');

    const infoGenerationPrompt = `あなたは生成AIに関する最新情報を収集・整理する専門家です。
以下の「【参照元記事】」の内容を基に、指定された期間（${targetDateString}から現在まで）における生成AIの最新情報、新サービス、技術革新、および関連する重要な動向をまとめてください。

**【厳守事項】**
1.  **参照元記事に記載されている情報のみを使用し、あなたの知っている知識や推測は一切含めないでください。**
2.  **特に信頼性が高く、日付が新しい情報（${targetDateString}以降の情報）を優先してまとめてください。**
3.  **セミナー開催、求人情報、イベント告知のような、生成AIの技術動向や新サービス・研究発表に関する直接的なニュースではない情報は、原則として含めないでください。**
4.  **各情報項目（例：新サービス名、技術革新の概要など）の直後には、その情報の出典となったWeb記事のタイトルとURLをHTMLリンク形式で必ず明記してください。** 例: <code>[情報概要] (<a href="URL">情報源タイトル</a>)</code>
5.  複数の情報源で同じ情報が言及されている場合は、最も詳細な情報を提供している、または一次情報源と思われるもの一つを厳選して記載してください。
6.  重複する情報や、重要性が低いと判断される情報は要約の過程で除外してください。

**【出力形式】**
各モダリティ（テキスト生成AI、画像生成AI、音声生成AI、その他マルチモーダルAI）ごとにセクションを設け、箇条書きで分かりやすくまとめてください。
特に重要な情報については、その背景や今後の影響についても簡潔に言及してください。
HTMLメールとして整形できるよう、適切なHTMLタグ（<h1>, <h2>, <ul>, <li>, <strong>, <br>など）を使用してください。
**絶対に、Markdownのコードブロック記法（例: \`\`\`html や \`\`\`）や、HTMLタグ以外のURLに関する具体的な記述を含めないでください。**

---
**【参照元記事】**
${formattedSourcesForGemini}
---
`;
    try {
      Logger.log('Geminiに情報生成を依頼 (RAG強化)...');
      collectedInfoHtml = callGeminiAPI(infoGenerationPrompt, GEMINI_API_KEY);
      Logger.log('Geminiによる情報生成完了。');
    } catch (e) {
      Logger.log(`Gemini情報生成エラー: ${e.message}`);
      collectedInfoHtml = `<p><strong>情報生成エラー:</strong><br>${e.message}</p>`;
    }
  }

  // --- 収集した情報の事実確認 (Geminiが生成した内容全体に対して行う) ---
  const plainTextCollectedInfo = collectedInfoHtml.replace(/<[^>]*>?/gm, '').replace(/&nbsp;/g, ' ').trim();
  const factCheckPrompt = `以下の情報はWeb検索結果を基にGeminiが生成した生成AIに関する最新情報です。この情報が事実に基づいているか、またその情報の信頼性について評価してください。
特に、情報の正確性、最新性、出典の適切性を確認し、信頼性の低い情報や誤っている可能性のある情報については、その旨を具体的に指摘し、理由を簡潔に述べてください。
情報が事実であると確認できる場合は、「信頼性：高」と評価し、特にコメントは不要です。

---
[収集された情報]
${plainTextCollectedInfo}
---

**【出力形式】**
評価結果は以下の形式で提供してください。

**評価:** [信頼性：高 / 信頼性：中 / 信頼性：低]
**コメント:** [信頼性が低い、または誤っている可能性がある場合の具体的な指摘と理由。事実確認が不要な場合は空欄]
`;

  try {
    Logger.log('収集した情報の事実確認を開始...');
    factCheckResult = callGeminiAPI(factCheckPrompt, GEMINI_API_KEY);
    Logger.log('事実確認完了。');
  } catch (e) {
    Logger.log('事実確認エラー: ' + e.message);
    factCheckResult = `事実確認中にエラーが発生しました: ${e.message}`;
  }

  // --- メール送信 ---
  const subject = `生成AI最新情報と事実確認 (${targetDateString}以降)`;

  const htmlBody = `
    <html>
      <body>
        <h1>${subject}</h1>
        <p>Web検索とGemini AIによる生成AIの最新情報をお届けします。</p>
        <hr>
        <h2>収集された最新情報</h2>
        ${collectedInfoHtml}
        <hr>
        <h2>情報の事実確認結果</h2>
        <pre>${factCheckResult}</pre>
        <hr>
        <p>このメールはGoogle Apps Scriptによって自動送信されています。</p>
      </body>
    </html>
  `;

  try {
    MailApp.sendEmail({
      to: RECIPIENT_EMAIL,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log('メールが正常に送信されました。');
  } catch (e) {
    Logger.log('メール送信エラー: ' + e.message);
    MailApp.sendEmail(RECIPIENT_EMAIL, 'GASエラー: メール送信失敗', '最新情報と事実確認結果のメール送信中にエラーが発生しました。詳細: ' + e.message);
  }
}