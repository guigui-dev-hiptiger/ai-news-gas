# Google Apps Script (GAS) 機能仕様  
## 生成AI最新情報収集＆メール送信スクリプト

---

## 📝 概要
本GASスクリプトは、**Web検索（Google Custom Search）** と **Google Gemini API** を活用し、生成AIに関する最新情報を収集。その情報の**事実確認を行い、HTML形式のレポートをメール送信**する自動化ツールです。

---

## 🔧 主要機能

### ✅ 設定項目の一元管理
- `GEMINI_API_KEY`：Google Gemini APIへのアクセスキー  
- `CUSTOM_SEARCH_API_KEY`：Google Custom Search APIへのアクセスキー  
- `CUSTOM_SEARCH_ENGINE_ID`：使用するカスタム検索エンジンID（cx）  
- `RECIPIENT_EMAIL`：スクリプト実行者のメールアドレス（自動取得）  
- `DAYS_PRIOR`：情報取得期間（日数、例：7日間）  
- `MAX_SEARCH_RESULTS_PER_QUERY`：1クエリあたり最大取得記事数（例：5件）

---

### ✅ APIキーのバリデーション
- 実行前に **必要なAPIキーがすべて設定されているかチェック**
- 未設定時はログ出力＆実行者にエラーメールを送信し、処理を中断

---

### ✅ 動的な検索キーワード生成
- 情報取得期間に基づいて `「2024年6月17日以降」` のように **日付付きクエリを自動生成**
- **生成AI全般＋Google AI（Gemini, DeepMind, Vertex AI等）に関するキーワードも網羅**

---

### ✅ Web検索機能（Google Custom Search API）
- 複数クエリでWeb検索を実行し、**最大数のタイトル／URL／抜粋を取得**
- **検索エラーはログ＆メールに記載**し、トラブルシュートも簡単

---

### ✅ 情報生成機能（Google Gemini API）
- Web検索結果をプロンプトに変換しGemini APIに渡す  
- 指定期間内の **最新動向／新技術／新サービス** を整理  
- **HTML形式＋出典URL付きの整形済み出力**を実現（メール対応）

---

### ✅ 事実確認機能（Google Gemini API）
- Geminiが生成した情報のプレーンテキストに対し、**再度Geminiで信頼性を検証**
- 各項目について、以下で評価：
  - `信頼性：高` / `中` / `低`  
  - 信頼性が低い項目は **理由を簡潔に記述**

---

### ✅ メール通知機能
- 情報収集結果 ＋ 事実確認結果をまとめてHTMLメールで送信  
- 件名に情報収集期間を明記  
- **メール送信失敗時にもエラーメール通知あり**

---

## 🔌 利用する外部サービス・API
- **Google Custom Search API**：Web検索用  
- **Google Gemini API**：情報生成＆事実確認用

### Google Apps Script 標準サービス
| サービス名        | 用途 |
|------------------|------|
| `PropertiesService` | APIキー管理 |
| `Session`          | 実行者のメール取得 |
| `UrlFetchApp`      | 外部APIアクセス |
| `MailApp`          | メール送信 |
| `Logger`           | ログ出力 |
| `Utilities`        | テキスト整形（プロンプト構築に使用） |

---

## ⚙️ 設定手順

1. Google Apps Scriptプロジェクトを作成  
2. コードを `.gs` ファイルに貼り付け  
3. スクリプトプロパティに以下を登録：
   - `GEMINI_API_KEY`
   - `CUSTOM_SEARCH_API_KEY`
   - `CUSTOM_SEARCH_ENGINE_ID`  
4. Google Cloud Consoleで以下のAPIを有効化：
   - Google Custom Search API  
   - Gemini API（Generative Language API）

---

## ▶️ 実行方法

- GASエディタで `getLatestAIInfoAndSendEmail` 関数を手動実行  
- または、**トリガーを設定して定期実行**も可能

---

## 👤 クレジット
このスクリプトは [@guigui](https://note.com/hip_tiger5987) によって設計・開発されました。
