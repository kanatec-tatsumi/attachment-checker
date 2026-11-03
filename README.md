# 添付ファイル忘れチェック（新しい Outlook 用アドイン）

本文（または件名）に「添付」と書かれているのに添付ファイルが無いメールを送信しようとすると、警告ダイアログを表示します。
Outlook の **Smart Alerts（`OnMessageSend` イベント）** を使っているため、ボタン操作なしで送信時に自動チェックされます。

## ファイル構成

```
attachment-checker/
├─ manifest.xml          … アドインの定義（URL と Id を書き換える）
├─ commands.html         … 実行環境（office.js とスクリプトを読み込む）
├─ src/launchevent.js    … 判定ロジック（キーワードやメッセージはここで変更）
└─ assets/icon-*.png     … アイコン
```

## 動作

| 状況 | 結果 |
|---|---|
| 本文に「添付」あり・添付ファイルなし | 警告を表示（「送信しない」/「そのまま送信」を選択） |
| 本文に「添付」あり・添付ファイルあり | そのまま送信 |
| 本文に「添付」なし | そのまま送信 |

- 返信・転送時の**引用部分は対象外**（相手の文中の「添付」には反応しません）
- 署名ロゴなどの**インライン画像は添付として数えません**
- 英語の `attached` / `attachment` も検出対象（`KEYWORDS` で変更可）

## セットアップ手順

### 1. ファイルを HTTPS で公開する
アドインのファイルは HTTPS で配信する必要があります。例:
- **Azure Static Web Apps / Azure Blob Storage の静的サイト**
- **GitHub Pages**
- 社内の Web サーバー（HTTPS 必須）

ローカルで試すだけなら Office の開発用証明書を使えます:
```bash
npx office-addin-dev-certs install
npx http-server -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000
```
この場合、URL は `https://localhost:3000` になります。

### 2. manifest.xml を書き換える
- `https://YOUR-HOST/attachment-checker` → 実際の公開 URL（すべて置換）
- `<Id>` → 新しい GUID（PowerShell で `[guid]::NewGuid()`）
- `<ProviderName>` → 会社名など

### 3. Outlook に読み込む（サイドロード）
1. ブラウザで <https://aka.ms/olksideload> を開く（新しい Outlook / Outlook on the web 共通）
2. 「**個人用アドイン**」→「**カスタム アドインの追加**」→「**ファイルから追加**」
3. `manifest.xml` を選択してインストール

### 4. 組織全体に配布する場合
Microsoft 365 管理センター →「設定」→「統合アプリ」→「カスタム アプリのアップロード」から `manifest.xml` をアップロードします。
**Smart Alerts は管理者による一括展開が推奨**です（反映まで数時間かかることがあります）。

## カスタマイズ

`src/launchevent.js` 冒頭の設定を変更します。

```js
const KEYWORDS = ["添付", "attached", "attachment"]; // 例: "別添", "同封" などを追加
const CHECK_SUBJECT = true;                           // 件名もチェックするか
const WARNING_MESSAGE = "...";                        // 警告文
```

送信のブロック強度は `manifest.xml` の `SendMode` で変えられます:

| SendMode | 挙動 |
|---|---|
| `PromptUser`（既定） | 警告するが「そのまま送信」も可能 |
| `SoftBlock` | 修正するまで送信不可（アドインが動かない場合は送信可） |
| `Block` | 修正するまで送信不可（アドインが動かない場合も送信不可） |

## 注意点
- 要件: Mailbox API 1.12 以上（新しい Outlook、Outlook on the web、Microsoft 365 版 Windows Outlook は対応）
- Microsoft 365（Exchange Online）アカウントが前提です。新しい Outlook に追加した Gmail 等の外部アカウントではアドインが動作しない場合があります
- 判定でエラーが起きた場合は、送信を妨げずにそのまま送信します
- 最初の送信時はアドインの読み込みに数秒かかることがあります

## 動作確認
1. 新規メールで本文に「資料を添付します」と書き、何も添付せずに送信 → 警告が出る
2. ファイルを添付して送信 → 警告なしで送信される
3. 「添付」を含まないメールを送信 → 警告なしで送信される
