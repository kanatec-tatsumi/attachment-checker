/*
 * 添付ファイル忘れチェック（Smart Alerts / OnMessageSend）
 *
 * 送信時に本文（と件名）に「添付」などのキーワードがあり、
 * かつ添付ファイルが 1 つも無い場合に警告ダイアログを出します。
 */

// ===== 設定 =====
// 検出するキーワード（必要に応じて追加・削除してください）
const KEYWORDS = ["添付", "attached", "attachment"];

// 件名もチェック対象にするか
const CHECK_SUBJECT = true;

// 返信・転送時の「引用部分」はチェック対象から外す（相手の文中の「添付」に反応しないように）
const QUOTE_MARKERS = [
  /^-{2,}\s*Original Message\s*-{2,}/im,
  /^-{2,}\s*元のメッセージ\s*-{2,}/m,
  /^_{10,}\s*$/m,                    // Outlook の返信区切り線
  /^[\s\u00a0\u200b*]*(From|差出人|送信者)[\s\u00a0*]*[:：]/m,
  /^[\s\u00a0\u200b*]*(Sent|送信日時|送信)[\s\u00a0*]*[:：]/m,
  /^On .+ wrote:\s*$/m,
  /^\d{4}年\d{1,2}月\d{1,2}日.*[:：]\s*$/m,
  /^>/m
];

// HTML 本文で「ここから下は引用」を示す目印（新しい Outlook / Outlook on the web / Windows 版）
const HTML_QUOTE_MARKERS = [
  /<div[^>]*id=["']?appendonsend/i,
  /<div[^>]*id=["']?divRplyFwdMsg/i,
  /<div[^>]*id=["']?mail-editor-reference-message-container/i,
  /<div[^>]*class=["']?[^"'>]*gmail_quote/i,
  /<blockquote/i,
  /<hr[\s>\/]/i
];

// 返信・転送の件名の接頭辞（この場合、件名の「添付」は元メールのものなのでチェックしない）
const REPLY_SUBJECT_PREFIX = /^\s*(re|fw|fwd|返信|転送)\s*[:：]/i;

// 警告メッセージ
const WARNING_MESSAGE =
  "本文に「添付」と書かれていますが、添付ファイルがありません。\n" +
  "このまま送信してよろしいですか？";
// ================

function stripQuotedText(text) {
  let cut = text.length;
  for (const re of QUOTE_MARKERS) {
    const m = re.exec(text);
    if (m && m.index < cut) cut = m.index;
  }
  return text.substring(0, cut);
}

// HTML 本文から、自分が書いた部分（引用より上）だけをテキストで取り出す
function extractOwnTextFromHtml(html) {
  let cut = html.length;
  for (const re of HTML_QUOTE_MARKERS) {
    const m = re.exec(html);
    if (m && m.index < cut) cut = m.index;
  }
  return html
    .substring(0, cut)
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>|<\/(p|div|li|tr|h\d)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function containsKeyword(text) {
  if (!text) return false;
  const lower = text.toLowerCase();
  return KEYWORDS.some((k) => lower.includes(k.toLowerCase()));
}

function getBody(item, type) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(type, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || "") : reject(r.error)
    );
  });
}

function getSubject(item) {
  return new Promise((resolve, reject) => {
    item.subject.getAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || "") : reject(r.error)
    );
  });
}

function getAttachments(item) {
  return new Promise((resolve, reject) => {
    item.getAttachmentsAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value || []) : reject(r.error)
    );
  });
}

function onMessageSendHandler(event) {
  console.log("[添付チェック] 送信イベント受信");

  // Outlook に結果を返すのは必ず 1 回だけ
  let done = false;
  const finish = (result) => {
    if (done) return;
    done = true;
    console.log("[添付チェック] 結果:", JSON.stringify(result));
    event.completed(result);
  };

  // 万一どこかで止まっても送信を妨げないよう、5 秒で必ず完了させる
  setTimeout(() => {
    console.warn("[添付チェック] タイムアウトのためそのまま送信");
    finish({ allowEvent: true });
  }, 5000);

  try {
    const item = Office.context.mailbox.item;
    Promise.all([
      getBody(item, Office.CoercionType.Html).catch(() => null),
      getBody(item, Office.CoercionType.Text).catch(() => ""),
      CHECK_SUBJECT ? getSubject(item) : Promise.resolve(""),
      getAttachments(item)
    ])
      .then(([html, text, subject, attachments]) => {
        // HTML が取れればそこから引用部分を除き、さらにテキストの目印でも切る
        const ownText = stripQuotedText(html ? extractOwnTextFromHtml(html) : text);
        const checkSubject = CHECK_SUBJECT && !REPLY_SUBJECT_PREFIX.test(subject);
        console.log("[添付チェック] チェック対象の本文:", ownText);

        const mentionsAttachment =
          containsKeyword(ownText) || (checkSubject && containsKeyword(subject));

        // 本文中に埋め込まれた画像（署名ロゴなど）は添付として数えない
        const realAttachments = attachments.filter((a) => !a.isInline);

        if (mentionsAttachment && realAttachments.length === 0) {
          finish({ allowEvent: false, errorMessage: WARNING_MESSAGE });
        } else {
          finish({ allowEvent: true });
        }
      })
      .catch((e) => {
        console.error("[添付チェック] エラー", e);
        finish({ allowEvent: true });
      });
  } catch (e) {
    console.error("[添付チェック] エラー", e);
    finish({ allowEvent: true });
  }
}

// イベントハンドラーの登録（manifest の FunctionName と一致させる）
console.log("[添付チェック] スクリプト読み込み完了 v1.0.4");
Office.onReady(() => console.log("[添付チェック] Office 準備完了"));
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
