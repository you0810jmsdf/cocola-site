/**
 * COCoLa DAO バックグラウンド通知 送信スクリプト（方式B）
 *
 * GitHub Actions 上で実行する。VAPID署名＋ペイロード暗号化は web-push が行う。
 *
 * 動作:
 *   - workflow_dispatch: 入力された title/body を全購読者へ手動送信
 *   - push（dao/data.json 変更時）: 直前コミットの data.json と比較し、
 *     新しい提案・新しいメンバーを検知したぶんだけ送信
 *
 * 必要な環境変数（GitHub Secrets）:
 *   VAPID_PUBLIC_KEY / VAPID_PRIVATE_KEY / VAPID_SUBJECT
 *   PUSH_GAS_URL / PUSH_API_TOKEN
 */

const webpush = require('web-push');
const fs = require('fs');
const { execSync } = require('child_process');

const {
  VAPID_PUBLIC_KEY,
  VAPID_PRIVATE_KEY,
  VAPID_SUBJECT,
  PUSH_GAS_URL,
  PUSH_API_TOKEN,
  EVENT_NAME,
  INPUT_TITLE,
  INPUT_BODY,
} = process.env;

const DAO_URL = 'https://you0810jmsdf.github.io/cocola-site/dao/';
const MEMBER_APP_URL = 'https://you0810jmsdf.github.io/cocola-site/member-app/html.html';
const COL_TS = 0;
const COL_NICKNAME = 3;

if (!VAPID_PUBLIC_KEY || !VAPID_PRIVATE_KEY || !PUSH_GAS_URL || !PUSH_API_TOKEN) {
  console.error('Missing required env vars (VAPID/PUSH).');
  process.exit(1);
}

webpush.setVapidDetails(
  VAPID_SUBJECT || 'mailto:you0810jmsdf@gmail.com',
  VAPID_PUBLIC_KEY,
  VAPID_PRIVATE_KEY
);

async function getSubscribers() {
  const url = `${PUSH_GAS_URL}?mode=list&token=${encodeURIComponent(PUSH_API_TOKEN)}&t=${Date.now()}`;
  const res = await fetch(url);
  const json = await res.json();
  if (!json.ok) throw new Error('list failed: ' + JSON.stringify(json));
  return json.subscribers || [];
}

async function deactivate(endpoint) {
  try {
    await fetch(`${PUSH_GAS_URL}?mode=unsubscribe&endpoint=${encodeURIComponent(endpoint)}&t=${Date.now()}`);
  } catch (e) {
    /* 失効通知の除去失敗は致命的でない */
  }
}

function buildMessages() {
  if (EVENT_NAME === 'workflow_dispatch') {
    return [{
      title: INPUT_TITLE || 'COCoLa DAO',
      body: INPUT_BODY || 'お知らせがあります',
      url: DAO_URL,
    }];
  }

  // push トリガー：data.json の差分検知
  let cur;
  try {
    cur = JSON.parse(fs.readFileSync('dao/data.json', 'utf8'));
  } catch (e) {
    console.log('current data.json not readable; skip.');
    return [];
  }

  let prev = null;
  try {
    prev = JSON.parse(execSync('git show HEAD~1:dao/data.json', { encoding: 'utf8' }));
  } catch (e) {
    console.log('no previous data.json (first commit); baseline only, skip.');
    return [];
  }

  const messages = [];
  const prevIds = (prev.proposals || []).map((p) => p.id);
  const newProposals = (cur.proposals || []).filter((p) => prevIds.indexOf(p.id) === -1);

  if (newProposals.length === 1) {
    messages.push({
      title: '新しい提案があります',
      body: newProposals[0].title || '新しい提案が公開されました',
      url: DAO_URL,
    });
  } else if (newProposals.length > 1) {
    messages.push({
      title: `新しい提案が${newProposals.length}件`,
      body: 'タップして提案一覧を確認できます',
      url: DAO_URL,
    });
  }

  const prevMembers = (prev.stats && prev.stats.members) || 0;
  const curMembers = (cur.stats && cur.stats.members) || 0;
  // prevMembers===0（パース失敗や初期状態）からの増加は誤検知のため通知しない。
  if (prevMembers > 0 && curMembers > prevMembers) {
    messages.push({
      title: 'メンバーが増えました',
      body: `新しく${curMembers - prevMembers}名が加わりました（現在 ${curMembers} 名）`,
      url: DAO_URL,
    });
  }

  // members/data.json の rows 差分検知（新規メンバー登録の即時通知）
  try {
    const curMem = JSON.parse(fs.readFileSync('members/data.json', 'utf8'));
    let prevMem = null;
    try {
      prevMem = JSON.parse(execSync('git show HEAD~1:members/data.json', { encoding: 'utf8' }));
    } catch (e) {
      prevMem = null;
    }
    if (prevMem && Array.isArray(prevMem.rows) && Array.isArray(curMem.rows)) {
      const prevKeys = new Set(prevMem.rows.map((r) => `${r[COL_TS]}|${r[COL_NICKNAME]}`));
      const newRows = curMem.rows.filter((r) => !prevKeys.has(`${r[COL_TS]}|${r[COL_NICKNAME]}`));
      if (newRows.length === 1) {
        const nick = newRows[0][COL_NICKNAME] || '新メンバー';
        messages.push({
          title: '新しいメンバーが登録されました',
          body: `${nick}さんがCOCoLaに参加しました`,
          url: MEMBER_APP_URL,
        });
      } else if (newRows.length > 1) {
        messages.push({
          title: `新しいメンバーが${newRows.length}名登録されました`,
          body: 'タップしてメンバー一覧を確認できます',
          url: MEMBER_APP_URL,
        });
      }
    }
  } catch (e) {
    console.log('members/data.json diff skipped:', e.message);
  }

  return messages;
}

(async () => {
  const messages = buildMessages();
  if (messages.length === 0) {
    console.log('No notifications to send.');
    return;
  }

  const subs = await getSubscribers();
  console.log(`Subscribers: ${subs.length}, Messages: ${messages.length}`);
  if (subs.length === 0) {
    console.log('No subscribers; nothing to send.');
    return;
  }

  let sent = 0;
  let removed = 0;
  for (const msg of messages) {
    const payload = JSON.stringify({
      title: msg.title,
      body: msg.body,
      url: msg.url,
      tag: 'cocola-' + Date.now(),
    });
    for (const sub of subs) {
      try {
        await webpush.sendNotification({ endpoint: sub.endpoint, keys: sub.keys }, payload);
        sent++;
      } catch (err) {
        if (err.statusCode === 404 || err.statusCode === 410) {
          await deactivate(sub.endpoint);
          removed++;
        } else {
          console.error('send error:', err.statusCode, err.body || err.message);
        }
      }
    }
  }
  console.log(`Done. Sent: ${sent}, Removed(expired): ${removed}`);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
