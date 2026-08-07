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
const EVENTS_URL = 'https://you0810jmsdf.github.io/cocola-site/events/';
const SCHEDULE_URL = 'https://you0810jmsdf.github.io/cocola-site/schedule/';
const DANGOTSUSIN_URL = 'https://you0810jmsdf.github.io/cocola-site/dangotsusin/';
const COL_TS = 0;
const COL_NICKNAME = 3;

// 旧コミットの JSON を安全に読む（無ければ null）
function readPrevJson(path) {
  try {
    return JSON.parse(execSync(`git show HEAD~1:${path}`, { encoding: 'utf8' }));
  } catch (e) {
    return null;
  }
}

// 現コミットの JSON を安全に読む（無ければ null）
function readCurJson(path) {
  try {
    return JSON.parse(fs.readFileSync(path, 'utf8'));
  } catch (e) {
    return null;
  }
}

function excerpt(value, maxLength) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (!text) return '';
  return text.length > maxLength ? text.slice(0, maxLength - 1) + '…' : text;
}

function proposalNotificationBody(proposal) {
  const detail = excerpt(proposal.detail || proposal.needs || '', 92);
  if (detail) return detail;
  return proposal.title || '新しい提案が公開されました';
}

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
      body: proposalNotificationBody(newProposals[0]),
      url: DAO_URL,
    });
  } else if (newProposals.length > 1) {
    const names = newProposals.slice(0, 3).map((p) => p.title || p.id).filter(Boolean).join(' / ');
    messages.push({
      title: `新しい提案が${newProposals.length}件`,
      body: names || 'タップして提案一覧を確認できます',
      url: DAO_URL,
    });
  }

  // 既存提案のステータス変化・投票数増加を検知
  const prevProposalMap = {};
  (prev.proposals || []).forEach((p) => { prevProposalMap[p.id] = p; });

  const STATUS_LABEL = { approved: '可決されました', expired: '期限切れになりました', voting: '投票受付中になりました' };

  (cur.proposals || []).forEach((p) => {
    const old = prevProposalMap[p.id];
    if (!old) return; // 新規提案は上の newProposals で処理済み

    // ステータス変化
    if (old.statusKey !== p.statusKey && p.statusKey !== 'voting') {
      const label = STATUS_LABEL[p.statusKey] || `ステータスが変わりました（${p.statusKey}）`;
      messages.push({
        title: `提案が${label}`,
        body: p.title || p.id,
        url: DAO_URL,
      });
    }

    // 投票数増加（voting中のみ通知。ステータス変化と重複しないよう old.statusKey も確認）
    if (p.statusKey === 'voting' && old.statusKey === 'voting' && typeof p.voteCount === 'number' && typeof old.voteCount === 'number' && p.voteCount > old.voteCount) {
      const diff = p.voteCount - old.voteCount;
      messages.push({
        title: `「${p.title || p.id}」に投票がありました`,
        body: `${diff > 1 ? diff + '票' : '1票'}追加（現在 ${p.voteCount} 票 / 賛成率 ${p.yesRate}%）`,
        url: DAO_URL,
      });
    }
  });

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

  // メンバー得点変動の検知（dao/data.json の members[].totalPoints 増加）
  // 新規メンバーは prevMemPts に存在しないため対象外（別途メンバー増加通知で扱う）。
  const prevMemPts = {};
  (prev.members || []).forEach((m) => { if (m && m.nickname) prevMemPts[m.nickname] = m.totalPoints || 0; });
  const gained = [];
  (cur.members || []).forEach((m) => {
    if (!m || !m.nickname) return;
    const before = prevMemPts[m.nickname];
    if (typeof before === 'number' && (m.totalPoints || 0) > before) {
      gained.push({ nickname: m.nickname, diff: (m.totalPoints || 0) - before, total: m.totalPoints || 0 });
    }
  });
  if (gained.length === 1) {
    const g = gained[0];
    messages.push({
      title: `${g.nickname}さんがポイントを獲得しました`,
      body: `+${g.diff} pt（累計 ${g.total} pt）`,
      url: DAO_URL,
    });
  } else if (gained.length > 1) {
    const names = gained.slice(0, 3).map((g) => g.nickname).join('・');
    const more = gained.length > 3 ? ` 他${gained.length - 3}名` : '';
    messages.push({
      title: `${gained.length}名がポイントを獲得しました`,
      body: `${names}${more}`,
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

  // events/data.json の新規イベント検知（rows、キー= タイトル|ts、過去日付は除外）
  try {
    const curEv = readCurJson('events/data.json');
    const prevEv = readPrevJson('events/data.json');
    if (prevEv && Array.isArray(prevEv.rows) && curEv && Array.isArray(curEv.rows)) {
      const prevKeys = new Set(prevEv.rows.map((r) => `${r.t}|${r.ts}`));
      const newEvents = curEv.rows.filter((r) => !r.past && !prevKeys.has(`${r.t}|${r.ts}`));
      if (newEvents.length === 1) {
        messages.push({
          title: '新しいイベント情報',
          body: newEvents[0].t || 'イベントが追加されました',
          url: EVENTS_URL,
        });
      } else if (newEvents.length > 1) {
        messages.push({
          title: `新しいイベントが${newEvents.length}件`,
          body: 'タップしてイベント一覧を確認できます',
          url: EVENTS_URL,
        });
      }
    }
  } catch (e) {
    console.log('events/data.json diff skipped:', e.message);
  }

  // schedule/data.json の新規予定検知（items、キー= id）
  try {
    const curSc = readCurJson('schedule/data.json');
    const prevSc = readPrevJson('schedule/data.json');
    if (prevSc && Array.isArray(prevSc.items) && curSc && Array.isArray(curSc.items)) {
      const prevIds = new Set(prevSc.items.map((it) => it.id));
      const newItems = curSc.items.filter((it) => !prevIds.has(it.id));
      if (newItems.length === 1) {
        messages.push({
          title: '新しい予定が登録されました',
          body: newItems[0].title || '予定が追加されました',
          url: SCHEDULE_URL,
        });
      } else if (newItems.length > 1) {
        messages.push({
          title: `新しい予定が${newItems.length}件`,
          body: 'タップしてスケジュールを確認できます',
          url: SCHEDULE_URL,
        });
      }
    }
  } catch (e) {
    console.log('schedule/data.json diff skipped:', e.message);
  }

  // dangotsusin/data.json の新規お知らせ検知（items、キー= id）
  try {
    const curDg = readCurJson('dangotsusin/data.json');
    const prevDg = readPrevJson('dangotsusin/data.json');
    if (prevDg && Array.isArray(prevDg.items) && curDg && Array.isArray(curDg.items)) {
      const prevIds = new Set(prevDg.items.map((it) => it.id));
      const newItems = curDg.items.filter((it) => !prevIds.has(it.id));
      if (newItems.length === 1) {
        messages.push({
          title: 'だんご通信に新着',
          body: newItems[0].subject || '新しいお知らせがあります',
          url: DANGOTSUSIN_URL,
        });
      } else if (newItems.length > 1) {
        messages.push({
          title: `だんご通信に新着${newItems.length}件`,
          body: 'タップして確認できます',
          url: DANGOTSUSIN_URL,
        });
      }
    }
  } catch (e) {
    console.log('dangotsusin/data.json diff skipped:', e.message);
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
