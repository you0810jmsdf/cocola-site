/**
 * COCoLa DAO points prototype.
 *
 * This script adds a lightweight DAO-style contribution ledger on top of the
 * existing Google Sheets workflow. It does not change the public TOPICS columns.
 *
 * First run:
 * 1. Paste this file into the bound Apps Script project.
 * 2. Run setupDaoSheets().
 * 3. Run processDaoPointsFromAllRows() against a test spreadsheet first.
 * 4. After verification, create an installable form-submit trigger for onDaoEventSubmit.
 */

var DAO_CONFIG = {
  spreadsheetId: '1Xma1V92uNPTcXmj1cDNzU0jePcxAtTR_4vEF2ZXRgdo',
  memberSpreadsheetId: '13FwQyLeZK_Kgvl6LqY4UhzMkxJJUkLLi01vhA2jwUcc',
  memberSourceSheetName: 'フォームの回答 1',

  // Prefer TOPICS if the normalized sheet exists. If not, raw form responses are supported.
  candidateEventSheets: ['TOPICS', 'フォームの回答 1'],
  candidateProposalSheets: ['DAO提案'],
  candidateProposalResponseSheets: ['DAO提案フォーム回答', 'DAO提案フォームの回答 1', '提案フォームの回答 1'],
  candidateVoteResponseSheets: ['DAO投票フォーム回答', 'DAO投票フォームの回答 1', '投票フォームの回答 1'],

  memberSheetName: 'DAOメンバー',
  ledgerSheetName: 'DAOポイント履歴',
  aliasSheetName: 'DAOメンバー別名',
  reviewSheetName: 'DAO名寄せ候補',
  proposalSheetName: 'DAO提案',
  voteSheetName: 'DAO投票',
  citizenLikeSheetName: 'DAO市民いいね',
  proposalFormTitle: 'COCoLa DAO 提案フォーム',
  voteFormTitle: 'COCoLa DAO 投票フォーム',
  proposalFormUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSdxfz3OyN2FJ6x89viHAyKU57glOMPMuFGdXNjp2O6motziWQ/viewform',
  voteFormUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSc9kOcabaSr8bjhE8pRDzIr1AWDEVodSc2lPWwkMAY029aAIw/viewform',
  proposalPageUrl: 'https://you0810jmsdf.github.io/cocola-site/dao/',
  proposalNotifyProperty: 'DAO_PROPOSAL_NOTIFY_TO',
  defaultAdminPassword: 'cocola2026',
  processedHeader: 'DAOポイント付与済',
  proposalProcessedHeader: 'DAO提案者ポイント付与済',
  proposalApprovedProcessedHeader: 'DAO提案承認ポイント付与済',
  proposalIdHeader: '提案ID',
  proposalStatusHeader: 'ステータス',
  voteProcessedHeader: 'DAO投票集計済',
  voteParticipationProcessedHeader: 'DAO投票参加ポイント付与済',
  backupPrefix: 'DAO_BACKUP_',

  // Organizer points are awarded only when the organizer text uniquely matches
  // a member nickname/name or an approved alias.
  awardOrganizerPoints: true,
  awardProposerPoints: true,
  awardProposalApprovalPoints: true,
  awardVoteParticipationPoints: true,
  proposalRequiresApproval: false,
  organizerPoints: 8,
  participantPoints: 1,
  proposerPoints: 3,
  proposalApprovalPoints: 5,
  voteParticipationPoints: 1,
  proposalApprovalThreshold: 0.5,

  activeMonths: 6,
  supporterMonths: 12,
};

var DAO_MEMBER_HEADERS = [
  'メンバーID',
  '氏名',
  '累計ポイント',
  '直近6ヶ月ポイント',
  '最終活動日',
  'ステータス',
  '投票権重み',
];

var DAO_LEDGER_HEADERS = [
  '記録日時',
  'イベントシート',
  'イベント行',
  '開催日時',
  'イベント名',
  '氏名',
  '活動区分',
  'ポイント',
  '備考',
];

var DAO_ALIAS_HEADERS = [
  '別名',
  '正式ニックネーム',
  '備考',
];

var DAO_REVIEW_HEADERS = [
  '記録日時',
  'イベントシート',
  'イベント行',
  'イベント名',
  '入力名',
  '活動区分',
  '候補',
  '理由',
  '承認ニックネーム',
  '処理済',
  '処理メモ',
];

var DAO_PROPOSAL_HEADERS = [
  '提案ID',
  '記録日時',
  '提案タイトル',
  'カテゴリ',
  '理由・詳細',
  '必要なもの',
  '提案者',
  'ステータス',
  '公開',
  '賛成重み',
  '反対重み',
  '有効投票権重み',
  '賛成率',
  '投票数',
  '判定日時',
  '処理メモ',
  'DAO提案者ポイント付与済',
  'DAO提案承認ポイント付与済',
];

var DAO_VOTE_HEADERS = [
  '記録日時',
  '提案ID',
  '提案タイトル',
  '投票者',
  '投票',
  'コメント',
  '名寄せ済み氏名',
  '投票者重み',
  'DAO投票集計済',
  'DAO投票参加ポイント付与済',
  '処理メモ',
];

var DAO_CITIZEN_LIKE_HEADERS = [
  '記録日時',
  '提案ID',
  'ブラウザキー',
  'ユーザーエージェント',
  '参照元',
];

var DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY = 'DAO_PROCESSED_PROPOSAL_RESPONSE_KEYS';
var DAO_VOTE_RESPONSE_KEYS_PROPERTY = 'DAO_PROCESSED_VOTE_RESPONSE_KEYS';
var DAO_PROPOSAL_PAGE_CACHE_KEY = 'DAO_PROPOSAL_PAGE_DATA_JSON';
var DAO_MEMBER_ONLY_FORM_NOTICE = '提案・投票はCOCoLaメンバー向けです。メンバー確認ができない回答は、提案一覧・投票結果・DAOポイントには反映されません。メンバー以外の方は、提案ページの「市民いいね」で応援や関心をお知らせください。';

function onOpen(e) {
  addDaoMenu_();
  addTopicsMenu_();
}

function addTopicsMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('TOPICS管理')
      .addItem('フォーム回答 → TOPICS を全件同期', 'menuRebuildTopicsFromResponses')
      .addToUi();
  } catch (error) {
    Logger.log('TOPICSメニュー作成エラー: ' + error);
  }
}

function menuRebuildTopicsFromResponses() {
  rebuildTopicsFromResponses();
  SpreadsheetApp.getUi().alert('TOPICS同期完了', 'フォームの回答をTOPICSシートへ同期しました。\n5分以内にサイトへ反映されます。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function doGetDao_(e) {
  var mode = e && e.parameter && e.parameter.mode;
  if (mode === 'json') {
    var forceRefresh = isTruthy_((e.parameter && (e.parameter.refresh || e.parameter.force)) || '');
    return ContentService
      .createTextOutput(getDaoProposalPageDataJson_(forceRefresh))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'cancel') {
    var id = ((e.parameter && e.parameter.id) || '').trim();
    var pw = (e.parameter && e.parameter.pw) || '';
    return ContentService
      .createTextOutput(JSON.stringify(cancelDaoProposal_(id, pw)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'restore') {
    var restorePw = (e.parameter && e.parameter.pw) || '';
    var restoreIds = ((e.parameter && e.parameter.ids) || '').trim();
    return ContentService
      .createTextOutput(JSON.stringify(restoreDaoProposals_(restoreIds, restorePw)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'like') {
    return ContentService
      .createTextOutput(JSON.stringify(recordDaoCitizenLike_(e && e.parameter ? e.parameter : {})))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'refreshFormNotices') {
    var noticePw = (e.parameter && e.parameter.pw) || '';
    return ContentService
      .createTextOutput(JSON.stringify(refreshDaoFormMemberOnlyNotices_(noticePw)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'debugRecentProposalResponses') {
    var debugPw = (e.parameter && e.parameter.pw) || '';
    var debugLimit = Number((e.parameter && e.parameter.limit) || 10);
    return ContentService
      .createTextOutput(JSON.stringify(debugRecentDaoProposalResponses_(debugPw, debugLimit)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createTemplateFromFile('dao-proposals')
    .evaluate()
    .setTitle('COCoLa DAO 提案と投票')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cancelDaoProposal_(id, pw) {
  var correctPw = getDaoAdminPassword_();
  if (!correctPw) return { ok: false, error: 'DAO_ADMIN_PASSWORD がスクリプトプロパティに設定されていません' };
  if (pw !== correctPw) return { ok: false, error: 'パスワードが違います' };
  if (!id) return { ok: false, error: '提案IDが指定されていません' };

  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var sheet = ss.getSheetByName(DAO_CONFIG.proposalSheetName);
  if (!sheet) return { ok: false, error: 'DAO提案シートが見つかりません' };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalIdHeader, '提案ID']);
  var statusCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalStatusHeader, 'ステータス']);
  if (idCol < 0 || statusCol < 0) return { ok: false, error: '列が見つかりません' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: '提案データがありません' };
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === id) {
      sheet.getRange(i + 2, statusCol + 1).setValue('終了');
      refreshDaoVoteFormChoicesOnly_(ss);
      clearDaoProposalPageCache_();
      return { ok: true, changed: 1 };
    }
  }
  return { ok: false, error: '提案が見つかりません: ' + id };
}

function restoreDaoProposals_(idsText, pw) {
  var correctPw = getDaoAdminPassword_();
  if (!correctPw) return { ok: false, error: 'DAO_ADMIN_PASSWORD がスクリプトプロパティに設定されていません' };
  if (pw !== correctPw) return { ok: false, error: 'パスワードが違います' };
  if (!idsText) return { ok: false, error: '復旧する提案IDが指定されていません' };

  var ids = {};
  idsText.split(',').forEach(function(id) {
    id = String(id || '').trim();
    if (id) ids[id] = true;
  });

  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var sheet = ss.getSheetByName(DAO_CONFIG.proposalSheetName);
  if (!sheet) return { ok: false, error: 'DAO提案シートが見つかりません' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalIdHeader, '提案ID']);
  var statusCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalStatusHeader, 'ステータス']);
  if (idCol < 0 || statusCol < 0) return { ok: false, error: '列が見つかりません' };

  var lastRow = sheet.getLastRow();
  var data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  var changed = 0;
  for (var i = 0; i < data.length; i++) {
    if (ids[String(data[i][idCol]).trim()]) {
      sheet.getRange(i + 2, statusCol + 1).setValue('投票中');
      changed += 1;
    }
  }
  refreshDaoVoteFormChoicesOnly_(ss);
  clearDaoProposalPageCache_();
  return { ok: true, changed: changed };
}

function getDaoProposalPageData(options) {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  if (options && options.refresh) {
    refreshDaoProposalPageData_(ss);
  }
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var voteWeights = buildDaoVotingWeightIndex_(ss);
  var forms = getDaoProposalPageFormUrls_();
  var citizenLikeCounts = buildDaoCitizenLikeCounts_(ensureDaoCitizenLikeSheet_(ss));
  var proposals = collectDaoProposalPageRows_(proposalSheet, citizenLikeCounts);

  var openCount = 0;
  var approvedCount = 0;
  var totalVotes = 0;
  var totalCitizenLikes = 0;
  proposals.forEach(function(proposal) {
    if (proposal.statusKey === 'approved') approvedCount += 1;
    if (proposal.statusKey === 'voting') openCount += 1;
    totalVotes += Number(proposal.voteCount) || 0;
    totalCitizenLikes += Number(proposal.citizenLikeCount) || 0;
  });

  return {
    generatedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    eligibleWeight: voteWeights.totalWeight || 0,
    approvalThresholdPercent: Math.round((DAO_CONFIG.proposalApprovalThreshold || 0.5) * 100),
    proposalFormUrl: forms.proposalFormUrl,
    voteFormUrl: forms.voteFormUrl,
    stats: {
      proposals: proposals.length,
      voting: openCount,
      approved: approvedCount,
      votes: totalVotes,
      citizenLikes: totalCitizenLikes,
    },
    proposals: proposals,
  };
}

function getDaoProposalPageDataJson_(forceRefresh) {
  var cache = CacheService.getScriptCache();
  var props = PropertiesService.getScriptProperties();
  if (!forceRefresh) {
    var cached = cache.get(DAO_PROPOSAL_PAGE_CACHE_KEY);
    if (cached) return cached;
    cached = props.getProperty(DAO_PROPOSAL_PAGE_CACHE_KEY);
    if (cached && !isDaoProposalPageJsonStale_(cached)) {
      cache.put(DAO_PROPOSAL_PAGE_CACHE_KEY, cached, 120);
      return cached;
    }
  }

  return storeDaoProposalPageDataJson_(getDaoProposalPageData({ refresh: !!forceRefresh }));
}

function refreshDaoProposalPageCache_() {
  return storeDaoProposalPageDataJson_(getDaoProposalPageData({ refresh: false }));
}

function storeDaoProposalPageDataJson_(data) {
  var json = JSON.stringify(data)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  var cache = CacheService.getScriptCache();
  var props = PropertiesService.getScriptProperties();
  cache.put(DAO_PROPOSAL_PAGE_CACHE_KEY, json, 120);
  props.setProperty(DAO_PROPOSAL_PAGE_CACHE_KEY, json);
  return json;
}

function isDaoProposalPageJsonStale_(json) {
  try {
    var data = JSON.parse(json);
    if (!data.generatedAt) return true;
    var match = String(data.generatedAt).match(/^(\d{4})\/(\d{2})\/(\d{2})\s+(\d{2}):(\d{2})$/);
    if (!match) return true;
    var generatedAt = new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4]),
      Number(match[5]),
      0
    );
    return (new Date().getTime() - generatedAt.getTime()) > 30 * 60 * 1000;
  } catch (error) {
    return true;
  }
}

function refreshDaoProposalPageData_(ss) {
  normalizeRecentDaoProposalResponses_(ss, 20);
  normalizeRecentDaoVoteResponses_(ss, 40);
  processDaoProposalVotesFast_(ss);
  processLatestDaoProposalPoint_(ss);
  refreshDaoVoteFormChoicesOnly_(ss);
  clearDaoProposalPageCache_();
}

function clearDaoProposalPageCache_() {
  CacheService.getScriptCache().remove(DAO_PROPOSAL_PAGE_CACHE_KEY);
  PropertiesService.getScriptProperties().deleteProperty(DAO_PROPOSAL_PAGE_CACHE_KEY);
}

function getDaoAdminPassword_() {
  var props = PropertiesService.getScriptProperties();
  var password = props.getProperty('DAO_ADMIN_PASSWORD');
  if (!password && DAO_CONFIG.defaultAdminPassword) {
    password = DAO_CONFIG.defaultAdminPassword;
    props.setProperty('DAO_ADMIN_PASSWORD', password);
  }
  return password || '';
}

function recordDaoCitizenLike_(params) {
  var proposalId = normalizeMemberName_(params.id || params.proposalId || '');
  var browserKey = normalizeDaoCitizenLikeKey_(params.key || params.clientKey || '');
  if (!proposalId) return { ok: false, error: '提案IDが指定されていません' };
  if (!browserKey) return { ok: false, error: 'ブラウザキーがありません。ページを再読み込みしてからもう一度押してください。' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'ただいま処理が混み合っています。少し待ってからもう一度押してください。' };

  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    var proposalSheet = ensureDaoProposalSheet_(ss);
    if (!isDaoPublicProposalId_(proposalSheet, proposalId)) {
      return { ok: false, error: '公開中の提案が見つかりません' };
    }

    var likeSheet = ensureDaoCitizenLikeSheet_(ss);
    var existing = findDaoCitizenLikeRow_(likeSheet, proposalId, browserKey);
    if (!existing) {
      likeSheet.appendRow([
        new Date(),
        proposalId,
        browserKey,
        String(params.ua || params.userAgent || '').slice(0, 300),
        String(params.ref || params.referrer || '').slice(0, 300),
      ]);
      clearDaoProposalPageCache_();
    }

    var count = countDaoCitizenLikesForProposal_(likeSheet, proposalId);
    return { ok: true, liked: true, duplicate: !!existing, id: proposalId, count: count };
  } finally {
    lock.releaseLock();
  }
}

function isDaoPublicProposalId_(sheet, proposalId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0];
  for (var row = 2; row <= values.length; row += 1) {
    var rowValues = values[row - 1];
    var currentId = normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['提案ID']));
    if (currentId !== proposalId) continue;
    var isPublic = isTruthy_(getFirstFromRowValues_(headers, rowValues, ['公開', '公開チェック', '掲載', '掲載可']));
    var status = normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['ステータス', '状態', '判定', '結果'])) || '投票中';
    return isPublic && getDaoProposalPageStatusKey_(status) !== 'closed';
  }
  return false;
}

function normalizeDaoCitizenLikeKey_(value) {
  return String(value || '').trim().replace(/[^a-zA-Z0-9._:-]/g, '').slice(0, 120);
}

function findDaoCitizenLikeRow_(sheet, proposalId, browserKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, 1, lastRow - 1, DAO_CITIZEN_LIKE_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i += 1) {
    if (normalizeMemberName_(values[i][1]) === proposalId && normalizeDaoCitizenLikeKey_(values[i][2]) === browserKey) {
      return i + 2;
    }
  }
  return 0;
}

function countDaoCitizenLikesForProposal_(sheet, proposalId) {
  var counts = buildDaoCitizenLikeCounts_(sheet);
  return counts[proposalId] || 0;
}

function buildDaoCitizenLikeCounts_(sheet) {
  var counts = {};
  var seen = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return counts;
  var values = sheet.getRange(2, 1, lastRow - 1, DAO_CITIZEN_LIKE_HEADERS.length).getValues();
  values.forEach(function(row) {
    var proposalId = normalizeMemberName_(row[1]);
    var browserKey = normalizeDaoCitizenLikeKey_(row[2]);
    if (!proposalId || !browserKey) return;
    var key = proposalId + '|' + browserKey;
    if (seen[key]) return;
    seen[key] = true;
    counts[proposalId] = (counts[proposalId] || 0) + 1;
  });
  return counts;
}

function collectDaoProposalPageRows_(sheet, citizenLikeCounts) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0];
  citizenLikeCounts = citizenLikeCounts || {};

  var rows = [];
  for (var row = 2; row <= values.length; row += 1) {
    var rowValues = values[row - 1];
    var isPublic = isTruthy_(getFirstFromRowValues_(headers, rowValues, ['公開', '公開チェック', '掲載', '掲載可']));
    if (!isPublic) continue;

    var proposalId = normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['提案ID'])) || ensureDaoProposalId_(sheet, row);
    var title = normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']));
    if (!proposalId || !title) continue;

    var status = normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['ステータス', '状態', '判定', '結果'])) || '投票中';
    var statusKey = getDaoProposalPageStatusKey_(status);
    if (statusKey === 'closed') continue;
    var yesWeight = numberFromRow_(getFirstFromRowValues_(headers, rowValues, ['賛成重み']));
    var noWeight = numberFromRow_(getFirstFromRowValues_(headers, rowValues, ['反対重み']));
    var eligibleWeight = numberFromRow_(getFirstFromRowValues_(headers, rowValues, ['有効投票権重み']));
    var yesRate = numberFromRow_(getFirstFromRowValues_(headers, rowValues, ['賛成率']));
    var voteCount = numberFromRow_(getFirstFromRowValues_(headers, rowValues, ['投票数']));
    var submittedAt = formatDaoPageDate_(getFirstFromRowValues_(headers, rowValues, ['記録日時', 'タイムスタンプ', 'Timestamp']));
    var decidedAt = formatDaoPageDate_(getFirstFromRowValues_(headers, rowValues, ['判定日時']));

    rows.push({
      id: proposalId,
      title: title,
      category: normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['カテゴリ'])) || 'その他',
      detail: String(getFirstFromRowValues_(headers, rowValues, ['理由・詳細', '理由', '詳細', 'なぜやりたいか']) || '').trim(),
      needs: String(getFirstFromRowValues_(headers, rowValues, ['必要なもの', '人・場所・費用など']) || '').trim(),
      proposer: normalizeMemberName_(getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム'])) || '未記入',
      status: status,
      statusKey: statusKey,
      yesWeight: yesWeight,
      noWeight: noWeight,
      eligibleWeight: eligibleWeight,
      yesRate: Math.round(yesRate * 1000) / 10,
      progressPercent: clampDaoPagePercent_(yesRate * 100),
      voteCount: voteCount,
      citizenLikeCount: citizenLikeCounts[proposalId] || 0,
      submittedAt: submittedAt,
      decidedAt: decidedAt,
      memo: String(getFirstFromRowValues_(headers, rowValues, ['処理メモ']) || '').trim(),
    });
  }

  rows = dedupeDaoProposalPageRows_(rows);
  rows.sort(function(a, b) {
    var statusOrder = { voting: 0, approved: 1, pending: 2, closed: 3 };
    if (statusOrder[a.statusKey] !== statusOrder[b.statusKey]) {
      return statusOrder[a.statusKey] - statusOrder[b.statusKey];
    }
    return a.id < b.id ? 1 : -1;
  });
  return rows;
}

function dedupeDaoProposalPageRows_(rows) {
  var byKey = {};
  rows.forEach(function(row) {
    var key = createDaoProposalContentKey_(row.title, row.proposer, row.category, row.detail, row.needs);
    if (!key) return;
    var current = byKey[key];
    if (!current || shouldPreferDaoProposalRow_(row, current)) {
      byKey[key] = row;
    }
  });
  return Object.keys(byKey).map(function(key) { return byKey[key]; });
}

function shouldPreferDaoProposalRow_(candidate, current) {
  var candidateVotes = Number(candidate.voteCount) || 0;
  var currentVotes = Number(current.voteCount) || 0;
  if (candidateVotes !== currentVotes) return candidateVotes > currentVotes;
  var candidateWeight = (Number(candidate.yesWeight) || 0) + (Number(candidate.noWeight) || 0);
  var currentWeight = (Number(current.yesWeight) || 0) + (Number(current.noWeight) || 0);
  if (candidateWeight !== currentWeight) return candidateWeight > currentWeight;
  return normalizeProposalKey_(candidate.id) < normalizeProposalKey_(current.id);
}

function createDaoProposalContentKey_(title, proposer, category, detail, needs) {
  var parts = [
    normalizeProposalDuplicateKey_(title),
    normalizeProposalDuplicateKey_(proposer),
    normalizeProposalDuplicateKey_(category),
    normalizeProposalDuplicateKey_(detail),
    normalizeProposalDuplicateKey_(needs),
  ].filter(function(part) { return Boolean(part); });
  return parts.length >= 2 ? parts.join('|') : '';
}

function normalizeProposalDuplicateKey_(value) {
  return normalizeHeader_(value).toLowerCase();
}

function getDaoProposalPageFormUrls_() {
  var props = PropertiesService.getScriptProperties();
  return {
    proposalFormUrl: props.getProperty('DAO_PROPOSAL_FORM_URL') || DAO_CONFIG.proposalFormUrl || '',
    voteFormUrl: props.getProperty('DAO_VOTE_FORM_URL') || DAO_CONFIG.voteFormUrl || '',
  };
}

function getDaoProposalPageStatusKey_(status) {
  if (isDaoProposalStatusApproved_(status)) return 'approved';
  var text = normalizeMemberName_(status);
  if (['投票中', '受付中', '募集中'].indexOf(text) !== -1) return 'voting';
  if (['下書き', '確認中', '未公開', '保留'].indexOf(text) !== -1) return 'pending';
  return 'closed';
}

function numberFromRow_(value) {
  var number = Number(value);
  return isNaN(number) ? 0 : number;
}

function clampDaoPagePercent_(value) {
  var number = Number(value) || 0;
  if (number < 0) return 0;
  if (number > 100) return 100;
  return Math.round(number * 10) / 10;
}

function formatDaoPageDate_(value) {
  var date = parseDaoDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function addDaoMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('DAO運営')
      .addItem('名寄せ承認を反映', 'menuApplyDaoReviewAliasApprovals')
      .addItem('未処理ポイントを付与', 'menuProcessDaoPointsFromAllRows')
      .addItem('未処理の提案者ポイントを付与', 'menuProcessDaoProposalPointsFromAllRows')
      .addItem('提案投票を集計', 'menuProcessDaoProposalVotes')
      .addSeparator()
      .addItem('提案・投票フォームを作成', 'menuSetupDaoProposalForms')
      .addItem('投票フォームの提案リスト更新', 'menuUpdateDaoVoteFormChoices')
      .addItem('提案・投票フォームの注意書き更新', 'menuRefreshDaoFormMemberOnlyNotices')
      .addItem('提案・投票の自動更新を設定', 'menuSetupDaoProposalAutomation')
      .addItem('付与予定をプレビュー', 'menuPreviewDaoPointsFromAllRows')
      .addItem('DAOシート初期設定', 'menuSetupDaoSheets')
      .addItem('手動バックアップ作成', 'menuCreateDaoBackup')
      .addToUi();
  } catch (error) {
    Logger.log('DAOメニュー作成エラー: ' + error);
  }
}

function menuApplyDaoReviewAliasApprovals() {
  applyDaoReviewAliasApprovals();
  showDaoMenuMessage_('名寄せ承認をDAOメンバー別名へ反映しました。続けて「未処理ポイントを付与」を実行してください。');
}

function menuProcessDaoPointsFromAllRows() {
  processDaoPointsFromAllRows();
  showDaoMenuMessage_('未処理行のポイント付与を実行しました。DAOメンバーとDAOポイント履歴を確認してください。');
}

function menuProcessDaoProposalPointsFromAllRows() {
  processDaoProposalPointsFromAllRows();
  showDaoMenuMessage_('未処理の提案者ポイント付与を実行しました。DAOメンバーとDAOポイント履歴を確認してください。');
}

function menuProcessDaoProposalVotes() {
  var result = processDaoProposalVotesFast();
  showDaoMenuMessage_('提案投票を集計しました。承認: ' + result.approved + '件 / 投票対象: ' + result.proposals + '件');
}

function menuSetupDaoProposalForms() {
  var result = setupDaoProposalForms();
  showDaoMenuMessage_('提案・投票フォームを作成しました。提案フォーム: ' + result.proposalFormUrl + ' / 投票フォーム: ' + result.voteFormUrl);
}

function menuUpdateDaoVoteFormChoices() {
  var count = updateDaoVoteFormChoices();
  showDaoMenuMessage_('投票フォームの提案リストを更新しました。選択肢: ' + count + '件');
}

function menuRefreshDaoFormMemberOnlyNotices() {
  var result = updateDaoFormMemberOnlyNotices();
  showDaoMenuMessage_('提案・投票フォームの注意書きを更新しました。提案フォーム: ' + (result.proposalUpdated ? '更新' : '未検出') + ' / 投票フォーム: ' + (result.voteUpdated ? '更新' : '未検出'));
}

function menuSetupDaoProposalAutomation() {
  var result = setupDaoProposalAutomationTrigger();
  showDaoMenuMessage_('提案・投票の自動更新を設定しました。トリガー: ' + result.triggers + '件');
}

function menuPreviewDaoPointsFromAllRows() {
  var rows = previewDaoPointsFromAllRows();
  showDaoMenuMessage_('付与予定プレビューを実行しました。Apps Scriptの実行ログを確認してください。対象者数: ' + rows.length);
}

function menuSetupDaoSheets() {
  setupDaoSheets();
  showDaoMenuMessage_('DAOシート初期設定を実行しました。DAO関連シートとバックアップを確認してください。');
}

function menuCreateDaoBackup() {
  var label = createDaoBackup('menu_manual');
  showDaoMenuMessage_('バックアップを作成しました: ' + label);
}

function showDaoMenuMessage_(message) {
  try {
    SpreadsheetApp.getActive().toast(message, 'DAO運営', 8);
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(message);
  }
}

/** 一度だけ実行してください。実行後はこの関数を削除してもOKです。 */
function setDaoAdminPassword() {
  PropertiesService.getScriptProperties().setProperty('DAO_ADMIN_PASSWORD', 'cocola2026');
  Logger.log('DAO_ADMIN_PASSWORD を設定しました');
}

function setupDaoSheets() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  createDaoBackup('setup');
  ensureDaoMemberSheet_(ss);
  ensureDaoLedgerSheet_(ss);
  ensureDaoAliasSheet_(ss);
  ensureDaoReviewSheet_(ss);
  ensureDaoProposalSheet_(ss);
  ensureDaoVoteSheet_(ss);
  ensureDaoCitizenLikeSheet_(ss);
  ensureProcessedColumns_(ss);
}

function setupDaoDailyTrigger() {
  removeDaoTriggers();
  ScriptApp.newTrigger('processDaoPointsFromAllRows')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log('DAOポイント日次トリガーを設定しました: 毎日9時ごろ');
}

function removeDaoTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (handler === 'processDaoPointsFromAllRows' || handler === 'onDaoEventSubmit' || handler === 'onDaoSpreadsheetFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function setupDaoProposalAutomationTrigger() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var exists = false;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onDaoSpreadsheetFormSubmit') exists = true;
  });
  if (!exists) {
    ScriptApp.newTrigger('onDaoSpreadsheetFormSubmit')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();
  }
  return {
    triggers: ScriptApp.getProjectTriggers().filter(function(trigger) {
      return trigger.getHandlerFunction() === 'onDaoSpreadsheetFormSubmit';
    }).length,
  };
}

function onDaoSpreadsheetFormSubmit(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    var submittedSheetName = getDaoSubmittedSheetName_(e);
    var isProposalResponse = isDaoSubmittedSheetName_(submittedSheetName, DAO_CONFIG.candidateProposalResponseSheets);
    var isVoteResponse = isDaoSubmittedSheetName_(submittedSheetName, DAO_CONFIG.candidateVoteResponseSheets);
    var proposalSheet = ensureDaoProposalSheet_(ss);
    var proposalLastRowBefore = proposalSheet.getLastRow();
    ensureDaoVoteSheet_(ss);

    if (isVoteResponse && !isProposalResponse) {
      if (normalizeLatestDaoVoteResponse_(ss)) {
        processDaoProposalVotesLatest_(ss);
      }
      refreshDaoProposalPageCache_();
      return;
    }

    if (isProposalResponse && !isVoteResponse) {
      if (normalizeLatestDaoProposalResponse_(ss)) {
        processLatestDaoProposalPoint_(ss);
        refreshDaoVoteFormChoicesOnly_(ss);
        notifyNewDaoProposalRows_(ss, proposalSheet, proposalLastRowBefore + 1, proposalSheet.getLastRow());
      }
      refreshDaoProposalPageCache_();
      return;
    }

    normalizeDaoProposalResponses_(ss);
    notifyNewDaoProposalRows_(ss, proposalSheet, proposalLastRowBefore + 1, proposalSheet.getLastRow());
    normalizeDaoVoteResponses_(ss);
    processDaoProposalVotesFast_(ss);
    refreshDaoVoteFormChoicesOnly_(ss);
    refreshDaoProposalPageCache_();
  } finally {
    lock.releaseLock();
  }
}

function getDaoSubmittedSheetName_(e) {
  try {
    if (e && e.range && e.range.getSheet) return e.range.getSheet().getName();
  } catch (error) {
    Logger.log('DAO submit sheet name detection failed: ' + error.message);
  }
  return '';
}

function isDaoSubmittedSheetName_(sheetName, candidateNames) {
  return Boolean(sheetName) && candidateNames.indexOf(sheetName) !== -1;
}

function setDaoProposalNotifyTo(address) {
  PropertiesService.getScriptProperties().setProperty(DAO_CONFIG.proposalNotifyProperty, String(address || '').trim());
}

function notifyNewDaoProposalRows_(ss, proposalSheet, startRow, endRow) {
  var recipients = getDaoProposalNotificationRecipients_();
  if (recipients.length === 0) {
    Logger.log('DAO proposal notification skipped: no member emails found.');
    return;
  }
  if (!proposalSheet || endRow < startRow) return;
  for (var row = startRow; row <= endRow; row += 1) {
    var proposal = getDaoProposalNotificationData_(proposalSheet, row);
    if (!proposal.id || !proposal.title) continue;
    try {
      sendDaoProposalNotification_(recipients, proposal);
    } catch (error) {
      Logger.log('DAO proposal notification send error: ' + error.message);
    }
  }
}

function getDaoProposalNotificationRecipients_() {
  var manual = PropertiesService.getScriptProperties().getProperty(DAO_CONFIG.proposalNotifyProperty);
  if (manual) return uniqueDaoEmails_(manual.split(/[,\s;]+/));

  var emails = [];
  try {
    var memberSs = SpreadsheetApp.openById(DAO_CONFIG.memberSpreadsheetId);
    var sheet = memberSs.getSheetByName(DAO_CONFIG.memberSourceSheetName) || memberSs.getSheets()[0];
    if (!sheet || sheet.getLastRow() < 2) return [];
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), 8)).getValues();
    rows.forEach(function(row) {
      if (!String(row[3] || '').trim()) return;
      emails.push(row[7]);
      emails.push(row[1]);
    });
  } catch (error) {
    Logger.log('DAO proposal notification recipients error: ' + error.message);
  }
  return uniqueDaoEmails_(emails);
}

function uniqueDaoEmails_(values) {
  var seen = {};
  var emails = [];
  values.forEach(function(value) {
    String(value || '').split(/[,\s;]+/).forEach(function(part) {
      var email = part.trim();
      if (!email || email.indexOf('@') === -1) return;
      var key = email.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      emails.push(email);
    });
  });
  return emails;
}

function getDaoProposalNotificationData_(proposalSheet, row) {
  var values = proposalSheet.getRange(row, 1, 1, DAO_PROPOSAL_HEADERS.length).getValues()[0];
  return {
    id: normalizeMemberName_(values[0]),
    submittedAt: formatDaoPageDate_(values[1]),
    title: normalizeMemberName_(values[2]),
    category: normalizeMemberName_(values[3]),
    detail: String(values[4] || '').trim(),
    needs: String(values[5] || '').trim(),
    proposer: normalizeMemberName_(values[6]),
  };
}

function sendDaoProposalNotification_(recipients, proposal) {
  var subject = '[COCoLa DAO] 新しい提案: ' + proposal.title;
  var body = [
    'COCoLa DAOに新しい提案が届きました。',
    '',
    '提案ID: ' + proposal.id,
    'タイトル: ' + proposal.title,
    '提案者: ' + (proposal.proposer || '未記入'),
    'カテゴリー: ' + (proposal.category || '未設定'),
    proposal.submittedAt ? '提出日: ' + proposal.submittedAt : '',
    '',
    '理由・詳細:',
    proposal.detail || '未記入',
    '',
    '必要なもの:',
    proposal.needs || '未記入',
    '',
    '提案と投票ページ:',
    DAO_CONFIG.proposalPageUrl,
    '',
    '投票フォーム:',
    getDaoProposalPageFormUrls_().voteFormUrl || DAO_CONFIG.voteFormUrl,
  ].filter(function(line) { return line !== ''; }).join('\n');

  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    bcc: recipients.join(','),
    subject: subject,
    body: body,
    name: 'COCoLa DAO',
  });
}

function applyDaoReviewAliasApprovals() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  createDaoBackup('alias_approval');
  var reviewSheet = ensureDaoReviewSheet_(ss);
  var aliasSheet = ensureDaoAliasSheet_(ss);
  var memberIndex = buildDaoMemberIndex_(ss);
  var canonicalSet = {};
  memberIndex.members.forEach(function(member) {
    canonicalSet[member.nickname] = true;
  });

  var lastRow = reviewSheet.getLastRow();
  if (lastRow < 2) return;

  var values = reviewSheet.getRange(2, 1, lastRow - 1, DAO_REVIEW_HEADERS.length).getValues();
  values.forEach(function(row, index) {
    var sheetRow = index + 2;
    var inputName = normalizeMemberName_(row[4]);
    var approvedNickname = normalizeMemberName_(row[8]);
    var processed = isTruthy_(row[9]);
    if (processed || !inputName || !approvedNickname) return;

    if (!canonicalSet[approvedNickname]) {
      reviewSheet.getRange(sheetRow, 11).setValue('承認ニックネームがメンバー紹介サイトに見つかりません');
      return;
    }

    appendDaoAliasIfMissing_(aliasSheet, inputName, approvedNickname, 'DAO名寄せ候補から承認');
    reviewSheet.getRange(sheetRow, 10).setValue(true);
    reviewSheet.getRange(sheetRow, 11).setValue('DAOメンバー別名へ反映済み');
  });
}

function processDaoPointsFromAllRows() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    var eventSheet = getDaoEventSheet_(ss);
    createDaoBackup('process');

    ensureDaoMemberSheet_(ss);
    ensureDaoLedgerSheet_(ss);
    ensureDaoAliasSheet_(ss);
    ensureDaoReviewSheet_(ss);

    if (eventSheet) {
      var processedCol = ensureProcessedColumn_(eventSheet);
      var lastRow = eventSheet.getLastRow();
      for (var row = 2; row <= lastRow; row += 1) {
        if (isTruthy_(eventSheet.getRange(row, processedCol).getValue())) continue;
        processDaoEventRow_(ss, eventSheet, row, processedCol);
      }
    }
    processDaoProposalPointsFromAllRows_(ss);
  } finally {
    lock.releaseLock();
  }
}

function processDaoProposalPointsFromAllRows() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    createDaoBackup('proposal_process');
    ensureDaoMemberSheet_(ss);
    ensureDaoLedgerSheet_(ss);
    ensureDaoAliasSheet_(ss);
    ensureDaoReviewSheet_(ss);
    processDaoProposalPointsFromAllRows_(ss);
  } finally {
    lock.releaseLock();
  }
}

function processDaoProposalVotes() {
  return processLatestDaoProposalVote();
}

function processLatestDaoProposalVote() {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    ensureDaoProposalSheet_(ss);
    ensureDaoVoteSheet_(ss);
    normalizeLatestDaoProposalResponse_(ss);
    normalizeLatestDaoVoteResponse_(ss);
    var result = processDaoProposalVotesLatest_(ss);
    clearDaoProposalPageCache_();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function processDaoProposalVotesFull() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    createDaoBackup('proposal_vote');
    ensureDaoMemberSheet_(ss);
    ensureDaoLedgerSheet_(ss);
    ensureDaoAliasSheet_(ss);
    ensureDaoReviewSheet_(ss);
    ensureDaoProposalSheet_(ss);
    ensureDaoVoteSheet_(ss);
    normalizeDaoProposalResponses_(ss);
    normalizeDaoVoteResponses_(ss);
    var result = processDaoProposalVotes_(ss);
    processDaoProposalPointsFromAllRows_(ss);
    clearDaoProposalPageCache_();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function processDaoProposalVotesFast() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    ensureDaoProposalSheet_(ss);
    ensureDaoVoteSheet_(ss);
    normalizeDaoProposalResponses_(ss);
    normalizeDaoVoteResponses_(ss);
    var result = processDaoProposalVotesFast_(ss);
    refreshDaoVoteFormChoicesOnly_(ss);
    clearDaoProposalPageCache_();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function setupDaoProposalForms() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  setupDaoSheets();
  var proposalForm = createDaoProposalForm_(ss);
  var voteForm = createDaoVoteForm_(ss);
  PropertiesService.getScriptProperties().setProperties({
    DAO_PROPOSAL_FORM_ID: proposalForm.getId(),
    DAO_VOTE_FORM_ID: voteForm.getId(),
    DAO_PROPOSAL_FORM_URL: proposalForm.getPublishedUrl(),
    DAO_VOTE_FORM_URL: voteForm.getPublishedUrl(),
  });
  return {
    proposalFormUrl: proposalForm.getPublishedUrl(),
    proposalFormEditUrl: proposalForm.getEditUrl(),
    voteFormUrl: voteForm.getPublishedUrl(),
    voteFormEditUrl: voteForm.getEditUrl(),
  };
}

function updateDaoVoteFormChoices() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  normalizeDaoProposalResponses_(ss);
  return refreshDaoVoteFormChoicesOnly_(ss);
}

function updateDaoFormMemberOnlyNotices() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var result = { proposalUpdated: false, voteUpdated: false };
  var proposalForm = getDaoProposalForm_(ss);
  if (proposalForm) {
    configureDaoProposalForm_(proposalForm);
    result.proposalUpdated = true;
  }
  var voteForm = getDaoVoteForm_(ss);
  if (voteForm) {
    configureDaoVoteForm_(voteForm, ss);
    result.voteUpdated = true;
  }
  return result;
}

function refreshDaoFormMemberOnlyNotices_(pw) {
  var correctPw = getDaoAdminPassword_();
  if (!correctPw) return { ok: false, error: 'DAO_ADMIN_PASSWORD がスクリプトプロパティに設定されていません' };
  if (pw !== correctPw) return { ok: false, error: 'パスワードが違います' };
  var result = updateDaoFormMemberOnlyNotices();
  result.ok = true;
  return result;
}

function debugRecentDaoProposalResponses_(pw, limit) {
  var correctPw = getDaoAdminPassword_();
  if (!correctPw) return { ok: false, error: 'DAO_ADMIN_PASSWORD がスクリプトプロパティに設定されていません' };
  if (pw !== correctPw) return { ok: false, error: 'パスワードが違います' };

  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var existing = buildExistingProposalResponseKeys_(proposalSheet);
  var existingContent = buildExistingDaoProposalContentKeys_(proposalSheet);
  var memberIndex = buildDaoMemberIndex_(ss);
  var result = [];
  getDaoProposalResponseSheets_(ss).forEach(function(responseSheet) {
    var lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return;
    var startRow = Math.max(2, lastRow - Math.max(Number(limit) || 10, 1) + 1);
    var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    var values = responseSheet.getRange(startRow, 1, lastRow - startRow + 1, responseSheet.getLastColumn()).getValues();
    for (var i = 0; i < values.length; i += 1) {
      var row = startRow + i;
      var rowValues = values[i];
      var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
      var title = getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
      var proposer = getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
      var category = getFirstFromRowValues_(headers, rowValues, ['カテゴリ']);
      var detail = getFirstFromRowValues_(headers, rowValues, ['理由・詳細', '理由', '詳細', 'なぜやりたいか']);
      var needs = getFirstFromRowValues_(headers, rowValues, ['必要なもの', '人・場所・費用など']);
      var key = createFormResponseKey_(responseSheet.getName(), row, timestamp, title, proposer);
      var match = matchDaoMember_(proposer, memberIndex);
      var contentKey = createDaoProposalContentKey_(title, proposer, category, detail, needs);
      result.push({
        sheet: responseSheet.getName(),
        row: row,
        timestamp: String(timestamp || ''),
        title: String(title || ''),
        proposer: String(proposer || ''),
        memberStatus: match.status,
        memberName: match.canonical || '',
        memberReason: match.reason || '',
        processed: !!existing[key],
        duplicatePublicContent: !!(contentKey && existingContent[contentKey]),
      });
    }
  });
  return { ok: true, rows: result };
}

function refreshDaoVoteFormChoicesOnly_(ss) {
  var form = getDaoVoteForm_(ss);
  if (!form) throw new Error('投票フォームが見つかりません。先に setupDaoProposalForms() を実行してください。');
  return configureDaoVoteForm_(form, ss);
}

function createDaoBackup(reason) {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var label = DAO_CONFIG.backupPrefix + stamp + '_' + sanitizeBackupLabel_(reason || 'manual');
  var targets = getDaoBackupTargetSheetNames_(ss);

  targets.forEach(function(sheetName) {
    var source = ss.getSheetByName(sheetName);
    if (!source) return;
    var copy = source.copyTo(ss);
    copy.setName(label + '__' + sheetName);
    copy.hideSheet();
  });

  Logger.log('DAOバックアップ作成: ' + label);
  return label;
}

function createManualBeforeSetupBackup() {
  return createDaoBackup('manual_before_setup');
}

function listDaoBackups() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var prefix = DAO_CONFIG.backupPrefix;
  var backups = {};
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf(prefix) !== 0) return;
    var parts = name.split('__');
    if (parts.length < 2) return;
    backups[parts[0]] = backups[parts[0]] || [];
    backups[parts[0]].push(parts.slice(1).join('__'));
  });
  Object.keys(backups).sort().forEach(function(label) {
    Logger.log(label + ' : ' + backups[label].join(', '));
  });
  return backups;
}

function restoreDaoBackup(backupLabel) {
  if (!backupLabel || backupLabel.indexOf(DAO_CONFIG.backupPrefix) !== 0) {
    throw new Error('復元するバックアップラベルを指定してください。例: DAO_BACKUP_20260416_120000_process');
  }

  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var backupSheets = ss.getSheets().filter(function(sheet) {
    return sheet.getName().indexOf(backupLabel + '__') === 0;
  });
  if (backupSheets.length === 0) {
    throw new Error('バックアップが見つかりません: ' + backupLabel);
  }

  createDaoBackup('before_restore');

  backupSheets.forEach(function(backupSheet) {
    var originalName = backupSheet.getName().substring((backupLabel + '__').length);
    var target = ss.getSheetByName(originalName) || ss.insertSheet(originalName);
    copyDaoSheetContents_(backupSheet, target);
  });

  Logger.log('DAOバックアップから復元しました: ' + backupLabel);
}

function getDaoBackupTargetSheetNames_(ss) {
  var names = [
    DAO_CONFIG.memberSheetName,
    DAO_CONFIG.ledgerSheetName,
    DAO_CONFIG.aliasSheetName,
    DAO_CONFIG.reviewSheetName,
    DAO_CONFIG.proposalSheetName,
    DAO_CONFIG.voteSheetName,
  ];
  DAO_CONFIG.candidateEventSheets.forEach(function(name) {
    if (names.indexOf(name) === -1) names.push(name);
  });
  DAO_CONFIG.candidateProposalSheets.forEach(function(name) {
    if (names.indexOf(name) === -1) names.push(name);
  });
  DAO_CONFIG.candidateProposalResponseSheets.forEach(function(name) {
    if (names.indexOf(name) === -1) names.push(name);
  });
  DAO_CONFIG.candidateVoteResponseSheets.forEach(function(name) {
    if (names.indexOf(name) === -1) names.push(name);
  });
  return names.filter(function(name) {
    return Boolean(ss.getSheetByName(name));
  });
}

function copyDaoSheetContents_(source, target) {
  target.clear({ contentsOnly: false });

  var sourceMaxRows = source.getMaxRows();
  var sourceMaxCols = source.getMaxColumns();
  if (target.getMaxRows() < sourceMaxRows) {
    target.insertRowsAfter(target.getMaxRows(), sourceMaxRows - target.getMaxRows());
  }
  if (target.getMaxRows() > sourceMaxRows) {
    target.deleteRows(sourceMaxRows + 1, target.getMaxRows() - sourceMaxRows);
  }
  if (target.getMaxColumns() < sourceMaxCols) {
    target.insertColumnsAfter(target.getMaxColumns(), sourceMaxCols - target.getMaxColumns());
  }
  if (target.getMaxColumns() > sourceMaxCols) {
    target.deleteColumns(sourceMaxCols + 1, target.getMaxColumns() - sourceMaxCols);
  }

  var range = source.getRange(1, 1, sourceMaxRows, sourceMaxCols);
  var targetRange = target.getRange(1, 1, sourceMaxRows, sourceMaxCols);
  range.copyTo(targetRange);
  target.setFrozenRows(source.getFrozenRows());
  target.setFrozenColumns(source.getFrozenColumns());
}

function sanitizeBackupLabel_(value) {
  return String(value || 'manual').replace(/[^A-Za-z0-9_]/g, '_').substring(0, 32);
}

function previewDaoPointsFromAllRows() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var eventSheet = getDaoEventSheet_(ss);
  if (!eventSheet) throw new Error('イベントシートが見つかりません');
  var memberIndex = buildDaoMemberIndex_(ss);

  var lastRow = eventSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('DAOポイント対象行はありません');
    return [];
  }

  var summary = {};
  for (var row = 2; row <= lastRow; row += 1) {
    var entries = collectDaoEntriesFromRow_(ss, eventSheet, row, memberIndex, false);
    entries.forEach(function(entry) {
      if (!summary[entry.name]) {
        summary[entry.name] = {
          name: entry.name,
          points: 0,
          count: 0,
          lastEvent: '',
        };
      }
      summary[entry.name].points += entry.points;
      summary[entry.name].count += 1;
      summary[entry.name].lastEvent = entry.eventTitle || summary[entry.name].lastEvent;
    });
  }

  getDaoProposalSheets_(ss).forEach(function(proposalSheet) {
    var proposalLastRow = proposalSheet.getLastRow();
    for (var proposalRow = 2; proposalRow <= proposalLastRow; proposalRow += 1) {
      var proposalEntry = collectDaoProposalEntryFromRow_(ss, proposalSheet, proposalRow, memberIndex, false);
      if (!proposalEntry || proposalEntry.waitingApproval || proposalEntry.unresolvedCount > 0) return;
      if (!summary[proposalEntry.name]) {
        summary[proposalEntry.name] = {
          name: proposalEntry.name,
          points: 0,
          count: 0,
          lastEvent: '',
        };
      }
      summary[proposalEntry.name].points += proposalEntry.points;
      summary[proposalEntry.name].count += 1;
      summary[proposalEntry.name].lastEvent = proposalEntry.proposalTitle || summary[proposalEntry.name].lastEvent;
      if (DAO_CONFIG.awardProposalApprovalPoints && proposalEntry.approved) {
        summary[proposalEntry.name].points += DAO_CONFIG.proposalApprovalPoints;
        summary[proposalEntry.name].count += 1;
      }
    }
  });

  var rows = Object.keys(summary).map(function(name) { return summary[name]; });
  rows.sort(function(a, b) {
    if (b.points !== a.points) return b.points - a.points;
    return a.name > b.name ? 1 : -1;
  });

  Logger.log('DAOポイント プレビュー: ' + rows.length + '名');
  rows.forEach(function(item) {
    Logger.log(item.name + ' / ' + item.points + 'pt / ' + item.count + '件 / 最終: ' + item.lastEvent);
  });
  return rows;
}

function onDaoEventSubmit(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
    ensureDaoMemberSheet_(ss);
    ensureDaoLedgerSheet_(ss);
    ensureDaoAliasSheet_(ss);
    ensureDaoReviewSheet_(ss);

    var eventSheet = null;
    var row = -1;
    if (e && e.range) {
      eventSheet = e.range.getSheet();
      row = e.range.getRow();
    } else {
      eventSheet = getDaoEventSheet_(ss);
      row = eventSheet ? eventSheet.getLastRow() : -1;
    }
    if (!eventSheet || row < 2) return;

    if (isDaoProposalSheet_(eventSheet)) {
      var proposalProcessedCol = ensureProposalProcessedColumn_(eventSheet);
      var proposalApprovedProcessedCol = ensureProposalApprovedProcessedColumn_(eventSheet);
      processDaoProposalRow_(ss, eventSheet, row, proposalProcessedCol, proposalApprovedProcessedCol, buildDaoMemberIndex_(ss));
      return;
    }

    var processedCol = ensureProcessedColumn_(eventSheet);
    if (isTruthy_(eventSheet.getRange(row, processedCol).getValue())) return;
    processDaoEventRow_(ss, eventSheet, row, processedCol);
  } finally {
    lock.releaseLock();
  }
}

function refreshDaoMemberStatuses() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.spreadsheetId);
  var memberSheet = ensureDaoMemberSheet_(ss);
  var ledgerSheet = ensureDaoLedgerSheet_(ss);
  var ledgerValues = ledgerSheet.getDataRange().getValues();
  var recentPointsByName = {};
  var cutoff = addMonths_(new Date(), -DAO_CONFIG.activeMonths);

  for (var i = 1; i < ledgerValues.length; i += 1) {
    var activityDate = parseDaoDate_(ledgerValues[i][3]) || parseDaoDate_(ledgerValues[i][0]);
    var name = normalizeMemberName_(ledgerValues[i][5]);
    var points = Number(ledgerValues[i][7]) || 0;
    if (!name || !activityDate || activityDate < cutoff) continue;
    recentPointsByName[name] = (recentPointsByName[name] || 0) + points;
  }

  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return;
  var values = memberSheet.getRange(2, 1, lastRow - 1, DAO_MEMBER_HEADERS.length).getValues();
  var updated = values.map(function(row) {
    var name = normalizeMemberName_(row[1]);
    row[3] = recentPointsByName[name] || 0;
    row[5] = getDaoStatus_(parseDaoDate_(row[4]));
    row[6] = getVotingWeight_(row[5], row[3], row[2]);
    return row;
  });
  memberSheet.getRange(2, 1, updated.length, DAO_MEMBER_HEADERS.length).setValues(updated);
}

function processDaoEventRow_(ss, eventSheet, row, processedCol) {
  var memberIndex = buildDaoMemberIndex_(ss);
  var entries = collectDaoEntriesFromRow_(ss, eventSheet, row, memberIndex, true);
  var eventTitle = entries.length ? entries[0].eventTitle : '';
  var eventDate = entries.length ? entries[0].eventDate : '';
  var ledgerSheet = ensureDaoLedgerSheet_(ss);

  if (entries.length === 0) {
    if (entries.unresolvedCount === 0) {
      eventSheet.getRange(row, processedCol).setValue(true);
    } else {
      Logger.log('未解決の名前があるためDAOポイント付与を保留: ' + eventSheet.getName() + ':' + row);
    }
    return;
  }

  var submittedDate = entries[0].submittedDate;
  var activityDate = parseDaoDate_(eventDate, submittedDate ? submittedDate.getFullYear() : null) || submittedDate || new Date();
  entries.forEach(function(entry) {
    if (hasLedgerEventKey_(ledgerSheet, entry.note)) return;
    addDaoPoints_(ss, entry.name, entry.points, activityDate);
    ledgerSheet.appendRow([
      new Date(),
      eventSheet.getName(),
      row,
      eventDate || activityDate,
      eventTitle || '',
      entry.name,
      entry.type,
      entry.points,
      entry.note,
    ]);
  });

  refreshDaoMemberStatuses();
  if (entries.unresolvedCount === 0) {
    eventSheet.getRange(row, processedCol).setValue(true);
  } else {
    Logger.log('一部の名前が未解決のため行は未処理のまま残します: ' + eventSheet.getName() + ':' + row);
  }
}

function processDaoProposalPointsFromAllRows_(ss) {
  if (!DAO_CONFIG.awardProposerPoints) return;
  var proposalSheets = getDaoProposalSheets_(ss);
  if (proposalSheets.length === 0) return;

  var memberIndex = buildDaoMemberIndex_(ss);
  proposalSheets.forEach(function(proposalSheet) {
    var processedCol = ensureProposalProcessedColumn_(proposalSheet);
    var approvedProcessedCol = ensureProposalApprovedProcessedColumn_(proposalSheet);
    var lastRow = proposalSheet.getLastRow();
    if (lastRow < 2) return;

    for (var row = 2; row <= lastRow; row += 1) {
      processDaoProposalRow_(ss, proposalSheet, row, processedCol, approvedProcessedCol, memberIndex);
    }
  });
}

function processLatestDaoProposalPoint_(ss) {
  if (!DAO_CONFIG.awardProposerPoints) return;
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var lastRow = proposalSheet.getLastRow();
  if (lastRow < 2) return;
  var processedCol = ensureProposalProcessedColumn_(proposalSheet);
  var approvedProcessedCol = ensureProposalApprovedProcessedColumn_(proposalSheet);
  processDaoProposalRow_(ss, proposalSheet, lastRow, processedCol, approvedProcessedCol, buildDaoMemberIndex_(ss));
}

function processDaoProposalRow_(ss, proposalSheet, row, processedCol, approvedProcessedCol, memberIndex) {
  var entry = collectDaoProposalEntryFromRow_(ss, proposalSheet, row, memberIndex, true);
  if (!entry) return;

  if (entry.unresolvedCount > 0) {
    Logger.log('提案者名が未解決のためDAO提案者ポイント付与を保留: ' + proposalSheet.getName() + ':' + row);
    return;
  }
  if (entry.waitingApproval) {
    Logger.log('提案が承認待ちのためDAO提案者ポイント付与を保留: ' + proposalSheet.getName() + ':' + row);
    return;
  }

  var ledgerSheet = ensureDaoLedgerSheet_(ss);
  var didAddPoints = false;
  if (!isTruthy_(proposalSheet.getRange(row, processedCol).getValue())) {
    didAddPoints = appendDaoPointEntryIfMissing_(ss, ledgerSheet, proposalSheet.getName(), row, entry);
    proposalSheet.getRange(row, processedCol).setValue(true);
  }
  if (DAO_CONFIG.awardProposalApprovalPoints
    && entry.approved
    && !isTruthy_(proposalSheet.getRange(row, approvedProcessedCol).getValue())) {
    var approvalEntry = createDaoProposalApprovalEntry_(entry);
    didAddPoints = appendDaoPointEntryIfMissing_(ss, ledgerSheet, proposalSheet.getName(), row, approvalEntry) || didAddPoints;
    proposalSheet.getRange(row, approvedProcessedCol).setValue(true);
  }
  if (didAddPoints) refreshDaoMemberStatuses();
}

function collectDaoProposalEntryFromRow_(ss, proposalSheet, row, memberIndex, writeReview) {
  var rowObject = getDaoRowObject_(proposalSheet, row);
  var timestamp = getFirstByAliases_(rowObject, ['タイムスタンプ', 'Timestamp', '提案日時', '送信日時']);
  var proposalTitle = getFirstByAliases_(rowObject, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
  var proposer = getFirstByAliases_(rowObject, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  var approved = getFirstByAliases_(rowObject, ['公開', '公開チェック', '掲載', '掲載可', '承認', '承認済', '管理者承認']);
  var submittedDate = parseDaoDate_(timestamp) || new Date();
  var eventKey = proposalSheet.getName() + ':' + row + ':' + String(proposalTitle || '');
  var proposerMatch = matchDaoMember_(proposer, memberIndex);

  if (DAO_CONFIG.proposalRequiresApproval && !isTruthy_(approved)) {
    return { waitingApproval: true, unresolvedCount: 0 };
  }
  if (proposerMatch.status === 'blank') return null;
  if (proposerMatch.status !== 'matched') {
    if (writeReview && shouldReviewName_(proposerMatch)) {
      appendDaoReview_(ss, proposalSheet.getName(), row, proposalTitle, proposer, '提案作成', proposerMatch);
    }
    return { unresolvedCount: 1 };
  }

  return {
    name: proposerMatch.canonical,
    type: '提案作成',
    points: DAO_CONFIG.proposerPoints,
    note: createLedgerEntryKey_(eventKey, '提案作成', proposerMatch.canonical),
    proposalTitle: proposalTitle,
    activityDate: submittedDate,
    approved: isTruthy_(approved) || isDaoProposalApprovedFromRow_(rowObject),
    unresolvedCount: 0,
    waitingApproval: false,
  };
}

function isDaoProposalApprovedFromRow_(rowObject) {
  var status = getFirstByAliases_(rowObject, ['ステータス', '状態', '判定', '結果']);
  return isDaoProposalStatusApproved_(status);
}

function createDaoProposalApprovalEntry_(entry) {
  return {
    name: entry.name,
    type: '提案承認',
    points: DAO_CONFIG.proposalApprovalPoints,
    note: entry.note.replace(':提案作成:', ':提案承認:'),
    proposalTitle: entry.proposalTitle,
    activityDate: new Date(),
  };
}

function appendDaoPointEntryIfMissing_(ss, ledgerSheet, sheetName, row, entry) {
  if (hasLedgerEventKey_(ledgerSheet, entry.note)) return false;
  addDaoPoints_(ss, entry.name, entry.points, entry.activityDate);
  ledgerSheet.appendRow([
    new Date(),
    sheetName,
    row,
    entry.activityDate,
    entry.proposalTitle || '',
    entry.name,
    entry.type,
    entry.points,
    entry.note,
  ]);
  return true;
}

function awardDaoVoteParticipationPoint_(ss, voteSheet, row, proposal, canonical, timestamp) {
  if (!DAO_CONFIG.awardVoteParticipationPoints || !proposal || !canonical) return false;
  var ledgerSheet = ensureDaoLedgerSheet_(ss);
  var note = createDaoVoteParticipationLedgerKey_(proposal, canonical);
  if (hasLedgerEventKey_(ledgerSheet, note)) return false;
  var activityDate = parseDaoDate_(timestamp) || new Date();
  addDaoPoints_(ss, canonical, DAO_CONFIG.voteParticipationPoints, activityDate);
  ledgerSheet.appendRow([
    new Date(),
    voteSheet.getName(),
    row,
    activityDate,
    proposal.title || proposal.id || '',
    canonical,
    'DAO投票参加',
    DAO_CONFIG.voteParticipationPoints,
    note,
  ]);
  refreshDaoMemberStatuses();
  return true;
}

function hasDaoVoteParticipationPoint_(ss, proposal, canonical) {
  if (!proposal || !canonical) return false;
  return hasLedgerEventKey_(ensureDaoLedgerSheet_(ss), createDaoVoteParticipationLedgerKey_(proposal, canonical));
}

function createDaoVoteParticipationLedgerKey_(proposal, canonical) {
  var proposalId = proposal.id || proposal.title || '';
  return createLedgerEntryKey_('DAO投票:' + proposalId, 'DAO投票参加', canonical);
}

function processDaoProposalVotes_(ss) {
  refreshDaoMemberStatuses();
  normalizeDaoProposalResponses_(ss);
  normalizeDaoVoteResponses_(ss);
  var proposalSheets = getDaoProposalSheets_(ss);
  var voteSheet = ensureDaoVoteSheet_(ss);
  var proposalIndex = buildDaoProposalIndex_(proposalSheets);
  var voterWeights = buildDaoVotingWeightIndex_(ss);
  var memberIndex = buildDaoMemberIndex_(ss);
  var voteSummary = collectDaoVotes_(ss, voteSheet, proposalIndex, voterWeights, memberIndex);
  var eligibleTotal = voterWeights.totalWeight;
  var approvedCount = 0;
  var proposalCount = 0;

  proposalSheets.forEach(function(proposalSheet) {
    ensureProposalSystemColumns_(proposalSheet);
    var lastRow = proposalSheet.getLastRow();
    if (lastRow < 2) return;

    for (var row = 2; row <= lastRow; row += 1) {
      var proposalId = ensureDaoProposalId_(proposalSheet, row);
      if (!proposalId) continue;
      proposalCount += 1;
      var summary = voteSummary[proposalId] || { yesWeight: 0, noWeight: 0, count: 0 };
      var yesRate = eligibleTotal > 0 ? summary.yesWeight / eligibleTotal : 0;
      var approved = eligibleTotal > 0 && summary.yesWeight > eligibleTotal * DAO_CONFIG.proposalApprovalThreshold;
      var status = getDaoProposalStatus_(proposalSheet, row);
      if (approved && !isDaoProposalStatusApproved_(status)) {
        setDaoProposalStatus_(proposalSheet, row, '承認');
        approvedCount += 1;
      } else if (!status) {
        setDaoProposalStatus_(proposalSheet, row, '投票中');
      } else if (isDaoProposalStatusApproved_(status)) {
        approvedCount += 1;
      }
      setDaoProposalAggregate_(proposalSheet, row, summary, eligibleTotal, yesRate, approved);
    }
  });

  return {
    proposals: proposalCount,
    approved: approvedCount,
    eligibleWeight: eligibleTotal,
  };
}

function processDaoProposalVotesFast_(ss) {
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var voteSheet = ensureDaoVoteSheet_(ss);
  var proposals = loadDaoProposalsForFastVote_(proposalSheet);
  var voterWeights = buildDaoVotingWeightIndex_(ss);
  var memberIndex = buildDaoMemberIndex_(ss);
  var voteSummary = collectDaoVotesFast_(ss, voteSheet, proposals.index, voterWeights, memberIndex);
  var eligibleTotal = voterWeights.totalWeight;
  var approvedCount = 0;
  var proposalCount = proposals.rows.length;
  var statusCol = ensureNamedColumn_(proposalSheet, DAO_CONFIG.proposalStatusHeader);
  var yesWeightCol = ensureNamedColumn_(proposalSheet, '賛成重み');
  var noWeightCol = ensureNamedColumn_(proposalSheet, '反対重み');
  var eligibleCol = ensureNamedColumn_(proposalSheet, '有効投票権重み');
  var yesRateCol = ensureNamedColumn_(proposalSheet, '賛成率');
  var voteCountCol = ensureNamedColumn_(proposalSheet, '投票数');
  var judgedAtCol = ensureNamedColumn_(proposalSheet, '判定日時');
  var memoCol = ensureNamedColumn_(proposalSheet, '処理メモ');

  proposals.rows.forEach(function(item) {
    var summary = voteSummary[item.id] || { yesWeight: 0, noWeight: 0, count: 0 };
    var yesRate = eligibleTotal > 0 ? summary.yesWeight / eligibleTotal : 0;
    var approved = eligibleTotal > 0 && summary.yesWeight > eligibleTotal * DAO_CONFIG.proposalApprovalThreshold;
    var nextStatus = item.status;
    if (approved && !isDaoProposalStatusApproved_(item.status)) nextStatus = '承認';
    if (!nextStatus) nextStatus = '投票中';
    if (isDaoProposalStatusApproved_(nextStatus)) approvedCount += 1;

    proposalSheet.getRange(item.row, statusCol).setValue(nextStatus);
    proposalSheet.getRange(item.row, yesWeightCol).setValue(summary.yesWeight || 0);
    proposalSheet.getRange(item.row, noWeightCol).setValue(summary.noWeight || 0);
    proposalSheet.getRange(item.row, eligibleCol).setValue(eligibleTotal || 0);
    proposalSheet.getRange(item.row, yesRateCol).setValue(yesRate || 0);
    proposalSheet.getRange(item.row, voteCountCol).setValue(summary.count || 0);
    proposalSheet.getRange(item.row, judgedAtCol).setValue(new Date());
    proposalSheet.getRange(item.row, memoCol).setValue(approved ? '重み付き過半数により承認' : '投票受付中');
  });

  return {
    proposals: proposalCount,
    approved: approvedCount,
    eligibleWeight: eligibleTotal,
  };
}

function processDaoProposalVotesLatest_(ss) {
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var voteSheet = ensureDaoVoteSheet_(ss);
  var lastVoteRow = voteSheet.getLastRow();
  if (lastVoteRow < 2) {
    return { proposals: Math.max(proposalSheet.getLastRow() - 1, 0), approved: 0, eligibleWeight: 0 };
  }

  var proposals = loadDaoProposalsForFastVote_(proposalSheet);
  var voterWeights = buildDaoVotingWeightIndex_(ss);
  var memberIndex = buildDaoMemberIndex_(ss);
  var voteResult = collectSingleDaoVote_(ss, voteSheet, lastVoteRow, proposals.index, voterWeights, memberIndex);
  if (!voteResult.proposalId) {
    return { proposals: proposals.rows.length, approved: 0, eligibleWeight: voterWeights.totalWeight };
  }

  var summary = summarizeDaoVotesForProposal_(ss, voteSheet, voteResult.proposalId, voterWeights, memberIndex);
  var proposal = proposals.index[normalizeProposalKey_(voteResult.proposalId)];
  if (proposal) {
    updateSingleDaoProposalAggregate_(proposalSheet, proposal.row, summary, voterWeights.totalWeight);
  }

  return {
    proposals: proposals.rows.length,
    approved: proposal && isDaoProposalStatusApproved_(getDaoProposalStatus_(proposalSheet, proposal.row)) ? 1 : 0,
    eligibleWeight: voterWeights.totalWeight,
  };
}

function collectSingleDaoVote_(ss, voteSheet, row, proposalIndex, voterWeights, memberIndex) {
  var processedCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteProcessedHeader);
  var participationCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteParticipationProcessedHeader);
  var canonicalCol = ensureNamedColumn_(voteSheet, '名寄せ済み氏名');
  var weightCol = ensureNamedColumn_(voteSheet, '投票者重み');
  var memoCol = ensureNamedColumn_(voteSheet, '処理メモ');
  var headers = voteSheet.getRange(1, 1, 1, voteSheet.getLastColumn()).getValues()[0];
  var rowValues = voteSheet.getRange(row, 1, 1, voteSheet.getLastColumn()).getValues()[0];
  var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp', '記録日時', '送信日時']);
  var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
  var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  var vote = normalizeDaoVote_(getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']));
  var proposal = proposalIndex[normalizeProposalKey_(parseDaoProposalSelectionId_(proposalInput))]
    || proposalIndex[normalizeProposalKey_(proposalInput)];
  var canonical = '';
  var weight = '';
  var processed = false;
  var memo = '';

  if (!proposal) {
    memo = '提案IDまたは提案タイトルが見つかりません';
  } else {
    var voterMatch = matchDaoMember_(voter, memberIndex);
    if (voterMatch.status !== 'matched') {
      memo = '投票者名を名寄せできません: ' + voterMatch.reason;
    } else if (!vote) {
      memo = '賛成または反対を判定できません';
    } else {
      canonical = voterMatch.canonical;
      weight = voterWeights.byName[canonical] || 0;
      processed = true;
      memo = weight > 0 ? '集計対象' : '投票権重み0のため集計外';
    }
  }

  voteSheet.getRange(row, canonicalCol).setValue(canonical);
  voteSheet.getRange(row, weightCol).setValue(weight);
  voteSheet.getRange(row, processedCol).setValue(processed);
  voteSheet.getRange(row, memoCol).setValue(memo);
  if (processed && weight > 0 && !isTruthy_(voteSheet.getRange(row, participationCol).getValue())) {
    var didAward = awardDaoVoteParticipationPoint_(ss, voteSheet, row, proposal, canonical, timestamp);
    voteSheet.getRange(row, participationCol).setValue(didAward || hasDaoVoteParticipationPoint_(ss, proposal, canonical));
  }
  return {
    proposalId: proposal ? proposal.id : '',
    canonical: canonical,
    vote: vote,
    weight: weight,
  };
}

function summarizeDaoVotesForProposal_(ss, voteSheet, proposalId, voterWeights, memberIndex) {
  var lastRow = voteSheet.getLastRow();
  var latest = {};
  if (lastRow < 2) return { yesWeight: 0, noWeight: 0, count: 0 };
  var values = voteSheet.getRange(1, 1, lastRow, voteSheet.getLastColumn()).getValues();
  var headers = values[0];
  for (var i = 1; i < values.length; i += 1) {
    var rowValues = values[i];
    var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
    var id = parseDaoProposalSelectionId_(proposalInput) || proposalInput;
    if (normalizeProposalKey_(id) !== normalizeProposalKey_(proposalId)) continue;
    var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
    var vote = normalizeDaoVote_(getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']));
    if (!vote) continue;
    var voterMatch = matchDaoMember_(voter, memberIndex);
    if (voterMatch.status !== 'matched') continue;
    var weight = voterWeights.byName[voterMatch.canonical] || 0;
    if (weight <= 0) continue;
    latest[voterMatch.canonical] = { vote: vote, weight: weight };
  }
  var summary = { yesWeight: 0, noWeight: 0, count: 0 };
  Object.keys(latest).forEach(function(name) {
    var item = latest[name];
    if (item.vote === 'yes') summary.yesWeight += item.weight;
    if (item.vote === 'no') summary.noWeight += item.weight;
    summary.count += 1;
  });
  return summary;
}

function updateSingleDaoProposalAggregate_(proposalSheet, row, summary, eligibleTotal) {
  var yesRate = eligibleTotal > 0 ? summary.yesWeight / eligibleTotal : 0;
  var approved = eligibleTotal > 0 && summary.yesWeight > eligibleTotal * DAO_CONFIG.proposalApprovalThreshold;
  var status = getDaoProposalStatus_(proposalSheet, row);
  if (approved && !isDaoProposalStatusApproved_(status)) {
    setDaoProposalStatus_(proposalSheet, row, '承認');
  } else if (!status) {
    setDaoProposalStatus_(proposalSheet, row, '投票中');
  }
  setDaoProposalAggregate_(proposalSheet, row, summary, eligibleTotal, yesRate, approved);
}

function loadDaoProposalsForFastVote_(sheet) {
  ensureProposalSystemColumns_(sheet);
  var lastRow = sheet.getLastRow();
  var result = { rows: [], index: {} };
  if (lastRow < 2) return result;
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0];
  var idCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalIdHeader]);
  var titleCol = findDaoHeaderIndex_(headers, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
  var statusCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalStatusHeader]);
  for (var i = 1; i < values.length; i += 1) {
    var rowNumber = i + 1;
    var id = idCol >= 0 ? normalizeMemberName_(values[i][idCol]) : '';
    var title = titleCol >= 0 ? normalizeMemberName_(values[i][titleCol]) : '';
    var status = statusCol >= 0 ? normalizeMemberName_(values[i][statusCol]) : '';
    if (!id) continue;
    var item = { id: id, title: title, status: status, row: rowNumber };
    result.rows.push(item);
    result.index[normalizeProposalKey_(id)] = item;
    if (title) result.index[normalizeProposalKey_(title)] = item;
  }
  return result;
}

function createDaoProposalForm_(ss) {
  var props = PropertiesService.getScriptProperties();
  var existingId = props.getProperty('DAO_PROPOSAL_FORM_ID');
  if (existingId) {
    try {
      var existingForm = FormApp.openById(existingId);
      configureDaoProposalForm_(existingForm);
      return existingForm;
    } catch (error) {
      props.deleteProperty('DAO_PROPOSAL_FORM_ID');
    }
  }

  var before = getSheetIds_(ss);
  var form = FormApp.create(DAO_CONFIG.proposalFormTitle);
  configureDaoProposalForm_(form);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, DAO_CONFIG.spreadsheetId);
  renameNewResponseSheet_(ss, before, 'DAO提案フォーム回答');
  return form;
}

function getDaoProposalForm_(ss) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DAO_PROPOSAL_FORM_ID');
  try {
    if (id) return FormApp.openById(id);
  } catch (error) {
    props.deleteProperty('DAO_PROPOSAL_FORM_ID');
  }
  var form = openDaoFormByUrl_(props.getProperty('DAO_PROPOSAL_FORM_URL') || DAO_CONFIG.proposalFormUrl);
  if (form) {
    props.setProperty('DAO_PROPOSAL_FORM_ID', form.getId());
    props.setProperty('DAO_PROPOSAL_FORM_URL', form.getPublishedUrl());
  }
  if (!form && ss) {
    form = getDaoFormFromResponseSheets_(ss, DAO_CONFIG.candidateProposalResponseSheets, 'DAO_PROPOSAL_FORM_ID', 'DAO_PROPOSAL_FORM_URL');
  }
  return form;
}

function configureDaoProposalForm_(form) {
  form.setTitle(DAO_CONFIG.proposalFormTitle);
  form.setDescription(DAO_MEMBER_ONLY_FORM_NOTICE + '\n\nCOCoLa DAOに提案したいことを入力してください。メンバー確認ができた提案だけが提案一覧に公開されます。');
  form.setCollectEmail(false);
  if (form.getItems().length > 0) return;
  form.addTextItem().setTitle('提案タイトル').setRequired(true);
  form.addListItem()
    .setTitle('カテゴリ')
    .setChoiceValues(['イベント', '場づくり', '広報', '学び', '制作', 'その他'])
    .setRequired(true);
  form.addParagraphTextItem().setTitle('理由・詳細').setRequired(true);
  form.addParagraphTextItem().setTitle('必要なもの').setRequired(false);
  form.addTextItem().setTitle('提案者').setRequired(true);
}

function createDaoVoteForm_(ss) {
  var props = PropertiesService.getScriptProperties();
  var existingId = props.getProperty('DAO_VOTE_FORM_ID');
  if (existingId) {
    try {
      var existingForm = FormApp.openById(existingId);
      configureDaoVoteForm_(existingForm, ss);
      return existingForm;
    } catch (error) {
      props.deleteProperty('DAO_VOTE_FORM_ID');
    }
  }

  var before = getSheetIds_(ss);
  var form = FormApp.create(DAO_CONFIG.voteFormTitle);
  configureDaoVoteForm_(form, ss);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, DAO_CONFIG.spreadsheetId);
  renameNewResponseSheet_(ss, before, 'DAO投票フォーム回答');
  return form;
}

function getDaoVoteForm_(ss) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DAO_VOTE_FORM_ID');
  try {
    if (id) return FormApp.openById(id);
  } catch (error) {
    props.deleteProperty('DAO_VOTE_FORM_ID');
  }
  var form = openDaoFormByUrl_(props.getProperty('DAO_VOTE_FORM_URL') || DAO_CONFIG.voteFormUrl);
  if (form) {
    props.setProperty('DAO_VOTE_FORM_ID', form.getId());
    props.setProperty('DAO_VOTE_FORM_URL', form.getPublishedUrl());
  }
  if (!form && ss) {
    form = getDaoFormFromResponseSheets_(ss, DAO_CONFIG.candidateVoteResponseSheets, 'DAO_VOTE_FORM_ID', 'DAO_VOTE_FORM_URL');
  }
  return form;
}

function openDaoFormByUrl_(url) {
  if (!url) return null;
  try {
    return FormApp.openByUrl(url);
  } catch (error) {
    Logger.log('フォームURLから開けませんでした: ' + url + ' / ' + error);
    return null;
  }
}

function getDaoFormFromResponseSheets_(ss, sheetNames, idProperty, urlProperty) {
  var props = PropertiesService.getScriptProperties();
  for (var i = 0; i < sheetNames.length; i += 1) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (!sheet || typeof sheet.getFormUrl !== 'function') continue;
    var formUrl = sheet.getFormUrl();
    if (!formUrl) continue;
    var form = openDaoFormByUrl_(formUrl);
    if (form) {
      props.setProperty(idProperty, form.getId());
      props.setProperty(urlProperty, form.getPublishedUrl());
      return form;
    }
  }
  return null;
}

function configureDaoVoteForm_(form, ss) {
  form.setTitle(DAO_CONFIG.voteFormTitle);
  form.setDescription(DAO_MEMBER_ONLY_FORM_NOTICE + '\n\nCOCoLa DAO提案への投票フォームです。提案を選び、投票者名と賛否を入力してください。有効な投票は、賛成・反対に関わらず1提案につき1ptの投票参加ポイント対象になります。');
  form.setCollectEmail(false);
  clearDaoFormItems_(form);
  var choices = buildDaoVoteProposalChoices_(ss);
  if (choices.length === 0) choices = ['提案がまだありません'];
  form.addListItem()
    .setTitle('投票する提案')
    .setChoiceValues(choices)
    .setRequired(true);
  form.addTextItem().setTitle('投票者').setRequired(true);
  form.addMultipleChoiceItem()
    .setTitle('投票')
    .setChoiceValues(['賛成', '反対'])
    .setRequired(true);
  form.addParagraphTextItem().setTitle('コメント').setRequired(false);
  return choices[0] === '提案がまだありません' ? 0 : choices.length;
}

function clearDaoFormItems_(form) {
  var items = form.getItems();
  for (var i = items.length - 1; i >= 0; i -= 1) {
    form.deleteItem(items[i]);
  }
}

function buildDaoVoteProposalChoices_(ss) {
  var sheet = ensureDaoProposalSheet_(ss);
  var lastRow = sheet.getLastRow();
  var choices = [];
  var seen = {};
  var seenContent = {};
  if (lastRow < 2) return choices;
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0];
  var idCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalIdHeader]);
  var titleCol = findDaoHeaderIndex_(headers, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
  var statusCol = findDaoHeaderIndex_(headers, [DAO_CONFIG.proposalStatusHeader, '状態', '判定', '結果']);
  var publicCol = findDaoHeaderIndex_(headers, ['公開', '公開チェック', '掲載', '掲載可']);
  var categoryCol = findDaoHeaderIndex_(headers, ['カテゴリ']);
  var detailCol = findDaoHeaderIndex_(headers, ['理由・詳細', '理由', '詳細', 'なぜやりたいか']);
  var needsCol = findDaoHeaderIndex_(headers, ['必要なもの', '人・場所・費用など']);
  var proposerCol = findDaoHeaderIndex_(headers, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  for (var i = 1; i < values.length; i += 1) {
    var rowValues = values[i];
    var id = idCol >= 0 ? normalizeMemberName_(rowValues[idCol]) : '';
    var title = titleCol >= 0 ? normalizeMemberName_(rowValues[titleCol]) : '';
    var status = statusCol >= 0 ? normalizeMemberName_(rowValues[statusCol]) : '';
    var contentKey = createDaoProposalContentKey_(
      title,
      proposerCol >= 0 ? rowValues[proposerCol] : '',
      categoryCol >= 0 ? rowValues[categoryCol] : '',
      detailCol >= 0 ? rowValues[detailCol] : '',
      needsCol >= 0 ? rowValues[needsCol] : ''
    );
    var isPublic = publicCol >= 0 ? isTruthy_(rowValues[publicCol]) : true;
    var statusKey = getDaoProposalPageStatusKey_(status || '投票中');
    if (!id || !title || !isPublic || statusKey !== 'voting' || seen[id] || (contentKey && seenContent[contentKey])) continue;
    choices.push(id + ': ' + title);
    seen[id] = true;
    if (contentKey) seenContent[contentKey] = true;
  }
  return choices;
}

function createDaoProposalRowFromResponse_(ss, proposalSheet, targetRow, responseSheetName, responseRow, headers, rowValues, key, memberIndex) {
  var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
  var title = getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
  var proposer = getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  var category = getFirstFromRowValues_(headers, rowValues, ['カテゴリ']);
  var detail = getFirstFromRowValues_(headers, rowValues, ['理由・詳細', '理由', '詳細', 'なぜやりたいか']);
  var needs = getFirstFromRowValues_(headers, rowValues, ['必要なもの', '人・場所・費用など']);
  var proposerMatch = matchDaoMember_(proposer, memberIndex);
  var isMember = proposerMatch.status === 'matched';
  var proposalId = 'P' + Utilities.formatString('%04d', Math.max(targetRow - 1, 1));
  var memo = 'フォーム回答: ' + key;
  if (!isMember) {
    memo += ' / メンバー確認できないため公開・集計外: ' + proposerMatch.reason;
    if (shouldReviewName_(proposerMatch)) {
      appendDaoReview_(ss, responseSheetName, responseRow, title, proposer, 'DAO提案メンバー確認', proposerMatch);
    }
  }

  return {
    contentKey: createDaoProposalContentKey_(title, proposer, category, detail, needs),
    row: [
      proposalId,
      timestamp || new Date(),
      title || '',
      category,
      detail,
      needs,
      isMember ? proposerMatch.canonical : (proposer || ''),
      isMember ? '投票中' : '保留',
      isMember,
      0,
      0,
      0,
      0,
      0,
      '',
      memo,
      false,
      false,
    ],
  };
}

function normalizeDaoProposalResponses_(ss) {
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var existing = buildExistingProposalResponseKeys_(proposalSheet);
  var existingContent = buildExistingDaoProposalContentKeys_(proposalSheet);
  var memberIndex = buildDaoMemberIndex_(ss);
  var newRows = [];
  getDaoProposalResponseSheets_(ss).forEach(function(responseSheet) {
    var lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return;
    var values = responseSheet.getRange(1, 1, lastRow, responseSheet.getLastColumn()).getValues();
    var headers = values[0];
    for (var row = 2; row <= values.length; row += 1) {
      var rowValues = values[row - 1];
      var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
      var title = getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
      var proposer = getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
      var key = createFormResponseKey_(responseSheet.getName(), row, timestamp, title, proposer);
      if (existing[key]) continue;
      var targetRow = proposalSheet.getLastRow() + newRows.length + 1;
      var entry = createDaoProposalRowFromResponse_(ss, proposalSheet, targetRow, responseSheet.getName(), row, headers, rowValues, key, memberIndex);
      if (entry.contentKey && existingContent[entry.contentKey]) {
        existing[key] = true;
        continue;
      }
      newRows.push(entry.row);
      existing[key] = true;
      if (entry.contentKey) existingContent[entry.contentKey] = true;
    }
  });
  if (newRows.length > 0) {
    proposalSheet.getRange(proposalSheet.getLastRow() + 1, 1, newRows.length, DAO_PROPOSAL_HEADERS.length).setValues(newRows);
  }
  saveDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY, existing);
}

function normalizeRecentDaoProposalResponses_(ss, limitPerSheet) {
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var existing = buildExistingProposalResponseKeys_(proposalSheet);
  var existingContent = buildExistingDaoProposalContentKeys_(proposalSheet);
  var memberIndex = buildDaoMemberIndex_(ss);
  var changed = false;
  getDaoProposalResponseSheets_(ss).forEach(function(responseSheet) {
    var lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return;
    var startRow = Math.max(2, lastRow - Math.max(Number(limitPerSheet) || 20, 1) + 1);
    var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    var values = responseSheet.getRange(startRow, 1, lastRow - startRow + 1, responseSheet.getLastColumn()).getValues();
    var newRows = [];
    for (var i = 0; i < values.length; i += 1) {
      var row = startRow + i;
      var rowValues = values[i];
      var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
      var title = getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
      var proposer = getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
      var key = createFormResponseKey_(responseSheet.getName(), row, timestamp, title, proposer);
      if (existing[key]) continue;
      var targetRow = proposalSheet.getLastRow() + newRows.length + 1;
      var entry = createDaoProposalRowFromResponse_(ss, proposalSheet, targetRow, responseSheet.getName(), row, headers, rowValues, key, memberIndex);
      if (entry.contentKey && existingContent[entry.contentKey]) {
        existing[key] = true;
        continue;
      }
      newRows.push(entry.row);
      existing[key] = true;
      if (entry.contentKey) existingContent[entry.contentKey] = true;
    }
    if (newRows.length > 0) {
      proposalSheet.getRange(proposalSheet.getLastRow() + 1, 1, newRows.length, DAO_PROPOSAL_HEADERS.length).setValues(newRows);
      changed = true;
    }
  });
  saveDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY, existing);
  return changed;
}

function normalizeLatestDaoProposalResponse_(ss) {
  var responseSheets = getDaoProposalResponseSheets_(ss);
  if (responseSheets.length === 0) return false;
  var responseSheet = responseSheets[0];
  var lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return false;
  var proposalSheet = ensureDaoProposalSheet_(ss);
  var existing = buildExistingProposalResponseKeys_(proposalSheet);
  var existingContent = buildExistingDaoProposalContentKeys_(proposalSheet);
  var memberIndex = buildDaoMemberIndex_(ss);
  var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var rowValues = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
  var title = getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']);
  var proposer = getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  var key = createFormResponseKey_(responseSheet.getName(), lastRow, timestamp, title, proposer);
  if (existing[key]) return false;
  var targetRow = proposalSheet.getLastRow() + 1;
  var entry = createDaoProposalRowFromResponse_(ss, proposalSheet, targetRow, responseSheet.getName(), lastRow, headers, rowValues, key, memberIndex);
  if (entry.contentKey && existingContent[entry.contentKey]) {
    existing[key] = true;
    saveDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY, existing);
    return false;
  }
  existing[key] = true;
  proposalSheet.getRange(targetRow, 1, 1, DAO_PROPOSAL_HEADERS.length).setValues([entry.row]);
  saveDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY, existing);
  return true;
}

function normalizeDaoVoteResponses_(ss) {
  var voteSheet = ensureDaoVoteSheet_(ss);
  var existing = buildExistingVoteResponseKeys_(voteSheet);
  var newRows = [];
  getDaoVoteResponseSheets_(ss).forEach(function(responseSheet) {
    var lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return;
    var values = responseSheet.getRange(1, 1, lastRow, responseSheet.getLastColumn()).getValues();
    var headers = values[0];
    for (var row = 2; row <= values.length; row += 1) {
      var rowValues = values[row - 1];
      var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
      var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
      var proposalId = parseDaoProposalSelectionId_(proposalInput) || proposalInput;
      var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
      var vote = getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']);
      var key = createFormResponseKey_(responseSheet.getName(), row, timestamp, proposalId, voter + ':' + vote);
      if (existing[key]) continue;
      newRows.push([
        timestamp || new Date(),
        proposalId || '',
        getFirstFromRowValues_(headers, rowValues, ['提案タイトル', '提案名', 'タイトル']) || parseDaoProposalSelectionTitle_(proposalInput),
        voter || '',
        vote || '',
        getFirstFromRowValues_(headers, rowValues, ['コメント', '理由', 'メモ']),
        '',
        '',
        false,
        false,
        'フォーム回答: ' + key,
      ]);
      existing[key] = true;
    }
  });
  if (newRows.length > 0) {
    voteSheet.getRange(voteSheet.getLastRow() + 1, 1, newRows.length, DAO_VOTE_HEADERS.length).setValues(newRows);
  }
  saveDaoProcessedResponseKeys_(DAO_VOTE_RESPONSE_KEYS_PROPERTY, existing);
}

function normalizeRecentDaoVoteResponses_(ss, limitPerSheet) {
  var voteSheet = ensureDaoVoteSheet_(ss);
  var existing = buildExistingVoteResponseKeys_(voteSheet);
  var changed = false;
  getDaoVoteResponseSheets_(ss).forEach(function(responseSheet) {
    var lastRow = responseSheet.getLastRow();
    if (lastRow < 2) return;
    var startRow = Math.max(2, lastRow - Math.max(Number(limitPerSheet) || 40, 1) + 1);
    var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
    var values = responseSheet.getRange(startRow, 1, lastRow - startRow + 1, responseSheet.getLastColumn()).getValues();
    var newRows = [];
    for (var i = 0; i < values.length; i += 1) {
      var row = startRow + i;
      var rowValues = values[i];
      var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
      var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
      var proposalId = parseDaoProposalSelectionId_(proposalInput) || proposalInput;
      var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
      var vote = getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']);
      var key = createFormResponseKey_(responseSheet.getName(), row, timestamp, proposalId, voter + ':' + vote);
      if (existing[key]) continue;
      newRows.push([
        timestamp || new Date(),
        proposalId || '',
        getFirstFromRowValues_(headers, rowValues, ['提案タイトル', '提案名', 'タイトル']) || parseDaoProposalSelectionTitle_(proposalInput),
        voter || '',
        vote || '',
        getFirstFromRowValues_(headers, rowValues, ['コメント', '理由', 'メモ']),
        '',
        '',
        false,
        false,
        'フォーム回答: ' + key,
      ]);
      existing[key] = true;
    }
    if (newRows.length > 0) {
      voteSheet.getRange(voteSheet.getLastRow() + 1, 1, newRows.length, DAO_VOTE_HEADERS.length).setValues(newRows);
      changed = true;
    }
  });
  saveDaoProcessedResponseKeys_(DAO_VOTE_RESPONSE_KEYS_PROPERTY, existing);
  return changed;
}

function normalizeLatestDaoVoteResponse_(ss) {
  var responseSheets = getDaoVoteResponseSheets_(ss);
  if (responseSheets.length === 0) return false;
  var responseSheet = responseSheets[0];
  var lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return false;
  var voteSheet = ensureDaoVoteSheet_(ss);
  var existing = buildExistingVoteResponseKeys_(voteSheet);
  var headers = responseSheet.getRange(1, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var rowValues = responseSheet.getRange(lastRow, 1, 1, responseSheet.getLastColumn()).getValues()[0];
  var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp']);
  var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
  var proposalId = parseDaoProposalSelectionId_(proposalInput) || proposalInput;
  var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
  var vote = getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']);
  var key = createFormResponseKey_(responseSheet.getName(), lastRow, timestamp, proposalId, voter + ':' + vote);
  if (existing[key]) return false;
  existing[key] = true;
  voteSheet.getRange(voteSheet.getLastRow() + 1, 1, 1, DAO_VOTE_HEADERS.length).setValues([[
    timestamp || new Date(),
    proposalId || '',
    getFirstFromRowValues_(headers, rowValues, ['提案タイトル', '提案名', 'タイトル']) || parseDaoProposalSelectionTitle_(proposalInput),
    voter || '',
    vote || '',
    getFirstFromRowValues_(headers, rowValues, ['コメント', '理由', 'メモ']),
    '',
    '',
    false,
    false,
    'フォーム回答: ' + key,
  ]]);
  saveDaoProcessedResponseKeys_(DAO_VOTE_RESPONSE_KEYS_PROPERTY, existing);
  return true;
}

function buildExistingProposalResponseKeys_(sheet) {
  var keys = loadDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  var memoCol = ensureNamedColumn_(sheet, '処理メモ');
  var values = sheet.getRange(2, memoCol, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    var text = String(row[0] || '');
    if (text.indexOf('フォーム回答: ') === 0) keys[text.substring('フォーム回答: '.length)] = true;
  });
  saveDaoProcessedResponseKeys_(DAO_PROPOSAL_RESPONSE_KEYS_PROPERTY, keys);
  return keys;
}

function buildExistingVoteResponseKeys_(sheet) {
  var keys = loadDaoProcessedResponseKeys_(DAO_VOTE_RESPONSE_KEYS_PROPERTY);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  var memoCol = ensureNamedColumn_(sheet, '処理メモ');
  var values = sheet.getRange(2, memoCol, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    var text = String(row[0] || '');
    if (text.indexOf('フォーム回答: ') === 0) keys[text.substring('フォーム回答: '.length)] = true;
  });
  saveDaoProcessedResponseKeys_(DAO_VOTE_RESPONSE_KEYS_PROPERTY, keys);
  return keys;
}

function buildExistingDaoProposalContentKeys_(sheet) {
  var keys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0];
  for (var i = 1; i < values.length; i += 1) {
    var rowValues = values[i];
    var isPublic = isTruthy_(getFirstFromRowValues_(headers, rowValues, ['公開', '公開チェック', '掲載', '掲載可']));
    if (!isPublic) continue;
    var key = createDaoProposalContentKey_(
      getFirstFromRowValues_(headers, rowValues, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']),
      getFirstFromRowValues_(headers, rowValues, ['提案者', '提案者名', 'メンバー名', '氏名', '名前', 'ニックネーム']),
      getFirstFromRowValues_(headers, rowValues, ['カテゴリ']),
      getFirstFromRowValues_(headers, rowValues, ['理由・詳細', '理由', '詳細', 'なぜやりたいか']),
      getFirstFromRowValues_(headers, rowValues, ['必要なもの', '人・場所・費用など'])
    );
    if (key) keys[key] = true;
  }
  return keys;
}

function loadDaoProcessedResponseKeys_(propertyName) {
  var text = PropertiesService.getScriptProperties().getProperty(propertyName) || '';
  var keys = {};
  text.split('\n').forEach(function(key) {
    key = String(key || '').trim();
    if (key) keys[key] = true;
  });
  return keys;
}

function saveDaoProcessedResponseKeys_(propertyName, keys) {
  var values = Object.keys(keys || {}).filter(function(key) { return Boolean(key); }).sort();
  var text = values.join('\n');
  // Script properties have a size limit. Keep the newest-looking tail if old test history gets large.
  if (text.length > 8000) {
    values = values.slice(Math.max(values.length - 120, 0));
    text = values.join('\n');
  }
  PropertiesService.getScriptProperties().setProperty(propertyName, text);
}

function buildDaoProposalIndex_(proposalSheets) {
  var index = {};
  proposalSheets.forEach(function(sheet) {
    ensureProposalSystemColumns_(sheet);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    for (var row = 2; row <= lastRow; row += 1) {
      var proposalId = ensureDaoProposalId_(sheet, row);
      var title = getDaoProposalTitle_(sheet, row);
      if (proposalId) {
        index[normalizeProposalKey_(proposalId)] = {
          id: proposalId,
          sheet: sheet,
          row: row,
          title: title,
        };
      }
      if (title) {
        index[normalizeProposalKey_(title)] = {
          id: proposalId,
          sheet: sheet,
          row: row,
          title: title,
        };
      }
    }
  });
  return index;
}

function collectDaoVotes_(ss, voteSheet, proposalIndex, voterWeights, memberIndex) {
  var lastRow = voteSheet.getLastRow();
  var summary = {};
  if (lastRow < 2) return summary;
  var processedCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteProcessedHeader);
  var participationCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteParticipationProcessedHeader);
  var canonicalCol = ensureNamedColumn_(voteSheet, '名寄せ済み氏名');
  var weightCol = ensureNamedColumn_(voteSheet, '投票者重み');
  var memoCol = ensureNamedColumn_(voteSheet, '処理メモ');
  var latestVotes = {};

  for (var row = 2; row <= lastRow; row += 1) {
    var rowObject = getDaoRowObject_(voteSheet, row);
    var timestamp = getFirstByAliases_(rowObject, ['タイムスタンプ', 'Timestamp', '記録日時', '送信日時']);
    var proposalInput = getFirstByAliases_(rowObject, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
    var voter = getFirstByAliases_(rowObject, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
    var vote = normalizeDaoVote_(getFirstByAliases_(rowObject, ['投票', '賛否', '意思表示', '回答']));
    var proposal = proposalIndex[normalizeProposalKey_(parseDaoProposalSelectionId_(proposalInput))]
      || proposalIndex[normalizeProposalKey_(proposalInput)];
    if (!proposal) {
      voteSheet.getRange(row, memoCol).setValue('提案IDまたは提案タイトルが見つかりません');
      continue;
    }
    var voterMatch = matchDaoMember_(voter, memberIndex);
    if (voterMatch.status !== 'matched') {
      if (shouldReviewName_(voterMatch)) {
        appendDaoReview_(ss, voteSheet.getName(), row, proposal.title || proposal.id, voter, '提案投票', voterMatch);
      }
      voteSheet.getRange(row, memoCol).setValue('投票者名を名寄せできません: ' + voterMatch.reason);
      continue;
    }
    if (!vote) {
      voteSheet.getRange(row, memoCol).setValue('賛成または反対を判定できません');
      continue;
    }

    var canonical = voterMatch.canonical;
    var weight = voterWeights.byName[canonical] || 0;
    voteSheet.getRange(row, canonicalCol).setValue(canonical);
    voteSheet.getRange(row, weightCol).setValue(weight);
    voteSheet.getRange(row, processedCol).setValue(true);
    voteSheet.getRange(row, memoCol).setValue(weight > 0 ? '集計対象' : '投票権重み0のため集計外');
    if (weight <= 0) continue;
    if (!isTruthy_(voteSheet.getRange(row, participationCol).getValue())) {
      var didAward = awardDaoVoteParticipationPoint_(ss, voteSheet, row, proposal, canonical, timestamp);
      voteSheet.getRange(row, participationCol).setValue(didAward || hasDaoVoteParticipationPoint_(ss, proposal, canonical));
    }

    latestVotes[proposal.id] = latestVotes[proposal.id] || {};
    latestVotes[proposal.id][canonical] = {
      vote: vote,
      weight: weight,
    };
  }

  Object.keys(latestVotes).forEach(function(proposalId) {
    summary[proposalId] = { yesWeight: 0, noWeight: 0, count: 0 };
    Object.keys(latestVotes[proposalId]).forEach(function(name) {
      var item = latestVotes[proposalId][name];
      if (item.vote === 'yes') summary[proposalId].yesWeight += item.weight;
      if (item.vote === 'no') summary[proposalId].noWeight += item.weight;
      summary[proposalId].count += 1;
    });
  });
  return summary;
}

function collectDaoVotesFast_(ss, voteSheet, proposalIndex, voterWeights, memberIndex) {
  var lastRow = voteSheet.getLastRow();
  var summary = {};
  if (lastRow < 2) return summary;
  var processedCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteProcessedHeader);
  var participationCol = ensureNamedCheckboxColumn_(voteSheet, DAO_CONFIG.voteParticipationProcessedHeader);
  var canonicalCol = ensureNamedColumn_(voteSheet, '名寄せ済み氏名');
  var weightCol = ensureNamedColumn_(voteSheet, '投票者重み');
  var memoCol = ensureNamedColumn_(voteSheet, '処理メモ');
  var lastColumn = voteSheet.getLastColumn();
  var values = voteSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var headers = values[0];
  var updates = {};
  var latestVotes = {};

  for (var i = 1; i < values.length; i += 1) {
    var row = i + 1;
    var rowValues = values[i];
    var timestamp = getFirstFromRowValues_(headers, rowValues, ['タイムスタンプ', 'Timestamp', '記録日時', '送信日時']);
    var proposalInput = getFirstFromRowValues_(headers, rowValues, ['投票する提案', '提案ID', '提案タイトル', '提案名', 'タイトル']);
    var voter = getFirstFromRowValues_(headers, rowValues, ['投票者', '投票者名', 'メンバー名', '氏名', '名前', 'ニックネーム']);
    var vote = normalizeDaoVote_(getFirstFromRowValues_(headers, rowValues, ['投票', '賛否', '意思表示', '回答']));
    var proposal = proposalIndex[normalizeProposalKey_(parseDaoProposalSelectionId_(proposalInput))]
      || proposalIndex[normalizeProposalKey_(proposalInput)];
    var canonical = '';
    var weight = '';
    var processed = false;
    var memo = '';
    var participationProcessed = isTruthy_(rowValues[participationCol - 1]);

    if (!proposal) {
      memo = '提案IDまたは提案タイトルが見つかりません';
    } else {
      var voterMatch = matchDaoMember_(voter, memberIndex);
      if (voterMatch.status !== 'matched') {
        memo = '投票者名を名寄せできません: ' + voterMatch.reason;
      } else if (!vote) {
        memo = '賛成または反対を判定できません';
      } else {
        canonical = voterMatch.canonical;
        weight = voterWeights.byName[canonical] || 0;
        processed = true;
        memo = weight > 0 ? '集計対象' : '投票権重み0のため集計外';
        if (weight > 0) {
          if (!isTruthy_(rowValues[participationCol - 1])) {
            var didAward = awardDaoVoteParticipationPoint_(ss, voteSheet, row, proposal, canonical, timestamp);
            participationProcessed = didAward || hasDaoVoteParticipationPoint_(ss, proposal, canonical);
          }
          latestVotes[proposal.id] = latestVotes[proposal.id] || {};
          latestVotes[proposal.id][canonical] = {
            vote: vote,
            weight: weight,
          };
        }
      }
    }

    updates[row] = {
      canonical: canonical,
      weight: weight,
      processed: processed,
      participationProcessed: participationProcessed,
      memo: memo,
    };
  }

  Object.keys(updates).forEach(function(rowText) {
    var row = Number(rowText);
    var item = updates[rowText];
    voteSheet.getRange(row, canonicalCol).setValue(item.canonical);
    voteSheet.getRange(row, weightCol).setValue(item.weight);
    voteSheet.getRange(row, processedCol).setValue(item.processed);
    voteSheet.getRange(row, participationCol).setValue(item.participationProcessed);
    voteSheet.getRange(row, memoCol).setValue(item.memo);
  });

  Object.keys(latestVotes).forEach(function(proposalId) {
    summary[proposalId] = { yesWeight: 0, noWeight: 0, count: 0 };
    Object.keys(latestVotes[proposalId]).forEach(function(name) {
      var item = latestVotes[proposalId][name];
      if (item.vote === 'yes') summary[proposalId].yesWeight += item.weight;
      if (item.vote === 'no') summary[proposalId].noWeight += item.weight;
      summary[proposalId].count += 1;
    });
  });
  return summary;
}

function buildDaoVotingWeightIndex_(ss) {
  var sheet = ensureDaoMemberSheet_(ss);
  var lastRow = sheet.getLastRow();
  var result = {
    byName: {},
    totalWeight: 0,
  };
  if (lastRow < 2) return result;
  var values = sheet.getRange(2, 1, lastRow - 1, DAO_MEMBER_HEADERS.length).getValues();
  values.forEach(function(row) {
    var name = normalizeMemberName_(row[1]);
    var weight = Number(row[6]) || 0;
    if (!name || weight <= 0) return;
    result.byName[name] = weight;
    result.totalWeight += weight;
  });
  result.totalWeight = Math.round(result.totalWeight * 10) / 10;
  return result;
}

function normalizeDaoVote_(value) {
  var text = String(value || '').trim().toLowerCase();
  if (!text) return '';
  if (['賛成', '承認', 'yes', 'y', 'agree', 'ok', 'true', '1'].indexOf(text) !== -1) return 'yes';
  if (['反対', '否認', 'no', 'n', 'disagree', 'ng', 'false', '0'].indexOf(text) !== -1) return 'no';
  if (text.indexOf('賛成') !== -1 || text.indexOf('承認') !== -1) return 'yes';
  if (text.indexOf('反対') !== -1 || text.indexOf('否認') !== -1) return 'no';
  return '';
}

function normalizeProposalKey_(value) {
  return normalizeHeader_(value).toLowerCase();
}

function parseDaoProposalSelectionId_(value) {
  var text = String(value || '').trim();
  var match = text.match(/^(P\d{4,})\s*[:：]/i);
  return match ? match[1] : '';
}

function parseDaoProposalSelectionTitle_(value) {
  var text = String(value || '').trim();
  var match = text.match(/^P\d{4,}\s*[:：]\s*(.+)$/i);
  return match ? match[1].trim() : '';
}

function ensureDaoProposalId_(sheet, row) {
  var col = ensureNamedColumn_(sheet, DAO_CONFIG.proposalIdHeader);
  var value = normalizeMemberName_(sheet.getRange(row, col).getValue());
  if (value) return value;
  value = 'P' + Utilities.formatString('%04d', Math.max(row - 1, 1));
  sheet.getRange(row, col).setValue(value);
  return value;
}

function getDaoProposalTitle_(sheet, row) {
  var rowObject = getDaoRowObject_(sheet, row);
  return normalizeMemberName_(getFirstByAliases_(rowObject, ['提案タイトル', 'タイトル', '何をしたいか', '提案名']));
}

function getDaoProposalStatus_(sheet, row) {
  var col = ensureNamedColumn_(sheet, DAO_CONFIG.proposalStatusHeader);
  return normalizeMemberName_(sheet.getRange(row, col).getValue());
}

function setDaoProposalStatus_(sheet, row, status) {
  var col = ensureNamedColumn_(sheet, DAO_CONFIG.proposalStatusHeader);
  sheet.getRange(row, col).setValue(status);
}

function setDaoProposalAggregate_(sheet, row, summary, eligibleTotal, yesRate, approved) {
  sheet.getRange(row, ensureNamedColumn_(sheet, '賛成重み')).setValue(summary.yesWeight || 0);
  sheet.getRange(row, ensureNamedColumn_(sheet, '反対重み')).setValue(summary.noWeight || 0);
  sheet.getRange(row, ensureNamedColumn_(sheet, '有効投票権重み')).setValue(eligibleTotal || 0);
  sheet.getRange(row, ensureNamedColumn_(sheet, '賛成率')).setValue(yesRate || 0);
  sheet.getRange(row, ensureNamedColumn_(sheet, '投票数')).setValue(summary.count || 0);
  sheet.getRange(row, ensureNamedColumn_(sheet, '判定日時')).setValue(new Date());
  sheet.getRange(row, ensureNamedColumn_(sheet, '処理メモ')).setValue(approved ? '重み付き過半数により承認' : '投票受付中');
}

function isDaoProposalStatusApproved_(status) {
  var text = normalizeMemberName_(status);
  return ['承認', '採択', '可決', '実行', '実行中', '完了'].indexOf(text) !== -1;
}

function getSheetIds_(ss) {
  var ids = {};
  ss.getSheets().forEach(function(sheet) {
    ids[sheet.getSheetId()] = true;
  });
  return ids;
}

function renameNewResponseSheet_(ss, beforeIds, targetName) {
  SpreadsheetApp.flush();
  var sheets = ss.getSheets();
  for (var i = sheets.length - 1; i >= 0; i -= 1) {
    var sheet = sheets[i];
    if (beforeIds[sheet.getSheetId()]) continue;
    var existing = ss.getSheetByName(targetName);
    if (existing && existing.getSheetId() !== sheet.getSheetId()) {
      targetName = targetName + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    }
    sheet.setName(targetName);
    return sheet;
  }
  return null;
}

function createFormResponseKey_(sheetName, row, timestamp, primary, secondary) {
  return [
    sheetName,
    row,
    String(timestamp || ''),
    String(primary || ''),
    String(secondary || ''),
  ].join(':');
}

function collectDaoEntriesFromRow_(ss, eventSheet, row, memberIndex, writeReview) {
  var rowObject = getDaoRowObject_(eventSheet, row);
  var timestamp = getFirstByAliases_(rowObject, ['タイムスタンプ', 'Timestamp']);
  var eventTitle = getFirstByAliases_(rowObject, ['イベント名', 'タイトル', 'イベントタイトル']);
  var eventDate = getFirstByAliases_(rowObject, ['開催日時', '日時', 'イベント日時', '開催日', '日付']);
  var organizer = getFirstByAliases_(rowObject, ['主催者', '主催', '主催団体', '主催者名']);
  var participants = getParticipantsFromRowObject_(rowObject);
  var submittedDate = parseDaoDate_(timestamp);
  var eventKey = eventSheet.getName() + ':' + row + ':' + String(eventTitle || '');
  var entries = [];

  var organizerMatch = matchDaoMember_(organizer, memberIndex);
  if (DAO_CONFIG.awardOrganizerPoints && organizerMatch.status === 'matched') {
    entries.push({
      name: organizerMatch.canonical,
      type: '主催・開催',
      points: DAO_CONFIG.organizerPoints,
      note: createLedgerEntryKey_(eventKey, '主催・開催', organizerMatch.canonical),
      eventTitle: eventTitle,
      eventDate: eventDate,
      submittedDate: submittedDate,
    });
  } else if (writeReview && DAO_CONFIG.awardOrganizerPoints && shouldReviewName_(organizerMatch)) {
    entries.unresolvedCount = (entries.unresolvedCount || 0) + 1;
    appendDaoReview_(ss, eventSheet.getName(), row, eventTitle, organizer, '主催・開催', organizerMatch);
  }

  var seenParticipants = {};
  participants.forEach(function(name) {
    var participantMatch = matchDaoMember_(name, memberIndex);
    if (participantMatch.status === 'blank') return;
    if (participantMatch.status !== 'matched') {
      entries.unresolvedCount = (entries.unresolvedCount || 0) + 1;
      if (writeReview) {
        appendDaoReview_(ss, eventSheet.getName(), row, eventTitle, name, 'イベント参加', participantMatch);
      }
      return;
    }
    var normalized = participantMatch.canonical;
    if (seenParticipants[normalized]) return;
    seenParticipants[normalized] = true;
    entries.push({
      name: normalized,
      type: 'イベント参加',
      points: DAO_CONFIG.participantPoints,
      note: createLedgerEntryKey_(eventKey, 'イベント参加', normalized),
      eventTitle: eventTitle,
      eventDate: eventDate,
      submittedDate: submittedDate,
    });
  });

  entries.unresolvedCount = entries.unresolvedCount || 0;
  return entries;
}

function shouldReviewName_(match) {
  return match.status === 'ambiguous' || match.status === 'review';
}

function createLedgerEntryKey_(eventKey, activityType, name) {
  return eventKey + ':' + activityType + ':' + name;
}

function ensureDaoMemberSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.memberSheetName) || ss.insertSheet(DAO_CONFIG.memberSheetName);
  ensureHeaders_(sheet, DAO_MEMBER_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDaoLedgerSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.ledgerSheetName) || ss.insertSheet(DAO_CONFIG.ledgerSheetName);
  ensureHeaders_(sheet, DAO_LEDGER_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDaoAliasSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.aliasSheetName) || ss.insertSheet(DAO_CONFIG.aliasSheetName);
  ensureHeaders_(sheet, DAO_ALIAS_HEADERS);
  sheet.setFrozenRows(1);
  seedDaoAliases_(sheet);
  return sheet;
}

function ensureDaoReviewSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.reviewSheetName) || ss.insertSheet(DAO_CONFIG.reviewSheetName);
  ensureHeaders_(sheet, DAO_REVIEW_HEADERS);
  sheet.setFrozenRows(1);
  var maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 10, maxRows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  return sheet;
}

function ensureDaoProposalSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.proposalSheetName) || ss.insertSheet(DAO_CONFIG.proposalSheetName);
  ensureHeaders_(sheet, DAO_PROPOSAL_HEADERS);
  ensureProposalSystemColumns_(sheet);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDaoVoteSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.voteSheetName) || ss.insertSheet(DAO_CONFIG.voteSheetName);
  ensureHeaders_(sheet, DAO_VOTE_HEADERS);
  ensureNamedCheckboxColumn_(sheet, DAO_CONFIG.voteProcessedHeader);
  ensureNamedCheckboxColumn_(sheet, DAO_CONFIG.voteParticipationProcessedHeader);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDaoCitizenLikeSheet_(ss) {
  var sheet = ss.getSheetByName(DAO_CONFIG.citizenLikeSheetName) || ss.insertSheet(DAO_CONFIG.citizenLikeSheetName);
  ensureHeaders_(sheet, DAO_CITIZEN_LIKE_HEADERS);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureProposalSystemColumns_(sheet) {
  ensureNamedColumn_(sheet, DAO_CONFIG.proposalIdHeader);
  ensureNamedColumn_(sheet, DAO_CONFIG.proposalStatusHeader);
  ensureNamedColumn_(sheet, '賛成重み');
  ensureNamedColumn_(sheet, '反対重み');
  ensureNamedColumn_(sheet, '有効投票権重み');
  ensureNamedColumn_(sheet, '賛成率');
  ensureNamedColumn_(sheet, '投票数');
  ensureNamedColumn_(sheet, '判定日時');
  ensureNamedColumn_(sheet, '処理メモ');
  ensureProposalProcessedColumn_(sheet);
  ensureProposalApprovedProcessedColumn_(sheet);
}

function seedDaoAliases_(sheet) {
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 6, DAO_ALIAS_HEADERS.length).setValues([
    ['ゆう', 'ゆうさん', '初期候補: 2026-04-13移行メモより'],
    ['ゆみちゃん', 'ゆみ', '初期候補: 2026-04-13移行メモより'],
    ['Yasu', 'Yasutaka', '初期候補: 2026-04-13移行メモより'],
    ['やすたか', 'Yasutaka', '初期候補: 表記ゆれ対策'],
    ['ひこ', 'ひこさん', '初期候補: 敬称ゆれ対策'],
    ['つるちゃん', 'つるちゃん', '初期候補: 明示登録'],
  ]);
}

function appendDaoAliasIfMissing_(sheet, alias, canonical, note) {
  var aliasKey = normalizeNameKey_(alias);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i += 1) {
      if (normalizeNameKey_(values[i][0]) === aliasKey) return;
    }
  }
  sheet.appendRow([alias, canonical, note || '']);
}

function ensureProcessedColumns_(ss) {
  DAO_CONFIG.candidateEventSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) ensureProcessedColumn_(sheet);
  });
  DAO_CONFIG.candidateProposalSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ensureProposalProcessedColumn_(sheet);
      ensureProposalApprovedProcessedColumn_(sheet);
    }
  });
}

function buildDaoMemberIndex_(daoSs) {
  var members = loadDaoMembersFromSource_();
  var aliasMap = loadDaoAliasMap_(daoSs);
  var exact = {};
  var canonicalSet = {};

  members.forEach(function(member) {
    canonicalSet[member.nickname] = true;
    addMemberKey_(exact, member.nickname, member.nickname, 'ニックネーム');
    addMemberKey_(exact, stripHonorific_(member.nickname), member.nickname, '敬称なしニックネーム');
    if (member.realName) {
      addMemberKey_(exact, member.realName, member.nickname, '本名');
      addMemberKey_(exact, stripSpaces_(member.realName), member.nickname, '空白なし本名');
    }
  });

  Object.keys(aliasMap).forEach(function(alias) {
    if (canonicalSet[aliasMap[alias]]) {
      addMemberKey_(exact, alias, aliasMap[alias], '承認済み別名');
    }
  });

  return {
    members: members,
    exact: exact,
    aliasMap: aliasMap,
  };
}

function loadDaoMembersFromSource_() {
  var ss = SpreadsheetApp.openById(DAO_CONFIG.memberSpreadsheetId);
  var sheet = ss.getSheetByName(DAO_CONFIG.memberSourceSheetName);
  if (!sheet) throw new Error('メンバー紹介シートが見つかりません: ' + DAO_CONFIG.memberSourceSheetName);
  var values = sheet.getDataRange().getValues();
  var members = [];
  var seen = {};

  for (var i = 1; i < values.length; i += 1) {
    var purpose = String(values[i][2] || '').trim();
    var nickname = normalizeMemberName_(values[i][3]);
    var realName = normalizeMemberName_(values[i][4]);
    if (!nickname || purpose === '情報の修正・更新') continue;
    seen[nickname] = {
      nickname: nickname,
      realName: realName,
    };
  }

  Object.keys(seen).forEach(function(nickname) {
    members.push(seen[nickname]);
  });
  return members;
}

function loadDaoAliasMap_(daoSs) {
  var sheet = ensureDaoAliasSheet_(daoSs);
  var lastRow = sheet.getLastRow();
  var map = {};
  if (lastRow < 2) return map;
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(function(row) {
    var alias = normalizeMemberName_(row[0]);
    var canonical = normalizeMemberName_(row[1]);
    if (alias && canonical) map[normalizeNameKey_(alias)] = canonical;
  });
  return map;
}

function addMemberKey_(exact, rawName, canonical, source) {
  var key = normalizeNameKey_(rawName);
  if (!key) return;
  if (!exact[key]) exact[key] = [];
  exact[key].push({
    canonical: canonical,
    source: source,
  });
}

function matchDaoMember_(rawName, memberIndex) {
  var name = normalizeMemberName_(rawName);
  if (!name) return { status: 'blank', canonical: '', candidates: [], reason: '空欄' };

  var keys = uniqueStrings_([
    normalizeNameKey_(name),
    normalizeNameKey_(stripHonorific_(name)),
    normalizeNameKey_(stripSpaces_(name)),
  ]);

  var matches = [];
  keys.forEach(function(key) {
    var found = memberIndex.exact[key] || [];
    found.forEach(function(item) {
      matches.push(item);
    });
  });
  matches = uniqueMatches_(matches);

  if (matches.length === 1) {
    return {
      status: 'matched',
      canonical: matches[0].canonical,
      candidates: [matches[0].canonical],
      reason: matches[0].source,
    };
  }
  if (matches.length > 1) {
    return {
      status: 'ambiguous',
      canonical: '',
      candidates: matches.map(function(item) { return item.canonical; }),
      reason: '複数候補',
    };
  }

  var fuzzy = fuzzyDaoCandidates_(name, memberIndex.members);
  if (fuzzy.length > 0) {
    return {
      status: 'review',
      canonical: '',
      candidates: fuzzy,
      reason: '表記ゆれ候補',
    };
  }
  return {
    status: 'not_found',
    canonical: '',
    candidates: [],
    reason: 'メンバー紹介サイトに一致なし',
  };
}

function fuzzyDaoCandidates_(rawName, members) {
  var key = normalizeNameKey_(stripHonorific_(rawName));
  if (!key || key.length < 2) return [];
  var candidates = [];
  members.forEach(function(member) {
    var nickKey = normalizeNameKey_(stripHonorific_(member.nickname));
    var realKey = normalizeNameKey_(stripSpaces_(member.realName));
    if (!nickKey && !realKey) return;
    if ((nickKey && (nickKey.indexOf(key) !== -1 || key.indexOf(nickKey) !== -1))
      || (realKey && (realKey.indexOf(key) !== -1 || key.indexOf(realKey) !== -1))) {
      candidates.push(member.nickname);
    }
  });
  return uniqueStrings_(candidates).slice(0, 5);
}

function appendDaoReview_(ss, sheetName, row, eventTitle, inputName, activityType, match) {
  var sheet = ensureDaoReviewSheet_(ss);
  var normalizedInputName = normalizeMemberName_(inputName);
  if (hasDaoReview_(sheet, sheetName, row, normalizedInputName, activityType)) return;
  sheet.appendRow([
    new Date(),
    sheetName,
    row,
    eventTitle || '',
    normalizedInputName,
    activityType,
    (match.candidates || []).join(', '),
    match.reason || match.status,
  ]);
}

function hasDaoReview_(sheet, sheetName, row, inputName, activityType) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var values = sheet.getRange(2, 2, lastRow - 1, 5).getValues();
  for (var i = 0; i < values.length; i += 1) {
    if (String(values[i][0] || '') === String(sheetName)
      && Number(values[i][1]) === Number(row)
      && normalizeMemberName_(values[i][3]) === normalizeMemberName_(inputName)
      && String(values[i][4] || '') === String(activityType)) {
      return true;
    }
  }
  return false;
}

function ensureProcessedColumn_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headers.length; i += 1) {
    if (normalizeHeader_(headers[i]) === normalizeHeader_(DAO_CONFIG.processedHeader)) return i + 1;
  }
  var newCol = lastColumn + 1;
  sheet.getRange(1, newCol).setValue(DAO_CONFIG.processedHeader);
  var maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, newCol, maxRows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  return newCol;
}

function ensureProposalProcessedColumn_(sheet) {
  return ensureNamedCheckboxColumn_(sheet, DAO_CONFIG.proposalProcessedHeader);
}

function ensureProposalApprovedProcessedColumn_(sheet) {
  return ensureNamedCheckboxColumn_(sheet, DAO_CONFIG.proposalApprovedProcessedHeader);
}

function ensureNamedCheckboxColumn_(sheet, header) {
  var newCol = ensureNamedColumn_(sheet, header);
  var maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, newCol, maxRows, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  return newCol;
}

function ensureNamedColumn_(sheet, header) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headers.length; i += 1) {
    if (normalizeHeader_(headers[i]) === normalizeHeader_(header)) return i + 1;
  }
  var newCol = lastColumn + 1;
  sheet.getRange(1, newCol).setValue(header);
  return newCol;
}

function ensureHeaders_(sheet, headers) {
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsWrite = false;
  for (var i = 0; i < headers.length; i += 1) {
    if (!current[i]) {
      current[i] = headers[i];
      needsWrite = true;
    }
  }
  if (needsWrite || sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([current]);
  }
}

function getDaoEventSheet_(ss) {
  for (var i = 0; i < DAO_CONFIG.candidateEventSheets.length; i += 1) {
    var sheet = ss.getSheetByName(DAO_CONFIG.candidateEventSheets[i]);
    if (sheet) return sheet;
  }
  return null;
}

function getDaoProposalSheets_(ss) {
  return DAO_CONFIG.candidateProposalSheets
    .map(function(sheetName) { return ss.getSheetByName(sheetName); })
    .filter(function(sheet) { return Boolean(sheet); });
}

function getDaoProposalResponseSheets_(ss) {
  return DAO_CONFIG.candidateProposalResponseSheets
    .map(function(sheetName) { return ss.getSheetByName(sheetName); })
    .filter(function(sheet) { return Boolean(sheet); });
}

function getDaoVoteResponseSheets_(ss) {
  return DAO_CONFIG.candidateVoteResponseSheets
    .map(function(sheetName) { return ss.getSheetByName(sheetName); })
    .filter(function(sheet) { return Boolean(sheet); });
}

function isDaoProposalSheet_(sheet) {
  if (!sheet) return false;
  return DAO_CONFIG.candidateProposalSheets.indexOf(sheet.getName()) !== -1;
}

function getDaoRowObject_(sheet, row) {
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  var object = {};
  headers.forEach(function(header, index) {
    var normalized = normalizeHeader_(header);
    if (!normalized) return;
    if (!object[normalized]) object[normalized] = [];
    object[normalized].push(values[index]);
  });
  return object;
}

function getFirstByAliases_(rowObject, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    var values = rowObject[normalizeHeader_(aliases[i])] || [];
    for (var j = 0; j < values.length; j += 1) {
      if (String(values[j] || '').trim()) return values[j];
    }
  }
  return '';
}

function getFirstFromRowValues_(headers, rowValues, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    var alias = normalizeHeader_(aliases[i]);
    for (var j = 0; j < headers.length; j += 1) {
      if (normalizeHeader_(headers[j]) !== alias) continue;
      if (String(rowValues[j] || '').trim()) return rowValues[j];
    }
  }
  return '';
}

function findDaoHeaderIndex_(headers, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    var alias = normalizeHeader_(aliases[i]);
    for (var j = 0; j < headers.length; j += 1) {
      if (normalizeHeader_(headers[j]) === alias) return j;
    }
  }
  return -1;
}

function getParticipantsFromRowObject_(rowObject) {
  var aliases = ['参加者', '参加者名', '参加メンバー', 'COCoLa参加者', '関わる人'];
  var values = [];
  aliases.forEach(function(alias) {
    var matches = rowObject[normalizeHeader_(alias)] || [];
    matches.forEach(function(value) {
      values = values.concat(splitMemberNames_(value));
    });
  });
  return values;
}

function splitMemberNames_(value) {
  return String(value || '')
    .replace(/[、，]/g, ',')
    .replace(/[／/]/g, ',')
    .replace(/\r?\n/g, ',')
    .split(',')
    .map(normalizeMemberName_)
    .filter(function(name) { return Boolean(name); });
}

function addDaoPoints_(ss, name, points, activityDate) {
  var sheet = ensureDaoMemberSheet_(ss);
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  if (lastRow >= 2) {
    var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < names.length; i += 1) {
      if (normalizeMemberName_(names[i][0]) === name) {
        targetRow = i + 2;
        break;
      }
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, DAO_MEMBER_HEADERS.length).setValues([[
      createMemberId_(targetRow),
      name,
      0,
      0,
      '',
      '',
      0,
    ]]);
  }

  var row = sheet.getRange(targetRow, 1, 1, DAO_MEMBER_HEADERS.length).getValues()[0];
  var currentTotal = Number(row[2]) || 0;
  var currentLastDate = parseDaoDate_(row[4]);
  var nextLastDate = currentLastDate && currentLastDate > activityDate ? currentLastDate : activityDate;
  row[2] = currentTotal + points;
  row[4] = nextLastDate;
  row[5] = getDaoStatus_(nextLastDate);
  row[6] = getVotingWeight_(row[5], row[3], row[2]);
  sheet.getRange(targetRow, 1, 1, DAO_MEMBER_HEADERS.length).setValues([row]);
}

function hasLedgerEventKey_(ledgerSheet, eventKey) {
  var lastRow = ledgerSheet.getLastRow();
  if (lastRow < 2) return false;
  var notes = ledgerSheet.getRange(2, 9, lastRow - 1, 1).getValues();
  for (var i = 0; i < notes.length; i += 1) {
    if (String(notes[i][0] || '') === eventKey) return true;
  }
  return false;
}

function getDaoStatus_(lastActivityDate) {
  if (!lastActivityDate) return '休眠';
  var now = new Date();
  if (lastActivityDate >= addMonths_(now, -DAO_CONFIG.activeMonths)) return 'アクティブ';
  if (lastActivityDate >= addMonths_(now, -DAO_CONFIG.supporterMonths)) return 'サポーター';
  return '休眠';
}

function getVotingWeight_(status, recentPoints, totalPoints) {
  if (status === '休眠') return 0;
  var base = status === 'アクティブ' ? 1 : 0.5;
  var recentBonus = Math.min(Math.floor((Number(recentPoints) || 0) / 10) * 0.1, 0.5);
  var totalBonus = Math.min(Math.floor((Number(totalPoints) || 0) / 50) * 0.1, 0.5);
  return Math.round((base + recentBonus + totalBonus) * 10) / 10;
}

function createMemberId_(row) {
  return 'M' + Utilities.formatString('%04d', Math.max(row - 1, 1));
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[「」『』【】［］\[\]（）()]/g, '')
    .trim();
}

function normalizeMemberName_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeNameKey_(value) {
  return stripSpaces_(stripHonorific_(value))
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(ch) {
      return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0);
    })
    .toLowerCase()
    .replace(/[ーｰ－―]/g, '-')
    .replace(/[・･.．、，,]/g, '')
    .trim();
}

function stripHonorific_(value) {
  return String(value || '')
    .replace(/(さん|ちゃん|くん|君|氏|様)$/g, '')
    .trim();
}

function stripSpaces_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}

function uniqueStrings_(values) {
  var seen = {};
  var result = [];
  values.forEach(function(value) {
    var text = String(value || '').trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    result.push(text);
  });
  return result;
}

function uniqueMatches_(matches) {
  var seen = {};
  var result = [];
  matches.forEach(function(match) {
    var key = match.canonical;
    if (seen[key]) return;
    seen[key] = true;
    result.push(match);
  });
  return result;
}

function parseDaoDate_(value, fallbackYear) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  var text = String(value).trim();
  var direct = new Date(text);
  if (!isNaN(direct.getTime())) return direct;

  var m = text.match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  var reiwa = text.match(/令和\s*(\d{1,2})年\s*(\d{1,2})月\s*(\d{1,2})日/);
  if (reiwa) return new Date(2018 + Number(reiwa[1]), Number(reiwa[2]) - 1, Number(reiwa[3]));

  var noYear = text.match(/(\d{1,2})月\s*(\d{1,2})日/);
  if (noYear && fallbackYear) return new Date(Number(fallbackYear), Number(noYear[1]) - 1, Number(noYear[2]));

  return null;
}

function addMonths_(date, months) {
  var d = new Date(date.getTime());
  d.setMonth(d.getMonth() + months);
  return d;
}

function isTruthy_(value) {
  if (value === true) return true;
  var text = String(value || '').trim().toLowerCase();
  return ['true', 'yes', 'y', '1', '済', '済み'].indexOf(text) !== -1;
}

// ---- GitHub Pages static JSON push ----------------------------------------
// Keeps dao/data.json on the GitHub repo up-to-date so the proposal page
// loads without needing a live GAS fetch (important for Instagram WebView).

var GITHUB_OWNER = 'you0810jmsdf';
var GITHUB_REPO  = 'cocola-site';
var GITHUB_DATA_PATH = 'dao/data.json';
var GITHUB_BRANCH    = 'main';
var GITHUB_PAT_KEY   = 'GITHUB_PAT';

function pushDaoDataToGithub_() {
  var pat = PropertiesService.getScriptProperties().getProperty(GITHUB_PAT_KEY);
  if (!pat) {
    Logger.log('[pushDaoDataToGithub_] GITHUB_PAT not set in Script Properties — skipping.');
    return;
  }

  var jsonStr = getDaoProposalPageDataJson_(true);
  var data;
  try { data = JSON.parse(jsonStr); } catch (e) {
    Logger.log('[pushDaoDataToGithub_] JSON parse error: ' + e);
    return;
  }
  data._generatedMs = Date.now();
  var newContent = JSON.stringify(data);

  var apiBase = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/' + GITHUB_DATA_PATH;
  var headers = {
    'Authorization': 'token ' + pat,
    'Accept': 'application/vnd.github.v3+json',
    'User-Agent': 'GAS-COCoLa'
  };

  var sha = '';
  try {
    var getResp = UrlFetchApp.fetch(apiBase + '?ref=' + GITHUB_BRANCH, {
      headers: headers,
      muteHttpExceptions: true
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha || '';
    }
  } catch (e) {
    Logger.log('[pushDaoDataToGithub_] GET error: ' + e);
  }

  var payload = {
    message: 'Auto-update DAO data [skip ci]',
    content: Utilities.base64Encode(newContent, Utilities.Charset.UTF_8),
    branch: GITHUB_BRANCH
  };
  if (sha) payload.sha = sha;

  try {
    var putResp = UrlFetchApp.fetch(apiBase, {
      method: 'put',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = putResp.getResponseCode();
    if (code === 200 || code === 201) {
      Logger.log('[pushDaoDataToGithub_] OK (' + code + ')');
    } else {
      Logger.log('[pushDaoDataToGithub_] PUT ' + code + ': ' + putResp.getContentText().slice(0, 200));
    }
  } catch (e) {
    Logger.log('[pushDaoDataToGithub_] PUT error: ' + e);
  }
}

function setupDaoStaticJsonTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pushDaoDataToGithub_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('pushDaoDataToGithub_')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('Trigger set: pushDaoDataToGithub_ every 10 minutes.');
}
