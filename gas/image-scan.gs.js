// ============================================
// COCoLa TOPICS Web App
// ============================================

var SHEET_ID = '1Xma1V92uNPTcXmj1cDNzU0jePcxAtTR_4vEF2ZXRgdo';
var SHEET_NAME = 'TOPICS';
var FORM_RESPONSE_SHEET_NAME = 'フォームの回答 1';
var FORM_ID = '1MXr-GdG5Jbu9aQE0xTR3xjWFl5XFkQXdl7asJNChMDQ';
var ADMIN_PASS = 'cocola2026';
var CALENDAR_ID = 'cocola.project@gmail.com'; // COCoLa Googleカレンダー
var ADMIN_NOTIFY_EMAIL = 'cocola.project@gmail.com';
var INZAI_SHEET = '印西イベント'; // 印西市イベント専用シート

// 巡回対象サイト一覧
// ※ まいぷれ印西は全国エリアリンクを誤取得するため一時除外
var INZAI_SOURCES = [
  { name: '印西市公式（今月）',   url: 'https://www.city.inzai.lg.jp/event2/0curr_1.html',        parser: 'inzai_city' },
  { name: '印西市公式（来月）',   url: 'https://www.city.inzai.lg.jp/event2/0next_1.html',        parser: 'inzai_city' },
  { name: '印西市観光協会',       url: 'https://www.city.inzai.lg.jp/soshiki/6-5-2-0-0_2.html',  parser: 'inzai_city' },
  { name: 'いんざいネット',       url: 'https://inzainet.com/wp-json/wp/v2/posts?per_page=50',    parser: 'inzainet'   },
  { name: 'いんざい子育てナビ',  url: 'https://www.city.inzai.lg.jp/kosodatenavi/category/18-2-0-0-0.html', parser: 'kosodatenavi' },
  { name: 'いんざい市民スクール', url: 'https://inzaiec.machikatsu.co.jp/',                                  parser: 'inzaiec',      skipDetailFetch: true },
  { name: '印西市文化ホール',   url: 'https://www.inzai-bunka.jp/event/',                                     parser: 'inzaibunka', skipDetailFetch: true },
  { name: 'アルカサール',       url: 'https://nt-alcazar.com/event-topics/',                                   parser: 'alcazar',    skipDetailFetch: true },
  { name: 'コスモスパレット',   url: 'https://inzai-cosmospalette.jp/wp-json/wp/v2/events?per_page=50&_fields=id,title,link,content', parser: 'cosmospalette', skipDetailFetch: true },
];

var CATEGORY_MAP = {
  'craft':     'ものをつくる',
  'machi':     'まちをつくる',
  'tsunagari': 'つながりをつくる',
  'food':      '食をつくる',
  'self':      '自分をつくる',
  'together':  'ともにつくる'
};

// ============================================
// Web App エントリーポイント
// ============================================
function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode ? e.parameter.mode : '';
  if (['json', 'cancel', 'restore', 'like', 'refreshFormNotices', 'debugRecentProposalResponses'].indexOf(mode) !== -1) {
    return doGetDao_(e);
  }
  if (mode === 'repairTopicsFormSync') {
    if (!e || !e.parameter || e.parameter.pass !== ADMIN_PASS) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var result = repairTopicsFormSync();
    result.ok = true;
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'debugTopicsFormSync') {
    if (!e || !e.parameter || e.parameter.pass !== ADMIN_PASS) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(debugTopicsFormSync_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'moveCocolaMorningCafe') {
    if (!e || !e.parameter || e.parameter.pass !== ADMIN_PASS) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify(moveCocolaMorningCafeIntoVisibleTopics_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (mode === 'fetchInzaiEvents') {
    if (!e || !e.parameter || e.parameter.pass !== ADMIN_PASS) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var before = countInzaiEventRows_();
    weeklyFetchInzaiEvents();
    var after = countInzaiEventRows_();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, before: before, after: after, added: after - before }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (mode === 'inzai_json') {
    return ContentService.createTextOutput(JSON.stringify(getInzaiEventsJson_()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (mode === 'pageview') {
    var p = (e && e.parameter && e.parameter.p) ? e.parameter.p : 'unknown';
    var vid = (e && e.parameter && e.parameter.vid) ? e.parameter.vid : '';
    try { logSiteAccess_(p, vid); } catch(err) {}
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  if (mode === 'record_balance') {
    var amt = parseInt((e && e.parameter && e.parameter.amount) ? e.parameter.amount : '0', 10);
    if (amt > 0) { try { logBalance_(amt); } catch(err) {} }
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'home';
  var cat  = (e && e.parameter && e.parameter.cat) ? e.parameter.cat : '';

  // adminページ以外はアクセスログを記録
  if (page !== 'admin') {
    try { logAccess_(page, cat); } catch(err) {}
  }

  var html;
  if (page === 'category') {
    html = buildCategoryHtml(cat);
  } else if (page === 'inzai') {
    html = buildInzaiEventsHtml();
  } else if (page === 'edit') {
    html = buildEventEditHtml_(e && e.parameter ? e.parameter.key : '');
  } else if (page === 'admin') {
    html = buildAdminHtml_(e);
  } else {
    html = buildHtml();
  }

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// アクセスログ記録
// ============================================================
function logAccess_(page, cat) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var LOG_SHEET = 'アクセスログ';
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['日時', 'ページ', '詳細']]);
    sheet.setFrozenRows(1);
  }
  var label = page === 'category' ? 'category:' + cat : page;
  sheet.appendRow([new Date(), page, label]);

  // ログが5000行を超えたら古い行を削除（先頭100行分）
  var last = sheet.getLastRow();
  if (last > 5001) {
    sheet.deleteRows(2, 100);
  }
}

function logSiteAccess_(page, vid) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var SITE_LOG = 'サイトアクセスログ';
  var sheet = ss.getSheetByName(SITE_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SITE_LOG);
    sheet.getRange(1, 1, 1, 3).setValues([['日時', 'ページ', '訪問者ID']]);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), page, vid || '']);
  var last = sheet.getLastRow();
  if (last > 10001) { sheet.deleteRows(2, 500); }
}

function logBalance_(amount) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('GCloud残高ログ');
  if (!sheet) {
    sheet = ss.insertSheet('GCloud残高ログ');
    sheet.getRange(1, 1, 1, 2).setValues([['日時', '残高（円）']]);
    sheet.setFrozenRows(1);
  }
  // 同日記録済みなら上書き
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var d = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'Asia/Tokyo', 'yyyy/MM/dd') : '';
    if (d === today) {
      sheet.getRange(i + 1, 2).setValue(amount);
      return;
    }
  }
  sheet.appendRow([new Date(), amount]);
}

// ============================================
// ホームページ（イベントカード表示）
// ============================================
function buildHtml() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var events = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var published = row[10]; // K列: 公開チェック
    if (!published) continue;

    events.push({
      timestamp:    row[0],
      datetime:     row[1],
      title:        row[2],
      location:     row[3],
      organizer:    row[4],
      imageUrl:     row[5],
      formUrl:      row[6],
      summary:      row[7],
      category:     row[8],
      participants: row[9],
    });
  }

  events.sort(function(a, b) {
    return new Date(a.datetime) - new Date(b.datetime);
  });

  var cards = events.map(function(ev) {
    return buildCard(ev);
  }).join('');

  if (cards === '') {
    cards = '<p style="text-align:center;color:#999;">現在公開中のイベントはありません</p>';
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>' + getStyle() + '</style>'
    + '</head><body>'
    + '<div class="topics-container">' + cards + '</div>'
    + '</body></html>';
}

function buildCard(ev) {
  var dateStr = formatDate(ev.datetime);
  var imageHtml = '';

  if (ev.imageUrl) {
    var imgSrc = getDriveImageUrl(ev.imageUrl);
    imageHtml = '<div class="card-image"><img src="' + imgSrc + '" alt="' + escapeHtml(ev.title) + '"></div>';
  }

  var formBtn = '';
  if (ev.formUrl) {
    formBtn = '<a class="form-btn" href="' + ev.formUrl + '" target="_blank">申込・詳細はこちら</a>';
  }

  var summaryHtml = ev.summary
    ? '<p class="summary">' + escapeHtml(String(ev.summary)) + '</p>'
    : '';

  var categoryHtml = ev.category
    ? '<span class="category">' + escapeHtml(String(ev.category)) + '</span>'
    : '';

  var participantsHtml = ev.participants
    ? '<p class="participants">参加者: ' + escapeHtml(String(ev.participants)) + '</p>'
    : '';

  return '<div class="card">'
    + imageHtml
    + '<div class="card-body">'
    + categoryHtml
    + '<p class="date">' + dateStr + '</p>'
    + '<h3 class="title">' + escapeHtml(String(ev.title)) + '</h3>'
    + '<p class="meta">📍 ' + escapeHtml(String(ev.location)) + '</p>'
    + '<p class="meta">主催: ' + escapeHtml(String(ev.organizer)) + '</p>'
    + summaryHtml
    + participantsHtml
    + formBtn
    + '</div>'
    + '</div>';
}

// ============================================
// カテゴリページ（活動実績テキスト表示）
// ============================================
function buildCategoryHtml(catKey) {
  var categoryName = CATEGORY_MAP[catKey] || '';
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var events = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var published = row[12]; // M列: カテゴリ掲載
    var category  = String(row[8] || ''); // I列: カテゴリ
    if (!published) continue;
    if (categoryName && category !== categoryName) continue;

    events.push({
      timestamp:    row[0],  // A: タイムスタンプ（ソート用）
      datetime:     row[1],  // B: 開催日時（表示用・テキスト）
      title:        row[2],  // C: イベント名
      organizer:    row[4],  // E: 主催者名
      participants: row[9],  // J: 参加者
      project:      row[11] || '', // L: プロジェクト名
    });
  }

  events.sort(function(a, b) {
    var da = parseEventDate(a.datetime);
    var db = parseEventDate(b.datetime);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });

  // プロジェクト別にグループ化
  var groups = {};
  var groupOrder = [];
  events.forEach(function(ev) {
    var proj = ev.project ? String(ev.project) : '（未分類）';
    if (!groups[proj]) {
      groups[proj] = [];
      groupOrder.push(proj);
    }
    groups[proj].push(ev);
  });

  var content = '';
  if (groupOrder.length === 0) {
    content = '<p class="empty">現在公開中の活動はありません</p>';
  } else {
    groupOrder.forEach(function(proj) {
      content += '<div class="project-section">';
      content += '<h3 class="project-name">【' + escapeHtml(proj) + '】</h3>';
      content += '<ul class="activity-list">';
      groups[proj].forEach(function(ev) {
        var dateStr = formatEventDateDisplay(ev.datetime);
        var participantsStr = ev.participants
          ? '（' + escapeHtml(String(ev.participants)) + '）' : '';
        content += '<li class="activity-item">'
          + '<span class="act-date">' + dateStr + '</span>'
          + '　<span class="act-title">' + escapeHtml(String(ev.title)) + '</span>'
          + '（<span class="act-organizer">' + escapeHtml(String(ev.organizer)) + '</span>）'
          + participantsStr
          + '</li>';
      });
      content += '</ul></div>';
    });
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>' + getCategoryStyle() + '</style>'
    + '</head><body>'
    + '<div class="category-container">' + content + '</div>'
    + '</body></html>';
}

// ============================================
// フォーム送信時に自動実行
// ============================================
function onFormSubmit(e) {
  if (!e || !e.values) {
    syncMissingFormResponsesToTopics();
    return;
  }
  appendTopicFromFormValues_(e.values, { sendEditMail: true, addCalendar: true });
}

function appendTopicFromFormValues_(values, options) {
  options = options || {};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  ensureTopicEditColumns_(topicsSheet);
  // フォームの回答1の列順:
  // values[0]=タイムスタンプ, [1]=開催日時, [2]=イベント名, [3]=場所
  // [4]=主催者名, [5]=COCoLaメンバー(スキップ), [6]=ポスター画像URL
  // [7]=申込・詳細リンク, [8]=イベント概要, [9]=カテゴリ
  // [10]=参加者, [11]=連絡先メール(スキップ), [12]=プロジェクト名

  var newRow = findFirstBlankTopicRow_(topicsSheet);

  var existingRow = findTopicRowBySubmittedValues_(topicsSheet, values);
  if (existingRow > 0) {
    Logger.log('TOPICS反映済みのためスキップ: row=' + existingRow + ', title=' + (values[2] || ''));
    return existingRow;
  }

  topicsSheet.getRange(newRow, 1, 1, 10).setValues([[
    values[0]  || '',  // A: タイムスタンプ
    values[1]  || '',  // B: 開催日時
    values[2]  || '',  // C: イベント名
    values[3]  || '',  // D: 場所
    values[4]  || '',  // E: 主催者名
    values[6]  || '',  // F: ポスター画像URL
    values[7]  || '',  // G: 申込・詳細リンク
    values[8]  || '',  // H: イベント概要
    values[9]  || '',  // I: カテゴリ
    values[10] || ''   // J: 参加者
  ]]);
  topicsSheet.getRange(newRow, 11).insertCheckboxes(); // K: 公開（ホームページ）
  topicsSheet.getRange(newRow, 11).setValue(true);
  topicsSheet.getRange(newRow, 12).setValue(values[12] || ''); // L: プロジェクト名
  topicsSheet.getRange(newRow, 13).insertCheckboxes(); // M: カテゴリ掲載
  var editKey = createEventEditKey_();
  var submitterEmail = String(values[11] || '').trim();
  topicsSheet.getRange(newRow, 15).setValue(editKey); // O: 編集キー
  topicsSheet.getRange(newRow, 16).setValue(submitterEmail); // P: 投稿者メール

  if (options.addCalendar !== false) {
    // ── Googleカレンダーに自動登録 ──
    var calEventId = addToCalendar_(
      values[2] || '',   // イベント名
      values[1] || '',   // 開催日時
      values[3] || '',   // 場所
      values[8] || '',   // イベント概要
      values[7] || ''    // 申込・詳細リンク
    );
    if (calEventId) {
      topicsSheet.getRange(newRow, 14).setValue(calEventId); // N列: カレンダーイベントID
    }
  }

  if (options.sendEditMail !== false) {
    sendEditLinkToSubmitter_(submitterEmail, editKey, values[2] || '');
  }

  Logger.log('TOPICSへ反映: row=' + newRow + ', title=' + (values[2] || ''));
  return newRow;
}

function syncMissingFormResponsesToTopics() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var responseSheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
  if (!responseSheet) throw new Error('フォーム回答シートが見つかりません: ' + FORM_RESPONSE_SHEET_NAME);
  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  ensureTopicEditColumns_(topicsSheet);

  var lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return { checked: 0, added: 0 };

  var values = responseSheet.getRange(2, 1, lastRow - 1, responseSheet.getLastColumn()).getValues();
  var added = 0;
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0] || !row[2]) continue;
    if (findTopicRowBySubmittedValues_(topicsSheet, row) > 0) continue;
    appendTopicFromFormValues_(row, { sendEditMail: false, addCalendar: true });
    added++;
  }
  return { checked: values.length, added: added };
}

function repairTopicsFormSync() {
  setTrigger();
  return syncMissingFormResponsesToTopics();
}

function moveCocolaMorningCafeIntoVisibleTopics_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var matches = findSheetRowsByText_(sheet, 'COCoLa朝カフェ', 3);
  if (matches.length === 0) return { ok: false, error: 'COCoLa朝カフェ not found' };

  var sourceRow = matches[0].row;
  var targetRow = findFirstBlankTopicRow_(sheet);
  if (sourceRow === targetRow) return { ok: true, moved: false, row: sourceRow };

  var lastColumn = Math.max(17, sheet.getLastColumn());
  var values = sheet.getRange(sourceRow, 1, 1, lastColumn).getValues();
  var validations = sheet.getRange(sourceRow, 1, 1, lastColumn).getDataValidations();
  sheet.getRange(targetRow, 1, 1, lastColumn).setValues(values);
  sheet.getRange(targetRow, 1, 1, lastColumn).setDataValidations(validations);
  sheet.getRange(sourceRow, 1, 1, lastColumn).clearContent();
  return { ok: true, moved: true, from: sourceRow, to: targetRow };
}

function debugTopicsFormSync_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  var responseSheet = ss.getSheetByName(FORM_RESPONSE_SHEET_NAME);
  var targetTitle = 'COCoLa朝カフェ';
  var responseMatch = findSheetRowsByText_(responseSheet, targetTitle, 3);
  var topicsMatch = findSheetRowsByText_(topicsSheet, targetTitle, 3);
  var lastRow = topicsSheet.getLastRow();
  var start = Math.max(2, lastRow - 10);
  var tail = lastRow >= 2
    ? topicsSheet.getRange(start, 1, lastRow - start + 1, Math.min(17, topicsSheet.getLastColumn())).getDisplayValues()
    : [];
  return {
    ok: true,
    spreadsheetId: SHEET_ID,
    topicsSheetId: topicsSheet.getSheetId(),
    topicsLastRow: lastRow,
    responseLastRow: responseSheet.getLastRow(),
    responseMatches: responseMatch,
    topicsMatches: topicsMatch,
    topicsTailStartRow: start,
    topicsTail: tail
  };
}

function findSheetRowsByText_(sheet, text, col) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, Math.min(17, sheet.getLastColumn())).getDisplayValues();
  var matches = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][col - 1] || '').indexOf(text) !== -1) {
      matches.push({ row: i + 2, values: values[i] });
    }
  }
  return matches;
}

function findFirstBlankTopicRow_(sheet) {
  var lastColumn = Math.max(17, sheet.getLastColumn());
  var maxRowsToCheck = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, maxRowsToCheck - 1, lastColumn).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (isBlankTopicRow_(values[i])) return i + 2;
  }
  return maxRowsToCheck + 1;
}

function isBlankTopicRow_(row) {
  var dataColumns = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 13, 14, 15, 16];
  for (var i = 0; i < dataColumns.length; i++) {
    var col = dataColumns[i];
    if (String(row[col] || '').trim() !== '') return false;
  }
  return true;
}

function findTopicRowBySubmittedValues_(sheet, values) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var submittedKey = buildTopicSubmissionKey_(values[0], values[1], values[2]);
  if (!submittedKey) return -1;

  var rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (buildTopicSubmissionKey_(rows[i][0], rows[i][1], rows[i][2]) === submittedKey) {
      return i + 2;
    }
  }
  return -1;
}

function buildTopicSubmissionKey_(timestamp, datetime, title) {
  var titleText = String(title || '').trim();
  if (!titleText) return '';
  return [
    normalizeTopicKeyValue_(timestamp),
    normalizeTopicKeyValue_(datetime),
    titleText
  ].join('|');
}

function normalizeTopicKeyValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '').replace(/\s+/g, '').trim();
}

function ensureTopicEditColumns_(sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 17)).getValues()[0];
  var required = {
    15: '編集キー',
    16: '投稿者メール',
    17: '最終修正日時'
  };
  Object.keys(required).forEach(function(col) {
    var c = parseInt(col, 10);
    if (!headers[c - 1]) {
      sheet.getRange(1, c).setValue(required[col]);
    }
  });
}

function createEventEditKey_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function sendEditLinkToSubmitter_(email, editKey, title) {
  if (!email || !editKey) return;
  var editUrl = ScriptApp.getService().getUrl() + '?page=edit&key=' + encodeURIComponent(editKey);
  var subject = '【COCoLa】イベント登録内容の修正リンク';
  var body = [
    'COCoLaのイベント登録ありがとうございます。',
    '',
    '登録内容を修正する場合は、以下のリンクから編集してください。',
    '修正後も公開状態は維持され、管理者へ変更内容が通知されます。',
    '',
    'イベント名: ' + (title || ''),
    '修正リンク: ' + editUrl,
    '',
    'このリンクを知っている人は登録内容を編集できます。必要な方以外には共有しないでください。'
  ].join('\n');
  try {
    MailApp.sendEmail(email, subject, body);
  } catch (err) {
    Logger.log('修正リンク送信エラー: ' + err);
  }
}

function findEventRowByEditKey_(sheet, editKey) {
  if (!editKey) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var keys = sheet.getRange(2, 15, lastRow - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '') === String(editKey)) return i + 2;
  }
  return -1;
}

function getEditableEventByKey_(editKey) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  ensureTopicEditColumns_(sheet);
  var rowIndex = findEventRowByEditKey_(sheet, editKey);
  if (rowIndex < 0) return null;
  var row = sheet.getRange(rowIndex, 1, 1, 17).getValues()[0];
  return {
    key: editKey,
    datetime: String(row[1] || ''),
    title: String(row[2] || ''),
    location: String(row[3] || ''),
    organizer: String(row[4] || ''),
    imageUrl: String(row[5] || ''),
    formUrl: String(row[6] || ''),
    summary: String(row[7] || ''),
    category: String(row[8] || ''),
    participants: String(row[9] || ''),
    project: String(row[11] || '')
  };
}

function updateEventByEditKey(editKey, data) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  ensureTopicEditColumns_(sheet);
  var rowIndex = findEventRowByEditKey_(sheet, editKey);
  if (rowIndex < 0) {
    return { ok: false, message: '修正対象のイベントが見つかりませんでした。' };
  }

  var before = sheet.getRange(rowIndex, 1, 1, 17).getValues()[0];
  var cleaned = {
    datetime: String(data.datetime || '').trim(),
    title: String(data.title || '').trim(),
    location: String(data.location || '').trim(),
    organizer: String(data.organizer || '').trim(),
    imageUrl: String(data.imageUrl || '').trim(),
    formUrl: String(data.formUrl || '').trim(),
    summary: String(data.summary || '').trim(),
    category: String(data.category || '').trim(),
    participants: String(data.participants || '').trim(),
    project: String(data.project || '').trim()
  };

  if (!cleaned.datetime || !cleaned.title || !cleaned.location || !cleaned.organizer) {
    return { ok: false, message: '開催日時、イベント名、場所、主催者名は入力してください。' };
  }

  sheet.getRange(rowIndex, 2, 1, 9).setValues([[
    cleaned.datetime,
    cleaned.title,
    cleaned.location,
    cleaned.organizer,
    cleaned.imageUrl,
    cleaned.formUrl,
    cleaned.summary,
    cleaned.category,
    cleaned.participants
  ]]);
  sheet.getRange(rowIndex, 11).setValue(true); // K: 修正後も公開のまま
  sheet.getRange(rowIndex, 12).setValue(cleaned.project);
  sheet.getRange(rowIndex, 17).setValue(new Date());

  notifyAdminEventEdited_(rowIndex, before, cleaned);
  return { ok: true, message: '修正を保存しました。公開状態は維持されています。' };
}

function notifyAdminEventEdited_(rowIndex, before, after) {
  var fields = [
    ['開催日時', before[1], after.datetime],
    ['イベント名', before[2], after.title],
    ['場所', before[3], after.location],
    ['主催者名', before[4], after.organizer],
    ['ポスター画像URL', before[5], after.imageUrl],
    ['申込・詳細リンク', before[6], after.formUrl],
    ['イベント概要', before[7], after.summary],
    ['カテゴリ', before[8], after.category],
    ['参加者', before[9], after.participants],
    ['プロジェクト名', before[11], after.project]
  ];
  var changed = fields.filter(function(f) {
    return String(f[1] || '') !== String(f[2] || '');
  });
  if (changed.length === 0) return;

  var body = [
    'イベント登録内容が投稿者により修正されました。',
    '',
    '対象行: ' + rowIndex,
    '公開状態: 公開のまま',
    '',
    '変更内容:'
  ];
  changed.forEach(function(f) {
    body.push('--- ' + f[0]);
    body.push('修正前: ' + String(f[1] || ''));
    body.push('修正後: ' + String(f[2] || ''));
  });
  body.push('');
  body.push('スプレッドシート: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit');

  try {
    MailApp.sendEmail(ADMIN_NOTIFY_EMAIL, '【COCoLa】イベント登録内容が修正されました', body.join('\n'));
  } catch (err) {
    Logger.log('管理者通知メール送信エラー: ' + err);
  }
}

function buildEventEditHtml_(editKey) {
  var ev = getEditableEventByKey_(editKey);
  if (!ev) {
    return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
      + '<meta name="viewport" content="width=device-width,initial-scale=1">'
      + '<title>イベント修正</title>'
      + '<style>' + getEventEditStyle_() + '</style></head><body>'
      + '<main class="edit-wrap"><h1>修正リンクが見つかりません</h1>'
      + '<p>URLが間違っているか、修正用のキーが無効です。</p></main></body></html>';
  }
  var evJson = JSON.stringify(ev);
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>イベント登録内容の修正</title>'
    + '<style>' + getEventEditStyle_() + '</style></head><body>'
    + '<main class="edit-wrap">'
    + '<p class="kicker">COCoLa EVENT EDIT</p>'
    + '<h1>イベント登録内容の修正</h1>'
    + '<p class="lead">修正後も公開状態は維持されます。保存すると管理者へ変更内容が通知されます。</p>'
    + '<form id="edit-form">'
    + editField_('datetime', '開催日時', true)
    + editField_('title', 'イベント名', true)
    + editField_('location', '場所', true)
    + editField_('organizer', '主催者名', true)
    + editField_('imageUrl', 'ポスター画像URL', false)
    + editField_('formUrl', '申込・詳細リンク', false)
    + editTextarea_('summary', 'イベント概要')
    + editField_('category', 'カテゴリ', false)
    + editField_('participants', '参加者', false)
    + editField_('project', 'プロジェクト名', false)
    + '<button type="submit">修正を保存する</button>'
    + '<p id="result" class="result"></p>'
    + '</form></main>'
    + '<script>var EV=' + evJson + ';'
    + 'Object.keys(EV).forEach(function(k){var el=document.getElementById(k);if(el)el.value=EV[k]||"";});'
    + 'document.getElementById("edit-form").addEventListener("submit",function(e){e.preventDefault();'
    + 'var d={};["datetime","title","location","organizer","imageUrl","formUrl","summary","category","participants","project"].forEach(function(k){d[k]=document.getElementById(k).value;});'
    + 'var r=document.getElementById("result");r.textContent="保存中です...";'
    + 'google.script.run.withSuccessHandler(function(res){r.textContent=res.message||"";r.className="result "+(res.ok?"ok":"ng");}).withFailureHandler(function(err){r.textContent="保存に失敗しました: "+err.message;r.className="result ng";}).updateEventByEditKey(EV.key,d);'
    + '});</script></body></html>';
}

function editField_(id, label, required) {
  return '<label for="' + id + '">' + label + (required ? '<span>必須</span>' : '') + '</label>'
    + '<input id="' + id + '" name="' + id + '"' + (required ? ' required' : '') + '>';
}

function editTextarea_(id, label) {
  return '<label for="' + id + '">' + label + '</label><textarea id="' + id + '" name="' + id + '" rows="5"></textarea>';
}

function getEventEditStyle_() {
  return [
    'body{margin:0;padding:16px;font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif;color:#333;background:#fff;}',
    '.edit-wrap{max-width:720px;margin:0 auto;}',
    '.kicker{margin:0 0 4px;font-size:12px;color:#999;font-weight:bold;letter-spacing:0.08em;}',
    'h1{margin:0 0 8px;font-size:24px;color:#c0392b;line-height:1.35;}',
    '.lead{margin:0 0 18px;color:#666;font-size:14px;line-height:1.7;}',
    'form{display:grid;gap:10px;}',
    'label{font-size:13px;font-weight:bold;color:#555;}',
    'label span{margin-left:6px;padding:1px 6px;border-radius:4px;background:#c0392b;color:#fff;font-size:10px;}',
    'input,textarea{width:100%;box-sizing:border-box;border:1px solid #ddd;border-radius:4px;padding:9px 10px;font:inherit;font-size:14px;}',
    'textarea{resize:vertical;}',
    'button{margin-top:8px;border:0;border-radius:4px;background:#c0392b;color:#fff;padding:10px 18px;font-weight:bold;font:inherit;cursor:pointer;}',
    'button:hover{background:#a93226;}',
    '.result{min-height:1.5em;font-weight:bold;}',
    '.result.ok{color:#1d9e75;}',
    '.result.ng{color:#c0392b;}'
  ].join('');
}


// ============================================
// カレンダー登録ヘルパー（重複チェック付き）
// ============================================
function addToCalendar_(title, datetimeStr, location, description, url) {
  if (!title || !datetimeStr) {
    Logger.log('タイトルまたは日時が空のためスキップ');
    return null;
  }

  // 重複チェック: タイトル＋日時のキーでPropertiesServiceに記録
  var props = PropertiesService.getScriptProperties();
  var dedupKey = 'cal_' + Utilities.base64Encode(title + '|' + String(datetimeStr));
  if (props.getProperty(dedupKey)) {
    Logger.log('重複のためスキップ: ' + title);
    return null;
  }

  // 日時をパース
  var startDate = parseCalDate_(String(datetimeStr));
  if (!startDate) {
    Logger.log('日時パース失敗: ' + datetimeStr);
    return null;
  }
  var endDate = new Date(startDate.getTime() + 60 * 60 * 1000); // 1時間後を終了時刻に設定

  try {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) {
      Logger.log('カレンダーが見つかりません: ' + CALENDAR_ID);
      return null;
    }
    var desc = description || '';
    if (url) desc += (desc ? '\n\n' : '') + '申込・詳細: ' + url;

    var event = calendar.createEvent(title, startDate, endDate, {
      location: location || '',
      description: desc,
    });

    var eventId = event.getId();
    props.setProperty(dedupKey, eventId); // 重複防止に記録
    Logger.log('カレンダー登録完了: ' + title + ' (' + eventId + ')');
    return eventId;

  } catch(err) {
    Logger.log('カレンダー登録エラー: ' + err);
    return null;
  }
}


// ============================================
// 日時文字列をDateオブジェクトにパース
// ============================================
function parseCalDate_(text) {
  if (!text) return null;
  var s = text.trim();

  // 時刻を抽出するヘルパー（最初の HH:MM を取得）
  function extractTime(str) {
    var m = str.match(/(\d{1,2}):(\d{2})/);
    return m ? { h: parseInt(m[1]), m: parseInt(m[2]) } : { h: 10, m: 0 };
  }

  // 標準フォーマット（2026/03/22 など）
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // YYYY.M.D または YYYY/M/D（時刻付き対応）
  var dotMatch = s.match(/^(\d{4})[.\/](\d{1,2})[.\/](\d{1,2})/);
  if (dotMatch) {
    var t = extractTime(s);
    return new Date(parseInt(dotMatch[1]), parseInt(dotMatch[2]) - 1, parseInt(dotMatch[3]), t.h, t.m, 0);
  }

  // YYYY年M月D日（時刻付き対応）
  var jaMatch = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (jaMatch) {
    var t = extractTime(s);
    return new Date(parseInt(jaMatch[1]), parseInt(jaMatch[2]) - 1, parseInt(jaMatch[3]), t.h, t.m, 0);
  }

  // 令和N年M月D日（時刻付き対応）
  var reiwaMatch = s.match(/令和(\d+)年(\d{1,2})月(\d{1,2})日/);
  if (reiwaMatch) {
    var t = extractTime(s);
    return new Date(2018 + parseInt(reiwaMatch[1]), parseInt(reiwaMatch[2]) - 1, parseInt(reiwaMatch[3]), t.h, t.m, 0);
  }

  // M月D日（年なし・曜日・時刻付き対応）例: 3月22日（日）10:00〜15:00
  var mdMatch = s.match(/(\d{1,2})月(\d{1,2})日/);
  if (mdMatch) {
    var year = new Date().getFullYear();
    var t = extractTime(s);
    return new Date(year, parseInt(mdMatch[1]) - 1, parseInt(mdMatch[2]), t.h, t.m, 0);
  }

  return null;
}

// ============================================
// STEP 4: setTrigger() を一度だけ実行してトリガーを設定する
// ============================================
function setTrigger() {
  // 旧トリガー(onFormSubmit)と新トリガー(onTopicsFormSubmitNormalized)の両方をクリアして
  // 新のみを登録する。両方が存在すると1回の送信で2行追加される。
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'onFormSubmit' || fn === 'onTopicsFormSubmitNormalized') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.newTrigger('onTopicsFormSubmitNormalized')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

// ============================================
// 印西イベントシートから不要行を削除して整理（一度だけ実行）
// NAV_SKIPに該当するタイトルの行を削除し、良いデータだけ残す
// ============================================
function cleanupInzaiSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INZAI_SHEET);
  if (!sheet) { Logger.log('シートが存在しません'); return; }

  var SKIP = /スキップ|お問い合わせ|組織一覧|サイトマップ|プライバシー|ホーム|トップ|メニュー|ログイン|検索|採用|会社概要|アクセス|もっと見る|開催日時|開催状況|開花状況|ご意見|業務報告|シティプロモーション|まっぷる|事業の終了|日常のすぐそば|調達情報|入札|予算|決算|条例|規則|議会|広報|通知|様式|申請|手続|補助金|助成|統計|窓口|証明|住民票|公園|神社|寺院|史跡|名所|観光スポット|道の駅|施設紹介|地図|駐車場/;

  var data = sheet.getDataRange().getValues();
  var removed = 0;

  // 後ろから走査して不要行を削除（行番号のズレを防ぐ）
  for (var i = data.length - 1; i >= 1; i--) {
    var title = String(data[i][1] || '').trim();
    if (!title) { sheet.deleteRow(i + 1); removed++; continue; }
    if (SKIP.test(title)) { sheet.deleteRow(i + 1); removed++; }
  }
  Logger.log('cleanupInzaiSheet 完了: ' + removed + '行削除');
}

// ============================================
// 印西イベントシートのデータをリセット（一度だけ実行）
// ============================================
function resetInzaiSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INZAI_SHEET);
  if (!sheet) { Logger.log('シートが存在しません'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearFormat();
  }
  Logger.log('印西イベントシートをリセットしました');
}

// ============================================
// 既存イベントをGoogleカレンダーに一括登録
// Apps Scriptエディタから手動で一度だけ実行してください
// N列にカレンダーイベントIDがない行のみ処理します（重複防止）
// ============================================
function syncAllToCalendar() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var added = 0;
  var skipped = 0;
  var failed = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var title       = String(row[2] || '').trim();  // C: イベント名
    var datetimeStr = String(row[1] || '').trim();  // B: 開催日時
    var location    = String(row[3] || '').trim();  // D: 場所
    var description = String(row[7] || '').trim();  // H: イベント概要
    var url         = String(row[6] || '').trim();  // G: 申込・詳細リンク
    var calEventId  = String(row[13] || '').trim(); // N: カレンダーイベントID

    // タイトルが空はスキップ
    if (!title) { skipped++; continue; }

    // N列にすでにIDがあればスキップ（登録済み）
    if (calEventId) {
      Logger.log('登録済みスキップ: ' + title);
      skipped++;
      continue;
    }

    // カレンダーに登録
    var eventId = addToCalendar_(title, datetimeStr, location, description, url);
    if (eventId) {
      sheet.getRange(i + 1, 14).setValue(eventId); // N列にイベントIDを記録
      added++;
      Utilities.sleep(200); // API制限対策
    } else {
      failed++;
    }
  }

  Logger.log('=== syncAllToCalendar 完了 ===');
  Logger.log('登録: ' + added + '件 / スキップ: ' + skipped + '件 / 失敗: ' + failed + '件');
}


// ============================================
// TOPICSシートのヘッダーを更新（一度だけ実行）
// ============================================
function updateHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  sheet.getRange(1, 1, 1, 17).setValues([[
    'タイムスタンプ', '開催日時', 'イベント名', '場所', '主催者名',
    'ポスター画像URL', '申込・詳細リンク', 'イベント概要', 'カテゴリ', '参加者', '公開', 'プロジェクト名',
    'カテゴリ掲載', 'カレンダーイベントID', '編集キー', '投稿者メール', '最終修正日時'
  ]]);
}

// ============================================================
// 印西市イベントページ
// ============================================================
function buildInzaiEventsHtml() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var now = new Date();
  var events = [];

  // ── 印西イベントシートから読み込み ──
  var inzaiSheet = ss.getSheetByName(INZAI_SHEET);
  if (inzaiSheet) {
    var data = inzaiSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[6] === false) continue; // G列: 公開=FALSE は非表示
      var title = String(row[1] || '').trim();
      if (!title) continue;
      var rawDate2 = String(row[2] || '');
      // 常に文字列から再パース（保存済みDateの年誤りを修正するため）
      var parsedDate = parseEventDate(rawDate2) || (row[7] instanceof Date ? row[7] : null);
      var endDate = parseEventEndDate_(rawDate2);
      events.push({
        title:       title,
        datetimeStr: formatEventDateDisplay(row[2]),
        location:    String(row[3] || ''),
        url:         String(row[4] || ''),
        description: String(row[5] || ''),
        source:      String(row[8] || ''),
        date:        parsedDate,
        endDate:     endDate,
      });
    }
  }

  // ── TOPICSシートから読み込み（K列=公開チェック済みのもの）──
  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  if (topicsSheet) {
    var tdata = topicsSheet.getDataRange().getValues();
    for (var j = 1; j < tdata.length; j++) {
      var tr = tdata[j];
      if (!tr[10]) continue; // K列: 公開チェックなしはスキップ
      var ttitle = String(tr[2] || '').trim();
      if (!ttitle) continue;
      var tdatetimeStr = formatEventDateDisplay(tr[1]);
      var tdate = parseEventDate(String(tr[1] || ''));
      var tendDate = parseEventEndDate_(String(tr[1] || ''));
      events.push({
        title:       ttitle,
        datetimeStr: tdatetimeStr,
        location:    String(tr[3] || ''),
        url:         String(tr[6] || ''),
        description: String(tr[7] || ''),
        source:      'COCoLa（' + String(tr[4] || '') + '）',
        date:        tdate,
        endDate:     tendDate,
      });
    }
  }

  // タイトル+日付でdedup
  var dedupSeen = {};
  events = events.filter(function(ev) {
    var key = ev.title + '|' + ev.datetimeStr;
    if (dedupSeen[key]) return false;
    dedupSeen[key] = true;
    return true;
  });

  // 日付順ソート（日付なしは末尾）
  events.sort(function(a, b) {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date - b.date;
  });

  // 30日以上前の古いイベントは除外
  var CUTOFF = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  // 不正タイトルフィルタ
  var BAD_TITLE = /^(お問い合わせ|開催日時|開催状況|もっと見る|アクセス|メニュー|ホーム|トップ|サイトマップ|日常のすぐそばに|旅がある|ご意見|業務報告|調達情報|リンク集|English)$/;
  events = events.filter(function(ev) {
    if (BAD_TITLE.test(ev.title.trim())) return false;
    var effectiveEnd = ev.endDate || ev.date;
    if (effectiveEnd && effectiveEnd < CUTOFF) return false;
    return true;
  });

  // イベントデータをJSON化してJSに渡す
  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var evJson = JSON.stringify(events.map(function(ev) {
    var effectiveEnd = ev.endDate || ev.date;
    var tm = ev.datetimeStr ? ev.datetimeStr.match(/(\d{1,2}):(\d{2})/) : null;
    return {
      t: ev.title, d: ev.datetimeStr, l: ev.location,
      u: ev.url, s: ev.source, desc: ev.description,
      ts: ev.date ? ev.date.getTime() : null,
      h: tm ? parseInt(tm[1]) : -1,
      mi: tm ? parseInt(tm[2]) : -1,
      past: !!(effectiveEnd) && effectiveEnd.getTime() < today.getTime()
    };
  }));

  // 巡回対象サイト一覧（上部に表示）
  var sourcesHtml = '';
  INZAI_SOURCES.forEach(function(src) {
    sourcesHtml += '✅ <a href="' + src.url + '" target="_blank" style="color:#c0392b;">' + src.name + '</a>　';
  });
  var sourcesNote = '情報収集元は、定期的な自動WEB巡回で情報を取得している公開サイトです。SNS、チラシ、個人・団体からのお知らせなど、自動巡回対象外の情報は「イベントを登録する」から個別に登録してください。';
  var sourcesSection = '<div style="padding:6px 0 10px;border-bottom:1px solid #eee;margin-bottom:8px;">'
    + '<p style="font-size:11px;color:#aaa;margin:0 0 4px;">📡 情報収集元</p>'
    + '<p style="font-size:12px;color:#666;line-height:1.6;margin:0 0 6px;">' + sourcesNote + '</p>'
    + '<div style="font-size:11px;color:#666;line-height:2;">' + sourcesHtml + '</div></div>';
  var entryFormUrl = 'https://docs.google.com/forms/d/' + FORM_ID + '/viewform';
  var adminSheetUrl = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=' + (inzaiSheet ? inzaiSheet.getSheetId() : 0);
  var entryButton = '<div class="entry-actions">'
    + '<a class="entry-btn" href="' + entryFormUrl + '" target="_blank" rel="noreferrer noopener">イベントを登録する</a>'
    + '<a class="admin-sheet-btn" href="' + adminSheetUrl + '" onclick="return _openAdminSheet(this.href)">管理者用：スプレッドシート</a>'
    + '</div>';
  var pageHeader = '<header class="inzai-page-header">'
    + '<p class="inzai-page-kicker">COCoLa EVENT LIST</p>'
    + '<h1>千葉ニュータウンイベント一覧</h1>'
    + '<p>千葉ニュータウン周辺のイベント情報をまとめています。</p>'
    + '</header>';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>千葉ニュータウンイベント一覧</title>'
    + '<style>' + getInzaiStyle_() + '</style>'
    + '</head><body><div class="inzai-wrap">'
    + pageHeader
    + sourcesSection
    + entryButton
    + '<div class="view-tabs">'
    + '<button id="btn-list" class="vbtn active" onclick="_setView(\'list\')">リスト</button>'
    + '<button id="btn-month" class="vbtn" onclick="_setView(\'month\')">月</button>'
    + '<button id="btn-week" class="vbtn" onclick="_setView(\'week\')">週</button>'
    + '</div>'
    + '<div id="view-area"></div>'
    + '</div>'
    + '<script>' + getInzaiScript_(evJson) + '<\/script>'
    + '</body></html>';
}

// ============================================================
// 印西イベント JSON エンドポイント (?mode=inzai_json)
// ============================================================
function getInzaiEventsJson_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var now = new Date();
  var CUTOFF = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  var events = [];

  var inzaiSheet = ss.getSheetByName(INZAI_SHEET);
  if (inzaiSheet) {
    var data = inzaiSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[6] === false) continue;
      var title = String(row[1] || '').trim();
      if (!title) continue;
      var rawDate2 = String(row[2] || '');
      var hDate = row[7] instanceof Date ? row[7] : null;
      var parsedDate = parseEventDate(rawDate2) || (hDate && hDate >= CUTOFF ? hDate : null);
      var endDate = parseEventEndDate_(rawDate2);
      events.push({ title: title, datetimeStr: formatEventDateDisplay(row[2]), location: String(row[3] || ''), url: String(row[4] || ''), description: String(row[5] || ''), source: String(row[8] || ''), date: parsedDate, endDate: endDate });
    }
  }

  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  if (topicsSheet) {
    var tdata = topicsSheet.getDataRange().getValues();
    for (var j = 1; j < tdata.length; j++) {
      var tr = tdata[j];
      if (!tr[10]) continue;
      var ttitle = String(tr[2] || '').trim();
      if (!ttitle) continue;
      var tdatetimeStr = formatEventDateDisplay(tr[1]);
      var tdate = parseEventDate(String(tr[1] || ''));
      var tendDate = parseEventEndDate_(String(tr[1] || ''));
      events.push({ title: ttitle, datetimeStr: tdatetimeStr, location: String(tr[3] || ''), url: String(tr[6] || ''), description: String(tr[7] || ''), source: 'COCoLa（' + String(tr[4] || '') + '）', date: tdate, endDate: tendDate });
    }
  }

  var dedupSeen = {};
  events = events.filter(function(ev) {
    var key = ev.title + '|' + ev.datetimeStr;
    if (dedupSeen[key]) return false;
    dedupSeen[key] = true;
    return true;
  });

  events.sort(function(a, b) {
    if (!a.date && !b.date) return 0;
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date - b.date;
  });

  var BAD_TITLE = /^(お問い合わせ|開催日時|開催状況|もっと見る|アクセス|メニュー|ホーム|トップ|サイトマップ|日常のすぐそばに|旅がある|ご意見|業務報告|調達情報|リンク集|English)$/;
  events = events.filter(function(ev) {
    if (BAD_TITLE.test(ev.title.trim())) return false;
    var effectiveEnd = ev.endDate || ev.date;
    if (effectiveEnd && effectiveEnd < CUTOFF) return false;
    return true;
  });

  var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return events.map(function(ev) {
    var effectiveEnd = ev.endDate || ev.date;
    var tm = ev.datetimeStr ? ev.datetimeStr.match(/(\d{1,2}):(\d{2})/) : null;
    return {
      t: ev.title, d: ev.datetimeStr, l: ev.location,
      u: ev.url, s: ev.source, desc: ev.description,
      ts: ev.date ? ev.date.getTime() : null,
      h: tm ? parseInt(tm[1]) : -1,
      mi: tm ? parseInt(tm[2]) : -1,
      past: !!(effectiveEnd) && effectiveEnd.getTime() < today.getTime()
    };
  });
}

// ============================================================
// 管理者ステータスページ (?page=admin)
// ============================================================
function buildAdminHtml_(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var now = new Date();
  var jst = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日 HH:mm');
  var excludeVid = (e && e.parameter && e.parameter.exclude) ? e.parameter.exclude : '';

  // ── サイトアクセスログ集計（GitHub Pages静的サイト）──
  var siteLogSheet = ss.getSheetByName('サイトアクセスログ');
  var sitePvToday = 0, sitePv7d = 0, sitePvTotal = 0;
  var siteVvToday = new Set ? new Set() : {};
  var siteVv7d   = new Set ? new Set() : {};
  var siteVvAll  = new Set ? new Set() : {};
  var sitePagePv = {}, sitePageVv = {};
  var siteRecent = [];
  var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
  var d7ago = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  var vidFreq = {}; // 訪問者ID別PV数（自分特定用）
  if (siteLogSheet) {
    var sdata = siteLogSheet.getDataRange().getValues();
    for (var si = 1; si < sdata.length; si++) {
      var sdate = sdata[si][0] instanceof Date ? sdata[si][0] : null;
      var spage = String(sdata[si][1] || '');
      var svid  = String(sdata[si][2] || '');
      if (!sdate) continue;
      if (excludeVid && svid === excludeVid) continue; // 自分を除外
      sitePvTotal++;
      if (svid) {
        siteVvAll[svid] = true;
        vidFreq[svid] = (vidFreq[svid] || 0) + 1;
      }
      sitePagePv[spage] = (sitePagePv[spage] || 0) + 1;
      if (!sitePageVv[spage]) sitePageVv[spage] = {};
      if (svid) sitePageVv[spage][svid] = true;
      var sdStr = Utilities.formatDate(sdate, 'Asia/Tokyo', 'yyyy/MM/dd');
      if (sdStr === todayStr) {
        sitePvToday++;
        if (svid) siteVvToday[svid] = true;
      }
      if (sdate >= d7ago) {
        sitePv7d++;
        if (svid) siteVv7d[svid] = true;
      }
    }
    var sstart = Math.max(1, sdata.length - 10);
    for (var sri = sdata.length - 1; sri >= sstart; sri--) {
      var srd = sdata[sri][0] instanceof Date ? Utilities.formatDate(sdata[sri][0], 'Asia/Tokyo', 'M/d HH:mm') : '';
      siteRecent.push({ time: srd, page: String(sdata[sri][1] || ''), vid: String(sdata[sri][2] || '') });
    }
  }
  function objKeys(o) { return Object.keys(o); }
  var siteVvTodayCount = objKeys(siteVvToday).length;
  var siteVv7dCount    = objKeys(siteVv7d).length;
  var siteVvAllCount   = objKeys(siteVvAll).length;
  // 訪問者ランキング（上位5件）
  var vidRanking = objKeys(vidFreq).sort(function(a,b){ return vidFreq[b]-vidFreq[a]; }).slice(0, 5);

  // ── アクセスログ集計（GASページ）──
  var logSheet = ss.getSheetByName('アクセスログ');
  var accessToday = 0, access7d = 0, accessTotal = 0;
  var pageCount = {};
  var recentRows = [];
  if (logSheet) {
    var ldata = logSheet.getDataRange().getValues();
    for (var li = 1; li < ldata.length; li++) {
      var ldate = ldata[li][0] instanceof Date ? ldata[li][0] : null;
      var lpage = String(ldata[li][1] || '');
      if (!ldate) continue;
      accessTotal++;
      pageCount[lpage] = (pageCount[lpage] || 0) + 1;
      var ldStr = Utilities.formatDate(ldate, 'Asia/Tokyo', 'yyyy/MM/dd');
      if (ldStr === todayStr) accessToday++;
      if (ldate >= d7ago) access7d++;
    }
    var start = Math.max(1, ldata.length - 10);
    for (var ri = ldata.length - 1; ri >= start; ri--) {
      var rd = ldata[ri][0] instanceof Date ? Utilities.formatDate(ldata[ri][0], 'Asia/Tokyo', 'M/d HH:mm') : '';
      recentRows.push({ time: rd, page: String(ldata[ri][1] || ''), detail: String(ldata[ri][2] || '') });
    }
  }

  // ── GCloud残高ログ集計 ──
  var balSheet = ss.getSheetByName('GCloud残高ログ');
  var balRows = [];   // [{date, amount}] 日付昇順
  var balLatest = null, balPrevDay = null;
  var TRIAL_END = new Date('2026-06-16T23:59:59+09:00');
  if (balSheet) {
    var bdata = balSheet.getDataRange().getValues();
    for (var bi = 1; bi < bdata.length; bi++) {
      var bd = bdata[bi][0] instanceof Date ? bdata[bi][0] : null;
      var ba = Number(bdata[bi][1]);
      if (!bd || isNaN(ba)) continue;
      balRows.push({ date: bd, amount: ba });
    }
    balRows.sort(function(a,b){ return a.date - b.date; });
    if (balRows.length > 0) balLatest = balRows[balRows.length - 1];
    if (balRows.length > 1) balPrevDay = balRows[balRows.length - 2];
  }
  // 消費ペース（直近7日の平均）
  var balPacePerDay = null, balDaysLeft = null, balDepletionDate = null;
  if (balRows.length >= 2) {
    var recent7 = balRows.slice(-8);
    var oldest = recent7[0], newest = recent7[recent7.length - 1];
    var daysDiff = (newest.date - oldest.date) / 86400000;
    if (daysDiff > 0) {
      balPacePerDay = (oldest.amount - newest.amount) / daysDiff;
      if (balPacePerDay > 0) {
        balDaysLeft = Math.floor(newest.amount / balPacePerDay);
        var dep = new Date(newest.date.getTime() + balDaysLeft * 86400000);
        balDepletionDate = Utilities.formatDate(dep, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
    }
  }
  // Chart.js用データ（日付ラベル・残高）
  var balLabels = JSON.stringify(balRows.map(function(r){ return Utilities.formatDate(r.date,'Asia/Tokyo','M/d'); }));
  var balData   = JSON.stringify(balRows.map(function(r){ return r.amount; }));
  var gasAppUrl = ScriptApp.getService().getUrl();
  var trialEndLabel = '6/16';

  // ── 印西イベントシート集計 ──
  var inzaiSheet = ss.getSheetByName(INZAI_SHEET);
  var inzaiTotal = 0;
  var inzaiVisible = 0;
  var inzaiLastFetch = '不明';
  var inzaiNoYear = 0;
  if (inzaiSheet) {
    var idata = inzaiSheet.getDataRange().getValues();
    inzaiTotal = Math.max(0, idata.length - 1);
    for (var i = 1; i < idata.length; i++) {
      if (idata[i][6] !== false) inzaiVisible++;
      if (String(idata[i][2]).indexOf('西暦未記載') !== -1) inzaiNoYear++;
    }
    // 最終収集日時: A列（収集日時）の最大値
    var maxDate = null;
    for (var k = 1; k < idata.length; k++) {
      var d = idata[k][0] instanceof Date ? idata[k][0] : null;
      if (d && (!maxDate || d > maxDate)) maxDate = d;
    }
    if (maxDate) inzaiLastFetch = Utilities.formatDate(maxDate, 'Asia/Tokyo', 'yyyy年M月d日 HH:mm');
  }

  // ── TOPICSシート集計 ──
  var topicsSheet = ss.getSheetByName(SHEET_NAME);
  var topicsTotal = 0;
  var topicsVisible = 0;
  if (topicsSheet) {
    var tdata = topicsSheet.getDataRange().getValues();
    topicsTotal = Math.max(0, tdata.length - 1);
    for (var j = 1; j < tdata.length; j++) {
      if (tdata[j][10]) topicsVisible++;
    }
  }

  // ── HTML生成 ──
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>COCoLa システム管理</title>'
    + '<style>'
    + 'body{margin:0;padding:16px;font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif;font-size:14px;color:#333;background:#f5f5f5;}'
    + 'h1{font-size:18px;color:#c0392b;border-bottom:2px solid #c0392b;padding-bottom:6px;margin-bottom:16px;}'
    + 'h2{font-size:14px;color:#555;background:#e8e8e8;padding:6px 10px;margin:20px 0 8px;border-left:3px solid #c0392b;}'
    + '.card{background:#fff;border-radius:8px;padding:14px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,0.1);}'
    + 'table{border-collapse:collapse;width:100%;}'
    + 'td{padding:5px 8px;border-bottom:1px solid #f0f0f0;}'
    + 'td:first-child{color:#888;width:160px;font-size:12px;}'
    + 'td:last-child{font-weight:bold;}'
    + '.badge-ok{background:#27ae60;color:#fff;padding:2px 8px;border-radius:10px;font-size:12px;}'
    + '.badge-warn{background:#e67e22;color:#fff;padding:2px 8px;border-radius:10px;font-size:12px;}'
    + '.badge-info{background:#2980b9;color:#fff;padding:2px 8px;border-radius:10px;font-size:12px;}'
    + 'a{color:#2980b9;}'
    + '.ts{font-size:11px;color:#aaa;text-align:right;margin-top:16px;}'
    + '.source-list{font-size:12px;color:#555;line-height:2;}'
    + '</style>'
    + '</head><body>'
    + '<h1>🛠 COCoLa システム管理ページ</h1>'

    // サイトアクセス統計（GitHub Pages）
    + '<h2>📊 サイトアクセス状況（GitHub Pages）' + (excludeVid ? ' <span style="font-size:11px;font-weight:normal;color:#e67e22;">自分除外中</span>' : '') + '</h2>'
    + '<div class="card">'
    + '<table>'
    + '<tr><th style="text-align:left;padding:4px 8px;color:#888;font-size:12px;font-weight:normal;width:120px">期間</th><th style="text-align:right;padding:4px 8px;color:#888;font-size:12px;font-weight:normal;">PV</th><th style="text-align:right;padding:4px 8px;color:#888;font-size:12px;font-weight:normal;">VV</th></tr>'
    + '<tr><td>今日</td><td style="text-align:right;font-weight:bold">' + sitePvToday + '</td><td style="text-align:right;font-weight:bold">' + siteVvTodayCount + '</td></tr>'
    + '<tr><td>直近7日間</td><td style="text-align:right;font-weight:bold">' + sitePv7d + '</td><td style="text-align:right;font-weight:bold">' + siteVv7dCount + '</td></tr>'
    + '<tr><td>累計</td><td style="text-align:right;font-weight:bold">' + sitePvTotal + '</td><td style="text-align:right;font-weight:bold">' + siteVvAllCount + '</td></tr>'
    + '</table>'
    + '<div style="margin-top:10px;font-size:12px;color:#888;">ページ別（累計）</div>'
    + '<table style="margin-top:4px">'
    + '<tr><th style="text-align:left;padding:4px 8px;color:#888;font-size:11px;font-weight:normal;">ページ</th><th style="text-align:right;padding:4px 8px;color:#888;font-size:11px;font-weight:normal;">PV</th><th style="text-align:right;padding:4px 8px;color:#888;font-size:11px;font-weight:normal;">VV</th></tr>'
    + Object.keys(sitePagePv).sort(function(a,b){ return sitePagePv[b]-sitePagePv[a]; }).map(function(p) {
        var label = p === 'home' ? 'ホーム' : p === 'inzai' ? '印西市イベント' : p === 'dao' ? 'DAO' : p === 'members' ? 'メンバー一覧' : p === 'member-app' ? 'メンバー登録' : p;
        var vvCount = sitePageVv[p] ? Object.keys(sitePageVv[p]).length : 0;
        return '<tr><td>' + label + '</td><td style="text-align:right">' + sitePagePv[p] + '</td><td style="text-align:right">' + vvCount + '</td></tr>';
      }).join('')
    + '</table>'
    + (vidRanking.length > 0
        ? '<div style="margin-top:10px;font-size:12px;color:#888;">訪問者別PV（上位5件）<span style="margin-left:8px;font-size:11px;">※自分のIDを特定して <code>?page=admin&exclude=あなたのID</code> で除外できます</span></div>'
          + '<table style="margin-top:4px">'
          + vidRanking.map(function(vid, idx) {
              var highlight = (vid === excludeVid) ? ' style="background:#fff3cd"' : '';
              return '<tr' + highlight + '><td style="font-size:11px;color:#aaa;width:28px">#' + (idx+1) + '</td><td style="font-family:monospace;font-size:11px">' + vid + '</td><td style="text-align:right">' + vidFreq[vid] + ' PV</td></tr>';
            }).join('')
          + '</table>'
        : '')
    + (siteRecent.length > 0
        ? '<div style="margin-top:10px;font-size:12px;color:#888;">直近アクセス</div>'
          + '<table style="margin-top:4px">'
          + siteRecent.map(function(r) {
              var label = r.page === 'home' ? 'ホーム' : r.page === 'inzai' ? '印西市イベント' : r.page === 'dao' ? 'DAO' : r.page === 'members' ? 'メンバー一覧' : r.page;
              return '<tr><td style="color:#aaa;font-size:11px;width:90px">' + r.time + '</td><td>' + label + '</td><td style="font-family:monospace;font-size:10px;color:#ccc">' + r.vid.slice(0,8) + '</td></tr>';
            }).join('')
          + '</table>'
        : '')
    + '</div>'

    // GASページアクセス統計
    + '<h2>📊 GASページアクセス状況</h2>'
    + '<div class="card"><table>'
    + '<tr><td>今日</td><td>' + accessToday + ' 回</td></tr>'
    + '<tr><td>直近7日間</td><td>' + access7d + ' 回</td></tr>'
    + '<tr><td>累計</td><td>' + accessTotal + ' 回</td></tr>'
    + '</table>'
    + '<div style="margin-top:10px;font-size:12px;color:#888;">ページ別（累計）</div>'
    + '<table style="margin-top:4px">'
    + Object.keys(pageCount).sort(function(a,b){ return pageCount[b]-pageCount[a]; }).map(function(p) {
        var label = p === 'home' ? 'ホーム' : p === 'inzai' ? '印西市イベント' : p === 'category' ? 'カテゴリ' : p;
        return '<tr><td>' + label + '</td><td>' + pageCount[p] + ' 回</td></tr>';
      }).join('')
    + '</table>'
    + (recentRows.length > 0
        ? '<div style="margin-top:10px;font-size:12px;color:#888;">直近アクセス</div>'
          + '<table style="margin-top:4px">'
          + recentRows.map(function(r) {
              var label = r.page === 'home' ? 'ホーム' : r.page === 'inzai' ? '印西市イベント' : r.page === 'category' ? 'カテゴリ(' + r.detail.replace('category:','') + ')' : r.page;
              return '<tr><td style="color:#aaa;font-size:11px;width:90px">' + r.time + '</td><td>' + label + '</td></tr>';
            }).join('')
          + '</table>'
        : '')
    + '</div>'

    // Google Cloud 課金管理
    + '<h2>☁️ Google Cloud 課金管理</h2>'
    + '<div class="card">'
    + '<table>'
    + '<tr><td>現在の残高</td><td>' + (balLatest ? '¥' + balLatest.amount.toLocaleString() + ' <span class="badge-ok">利用可能</span>' : '<span class="badge-warn">未記録</span>') + '</td></tr>'
    + '<tr><td>前回比</td><td>' + (balLatest && balPrevDay ? '▼¥' + (balPrevDay.amount - balLatest.amount).toLocaleString() + '／日' : '—') + '</td></tr>'
    + '<tr><td>平均消費（直近）</td><td>' + (balPacePerDay !== null ? '¥' + Math.round(balPacePerDay).toLocaleString() + '／日' : '—') + '</td></tr>'
    + '<tr><td>推定枯渇日</td><td>' + (balDepletionDate ? balDepletionDate : '—') + '</td></tr>'
    + '<tr><td>無料期限</td><td>2026年6月16日</td></tr>'
    + '<tr><td>Gemini API</td><td><span class="badge-info">gemini-2.5-flash 使用中</span></td></tr>'
    + '</table>'
    // 残高入力フォーム
    + '<div style="margin-top:14px;padding-top:12px;border-top:1px solid #f0f0f0;">'
    + '<div style="font-size:12px;color:#888;margin-bottom:6px;">📝 今日の残高を記録</div>'
    + '<div style="display:flex;gap:8px;align-items:center;">'
    + '<span style="font-size:13px;">¥</span>'
    + '<input id="balInput" type="number" min="0" max="999999" placeholder="例：46953" style="flex:1;padding:6px 8px;border:1px solid #ddd;border-radius:4px;font-size:14px;">'
    + '<button onclick="recordBalance()" style="padding:6px 14px;background:#c0392b;color:#fff;border:none;border-radius:4px;font-size:13px;cursor:pointer;">記録</button>'
    + '</div>'
    + '<div id="balMsg" style="font-size:12px;margin-top:6px;color:#888;"></div>'
    + '</div>'
    // Chart.js グラフ
    + (balRows.length > 0
        ? '<div style="margin-top:14px;padding-top:12px;border-top:1px solid #f0f0f0;">'
          + '<div style="font-size:12px;color:#888;margin-bottom:8px;">📈 残高推移</div>'
          + '<canvas id="balChart" style="width:100%;max-height:200px;"></canvas>'
          + '</div>'
        : '<div style="font-size:12px;color:#aaa;margin-top:10px;">残高を記録するとグラフが表示されます。</div>')
    + '<hr style="border:none;border-top:1px solid #f0f0f0;margin:12px 0;">'
    + '<div style="font-size:12px;color:#888;margin-bottom:6px;">💰 予算アラート設定</div>'
    + '<table>'
    + '<tr><td>月次予算上限</td><td>¥1,000 / 月</td></tr>'
    + '<tr><td>アラートしきい値</td><td>50% ／ 90% ／ 100%</td></tr>'
    + '<tr><td>通知メール</td><td>you0810jmsdf@gmail.com</td></tr>'
    + '<tr><td>予算とアラート</td><td><a href="https://console.cloud.google.com/billing/01F9AA-70146E-8EBB5D/budgets" target="_blank">GCP コンソールで確認 →</a></td></tr>'
    + '<tr><td>クレジット残高</td><td><a href="https://console.cloud.google.com/billing/01F9AA-70146E-8EBB5D/credits" target="_blank">クレジット残高を確認 →</a></td></tr>'
    + '</table></div>'
    // Chart.js + 記録スクリプト
    + '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>'
    + '<script>'
    + 'var GAS_URL="' + gasAppUrl + '";'
    + 'function recordBalance(){'
    +   'var v=parseInt(document.getElementById("balInput").value,10);'
    +   'if(!v||v<=0){document.getElementById("balMsg").textContent="金額を入力してください";return;}'
    +   'document.getElementById("balMsg").textContent="記録中...";'
    +   'fetch(GAS_URL+"?mode=record_balance&amount="+v,{mode:"no-cors",cache:"no-store"})'
    +     '.then(function(){document.getElementById("balMsg").textContent="✅ 記録しました（ページを再読み込みすると反映されます）";})'
    +     '.catch(function(){document.getElementById("balMsg").textContent="記録しました（no-cors）";});'
    + '}'
    + (balRows.length > 0
        ? 'window.addEventListener("load",function(){'
          + 'var ctx=document.getElementById("balChart").getContext("2d");'
          + 'var labels=' + balLabels + ';'
          + 'var data=' + balData + ';'
          + 'new Chart(ctx,{type:"line",data:{labels:labels,datasets:[{label:"残高（円）",data:data,borderColor:"#c0392b",backgroundColor:"rgba(192,57,43,0.08)",tension:0.3,pointRadius:3,fill:true}]},options:{responsive:true,plugins:{legend:{display:false},annotation:{}},scales:{y:{ticks:{callback:function(v){return"¥"+v.toLocaleString();}},min:0},x:{ticks:{maxTicksLimit:10}}}}});'
          + '});'
        : '')
    + '</script>'

    // 広報いんざいPDF解析
    + '<h2>📄 広報いんざいPDF解析</h2>'
    + '<div class="card"><table>'
    + '<tr><td>モデル</td><td>gemini-2.5-flash（v1beta）</td></tr>'
    + '<tr><td>自動実行</td><td>毎月1日・15日 10:00</td></tr>'
    + '<tr><td>今回の抽出結果</td><td>40件抽出 → 38件追加（2026年4月号）</td></tr>'
    + '<tr><td>APIキー</td><td>creators-map プロジェクト</td></tr>'
    + '</table></div>'

    // 印西イベントシート
    + '<h2>📋 印西イベントシート</h2>'
    + '<div class="card"><table>'
    + '<tr><td>総件数</td><td>' + inzaiTotal + '件</td></tr>'
    + '<tr><td>公開中</td><td>' + inzaiVisible + '件</td></tr>'
    + '<tr><td>西暦未記載</td><td>' + (inzaiNoYear > 0 ? '<span class="badge-warn">' + inzaiNoYear + '件 要確認</span>' : '<span class="badge-ok">なし</span>') + '</td></tr>'
    + '<tr><td>最終収集日時</td><td>' + inzaiLastFetch + '</td></tr>'
    + '<tr><td>次回自動収集</td><td>毎週月曜 6:00</td></tr>'
    + '<tr><td>スプレッドシート</td><td><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit" target="_blank">開く →</a></td></tr>'
    + '</table></div>'

    // TOPICSシート
    + '<h2>📝 COCoLa TOPICSシート</h2>'
    + '<div class="card"><table>'
    + '<tr><td>総件数</td><td>' + topicsTotal + '件</td></tr>'
    + '<tr><td>公開中（K列チェック）</td><td>' + topicsVisible + '件</td></tr>'
    + '</table></div>'

    // 巡回対象サイト
    + '<h2>🌐 巡回対象サイト</h2>'
    + '<div class="card"><div class="source-list">';

  INZAI_SOURCES.forEach(function(src) {
    html += '✅ <a href="' + src.url + '" target="_blank">' + src.name + '</a><br>';
  });

  html += '</div></div>'
    + '<p class="ts">最終確認: ' + jst + '</p>'
    + '</body></html>';

  return html;
}

function getInzaiScript_(evJson) {
  var js = [];
  js.push('var EVTS=' + evJson + ';');
  js.push('var ADMIN_PASS=' + JSON.stringify(ADMIN_PASS) + ';');
  js.push('var _now=new Date();');
  js.push('var _view="list";');
  js.push('var _mo=new Date(_now.getFullYear(),_now.getMonth(),1);');
  js.push('var _wk=_monOf(_now);');
  js.push('function _monOf(d){var r=new Date(d.getFullYear(),d.getMonth(),d.getDate());var w=r.getDay();r.setDate(r.getDate()-(w===0?6:w-1));return r;}');
  js.push('function _prevMo(){_mo=new Date(_mo.getFullYear(),_mo.getMonth()-1,1);_render();}');
  js.push('function _nextMo(){_mo=new Date(_mo.getFullYear(),_mo.getMonth()+1,1);_render();}');
  js.push('function _prevWk(){_wk=new Date(_wk.getTime()-7*86400000);_render();}');
  js.push('function _nextWk(){_wk=new Date(_wk.getTime()+7*86400000);_render();}');
  js.push('function _setView(v){_view=v;["list","month","week"].forEach(function(x){var b=document.getElementById("btn-"+x);if(b)b.className="vbtn"+(x===v?" active":"");});_render();}');
  js.push('function _esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}');
  js.push('function _ed(ev){return ev.ts?new Date(ev.ts):null;}');
  js.push('function _sd(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();}');
  js.push('function _card(ev){');
  js.push('  var url=ev.u?`<a href="${_esc(ev.u)}" target="_blank">${_esc(ev.t)}</a>`:_esc(ev.t);');
  js.push('  var s=`<div class="ev-item${ev.past?" past":""}"><p class="ev-date">${_esc(ev.d)}</p><p class="ev-title">${url}</p>`;');
  js.push('  if(ev.l)s+=`<p class="ev-meta">📍 ${_esc(ev.l)}</p>`;');
  js.push('  if(ev.desc)s+=`<p class="ev-desc">${_esc(ev.desc)}</p>`;');
  js.push('  if(ev.s)s+=`<p class="ev-source">情報元: ${_esc(ev.s)}</p>`;');
  js.push('  return s+"</div>";');
  js.push('}');
  js.push('function _renderList(el){');
  js.push('  if(!EVTS.length){el.innerHTML=\'<p class="empty">現在公開中のイベントはありません</p>\';return;}');
  js.push('  var groups={},groupKeys=[];');
  js.push('  EVTS.forEach(function(ev){');
  js.push('    var d=ev.ts?new Date(ev.ts):null;');
  js.push('    var key=d?(d.getFullYear()+"_"+(d.getMonth()+1)):"nodate";');
  js.push('    var label=d?(d.getFullYear()+"年"+(d.getMonth()+1)+"月"):"日付不明";');
  js.push('    if(!groups[key]){groups[key]={label:label,evts:[],year:d?d.getFullYear():9999,month:d?d.getMonth():0};groupKeys.push(key);}');
  js.push('    groups[key].evts.push(ev);');
  js.push('  });');
  js.push('  var nowY=_now.getFullYear(),nowM=_now.getMonth();');
  js.push('  var html="";');
  js.push('  groupKeys.forEach(function(key){');
  js.push('    var g=groups[key];');
  js.push('    var open=(key==="nodate")||(g.year>nowY)||(g.year===nowY&&g.month>=nowM);');
  js.push('    html+=\'<details class="month-group"\'+(open?" open":"")+\'>\';');
  js.push('    html+=\'<summary class="month-summary">\'+g.label+\'<span class="month-count">\'+g.evts.length+\'件</span></summary>\';');
  js.push('    html+=g.evts.map(_card).join("");');
  js.push('    html+="</details>";');
  js.push('  });');
  js.push('  el.innerHTML=html;');
  js.push('}');
  js.push('function _renderMonth(el){');
  js.push('  var y=_mo.getFullYear(),m=_mo.getMonth();');
  js.push('  var first=new Date(y,m,1),last=new Date(y,m+1,0);');
  js.push('  var nav=\'<div class="cal-nav"><button class="nav-btn" onclick="_prevMo()">◀</button><span class="cal-title">\'+y+\'年\'+(m+1)+\'月</span><button class="nav-btn" onclick="_nextMo()">▶</button></div>\';');
  js.push('  var grid=\'<div class="cal-grid">\';');
  js.push('  ["日","月","火","水","木","金","土"].forEach(function(d){grid+=\'<div class="cal-hd">\'+d+\'</div>\';});');
  js.push('  for(var i=0;i<first.getDay();i++)grid+=\'<div class="cal-cell empty"></div>\';');
  js.push('  for(var d=1;d<=last.getDate();d++){');
  js.push('    var dt=new Date(y,m,d);');
  js.push('    var evs=EVTS.filter(function(ev){return _sd(_ed(ev),dt);});');
  js.push('    var cls="cal-cell"+(_sd(dt,_now)?" today":"");');
  js.push('    var dot=evs.length?\'<span class="cal-dot"></span>\':\'\';');
  js.push('    grid+=\'<div class="\'+cls+\'" onclick="_showDay(\'+y+\',\'+(m+1)+\',\'+d+\')"><span class="cal-dn">\'+d+\'</span>\'+dot+\'</div>\';');
  js.push('  }');
  js.push('  grid+="</div>";');
  js.push('  el.innerHTML=nav+grid+\'<div id="dpanel"></div>\';');
  js.push('}');
  js.push('function _showDay(y,m,d){');
  js.push('  var dt=new Date(y,m-1,d);');
  js.push('  var evs=EVTS.filter(function(ev){return _sd(_ed(ev),dt);});');
  js.push('  var out=\'<div class="day-panel"><p class="dp-title">\'+y+\'年\'+m+\'月\'+d+\'日のイベント</p>\';');
  js.push('  out+=evs.length?evs.map(_card).join(""):\'<p class="empty">この日のイベントはありません</p>\';');
  js.push('  out+="</div>";');
  js.push('  var p=document.getElementById("dpanel");');
  js.push('  if(p){p.innerHTML=out;try{google.script.host.setHeight(document.body.scrollHeight+40);}catch(e){}}');
  js.push('}');
  js.push('function _renderWeek(el){');
  js.push('  var ms=_wk;');
  js.push('  var ds=["月","火","水","木","金","土","日"];');
  js.push('  var nav=\'<div class="cal-nav"><button class="nav-btn" onclick="_prevWk()">◀</button><span class="cal-title">\'+ms.getFullYear()+\'年\'+(ms.getMonth()+1)+\'月\'+ms.getDate()+\'日〜</span><button class="nav-btn" onclick="_nextWk()">▶</button></div>\';');
  js.push('  var html=nav+\'<div class="week-grid">\';');
  js.push('  for(var i=0;i<7;i++){');
  js.push('    var dt=new Date(ms.getTime()+i*86400000);');
  js.push('    var evs=EVTS.filter(function(ev){return _sd(_ed(ev),dt);});');
  js.push('    var isT=_sd(dt,_now);');
  js.push('    html+=\'<div class="week-col\'+(isT?" today":"")+\'"><div class="week-hd">\'+ds[i]+\'<br><span class="wdate">\'+(dt.getMonth()+1)+\'/\'+dt.getDate()+\'</span></div><div class="week-evs">\';');
  js.push('    evs.forEach(function(ev){');
  js.push('      var tim=ev.h>=0?("0"+ev.h).slice(-2)+":"+("0"+ev.mi).slice(-2):"";');
  js.push('      var ttl=ev.u?\'<a href="\'+_esc(ev.u)+\'" target="_blank">\'+_esc(ev.t)+\'</a>\':_esc(ev.t);');
  js.push('      html+=\'<div class="week-ev\'+(ev.past?" past":"")+\'">\'+( tim?\'<span class="wtim">\'+tim+\'</span>\':"")+\'<span class="wtit">\'+ttl+\'</span></div>\';');
  js.push('    });');
  js.push('    html+="</div></div>";');
  js.push('  }');
  js.push('  html+="</div>";');
  js.push('  el.innerHTML=html;');
  js.push('}');
  js.push('function _render(){');
  js.push('  var el=document.getElementById("view-area");');
  js.push('  if(!el)return;');
  js.push('  if(_view==="list")_renderList(el);');
  js.push('  else if(_view==="month")_renderMonth(el);');
  js.push('  else _renderWeek(el);');
  js.push('  try{google.script.host.setHeight(document.body.scrollHeight+40);}catch(e){}');
  js.push('}');
  js.push('function _openAdminSheet(url){var p=prompt("管理者PASSを入力してください");if(p===null)return false;if(p===ADMIN_PASS){window.open(url,"_blank","noopener");return false;}alert("PASSが違います");return false;}');
  js.push('window.onload=function(){_render();};');
  return js.join('\n');
}

function getInzaiStyle_() {
  return [
    'body{margin:0;padding:12px;font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif;font-size:14px;color:#333;}',
    '.inzai-wrap{max-width:800px;}',
    '.inzai-page-header{padding:10px 0 12px;margin-bottom:10px;border-bottom:2px solid #c0392b;}',
    '.inzai-page-kicker{margin:0 0 3px;font-size:11px;color:#999;font-weight:bold;letter-spacing:0.08em;}',
    '.inzai-page-header h1{margin:0 0 4px;font-size:22px;line-height:1.35;color:#c0392b;}',
    '.inzai-page-header p{margin:0;font-size:13px;color:#666;}',
    '.month-group{margin-bottom:6px;border:1px solid #eee;border-radius:6px;overflow:hidden;}',
    '.month-summary{padding:10px 14px;cursor:pointer;font-weight:bold;font-size:15px;color:#c0392b;background:#fff5f5;display:flex;align-items:center;justify-content:space-between;list-style:none;user-select:none;}',
    '.month-summary::-webkit-details-marker{display:none;}',
    '.month-summary::before{content:"▶";font-size:11px;margin-right:8px;display:inline-block;transition:transform 0.2s;}',
    'details[open]>.month-summary::before{transform:rotate(90deg);}',
    '.month-count{font-size:12px;color:#999;font-weight:normal;margin-left:auto;}',
    '.month-group .ev-item{padding:12px 14px;border-bottom:1px solid #eee;}',
    '.month-group .ev-item:last-child{border-bottom:none;}',
    '.ev-item{padding:12px 0;border-bottom:1px solid #eee;}',
    '.ev-item.past{opacity:0.45;}',
    '.ev-date{margin:0 0 3px;font-size:12px;color:#c0392b;font-weight:bold;}',
    '.ev-title{margin:0 0 4px;font-size:15px;font-weight:bold;}',
    '.ev-title a{color:#c0392b;text-decoration:underline;}',
    '.ev-title a:hover{color:#a93226;}',
    '.ev-meta{margin:0 0 3px;font-size:13px;color:#555;}',
    '.ev-desc{margin:0 0 3px;font-size:13px;color:#666;line-height:1.5;}',
    '.ev-source{margin:0;font-size:11px;color:#bbb;}',
    '.no-year{font-size:11px;color:#e67e22;font-weight:normal;margin-left:4px;}',
    '.empty{color:#999;text-align:center;padding:20px;}',
    '.entry-actions{display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-start;margin:8px 0 12px;}',
    '.entry-btn{display:inline-flex;align-items:center;justify-content:center;padding:8px 18px;background:#c0392b;color:#fff;border-radius:4px;text-decoration:none;font-size:13px;font-weight:bold;line-height:1.4;}',
    '.entry-btn:hover{background:#a93226;}',
    '.admin-sheet-btn{display:inline-flex;align-items:center;justify-content:center;padding:8px 14px;background:#fff;color:#c0392b;border:1px solid #c0392b;border-radius:4px;text-decoration:none;font-size:13px;font-weight:bold;line-height:1.4;}',
    '.admin-sheet-btn:hover{background:#fff5f5;}',
    // ビュー切り替えタブ
    '.view-tabs{display:flex;gap:0;margin-bottom:12px;border-bottom:2px solid #c0392b;}',
    '.vbtn{padding:7px 18px;border:1px solid #ddd;border-bottom:none;background:#f5f5f5;cursor:pointer;font-size:13px;color:#666;border-radius:4px 4px 0 0;font-family:inherit;line-height:1.4;}',
    '.vbtn.active{background:#fff;color:#c0392b;font-weight:bold;border-color:#c0392b;margin-bottom:-2px;}',
    '.vbtn:hover{background:#fef0f0;}',
    // カレンダー共通
    '.cal-nav{display:flex;align-items:center;gap:8px;margin-bottom:10px;}',
    '.cal-title{font-weight:bold;font-size:15px;flex:1;text-align:center;}',
    '.nav-btn{background:#fff;border:1px solid #ccc;border-radius:4px;padding:5px 12px;cursor:pointer;font-size:14px;color:#555;font-family:inherit;}',
    '.nav-btn:hover{background:#fef0f0;border-color:#c0392b;color:#c0392b;}',
    // 月カレンダー
    '.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;margin-bottom:12px;}',
    '.cal-hd{text-align:center;font-size:11px;color:#999;padding:4px 0;font-weight:bold;}',
    '.cal-cell{min-height:44px;border:1px solid #e8e8e8;border-radius:4px;padding:4px 2px;cursor:pointer;text-align:center;}',
    '.cal-cell:hover{background:#fef0f0;border-color:#c0392b;}',
    '.cal-cell.empty{border:none;background:none;cursor:default;}',
    '.cal-cell.today{background:#fff5f5;border-color:#c0392b;}',
    '.cal-dn{font-size:13px;display:block;}',
    '.cal-dot{display:block;width:7px;height:7px;background:#c0392b;border-radius:50%;margin:3px auto 0;}',
    '.day-panel{margin-top:12px;padding:10px 0;border-top:2px solid #eee;}',
    '.dp-title{font-weight:bold;margin:0 0 8px;font-size:14px;color:#555;}',
    // 週カレンダー
    '.week-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;}',
    '.week-col{border:1px solid #e8e8e8;border-radius:4px;min-height:100px;}',
    '.week-col.today{background:#fff5f5;border-color:#c0392b;}',
    '.week-hd{text-align:center;padding:6px 2px;border-bottom:1px solid #eee;font-size:12px;font-weight:bold;color:#888;}',
    '.week-col.today .week-hd{color:#c0392b;}',
    '.wdate{font-size:11px;font-weight:normal;}',
    '.week-evs{padding:4px 3px;}',
    '.week-ev{font-size:11px;padding:3px 4px;margin-bottom:3px;background:#fafafa;border-left:3px solid #c0392b;border-radius:0 2px 2px 0;}',
    '.week-ev.past{opacity:0.5;}',
    '.wtim{display:block;font-size:10px;color:#c0392b;font-weight:bold;}',
    '.wtit a{color:#333;text-decoration:none;}',
    '.wtit a:hover{color:#c0392b;text-decoration:underline;}',
    '@media(max-width:600px){.vbtn{padding:6px 11px;font-size:12px;}.week-hd{font-size:10px;padding:4px 1px;}.week-ev{font-size:10px;padding:2px 3px;}.cal-cell{min-height:36px;}}',
  ].join('');
}


// ============================================================
// 週次WEB巡回 — 印西市イベント収集
// Apps Script のトリガーで weeklyFetchInzaiEvents を週1回自動実行
// setupInzaiWeeklyTrigger() を一度だけ手動実行してトリガーを設定すること
// ============================================================
function weeklyFetchInzaiEvents() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INZAI_SHEET);

  // シートがなければ新規作成
  if (!sheet) {
    sheet = ss.insertSheet(INZAI_SHEET);
    sheet.getRange(1, 1, 1, 9).setValues([[
      '収集日時', 'イベント名', '開催日時', '場所', 'URL', '概要', '公開', '開催日時（パース）', '情報元'
    ]]);
    Logger.log('シート「' + INZAI_SHEET + '」を作成しました');
  }

  // URLをメインキーに使用（C列を手修正してもURLは変わらないため重複しない）
  // 旧キー（タイトル|日時）も互換のため保持
  var allRows = sheet.getDataRange().getValues();
  var existingKeys = new Set();
  for (var ri = 1; ri < allRows.length; ri++) {
    var r = allRows[ri];
    if (!String(r[1] || '').trim()) continue; // イベント名が空の行はスキップ
    var eUrl = String(r[4] || '').trim();
    if (eUrl) existingKeys.add(eUrl);                                    // URLキー（メイン）
    existingKeys.add(String(r[1]).trim() + '|' + String(r[2]).trim());  // 旧キー（互換）
  }

  var totalAdded = 0;

  INZAI_SOURCES.forEach(function(src) {
    try {
      var fetchHeaders = {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3'
      };
      if (src.referer) fetchHeaders['Referer'] = src.referer;
      var res = UrlFetchApp.fetch(src.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: fetchHeaders
      });
      var code = res.getResponseCode();
      if (code !== 200) {
        Logger.log('取得失敗(HTTP ' + code + '): ' + src.name + ' ' + src.url);
        return;
      }
      var html = res.getContentText('UTF-8');
      Logger.log(src.name + ': HTMLサイズ=' + html.length + '文字');

      var events = (src.parser === 'inzai_city')
        ? parseInzaiCityHtml_(html, src.name, src.url)
        : (src.parser === 'inzainet')
          ? parseInzaiNetHtml_(html, src.name, src.url)
          : (src.parser === 'kosodatenavi')
            ? fetchKosodatenaviEvents_(html, src.name, src.url)
            : (src.parser === 'inzaibunka')
              ? parseInzaiBunkaHtml_(html, src.name, src.url)
              : (src.parser === 'inzaiec')
                ? parseInzaiecHtml_(html, src.name, src.url)
                : (src.parser === 'alcazar')
                  ? parseAlcazarHtml_(html, src.name, src.url)
                  : (src.parser === 'cosmospalette')
                    ? parseCosmospaletteJson_(html, src.name, src.url)
                    : parseInzaiHtml_(html, src.name, src.url);

      events.forEach(function(ev) {
        var urlKey = (ev.url || '').trim();
        var titleDateKey = ev.title + '|' + ev.datetime;
        // URLキー（メイン）または旧タイトル|日時キーで重複チェック
        if ((urlKey && existingKeys.has(urlKey)) || existingKeys.has(titleDateKey)) return;
        // 詳細ページから実際のイベント名を取得（skipDetailFetchフラグがあればスキップ）
        var detailTitle = src.skipDetailFetch ? null : fetchDetailTitle_(ev.url, ev.title);
        var displayTitle = detailTitle || ev.title;
        var displayUrlKey = urlKey; // URLは変わらないのでdisplayTitleが変わっても同じキー
        var displayTitleDateKey = displayTitle + '|' + ev.datetime;
        if ((displayUrlKey && existingKeys.has(displayUrlKey)) || existingKeys.has(displayTitleDateKey)) return;
        var parsedDate = parseCalDate_(ev.datetime) || '';
        // 年なし日付は「西暦未記載」をシートに明示
        var hasYear = /\d{4}年|令和\d+年/.test(ev.datetime);
        var storedDatetime = hasYear ? ev.datetime : ev.datetime + '　西暦未記載';
        // A列の最終データ行を探して追記（チェックボックス空行を飛ばす）
        var colA = sheet.getRange('A:A').getValues();
        var newRow = 2;
        for (var k = colA.length - 1; k >= 0; k--) {
          if (colA[k][0] !== '') { newRow = k + 2; break; }
        }
        sheet.getRange(newRow, 3).setNumberFormat('@STRING@'); // C列を先にテキスト形式に（日付の自動変換防止）
        sheet.getRange(newRow, 1, 1, 9).setValues([[
          new Date(), displayTitle, storedDatetime, ev.location, ev.url, ev.description, true, parsedDate, src.name
        ]]);
        sheet.getRange(newRow, 7).insertCheckboxes();
        sheet.getRange(newRow, 7).setValue(true);
        // 両方のキーを登録
        if (urlKey) existingKeys.add(urlKey);
        existingKeys.add(displayTitleDateKey);
        totalAdded++;
        Utilities.sleep(500); // 詳細ページ取得があるため余裕を持たせる
      });

      Logger.log(src.name + ': ' + events.length + '件取得');
    } catch(err) {
      Logger.log('エラー(' + src.name + '): ' + err);
    }
  });

  Logger.log('=== weeklyFetchInzaiEvents 完了: 合計' + totalAdded + '件追加 ===');
}

function countInzaiEventRows_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(INZAI_SHEET);
  if (!sheet) return 0;
  var rowCount = sheet.getLastRow() - 1;
  if (rowCount <= 0) return 0;
  var values = sheet.getRange(2, 2, rowCount, 1).getValues();
  var count = 0;
  values.forEach(function(row) {
    if (String(row[0] || '').trim()) count++;
  });
  return count;
}

// ============================================================
// いんざいネット専用パーサー
// 対象: https://inzainet.com/event/
// 構造: li > a[href] > h3(タイトル) + p(地域) + p(日付)
// 印西市のイベントのみ取得
// ============================================================
function parseInzaiNetHtml_(json, sourceName, sourceBaseUrl) {
  var events = [];

  // WordPress REST API JSON をパース
  var posts;
  try {
    posts = JSON.parse(json);
  } catch(e) {
    Logger.log(sourceName + ': JSONパースエラー: ' + e);
    return events;
  }
  if (!Array.isArray(posts)) {
    Logger.log(sourceName + ': 予期しないレスポンス形式');
    return events;
  }

  function cleanTag(s) {
    return s.replace(/<[^>]+>/g, '')
            .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
            .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
            .replace(/&#\d+;/g, '').replace(/\s+/g, ' ').trim();
  }

  for (var i = 0; i < posts.length; i++) {
    var post = posts[i];
    var title   = cleanTag((post.title   || {}).rendered || '');
    var url     = post.link || '';
    var content = cleanTag((post.content || {}).rendered || '');
    var excerpt = cleanTag((post.excerpt || {}).rendered || '');

    // 印西フィルター
    if (content.indexOf('印西') === -1 && title.indexOf('印西') === -1 && excerpt.indexOf('印西') === -1) continue;

    // 日付抽出（令和・西暦・年なし に対応）
    var dateM = content.match(/令和\d+年\d{1,2}月\d{1,2}日|\d{4}年\d{1,2}月\d{1,2}日|\d{1,2}月\d{1,2}日/);
    var dateStr = dateM ? dateM[0] : '';

    // 日付がない記事（企業紹介・広告等）はスキップ
    if (!dateStr) continue;

    // 地域抽出
    var regionM = content.match(/印西市|白井市|成田市|佐倉市|四街道市/);
    var region  = regionM ? regionM[0] : '印西市';

    events.push({ title: title, datetime: dateStr, location: region, url: url, description: content.substring(0, 120) });
  }

  Logger.log(sourceName + ': ' + events.length + '件取得（印西市のみ）');
  return events;
}

// ============================================================
// 印西市公式サイト専用パーサー
// 対象: https://www.city.inzai.lg.jp/event2/0curr_1.html
// 構造: article > ul > li に日付テキスト＋<a>リンク
// ============================================================
function parseInzaiCityHtml_(html, sourceName, sourceBaseUrl) {
  var events = [];
  var seen = new Set();
  var base = 'https://www.city.inzai.lg.jp';

  // 日付パターン（年あり・年なし両対応）
  var DATE_PATTERN = /(\d{4}年\d{1,2}月\d{1,2}日[（(]?[日月火水木金土曜日]*[）)]?|\d{1,2}月\d{1,2}日[（(]?[日月火水木金土曜日]*[）)]?)/;

  function toAbsUrl(href) {
    if (!href) return '';
    href = href.trim();
    var full;
    if (href.startsWith('http')) {
      full = href;
    } else if (href.startsWith('/')) {
      full = base + href;
    } else {
      // 相対パス (例: ../bousaiportal/0000021300.html)
      full = base + '/event2/' + href;
    }
    // /xxx/../ を解決して正規化 (/event2/../bousaiportal/ → /bousaiportal/)
    var prev = '';
    while (prev !== full) { prev = full; full = full.replace(/\/[^\/]+\/\.\.\//g, '/'); }
    return full;
  }

  var NAV_SKIP = /スキップ|お問い合わせ|組織一覧|サイトマップ|プライバシー|ホーム|トップ|メニュー|ログイン|検索|採用|会社概要|アクセス|もっと見る|開催日時|開催状況|ご意見|業務報告|シティプロモーション|まっぷる|事業の終了|日常のすぐそば|調達情報|入札|予算|決算|条例|規則|議会|広報|通知|様式|申請|手続|補助金|助成|統計|窓口|証明|住民票|公園|神社|寺院|史跡|名所|観光スポット|道の駅|施設紹介|アクセス|地図|駐車場/;

  function extractFromLi(liContent) {
    var linkM = /<a\s[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i.exec(liContent);
    if (!linkM) return null;
    var href  = linkM[1].trim();
    var title = linkM[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    if (!title || title.length < 2) return null;
    if (NAV_SKIP.test(title)) return null;
    var liText  = liContent.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    var dateM   = DATE_PATTERN.exec(liText);
    return { title: title, datetime: dateM ? dateM[1].trim() : '', url: toAbsUrl(href) };
  }

  // ① article 内の li を対象
  var articleRe = /<article[^>]*>([\s\S]*?)<\/article>/gi;
  var articleMatch;
  while ((articleMatch = articleRe.exec(html)) !== null) {
    var liRe2 = /<li[^>]*>([\s\S]*?)<\/li>/gi;
    var liMatch2;
    while ((liMatch2 = liRe2.exec(articleMatch[1])) !== null) {
      var ev = extractFromLi(liMatch2[1]);
      if (!ev || seen.has(ev.title)) continue;
      seen.add(ev.title);
      events.push({ title: ev.title, datetime: ev.datetime, location: '印西市', url: ev.url, description: '' });
    }
  }
  Logger.log(sourceName + ': ①articleパーサー → ' + events.length + '件');

  // ② article で 0 件 → dl/dd 構造を試す
  if (events.length === 0) {
    var ddRe = /<dd[^>]*>([\s\S]*?)<\/dd>/gi;
    var ddMatch;
    while ((ddMatch = ddRe.exec(html)) !== null) {
      var liRe3 = /<li[^>]*>([\s\S]*?)<\/li>/gi;
      var liMatch3;
      while ((liMatch3 = liRe3.exec(ddMatch[1])) !== null) {
        var ev2 = extractFromLi(liMatch3[1]);
        if (!ev2 || seen.has(ev2.title)) continue;
        seen.add(ev2.title);
        events.push({ title: ev2.title, datetime: ev2.datetime, location: '印西市', url: ev2.url, description: '' });
      }
    }
    Logger.log(sourceName + ': ②dl/ddパーサー → ' + events.length + '件');
  }

  // ③ それでも 0 件 → 全 li を対象（日付がある li のみ採用）
  if (events.length === 0) {
    var liRe4 = /<li[^>]*>([\s\S]*?)<\/li>/gi;
    var liMatch4;
    while ((liMatch4 = liRe4.exec(html)) !== null) {
      var ev3 = extractFromLi(liMatch4[1]);
      if (!ev3 || !ev3.datetime || seen.has(ev3.title)) continue; // 日付なしはスキップ
      seen.add(ev3.title);
      events.push({ title: ev3.title, datetime: ev3.datetime, location: '印西市', url: ev3.url, description: '' });
    }
    Logger.log(sourceName + ': ③全liパーサー（日付あり限定）→ ' + events.length + '件');
  }

  return events;
}

// ============================================================
// いんざい子育てナビ専用クローラー
// カテゴリインデックスページから sub-category URLを抽出し、
// 各サブページを fetchして parseInzaiCityHtml_ でイベントを収集する
// ============================================================
function fetchKosodatenaviEvents_(indexHtml, sourceName, indexUrl) {
  var events = [];
  var base = 'https://www.city.inzai.lg.jp';

  // インデックスページの article > h2 > a から sub-category URL を抽出
  var subUrls = [];
  var articleRe = /<article[^>]*>([\s\S]*?)<\/article>/gi;
  var articleMatch;
  while ((articleMatch = articleRe.exec(indexHtml)) !== null) {
    var h2Re = /<h2[^>]*>[\s\S]*?<a\s[^>]*href=["']([^"']+)["'][^>]*>/gi;
    var h2m;
    while ((h2m = h2Re.exec(articleMatch[1])) !== null) {
      var href = h2m[1].trim();
      var absUrl;
      if (href.startsWith('http')) {
        absUrl = href;
      } else if (href.startsWith('/')) {
        absUrl = base + href;
      } else {
        // 相対パス (./../category/18-2-1-0-0.html)
        absUrl = base + '/kosodatenavi/category/' + href.replace(/^.*\//, '');
      }
      // /xxx/../ 正規化
      var prev = '';
      while (prev !== absUrl) { prev = absUrl; absUrl = absUrl.replace(/\/[^\/]+\/\.\.\//g, '/'); }
      if (subUrls.indexOf(absUrl) === -1) subUrls.push(absUrl);
    }
  }
  Logger.log(sourceName + ': サブカテゴリURL ' + subUrls.length + '件: ' + subUrls.join(', '));

  // 各サブページをfetch & parse
  var seen = new Set();
  subUrls.forEach(function(subUrl) {
    try {
      var res = UrlFetchApp.fetch(subUrl, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
          'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3'
        }
      });
      if (res.getResponseCode() !== 200) {
        Logger.log(sourceName + ': サブページ取得失敗(HTTP ' + res.getResponseCode() + '): ' + subUrl);
        return;
      }
      var subHtml = res.getContentText('UTF-8');
      var subEvents = parseInzaiCityHtml_(subHtml, sourceName + '[' + subUrl.replace(/.*\//, '') + ']', subUrl);
      subEvents.forEach(function(ev) {
        var key = ev.title + '|' + ev.datetime;
        if (seen.has(key)) return;
        seen.add(key);
        events.push(ev);
      });
      Utilities.sleep(300);
    } catch(err) {
      Logger.log(sourceName + ': サブページエラー(' + subUrl + '): ' + err);
    }
  });

  Logger.log(sourceName + ': 合計 ' + events.length + '件取得');
  return events;
}

// ============================================================
// 詳細ページから実際のイベント名を取得するヘルパー
// h1 → h2 → <title> の順で見出しを探す
// ============================================================
function fetchDetailTitle_(url, originalTitle) {
  if (!url) return null;
  try {
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja' }
    });
    if (res.getResponseCode() !== 200) return null;
    var html = res.getContentText('UTF-8');

    function cleanTag(s) {
      return s.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
    }

    // h1 を取得
    var h1m = /<h1[^>]*>([\s\S]*?)<\/h1>/i.exec(html);
    var h1Text = h1m ? cleanTag(h1m[1]) : null;
    if (h1Text && h1Text.length < 4) h1Text = null;

    // h2 を全件取得（最初の有効なもの）
    var h2Text = null;
    var h2Re = /<h2[^>]*>([\s\S]*?)<\/h2>/gi;
    var h2m;
    while ((h2m = h2Re.exec(html)) !== null) {
      var t2 = cleanTag(h2m[1]);
      if (t2.length >= 4) { h2Text = t2; break; }
    }

    // h1 が元タイトルと異なれば採用、同じなら h2 を優先
    if (h1Text && h1Text !== originalTitle) return h1Text;
    if (h2Text && h2Text !== originalTitle) return h2Text;
    if (h1Text) return h1Text;

    // <title> タグ（「| 印西市」等の接尾辞を除去）
    var tlm = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html);
    if (tlm) {
      var t3 = cleanTag(tlm[1]).replace(/[\|｜\/].*/g, '').trim();
      if (t3.length >= 4 && t3 !== originalTitle) return t3;
    }
    return null;
  } catch(e) {
    return null;
  }
}

// HTMLからイベント情報を抽出（汎用パーサー）
function parseInzaiHtml_(html, sourceName, sourceBaseUrl) {
  var events = [];
  var seen = new Set();

  // <a href="...">タイトル</a> を抽出してイベント候補にする
  var linkRe = /<a\s[^>]*href=["']([^"'#?][^"']*)["'][^>]*>\s*([\s\S]{4,80}?)\s*<\/a>/gi;
  var skipRe = /ホーム|トップ|メニュー|ログイン|検索|プライバシー|サイトマップ|お問い合わせ|もっと見る|一覧|アクセス|採用|会社概要/;
  var dateRe = /\d{4}[年\/\.]\d{1,2}[月\/\.]\d{1,2}|[1-9]\d?月\d{1,2}日/;

  // ページ全体から日付候補をまとめて抽出（タイトル付近の日付を紐付けに使う）
  var allDates = [];
  var dr;
  var dateExtRe = /(\d{4}[年\/]\d{1,2}[月\/]\d{1,2}日?|\d{1,2}月\d{1,2}日[（(]?[日月火水木金土]?[）)]?(?:[\s　]*\d{1,2}:\d{2}(?:[〜～~]\d{1,2}:\d{2})?)?)/g;
  while ((dr = dateExtRe.exec(html)) !== null) allDates.push({ idx: dr.index, val: dr[1] });

  var match;
  while ((match = linkRe.exec(html)) !== null) {
    var href  = match[1].trim();
    var title = match[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();

    if (title.length < 4 || skipRe.test(title)) continue;
    if (seen.has(title)) continue;
    seen.add(title);

    // 絶対URLに変換
    var fullUrl = href;
    if (!href.startsWith('http')) {
      try {
        var base = sourceBaseUrl.match(/^https?:\/\/[^\/]+/)[0];
        fullUrl = href.startsWith('/') ? base + href : sourceBaseUrl.replace(/\/[^\/]*$/, '/') + href;
      } catch(e) { fullUrl = href; }
    }

    // リンク位置に最も近い日付を開催日時として採用
    var pos = match.index;
    var nearestDate = '';
    var minDist = 3000;
    allDates.forEach(function(d) {
      var dist = Math.abs(d.idx - pos);
      if (dist < minDist) { minDist = dist; nearestDate = d.val; }
    });

    events.push({
      title:       title,
      datetime:    nearestDate,
      location:    '印西市',
      url:         fullUrl,
      description: '',
    });

    if (events.length >= 30) break; // 1サイトあたり最大30件
  }

  return events;
}

// ============================================================
// デバッグ用: 印西市公式サイトのHTMLを確認する
// Apps Script エディタから手動で実行 → ログで先頭2000文字を確認
// ============================================================
// ============================================================
// 印西市文化ホール パーサー
// https://www.inzai-bunka.jp/event/
// 構造: li > date + h2(a[href]) + table(時間)
// 日付は YYYY年MM月DD日(曜) 形式で明記されている。
// 過去日は除外する。
// ============================================================
function parseInzaiBunkaHtml_(html, sourceName, sourceUrl) {
  var events = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // <li> ブロックを抽出（入れ子対策で末尾タグまでを取る）
  var liRe = /<li[^>]*>([\s\S]*?)<\/li>/gi;
  var m;
  while ((m = liRe.exec(html)) !== null) {
    var li = m[1];

    // 詳細ページURL（/event/数字/ 形式のみ）
    var urlM = li.match(/href="(https?:\/\/www\.inzai-bunka\.jp\/event\/\d+\/)"/i);
    if (!urlM) continue;
    var url = urlM[1];

    // 日付を抽出（範囲も対応：YYYY年MM月DD日(曜) 〜 MM月DD日(曜)）
    var dateM = li.match(/(\d{4}年\d{1,2}月\d{1,2}日[^<]{0,20})/);
    if (!dateM) continue;
    var datetime = dateM[1].trim().replace(/\s+/g, ' ');

    // タイトル（h2/h3内のリンクテキスト）
    var titleM = li.match(/<h[23][^>]*>[\s\S]*?<a[^>]*>([\s\S]*?)<\/a>/i)
              || li.match(/<h[23][^>]*>([\s\S]*?)<\/h[23]>/i);
    if (!titleM) continue;
    var title = titleM[1].replace(/<[^>]+>/g, '').trim();
    if (!title || title.length < 4) continue;

    // 過去イベントを除外
    var evDate = parseEventDate(datetime);
    if (evDate && evDate < today) continue;

    events.push({
      title: title,
      datetime: datetime,
      location: '印西市文化ホール',
      url: url,
      description: ''
    });
  }

  Logger.log(sourceName + ': ' + events.length + '件取得');
  return events;
}

// ============================================================
// いんざい市民スクール 講座一覧パーサー
// https://inzaiec.machikatsu.co.jp/
// 構造: .eventbox > article 内に講座名、開催日、施設名、概要、詳細URL
// ============================================================
function parseInzaiecHtml_(html, sourceName, sourceUrl) {
  var events = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  function cleanText(s) {
    return String(s || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/&quot;/g, '"')
      .replace(/&amp;/g, '&')
      .replace(/&nbsp;/g, ' ')
      .replace(/&#39;/g, "'")
      .replace(/<[^>]+>/g, ' ')
      .replace(/[ \t]+/g, ' ')
      .replace(/\n\s+/g, '\n')
      .replace(/\s+\n/g, '\n')
      .trim();
  }

  function toAbsUrl(href) {
    if (!href) return '';
    href = href.trim();
    if (href.indexOf('http') === 0) return href;
    if (href.charAt(0) === '/') return 'https://inzaiec.machikatsu.co.jp' + href;
    return sourceUrl.replace(/\/[^\/]*$/, '/') + href;
  }

  var boxRe = /<div[^>]*class=["'][^"']*eventbox[^"']*["'][^>]*>([\s\S]*?)<!-- End Part/gi;
  var m;
  while ((m = boxRe.exec(html)) !== null) {
    var box = m[1];

    var titleM = box.match(/<h3[^>]*>([\s\S]*?)<\/h3>/i);
    var dateM = box.match(/開催日｜\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    var urlM = box.match(/<a\s+[^>]*href=["']([^"']*course\/detail\.html\?no=\d+[^"']*)["'][^>]*>/i);
    if (!titleM || !dateM || !urlM) continue;

    var title = cleanText(titleM[1]);
    if (!title || title.length < 2) continue;

    var datetime = dateM[1] + '年' + parseInt(dateM[2], 10) + '月' + parseInt(dateM[3], 10) + '日';
    var evDate = parseEventDate(datetime);
    if (evDate && evDate < today) continue;

    var locationM = box.match(/<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i);
    var descM = box.match(/<p[^>]*class=["'][^"']*ellipsis[^"']*["'][^>]*>([\s\S]*?)<\/p>/i);
    var genreM = box.match(/<ul[^>]*class=["'][^"']*mark[^"']*["'][^>]*>[\s\S]*?<li[^>]*>([\s\S]*?)<\/li>/i);
    var description = cleanText(descM ? descM[1] : '');
    var genre = cleanText(genreM ? genreM[1] : '');
    if (genre && description.indexOf(genre) === -1) {
      description = genre + (description ? '\n' + description : '');
    }

    events.push({
      title: title,
      datetime: datetime,
      location: cleanText(locationM ? locationM[1] : '') || '印西市',
      url: toAbsUrl(urlM[1]),
      description: description
    });
  }

  Logger.log(sourceName + ': ' + events.length + '件取得');
  return events;
}

// ============================================================
// アルカサール（千葉NT）イベントページ パーサー
// ============================================================
function parseAlcazarHtml_(html, sourceName, sourceUrl) {
  var events = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // <li class="p-event-topics-top-card">...から日付・タイトル・URLを抽出
  var liRe = /<li[^>]*p-event-topics-top-card[^>]*>([\s\S]*?)<\/li>/gi;
  var m;
  while ((m = liRe.exec(html)) !== null) {
    var li = m[1];

    // URL
    var urlM = li.match(/href="(https?:\/\/nt-alcazar\.com\/event-topics\/[^"]+)"/i);
    if (!urlM) continue;
    var url = urlM[1];

    // 日付 <time>YYYY/MM/DD</time>
    var dateM = li.match(/<time[^>]*>(\d{4})\/(\d{1,2})\/(\d{1,2})<\/time>/i);
    if (!dateM) continue;
    var datetime = dateM[1] + '年' + parseInt(dateM[2]) + '月' + parseInt(dateM[3]) + '日';

    // カテゴリ（「その他」は除外）
    var catM = li.match(/<p[^>]*m-cat[^>]*>([^<]+)<\/p>/i);
    var category = catM ? catM[1].trim() : '';
    if (category === 'その他') continue;

    // 過去イベントを除外
    var evDate = parseEventDate(datetime);
    if (evDate && evDate < today) continue;

    // タイトル <h3>
    var titleM = li.match(/<h3[^>]*>([\s\S]*?)<\/h3>/i);
    if (!titleM) continue;
    var title = titleM[1].replace(/<[^>]+>/g, '').trim();
    if (!title || title.length < 2) continue;

    events.push({
      title:       title,
      datetime:    datetime,
      location:    '千葉ニュータウン中央',
      url:         url,
      description: ''
    });
  }

  Logger.log(sourceName + ': ' + events.length + '件取得');
  return events;
}

// ============================================================
// コスモスパレット パーサー（WordPress REST API: /wp/v2/events）
// ============================================================
function parseCosmospaletteJson_(json, sourceName, sourceBaseUrl) {
  var events = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var posts;
  try {
    posts = JSON.parse(json);
  } catch(e) {
    Logger.log(sourceName + ': JSONパースエラー: ' + e);
    return events;
  }
  if (!Array.isArray(posts)) {
    Logger.log(sourceName + ': 予期しないレスポンス形式');
    return events;
  }

  function cleanTag(s) {
    return s.replace(/<[^>]+>/g, '')
            .replace(/&amp;/g, '&').replace(/&nbsp;/g, ' ')
            .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
            .replace(/&#\d+;/g, '').replace(/\s+/g, ' ').trim();
  }

  for (var i = 0; i < posts.length; i++) {
    var post = posts[i];
    var title   = cleanTag((post.title   || {}).rendered || '').replace(/^\d{1,2}月\d{1,2}日[（(][^）)]*[）)]\s*/, '').replace(/^\d{4}年\d{1,2}月\d{1,2}日[（(][^）)]*[）)]\s*/, '').trim();
    var url     = post.link || '';
    var content = cleanTag((post.content || {}).rendered || '');
    if (!title || !url) continue;

    // 日時：YYYY年M月D日（曜）HH:MM ～ HH:MM 形式から抽出
    var dateM = content.match(/日時[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)/);
    var datetime = dateM ? dateM[1] : '';

    // 年なしタイトル内の月日も試みる（例: "5月23日（土）"）
    if (!datetime) {
      var titleDateM = title.match(/(\d{1,2})月(\d{1,2})日/);
      if (titleDateM) {
        var year = today.getFullYear();
        var m = parseInt(titleDateM[1]);
        var d = parseInt(titleDateM[2]);
        // 月が過去の場合は来年として扱う
        if (m < today.getMonth() + 1 || (m === today.getMonth() + 1 && d < today.getDate())) {
          year++;
        }
        datetime = year + '年' + m + '月' + d + '日';
      }
    }
    if (!datetime) continue;

    // 過去イベントを除外
    var evDate = parseEventDate(datetime);
    if (evDate && evDate < today) continue;

    // 会場：〜 から location を抽出
    var locM = content.match(/会場[：:]\s*([^。\n]+)/);
    var location = locM ? locM[1].trim().replace(/[。、]$/, '') : 'コスモスパレット';

    // 概要（日時・会場行を除いた残り）
    var desc = content.replace(/日時[：:][^\n。]+[。\n]?/g, '')
                      .replace(/会場[：:][^\n。]+[。\n]?/g, '')
                      .replace(/\s+/g, ' ').trim().slice(0, 120);

    events.push({ title: title, datetime: datetime, location: location, url: url, description: desc });
  }

  Logger.log(sourceName + ': ' + events.length + '件取得');
  return events;
}

// ============================================================
// 印西とぴっく パーサー（WordPress REST API版）※現在未使用
// ============================================================
function parseInzaiTopicHtml_(jsonStr, sourceName, indexUrl) {
  var events = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var posts;
  try {
    posts = JSON.parse(jsonStr);
  } catch(e) {
    Logger.log(sourceName + ': JSONパース失敗 - ' + e.message);
    return events;
  }
  if (!Array.isArray(posts)) {
    Logger.log(sourceName + ': レスポンスが配列ではありません');
    return events;
  }

  // イベントカテゴリのスラッグ確認用（全投稿から取得）
  posts.forEach(function(post) {
    // タイトル
    var title = (post.title && post.title.rendered)
      ? post.title.rendered.replace(/<[^>]+>/g, '').trim()
      : '';
    if (!title || title.length < 4) return;

    // URL
    var url = post.link || '';
    if (!url) return;

    // 本文をプレーンテキストに変換
    var bodyHtml = (post.content && post.content.rendered) ? post.content.rendered : '';
    var bodyText = bodyHtml
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>');

    // 開催日を抽出（■開催日、開催日：、日時：）
    var datetime = '';
    var datePatterns = [
      /■開催日[\s\S]*?(\d{4}年\d{1,2}月\d{1,2}日[^\n■]{0,30})/,
      /開催日[時]?[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日[^\n]{0,30})/,
      /日時[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日[^\n]{0,30})/
    ];
    for (var p = 0; p < datePatterns.length; p++) {
      var dm = bodyText.match(datePatterns[p]);
      if (dm) { datetime = dm[1].trim().replace(/\s+/g, ' '); break; }
    }

    // 本文に年付き日付がなければタイトルから抽出
    if (!datetime) {
      var tdm = title.match(/(\d{4}年\d{1,2}月\d{1,2}日[^\s　。、！？]*)/)
             || title.match(/(\d{1,2}月\d{1,2}日[^\s　。、！？]*)/);
      if (tdm) datetime = tdm[1];
    }

    // 過去イベントを除外
    if (datetime) {
      var evDate = parseEventDate(datetime);
      if (evDate && evDate < today) return;
    }

    // 開催場所を抽出
    var location = '印西市';
    var locPatterns = [
      /■開催場所[\s\S]*?\n([^\n■]{4,60})/,
      /場所[：:]\s*([^\n]{4,60})/
    ];
    for (var lp = 0; lp < locPatterns.length; lp++) {
      var lm = bodyText.match(locPatterns[lp]);
      if (lm) { location = lm[1].trim().replace(/\s+/g, ' ').substring(0, 50); break; }
    }

    events.push({
      title: title,
      datetime: datetime,
      location: location,
      url: url,
      description: ''
    });
  });

  Logger.log(sourceName + ': ' + events.length + '件取得');
  return events;
}

function debugFetchInzaiCity() {
  var url = 'https://www.city.inzai.lg.jp/event2/0curr_1.html';
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml',
      'Accept-Language': 'ja,en-US;q=0.7,en;q=0.3'
    }
  });
  Logger.log('HTTP: ' + res.getResponseCode());
  var html = res.getContentText('UTF-8');
  Logger.log('HTMLサイズ: ' + html.length);
  Logger.log('先頭2000文字:\n' + html.substring(0, 2000));
  Logger.log('--- 末尾500文字 ---\n' + html.substring(Math.max(0, html.length - 500)));

  // パーサーテスト
  var events = parseInzaiCityHtml_(html, '印西市公式', url);
  Logger.log('パーサー結果: ' + events.length + '件');
  events.slice(0, 5).forEach(function(ev) {
    Logger.log('  ・' + ev.datetime + ' / ' + ev.title + ' / ' + ev.url);
  });
}

// ============================================================
// 週次トリガー設定（一度だけ手動実行）
// 毎週月曜日 朝6時に weeklyFetchInzaiEvents を自動実行
// ============================================================
function setupInzaiWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'weeklyFetchInzaiEvents') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyFetchInzaiEvents')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();
  Logger.log('週次トリガー設定完了: 毎週月曜日 6時に weeklyFetchInzaiEvents を実行');
}


// ============================================
// ユーティリティ
// ============================================
function getDriveImageUrl(url) {
  var match = url.match(/(?:id=|\/d\/)([a-zA-Z0-9_-]+)/);
  if (match) {
    return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w600';
  }
  return url;
}

function formatDate(datetime) {
  if (!datetime) return '';
  var d = new Date(datetime);
  if (isNaN(d.getTime())) return String(datetime);
  var days = ['日', '月', '火', '水', '木', '金', '土'];
  return d.getFullYear() + '年'
    + (d.getMonth() + 1) + '月'
    + d.getDate() + '日（' + days[d.getDay()] + '）'
    + ' ' + zeroPad(d.getHours()) + ':' + zeroPad(d.getMinutes());
}

function formatEventDateDisplay(text) {
  if (!text) return '';
  // スプレッドシートからDateオブジェクトとして取得された場合
  if (text instanceof Date) {
    var days = ['日', '月', '火', '水', '木', '金', '土'];
    return text.getFullYear() + '年' + (text.getMonth() + 1) + '月'
      + text.getDate() + '日（' + days[text.getDay()] + '）';
  }
  var s = String(text);
  // シート上の「西暦未記載」マーカーを検出して除去（後でspanとして出力）
  var noYearMarked = /　西暦未記載$/.test(s);
  if (noYearMarked) s = s.replace(/　西暦未記載$/, '').trim();
  var days = ['日', '月', '火', '水', '木', '金', '土'];
  // 時間範囲（日付 + HH:MM〜HH:MM）の場合は日付部分だけ漢字変換してから時間を付加
  var timeRangeMatch = s.match(/^(.*?)\s+(\d{1,2}:\d{2})[〜～~](\d{1,2}:\d{2})\s*$/);
  if (timeRangeMatch) {
    var datePart = formatEventDateDisplay(timeRangeMatch[1].trim());
    return datePart + '　' + timeRangeMatch[2] + '〜' + timeRangeMatch[3];
  }
  // 日付期間表記（〜 ～ ~ を含む）は各部分を漢字変換して結合
  var rangeChar = s.indexOf('〜') !== -1 ? '〜' : (s.indexOf('～') !== -1 ? '～' : (s.indexOf('~') !== -1 ? '~' : null));
  if (rangeChar) {
    var parts = s.split(rangeChar);
    var convertPart = function(part) {
      part = part.trim();
      // YYYY.M.D または YYYY/M/D 形式
      var dm = part.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})/);
      if (dm) return dm[1] + '年' + parseInt(dm[2]) + '月' + parseInt(dm[3]) + '日';
      // M.D または M/D 形式（年なし）
      var md = part.match(/^(\d{1,2})[./](\d{1,2})$/);
      if (md) return parseInt(md[1]) + '月' + parseInt(md[2]) + '日';
      // D日のみ
      var d2 = part.match(/^(\d{1,2})日?$/);
      if (d2) return parseInt(d2[1]) + '日';
      return part;
    };
    var converted = parts.map(convertPart).join('〜');
    // 年が含まれていない場合は「西暦未記載」を付加
    if (noYearMarked || (!converted.match(/\d{4}年/) && !converted.match(/令和/) && converted.match(/\d{1,2}月/))) {
      return escapeHtml(converted) + '<span class="no-year">西暦未記載</span>';
    }
    return escapeHtml(converted);
  }
  // YYYY.M.D または YYYY/M/D 形式
  var dotMatch = s.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})/);
  if (dotMatch) {
    var dd = new Date(parseInt(dotMatch[1]), parseInt(dotMatch[2]) - 1, parseInt(dotMatch[3]));
    return dotMatch[1] + '年' + parseInt(dotMatch[2]) + '月'
      + parseInt(dotMatch[3]) + '日（' + days[dd.getDay()] + '）';
  }
  // 標準フォーマットでパース可能な場合は整形
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.getFullYear() + '年' + (d.getMonth() + 1) + '月'
      + d.getDate() + '日（' + days[d.getDay()] + '）';
  }
  // 年が含まれていない日本語日付は「西暦未記載」を付加
  if (noYearMarked || (s.match(/\d{1,2}月/) && !s.match(/\d{4}年/) && !s.match(/令和/))) {
    return escapeHtml(s) + '<span class="no-year">西暦未記載</span>';
  }
  return escapeHtml(s);
}

// 日付範囲文字列から終了日を返す（〜がなければnull）
function parseEventEndDate_(text) {
  if (!text) return null;
  var s = String(text);
  var idx = s.indexOf('〜');
  if (idx === -1) idx = s.indexOf('～');
  if (idx === -1) idx = s.indexOf('~');
  if (idx === -1) return null;
  var endPart = s.substring(idx + 1).trim();
  if (!endPart) return null;
  // 終了部分に月がなければ開始部分から月を補う（例: 3月17日〜31日）
  if (!endPart.match(/\d{1,2}月/)) {
    var monthMatch = s.match(/(\d{1,2})月/);
    if (monthMatch) endPart = monthMatch[1] + '月' + endPart;
  }
  // 終了部分に年がなければ開始部分から年を補う
  if (!endPart.match(/\d{4}年/) && !endPart.match(/令和/)) {
    var yearMatch = s.match(/(\d{4})年/);
    var reiwaMatch = s.match(/令和(\d+)年/);
    if (yearMatch) endPart = yearMatch[1] + '年' + endPart;
    else if (reiwaMatch) endPart = s.substring(0, idx + 1) + endPart;
  }
  return parseEventDate(endPart);
}


function parseEventDate(text) {
  if (!text) return null;
  var s = String(text);

  // 標準フォーマット（2026/03/22 など）
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // YYYY.M.D または YYYY/M/D 形式
  var dotMatch = s.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})/);
  if (dotMatch) {
    return new Date(parseInt(dotMatch[1]), parseInt(dotMatch[2]) - 1, parseInt(dotMatch[3]));
  }

  var year = null, month = null, day = null;

  // 令和N年
  var reiwa = s.match(/令和(\d+)年/);
  if (reiwa) year = 2018 + parseInt(reiwa[1]);

  // YYYY年
  if (!year) {
    var y = s.match(/(\d{4})年/);
    if (y) year = parseInt(y[1]);
  }

  // M月 D日
  var m = s.match(/(\d{1,2})月/);
  var dy = s.match(/(\d{1,2})日/);
  if (m) month = parseInt(m[1]);
  if (dy) day = parseInt(dy[1]);

  if (month && day) {
    if (!year) return null; // 年情報なし → nullで呼び出し側に委ねる
    return new Date(year, month - 1, day);
  }

  return null;
}

function debugInzaiRows() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INZAI_SHEET);
  if (!sheet) { Logger.log('シートなし'); return; }
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var CUTOFF = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  Logger.log('CUTOFF: ' + CUTOFF.toISOString());
  Logger.log('総行数: ' + (data.length - 1));
  var start = Math.max(1, data.length - 10);
  for (var i = start; i < data.length; i++) {
    var row = data[i];
    var title = String(row[1] || '').trim();
    var rawDate = String(row[2] || '');
    var g = row[6];
    var h = row[7];
    var parsed = parseEventDate(rawDate);
    var hDate = h instanceof Date ? h : null;
    var effectiveParsed = parsed || (hDate && hDate >= CUTOFF ? hDate : null);
    var effectiveEnd = effectiveParsed;
    var filtered = effectiveEnd && effectiveEnd < CUTOFF;
    Logger.log('行' + (i+1) + ': title=[' + title + '] C=[' + rawDate + '] G=[' + g + '/' + (typeof g) + '] H=[' + h + '/' + (h instanceof Date) + '] parseEventDate=[' + parsed + '] effectiveParsed=[' + effectiveParsed + '] filtered=' + filtered);
  }
}

function formatDateShort(datetime) {
  if (!datetime) return '';
  var d = new Date(datetime);
  if (isNaN(d.getTime())) return String(datetime);
  return d.getFullYear() + '.'
    + zeroPad(d.getMonth() + 1) + '.'
    + zeroPad(d.getDate());
}

function zeroPad(n) {
  return n < 10 ? '0' + n : String(n);
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
}

// ============================================
// スタイル（ホームページ）
// ============================================
function getStyle() {
  return [
    'body { margin:0; padding:0; font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif; background:#fff; }',
    '.topics-container { display:flex; flex-wrap:wrap; gap:16px; padding:16px; justify-content:flex-start; }',
    '.card { background:#fff; border:1px solid #e0e0e0; border-radius:10px; overflow:hidden;',
    '  width:calc(50% - 8px); box-shadow:0 2px 6px rgba(0,0,0,0.08); box-sizing:border-box; }',
    '@media(max-width:600px){ .card { width:100%; } }',
    '.card-image { aspect-ratio:210 / 297; overflow:hidden; background:#f7f7f7; }',
    '.card-image img { width:100%; height:100%; object-fit:cover; object-position:center top; display:block; }',
    '.card-body { padding:14px 16px; }',
    '.category { display:inline-block; margin-bottom:6px; padding:2px 10px;',
    '  background:#f0e6e6; color:#c0392b; border-radius:12px; font-size:12px; font-weight:bold; }',
    '.date { margin:0 0 4px; font-size:13px; color:#888; }',
    '.title { margin:0 0 8px; font-size:16px; font-weight:bold; color:#333; }',
    '.meta { margin:0 0 6px; font-size:13px; color:#555; }',
    '.summary { margin:0 0 6px; font-size:13px; color:#555; line-height:1.5; }',
    '.participants { margin:0 0 10px; font-size:12px; color:#777; }',
    '.form-btn { display:inline-block; margin-top:8px; padding:8px 16px;',
    '  background:#c0392b; color:#fff; border-radius:20px; text-decoration:none;',
    '  font-size:13px; font-weight:bold; }',
    '.form-btn:hover { background:#a93226; }',
    '.inzai-nav { padding:10px 16px; background:#fff8f8; border-bottom:2px solid #f0d0d0; }',
    '.inzai-nav a { color:#c0392b; font-size:14px; font-weight:bold; text-decoration:none; }',
    '.inzai-nav a:hover { text-decoration:underline; }',
  ].join('');
}

// ============================================
// スタイル（カテゴリページ）
// ============================================
function getCategoryStyle() {
  return [
    'body { margin:0; padding:12px; font-family:"Hiragino Kaku Gothic ProN","Meiryo",sans-serif; font-size:14px; color:#333; }',
    '.category-container { max-width:800px; }',
    '.project-section { margin-bottom:24px; }',
    '.project-name { font-size:15px; font-weight:bold; color:#333; margin:0 0 8px;',
    '  border-bottom:2px solid #c0392b; padding-bottom:4px; }',
    '.activity-list { list-style:none; padding:0; margin:0; }',
    '.activity-item { padding:5px 0; border-bottom:1px dotted #eee; line-height:1.6; }',
    '.act-date { color:#888; font-size:13px; font-family:monospace; }',
    '.act-title { font-weight:bold; }',
    '.act-organizer { color:#666; font-size:13px; }',
    '.empty { color:#999; text-align:center; padding:20px; }',
  ].join('');
}

// ============================================================
// 手動イベント一括追加（一度だけ実行して削除してOK）
// ============================================================
function addManualEvents() {
  var ss = SpreadsheetApp.openById('1Xma1V92uNPTcXmj1cDNzU0jePcxAtTR_4vEF2ZXRgdo');
  var sheet = ss.getSheetByName('印西イベント');
  var now = new Date();

  var events = [
    ['吉田ガーデン・ルリビタキの花園 春のオープンガーデン', '〜2026年5月10日(日) 10時〜17時30分',       '吉田ガーデン・ルリビタキの花園（佐倉市大佐倉310-5）',   '', '料金：無料 ※事前予約 TEL:090-5495-5425 4/4・4/25ガーデンコンサート、5/10ギター弾き語りあり', new Date(2026,4,10)],
    ['第3回 さくら霊園落語会',                              '2026年3月21日(土) 13時30分開場 14時開演',   'さくら霊園（佐倉市小竹699-3）',                         '', '料金：500円（小中学生無料）定員70人 TEL:043-461-7733',                                       new Date(2026,2,21)],
    ['第4回 朗読「どんぐりの会」',                          '2026年3月31日(火) 13時20分〜14時45分',      '習志野市民プラザ大久保スタジオ1',                       '', '料金：無料 先着25人（成人対象） TEL:090-2176-1749',                                           new Date(2026,2,31)],
    ['写真展（会員によるテーマ作品と自由作品）',            '2026年4月8日(水)〜13日(月) 10〜17時（最終日〜16時）', '勝田台ステーションギャラリー',                    '', '料金：無料 TEL:090-8505-0096',                                                                new Date(2026,3,13)],
    ['船橋市立海神中学校吹奏楽部 第33回定期演奏会',         '2026年3月24日(火) 17時30分開場 18時開演',   'かつしかシンフォニーヒルズ モーツァルトホール',          '', '料金：無料 ※未就学児入場不可 TEL:047-431-3074',                                               new Date(2026,2,24)],
    ['船橋市立法田中学校吹奏楽部 第38回定期演奏会',         '2026年3月26日(木) 14時30分開場 15時開演',   '松戸市民会館',                                          '', '料金：無料 TEL:047-438-3026',                                                                 new Date(2026,2,26)],
    ['はじめてのモルック教室',                              '2026年4月5日(日) 10〜12時 ※雨天中止',       '八千代総合運動公園多目的広場',                          '', '料金：1,000円 先着48人（小3以上）申込:yachiyopark-sun@sunwax.jp TEL:047-406-3010',            new Date(2026,3,5)],
  ];

  events.forEach(function(ev) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 9).setValues([[
      now, ev[0], ev[1], ev[2], ev[3], ev[4], true, ev[5], '手動入力'
    ]]);
    sheet.getRange(newRow, 7).insertCheckboxes();
    sheet.getRange(newRow, 7).setValue(true);
    sheet.getRange(newRow, 3).setNumberFormat('@STRING@');
  });

  Logger.log('7件追加完了');
}

// ============================================================
// 広報いんざい 令和8年4月号 No.1040 イベント一括追加
// 14面「おでかけインフォメーション」+ 15面「文化ホール情報」
// ============================================================
function addKouhouEvents() {
  var ss = SpreadsheetApp.openById('1Xma1V92uNPTcXmj1cDNzU0jePcxAtTR_4vEF2ZXRgdo');
  var sheet = ss.getSheetByName('印西イベント');
  var now = new Date();
  var sourceUrl = 'https://www.city.inzai.lg.jp/0000021451.html';

  var events = [
    // 14面 おでかけインフォメーション
    ['ほくそう春まつり2026（市制施行30周年記念）',
     '2026年4月19日(日) 10時〜15時 ※雨天決行・荒天中止',
     'イオンモール千葉ニュータウン提携駐車場（中央北第1駐車場）',
     sourceUrl,
     'プロアーティストステージ・各鉄道グッズ販売・ミニ電車・飲食物販あり TEL:33-4415',
     new Date(2026,3,19)],
    ['春の全国交通安全運動 出動式',
     '2026年4月5日(日) 14時〜16時（運動期間4月6日〜15日）',
     'イオンモール千葉ニュータウン コスモス広場',
     sourceUrl,
     'TEL:33-7222（市民活動推進課）/ 42-0110（印西警察署）',
     new Date(2026,3,5)],
    ['柏レイソル 印西市ホームタウンデー',
     '2026年4月29日(水・祝) 16時〜',
     '三協フロンテア柏スタジアム（柏市）',
     sourceUrl,
     '市内在住小学生無料（1家族3人まで）同伴保護者5,000円 定員150人(抽選) 申込締切4/15 TEL:04-7162-2250',
     new Date(2026,3,29)],
    ['NECグリーンロケッツ東葛 印西市ホストタウンデー',
     '2026年5月2日(土) 14時30分〜',
     '柏の葉公園総合競技場（柏市）',
     sourceUrl,
     '無料（市内在住・在勤・在学者対象 1組2人まで） 定員1,500組 申込3/27〜5/1 要JapanRugbyID登録',
     new Date(2026,4,2)],
    ['第31回印西市民文化祭 合唱の集い 参加者募集',
     '2026年11月3日(火・祝)',
     '文化ホール',
     sourceUrl,
     '参加費1団体1,500円程度 申込4/1〜4/24 TEL:33-4836 Mail:bunkashinko@city.inzai.chiba.jp',
     new Date(2026,10,3)],
    ['ニュースポーツ教室（ピックルボール・ボッチャ）',
     '2026年5月15日〜6月19日 各金曜日 19時〜21時',
     '松山下公園総合体育館',
     sourceUrl,
     '無料 定員50人 対象:市内在住等小学生以上 申込4/1〜5/8 TEL:42-8417',
     new Date(2026,5,19)],
    ['クライミング教室',
     '2026年5月16日〜30日 各土曜日 17時30分〜19時',
     '松山下公園総合体育館',
     sourceUrl,
     '1日500円 定員20人 対象:小3以上 申込4/1〜4/30(抽選) TEL:42-8417',
     new Date(2026,4,30)],
    // 15面 文化ホール情報
    ['フライデーナイトコンサートVol.13 愛と幻想〜歌とピアノで紡ぐ夜〜',
     '2026年5月1日(金) 19時30分〜',
     '文化ホール（印西市大森2535）',
     'https://www.inzai-bunka.jp/event/37692/',
     '一般1,000円/高校生以下500円/小学生無料 未就学児入場不可 TEL:42-8811',
     new Date(2026,4,1)],
    ['こどもパフォーマー研究所（ラボ）',
     '2026年5月30日〜7月25日 全6回（体験会5/30）',
     '文化ホール（印西市大森2535）',
     'https://www.inzai-bunka.jp/event/',
     '体験会1,000円/全6回5,000円 対象:小中学生 申込4/11〜 TEL:42-8811',
     new Date(2026,6,25)],
    ['コスモスパレットマルシェ 出店者募集',
     '2026年5月23日(土) 10時〜15時 ※小雨決行',
     'コスモスパレット パレットII',
     sourceUrl,
     '出店費屋内4,000円/屋外3,500円 各20ブース 申込締切4/10 TEL:33-6017',
     new Date(2026,4,23)],
  ];

  events.forEach(function(ev) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 9).setValues([[
      now, ev[0], ev[1], ev[2], ev[3], ev[4], true, ev[5], '広報いんざい4月号'
    ]]);
    sheet.getRange(newRow, 7).insertCheckboxes();
    sheet.getRange(newRow, 7).setValue(true);
    sheet.getRange(newRow, 3).setNumberFormat('@STRING@');
  });

  Logger.log('10件追加完了');
}

// ============================================================
// 広報いんざい PDF自動取得（Gemini API版）
// 毎月1日・15日にトリガー実行
// 14面「おでかけインフォメーション」+ 15面「文化ホール情報」
// ============================================================

// 毎日実行。1日・15日(前後1日許容)のみPDF取得処理を実行する。
// setupKouhouTrigger() で日次トリガーを設定すること。
function fetchKouhouInzaiEvents() {
  var today = new Date();
  var day = today.getDate();
  if (day !== 1 && day !== 2 && day !== 15 && day !== 16) {
    Logger.log('広報PDF: 本日(' + day + '日)は実行スキップ');
    return;
  }
  fetchKouhouInzaiEvents_();
}

// 手動実行・テスト用（日付チェックなし）
function fetchKouhouInzaiEventsNow() {
  fetchKouhouInzaiEvents_();
}

function fetchKouhouInzaiEvents_() {
  var props = PropertiesService.getScriptProperties();

  // Step1: 一覧ページから最新記事URLを取得
  var listUrl = 'https://www.city.inzai.lg.jp/category/2-6-19-0-0.html';
  var listRes = UrlFetchApp.fetch(listUrl, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
  });
  if (listRes.getResponseCode() !== 200) {
    Logger.log('広報PDF: 一覧ページ取得失敗 HTTP ' + listRes.getResponseCode());
    return;
  }
  var listHtml = listRes.getContentText('UTF-8');

  // 最新記事URL（/0000XXXXX.html 形式）
  var articleUrlM = listHtml.match(/href="(https?:\/\/www\.city\.inzai\.lg\.jp\/\d+\.html)"/);
  if (!articleUrlM) {
    var relM = listHtml.match(/href="(\/\d{10}\.html)"/);
    if (relM) articleUrlM = ['', 'https://www.city.inzai.lg.jp' + relM[1]];
  }
  if (!articleUrlM) {
    Logger.log('広報PDF: 記事URLが見つかりません');
    return;
  }
  var articleUrl = articleUrlM[1];
  Logger.log('広報PDF: 最新記事URL = ' + articleUrl);

  // Step2: 記事ページからPDF URLを取得
  var articleRes = UrlFetchApp.fetch(articleUrl, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
  });
  if (articleRes.getResponseCode() !== 200) {
    Logger.log('広報PDF: 記事ページ取得失敗 HTTP ' + articleRes.getResponseCode());
    return;
  }
  var articleHtml = articleRes.getContentText('UTF-8');
  var pdfUrlM = articleHtml.match(/href="([^"]*cmsfiles[^"]*\.pdf)"/i);
  if (!pdfUrlM) {
    Logger.log('広報PDF: PDF URLが見つかりません');
    return;
  }
  var pdfUrl = pdfUrlM[1];
  if (pdfUrl.indexOf('http') !== 0) {
    pdfUrl = 'https://www.city.inzai.lg.jp' + (pdfUrl.charAt(0) === '/' ? '' : '/') + pdfUrl;
  }
  Logger.log('広報PDF: PDF URL = ' + pdfUrl);

  // Step3: 同じPDFを既処理済みならスキップ
  var processedKey = 'kouhou_pdf_' + pdfUrl;
  if (props.getProperty(processedKey)) {
    Logger.log('広報PDF: 処理済みスキップ ' + pdfUrl);
    return;
  }

  // Step4: PDFをダウンロードしてBase64化
  Logger.log('広報PDF: ダウンロード開始...');
  var pdfRes = UrlFetchApp.fetch(pdfUrl, {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
  });
  if (pdfRes.getResponseCode() !== 200) {
    Logger.log('広報PDF: PDFダウンロード失敗 HTTP ' + pdfRes.getResponseCode());
    return;
  }
  var pdfBase64 = Utilities.base64Encode(pdfRes.getContent());
  Logger.log('広報PDF: ダウンロード完了 Base64サイズ = ' + Math.round(pdfBase64.length / 1024) + 'KB');

  // Step5: Gemini APIでイベント情報を抽出
  var geminiKey = props.getProperty('GEMINI_API_KEY');
  if (!geminiKey) {
    Logger.log('広報PDF: GEMINI_API_KEY が未設定です');
    Logger.log('設定方法: GASエディタ「プロジェクトの設定」→「スクリプトプロパティ」→ GEMINI_API_KEY を追加');
    return;
  }

  var prompt = 'この「広報いんざい」PDFの14ページ「おでかけインフォメーション」と'
    + '15ページ「文化ホール情報」に掲載されているイベント・催しを全て抽出してください。\n'
    + '各種募集（出店者募集・参加者募集なども含む）も含めてください。\n'
    + '令和・和暦は西暦に変換してください。年が明記されていない月日は文脈から判断して西暦を補ってください。\n'
    + '以下のJSON配列形式のみで返してください（説明文・マークダウン不要）:\n'
    + '[{"title":"イベント名","datetime":"開催日時","location":"場所","description":"料金・定員・申込方法・問合せTELなど"}]';

  var payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'application/pdf', data: pdfBase64 } }
      ]
    }],
    generationConfig: { temperature: 0.1 }
  };

  Logger.log('広報PDF: Gemini APIに送信中...');
  var geminiRes = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + geminiKey,
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  if (geminiRes.getResponseCode() !== 200) {
    Logger.log('広報PDF: Gemini API失敗 HTTP ' + geminiRes.getResponseCode());
    Logger.log(geminiRes.getContentText().substring(0, 500));
    return;
  }

  // Step6: Geminiレスポンスを解析
  var geminiJson = JSON.parse(geminiRes.getContentText());
  var text = '';
  try {
    text = geminiJson.candidates[0].content.parts[0].text;
  } catch(e) {
    Logger.log('広報PDF: Geminiレスポンス解析失敗 ' + e);
    return;
  }

  // JSONブロックを抽出（```json ... ``` 形式にも対応）
  var jsonStr = text.replace(/```json\s*/i, '').replace(/```\s*$/m, '').trim();
  var jsonM = jsonStr.match(/\[[\s\S]*\]/);
  if (!jsonM) {
    Logger.log('広報PDF: JSON抽出失敗\n' + text.substring(0, 500));
    return;
  }

  var events;
  try {
    events = JSON.parse(jsonM[0]);
  } catch(e) {
    Logger.log('広報PDF: JSON解析エラー ' + e.message);
    return;
  }
  Logger.log('広報PDF: Geminiが ' + events.length + '件を抽出');

  // Step7: スプレッドシートに追加（タイトル+広報キーで重複チェック）
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(INZAI_SHEET);
  var allRows = sheet.getDataRange().getValues();
  var existingKeys = {};
  for (var ri = 1; ri < allRows.length; ri++) {
    var existTitle = String(allRows[ri][1] || '').trim();
    if (existTitle) existingKeys[existTitle + '|kouhou'] = true;
  }

  var now = new Date();
  var cutoff = new Date();
  cutoff.setHours(0, 0, 0, 0);
  var addedCount = 0;

  events.forEach(function(ev) {
    if (!ev || !ev.title || ev.title.trim().length < 3) return;
    var title    = ev.title.trim();
    var datetime = (ev.datetime  || '').trim();
    var location = (ev.location  || '').trim();
    var desc     = (ev.description || '').trim();

    // 重複チェック
    var dedupKey = title + '|kouhou';
    if (existingKeys[dedupKey]) return;

    // 過去イベントを除外
    var parsedDate = parseCalDate_(datetime) || null;
    if (parsedDate && parsedDate < cutoff) return;

    // 年なし日付に西暦未記載フラグ
    var hasYear = /\d{4}年/.test(datetime);
    var storedDatetime = hasYear ? datetime : (datetime ? datetime + '　西暦未記載' : '');

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 9).setValues([[
      now, title, storedDatetime, location, articleUrl, desc, true, parsedDate || '', '広報いんざい'
    ]]);
    sheet.getRange(newRow, 7).insertCheckboxes();
    sheet.getRange(newRow, 7).setValue(true);
    sheet.getRange(newRow, 3).setNumberFormat('@STRING@');

    existingKeys[dedupKey] = true;
    addedCount++;
  });

  // Step8: 処理済みマーク（同じPDFを再処理しない）
  props.setProperty(processedKey, 'done_' + now.toISOString());
  Logger.log('広報PDF: ' + addedCount + '件追加完了（' + pdfUrl + '）');
}

// ============================================================
// 広報いんざい 日次トリガー設定
// このfunctionを一度だけ手動実行してください
// ============================================================
function setupKouhouTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'fetchKouhouInzaiEvents') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎日10:00に実行（1日・2日・15日・16日のみ処理）
  ScriptApp.newTrigger('fetchKouhouInzaiEvents')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();
  Logger.log('広報いんざいトリガー設定完了: 毎日10時実行（1日・15日のみ処理）');
}
