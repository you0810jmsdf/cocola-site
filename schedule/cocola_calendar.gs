// ============================================
// COCoLa schedule Web App
// 仕様: schedule/HANDOFF.md
// ============================================

var CALENDAR_ID = 'cocola.project@gmail.com';
var DEFAULT_DURATION_MIN = 60;
var LIST_RANGE_DAYS = 30;
var TIMEZONE = 'Asia/Tokyo';

function getAdminPass_() {
  var pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS');
  if (!pass) throw new Error('ADMIN_PASS is not set in script properties');
  return pass;
}

// ============================================
// エントリーポイント
// ============================================
function doGet(e) {
  var mode = (e && e.parameter && e.parameter.mode) || '';

  if (mode === 'list') {
    return jsonResponse_({ ok: true, events: getUpcomingEvents_() });
  }

  var tpl = HtmlService.createTemplateFromFile('schedule');
  return tpl.evaluate()
    .setTitle('COCoLa スケジュール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = parsePayload_(e);
    if (payload.pass !== getAdminPass_()) {
      return jsonResponse_({ ok: false, error: 'unauthorized' });
    }
    var id = createEvent_(payload);
    return jsonResponse_({ ok: true, id: id });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message || err) });
  }
}

// ============================================
// HTML から呼ばれる関数 (google.script.run)
// ============================================
function createEventFromForm(payload) {
  try {
    if (!payload || payload.pass !== getAdminPass_()) {
      return { ok: false, error: 'unauthorized' };
    }
    var id = createEvent_(payload);
    return { ok: true, id: id };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function listUpcomingEvents() {
  return getUpcomingEvents_();
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ============================================
// カレンダー操作
// ============================================
function createEvent_(p) {
  if (!p.title) throw new Error('title is required');
  if (!p.start) throw new Error('start is required');

  var start = new Date(p.start);
  if (isNaN(start.getTime())) throw new Error('invalid start');

  var end = p.end ? new Date(p.end) : new Date(start.getTime() + DEFAULT_DURATION_MIN * 60 * 1000);
  if (isNaN(end.getTime())) throw new Error('invalid end');

  var description = p.description || '';
  if (p.url) {
    description = (description ? description + '\n\n' : '') + p.url;
  }

  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error('calendar not found: ' + CALENDAR_ID);

  var event = calendar.createEvent(p.title, start, end, {
    location: p.location || '',
    description: description
  });
  return event.getId();
}

function getUpcomingEvents_() {
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error('calendar not found: ' + CALENDAR_ID);

  var now = new Date();
  var until = new Date(now.getTime() + LIST_RANGE_DAYS * 24 * 60 * 60 * 1000);

  return calendar.getEvents(now, until).map(function (ev) {
    return {
      id: ev.getId(),
      title: ev.getTitle(),
      start: Utilities.formatDate(ev.getStartTime(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
      end: Utilities.formatDate(ev.getEndTime(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
      location: ev.getLocation() || '',
      description: ev.getDescription() || ''
    };
  });
}

// ============================================
// 内部ユーティリティ
// ============================================
function parsePayload_(e) {
  if (e && e.postData && e.postData.type === 'application/json') {
    return JSON.parse(e.postData.contents);
  }
  return (e && e.parameter) || {};
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
