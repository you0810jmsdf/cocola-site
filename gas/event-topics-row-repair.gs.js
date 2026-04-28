/**
 * COCoLa_TOPICSイベント登録: TOPICS sheet row repair helpers.
 *
 * Emergency-only helper.
 * Prefer event-topics-normalized-submit.gs for the root fix.
 *
 * Usage in the bound Apps Script project:
 * 1. Paste this file into the Apps Script project for COCoLa_TOPICSイベント登録.
 * 2. Run repairTopicsRowsFrom59() once.
 * 3. Create an installable "On form submit" trigger for normalizeLatestTopicsRowOnSubmit.
 */

const TOPICS_REPAIR_CONFIG = {
  sheetName: 'TOPICS',
  firstBrokenRow: 59,
};

const TOPICS_COL = {
  timestamp: 1,
  date: 2,
  title: 3,
  place: 4,
  organizer: 5,
  imageUrl: 6,
  detailUrl: 7,
  summary: 8,
  tsukuruCategory: 9,
  participants: 10,
  publishTopics: 11,
  projectName: 12,
  publishAchievement: 13,
  editKey: 14,
};

const TSUKURU_CATEGORIES = [
  '自分をつくる',
  'ものをつくる',
  '場をつくる',
  'まちをつくる',
  'つながりをつくる',
  '未来をつくる',
  'その他',
];

function repairTopicsRowsFrom59() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOPICS_REPAIR_CONFIG.sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + TOPICS_REPAIR_CONFIG.sheetName);

  const startRow = TOPICS_REPAIR_CONFIG.firstBrokenRow;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const rowCount = lastRow - startRow + 1;
  const range = sheet.getRange(startRow, 1, rowCount, TOPICS_COL.editKey);
  const values = range.getValues();
  const repaired = values.map(normalizeTopicsRow_);
  range.setValues(repaired);
}

function normalizeLatestTopicsRowOnSubmit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOPICS_REPAIR_CONFIG.sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + TOPICS_REPAIR_CONFIG.sheetName);

  const row = sheet.getLastRow();
  if (row < TOPICS_REPAIR_CONFIG.firstBrokenRow) return;

  const range = sheet.getRange(row, 1, 1, TOPICS_COL.editKey);
  const value = range.getValues()[0];
  range.setValues([normalizeTopicsRow_(value)]);
}

function normalizeTopicsRow_(row) {
  const fixed = row.slice();

  const detailUrl = stringValue_(fixed[TOPICS_COL.detailUrl - 1]);
  const summary = stringValue_(fixed[TOPICS_COL.summary - 1]);
  const category = stringValue_(fixed[TOPICS_COL.tsukuruCategory - 1]);

  // Rows submitted without an image can shift detail URL, summary, and category one column left.
  if (!category && isTsukuruCategory_(summary)) {
    fixed[TOPICS_COL.tsukuruCategory - 1] = summary;
    fixed[TOPICS_COL.summary - 1] = detailUrl;

    const imageOrUrl = stringValue_(fixed[TOPICS_COL.imageUrl - 1]);
    if (looksLikeUrl_(imageOrUrl) && !looksLikeDriveImageUrl_(imageOrUrl)) {
      fixed[TOPICS_COL.detailUrl - 1] = imageOrUrl;
      fixed[TOPICS_COL.imageUrl - 1] = '';
    }
  }

  const participants = stringValue_(fixed[TOPICS_COL.participants - 1]);
  const projectName = stringValue_(fixed[TOPICS_COL.projectName - 1]);

  // In the broken rows, participant names landed in the project-name column.
  if (!participants && projectName && looksLikeParticipantList_(projectName)) {
    fixed[TOPICS_COL.participants - 1] = projectName;
    fixed[TOPICS_COL.projectName - 1] = '';
  }

  if (!stringValue_(fixed[TOPICS_COL.projectName - 1])) {
    fixed[TOPICS_COL.projectName - 1] = inferProjectName_(fixed);
  }

  if (stringValue_(fixed[TOPICS_COL.participants - 1]) && isBlankCheckbox_(fixed[TOPICS_COL.publishAchievement - 1])) {
    fixed[TOPICS_COL.publishAchievement - 1] = true;
  }

  return fixed;
}

function inferProjectName_(row) {
  const organizer = stringValue_(row[TOPICS_COL.organizer - 1]);
  const title = stringValue_(row[TOPICS_COL.title - 1]);
  return organizer || title;
}

function isTsukuruCategory_(value) {
  return TSUKURU_CATEGORIES.indexOf(stringValue_(value)) !== -1;
}

function looksLikeParticipantList_(value) {
  const text = stringValue_(value);
  if (!text) return false;
  if (isTsukuruCategory_(text)) return false;
  if (looksLikeUrl_(text)) return false;
  return text.length <= 80;
}

function looksLikeUrl_(value) {
  return /^https?:\/\//i.test(stringValue_(value));
}

function looksLikeDriveImageUrl_(value) {
  const text = stringValue_(value);
  return /drive\.google\.com/i.test(text) || /googleusercontent\.com/i.test(text);
}

function isBlankCheckbox_(value) {
  return value === '' || value === null || value === undefined || value === false;
}

function stringValue_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}
