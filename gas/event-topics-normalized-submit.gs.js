/**
 * COCoLa TOPICS event form normalizer.
 *
 * Root fix:
 * - Google Forms writes raw responses to "フォームの回答 1".
 * - This script maps raw response headers to fixed TOPICS columns by header name.
 * - TOPICS column order stays stable for public pages and achievement aggregation.
 */

const TOPICS_NORMALIZE_CONFIG = {
  responseSheetName: 'フォームの回答 1',
  topicsSheetName: 'TOPICS',
  editKeyPrefix: 'evt',
};

const TOPICS_HEADERS = [
  'タイムスタンプ',
  '開催日時',
  'イベント名',
  '場所',
  '主催者',
  '画像URL',
  '詳細URL',
  '概要',
  '6つのつくる分類',
  '参加者',
  'TOPICS掲載',
  'プロジェクト名',
  '実績反映',
  '編集キー',
];

const FORM_FIELD_ALIASES = {
  mode: ['処理種別', '申請種別', '登録種別', '回答種別', '新規・修正', '新規登録・修正'],
  timestamp: ['タイムスタンプ', 'Timestamp'],
  date: ['開催日時', '日時', 'イベント日時', '日付', '開催日'],
  title: ['イベント名', 'タイトル', 'イベントタイトル'],
  place: ['場所', '会場', '開催場所'],
  organizer: ['主催者', '主催', '主催団体', '主催者名'],
  imageUrl: ['画像URL', '画像', 'チラシ画像', '画像リンク', 'チラシURL', 'ポスター画像のURL', 'ポスター画像URL'],
  detailUrl: ['詳細URL', '申込URL', 'リンク', 'イベントURL', '詳細リンク', '申込・詳細URL', '申込・詳細リンク（任意）', '申込・詳細リンク'],
  summary: ['概要', '説明', 'イベント概要', '内容', '紹介文'],
  tsukuruCategory: ['6つのつくる分類', 'つくる分類', '分類', 'カテゴリ', 'カテゴリー'],
  participants: ['参加者', '参加メンバー', 'COCoLa参加者', '関わる人'],
  publishTopics: ['TOPICS掲載', 'TOPICSに掲載', '掲載する', '公開する'],
  projectName: ['プロジェクト名', 'プロジェクト', '実績プロジェクト名'],
  publishAchievement: ['実績反映', '実績として反映', '実績に反映', '6つのつくるページの実績として反映'],
  editKey: ['編集キー', 'editKey'],
};

function onTopicsFormSubmitNormalized(e) {
  // DAOフォーム（提案・投票）の送信は別シートに書かれるためスキップ
  if (e && e.range) {
    const submittedSheet = e.range.getSheet().getName();
    if (submittedSheet !== TOPICS_NORMALIZE_CONFIG.responseSheetName) {
      Logger.log('onTopicsFormSubmitNormalized: スキップ (シート=' + submittedSheet + ')');
      return;
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = spreadsheet.getSheetByName(TOPICS_NORMALIZE_CONFIG.responseSheetName);
  if (!responseSheet) throw new Error('シートが見つかりません: ' + TOPICS_NORMALIZE_CONFIG.responseSheetName);

  const topicsSheet = ensureTopicsSheet_();
  const responseRow = getSubmittedResponseRow_(e, responseSheet);
  const responseObject = getResponseObject_(responseSheet, responseRow);
  upsertTopicsRow_(topicsSheet, responseObject);
}

/**
 * TOPICS_誤混入_backup の全行を TOPICSシートに復元する（緊急用）。
 * GASエディタから手動実行してください。
 */
function restoreTopicsFromBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = 'TOPICS_誤混入_backup';
  const backupSheet = ss.getSheetByName(backupName);
  if (!backupSheet) {
    Logger.log('バックアップシートが見つかりません: ' + backupName);
    return;
  }

  const lastRow = backupSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('バックアップシートにデータがありません');
    return;
  }

  const numCols = backupSheet.getLastColumn();
  const data = backupSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const topicsSheet = ss.getSheetByName(TOPICS_NORMALIZE_CONFIG.topicsSheetName);
  if (!topicsSheet) {
    Logger.log('TOPICSシートが見つかりません');
    return;
  }

  topicsSheet.getRange(topicsSheet.getLastRow() + 1, 1, data.length, numCols).setValues(data);
  Logger.log('復元完了: ' + data.length + '行をTOPICSシートに追加しました');
}

/**
 * TOPICSシートに誤混入したDAO提案・投票行をバックアップして削除する。
 * GASエディタから手動実行してください。実行後ログで結果を確認のこと。
 *
 * 判定基準（いずれかに該当する行を誤混入とみなす）:
 *   1. 編集キーが evt-[16桁英数字] 形式でない
 *   2. イベント名・開催日時が両方空（イベントの実態がない）
 */
function cleanupDaoRowsFromTopics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const topicsSheet = ss.getSheetByName(TOPICS_NORMALIZE_CONFIG.topicsSheetName);
  if (!topicsSheet) {
    Logger.log('TOPICSシートが見つかりません');
    return;
  }

  const lastRow = topicsSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('TOPICSシートにデータ行がありません');
    return;
  }

  const numCols = Math.max(topicsSheet.getLastColumn(), TOPICS_HEADERS.length);
  const allData = topicsSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const editKeyPattern = /^evt-[a-f0-9]{16}$/;

  const daoRowIndices = [];
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const date     = String(row[1]  || '').trim();
    const title    = String(row[2]  || '').trim();
    const editKey  = String(row[13] || '').trim();
    const hasEventInfo   = title !== '' || date !== '';
    const hasValidKey    = editKeyPattern.test(editKey);

    if (!hasEventInfo || !hasValidKey) {
      daoRowIndices.push(i + 2); // 1始まりのシート行番号
    }
  }

  if (daoRowIndices.length === 0) {
    Logger.log('誤混入行は見つかりませんでした');
    return;
  }

  // バックアップシートに退避
  const backupName = 'TOPICS_誤混入_backup';
  let backupSheet = ss.getSheetByName(backupName);
  if (!backupSheet) {
    backupSheet = ss.insertSheet(backupName);
    backupSheet.getRange(1, 1, 1, TOPICS_HEADERS.length).setValues([TOPICS_HEADERS]);
  }
  const backupData = daoRowIndices.map(function(rowIdx) {
    return allData[rowIdx - 2].slice(0, numCols);
  });
  const backupStart = backupSheet.getLastRow() + 1;
  backupSheet.getRange(backupStart, 1, backupData.length, numCols).setValues(backupData);

  // 下から順に削除（行番号ずれ防止）
  daoRowIndices.slice().reverse().forEach(function(rowIdx) {
    topicsSheet.deleteRow(rowIdx);
  });

  Logger.log('完了: ' + daoRowIndices.length + '行を退避・削除しました → ' + backupName);
  Logger.log('対象行(元の行番号): ' + JSON.stringify(daoRowIndices));
  backupData.forEach(function(r, i) {
    Logger.log('  行' + daoRowIndices[i] + ': editKey=' + r[13] + ' title=' + r[2] + ' date=' + r[1]);
  });
}

/**
 * 指定タイトルの行をTOPICSシートから削除する（ピンポイント削除用）。
 * GASエディタから手動実行してください。
 */
function deleteTopicsRowsByTitle() {
  var TITLES_TO_DELETE = ['焼き鳥交流会', '健活！カラダの歪み改善'];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TOPICS_NORMALIZE_CONFIG.topicsSheetName);
  if (!sheet) { Logger.log('TOPICSシートが見つかりません'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('データ行がありません'); return; }

  var titles = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // C列=イベント名
  var toDelete = [];
  for (var i = 0; i < titles.length; i++) {
    if (TITLES_TO_DELETE.indexOf(String(titles[i][0]).trim()) !== -1) {
      toDelete.push(i + 2);
    }
  }

  if (toDelete.length === 0) { Logger.log('対象行が見つかりませんでした'); return; }

  toDelete.slice().reverse().forEach(function(row) { sheet.deleteRow(row); });
  Logger.log('削除完了: ' + toDelete.length + '行 → 行番号: ' + JSON.stringify(toDelete));
}

function setupTopicsHeaders() {
  ensureTopicsSheet_();
}

function rebuildTopicsFromResponses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = spreadsheet.getSheetByName(TOPICS_NORMALIZE_CONFIG.responseSheetName);
  if (!responseSheet) throw new Error('シートが見つかりません: ' + TOPICS_NORMALIZE_CONFIG.responseSheetName);

  const topicsSheet = ensureTopicsSheet_();
  const lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return;

  const rows = [];
  for (let row = 2; row <= lastRow; row += 1) {
    const responseObj = getResponseObject_(responseSheet, row);
    const title = getByAliases_(responseObj, FORM_FIELD_ALIASES.title);
    const date  = getByAliases_(responseObj, FORM_FIELD_ALIASES.date);
    if (!title && !date) continue;
    rows.push(buildTopicsRow_(responseObj, null, false));
  }

  // 行3以降を削除（行1=ヘッダー・行2は残す必要があるため）
  const maxRows = topicsSheet.getMaxRows();
  if (maxRows > 2) {
    topicsSheet.deleteRows(3, maxRows - 2);
  }
  // 行2をコンテンツ・バリデーションごとクリア
  topicsSheet.getRange(2, 1, 1, topicsSheet.getMaxColumns()).clearContent();
  topicsSheet.getRange(2, 1, 1, topicsSheet.getMaxColumns()).clearDataValidations();

  // データ行数分だけ追加（行2を1行目として使うので rows.length-1 を挿入）
  if (rows.length > 1) {
    topicsSheet.insertRowsAfter(1, rows.length - 1);
  }
  topicsSheet.getRange(2, 1, rows.length, TOPICS_HEADERS.length).setValues(rows);

  // チェックボックスをデータ行のみに適用（全行適用だと行数誤検知の原因になる）
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  topicsSheet.getRange(2, 11, rows.length, 1).setDataValidation(checkboxRule);
  topicsSheet.getRange(2, 13, rows.length, 1).setDataValidation(checkboxRule);
}

function ensureTopicsSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(TOPICS_NORMALIZE_CONFIG.topicsSheetName)
    || spreadsheet.insertSheet(TOPICS_NORMALIZE_CONFIG.topicsSheetName);

  sheet.getRange(1, 1, 1, TOPICS_HEADERS.length).setValues([TOPICS_HEADERS]);
  sheet.setFrozenRows(1);
  applyTopicsCheckboxes_(sheet);
  return sheet;
}

function applyTopicsCheckboxes_(sheet) {
  // getLastRow() のみ使用。getMaxRows() は空行にもチェックボックス値を書き込み
  // getLastRow() を膨張させて appendRow の挿入位置がずれる原因になる。
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const dataRows = lastRow - 1;
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  sheet.getRange(2, 11, dataRows, 1).setDataValidation(checkboxRule);
  sheet.getRange(2, 13, dataRows, 1).setDataValidation(checkboxRule);
}

function applyTopicsCheckboxesToRow_(sheet, rowIndex) {
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(rowIndex, 11, 1, 1).setDataValidation(checkboxRule);
  sheet.getRange(rowIndex, 13, 1, 1).setDataValidation(checkboxRule);
}

function getSubmittedResponseRow_(e, responseSheet) {
  if (e && e.range && e.range.getSheet().getName() === responseSheet.getName()) {
    return e.range.getRow();
  }
  return responseSheet.getLastRow();
}

function getResponseObject_(sheet, row) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(normalizeHeader_);
  const values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];

  return headers.reduce((object, header, index) => {
    if (header) object[header] = values[index];
    return object;
  }, {});
}

function upsertTopicsRow_(topicsSheet, response) {
  const title = getByAliases_(response, FORM_FIELD_ALIASES.title);
  const date  = getByAliases_(response, FORM_FIELD_ALIASES.date);
  if (!title && !date) {
    Logger.log('upsertTopicsRow_: イベント名・日時が両方空のためスキップ');
    return;
  }

  const editKey = getByAliases_(response, FORM_FIELD_ALIASES.editKey);
  const targetRow = editKey ? findTopicsRowByEditKey_(topicsSheet, editKey) : -1;

  if (targetRow === -1) {
    topicsSheet.appendRow(buildTopicsRow_(response, null, false));
    applyTopicsCheckboxesToRow_(topicsSheet, topicsSheet.getLastRow());
    return;
  }

  const existing = topicsSheet.getRange(targetRow, 1, 1, TOPICS_HEADERS.length).getValues()[0];
  const updated = buildTopicsRow_(response, existing, true);
  topicsSheet.getRange(targetRow, 1, 1, TOPICS_HEADERS.length).setValues([updated]);
}

function buildTopicsRow_(response, existingRow, preserveBlankValues) {
  const value = (key) => getByAliases_(response, FORM_FIELD_ALIASES[key]);
  const existing = (index) => existingRow ? existingRow[index] : '';
  const pick = (key, index, fallback) => {
    const incoming = value(key);
    if (incoming !== '' && incoming !== null && incoming !== undefined) return incoming;
    if (preserveBlankValues) return existing(index);
    return fallback || '';
  };

  const timestamp = pick('timestamp', 0, new Date()) || new Date();
  const title = value('title');
  const organizer = value('organizer');
  const participants = pick('participants', 9, '');
  const projectName = pick('projectName', 11, '') || organizer || title;

  return [
    timestamp,
    pick('date', 1, ''),
    pick('title', 2, ''),
    pick('place', 3, ''),
    pick('organizer', 4, ''),
    pick('imageUrl', 5, ''),
    pick('detailUrl', 6, ''),
    pick('summary', 7, ''),
    pick('tsukuruCategory', 8, ''),
    participants,
    toCheckboxValue_(value('publishTopics'), preserveBlankValues ? existing(10) : true),
    projectName,
    toCheckboxValue_(value('publishAchievement'), preserveBlankValues ? existing(12) : Boolean(participants)),
    pick('editKey', 13, '') || createEditKey_(),
  ];
}

function findTopicsRowByEditKey_(sheet, editKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 14, lastRow - 1, 1).getValues();
  const normalizedEditKey = String(editKey).trim();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0] || '').trim() === normalizedEditKey) {
      return i + 2;
    }
  }
  return -1;
}

function getByAliases_(response, aliases) {
  for (let i = 0; i < aliases.length; i += 1) {
    const key = normalizeHeader_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(response, key)) {
      const value = response[key];
      if (value !== '' && value !== null && value !== undefined) return value;
    }
  }
  return '';
}

function toCheckboxValue_(value, defaultValue) {
  if (value === '' || value === null || value === undefined) return defaultValue;
  if (value === true || value === false) return value;
  const text = String(value).trim().toLowerCase();
  return ['true', 'yes', 'y', '1', 'する', 'はい', '掲載する', '反映する'].indexOf(text) !== -1;
}

function createEditKey_() {
  return TOPICS_NORMALIZE_CONFIG.editKeyPrefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[「」『』【】［］\[\]（）()]/g, '')
    .trim();
}
