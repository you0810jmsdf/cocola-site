/**
 * COCoLa メンバー紹介 Webアプリ
 * 表示は html.html、データ配信とフォーム更新処理はこのファイルで担当する。
 */

const CONFIG = {
  sheetName: 'フォームの回答 1',
  colPurpose: 3,
  colNickname: 4,
  updateLabel: '情報の修正・更新',
  skipEmpty: true,
  adminEmail: 'you0810jmsdf@gmail.com',
  formUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSfZYSm7QglMdTIQjfRFzbTuRWxnMlXJfpby3uHEAGTQ-6lc3A/viewform',
};

function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'json') {
    const callback = e.parameter.callback;
    const json = JSON.stringify(getMembers());
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  const template = HtmlService.createTemplateFromFile('html');
  template.APP_URL = ScriptApp.getService().getUrl();
  template.FORM_URL = CONFIG.formUrl;
  template.HOME_URL = 'https://you0810jmsdf.github.io/cocola-site/index.html';

  return template.evaluate()
    .setTitle('COCoLa メンバー紹介')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  if (!sheet) return [];

  return sheet.getDataRange().getValues()
    .slice(1)
    .filter((row) => {
      const purpose = String(row[2] || '').trim();
      const nickname = String(row[3] || '').trim();
      return purpose !== CONFIG.updateLabel && nickname !== '';
    })
    .map((row) => row.map((cell) => {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      }
      return cell === null || cell === undefined ? '' : String(cell);
    }));
}

function onFormSubmit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    Logger.log('シートが見つかりません: ' + CONFIG.sheetName);
    return;
  }

  const allData = sheet.getDataRange().getValues();
  const newRow = allData.length;
  const newData = allData[newRow - 1];

  const purpose = String(newData[CONFIG.colPurpose - 1] || '').trim();
  const nickname = String(newData[CONFIG.colNickname - 1] || '').trim();

  Logger.log('回答種別: ' + purpose + ' / ニックネーム: ' + nickname);

  if (purpose !== CONFIG.updateLabel) {
    Logger.log('新規登録のため更新処理をスキップ');
    return;
  }

  if (!nickname) {
    Logger.log('ニックネーム未入力のため更新を中断');
    addNote(sheet, newRow, '※ニックネーム未入力のため更新できません。');
    return;
  }

  let targetRow = -1;
  for (let i = 1; i < allData.length - 1; i += 1) {
    const existingNickname = String(allData[i][CONFIG.colNickname - 1] || '').trim();
    if (existingNickname === nickname) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    Logger.log('更新対象のニックネームが見つかりません: ' + nickname);
    addNote(sheet, newRow, '※ニックネーム「' + nickname + '」が見つからないため更新できません。');
    return;
  }

  for (let col = 1; col <= newData.length; col += 1) {
    if (col <= 2 || col === CONFIG.colPurpose) continue;
    const newValue = newData[col - 1];
    const isEmpty = newValue === null || newValue === '' || newValue === undefined;
    if (CONFIG.skipEmpty && isEmpty) continue;
    sheet.getRange(targetRow, col).setValue(newValue);
  }

  sheet.getRange(targetRow, 1).setValue(new Date());
  addNote(
    sheet,
    newRow,
    '✓ ' + targetRow + '行目に更新済み (' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') + ')'
  );

  Logger.log('更新完了: ' + targetRow + '行目');
  sendNotification(nickname, targetRow);
}

function addNote(sheet, row, note) {
  sheet.getRange(row, 1).setNote(note);
}

function sendNotification(nickname, rowIndex) {
  try {
    MailApp.sendEmail(
      CONFIG.adminEmail,
      '[COCoLa] メンバー情報が更新されました',
      'COCoLaメンバー情報が更新されました。\n\n'
      + 'ニックネーム: ' + nickname + '\n'
      + '更新先: スプレッドシート ' + rowIndex + ' 行目\n\n'
      + '内容を確認してください。'
    );
  } catch (error) {
    Logger.log('メール通知エラー: ' + error.message);
  }
}
