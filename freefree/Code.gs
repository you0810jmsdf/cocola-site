/**
 * Code.gs - エントリーポイント
 *
 * URL例:
 *   ?page=index     公開掲示板 (デフォルト)
 *   ?page=mypage    事業主マイページ
 *   ?page=signup    新規登録
 *
 * POSTは doPost で JSON API として処理
 */

function doGet(e) {
  try {
    const page = (e && e.parameter && e.parameter.page) || 'index';

    // JSON API モード: ?page=api
    if (page === 'api') {
      const type = (e.parameter && e.parameter.type) || '';
      if (type === 'businesses') {
        const data = getPublicBusinesses_();
        return ContentService.createTextOutput(JSON.stringify({ ok: true, businesses: data }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const filter = {
        category: (e.parameter && e.parameter.category) || '',
        area: (e.parameter && e.parameter.area) || ''
      };
      const data = getPublicPosters_(filter);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, posters: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let templateName;
    switch (page) {
      case 'mypage': templateName = 'mypage'; break;
      case 'signup': templateName = 'signup'; break;
      case 'help':   templateName = 'help'; break;
      default:       templateName = 'index';
    }
    const tpl = HtmlService.createTemplateFromFile(templateName);
    const siteName = getSetting_('SITE_NAME') || '印西市地域掲示板';
    tpl.siteName = siteName;
    tpl.userEmail = '';
    tpl.isAdmin = false;
    tpl.webAppUrl = '';
    try { tpl.userEmail = Session.getActiveUser().getEmail() || ''; } catch(e2) {}
    try { tpl.isAdmin = tpl.userEmail === getSetting_('ADMIN_EMAIL'); } catch(e2) {}
    try { tpl.webAppUrl = ScriptApp.getService().getUrl(); } catch(e2) {}
    try { tpl.categoriesJson = JSON.stringify(String(getSetting_('CATEGORIES')).split(',').map(s => s.trim())); } catch(e2) { tpl.categoriesJson = '[]'; }
    try { tpl.areasJson = JSON.stringify(String(getSetting_('AREAS')).split(',').map(s => s.trim())); } catch(e2) { tpl.areasJson = '[]'; }
    try {
      const _areas = String(getSetting_('AREAS')).split(',').map(s => s.trim());
      tpl.areaOptionsHtml = _areas.map(a => '<option value="' + a + '">' + a + '</option>').join('');
      const _cats = String(getSetting_('CATEGORIES')).split(',').map(s => s.trim());
      tpl.categoryOptionsHtml = _cats.map(c => '<option value="' + c + '">' + c + '</option>').join('');
    } catch(e2) { tpl.areaOptionsHtml = ''; tpl.categoryOptionsHtml = ''; }
    return tpl.evaluate()
      .setTitle(siteName)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    return HtmlService.createHtmlOutput(
      '<h2 style="color:red;font-family:sans-serif">エラー</h2><pre>' + err.toString() + '</pre>'
    );
  }
}

/** HTMLファイルのインクルード (<?!= include('style') ?>) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * クライアント側から呼ばれる関数 (google.script.run で呼び出し)
 * 認証は各関数内で実施
 */

// ========== 公開画面用 ==========

function api_getPublicPosters(filter) {
  return getPublicPosters_(filter);
}

function api_logEvent(event) {
  return logEvent_(event);
}

// ========== 事業主用 ==========

function api_getMe(email) {
  if (!email) return { loggedIn: false, registered: false };
  const biz = getMyBusiness_(email);
  if (!biz) return { loggedIn: true, registered: false, email: email };
  try {
    const posters = sheetToObjects_('Posters').filter(p => p.business_id === biz.business_id && p.status !== '終了');
    const sanitize_ = obj => JSON.parse(JSON.stringify(obj, (k, v) => v instanceof Date ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') : v));
    return { loggedIn: true, registered: true, business: sanitize_(biz), posters: sanitize_(posters) };
  } catch(e) {
    return { loggedIn: true, registered: true, business: String(biz.business_id), posters: [], posterError: e.toString() };
  }
}

function api_registerBusiness(data) {
  return registerBusiness_(data);
}

function api_updateBusiness(data) {
  return updateBusiness_(data);
}

function api_savePoster(data) {
  return savePoster_(data);
}

function api_deletePoster(posterId, email) {
  return deletePoster_(posterId, email);
}

function api_getMyStats(email) {
  return getMyStats_(email);
}

function api_getPublicBusinesses() {
  return getPublicBusinesses_();
}

function api_getSettings() {
  return {
    categories: String(getSetting_('CATEGORIES')).split(',').map(s => s.trim()),
    areas: String(getSetting_('AREAS')).split(',').map(s => s.trim()),
    siteName: getSetting_('SITE_NAME'),
    freePhase: String(getSetting_('FREE_PHASE')) === 'true'
  };
}
