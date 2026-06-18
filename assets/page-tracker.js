(function () {
  var s = document.currentScript;
  var page = s ? s.getAttribute('data-page') : '';
  var counterBase = window.COCOLA_ACCESS_COUNTER_URL;
  if (!page || !counterBase) return;

  var VID_KEY = 'cocola_vid';
  var vid = localStorage.getItem(VID_KEY);
  if (!vid) {
    vid = 'v' + Math.random().toString(36).slice(2, 10) + Math.random().toString(36).slice(2, 10);
    localStorage.setItem(VID_KEY, vid);
  }

  function addBadge(count) {
    if (document.getElementById('cocola-access-counter')) return;
    var badge = document.createElement('div');
    badge.id = 'cocola-access-counter';
    badge.textContent = 'アクセス ' + Number(count || 0).toLocaleString('ja-JP');
    var style = badge.style;
    style.position = 'fixed';
    style.left = '50%';
    style.transform = 'translateX(-50%)';
    style.bottom = '8px';
    style.zIndex = '900';
    style.padding = '4px 10px';
    style.border = '1px solid rgba(120, 86, 60, .22)';
    style.borderRadius = '999px';
    style.background = 'rgba(255, 255, 255, .82)';
    style.color = '#6f4e37';
    style.fontSize = '11px';
    style.fontWeight = '600';
    style.letterSpacing = '0';
    style.boxShadow = '0 2px 8px rgba(0,0,0,.06)';
    style.backdropFilter = 'blur(6px)';
    style.pointerEvents = 'none';
    document.body.appendChild(badge);
  }

  var url = counterBase
    + '?mode=track'
    + '&page=' + encodeURIComponent(page)
    + '&title=' + encodeURIComponent(document.title || '')
    + '&vid=' + encodeURIComponent(vid);

  fetch(url, { cache: 'no-store' })
    .then(function (res) { return res.ok ? res.json() : null; })
    .then(function (data) {
      if (data && data.ok && typeof data.count !== 'undefined') {
        if (document.readyState === 'loading') {
          document.addEventListener('DOMContentLoaded', function () { addBadge(data.count); });
        } else {
          addBadge(data.count);
        }
      }
    })
    .catch(function () { /* カウンタは装飾。失敗してもページ機能は止めない */ });
})();
