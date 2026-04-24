(function () {
  var s = document.currentScript;
  var page = s ? s.getAttribute('data-page') : '';
  var base = window.COCOLA_TOPICS_WEBAPP_URL;
  if (!page || !base) return;

  var VID_KEY = 'cocola_vid';
  var vid = localStorage.getItem(VID_KEY);
  if (!vid) {
    vid = 'v' + Math.random().toString(36).slice(2, 10) + Math.random().toString(36).slice(2, 10);
    localStorage.setItem(VID_KEY, vid);
  }

  fetch(base + '?mode=pageview&p=' + encodeURIComponent(page) + '&vid=' + encodeURIComponent(vid), {
    mode: 'no-cors',
    cache: 'no-store'
  }).catch(function () {});
})();
