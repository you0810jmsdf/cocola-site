(function () {
  var s = document.currentScript;
  var page = s ? s.getAttribute('data-page') : '';
  var base = window.COCOLA_TOPICS_WEBAPP_URL;
  if (!page || !base) return;
  fetch(base + '?mode=pageview&p=' + encodeURIComponent(page), {
    mode: 'no-cors',
    cache: 'no-store'
  }).catch(function () {});
})();
