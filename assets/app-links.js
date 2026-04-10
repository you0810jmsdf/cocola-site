(function () {
  var links = window.COCOLA_APP_LINKS || {};
  var items = document.querySelectorAll('[data-app-link]');
  if (!items.length) return;

  items.forEach(function (el) {
    var key = el.getAttribute('data-app-link');
    var url = (links[key] || '').trim();
    var missing = el.getAttribute('data-app-missing');

    if (url) {
      if (el.tagName === 'A') {
        el.href = url;
      }
      el.hidden = false;
    } else if (missing === 'hide') {
      el.hidden = true;
    } else {
      el.hidden = false;
      if (el.tagName === 'A') {
        el.removeAttribute('href');
        el.setAttribute('aria-disabled', 'true');
      }
    }
  });
})();
