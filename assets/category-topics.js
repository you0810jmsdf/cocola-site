(function () {
  var baseUrl = (window.COCOLA_TOPICS_WEBAPP_URL || '').trim();
  var embeds = document.querySelectorAll('[data-topics-embed]');
  if (!embeds.length) return;

  embeds.forEach(function (root) {
    var catKey = root.getAttribute('data-cat-key');
    var mode = root.getAttribute('data-topics-mode') || 'category';
    var frame = root.querySelector('[data-topics-frame]');
    var pageLink = root.querySelector('[data-topics-link]');
    var emptyNote = root.querySelector('[data-topics-empty]');

    if (!pageLink || !emptyNote) return;

    if (!baseUrl) {
      if (frame) frame.hidden = true;
      pageLink.hidden = true;
      emptyNote.hidden = false;
      return;
    }

    var url = baseUrl;
    if (mode === 'category' && catKey) {
      url += '?page=category&cat=' + encodeURIComponent(catKey);
    }

    if (frame) {
      frame.src = url;
      frame.hidden = false;
    }
    pageLink.href = url;
    pageLink.hidden = false;
    emptyNote.hidden = true;
  });
})();
