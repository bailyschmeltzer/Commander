(function () {
  const user = String(localStorage.getItem('commanderTrackerSyncUser') || '').trim();
  const token = String(localStorage.getItem('commanderTrackerSyncToken') || '').trim();
  const adminLinks = Array.from(document.querySelectorAll('.page-link[href$="admin-logs.html"]'));

  if (!user || !token || !adminLinks.length) {
    return;
  }

  fetch('/api/session', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ user, token }),
  })
    .then((response) => response.ok ? response.json() : null)
    .then((payload) => {
      if (String(payload?.auth?.role || '').toLowerCase() !== 'admin') {
        return;
      }
      adminLinks.forEach((link) => {
        link.classList.add('admin-link-visible');
        link.hidden = false;
        link.setAttribute('aria-hidden', 'false');
        link.removeAttribute('tabindex');
      });
    })
    .catch(() => null);
}());