// instant.js — Prefetch on hover/touch to make links feel instant

(function () {
  const supportsLinkPrefetch = document.createElement('link').relList?.supports?.('prefetch');
  const prefetch = (url) => {
    try {
      if (supportsLinkPrefetch) {
        const link = document.createElement('link');
        link.rel = 'prefetch';
        link.href = url;
        link.as = 'document';
        document.head.appendChild(link);
      } else if ('requestIdleCallback' in window) {
        requestIdleCallback(() => fetch(url, { credentials: 'include' }).catch(() => {}));
      } else {
        setTimeout(() => fetch(url, { credentials: 'include' }).catch(() => {}), 0);
      }
    } catch (e) {}
  };

  const mark = new WeakSet();

  const handle = (e) => {
    const a = e.target.closest('a');
    if (!a) return;
    const href = a.getAttribute('href');
    if (!href || href.startsWith('mailto:') || href.startsWith('tel:') || href.startsWith('#')) return;
    if (mark.has(a)) return;
    mark.add(a);
    prefetch(href);
  };

  // Hover بالمؤشر ولمس الإصبع
  document.addEventListener('mouseover', handle, { passive: true, capture: true });
  document.addEventListener('touchstart', handle, { passive: true, capture: true });

  // Prefetch أول شاشة (Critical routes) — عدّل الروابط المهمة لديك هنا إن لزم
  window.addEventListener('load', () => {
    const critical = ['/', '/engines.html', '/generators.html'];
    critical.forEach((p) => prefetch(p));
  });
})();
