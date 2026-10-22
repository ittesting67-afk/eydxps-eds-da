const DATE_RE = /^(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|\d{4}-\d{2}-\d{2}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})$/;

function shareLinks(title) {
  const url = window.location.href.split('#')[0];
  const wrap = document.createElement('div');
  wrap.className = 'ey-article-share';

  const mail = document.createElement('a');
  mail.className = 'ey-share-mail';
  mail.title = 'Send Email';
  mail.href = `mailto:?subject=${encodeURIComponent(`Discover - ${title}`)}&body=${encodeURIComponent(`Open: < ${url} >`)}`;
  mail.textContent = 'Share';

  const teams = document.createElement('a');
  teams.className = 'ey-share-teams';
  teams.title = 'Share to Teams';
  teams.target = '_blank';
  teams.rel = 'noopener';
  teams.href = `https://teams.microsoft.com/share?href=${encodeURIComponent(url)}&msgText=${encodeURIComponent(title)}`;
  teams.textContent = 'Share to Teams';

  wrap.append(mail, teams);
  return wrap;
}

/**
 * ey-article-header - article/event/topic page header: parent breadcrumb link(s) with angle-right,
 * title, publication date (right aligned) and Share / Share to Teams links (generated).
 * Rows are classified by content: list = breadcrumb, heading = title, date-like text = date.
 * Variant `no-share` hides the share links.
 * @param {Element} block
 */
export default function decorate(block) {
  let crumbs;
  let titleEl;
  let dateText = '';
  const cells = [...block.children].map((row) => row.firstElementChild || row);

  cells.forEach((cell) => {
    const text = cell.textContent.trim();
    if (!text) return;
    if (!crumbs && cell.querySelector('ul, ol')) {
      crumbs = cell.querySelector('ul, ol');
    } else if (!titleEl && cell.querySelector('h1, h2, h3, h4, h5, h6')) {
      titleEl = cell.querySelector('h1, h2, h3, h4, h5, h6');
    } else if (!dateText && DATE_RE.test(text)) {
      dateText = text;
    } else if (!titleEl) {
      titleEl = document.createElement('h1');
      titleEl.textContent = text;
    }
  });

  block.textContent = '';

  if (crumbs) {
    const nav = document.createElement('nav');
    nav.className = 'ey-article-crumbs';
    nav.setAttribute('aria-label', 'Breadcrumb');
    [...crumbs.querySelectorAll(':scope > li')].forEach((li) => {
      const a = li.querySelector('a');
      if (a) nav.append(a.cloneNode(true));
      else {
        const span = document.createElement('span');
        span.textContent = li.textContent.trim();
        nav.append(span);
      }
    });
    block.append(nav);
  }

  const head = document.createElement('div');
  head.className = 'ey-article-head';
  const title = titleEl || document.createElement('h1');
  if (!titleEl) title.textContent = document.title.split('|')[0].trim();
  head.append(title);
  if (dateText) {
    const date = document.createElement('time');
    date.className = 'ey-article-date';
    date.textContent = dateText;
    const parsed = new Date(dateText);
    if (!Number.isNaN(parsed.getTime())) date.dateTime = parsed.toISOString().slice(0, 10);
    head.append(date);
  }
  block.append(head);

  if (!block.classList.contains('no-share')) {
    block.append(shareLinks(title.textContent.trim()));
  }
}
