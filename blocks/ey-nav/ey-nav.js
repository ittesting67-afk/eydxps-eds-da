import { getMetadata } from '../../scripts/aem.js';

/** normalise a path for comparison: lowercase, no origin, no .html/.aspx, no trailing slash, no /index */
function norm(p) {
  return (p || '')
    .toLowerCase()
    .replace(/[?#].*$/, '')
    .replace(/\.(html|aspx)$/, '')
    .replace(/\/index$/, '')
    .replace(/\/+$/, '');
}

function currentPath() {
  // window.eyPagePath is an optional override (used by the local preview harness)
  return norm(window.eyPagePath || window.location.pathname);
}

/**
 * ey-nav - horizontal site tab bar (SharePoint "eySecNav"). Content model: one list of links;
 * the link whose text is "Home" renders as a house icon. The active tab is the link whose path
 * equals the current page, or the one named by the page metadata `nav-tab` (link text or path slug).
 * @param {Element} block
 */
export default function decorate(block) {
  const list = block.querySelector('ul, ol');
  if (!list) return;
  const here = currentPath();
  const tabMeta = (getMetadata('nav-tab') || '').trim().toLowerCase();

  const nav = document.createElement('nav');
  nav.className = 'ey-nav-bar';
  nav.setAttribute('aria-label', 'Site navigation');
  const ul = document.createElement('ul');
  let active;

  [...list.children].forEach((li) => {
    const a = li.querySelector('a');
    const item = document.createElement('li');
    const label = li.textContent.trim();
    if (a) {
      const link = a.cloneNode(true);
      item.append(link);
      let path = '';
      try { path = norm(new URL(a.href, window.location.href).pathname); } catch (e) { /* ignore */ }
      if (/^home$/i.test(label)) {
        item.classList.add('ey-nav-home');
        link.setAttribute('aria-label', 'Home');
      }
      const slug = path.split('/').pop();
      const byMeta = tabMeta && (tabMeta === label.toLowerCase() || tabMeta === slug || tabMeta === path);
      if (!active && ((!tabMeta && path && path === here) || byMeta)) active = item;
    } else {
      item.textContent = label;
    }
    ul.append(item);
  });

  if (active) {
    active.classList.add('active');
    active.firstElementChild?.setAttribute('aria-current', 'page');
  }
  nav.append(ul);
  block.textContent = '';
  block.append(nav);

  // keep the active tab visible when the bar scrolls horizontally (mobile)
  if (active) {
    requestAnimationFrame(() => {
      const target = active.offsetLeft - (ul.clientWidth - active.offsetWidth) / 2;
      if (ul.scrollWidth > ul.clientWidth) ul.scrollLeft = Math.max(0, target);
    });
  }
}
