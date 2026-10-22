/**
 * ey-banner - EY Discover site/page banner: hero image with dark gradient overlay, optional
 * breadcrumb ("Sectors > Health") and a white title. Content model (rows are classified by content,
 * order is free): a row with a picture = background image; a row with a list = breadcrumb;
 * any other row = title (heading keeps its level, plain text/link renders as a div).
 * @param {Element} block
 */
export default function decorate(block) {
  let picture;
  let crumbs;
  let titleEl;

  [...block.children].forEach((row) => {
    const cell = row.firstElementChild || row;
    if (!picture && cell.querySelector('picture, img')) {
      picture = cell.querySelector('picture') || cell.querySelector('img');
    } else if (!crumbs && cell.querySelector('ul, ol')) {
      crumbs = cell.querySelector('ul, ol');
    } else if (!titleEl && cell.textContent.trim()) {
      titleEl = cell.querySelector('h1, h2, h3, h4, h5, h6');
      if (!titleEl) {
        titleEl = document.createElement('div');
        titleEl.innerHTML = cell.querySelector('p') ? cell.querySelector('p').innerHTML : cell.innerHTML;
      }
    }
  });

  block.textContent = '';

  if (picture) {
    const media = document.createElement('div');
    media.className = 'ey-banner-media';
    media.append(picture);
    const img = picture.tagName === 'IMG' ? picture : picture.querySelector('img');
    if (img) {
      img.loading = 'eager';
      img.setAttribute('fetchpriority', 'high');
      if (!img.alt) img.alt = '';
    }
    block.append(media);
  }

  const overlay = document.createElement('div');
  overlay.className = 'ey-banner-overlay';
  block.append(overlay);

  const content = document.createElement('div');
  content.className = 'ey-banner-content';
  const inner = document.createElement('div');
  inner.className = 'ey-banner-inner';

  if (crumbs) {
    const nav = document.createElement('nav');
    nav.className = 'ey-banner-crumbs';
    nav.setAttribute('aria-label', 'Breadcrumb');
    const ol = document.createElement('ol');
    [...crumbs.children].forEach((li) => {
      const item = document.createElement('li');
      item.innerHTML = li.innerHTML;
      ol.append(item);
    });
    nav.append(ol);
    inner.append(nav);
  }

  if (titleEl) {
    titleEl.classList.add('ey-banner-title');
    inner.append(titleEl);
  }

  content.append(inner);
  block.append(content);
}
