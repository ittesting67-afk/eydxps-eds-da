/*
 * ey-tiles - one flexible tile / card block for the EY Discover (SharePoint) look.
 *
 * Authoring table: first row = block name + variants, e.g. `ey-tiles (image outlined)`.
 * One row per tile. Cells are recognised by content, not by column position:
 *   - a cell holding a picture            -> tile image
 *   - 1st remaining cell                  -> title (optionally a link)
 *   - 2nd / 3rd remaining cell            -> description / meta (date, sector tag)
 *   - a row that holds ONLY a picture     -> shared background image of the whole block
 *     (variants box, panorama, flip)
 *   - cards only: a row with a single link cell and no picture -> "view more" button
 * Variant `reveal`: a tile that has a description cell does not link anywhere; clicking it
 * reveals the cell content as a panel below the tiles (one panel open at a time).
 *
 * Visual variants: cards, flip, buttons, image (+outlined), tab, bar, slide, overlay, box
 * (+labelbar), panorama, panels, media. Optional modifiers: reveal, cols-2 .. cols-6,
 * short, light (reveal panel in the lighter "tile-content" style).
 */

const WHOLE_LINK = ['flip', 'buttons', 'image', 'tab', 'bar', 'slide', 'overlay', 'box', 'panorama'];
const SHARED_BG = ['box', 'panorama', 'flip'];
const NO_PIC = ['buttons', 'panels', 'box', 'panorama', 'flip']; // pictures are ignored per tile
const ROLES = {
  cards: ['title', 'desc', 'meta'],
  overlay: ['title', 'tag'],
  default: ['title', 'desc', 'meta'],
};

let uid = 0;

const hasPic = (cell) => !!cell.querySelector('picture, img');
const textOf = (el) => (el ? el.textContent.replace(/\s+/g, ' ').trim() : '');

/** Split a DA row into {pic, texts[]}. */
function parseRow(row) {
  const cells = [...row.children];
  const pic = cells.map((c) => c.querySelector('picture') || c.querySelector('img')).find(Boolean) || null;
  const texts = cells.filter((c) => !hasPic(c) || textOf(c)).map((c) => {
    if (!hasPic(c)) return c;
    const copy = c.cloneNode(true);
    copy.querySelectorAll('picture, img').forEach((p) => p.remove());
    return copy;
  });
  // drop trailing empty cells but keep empty ones in the middle (they keep the column roles)
  while (texts.length && !textOf(texts[texts.length - 1]) && !texts[texts.length - 1].querySelector('a, ul, ol')) texts.pop();
  return { pic, texts };
}

function plainParagraphs(cell) {
  const wrap = document.createElement('div');
  wrap.className = 'ey-tile-desc';
  const ps = [...cell.querySelectorAll('p, li')];
  const parts = ps.length ? ps.map(textOf) : [textOf(cell)];
  parts.filter(Boolean).forEach((t) => {
    const p = document.createElement('p');
    p.textContent = t;
    wrap.append(p);
  });
  return wrap;
}

function makeFace(item, opts) {
  const {
    link, reveal, panelId, wholeLink,
  } = opts;
  let face;
  if (reveal) {
    face = document.createElement('button');
    face.type = 'button';
    face.setAttribute('aria-expanded', 'false');
    face.setAttribute('aria-controls', panelId);
  } else if (wholeLink && link) {
    face = document.createElement('a');
    face.href = link.getAttribute('href');
    ['target', 'rel', 'title'].forEach((a) => { if (link.getAttribute(a)) face.setAttribute(a, link.getAttribute(a)); });
    if (face.target === '_blank' && !face.rel) face.rel = 'noopener';
  } else {
    face = document.createElement('div');
  }
  face.classList.add('ey-tile-face');
  return face;
}

function buildTile(item, variant, reveal, panels) {
  const roles = ROLES[variant] || ROLES.default;
  const wholeLink = WHOLE_LINK.includes(variant);
  const { pic, texts } = item;
  const byRole = {};
  texts.forEach((cell, i) => { byRole[roles[i] || `extra${i}`] = cell; });

  const titleCell = byRole.title;
  const link = titleCell ? titleCell.querySelector('a') : null;
  const titleText = textOf(titleCell);
  const descCell = byRole.desc;
  const isReveal = reveal && !!descCell;
  const panelId = isReveal ? `ey-tile-panel-${(uid += 1)}` : '';

  const li = document.createElement('li');
  li.className = 'ey-tile';
  const face = makeFace(item, {
    link, reveal: isReveal, panelId, wholeLink,
  });
  li.append(face);

  if (pic && !NO_PIC.includes(variant)) {
    const pw = document.createElement('div');
    pw.className = 'ey-tile-pic';
    pw.append(pic.tagName === 'PICTURE' ? pic : (() => { const p = document.createElement('picture'); p.append(pic); return p; })());
    const img = pw.querySelector('img');
    if (img && !img.getAttribute('loading')) img.loading = 'lazy';
    face.append(pw);
  }

  if (byRole.tag) {
    const tag = document.createElement('span');
    tag.className = 'ey-tile-tag';
    tag.textContent = textOf(byRole.tag);
    face.append(tag);
  }

  if (titleText || link) {
    const label = document.createElement(wholeLink ? 'span' : 'h4');
    label.className = 'ey-tile-label';
    if (!wholeLink && link) {
      const a = link.cloneNode(false);
      a.removeAttribute('class'); // aem.js turns bold links into a.button; keep it a plain link
      a.textContent = titleText;
      label.append(a);
    } else {
      label.textContent = titleText;
    }
    face.append(label);
  }

  if (descCell && !isReveal) {
    if (wholeLink) face.append(plainParagraphs(descCell));
    else {
      const d = document.createElement('div');
      d.className = 'ey-tile-desc';
      d.append(...descCell.childNodes);
      face.append(d);
    }
  }
  if (variant === 'flip') {
    let d = face.querySelector('.ey-tile-desc');
    if (!d) { d = document.createElement('div'); d.className = 'ey-tile-desc'; face.append(d); }
    d.setAttribute('data-title', titleText);
  }
  if (byRole.meta) {
    const m = document.createElement('div');
    m.className = 'ey-tile-meta';
    m.textContent = textOf(byRole.meta);
    face.append(m);
  }

  if (variant === 'media') {
    const text = document.createElement('div');
    text.className = 'ey-tile-text';
    text.append(...face.querySelectorAll('.ey-tile-label, .ey-tile-desc, .ey-tile-meta'));
    face.append(text);
  }

  if (isReveal) {
    li.classList.add('is-reveal');
    const panel = document.createElement('section');
    panel.className = 'ey-tile-panel';
    panel.id = panelId;
    panel.hidden = true;
    const head = document.createElement('div');
    head.className = 'ey-tile-panel-head';
    const h = document.createElement('h4');
    h.textContent = titleText;
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'ey-tile-panel-close';
    close.setAttribute('aria-label', 'Close');
    close.textContent = '\u00d7';
    head.append(h, close);
    const body = document.createElement('div');
    body.className = 'ey-tile-panel-body';
    // every cell after the title is panel content
    texts.slice(1).forEach((c) => body.append(...c.childNodes));
    panel.append(head, body);
    panels.push({ panel, face, close });
  }
  return li;
}

function wireReveal(block, panels) {
  const closeAll = () => panels.forEach(({ panel, face }) => {
    panel.hidden = true;
    face.setAttribute('aria-expanded', 'false');
    face.classList.remove('active');
    face.closest('.ey-tile').classList.remove('active');
  });
  panels.forEach(({ panel, face, close }) => {
    face.addEventListener('click', () => {
      const wasOpen = !panel.hidden;
      closeAll();
      if (!wasOpen) {
        panel.hidden = false;
        face.setAttribute('aria-expanded', 'true');
        face.classList.add('active');
        face.closest('.ey-tile').classList.add('active');
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
    close.addEventListener('click', () => {
      closeAll();
      face.focus();
    });
  });
  block.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && panels.some((p) => !p.panel.hidden)) closeAll();
  });
}

export default function decorate(block) {
  const known = ['cards', 'flip', 'buttons', 'image', 'tab', 'bar', 'slide', 'overlay', 'box', 'panorama', 'panels', 'media'];
  let variant = known.find((v) => block.classList.contains(v));
  if (!variant) {
    variant = 'buttons';
    block.classList.add(variant);
  }
  const reveal = block.classList.contains('reveal');
  const rows = [...block.children];

  let bgPic = null;
  let more = null;
  const items = [];
  rows.forEach((row) => {
    const item = parseRow(row);
    if (item.pic && !item.texts.length && SHARED_BG.includes(variant) && !items.length && !bgPic) {
      bgPic = item.pic;
    } else if (variant === 'cards' && !item.pic && item.texts.length === 1 && item.texts[0].querySelector('a')) {
      more = item.texts[0].querySelector('a');
    } else if (item.pic || item.texts.length) {
      items.push(item);
    }
  });

  const list = document.createElement('ul');
  list.className = 'ey-tiles-list';
  const panels = [];
  items.forEach((it) => list.append(buildTile(it, variant, reveal, panels)));
  list.style.setProperty('--ey-tiles-count', items.length);

  block.textContent = '';
  if (bgPic) {
    const bg = document.createElement('div');
    bg.className = 'ey-tiles-bg';
    bg.append(bgPic.tagName === 'PICTURE' ? bgPic : (() => { const p = document.createElement('picture'); p.append(bgPic); return p; })());
    block.append(bg);
  }
  block.append(list);

  if (more) {
    const wrap = document.createElement('div');
    wrap.className = 'ey-tiles-more';
    const a = document.createElement('a');
    a.href = more.getAttribute('href');
    a.textContent = textOf(more);
    wrap.append(a);
    block.append(wrap);
  }

  if (panels.length) {
    const holder = document.createElement('div');
    holder.className = 'ey-tiles-panels';
    panels.forEach(({ panel }) => holder.append(panel));
    block.append(holder);
    wireReveal(block, panels);
  }
}
