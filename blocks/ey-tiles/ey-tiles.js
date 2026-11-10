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
 * Button layouts: yellow (rounded yellow link buttons, 3 per row, left-aligned; add `centered` to
 * centre the last row, `short` for the compact 38px size), sub-service-lines (= yellow centered),
 * hot-topics (tall six-column tiles).
 * Variant `wheel`: tiles become donut segments (clockwise from 12 o'clock); a first row with a
 * single cell is the centre label. With `reveal`, panels open beside the wheel.
 */

const WHOLE_LINK = ['flip', 'buttons', 'image', 'tab', 'bar', 'slide', 'overlay', 'box', 'panorama', 'wheel'];
const SHARED_BG = ['box', 'panorama', 'flip'];
const NO_PIC = ['buttons', 'panels', 'box', 'panorama', 'flip', 'wheel']; // pictures are ignored per tile
// SharePoint CCaSS wheel colours, in authoring order
const WHEEL_COLORS = ['#161d5c', '#b82a24', '#e3500f', '#1f8a3c', '#1d5db4'];
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

/**
 * A tile description may start with a short italic teaser paragraph (<p><em>...</em></p>).
 * Returns it so the bar variant can show it on hover (SharePoint overlay text).
 */
function findTeaser(cell) {
  const p = cell && cell.querySelector(':scope > p');
  if (!p || p.children.length !== 1 || p.firstElementChild.tagName !== 'EM') return null;
  const text = textOf(p);
  return text && text === textOf(p.firstElementChild) ? { p, text } : null;
}

/**
 * Inside a reveal panel, two or more <h5> groups (heading + the content that follows it) become a
 * vertical tab list with one pane per group (SharePoint "erf-tab" widget). Content before the first h5 stays above.
 */
function buildTabs(groups) {
  const wrap = document.createElement('div');
  wrap.className = 'ey-tabs';
  const list = document.createElement('div');
  list.className = 'ey-tablist';
  list.setAttribute('role', 'tablist');
  list.setAttribute('aria-orientation', 'vertical');
  const tabs = [];
  const panes = [];
  const select = (i, focus) => {
    tabs.forEach((t, k) => {
      t.setAttribute('aria-selected', k === i ? 'true' : 'false');
      t.tabIndex = k === i ? 0 : -1;
      panes[k].hidden = k !== i;
    });
    if (focus) tabs[i].focus();
  };
  groups.forEach((g, i) => {
    uid += 1;
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'ey-tab';
    tab.id = `ey-tab-${uid}`;
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-controls', `ey-tabpane-${uid}`);
    tab.textContent = g.title;
    const pane = document.createElement('div');
    pane.className = 'ey-tabpane';
    pane.id = `ey-tabpane-${uid}`;
    pane.setAttribute('role', 'tabpanel');
    pane.setAttribute('aria-labelledby', tab.id);
    pane.append(...g.nodes);
    tab.addEventListener('click', () => select(i));
    tab.addEventListener('keydown', (e) => {
      const keys = {
        ArrowDown: i + 1, ArrowRight: i + 1, ArrowUp: i - 1, ArrowLeft: i - 1, Home: 0, End: groups.length - 1,
      };
      if (!(e.key in keys)) return;
      e.preventDefault();
      select((keys[e.key] + groups.length) % groups.length, true);
    });
    tabs.push(tab);
    panes.push(pane);
    list.append(tab);
  });
  wrap.append(list, ...panes);
  select(0);
  return wrap;
}

function tabify(body) {
  if ([...body.children].filter((el) => el.tagName === 'H5').length < 2) return;
  const intro = [];
  const groups = [];
  [...body.childNodes].forEach((n) => {
    if (n.nodeType === 1 && n.tagName === 'H5') groups.push({ title: n.textContent.trim(), nodes: [] });
    else if (groups.length) groups[groups.length - 1].nodes.push(n);
    else intro.push(n);
  });
  body.replaceChildren(...intro, buildTabs(groups));
}

/**
 * Standalone tabs: `ey-tiles (tabs)`. One row per tab: first cell = tab title, remaining cells = tab content.
 * Renders the same vertical tab list as the tabs inside a reveal panel (SharePoint erf-tab).
 */
function decorateTabs(block) {
  const groups = [...block.children].map((row) => {
    const cells = [...row.children];
    return { title: textOf(cells[0]), nodes: cells.slice(1).flatMap((c) => [...c.childNodes]) };
  }).filter((g) => g.title);
  if (!groups.length) return;
  block.replaceChildren(buildTabs(groups));
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
  } else if (item.emptyPic) {
    const pw = document.createElement('div');
    pw.className = 'ey-tile-pic';
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
    const teaser = findTeaser(descCell);
    if (teaser) {
      face.dataset.teaser = teaser.text; // shown on tile hover (bar variant)
      teaser.p.classList.add('ey-tile-teaser');
    }
    texts.slice(1).forEach((c) => body.append(...c.childNodes));
    tabify(body);
    panel.append(head, body);
    panels.push({ panel, face, close });
  }
  return li;
}

function buildWheel(list, center) {
  const tiles = [...list.children];
  const step = 360 / tiles.length;
  const point = (deg, r) => {
    const rad = (deg * Math.PI) / 180;
    return `${(50 + r * Math.sin(rad)).toFixed(2)}% ${(50 - r * Math.cos(rad)).toFixed(2)}%`;
  };
  tiles.forEach((tile, i) => {
    const from = i * step;
    const arc = [];
    for (let d = from; d < from + step; d += 5) arc.push(point(d, 50));
    arc.push(point(from + step, 50));
    tile.style.clipPath = `polygon(50% 50%, ${arc.join(', ')})`;
    const [x, y] = point(from + step / 2, 34.5).split(' ');
    tile.style.setProperty('--ey-wheel-x', x);
    tile.style.setProperty('--ey-wheel-y', y);
    tile.style.setProperty('--ey-wheel-color', WHEEL_COLORS[i % WHEEL_COLORS.length]);
  });

  const wheel = document.createElement('div');
  wheel.className = 'ey-wheel';
  wheel.append(list);
  if (center) {
    const label = document.createElement('p');
    label.className = 'ey-wheel-center';
    label.textContent = textOf(center.texts[0]);
    wheel.append(label);
  }
  return wheel;
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
  if (block.classList.contains('tabs')) {
    decorateTabs(block);
    return;
  }
  const known = ['cards', 'flip', 'buttons', 'image', 'tab', 'bar', 'slide', 'overlay', 'box', 'panorama', 'panels', 'media', 'wheel'];
  let variant = known.find((v) => block.classList.contains(v));
  if (!variant) {
    variant = 'buttons';
    block.classList.add(variant);
  }
  const pagePath = window.eyPagePath || window.location.pathname;
  if (variant === 'buttons' && /^\/ds_tax(?:\/|$)/i.test(pagePath)
    && !block.classList.contains('sub-service-lines') && !block.classList.contains('hot-topics')) {
    const previous = block.previousElementSibling
      || block.closest('.ey-tiles-wrapper')?.previousElementSibling;
    const heading = previous?.matches('h2, h3, h4') ? previous : previous?.lastElementChild;
    if (heading?.id === 'sub-service-lines') block.classList.add('sub-service-lines');
    if (heading?.id === 'tax-hot-topics') block.classList.add('hot-topics');
  }
  const reveal = block.classList.contains('reveal');
  const rows = [...block.children];

  let bgPic = null;
  let more = null;
  const items = [];
  rows.forEach((row) => {
    const item = parseRow(row);
    // media: an empty first cell is a missing photo, not an empty title
    if (variant === 'media' && !item.pic && item.texts.length > 1 && !textOf(item.texts[0])
      && !item.texts[0].querySelector('a, ul, ol')) {
      item.texts.shift();
      item.emptyPic = true;
    }
    if (item.pic && !item.texts.length && SHARED_BG.includes(variant) && !items.length && !bgPic) {
      bgPic = item.pic;
    } else if (variant === 'cards' && !item.pic && item.texts.length === 1 && item.texts[0].querySelector('a')) {
      more = item.texts[0].querySelector('a');
    } else if (item.pic || item.texts.length) {
      items.push(item);
    }
  });

  const center = variant === 'wheel' && items[0] && !items[0].pic && items[0].texts.length === 1
    ? items.shift() : null;

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
  block.append(variant === 'wheel' ? buildWheel(list, center) : list);

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
    panels.forEach(({ panel, face }) => {
      const color = face.closest('.ey-tile').style.getPropertyValue('--ey-wheel-color');
      if (color) panel.style.setProperty('--ey-wheel-color', color);
      holder.append(panel);
    });
    block.append(holder);
    wireReveal(block, panels);
  }
}
