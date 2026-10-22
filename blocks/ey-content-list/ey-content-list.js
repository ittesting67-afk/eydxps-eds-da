/*
 * ey-content-list - list/feed block for the EY Discover (SharePoint) migration.
 *
 * One block, several looks (variants), two data modes:
 *   (a) authored rows   - one table row per item
 *   (b) data sheet      - a single-cell row with a link to a DA/EDS JSON sheet (.json)
 * Spec: /workspace/ch-migration/blocks/ey-content-list.md
 *
 * Layout variants (class on the block):  cards | tiles | rows | events | rail | meta
 * Presets (set a default layout + label formatting): news, featured-topics -> cards;
 *   announcements -> tiles; events -> events;  `list` forces `rows`.
 * Modifiers: limit-N, page-N, by-priority, sort-priority, sort-date, sort-asc, upcoming, past,
 *   group-month, rule, compact, mobile-image.
 */
import { createOptimizedPicture } from '../../scripts/aem.js';

const LAYOUTS = ['cards', 'tiles', 'rows', 'events', 'rail', 'meta'];
const PRESET_LAYOUT = {
  news: 'cards',
  'featured-topics': 'cards',
  announcements: 'tiles',
  events: 'events',
};
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December'];
const DATE_RE = /(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})/;
const TIME_RE = /^\s*\d{1,2}:\d{2}/;

/* ------------------------------------------------------------------ dates */
function parseDate(value) {
  if (value === undefined || value === null || value === '') return null;
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === 'number' || /^\d{9,13}$/.test(String(value).trim())) {
    let n = Number(value);
    if (n < 1e11) n *= 1000; // EDS sheets store lastModified in seconds
    return new Date(n);
  }
  const s = String(value).trim();
  // ISO without zone -> treat as UTC calendar date (avoid off-by-one in other time zones)
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
  m = s.match(DATE_RE);
  if (m) {
    const mi = MONTHS.findIndex((x) => x.toLowerCase() === m[2].slice(0, 3).toLowerCase());
    if (mi > -1) return new Date(Date.UTC(+m[3], mi, +m[1]));
  }
  // `June 8, 2026` / `8/6/2026` style fallbacks; anything else is NOT a date (e.g. tags)
  if (/^[A-Za-z]{3,9}\.? \d{1,2},? \d{4}/.test(s) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
    const t = Date.parse(s);
    if (!Number.isNaN(t)) {
      const d = new Date(t);
      return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    }
  }
  return null;
}

/** `8 Jun 2026` (the SharePoint news script writes `Sept`; announcements/events `Sep`). */
function formatDate(d, long) {
  if (!d) return '';
  let mon = MONTHS[d.getUTCMonth()];
  if (long && mon === 'Sep') mon = 'Sept';
  return `${d.getUTCDate()} ${mon} ${d.getUTCFullYear()}`;
}

/* ------------------------------------------------------------ item sources */
function textOf(el) {
  return (el?.textContent || '').replace(/\s+/g, ' ').trim();
}

/** `meta` rows: [start date | time text | add-to-calendar link | end date] - cells typed by content. */
function metaFromRow(row) {
  const item = {};
  [...row.children].forEach((cell) => {
    const link = cell.querySelector('a[href]');
    const lines = [...cell.querySelectorAll('p')].map(textOf).filter(Boolean);
    if (!lines.length && textOf(cell)) lines.push(textOf(cell));
    if (link && !parseDate(textOf(link))) {
      item.calendarLink = link.getAttribute('href');
      item.calendarText = textOf(link);
      return;
    }
    lines.forEach((line) => {
      if (TIME_RE.test(line)) item.time = item.time ? `${item.time} ${line}` : line;
      else if (parseDate(line) && !item.date) { item.date = parseDate(line); item.dateText = line; }
      else if (parseDate(line)) item.endDate = parseDate(line);
      else if (!item.time) item.time = line;
    });
  });
  return item;
}

/** Turn one authored row (cells: image | text | date | tag) into an item object. */
function itemFromRow(row) {
  const cells = [...row.children];
  const item = {};
  const pic = row.querySelector('picture, img');
  if (pic) item.picture = pic.closest('picture') || pic;
  // text cell = first cell that is not only the picture
  const candidates = cells.filter((c) => textOf(c) && !(c.querySelector('picture, img') && !c.querySelector('a, p, h1, h2, h3, h4, h5, h6')));
  // the text cell is the one holding the title link/heading; else the first cell that is not a bare date/time
  const textCell = candidates.find((c) => c.querySelector('a[href], h1, h2, h3, h4, h5, h6, strong'))
    || candidates.find((c) => !parseDate(textOf(c)) && !TIME_RE.test(textOf(c)))
    || candidates[0];
  const rest = cells.filter((c) => c !== textCell && textOf(c) && !c.querySelector('picture, img'));
  if (textCell) {
    const a = textCell.querySelector('a[href]');
    const head = textCell.querySelector('h1, h2, h3, h4, h5, h6, strong');
    if (a) {
      item.url = a.getAttribute('href');
      item.title = textOf(a);
    } else if (head) {
      item.title = textOf(head);
    }
    let paras = [...textCell.querySelectorAll('p, li, div')].filter((p) => !p.querySelector('p, li, div') && textOf(p));
    if (!paras.length && textOf(textCell)) paras = [textCell];
    const skip = (p) => (a && p.contains(a)) || (head && (p === head || p.contains(head) || head.contains(p)));
    const em = textCell.querySelector('em, time');
    if (em && parseDate(textOf(em))) {
      item.dateText = textOf(em);
      item.date = parseDate(item.dateText);
    }
    const desc = [];
    paras.forEach((p) => {
      if (skip(p) || (em && p.contains(em) && textOf(p) === textOf(em))) return;
      const t = textOf(p);
      if (TIME_RE.test(t) && !item.time) item.time = t;
      else desc.push(t);
    });
    if (!item.title && desc.length) item.title = desc.shift();
    item.description = desc.join(' ');
  }
  // extra cells: first parsable date, optional end date / tag
  rest.forEach((c) => {
    const t = textOf(c);
    const link = c.querySelector('a[href]');
    const lines = [...c.querySelectorAll('p')].map(textOf).filter(Boolean);
    const parts = lines.length ? lines : [t];
    parts.forEach((part) => {
      if (link && !parseDate(part) && !item.calendarLink && /calendar|ics/i.test(`${part} ${link.href}`)) {
        item.calendarLink = link.getAttribute('href');
      } else if (!item.date && parseDate(part)) {
        item.dateText = part;
        item.date = parseDate(part);
      } else if (item.date && !item.endDate && parseDate(part) && DATE_RE.test(part)) {
        item.endDate = parseDate(part);
      } else if (TIME_RE.test(part) && !item.time) item.time = part;
      else if (!item.tag) item.tag = part;
    });
  });
  return item;
}

const KEYS = {
  title: ['title', 'name'],
  url: ['link', 'path', 'url', 'targetlink', 'href'],
  description: ['description', 'summary', 'desc', 'abstract'],
  image: ['image', 'imageurl', 'img', 'thumbnail'],
  alt: ['imagealt', 'alt'],
  date: ['date', 'publisheddate', 'published', 'eventstartdate', 'startdate', 'start', 'lastmodified'],
  endDate: ['enddate', 'end'],
  time: ['time', 'timedisplay', 'eventtime'],
  tag: ['tag', 'label', 'type', 'category'],
  priority: ['priority', 'hspriority'],
  calendarLink: ['calendarlink', 'addtocalendar', 'ics'],
};

function itemFromRecord(rec) {
  const low = {};
  Object.keys(rec).forEach((k) => { low[k.toLowerCase()] = rec[k]; });
  const get = (name) => {
    const k = KEYS[name].find((x) => low[x] !== undefined && String(low[x]).trim() !== '');
    return k ? String(low[k]).trim() : '';
  };
  const date = parseDate(get('date'));
  return {
    title: get('title'),
    url: get('url'),
    description: get('description'),
    image: get('image'),
    alt: get('alt'),
    date,
    dateText: '',
    endDate: parseDate(get('endDate')),
    time: get('time'),
    tag: get('tag'),
    priority: get('priority') === '' ? null : Number(get('priority')),
    calendarLink: get('calendarLink'),
    raw: low,
  };
}

async function loadSheet(href, preset) {
  const u = new URL(href, window.location.href);
  const sheet = u.searchParams.get('sheet');
  const filters = {};
  u.searchParams.forEach((v, k) => {
    if (!['sheet', 'limit', 'offset'].includes(k)) filters[k.toLowerCase()] = v;
  });
  const q = new URLSearchParams({ limit: '1000' });
  if (sheet) q.set('sheet', sheet);
  const resp = await fetch(`${u.pathname}?${q}`);
  if (!resp.ok) throw new Error(`${resp.status} ${u.pathname}`);
  const json = await resp.json();
  let rows = json.data;
  if (!rows && json[':names']) {
    const names = json[':names'];
    const pick = sheet || names.find((n) => n === preset) || names[0];
    rows = json[pick]?.data;
  }
  rows = (rows || []).map(itemFromRecord);
  Object.keys(filters).forEach((k) => {
    const wanted = filters[k].toLowerCase().split(',').map((x) => x.trim());
    rows = rows.filter((r) => wanted.includes(String(r.raw[k] ?? '').toLowerCase()));
  });
  return rows;
}

/* ------------------------------------------------------------ item helpers */
function isExternal(href) {
  try {
    const u = new URL(href, window.location.href);
    if (u.host === window.location.host) return false;
    return !/sites\.ey\.com\/sites\/DS_/i.test(u.href);
  } catch (e) {
    return false;
  }
}

/** data comes from authors/sheets: only allow relative, http(s), mailto and hash targets */
function safeUrl(href) {
  const h = (href || '').trim();
  if (!h) return '';
  if (/^(https?:|mailto:|\/|#|\.{1,2}\/)/i.test(h) || !/^[a-z][a-z0-9+.-]*:/i.test(h)) return h;
  return '';
}

function makeLink(rawHref, text) {
  const a = document.createElement('a');
  const href = safeUrl(rawHref) || (String(rawHref).startsWith('data:text/calendar') ? rawHref : '');
  if (href) {
    a.href = href;
    if (/^https?:/i.test(href) && isExternal(href)) {
      a.target = '_blank';
      a.rel = 'noopener';
    }
  }
  a.textContent = text;
  return a;
}

function makePicture(item) {
  if (item.picture) {
    const pic = item.picture.cloneNode(true);
    const img = pic.tagName === 'IMG' ? pic : pic.querySelector('img');
    if (img) img.loading = 'lazy';
    return pic;
  }
  if (!item.image || !safeUrl(item.image)) return null;
  const alt = item.alt || '';
  if (item.image.startsWith('/') || item.image.startsWith(window.location.origin)) {
    return createOptimizedPicture(item.image, alt, false, [{ media: '(min-width: 600px)', width: '600' }, { width: '400' }]);
  }
  const pic = document.createElement('picture');
  const img = document.createElement('img');
  img.src = item.image;
  img.alt = alt;
  img.loading = 'lazy';
  pic.append(img);
  return pic;
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text) e.textContent = text;
  return e;
}

function titleEl(item, level = 'h4') {
  const h = el(level, 'ey-cl-title');
  if (item.url) h.append(makeLink(item.url, item.title));
  else h.textContent = item.title;
  return h;
}

/* ---------------------------------------------------------- item rendering */
function calendarBox(date) {
  const c = el('div', 'ey-cl-calendar');
  c.append(el('div', 'ey-cl-cal-month', date ? MONTHS[date.getUTCMonth()] : ''));
  c.append(el('div', 'ey-cl-cal-day', date ? String(date.getUTCDate()) : ''));
  return c;
}

function renderItem(item, cfg) {
  const li = el('li', 'ey-cl-item');
  const { layout } = cfg;
  const dateStr = item.date ? formatDate(item.date, cfg.longSep) : item.dateText || '';

  if (layout === 'events') {
    li.append(calendarBox(item.date));
    const body = el('div', 'ey-cl-body');
    body.append(titleEl(item));
    if (item.time) {
      const t = el('div', 'ey-cl-time');
      t.append(item.url ? makeLink(item.url, item.time) : document.createTextNode(item.time));
      body.append(t);
    }
    if (item.description) body.append(el('p', 'ey-cl-desc', item.description));
    li.append(body);
    return li;
  }

  const pic = layout === 'cards' && !item.picture && !item.image ? null : makePicture(item);
  if (pic) {
    const wrap = el('div', 'ey-cl-img');
    wrap.append(pic);
    li.append(wrap);
  }
  const body = el('div', 'ey-cl-body');
  if (layout === 'rail' && item.tag) {
    const top = el('div', 'ey-cl-top');
    top.append(el('span', 'ey-cl-tag', item.tag));
    if (cfg.viewAll) {
      const va = makeLink(cfg.viewAll.href, `${cfg.viewAll.text} (${cfg.tagCount})`);
      va.className = 'ey-cl-viewall';
      top.append(va);
    }
    body.append(top);
  }
  if (layout === 'rows') {
    const head = el('div', 'ey-cl-head');
    head.append(titleEl(item));
    if (dateStr) head.append(el('div', 'ey-cl-date', dateStr));
    body.append(head);
    if (item.description) body.append(el('div', 'ey-cl-desc', item.description));
  } else {
    body.append(titleEl(item));
    if (layout === 'cards') {
      // source order: title, description, date (date pinned below the fixed-height description)
      body.append(el('div', 'ey-cl-desc', item.description || ''));
      body.append(el('div', 'ey-cl-date', dateStr));
    } else {
      if (dateStr && layout !== 'rail') body.append(el('div', 'ey-cl-date', dateStr));
      if (item.description) body.append(el('p', 'ey-cl-desc', item.description));
    }
  }
  li.append(body);
  return li;
}

/* ------------------------------------------------------------- ics (meta) */
function icsStamp(d) {
  return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
}

function icsHref(title, start, end) {
  const s = start;
  const e = end || new Date(start.getTime() + 3600000);
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//EY Discover//EN', 'BEGIN:VEVENT',
    `UID:${icsStamp(s)}-${encodeURIComponent(title).slice(0, 30)}@discover`,
    `DTSTAMP:${icsStamp(new Date())}`, `DTSTART:${icsStamp(s)}`, `DTEND:${icsStamp(e)}`,
    `SUMMARY:${title.replace(/[,;]/g, ' ')}`, 'END:VEVENT', 'END:VCALENDAR'];
  return `data:text/calendar;charset=utf-8,${encodeURIComponent(lines.join('\r\n'))}`;
}

function renderMeta(item, cfg) {
  const wrap = el('div', 'ey-cl-meta');
  wrap.append(calendarBox(item.date));
  const c = el('div', 'ey-cl-meta-text');
  c.append(el('span', '', item.date ? formatDate(item.date) : item.dateText || ''));
  if (item.time) {
    c.append(document.createElement('br'));
    c.append(el('span', '', item.time));
  }
  wrap.append(c);
  if (cfg.moreLink || item.calendarLink) {
    const add = el('div', 'ey-cl-addcal');
    const label = item.calendarText || cfg.moreLink?.text || 'Add to calendar';
    let href = safeUrl(item.calendarLink) || safeUrl(cfg.moreLink?.href) || '';
    if ((!href || href === '#' || /\/ics$/i.test(href)) && item.date) {
      href = icsHref(document.title || label, item.date, item.endDate);
    }
    const a = el('a', '', '');
    a.href = href;
    if (href.startsWith('data:')) a.download = 'event.ics';
    a.append(el('i', 'ey-cl-icon-cal'), el('span', '', label));
    add.append(a);
    wrap.append(add);
  }
  return wrap;
}

/* ------------------------------------------------------------ list pipeline */
function startOfToday() {
  const n = new Date();
  return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate());
}

function arrange(items, flags, layout) {
  let list = layout === 'meta' ? items : items.filter((i) => i.title || i.picture || i.image);
  if (flags.upcoming) list = list.filter((i) => i.date && (i.endDate || i.date).getTime() >= startOfToday());
  if (flags.past) list = list.filter((i) => i.date && (i.endDate || i.date).getTime() < startOfToday());
  const byDate = (dir) => (a, b) => dir * ((a.date?.getTime() || 0) - (b.date?.getTime() || 0));
  if (flags.byPriority) {
    // SharePoint home 'What's New': newest item for each Priority value 1,2,3
    const seen = new Set();
    const picked = [];
    [...list].sort(byDate(-1)).forEach((i) => {
      if (i.priority && !seen.has(i.priority)) { seen.add(i.priority); picked.push(i); }
    });
    list = picked.sort((a, b) => a.priority - b.priority);
  } else if (flags.sortPriority) {
    list = [...list].sort((a, b) => ((a.priority || 0) - (b.priority || 0)) || byDate(-1)(a, b));
  } else if (flags.sortDate || flags.sortAsc || layout === 'events') {
    const asc = flags.sortAsc || (layout === 'events' && !flags.past);
    list = [...list].sort(byDate(asc ? 1 : -1));
  } else if (flags.fromSheet) {
    // data-sheet rows default to newest first; authored rows keep their position (first)
    list = [...list.filter((i) => !i.fromSheet), ...list.filter((i) => i.fromSheet).sort(byDate(-1))];
  }
  return list;
}

function pageControls(state, onPage) {
  const nav = el('div', 'ey-cl-pager');
  nav.setAttribute('aria-label', 'Pagination');
  const { page, size, total } = state;
  const pages = Math.ceil(total / size);
  if (pages < 2) return nav;
  const start = (page - 1) * size;
  const end = Math.min(start + size, total);
  if (page > 1) {
    const b = el('button', 'ey-cl-page ey-cl-prev');
    b.type = 'button';
    b.setAttribute('aria-label', 'Previous page');
    b.addEventListener('click', () => onPage(page - 1));
    nav.append(b);
  }
  nav.append(document.createTextNode(`Displaying items ${start + 1} - ${end} of ${total} `));
  if (page < pages) {
    const b = el('button', 'ey-cl-page ey-cl-next');
    b.type = 'button';
    b.setAttribute('aria-label', 'Next page');
    b.addEventListener('click', () => onPage(page + 1));
    nav.append(b);
  }
  return nav;
}

function renderGroups(list, cfg) {
  const container = el('div', 'ey-cl-items');
  if (cfg.layout === 'events' && cfg.groupMonth) {
    let cur = '';
    let ul;
    list.forEach((item) => {
      const key = item.date ? `${MONTH_NAMES[item.date.getUTCMonth()]} ${item.date.getUTCFullYear()}` : '';
      if (key !== cur || !ul) {
        cur = key;
        if (key) container.append(el('h2', 'ey-cl-group', key));
        ul = el('ul', 'ey-cl-list');
        container.append(ul);
      }
      ul.append(renderItem(item, cfg));
    });
  } else {
    const ul = el('ul', 'ey-cl-list');
    list.forEach((item) => ul.append(renderItem(item, cfg)));
    container.append(ul);
    if (cfg.rule) container.append(el('hr', 'ey-cl-rule'));
  }
  return container;
}

/* ------------------------------------------------------------------ decorate */
export default async function decorate(block) {
  const cls = [...block.classList];
  const num = (prefix) => {
    const c = cls.find((x) => x.startsWith(`${prefix}-`) && /^\d+$/.test(x.slice(prefix.length + 1)));
    return c ? Number(c.slice(prefix.length + 1)) : 0;
  };
  const preset = Object.keys(PRESET_LAYOUT).find((p) => cls.includes(p)) || '';
  let layout = LAYOUTS.find((l) => cls.includes(l));
  if (!layout && cls.includes('list')) layout = 'rows';
  if (!layout) layout = PRESET_LAYOUT[preset] || 'tiles';
  if (!block.classList.contains(layout)) block.classList.add(layout);

  const flags = {
    byPriority: cls.includes('by-priority'),
    sortPriority: cls.includes('sort-priority'),
    sortDate: cls.includes('sort-date'),
    sortAsc: cls.includes('sort-asc'),
    upcoming: cls.includes('upcoming'),
    past: cls.includes('past'),
  };
  const cfg = {
    layout,
    groupMonth: cls.includes('group-month'),
    rule: cls.includes('rule'),
    longSep: preset === 'news' || preset === 'featured-topics',
  };
  const limit = num('limit');
  const pageSize = num('page');

  // ---- split rows: item rows (>=2 cells) vs control rows (1 cell)
  const rows = [...block.children];
  let title = null;
  let sheetHref = '';
  let more = null;
  const itemRows = [];
  rows.forEach((row) => {
    const cells = [...row.children];
    const a = row.querySelector('a[href]');
    if (cells.length === 1) {
      const cell = cells[0];
      const heading = cell.querySelector('h1, h2, h3, h4, h5, h6');
      if (heading && !a) title = heading.cloneNode(true);
      else if (a && /\.json(\?|#|$)/i.test(a.getAttribute('href'))) sheetHref = a.getAttribute('href');
      else if (a && !cell.querySelector('picture')) more = { href: a.getAttribute('href'), text: textOf(a) };
      else if (cell.querySelector('picture, img') || textOf(cell)) itemRows.push(row);
    } else if (textOf(row) || row.querySelector('picture, img')) {
      itemRows.push(row);
    }
  });
  if (more) cfg.moreLink = more;
  if (layout === 'rail' && more) cfg.viewAll = more;

  let items = itemRows.map(layout === 'meta' ? metaFromRow : itemFromRow);
  if (sheetHref) {
    try {
      items = items.concat((await loadSheet(sheetHref, preset)).map((i) => ({ ...i, fromSheet: true })));
      flags.fromSheet = true;
    } catch (e) {
      // eslint-disable-next-line no-console
      console.warn('ey-content-list: data sheet failed', e);
      block.dataset.error = 'sheet';
    }
  }

  let list = arrange(items, flags, layout);
  cfg.tagCount = list.filter((i) => i.tag).length;
  if (!list.length) block.classList.add('ey-cl-empty');

  block.textContent = '';
  if (title || (more && layout === 'events')) {
    const head = el('div', 'ey-cl-block-head');
    if (title) head.append(title);
    if (more && layout === 'events') {
      const m = makeLink(more.href, more.text);
      m.className = 'ey-cl-top-link';
      m.prepend(el('i', 'ey-cl-icon-history'));
      head.append(m);
    }
    block.append(head);
  }

  if (layout === 'meta') {
    if (list[0]) block.append(renderMeta(list[0], cfg));
    return;
  }

  if (limit) list = list.slice(0, limit);
  const body = el('div', 'ey-cl-wrap');
  block.append(body);
  const draw = (page) => {
    body.textContent = '';
    let view = list;
    if (pageSize) view = list.slice((page - 1) * pageSize, page * pageSize);
    body.append(renderGroups(view, cfg));
    if (pageSize) {
      body.append(pageControls({ page, size: pageSize, total: list.length }, (p) => {
        draw(p);
        block.scrollIntoView({ block: 'start' });
      }));
    }
  };
  draw(1);

  if (more && layout !== 'events' && layout !== 'rail') {
    const m = makeLink(more.href, more.text);
    m.className = 'ey-cl-more';
    block.append(m);
  }
}
