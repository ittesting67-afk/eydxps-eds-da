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

function linkPath(a) {
  try { return norm(new URL(a.getAttribute('href'), window.location.href).pathname); } catch (e) { return ''; }
}

/** text of an <li> excluding nested lists */
function ownLabel(li) {
  const clone = li.cloneNode(true);
  clone.querySelectorAll('ul, ol').forEach((n) => n.remove());
  return clone;
}

/** parse authored list into groups: [{ title: Element|null, items: [<a>|text] }] */
function parse(list) {
  const groups = [];
  let loose;
  [...list.children].forEach((li) => {
    const sub = li.querySelector(':scope > ul, :scope > ol');
    if (sub) {
      loose = null;
      groups.push({
        title: ownLabel(li),
        items: [...sub.children].map((c) => ownLabel(c)),
      });
    } else {
      if (!loose) {
        loose = { title: null, items: [] };
        groups.push(loose);
      }
      loose.items.push(ownLabel(li));
    }
  });
  return groups;
}

async function loadList(block) {
  const list = block.querySelector('ul, ol');
  if (list) return list;
  const a = block.querySelector('a');
  if (!a) return null;
  const path = new URL(a.href, window.location.href).pathname.replace(/\.html$/, '');
  try {
    const resp = await fetch(`${path}.plain.html`);
    if (!resp.ok) return null;
    const doc = new DOMParser().parseFromString(await resp.text(), 'text/html');
    return doc.querySelector('ul, ol');
  } catch (e) {
    return null;
  }
}

/**
 * ey-side-nav - left navigation (SharePoint eyParentMainLeftNav). Content model: a nested list
 * (top-level item = group heading, nested list = pages) OR a single link to a fragment that holds
 * that list. The current page is highlighted by path match. Marks its section `ey-side` so the
 * theme lays the page out as left column + content. On mobile every group is a collapsible panel.
 * @param {Element} block
 */
export default async function decorate(block) {
  const list = await loadList(block);
  if (!list) return;
  const here = currentPath();
  const groups = parse(list);

  block.textContent = '';
  const nav = document.createElement('nav');
  nav.setAttribute('aria-label', 'Section navigation');

  groups.forEach((g, gi) => {
    const group = document.createElement('div');
    group.className = 'ey-side-nav-group';
    const ul = document.createElement('ul');
    ul.id = `ey-side-nav-${gi}-${Math.random().toString(36).slice(2, 7)}`;
    let hasActive = false;

    g.items.forEach((item) => {
      const li = document.createElement('li');
      const a = item.querySelector('a');
      if (a) {
        const link = a.cloneNode(true);
        li.append(link);
        if (linkPath(a) && linkPath(a) === here) {
          li.classList.add('active');
          link.setAttribute('aria-current', 'page');
          hasActive = true;
        }
      } else {
        li.textContent = item.textContent.trim();
      }
      ul.append(li);
    });

    if (g.title) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'ey-side-nav-title';
      btn.setAttribute('aria-controls', ul.id);
      btn.textContent = g.title.textContent.trim();
      group.dataset.hasActive = hasActive;
      btn.setAttribute('aria-expanded', 'true');
      btn.addEventListener('click', () => {
        // collapsing only exists on mobile; on desktop all groups are always shown
        if (!window.matchMedia('(max-width: 767px)').matches) return;
        const isOpen = group.classList.toggle('open');
        btn.setAttribute('aria-expanded', isOpen);
      });
      group.classList.add('open');
      group.append(btn);
    } else {
      group.classList.add('open', 'untitled');
    }
    group.append(ul);
    nav.append(group);
  });

  // on mobile: when the current page is in one group, only that group starts open
  if (nav.querySelector('.ey-side-nav-group[data-has-active="true"]')) {
    nav.querySelectorAll('.ey-side-nav-group[data-has-active="false"]').forEach((el) => {
      el.classList.remove('open');
      el.querySelector('button')?.setAttribute('aria-expanded', 'false');
    });
  }

  block.append(nav);
  block.closest('.section')?.classList.add('ey-side');
}
