import {
  sampleRUM,
  buildBlock,
  loadHeader,
  loadFooter,
  decorateButtons,
  addLoadingToHeader,
  decorateIcons,
  decorateSections,
  decorateBlocks,
  decorateTemplateAndTheme,
  waitForLCP,
  loadBlocks,
  loadCSS,
} from './aem.js';

const LCP_BLOCKS = []; // add your LCP blocks to the list

/**
 * Builds hero block and prepends to main in a new section.
 * @param {Element} main The container element
 */
function buildHeroBlock(main) {
  const h1 = main.querySelector('h1');
  const picture = main.querySelector('picture');
  // eslint-disable-next-line no-bitwise
  if (h1 && picture && (h1.compareDocumentPosition(picture) & Node.DOCUMENT_POSITION_PRECEDING)) {
    const section = document.createElement('div');
    section.append(buildBlock('hero', { elems: [picture, h1] }));
    main.prepend(section);
  }
}

/**
 * load fonts.css and set a session storage flag
 */
async function loadFonts() {
  await loadCSS(`${window.hlx.codeBasePath}/styles/fonts.css`);
  try {
    if (!window.location.hostname.includes('localhost')) sessionStorage.setItem('fonts-loaded', 'true');
  } catch (e) {
    // do nothing
  }
}

/**
 * Builds all synthetic blocks in a container element.
 * @param {Element} main The container element
 */
function buildAutoBlocks(main) {
  try {
    buildHeroBlock(main);
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error('Auto Blocking failed', error);
  }
}

/**
 * Decorates the main element.
 * @param {Element} main The main element
 */
function decorateAssuranceHome(main) {
  const pagePath = window.eyPagePath || window.location.pathname;
  if (!/^\/ds_assurance\/pages\/home(?:\.html)?\/?$/i.test(pagePath)) return;

  main.classList.add('assurance-home');
  const image = main.querySelector('img[alt="Assurance at EY: We\'re All in"]');
  const imageParagraph = image?.closest('p');
  const linkParagraph = imageParagraph?.nextElementSibling;
  const videoLink = linkParagraph?.querySelector('a[href*="/embed.aspx"]');
  const quote = linkParagraph?.nextElementSibling;
  if (imageParagraph && videoLink && quote?.tagName === 'BLOCKQUOTE') {
    const media = document.createElement('div');
    media.className = 'assurance-intro-media';
    imageParagraph.before(media);
    videoLink.setAttribute('aria-label', videoLink.textContent.trim());
    videoLink.classList.remove('button', 'primary', 'secondary');
    videoLink.replaceChildren(...imageParagraph.childNodes);
    imageParagraph.replaceChildren(videoLink);
    imageParagraph.className = 'assurance-video';
    quote.classList.add('assurance-video-quote');
    media.append(imageParagraph, quote);
    linkParagraph.remove();
  }

  main.querySelector('.ey-tiles.box.labelbox.reveal')?.classList.add('assurance-strategy');
  const priorities = [...main.querySelectorAll('a')].find((link) => /^All in:\s*Our Assurance Priorities$/i.test(link.textContent.trim()));
  if (priorities) {
    priorities.classList.add('assurance-priorities');
    priorities.closest('p')?.classList.add('assurance-priorities-container');
  }
}

let profileGroupId = 0;

function buildProfileGroups(run) {
  const groups = document.createElement('div');
  groups.className = 'ey-profile-groups';
  groups.style.setProperty('--ey-profile-groups', run.length);
  const tabs = document.createElement('div');
  tabs.className = 'ey-profile-groups-tabs';
  run[0].head.before(groups);
  groups.append(tabs);

  const entries = run.map(({ head, h5, body }) => {
    profileGroupId += 1;
    const panel = document.createElement('section');
    panel.className = 'ey-profile-group';
    panel.id = `ey-profile-group-${profileGroupId}`;
    panel.hidden = true;
    if (h5.id) panel.setAttribute('aria-labelledby', h5.id);
    const tab = document.createElement('button');
    tab.type = 'button';
    tab.className = 'ey-profile-group-tab';
    tab.textContent = h5.textContent.trim();
    tab.setAttribute('aria-expanded', 'false');
    tab.setAttribute('aria-controls', panel.id);
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'ey-profile-group-close';
    close.setAttribute('aria-label', `Close ${tab.textContent}`);
    close.textContent = '\u00d7';
    const bar = document.createElement('div');
    bar.className = 'ey-profile-group-head';
    bar.append(h5, close);
    panel.append(bar, body);
    head.remove();
    tabs.append(tab);
    groups.append(panel);
    return { tab, panel, close };
  });

  const show = (target) => entries.forEach(({ tab, panel }) => {
    panel.hidden = panel !== target;
    tab.setAttribute('aria-expanded', String(panel === target));
  });
  entries.forEach(({ tab, panel, close }) => {
    tab.addEventListener('click', () => show(panel.hidden ? panel : null));
    close.addEventListener('click', () => {
      show(null);
      tab.focus();
    });
  });
  groups.addEventListener('keydown', (e) => {
    const open = entries.find(({ panel }) => !panel.hidden);
    if (e.key === 'Escape' && open) {
      show(null);
      open.tab.focus();
    }
  });
}

/**
 * Media blocks listing people.ey.com profiles get round photos; two or more consecutive
 * "h5 + profile block" pairs become toggle groups (SharePoint leadership tabs).
 * @param {Element} main The main element
 */
function decorateProfileGroups(main) {
  main.querySelectorAll('.ey-tiles.media').forEach((block) => {
    const links = [...block.querySelectorAll('a[href]')];
    if (links.length && links.every((a) => /^https:\/\/people\.ey\.com\//i.test(a.href))) {
      block.classList.add('profiles');
    }
  });

  main.querySelectorAll('.section').forEach((section) => {
    const children = [...section.children];
    let run = [];
    const flush = () => {
      if (run.length > 1) buildProfileGroups(run);
      run = [];
    };
    for (let i = 0; i < children.length; i += 1) {
      const head = children[i];
      const body = children[i + 1];
      const h5 = head.classList.contains('default-content-wrapper') && head.children.length === 1
        && head.querySelector(':scope > h5');
      if (h5 && body?.querySelector(':scope > .ey-tiles.profiles')) {
        run.push({ head, h5, body });
        i += 1;
      } else {
        flush();
      }
    }
    flush();
  });
}

function isSlide(el) {
  if (!el?.matches('p') || el.children.length !== 1 || el.textContent.trim()) return false;
  const img = el.querySelector(':scope > picture > img');
  const ratio = img && img.getAttribute('width') / img.getAttribute('height');
  return ratio > 1.6 && ratio < 1.95;
}

/**
 * Two or more consecutive "h5 + 16:9 slide image" pairs become a thumbnail grid
 * (SharePoint DS_PAS strategy slides, enlarged on hover).
 * @param {Element} main The main element
 */
function decorateSlideGrids(main) {
  main.querySelectorAll('.default-content-wrapper').forEach((wrapper) => {
    const children = [...wrapper.children];
    let run = [];
    const flush = () => {
      if (run.length > 1) {
        const grid = document.createElement('div');
        grid.className = 'ey-slides';
        run[0][0].before(grid);
        run.forEach(([heading, frame]) => {
          const slide = document.createElement('div');
          slide.className = 'ey-slide';
          frame.classList.add('ey-slide-frame');
          frame.tabIndex = 0;
          slide.append(heading, frame);
          grid.append(slide);
        });
      }
      run = [];
    };
    for (let i = 0; i < children.length; i += 1) {
      if (children[i].tagName === 'H5' && isSlide(children[i + 1])) {
        run.push([children[i], children[i + 1]]);
        i += 1;
      } else {
        flush();
      }
    }
    flush();
  });
}

/**
 * Callout bar (SharePoint: <p style="border-left:8px solid #fee600; background:#efefef"><a>...</a></p>).
 * A paragraph that holds ONLY a link whose text is a full sentence (ends with a full stop) is a
 * callout, not a button: undo the button decoration and mark it so ey-theme.css can style it.
 * Short link labels (no trailing full stop) stay buttons. Author it as a plain paragraph + link.
 * @param {Element} main
 */
function decorateCallouts(main) {
  main.querySelectorAll('p.button-container').forEach((p) => {
    const a = p.querySelector(':scope > a.button');
    if (!a || p.children.length !== 1) return;
    const text = a.textContent.trim();
    if (text.length < 20 || !text.endsWith('.')) return;
    a.className = '';
    p.className = 'ey-callout';
  });
}

// eslint-disable-next-line import/prefer-default-export
export function decorateMain(main) {
  // hopefully forward compatible button decoration
  decorateButtons(main);
  decorateCallouts(main);
  decorateIcons(main);
  // buildAutoBlocks(main);
  decorateSections(main);
  decorateBlocks(main);
  decorateProfileGroups(main);
  decorateSlideGrids(main);
  decorateAssuranceHome(main);
}

/**
 * Loads everything needed to get to LCP.
 * @param {Element} doc The container element
 */
async function loadEager(doc) {
  document.documentElement.lang = 'en';
  decorateTemplateAndTheme();
  const main = doc.querySelector('main');
  if (main) {
    decorateMain(main);
    document.body.classList.add('appear');
    await waitForLCP(LCP_BLOCKS);
  }

  try {
    /* if desktop (proxy for fast connection) or fonts already loaded, load fonts.css */
    if (window.innerWidth >= 900 || sessionStorage.getItem('fonts-loaded')) {
      loadFonts();
    }
  } catch (e) {
    // do nothing
  }
}

/**
 * Loads everything that doesn't need to be delayed.
 * @param {Element} doc The container element
 */
async function loadLazy(doc) {
  const spinner = addLoadingToHeader(doc.querySelector('header'))
  const main = doc.querySelector('main');
  await loadBlocks(main);

  const { hash } = window.location;
  const element = hash ? doc.getElementById(hash.substring(1)) : false;
  if (hash && element) element.scrollIntoView();

  loadHeader(doc.querySelector('header')).then(() => {
    spinner.remove()
  });;
  loadFooter(doc.querySelector('footer'));

  loadCSS(`${window.hlx.codeBasePath}/styles/lazy-styles.css`);
  loadFonts();

  sampleRUM('lazy');
  sampleRUM.observe(main.querySelectorAll('div[data-block-name]'));
  sampleRUM.observe(main.querySelectorAll('picture > img'));
}

/**
 * Loads everything that happens a lot later,
 * without impacting the user experience.
 */
function loadDelayed() {
  // eslint-disable-next-line import/no-cycle
  window.setTimeout(() => import('./delayed.js'), 3000);
  // load anything that can be postponed to the latest here
}

async function loadPage() {
  await loadEager(document);
  await loadLazy(document);
  loadDelayed();
}

loadPage();
