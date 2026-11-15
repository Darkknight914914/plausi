// Plausi extension — toolbar button and right-click menus.
const CHECKER = 'checker.html';

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({ id: 'plausi-tables', title: 'Check tables on this page with Plausi', contexts: ['page', 'selection', 'frame'] });
    chrome.contextMenus.create({ id: 'plausi-link', title: 'Check linked file with Plausi', contexts: ['link'] });
  });
});

chrome.action.onClicked.addListener(() => chrome.tabs.create({ url: CHECKER }));

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  let pending;
  if (info.menuItemId === 'plausi-tables') {
    try {
      const target = { tabId: tab.id };
      if (info.frameId) target.frameIds = [info.frameId];
      const [res] = await chrome.scripting.executeScript({ target, func: grabTables });
      pending = { kind: 'tables', source: (tab.title || 'Web page').slice(0, 80), url: tab.url, tables: res.result || [] };
    } catch (e) {
      pending = { kind: 'error', message: 'Plausi can’t read this page. Browser pages such as settings or the Chrome Web Store are off limits to extensions. (' + e.message + ')' };
    }
  } else if (info.menuItemId === 'plausi-link') {
    pending = { kind: 'link', url: info.linkUrl };
  } else {
    return;
  }
  await chrome.storage.session.set({ pending });
  chrome.tabs.create({ url: CHECKER + '#pending', index: tab ? tab.index + 1 : undefined });
});

/**
 * Runs inside the web page: collects visible HTML tables and ARIA grids as rows of text.
 * Must stay self-contained (it is serialised and injected).
 */
function grabTables() {
  const MAX_TABLES = 20, MAX_CELLS = 200000;
  let cells = 0;
  const clean = (t) => (t || '').replace(/ /g, ' ').replace(/[ \t]*\n[ \t]*/g, ' ').trim();
  const visible = (el) => !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length);
  const usedNames = {};
  const nameFor = (el, i) => {
    let n = '';
    if (el.caption) n = el.caption.innerText;
    n = n || el.getAttribute('aria-label') || '';
    if (!n) {
      const lb = el.getAttribute('aria-labelledby');
      if (lb) n = lb.split(/\s+/).map((id) => (document.getElementById(id) || {}).innerText || '').join(' ');
    }
    if (!n) {
      let p = el;
      for (let up = 0; up < 6 && p && !n; up++) {
        let s = p.previousElementSibling;
        for (let k = 0; k < 4 && s && !n; k++, s = s.previousElementSibling) {
          const h = s.matches('h1,h2,h3,h4,h5,h6,legend,[role=heading]') ? s : s.querySelector('h1,h2,h3,h4,h5,h6,[role=heading]');
          if (h && h.innerText.trim()) n = h.innerText;
        }
        p = p.parentElement;
      }
    }
    n = clean(n).replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 28).trim() || 'Table ' + (i + 1);
    let out = n, k = 2;
    while (usedNames[out.toLowerCase()]) out = n.slice(0, 26) + ' ' + k++;
    usedNames[out.toLowerCase()] = true;
    return out;
  };

  const out = [];
  const tables = Array.from(document.querySelectorAll('table')).filter(visible);
  for (const tbl of tables) {
    if (out.length >= MAX_TABLES || cells > MAX_CELLS) break;
    const rows = [], spans = [];
    for (const tr of Array.from(tbl.rows)) {
      if (!visible(tr)) continue;
      const row = [];
      let c = 0;
      const skip = () => { while (spans[c] && spans[c] > 0) { row[c] = ''; spans[c]--; c++; } };
      for (const cell of Array.from(tr.cells)) {
        skip();
        row[c] = clean(cell.innerText);
        const cs = Math.min(cell.colSpan || 1, 50), rs = Math.min(cell.rowSpan || 1, 1000);
        for (let k = 0; k < cs; k++) { if (k > 0) row[c + k] = ''; if (rs > 1) spans[c + k] = rs - 1; }
        c += cs;
      }
      skip();
      cells += row.length;
      rows.push(row);
    }
    const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
    if (rows.length >= 2 && width >= 2) out.push({ name: nameFor(tbl, out.length), rows });
  }
  // Grids built from <div>s (common in web apps)
  const grids = Array.from(document.querySelectorAll('[role=grid],[role=treegrid],[role=table]')).filter((g) => visible(g) && g.tagName !== 'TABLE' && !g.closest('table'));
  for (const g of grids) {
    if (out.length >= MAX_TABLES || cells > MAX_CELLS) break;
    const rows = Array.from(g.querySelectorAll('[role=row]')).filter((r) => r.closest('[role=grid],[role=treegrid],[role=table]') === g).map((r) =>
      Array.from(r.querySelectorAll('[role=columnheader],[role=rowheader],[role=gridcell],[role=cell]')).map((c) => clean(c.innerText)));
    cells += rows.reduce((s, r) => s + r.length, 0);
    if (rows.length >= 2 && rows.some((r) => r.length >= 2)) out.push({ name: nameFor(g, out.length), rows });
  }
  return out;
}
