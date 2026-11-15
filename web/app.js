(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var E = PlausiEngine, W = PlausiWeb;

  var PREVIEW_ROWS = 200, ISSUE_PAGE = 300;
  var EXCEL_EXT = ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'xlt', 'ods', 'fods'];
  var TEXT_EXT = ['csv', 'tsv', 'txt'];

  var state = {
    file: { name: 'Sample payouts', format: 'sample', sample: true },
    sheets: [], current: 0, rules: [], result: null,
    filter: 'all', shown: ISSUE_PAGE, previewStart: null, flash: null, problem: ''
  };
  var sample = null;          // Claude, when the viewer's app offers it
  var downloads;              // undefined = not checked yet, null = unavailable

  // ---------------------------------------------------------------------------
  // helpers
  // ---------------------------------------------------------------------------
  function el(tag, attrs, kids) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === 'text') e.textContent = attrs[k];
      else if (k === 'class') e.className = attrs[k];
      else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] !== false && attrs[k] != null) e.setAttribute(k, attrs[k] === true ? '' : attrs[k]);
    }
    (kids || []).forEach(function (c) { if (c != null) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return e;
  }
  function plural(n, w) { return n + ' ' + w + (n === 1 ? '' : 's'); }
  function sheet() { return state.sheets[state.current]; }
  function width(s) { return s.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0); }
  function headers(s) { return W.namedHeaders(s.rows[s.headerRow] || [], width(s)); }
  function headersBySheet() {
    var o = {};
    state.sheets.forEach(function (s) { o[s.name] = headers(s).filter(Boolean); });
    return o;
  }
  function signature() {
    return state.sheets.map(function (s) { return s.name + ':' + headers(s).join('|'); }).join('||');
  }
  function store(key, val) {
    try { if (val === undefined) return JSON.parse(localStorage.getItem(key) || 'null'); localStorage.setItem(key, JSON.stringify(val)); }
    catch (e) { return null; }
  }
  function cellRef(i) { return i.row ? (i.colIndex >= 0 ? W.columnLetter(i.colIndex + 1) : 'A') + i.row : ''; }
  function rulesFor(s) {
    var single = state.sheets.length === 1;
    return state.rules.filter(function (r) { return r.sheet === s.name || (!r.sheet && single); });
  }

  // ---------------------------------------------------------------------------
  // loading
  // ---------------------------------------------------------------------------
  function loadSample() {
    var d = PlausiDemo.build(new Date());
    state.file = { name: 'Sample payouts', format: 'sample', sample: true };
    state.sheets = [{ name: d.sheetName, rows: [d.headers].concat(d.rows), headerRow: 0, textInput: false }];
    state.rules = d.rules.map(function (r) { var o = {}; for (var k in r) o[k] = r[k]; o.enabled = true; return o; });
    state.current = 0;
    state.problem = '';
    resetView();
  }

  function resetView() { state.filter = 'all'; state.shown = ISSUE_PAGE; state.previewStart = null; state.flash = null; run(); }

  function ext(name) { var m = /\.([^.]+)$/.exec(name || ''); return m ? m[1].toLowerCase() : ''; }

  function readText(file) {
    return file.arrayBuffer().then(function (buf) {
      var text = new TextDecoder('utf-8').decode(buf);
      if (text.indexOf('\uFFFD') > -1) {
        try { text = new TextDecoder('windows-1252').decode(buf); } catch (e) { /* keep utf-8 */ }
      }
      return text;
    });
  }

  function openFile(file) {
    var x = ext(file.name);
    var base = file.name.replace(/\.[^.]+$/, '');
    var p;
    if (TEXT_EXT.indexOf(x) > -1) {
      p = readText(file).then(function (text) {
        var parsed = W.parseCsv(text);
        return [{ name: base.slice(0, 31) || 'Data', rows: parsed.rows, textInput: true }];
      });
    } else if (EXCEL_EXT.indexOf(x) > -1 || !x) {
      if (!window.XLSX) return showProblem(window.PLAUSI_XLSX_HINT || 'The Excel reader could not load, so only CSV files work right now. Check your connection and reload the page.');
      p = file.arrayBuffer().then(function (buf) {
        var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellNF: true, cellDates: false });
        var d1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
        return wb.SheetNames.map(function (n) { return { name: n, rows: W.sheetToRows(XLSX, wb.Sheets[n], d1904), textInput: false }; });
      });
    } else {
      return showProblem('"' + file.name + '" is not a spreadsheet. Open an .xlsx, .xlsm, .xlsb, .xls, .ods or .csv file.');
    }
    $('fileName').textContent = 'Reading ' + file.name + '…';
    loadSheets(file.name, x || 'file', p);
  }

  /** Shows sheets from any source: a file, a web page's tables, a downloaded link. */
  function loadSheets(name, format, sheetsOrPromise) {
    Promise.resolve(sheetsOrPromise).then(function (sheets) {
      var rulesSheet = sheets.filter(function (s) { return s.name === W.RULES_SHEET; })[0];
      var data = sheets.filter(function (s) { return s.name !== W.RULES_SHEET && s.name !== W.REPORT_SHEET; });
      if (!data.length || data.every(function (s) { return !s.rows.length; })) throw new Error('This file has no data in it.');
      data.forEach(function (s) { s.headerRow = W.detectHeaderRow(s.rows); });
      state.file = { name: name, format: format, sample: false, rulesFrom: '' };
      state.sheets = data;
      state.current = Math.max(0, data.findIndex(function (s) { return s.rows.length > 0; }));
      var fromFile = rulesSheet ? W.rulesFromRows(rulesSheet.rows) : [];
      var remembered = store('plausi.rules.' + signature());
      if (fromFile.length) { state.rules = fromFile; state.file.rulesFrom = 'file'; }
      else if (remembered && remembered.length) { state.rules = remembered; state.file.rulesFrom = 'memory'; }
      else { state.rules = suggestAll(); state.file.rulesFrom = state.rules.length ? 'suggested' : ''; }
      state.problem = '';
      resetView();
    }).catch(function (e) {
      var msg = String(e && e.message || e);
      if (/password|encrypt/i.test(msg)) msg = 'This file is password-protected. Save an unprotected copy in Excel and open that.';
      showProblem('Could not read "' + name + '": ' + msg);
      render();
    });
  }

  function suggestAll() {
    var out = [];
    state.sheets.forEach(function (s) {
      out = out.concat(W.suggestForSheet(s.name, s.rows, s.headerRow || 0));
    });
    return out;
  }

  function showProblem(msg) { state.problem = msg; renderBanner(); }

  // ---------------------------------------------------------------------------
  // checking
  // ---------------------------------------------------------------------------
  function run() {
    var now = new Date();
    var issues = [], stats = { rules: 0, rows: 0, errors: 0, warnings: 0, config: 0 };
    var names = {};
    state.sheets.forEach(function (s) {
      names[s.name] = true;
      var rs = rulesFor(s).filter(function (r) { return r.enabled !== false; });
      var h = s.headerRow || 0;
      rs = rs.map(function (r) {
        if (r.check !== 'sum_equals' || !/^\$?[A-Z]{1,3}\$?\d+$/i.test(String(r.param).trim())) return r;
        var o = {}; for (var k in r) o[k] = r[k];
        o.resolvedParam = W.cellValue(s.rows, r.param);
        return o;
      });
      var res = E.runRules({ sheetName: s.name, headers: headers(s), rows: W.dataOnly(s.rows, h), firstDataRow: h + 2 }, rs, { now: now, textInput: s.textInput });
      res.issues.forEach(function (i) { i.sheet = s.name; });
      issues = issues.concat(res.issues);
      stats.rules += res.stats.rules; stats.rows += res.stats.rows;
      stats.errors += res.stats.errors; stats.warnings += res.stats.warnings; stats.config += res.stats.config;
    });
    state.rules.forEach(function (r) {
      if (r.enabled === false || !r.sheet || names[r.sheet]) return;
      issues.push(E.configIssue(r, 'Sheet "' + r.sheet + '" is not in this file'));
      stats.config++; stats.rules++;
    });
    var order = { config: 0, error: 1, warning: 2 };
    issues.sort(function (a, b) {
      return (order[a.severity] - order[b.severity]) || (a.sheet < b.sheet ? -1 : a.sheet > b.sheet ? 1 : 0) || ((a.row || 0) - (b.row || 0));
    });
    state.result = { issues: issues, stats: stats };
    if (!state.file.sample) store('plausi.rules.' + signature(), state.rules);
    render();
  }

  // ---------------------------------------------------------------------------
  // rendering
  // ---------------------------------------------------------------------------
  function render() {
    renderFile(); renderBanner(); renderSheetControls(); renderRules(); renderSummary(); renderIssues(); renderPreview(); renderDownloads();
  }

  function renderFile() {
    $('fileName').textContent = state.file.name;
    $('fileFormat').textContent = state.file.sample ? 'Sample' : state.file.format;
  }

  function renderBanner() {
    var b = $('banner');
    b.innerHTML = '';
    if (state.problem) {
      b.className = 'note problem';
      b.appendChild(el('span', { text: state.problem }));
      return;
    }
    if (state.file.sample) {
      b.className = 'note sample';
      b.appendChild(el('span', {}, [el('b', { text: 'This is sample data: ' }), 'a payouts export with 13 planted problems and 15 rules. Open your own file or drop it anywhere on the page.']));
    } else {
      b.className = 'note';
      var msg = state.file.rulesFrom === 'file' ? 'Rules loaded from the file’s “Plausi Rules” sheet.'
        : state.file.rulesFrom === 'memory' ? 'Rules restored from the last time you checked a file with these columns.'
        : state.file.rulesFrom === 'suggested' ? 'This file had no rules, so Plausi suggested ' + plural(state.rules.length, 'starter rule') + ' from your columns. Check them on the left, delete what doesn’t fit, and add your own.'
        : 'No rules yet. Describe what the data should look like, on the left.';
      b.appendChild(el('span', {}, [el('b', { text: plural(state.sheets.length, 'sheet') + ' read. ' }), msg]));
    }
    var f = el('span', { class: 'formats' });
    ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv', 'tsv'].forEach(function (x) { f.appendChild(el('span', { text: '.' + x })); });
    b.appendChild(f);
  }

  function renderSheetControls() {
    var sel = $('sheetSelect');
    sel.innerHTML = '';
    state.sheets.forEach(function (s, i) {
      sel.appendChild(el('option', { value: i, text: s.name + (s.rows.length ? '' : ' (empty)') }));
    });
    sel.value = state.current;
    sel.disabled = state.sheets.length < 2;
    var s = sheet();
    $('headerRow').value = s ? s.headerRow + 1 : 1;
    $('headerRow').max = s ? Math.max(1, s.rows.length) : 1;

    var hs = s ? headers(s).filter(Boolean) : [];
    var bc = $('bColumn'); bc.innerHTML = '';
    hs.forEach(function (h) { bc.appendChild(el('option', { value: h, text: h })); });
    var ex = $('examples'); ex.innerHTML = '';
    exampleSentences(hs).forEach(function (t) {
      ex.appendChild(el('button', { type: 'button', text: t, onclick: function () { $('ruleText').value = t; $('ruleText').focus(); } }));
    });
  }

  function exampleSentences(hs) {
    var find = function (re) { return hs.filter(function (h) { return re.test(h); })[0]; };
    var out = [];
    var id = find(/\b(id|nr|no|number|nummer)\b/i) || hs[0];
    if (id) out.push(id + ' must be unique');
    var iban = find(/iban/i); if (iban) out.push(iban + ' must be valid');
    var amt = find(/amount|betrag|balance|saldo|interest|zins|tax|steuer/i); if (amt) out.push(amt + ' must not be negative');
    var name = find(/name|customer|kunde|bank/i); if (name && out.length < 4) out.push('Warn if ' + name + ' values look similar');
    return out.slice(0, 4);
  }

  function renderRules() {
    var s = sheet();
    var list = $('ruleList'); list.innerHTML = '';
    var rs = s ? rulesFor(s) : [];
    $('ruleCount').textContent = plural(state.rules.filter(function (r) { return r.enabled !== false; }).length, 'rule') + ' active';
    if (!rs.length) { list.appendChild(el('li', { class: 'empty-rules', text: 'No rules for this sheet yet. Use “Suggest rules from my data” or describe one above.' })); return; }
    rs.forEach(function (r) {
      var off = r.enabled === false;
      list.appendChild(el('li', { class: 'rule' + (off ? ' off' : '') }, [
        el('div', { class: 'what' }, [
          el('span', { class: 'col', text: r.column || 'Whole sheet' }),
          el('span', { class: 'chip', text: r.check }),
          r.param ? el('span', { class: 'param', text: String(r.param) }) : null,
          r.suggested ? el('span', { class: 'sugg', text: 'suggested' }) : null
        ]),
        el('div', { class: 'acts' }, [
          el('button', { type: 'button', class: 'sev ' + (r.severity === 'warning' ? 'warning' : 'error'), text: r.severity === 'warning' ? 'warn' : 'error',
            title: 'Switch between error and warning', onclick: function () { r.severity = r.severity === 'warning' ? 'error' : 'warning'; run(); } }),
          el('button', { type: 'button', class: 'icon-btn', text: off ? 'Turn on' : 'Turn off', onclick: function () { r.enabled = off; run(); } }),
          el('button', { type: 'button', class: 'icon-btn', text: 'Delete', 'aria-label': 'Delete rule ' + (r.description || r.check),
            onclick: function () { state.rules.splice(state.rules.indexOf(r), 1); run(); } })
        ]),
        r.description ? el('div', { class: 'desc', text: r.description }) : null
      ]));
    });
  }

  function renderSummary() {
    var st = state.result.stats, box = $('summary');
    box.innerHTML = '';
    var clean = !st.errors && !st.warnings && !st.config;
    [['e', st.errors, 'errors'], ['w', st.warnings, 'warnings'], ['c', st.config, 'rule problems'], [clean && st.rules ? 'ok' : '', st.rows, 'rows checked']].forEach(function (t) {
      box.appendChild(el('div', { class: 'tile ' + t[0] }, [el('b', { text: t[1].toLocaleString() }), el('span', { text: t[2] })]));
    });
  }

  function filtered() {
    var f = state.filter;
    return state.result.issues.filter(function (i) { return f === 'all' || i.severity === f; });
  }

  function renderIssues() {
    var st = state.result.stats;
    var fb = $('filters'); fb.innerHTML = '';
    [['all', 'All', st.errors + st.warnings + st.config], ['error', 'Errors', st.errors], ['warning', 'Warnings', st.warnings], ['config', 'Rule problems', st.config]].forEach(function (f) {
      if (f[0] === 'config' && !f[2]) return;
      fb.appendChild(el('button', { type: 'button', 'aria-pressed': String(state.filter === f[0]), text: f[1] + ' ' + f[2],
        onclick: function () { state.filter = f[0]; state.shown = ISSUE_PAGE; renderIssues(); } }));
    });
    var ul = $('issues'); ul.innerHTML = '';
    var items = filtered();
    if (!state.rules.length) { ul.appendChild(el('li', { class: 'clean', style: 'color:var(--muted)', text: 'Add a rule to start checking.' })); return; }
    if (!items.length) { ul.appendChild(el('li', { class: 'clean', text: 'No problems found.' })); return; }
    items.slice(0, state.shown).forEach(function (i) {
      var ref = cellRef(i);
      var li = el('li', { class: 'issue ' + i.severity, tabindex: '0' }, [
        el('span', { class: 'stripe' }),
        el('span', { class: 'ref' }, [ref || (i.severity === 'config' ? 'Rule' : 'Sheet'), state.sheets.length > 1 ? el('small', { text: i.sheet }) : null]),
        el('div', {}, [
          el('div', { class: 'msg', text: i.message }),
          el('div', { class: 'meta' }, [
            i.column ? el('span', { text: i.column }) : null,
            i.value !== '' && i.value != null ? el('span', { class: 'val', text: E.display(i.value) }) : null,
            el('span', { text: i.rule })
          ])
        ])
      ]);
      if (i.row) {
        var go = function () { jumpTo(i); };
        li.addEventListener('click', go);
        li.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
      }
      ul.appendChild(li);
    });
    if (items.length > state.shown) {
      ul.appendChild(el('li', { class: 'more' }, [el('button', { type: 'button', class: 'btn', text: 'Show ' + Math.min(ISSUE_PAGE, items.length - state.shown) + ' more of ' + (items.length - state.shown).toLocaleString(),
        onclick: function () { state.shown += ISSUE_PAGE; renderIssues(); } })]));
    }
  }

  function jumpTo(i) {
    var idx = state.sheets.findIndex(function (s) { return s.name === i.sheet; });
    if (idx < 0) return;
    state.current = idx;
    var s = sheet();
    var r0 = i.row - 1;                                  // zero-based row in s.rows
    var start = state.previewStart == null ? s.headerRow + 1 : state.previewStart;
    if (r0 < start || r0 >= start + PREVIEW_ROWS) state.previewStart = Math.max(s.headerRow + 1, r0 - 10);
    state.flash = r0 + ',' + i.colIndex;
    renderSheetControls(); renderRules(); renderPreview();
    var cell = document.querySelector('td[data-k="' + state.flash + '"]');
    if (cell) cell.scrollIntoView({ block: 'center', inline: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
  }

  function renderPreview() {
    var tabs = $('tabs'); tabs.innerHTML = '';
    var counts = {};
    state.result.issues.forEach(function (i) { if (i.severity !== 'config') counts[i.sheet] = (counts[i.sheet] || 0) + 1; });
    state.sheets.forEach(function (s, idx) {
      tabs.appendChild(el('button', { type: 'button', role: 'tab', 'aria-selected': String(idx === state.current),
        onclick: function () { state.current = idx; state.previewStart = null; state.flash = null; renderSheetControls(); renderRules(); renderPreview(); } },
        [s.name, counts[s.name] ? el('span', { class: 'cnt', text: counts[s.name] }) : null]));
    });

    var s = sheet(), wrap = $('gridWrap'), foot = $('previewFoot');
    wrap.innerHTML = ''; foot.innerHTML = '';
    if (!s || !s.rows.length) { wrap.appendChild(el('p', { class: 'more muted', text: 'This sheet is empty.' })); return; }
    var marks = {};
    state.result.issues.forEach(function (i) {
      if (i.sheet !== s.name || !i.row || i.colIndex < 0) return;
      var k = (i.row - 1) + ',' + i.colIndex;
      var m = marks[k] || (marks[k] = { sev: 'w', msgs: [] });
      if (i.severity !== 'warning') m.sev = 'e';
      m.msgs.push(i.message);
    });
    var width = s.rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var start = state.previewStart == null ? s.headerRow + 1 : state.previewStart;
    var end = Math.min(s.rows.length, start + PREVIEW_ROWS);

    var table = el('table', { class: 'sheet' });
    var head = el('tr', {}, [el('th', { text: '' })]);
    for (var c = 0; c < width; c++) head.appendChild(el('th', { text: W.columnLetter(c + 1) }));
    table.appendChild(el('thead', {}, [head]));
    var body = el('tbody');
    var rowsToShow = [];
    for (var h = 0; h <= s.headerRow; h++) rowsToShow.push(h);
    for (var r = start; r < end; r++) rowsToShow.push(r);
    var kinds = W.classifyRows(s.rows, s.headerRow);
    rowsToShow.forEach(function (r) {
      var kind = r > s.headerRow ? kinds[r - s.headerRow - 1] : '';
      var tr = el('tr', { class: r === s.headerRow ? 'hdr' : (kind === 'total' || kind === 'subtotal' ? 'tot' : kind === 'note' ? 'note-row' : ''),
        title: kind === 'total' || kind === 'subtotal' ? 'Total row: not checked as data' : kind === 'note' ? 'Note row: not checked as data' : false }, [el('th', { text: r + 1 })]);
      for (var c = 0; c < width; c++) {
        var v = s.rows[r][c];
        var k = r + ',' + c, m = marks[k];
        var cls = (typeof v === 'number' ? 'num ' : '') + (m ? m.sev + ' ' : '') + (state.flash === k ? 'flash' : '');
        tr.appendChild(el('td', { class: cls.trim(), 'data-k': k, title: m ? m.msgs.join('\n') : false, text: v === '' || v == null ? '' : E.display(v) }));
      }
      body.appendChild(tr);
    });
    table.appendChild(body);
    wrap.appendChild(table);

    var dataRows = kinds.filter(function (k) { return k === 'data'; }).length;
    var skipped = kinds.filter(function (k) { return k === 'total' || k === 'subtotal' || k === 'note'; }).length;
    foot.appendChild(el('span', { text: 'Rows ' + (start + 1) + '–' + end + ' of ' + s.rows.length + ' (' + plural(Math.max(0, dataRows), 'data row') + (skipped ? ', ' + plural(skipped, 'total or note row') + ' not checked as data' : '') + '). Colored cells have problems; hover for details.' }));
    if (start > s.headerRow + 1) foot.appendChild(el('button', { type: 'button', class: 'btn', text: 'Previous ' + PREVIEW_ROWS, onclick: function () { state.previewStart = Math.max(s.headerRow + 1, start - PREVIEW_ROWS); renderPreview(); } }));
    if (end < s.rows.length) foot.appendChild(el('button', { type: 'button', class: 'btn', text: 'Next ' + PREVIEW_ROWS, onclick: function () { state.previewStart = end; renderPreview(); } }));
  }

  function renderDownloads() {
    var none = downloads === null;
    $('dlXlsx').disabled = none; $('dlCsv').disabled = none;
    if (none) $('dlHint').textContent = 'Downloads are not available in this view. Open the page in the Claude app or a browser to save the marked-up workbook.';
  }

  // ---------------------------------------------------------------------------
  // adding rules
  // ---------------------------------------------------------------------------
  function status(ok, text) { var s = $('ruleStatus'); s.hidden = false; s.className = 'status ' + (ok ? 'ok' : 'bad'); s.textContent = text; }

  function describeAdded(rs) {
    return rs.map(function (r) { return (r.column || 'sheet') + ' → ' + r.check + (r.param ? ' ' + r.param : ''); }).join('; ');
  }

  function addFromText() {
    var text = $('ruleText').value.trim();
    var s = sheet();
    if (!text || !s) return;
    var btn = $('addRule');
    var local = function (note) {
      var r = E.parseRuleText(text, headers(s).filter(Boolean));
      if (!r) {
        status(false, (note ? note + ' ' : '') + 'Could not turn that into a rule. Use a column name exactly as it appears in the header, for example "' + (exampleSentences(headers(s).filter(Boolean))[0] || 'Amount must not be negative') + '", or build the rule by hand below.');
        return;
      }
      r.sheet = s.name; r.enabled = true;
      state.rules.push(r);
      $('ruleText').value = '';
      status(true, (note ? note + ' ' : '') + 'Added: ' + describeAdded([r]));
      run();
    };
    if (!sample || !$('useClaude').checked) return local();

    btn.disabled = true; btn.textContent = 'Reading…';
    var hb = headersBySheet();
    var prompt = [
      'You translate data-quality requirements written by business users (English or German) into rules for a spreadsheet checker.',
      'Available checks (name [scope] param — meaning):',
      E.checkList().map(function (c) { return '- ' + c.name + ' [' + c.scope + '] param: ' + (c.param || 'none') + ' — ' + c.help; }).join('\n'),
      '', 'Sheets and their columns:',
      Object.keys(hb).map(function (n) { return '- "' + n + '": ' + hb[n].map(function (h) { return '"' + h + '"'; }).join(', '); }).join('\n'),
      '', 'The user is looking at sheet "' + s.name + '"; prefer it when the sheet is not named.',
      'Use only the checks above and sheet/column names exactly as listed. One sentence may need several rules. For min_rows leave column empty.',
      'severity is "warning" only if the user says warn/should/hint, else "error". description: a short restatement of that one rule in the user\'s language.',
      'Reply with JSON only: {"rules":[{"sheet":"","column":"","check":"","param":"","severity":"error","description":""}],"unclear":""}',
      '', 'Requirement: ' + text
    ].join('\n');
    sample.json(prompt, { modelTier: 'quick' }).then(function (out) {
      var rs = E.validateRules(out && out.rules || [], hb);
      if (!rs.length) return local(out && out.unclear ? 'Claude: ' + out.unclear : '');
      rs.forEach(function (r) { r.enabled = true; state.rules.push(r); });
      $('ruleText').value = '';
      status(true, 'Added ' + plural(rs.length, 'rule') + ': ' + describeAdded(rs));
      run();
    }).catch(function (e) {
      if (e && e.code === 'not_granted') { sample = null; $('claudeToggle').hidden = true; }
      local(e && e.code === 'rate_limited' ? 'Claude is busy, used the built-in reader.' : '');
    }).then(function () { btn.disabled = false; btn.textContent = 'Add rule'; });
  }

  function fillChecks() {
    var sel = $('bCheck');
    E.checkList().forEach(function (c) { sel.appendChild(el('option', { value: c.name, text: c.name.replace(/_/g, ' ') })); });
    var upd = function () {
      var c = E.CHECKS[sel.value];
      $('bParam').disabled = !c.param; $('bParam').placeholder = c.param || 'No parameter needed';
      $('bHelp').textContent = c.help;
      $('bColumn').disabled = c.scope === 'table';
    };
    sel.addEventListener('change', upd); upd();
  }

  function addManual() {
    var s = sheet(); if (!s) return;
    var check = $('bCheck').value, def = E.CHECKS[check];
    var r = { sheet: s.name, column: def.scope === 'table' ? '' : $('bColumn').value, check: check,
      param: def.param ? $('bParam').value.trim() : '', severity: $('bSeverity').value, description: '', enabled: true };
    if (def.param && !r.param && check !== 'unique') { status(false, 'This check needs a parameter: ' + def.param + '.'); return; }
    state.rules.push(r);
    $('bParam').value = '';
    status(true, 'Added: ' + describeAdded([r]));
    run();
  }

  // ---------------------------------------------------------------------------
  // downloads
  // ---------------------------------------------------------------------------
  function getDownloads() {
    if (downloads !== undefined) return Promise.resolve(downloads);
    if (!window.claude || typeof window.claude.use !== 'function') { downloads = 'local'; return Promise.resolve(downloads); }
    return window.claude.use('downloads').then(function (d) { downloads = d || null; renderDownloads(); return downloads; }, function () { downloads = null; renderDownloads(); return null; });
  }

  function save(filename, data, type) {
    return getDownloads().then(function (d) {
      if (!d) return;
      if (d === 'local') {
        var url = URL.createObjectURL(new Blob([data], { type: type }));
        var a = el('a', { href: url, download: filename });
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
        return;
      }
      return d.save({ filename: filename, data: new Blob([data], { type: type }) }).catch(function (e) {
        if (e && e.code === 'declined') return;
        $('dlHint').textContent = e && e.code === 'rate_limited' ? 'A save prompt is already open.' : 'The file could not be saved here (' + (e && e.code || 'error') + ').';
      });
    });
  }

  function baseName() { return (state.file.sample ? 'sample-payouts' : state.file.name.replace(/\.[^.]+$/, '')) + '-plausi'; }

  $('dlXlsx').addEventListener('click', function () {
    var bytes = W.buildReport({ sheets: state.sheets }, state.rules, state.result.issues);
    save(baseName() + '.xlsx', bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  });
  $('dlCsv').addEventListener('click', function () {
    save(baseName() + '-report.csv', W.reportCsv(state.result.issues), 'text/csv;charset=utf-8');
  });

  // ---------------------------------------------------------------------------
  // events
  // ---------------------------------------------------------------------------
  $('fileInput').addEventListener('change', function () { if (this.files[0]) openFile(this.files[0]); this.value = ''; });
  $('sheetSelect').addEventListener('change', function () { state.current = +this.value; state.previewStart = null; state.flash = null; render(); });
  $('headerRow').addEventListener('change', function () {
    var s = sheet(); var v = Math.round(+this.value) - 1;
    if (!s || isNaN(v)) return;
    s.headerRow = Math.max(0, Math.min(s.rows.length - 1, v)); state.previewStart = null; run();
  });
  $('addRule').addEventListener('click', addFromText);
  $('ruleText').addEventListener('keydown', function (e) { if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) addFromText(); });
  $('bAdd').addEventListener('click', addManual);
  $('suggestBtn').addEventListener('click', function () {
    var s = sheet(); if (!s) return;
    var have = {};
    state.rules.forEach(function (r) { have[[r.sheet, r.column, r.check, r.param].join('|')] = true; });
    var fresh = W.suggestForSheet(s.name, s.rows, s.headerRow || 0).filter(function (r) { return !have[[r.sheet, r.column, r.check, r.param].join('|')]; });
    fresh.forEach(function (r) { state.rules.push(r); });
    status(!!fresh.length, fresh.length ? 'Added ' + plural(fresh.length, 'suggested rule') + ' for ' + s.name + '.' : 'No new suggestions for this sheet. Describe a rule instead.');
    run();
  });

  var dragDepth = 0;
  window.addEventListener('dragenter', function (e) { if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types, 'Files') > -1) { dragDepth++; $('dropzone').hidden = false; } });
  window.addEventListener('dragleave', function () { if (--dragDepth <= 0) { dragDepth = 0; $('dropzone').hidden = true; } });
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    e.preventDefault(); dragDepth = 0; $('dropzone').hidden = true;
    var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) openFile(f);
  });

  fillChecks();
  loadSample();

  // Hooks for the browser extension (tables from web pages, files from links).
  window.PlausiApp = {
    openFile: openFile,
    loadSheets: loadSheets,
    showProblem: function (msg) { showProblem(msg); }
  };

  if (window.claude && typeof window.claude.use === 'function') {
    window.claude.use('sample').then(function (s) { if (s) { sample = s; $('claudeToggle').hidden = false; } }, function () {});
    getDownloads();
  } else {
    downloads = 'local';
  }
})();
