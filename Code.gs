/**
 * Plausi — Google Sheets add-on: menu, rules sheet, running checks, report.
 */
var PLAUSI = {
  RULES_SHEET: 'Plausi Rules',
  REPORT_SHEET: 'Plausi Report',
  RULE_HEADERS: ['Enabled', 'Sheet', 'Column', 'Check', 'Parameter', 'Severity', 'Description'],
  MAX_REPORT_ROWS: 5000,
  MAX_SIDEBAR_ISSUES: 200
};

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------
function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('Plausi')
    .addItem('Open Plausi', 'showSidebar')
    .addItem('Run checks now', 'menuRunChecks')
    .addSeparator()
    .addItem('Create rules sheet', 'menuSetupRules')
    .addItem('Load demo data', 'menuCreateDemo')
    .addToUi();
}

function onInstall(e) { onOpen(e); }

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Plausi · data checks');
  SpreadsheetApp.getUi().showSidebar(html);
}

function menuRunChecks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = runChecks_(ss);
  ss.toast(res.noRules ? 'No rules yet. Open Plausi to add some.' : summaryText_(res.stats), 'Plausi', 8);
}

function menuSetupRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ensureRulesSheet_(ss));
}

function menuCreateDemo() {
  createDemo_(SpreadsheetApp.getActiveSpreadsheet());
  showSidebar();
}

// ---------------------------------------------------------------------------
// Rules sheet
// ---------------------------------------------------------------------------
function ensureRulesSheet_(ss) {
  var sh = ss.getSheetByName(PLAUSI.RULES_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(PLAUSI.RULES_SHEET);
  var h = PLAUSI.RULE_HEADERS;
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#e8eef7');
  sh.setFrozenRows(1);
  [70, 140, 160, 130, 180, 90, 320].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  var n = sh.getMaxRows() - 1;
  sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(PlausiEngine.checkNames(), true).setAllowInvalid(false)
    .setHelpText('Pick a check type. See the Rules tab in the Plausi sidebar for what each one does.').build());
  sh.getRange(2, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['error', 'warning'], true).setAllowInvalid(false).build());
  sh.getRange(2, 5, n, 1).setNumberFormat('@');
  sh.getRange(1, 5).setNote('Parameter formats:\n' + PlausiEngine.checkList()
    .filter(function (c) { return c.param; })
    .map(function (c) { return c.name + ': ' + c.param; }).join('\n'));
  return sh;
}

function readRules_(ss) {
  var sh = ss.getSheetByName(PLAUSI.RULES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, PLAUSI.RULE_HEADERS.length).getValues();
  var rules = [];
  vals.forEach(function (r, i) {
    var check = String(r[3]).trim();
    if (!check || r[0] === false) return;           // empty row or unticked
    rules.push({
      ruleRow: i + 2,
      sheet: String(r[1]).trim(),
      column: String(r[2]).trim(),
      check: check,
      param: r[4] === '' || r[4] === null ? '' : String(r[4]).trim(),
      severity: String(r[5]).trim().toLowerCase() || 'error',
      description: String(r[6]).trim()
    });
  });
  return rules;
}

function appendRules_(ss, rules) {
  var sh = ensureRulesSheet_(ss);
  var start = sh.getLastRow() + 1;
  var vals = rules.map(function (r) {
    var param = r.param === undefined || r.param === null ? '' : String(r.param);
    if (/^[=+@]/.test(param)) param = "'" + param;    // never let a parameter become a formula
    return [true, r.sheet || '', r.column || '', r.check, param, r.severity || 'error', r.description || ''];
  });
  sh.getRange(start, 1, vals.length, 1).insertCheckboxes();
  sh.getRange(start, 5, vals.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, vals.length, PLAUSI.RULE_HEADERS.length).setValues(vals);
}

function dataSheets_(ss) {
  return ss.getSheets().filter(function (s) {
    var n = s.getName();
    return n !== PLAUSI.RULES_SHEET && n !== PLAUSI.REPORT_SHEET;
  });
}

function headersBySheet_(ss) {
  var out = {};
  dataSheets_(ss).forEach(function (s) {
    if (s.getLastColumn() < 1) return;
    out[s.getName()] = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim(); }).filter(Boolean);
  });
  return out;
}

// ---------------------------------------------------------------------------
// Running checks
// ---------------------------------------------------------------------------
function runChecks_(ss) {
  var now = new Date();
  var rules = readRules_(ss);
  var all = { issues: [], stats: { rules: 0, rows: 0, errors: 0, warnings: 0, config: 0, sheets: 0 }, ranAt: now };
  if (!rules.length) { all.noRules = true; return all; }

  var data = dataSheets_(ss);
  var groups = {};
  rules.forEach(function (r) {
    if (!r.sheet && data.length === 1) r.sheet = data[0].getName();
    (groups[r.sheet] = groups[r.sheet] || []).push(r);
  });

  Object.keys(groups).forEach(function (name) {
    var rs = groups[name];
    var sh = name ? ss.getSheetByName(name) : null;
    all.stats.rules += rs.length;
    if (!sh) {
      rs.forEach(function (r) {
        all.issues.push(PlausiEngine.configIssue(r, name ? 'Sheet "' + name + '" not found' : 'No sheet given (this file has several sheets)'));
        all.stats.config++;
      });
      return;
    }
    rs.forEach(function (r) { resolveParam_(ss, r); });
    var values = sh.getDataRange().getValues();
    var res = PlausiEngine.runRules({
      sheetName: name, headers: values[0] || [], rows: values.slice(1), firstDataRow: 2
    }, rs, { now: now });
    all.issues = all.issues.concat(res.issues);
    all.stats.rows += res.stats.rows;
    all.stats.errors += res.stats.errors;
    all.stats.warnings += res.stats.warnings;
    all.stats.config += res.stats.config;
    all.stats.sheets++;
  });

  var order = { config: 0, error: 1, warning: 2 };
  all.issues.sort(function (a, b) {
    return (order[a.severity] - order[b.severity]) || (a.sheet < b.sheet ? -1 : a.sheet > b.sheet ? 1 : 0) || ((a.row || 0) - (b.row || 0));
  });
  writeReport_(ss, all);
  PropertiesService.getDocumentProperties().setProperty('lastRun', JSON.stringify({ at: now.toISOString(), stats: all.stats }));
  return all;
}

/** sum_equals may point at a cell, e.g. "Bank!B2" or "B2" (same sheet). */
function resolveParam_(ss, rule) {
  if (rule.check !== 'sum_equals' || !rule.param || !isNaN(PlausiEngine.toNumber(rule.param))) return;
  try {
    var ref = rule.param.indexOf('!') > -1 ? rule.param : "'" + rule.sheet.replace(/'/g, "''") + "'!" + rule.param;
    rule.resolvedParam = ss.getRange(ref).getValue();
  } catch (e) { /* engine reports it as a rule problem */ }
}

function writeReport_(ss, result) {
  var sh = ss.getSheetByName(PLAUSI.REPORT_SHEET) || ss.insertSheet(PLAUSI.REPORT_SHEET);
  sh.clear();
  var head = ['Severity', 'Sheet', 'Row', 'Column', 'Value', 'Problem', 'Rule', 'Link'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
  sh.setFrozenRows(1);
  sh.getRange(1, 10).setValue('Last run: ' + Utilities.formatDate(result.ranAt, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm') +
    ' · ' + summaryText_(result.stats));

  var gids = {};
  ss.getSheets().forEach(function (s) { gids[s.getName()] = s.getSheetId(); });

  var issues = result.issues.slice(0, PLAUSI.MAX_REPORT_ROWS);
  if (!issues.length) {
    sh.getRange(2, 1).setValue('No problems found 🎉');
    return;
  }
  var colors = { config: '#fde2e1', error: '#fdecea', warning: '#fff7e0' };
  var rows = [], bgs = [];
  issues.forEach(function (i) {
    var link = '';
    if (i.row && gids[i.sheet] !== undefined) {
      var a1 = columnLetter_(i.colIndex >= 0 ? i.colIndex + 1 : 1) + i.row;
      link = '=HYPERLINK("#gid=' + gids[i.sheet] + '&range=' + a1 + '","Go to cell")';
    }
    var v = i.value;
    if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;   // show text, never evaluate it
    rows.push([i.severity === 'config' ? 'rule problem' : i.severity, i.sheet, i.row || '', i.column || '', v, i.message, i.rule, link]);
    bgs.push(new Array(head.length).fill(colors[i.severity] || '#ffffff'));
  });
  var rng = sh.getRange(2, 1, rows.length, head.length);
  rng.setValues(rows).setBackgrounds(bgs).setVerticalAlignment('top');
  sh.getRange(2, 5, rows.length, 1).setNumberFormat('@STRING@');
  [90, 110, 50, 130, 170, 360, 240, 90].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });
  if (result.issues.length > PLAUSI.MAX_REPORT_ROWS) {
    sh.getRange(rows.length + 2, 1).setValue('… ' + (result.issues.length - PLAUSI.MAX_REPORT_ROWS) + ' more problems not shown');
  }
}

function columnLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function summaryText_(st) {
  var parts = [st.errors + ' error' + (st.errors === 1 ? '' : 's'), st.warnings + ' warning' + (st.warnings === 1 ? '' : 's')];
  if (st.config) parts.push(st.config + ' rule problem' + (st.config === 1 ? '' : 's'));
  return parts.join(', ') + ' in ' + st.rows + ' rows (' + st.rules + ' rules)';
}

// ---------------------------------------------------------------------------
// Demo
// ---------------------------------------------------------------------------
function createDemo_(ss) {
  var demo = PlausiDemo.build(new Date());
  var sh = ss.getSheetByName(demo.sheetName);
  if (sh) sh.clear(); else sh = ss.insertSheet(demo.sheetName, 0);
  var rows = demo.rows.map(function (r) {
    return r.map(function (v) { return typeof v === 'string' && /^-?[\d.,]+$/.test(v) ? "'" + v : v; });
  });
  sh.getRange(1, 1, 1, demo.headers.length).setValues([demo.headers]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, demo.headers.length).setValues(rows);
  sh.getRange(2, 8, rows.length, 3).setNumberFormat('yyyy-mm-dd');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, demo.headers.length);

  var rulesSheet = ensureRulesSheet_(ss);
  if (rulesSheet.getLastRow() > 1) rulesSheet.getRange(2, 1, rulesSheet.getLastRow() - 1, PLAUSI.RULE_HEADERS.length).clearContent().removeCheckboxes();
  appendRules_(ss, demo.rules);
  ss.setActiveSheet(sh);
  ss.toast('Demo loaded: a payouts export with 13 planted problems. Click "Run checks".', 'Plausi', 8);
}

// ---------------------------------------------------------------------------
// Sidebar API (called with google.script.run)
// ---------------------------------------------------------------------------
function sidebarState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var last = PropertiesService.getDocumentProperties().getProperty('lastRun');
  return {
    settings: getSettings_(),
    checks: PlausiEngine.checkList(),
    ruleCount: readRules_(ss).length,
    sheets: Object.keys(headersBySheet_(ss)),
    lastRun: last ? JSON.parse(last) : null
  };
}

function sidebarRun() {
  var res = runChecks_(SpreadsheetApp.getActiveSpreadsheet());
  return {
    noRules: !!res.noRules,
    stats: res.stats,
    summary: res.noRules ? '' : summaryText_(res.stats),
    total: res.issues.length,
    issues: res.issues.slice(0, PLAUSI.MAX_SIDEBAR_ISSUES).map(function (i) {
      return { severity: i.severity, sheet: i.sheet, row: i.row, column: i.column, colIndex: i.colIndex,
        value: PlausiEngine.display(i.value), message: i.message, rule: i.rule };
    })
  };
}

function sidebarGoTo(sheetName, row, colIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return;
  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(row || 1, colIndex >= 0 ? colIndex + 1 : 1));
}

function sidebarOpenSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = name === 'rules' ? ensureRulesSheet_(ss) : ss.getSheetByName(PLAUSI.REPORT_SHEET);
  if (sh) ss.setActiveSheet(sh);
}

function sidebarLoadDemo() {
  createDemo_(SpreadsheetApp.getActiveSpreadsheet());
  return sidebarState();
}

/** Turns a sentence into one or more rules: Claude if a key is set, otherwise the offline parser. */
function sidebarAddRule(text) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var headers = headersBySheet_(ss);
  if (!Object.keys(headers).length) return { added: 0, message: 'Add a sheet with a header row first.' };
  var settings = getSettings_();
  var rules = [], source = 'local', aiError = '';

  if (settings.hasApiKey) {
    try {
      rules = aiRulesFromText_(text, headers);
      source = 'claude';
    } catch (e) {
      aiError = e.message;
    }
  }
  if (!rules.length) {
    source = 'local';
    var active = ss.getActiveSheet().getName();
    var names = Object.keys(headers).sort(function (a, b) { return (b === active) - (a === active); });
    for (var k = 0; k < names.length && !rules.length; k++) {
      var r = PlausiEngine.parseRuleText(text, headers[names[k]]);
      if (r && (r.column || r.check === 'min_rows')) { r.sheet = names[k]; rules.push(r); }
    }
  }
  if (!rules.length) {
    return { added: 0, aiError: aiError, message: settings.hasApiKey
      ? 'Claude could not turn that into a rule. Try naming the column exactly as in the header.'
      : 'Could not understand that. Name the column exactly as in the header, e.g. "IBAN must be valid". Adding a Claude API key in Settings understands much more.' };
  }
  appendRules_(ss, rules);
  return { added: rules.length, source: source, aiError: aiError, rules: rules };
}
