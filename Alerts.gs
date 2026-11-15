/**
 * Plausi — settings, scheduled runs and alerts (e-mail, Slack, Jira).
 * Only NEW errors are alerted, so a known problem doesn't ping you every hour.
 */
function getSettings_() {
  var dp = PropertiesService.getDocumentProperties();
  var up = PropertiesService.getUserProperties();
  return {
    emailAlerts: dp.getProperty('emailAlerts') === 'true',
    alertWarnings: dp.getProperty('alertWarnings') === 'true',
    slackWebhook: dp.getProperty('slackWebhook') || '',
    jiraWebhook: dp.getProperty('jiraWebhook') || '',
    schedule: dp.getProperty('schedule') || 'off',
    model: up.getProperty('claudeModel') || 'claude-haiku-5-5',
    hasApiKey: !!up.getProperty('anthropicApiKey')
  };
}

function sidebarSaveSettings(s) {
  var dp = PropertiesService.getDocumentProperties();
  var up = PropertiesService.getUserProperties();
  ['slackWebhook', 'jiraWebhook'].forEach(function (k) {
    var v = String(s[k] || '').trim();
    if (v && !/^https:\/\//.test(v)) throw new Error('Webhook URLs must start with https://');
    dp.setProperty(k, v);
  });
  dp.setProperty('emailAlerts', String(!!s.emailAlerts));
  dp.setProperty('alertWarnings', String(!!s.alertWarnings));
  if (s.model) up.setProperty('claudeModel', String(s.model).trim());
  if (s.apiKey) up.setProperty('anthropicApiKey', String(s.apiKey).trim());
  if (s.clearApiKey) up.deleteProperty('anthropicApiKey');
  setSchedule_(s.schedule || 'off');
  dp.setProperty('schedule', s.schedule || 'off');
  return getSettings_();
}

function setSchedule_(freq) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledRun') ScriptApp.deleteTrigger(t);
  });
  var b = ScriptApp.newTrigger('scheduledRun').timeBased();
  if (freq === 'hourly') b.everyHours(1).create();
  else if (freq === 'daily') b.everyDays(1).atHour(7).create();
  else if (freq === 'weekly') b.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

/** Time-driven entry point. */
function scheduledRun() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = runChecks_(ss);
  if (!res.noRules) sendAlerts_(ss, res, false);
}

function sidebarTestAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sample = {
    issues: [{ severity: 'error', sheet: 'Example', row: 2, column: 'IBAN', value: 'DE00…', message: 'Test alert from Plausi', rule: 'Test', check: 'iban' }],
    stats: { errors: 1, warnings: 0, config: 0, rows: 1, rules: 1 }
  };
  return sendAlerts_(ss, sample, true);
}

function hashKey_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .slice(0, 6).map(function (b) { return ((b + 256) % 256).toString(16).replace(/^(.)$/, '0$1'); }).join('');
}

function sendAlerts_(ss, res, isTest) {
  var s = getSettings_();
  var dp = PropertiesService.getDocumentProperties();
  var relevant = res.issues.filter(function (i) { return i.severity !== 'warning' || s.alertWarnings; });

  var fresh = relevant;
  if (!isTest) {
    var prev = {};
    JSON.parse(dp.getProperty('alertedKeys') || '[]').forEach(function (k) { prev[k] = true; });
    var keys = relevant.map(function (i) { return hashKey_(PlausiEngine.issueKey(i)); });
    fresh = relevant.filter(function (i, idx) { return !prev[keys[idx]]; });
    dp.setProperty('alertedKeys', JSON.stringify(keys.slice(0, 500)));
  }
  if (!fresh.length) return { sent: [], newIssues: 0 };

  var url = ss.getUrl() + '#gid=' + (ss.getSheetByName(PLAUSI.REPORT_SHEET) || ss.getSheets()[0]).getSheetId();
  var title = (isTest ? '[TEST] ' : '') + 'Plausi: ' + fresh.length + ' new data problem' + (fresh.length === 1 ? '' : 's') + ' in "' + ss.getName() + '"';
  var lines = fresh.slice(0, 15).map(function (i) {
    return '• ' + i.sheet + (i.row ? ' row ' + i.row : '') + (i.column ? ', ' + i.column : '') + ': ' + i.message;
  });
  if (fresh.length > 15) lines.push('… and ' + (fresh.length - 15) + ' more');
  var body = title + '\n\n' + lines.join('\n') + '\n\nFull report: ' + url;
  var sent = [];

  if (s.emailAlerts) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), title, body);
    sent.push('e-mail');
  }
  if (s.slackWebhook) {
    postJson_(s.slackWebhook, { text: '*' + title + '*\n' + lines.join('\n') + '\n<' + url + '|Open report>' });
    sent.push('Slack');
  }
  if (s.jiraWebhook) {
    // Shape for a Jira Automation "Incoming webhook" rule: use {{webhookData.summary}} etc.
    postJson_(s.jiraWebhook, {
      summary: title,
      description: body,
      spreadsheet: ss.getName(),
      url: url,
      newIssues: fresh.length,
      errors: res.stats.errors,
      warnings: res.stats.warnings,
      issues: fresh.slice(0, 25).map(function (i) {
        return { sheet: i.sheet, row: i.row, column: i.column, value: PlausiEngine.display(i.value), problem: i.message, rule: i.rule, severity: i.severity };
      })
    });
    sent.push('Jira');
  }
  return { sent: sent, newIssues: fresh.length };
}

function postJson_(url, payload) {
  var r = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (r.getResponseCode() >= 300) throw new Error('Webhook returned ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
}
