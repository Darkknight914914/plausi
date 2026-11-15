/**
 * Plausi — turn plain-language rules into structured rules with Claude.
 *
 * Privacy: only the sentence the user typed and the sheet/column NAMES are
 * sent. No cell values ever leave the spreadsheet.
 */
function aiRulesFromText_(text, headersBySheet) {
  var up = PropertiesService.getUserProperties();
  var apiKey = up.getProperty('anthropicApiKey');
  if (!apiKey) throw new Error('No Claude API key set');
  var model = up.getProperty('claudeModel') || 'claude-haiku-5-5';

  var checks = PlausiEngine.checkList().map(function (c) {
    return '- ' + c.name + ' [' + c.scope + '] param: ' + (c.param || 'none') + ' — ' + c.help;
  }).join('\n');
  var sheets = Object.keys(headersBySheet).map(function (n) {
    return '- "' + n + '": ' + headersBySheet[n].map(function (h) { return '"' + h + '"'; }).join(', ');
  }).join('\n');

  var system = [
    'You translate data-quality requirements written by business users (English or German) into rules for a spreadsheet checker.',
    'Available checks:', checks,
    '',
    'Sheets and their columns:', sheets,
    '',
    'Rules:',
    '- Use only the checks above and only sheet/column names exactly as listed.',
    '- One sentence may need several rules (e.g. "IBAN is mandatory and valid" = required + iban).',
    '- For min_rows leave column empty.',
    '- severity is "warning" only if the user says warn/hint/should; otherwise "error".',
    '- description: a short English restatement of that single rule.',
    '- If the request cannot be expressed with these checks, return no rules and explain in "unclear".',
    'Reply with JSON only, no prose: {"rules":[{"sheet":"","column":"","check":"","param":"","severity":"error","description":""}],"unclear":""}'
  ].join('\n');

  var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: model,
      max_tokens: 1500,
      system: system,
      messages: [{ role: 'user', content: text }]
    }),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Claude API error ' + code + ': ' + resp.getContentText().slice(0, 300));
  var body = JSON.parse(resp.getContentText());
  var out = (body.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('');
  var start = out.indexOf('{'), end = out.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('Claude did not return JSON');
  var parsed = JSON.parse(out.slice(start, end + 1));
  return PlausiEngine.validateRules(parsed.rules || [], headersBySheet);
}
