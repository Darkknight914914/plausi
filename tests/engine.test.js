// Run with: node tests/engine.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'Engine.gs'), 'utf8'));
vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'Demo.gs'), 'utf8'));

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}

const now = new Date(2026, 9, 18, 18, 0, 0);
const demo = PlausiDemo.build(now);
const result = PlausiEngine.runRules(
  { sheetName: demo.sheetName, headers: demo.headers, rows: demo.rows, firstDataRow: 2 },
  demo.rules, { now });
const has = (row, check, re) => result.issues.some(i => i.row === row && i.check === check && (!re || re.test(i.message)));

console.log('Demo run');
test('no rule configuration problems', () => assert.strictEqual(result.stats.config, 0, JSON.stringify(result.issues.filter(i => i.severity === 'config'))));
test('valid IBANs pass', () => assert.ok(!has(2, 'iban') && !has(12, 'iban')));
test('IBAN typo caught (row 6)', () => assert.ok(has(6, 'iban', /checksum/)));
test('missing IBAN caught (row 7)', () => assert.ok(has(7, 'required')));
test('invalid e-mail caught (row 4)', () => assert.ok(has(4, 'email')));
test('duplicate payout ID caught (row 8)', () => assert.ok(has(8, 'unique', /row 4/)));
test('Müller & Söhne GmbH ~ Mueller und Soehne (row 5)', () => assert.ok(has(5, 'fuzzy_duplicate', /row 3/)));
test('Schmidt Logistik ~ Schmidt Logistik GmbH (row 9)', () => assert.ok(has(9, 'fuzzy_duplicate', /row 2/)));
test('EURO not an allowed currency (row 10)', () => assert.ok(has(10, 'allowed')));
test('negative amount (row 11)', () => assert.ok(has(11, 'range', /minimum/)));
test('3 decimals (row 12)', () => assert.ok(has(12, 'decimals')));
test('number stored as text (row 13)', () => assert.ok(has(13, 'number', /stored as text/)));
test('text number still range-checked fine (row 13)', () => assert.ok(!has(13, 'range')));
test('interest rate 12 % out of range (row 14)', () => assert.ok(has(14, 'range', /maximum/)));
test('end date before start date (row 15)', () => assert.ok(has(15, 'compare')));
test('created in the future (row 16)', () => assert.ok(has(16, 'not_future')));
test('fresh export and enough rows', () => assert.ok(!result.issues.some(i => i.check === 'freshness' || i.check === 'min_rows')));
test('clean row has no issues (row 17)', () => assert.ok(!result.issues.some(i => i.row === 17)));
test('exact issue counts', () => {
  assert.strictEqual(result.stats.errors, 10, JSON.stringify(result.issues.map(i => [i.row, i.check, i.message])));
  assert.strictEqual(result.stats.warnings, 3);
});

console.log('Value parsing');
test('German number text', () => assert.strictEqual(PlausiEngine.toNumber('1.250,00'), 1250));
test('English number text', () => assert.strictEqual(PlausiEngine.toNumber('1,250.50'), 1250.5));
test('decimal comma', () => assert.strictEqual(PlausiEngine.toNumber('3,5'), 3.5));
test('invalid date rejected', () => assert.strictEqual(PlausiEngine.toDate('31.02.2026'), null));
test('German date parsed', () => assert.strictEqual(PlausiEngine.toDate('01.03.2026').getMonth(), 2));
test('known IBAN valid', () => assert.strictEqual(PlausiEngine.ibanProblem('DE89 3704 0044 0532 0130 00'), null));
test('wrong-length IBAN', () => assert.ok(/22 characters/.test(PlausiEngine.ibanProblem('DE8937040044053201300'))));

console.log('Column rules');
test('sum_equals reconciliation', () => {
  const r = PlausiEngine.runRules({ sheetName: 'S', headers: ['Amount'], rows: [[100], [50.5], ['']] },
    [{ column: 'Amount', check: 'sum_equals', param: '160' }], { now });
  assert.ok(/difference -9.5/.test(r.issues[0].message), r.issues[0].message);
});
test('unique across a column combination', () => {
  const r = PlausiEngine.runRules({ sheetName: 'S', headers: ['ID', 'Currency'], rows: [['A', 'EUR'], ['A', 'USD'], ['A', 'EUR']] },
    [{ column: 'ID', check: 'unique', param: 'Currency' }], { now });
  assert.strictEqual(r.issues.length, 1); assert.strictEqual(r.issues[0].row, 4);
});
test('stale export flagged', () => {
  const r = PlausiEngine.runRules({ sheetName: 'S', headers: ['Day'], rows: [[new Date(2026, 9, 1)]] },
    [{ column: 'Day', check: 'freshness', param: 7 }], { now });
  assert.ok(/17 days old/.test(r.issues[0].message), r.issues[0].message);
});
test('unknown column becomes a rule problem, not a crash', () => {
  const r = PlausiEngine.runRules({ sheetName: 'S', headers: ['A'], rows: [[1]] }, [{ column: 'B', check: 'required' }], { now });
  assert.strictEqual(r.issues[0].severity, 'config');
});

console.log('Plain-language rules (offline parser)');
const H = demo.headers;
const cases = [
  ['IBAN must be valid', 'iban', 'IBAN', ''],
  ['IBAN is mandatory', 'required', 'IBAN', ''],
  ['Payout ID must be unique', 'unique', 'Payout ID', ''],
  ['Interest rate % must be between 0 and 5', 'range', 'Interest rate %', '0..5'],
  ['Amount must not be negative', 'range', 'Amount', '0..'],
  ['Amount at most 10,000', 'range', 'Amount', '..10000'],
  ['Currency must be one of EUR, USD or GBP', 'allowed', 'Currency', 'EUR, USD, GBP'],
  ['End date must be on or after Start date', 'compare', 'End date', '>= Start date'],
  ['Created at must not be in the future', 'not_future', 'Created at', ''],
  ['Warn if customer names look similar', 'fuzzy_duplicate', 'Customer', ''],
  ['Warn if Customer names look similar', 'fuzzy_duplicate', 'Customer', ''],
  ['Amount max 2 decimal places', 'decimals', 'Amount', '2'],
  ['Sheet should have at least 10 rows', 'min_rows', '', '10'],
  ['Created at should be updated within 3 days', 'freshness', 'Created at', '3'],
  ['Summe von Amount muss Bank!B2 entsprechen', 'sum_equals', 'Amount', 'Bank!B2'],
  ['Zinssatz egal', null, null, null],
  ['Currency darf nicht leer sein', 'required', 'Currency', ''],
  ['Email muss gültig sein', 'email', 'Email', ''],
  ['Payout ID must match ^P-\\d{4}$', 'regex', 'Payout ID', '^P-\\d{4}$']
];
cases.forEach(([text, check, column, param]) => {
  test(text, () => {
    const r = PlausiEngine.parseRuleText(text, H);
    if (check === null) { assert.strictEqual(r, null, JSON.stringify(r)); return; }
    assert.ok(r, 'no rule parsed');
    assert.deepStrictEqual([r.check, r.column, r.param], [check, column, param]);
  });
});
test('warning severity detected', () => assert.strictEqual(PlausiEngine.parseRuleText('Warn if Customer names look similar', H).severity, 'warning'));

console.log('Validating rules from Claude');
test('drops unknown checks and columns, fixes case', () => {
  const v = PlausiEngine.validateRules([
    { sheet: 'payouts', column: 'iban', check: 'IBAN' },
    { sheet: 'Payouts', column: 'Nope', check: 'required' },
    { sheet: 'Payouts', column: 'Amount', check: 'teleport' }
  ], { Payouts: H });
  assert.deepStrictEqual(v.map(r => [r.sheet, r.column, r.check]), [['Payouts', 'IBAN', 'iban']]);
});

console.log('\n' + passed + ' tests passed' + (process.exitCode ? ', some FAILED' : ''));
