// Run with: node tests/web.test.js [output.xlsx]
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

for (const f of ['Engine.gs', 'Demo.gs', 'web/core.js']) vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', f), 'utf8'));

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}

console.log('CSV');
test('semicolon CSV with BOM, quotes and German numbers', () => {
  const r = PlausiWeb.parseCsv('﻿ID;Name;Amount\r\n1;"Müller; Söhne";"1.250,00"\r\n2;"He said ""hi""";3,5\r\n');
  assert.strictEqual(r.delimiter, ';');
  assert.deepStrictEqual(r.rows, [['ID', 'Name', 'Amount'], ['1', 'Müller; Söhne', '1.250,00'], ['2', 'He said "hi"', '3,5']]);
});
test('comma CSV, no trailing newline, multiline field', () => {
  const r = PlausiWeb.parseCsv('a,b\n1,"x\ny"');
  assert.strictEqual(r.delimiter, ',');
  assert.deepStrictEqual(r.rows, [['a', 'b'], ['1', 'x\ny']]);
});
test('tab separated', () => assert.strictEqual(PlausiWeb.parseCsv('a\tb\tc\n1\t2\t3').delimiter, '\t'));
test('Excel sep= line', () => {
  const r = PlausiWeb.parseCsv('sep=,\na;b,c\n');
  assert.deepStrictEqual(r.rows, [['a;b', 'c']]);
});
test('CSV rows check cleanly in text mode', () => {
  const r = PlausiWeb.parseCsv('Amount;Start;End\n1.250,00;01.03.2026;31.03.2026\n-5;01.03.2026;01.02.2026\nabc;31.02.2026;\n');
  const res = PlausiEngine.runRules({ sheetName: 'csv', headers: r.rows[0], rows: r.rows.slice(1) }, [
    { column: 'Amount', check: 'number' }, { column: 'Amount', check: 'range', param: '0..' },
    { column: 'End', check: 'compare', param: '>= Start' }, { column: 'Start', check: 'date' }], { textInput: true });
  const got = res.issues.map(i => [i.row, i.check]);
  assert.deepStrictEqual(got.sort(), [[3, 'compare'], [3, 'range'], [4, 'date'], [4, 'number'], [4, 'range']].sort());
});

console.log('Header detection & rules sheet');
test('skips title rows', () => {
  assert.strictEqual(PlausiWeb.detectHeaderRow([['Payout export Oct 2026'], [''], ['ID', 'Amount'], ['P-1', 5]]), 2);
});
test('reads a Plausi Rules sheet (add-on format)', () => {
  const rules = PlausiWeb.rulesFromRows([
    ['Enabled', 'Sheet', 'Column', 'Check', 'Parameter', 'Severity', 'Description'],
    [true, 'Payouts', 'IBAN', 'iban', '', 'error', 'IBAN valid'],
    [false, 'Payouts', 'Amount', 'range', '0..', 'warning', ''],
    ['', '', '', '', '', '', '']]);
  assert.strictEqual(rules.length, 2);
  assert.strictEqual(rules[1].enabled, false);
  assert.strictEqual(rules[1].severity, 'warning');
});

console.log('SheetJS adapter (with a stand-in for the library)');
const fakeXLSX = {
  utils: {
    decode_range: () => ({ s: { r: 0, c: 0 }, e: { r: 2, c: 3 } }),
    encode_cell: ({ r, c }) => String.fromCharCode(65 + c) + (r + 1)
  },
  SSF: {
    is_date: z => /y|d/.test(z),
    parse_date_code: v => (v === 46313 ? { y: 2026, m: 10, d: 18, H: 0, M: 0, S: 0 } : null)
  }
};
test('converts date-formatted numbers, booleans, errors; drops empty trailing rows', () => {
  const ws = {
    '!ref': 'A1:D3',
    A1: { t: 's', v: 'Day' }, B1: { t: 's', v: 'Amt' }, C1: { t: 's', v: 'Active' }, D1: { t: 's', v: 'Calc' },
    A2: { t: 'n', v: 46313, z: 'dd.mm.yyyy' }, B2: { t: 'n', v: 12.5, z: '#,##0.00' }, C2: { t: 'b', v: true }, D2: { t: 'e', v: 15, w: '#VALUE!' }
  };
  const rows = PlausiWeb.sheetToRows(fakeXLSX, ws, false);
  assert.strictEqual(rows.length, 2);
  assert.strictEqual(PlausiEngine.display(rows[1][0]), '2026-10-18');
  assert.deepStrictEqual(rows[1].slice(1), [12.5, true, '#VALUE!']);
});

console.log('Excel writer');
const now = new Date(2026, 9, 18, 12);
const demo = PlausiDemo.build(now);
const book = { sheets: [{ name: demo.sheetName, rows: [demo.headers].concat(demo.rows), headerRow: 0 }] };
const res = PlausiEngine.runRules({ sheetName: demo.sheetName, headers: demo.headers, rows: demo.rows }, demo.rules, { now });
const bytes = PlausiWeb.buildReport(book, demo.rules, res.issues);
test('produces a zip', () => assert.deepStrictEqual(Array.from(bytes.slice(0, 4)), [0x50, 0x4b, 0x03, 0x04]));
test('crc32 known value', () => assert.strictEqual(PlausiWeb.crc32(Buffer.from('hello')), 0x3610a686));
test('sheet names are sanitised and unique', () => {
  const { names } = PlausiWeb.buildXlsx([{ name: 'a/b', rows: [] }, { name: 'a b', rows: [] }, { name: 'x'.repeat(40), rows: [] }]);
  assert.deepStrictEqual(names, ['a b', 'a b (2)', 'x'.repeat(31)]);
});
test('report CSV escapes and neutralises formulas', () => {
  const csv = PlausiWeb.reportCsv([{ severity: 'error', sheet: 'S', row: 2, colIndex: 0, column: 'A', value: '=1+1', message: 'x, "y"', rule: 'r' }]);
  assert.ok(csv.includes("'=1+1") && csv.includes('"x, ""y"""'), csv);
});

const out = process.argv[2];
if (out) { fs.writeFileSync(out, bytes); console.log('wrote ' + out + ' (' + res.issues.length + ' issues)'); }

console.log('Suggested rules (data shaped like the deposits screenshot)');
{
  const H = ['Sr no', 'Deposit ID', 'Registration Number', 'Bank Name', 'Is Active Deposit', 'Deposit Balance Year End', 'Interest Earned', 'Interest Earned Eur', 'Paid Source Tax'];
  const R = [
    [1, 'ACCT_856_976_972_983', 'CUST_300_942_462_988', 'Bank Name 1001', true, 42354.82, 349.61, 2490.7, 395.4],
    [2, 'ACCT_393_972_920_969', 'CUST_568_992_436_986', 'Bank Name 1002', false, 32781.56, 249.54, 1170.15, 256.57],
    [3, 'ACCT_393_972_920_969', 'CUST_568_992_436_986', 'Bank Name 1003', false, 32781.56, 249.54, 1170.15, 256.57],
    [4, 'ACCT_572_953_947_866', 'CUST_937_985_499_991', 'Bank Name 1004', false, 31269.99, 477.02, 1770.78, 245.62],
    [5, 'ACCT_989_990_992_876', 'CUST_842_902_447_927', 'Bank Name 1005', false, 0, 409.91, 1008.43, 367.07],
    [6, 'ACCT_300_966_977_860', 'CUST_902_962_475_972', 'Bank Name 1006', false, 5753.56, 4718, 0, 480.3],
    [7, 'ACCT_111_222_333_444', 'CUST_111_222_333_444', 'Bank Name 1007', true, -12.5, 10.123, 5, '']
  ];
  const sug = PlausiWeb.suggestRules('Sheet1', H, R);
  const has = (col, check) => sug.some(r => r.column === col && r.check === check);
  test('suggests unique Deposit ID and Sr no', () => assert.ok(has('Deposit ID', 'unique') && has('Sr no', 'unique')));
  test('TRUE/FALSE column gets allowed list', () => assert.ok(has('Is Active Deposit', 'allowed')));
  test('money columns get not-negative and 2 decimals', () => assert.ok(has('Interest Earned', 'range') && has('Interest Earned', 'decimals')));
  test('mostly-filled column gets required', () => assert.ok(has('Paid Source Tax', 'required')));
  const res = PlausiEngine.runRules({ sheetName: 'Sheet1', headers: H, rows: R }, sug, { now: new Date() });
  const hit = (row, col, check) => res.issues.some(i => i.row === row && i.column === col && i.check === check);
  test('finds the duplicated deposit (row 4)', () => assert.ok(hit(4, 'Deposit ID', 'unique')));
  test('finds the missing tax (row 8)', () => assert.ok(hit(8, 'Paid Source Tax', 'required')));
  test('finds 3 decimals (row 8)', () => assert.ok(hit(8, 'Interest Earned', 'decimals')));
  test('no rule problems', () => assert.strictEqual(res.stats.config, 0));
  test('nothing for an empty sheet', () => assert.deepStrictEqual(PlausiWeb.suggestRules('S', ['A'], []), []));
  console.log('       ' + sug.length + ' suggestions → ' + res.issues.length + ' findings');
}
console.log('Suggestions for other kinds of data');
{
  const H = ['Invoice no', 'Customer', 'Currency', 'Term', 'Bic', 'Start date', 'End date'];
  const R = [];
  for (let i = 0; i < 30; i++) R.push([`R-${i}`, i % 3 ? 'Acme GmbH' : 'Beta AG', 'EUR', [12, 24, 36][i % 3], 'DEUTDEFF', new Date(2026, 0, 1), new Date(2026, 5, 1)]);
  R[4][1] = 'ACME Gmbh.'; R[5][2] = 'EURO'; R[6][3] = 13; R[7][4] = 'DEUT'; R[8][6] = new Date(2025, 0, 1); R[9][0] = 'R-1';
  const sug = PlausiWeb.suggestRules('S', H, R);
  const res = PlausiEngine.runRules({ sheetName: 'S', headers: H, rows: R }, sug, { now: new Date(2026, 9, 18) });
  const hit = (row, col, check) => res.issues.some(i => i.row === row && i.column === col && i.check === check);
  test('invoice number duplicate', () => assert.ok(hit(11, 'Invoice no', 'unique')));
  test('customer spelled differently, repeats not flagged', () => {
    assert.ok(hit(6, 'Customer', 'fuzzy_duplicate'));
    assert.strictEqual(res.issues.filter(i => i.check === 'fuzzy_duplicate').length, 1);
  });
  test('currency code typo', () => assert.ok(hit(7, 'Currency', 'allowed')));
  test('unusual term value', () => assert.ok(hit(8, 'Term', 'allowed')));
  test('bad BIC', () => assert.ok(hit(9, 'Bic', 'bic')));
  test('end date before start date', () => assert.ok(hit(10, 'End date', 'compare')));
}
console.log('Messy reports (any columns)');
{
  const rows = [
    ['Regional sales report', '', '', '', '', ''],
    ['Generated 18.10.2026 by Finance', '', '', '', '', ''],
    [],
    ['Order ref', 'Region', '', 'Units', 'Revenue', 'Revenue', 'Margin %'],
  ];
  const regions = ['North', 'South', 'East', 'West'];
  for (let i = 0; i < 24; i++) rows.push([`SO-${2000 + i}`, regions[i % 4], 'x', 10 + (i % 7), 1000 + i * 10, 900 + i * 10, 0.2 + (i % 5) / 100]);
  rows[6][0] = 'SO-20O2';           // letter O instead of zero
  rows[8][1] = ' South';            // stray space
  rows[10][3] = 900;                // outlier
  rows[12][6] = 25;                 // percent entered as 25 instead of 0.25
  const sum = rows.slice(4).reduce((a, r) => a + r[4], 0);
  rows.push(['Total', '', '', '', sum + 50, '', '']);   // total that doesn't add up
  rows.push(['* Revenue excludes VAT. Figures are preliminary and subject to change.']);
  const h = PlausiWeb.detectHeaderRow(rows);
  test('finds the header under the title rows', () => assert.strictEqual(h, 3));
  test('names blank and repeated columns', () => assert.deepStrictEqual(PlausiWeb.namedHeaders(rows[h], 7), ['Order ref', 'Region', 'Column C', 'Units', 'Revenue', 'Revenue (2)', 'Margin %']));
  const kinds = PlausiWeb.classifyRows(rows, h);
  test('total and note rows are not data', () => assert.deepStrictEqual(kinds.slice(-2), ['total', 'note']));
  const sug = PlausiWeb.suggestForSheet('Sales', rows, h);
  const named = PlausiWeb.namedHeaders(rows[h], 7);
  const resolved = sug.map(r => r.check === 'sum_equals' ? Object.assign({}, r, { resolvedParam: PlausiWeb.cellValue(rows, r.param) }) : r);
  const res = PlausiEngine.runRules({ sheetName: 'Sales', headers: named, rows: PlausiWeb.dataOnly(rows, h), firstDataRow: h + 2 }, resolved, { now: new Date() });
  const hit = (row, col, check) => res.issues.some(i => (row == null || i.row === row) && i.column === col && i.check === check);
  test('format break SO-20O2', () => assert.ok(hit(7, 'Order ref', 'regex')));
  test('stray space', () => assert.ok(hit(9, 'Region', 'no_extra_spaces')));
  test('unusual units value', () => assert.ok(hit(11, 'Units', 'outlier') || hit(11, 'Units', 'allowed')));
  test('percentage out of range', () => assert.ok(hit(13, 'Margin %', 'range')));
  test('total row does not add up', () => assert.ok(hit(null, 'Revenue', 'sum_equals')));
  test('total row itself is not flagged as a data problem', () => assert.ok(!res.issues.some(i => i.row === rows.length - 1)));
  test('no rule problems', () => assert.strictEqual(res.stats.config, 0, JSON.stringify(res.issues.filter(i => i.severity === 'config'))));
  console.log('       ' + sug.length + ' suggestions → ' + res.issues.length + ' findings: ' + res.issues.map(i => i.column + ':' + i.check + '@' + (i.row || '-')).join(', '));
}
console.log('\n' + passed + ' tests passed' + (process.exitCode ? ', some FAILED' : ''));
