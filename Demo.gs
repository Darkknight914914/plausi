/**
 * Plausi — demo data: a payouts sheet with planted problems that the default
 * rules catch. Pure JavaScript (no Apps Script services) so tests can use it.
 */
var PlausiDemo = (function () {
  'use strict';

  function build(now) {
    now = now || new Date();
    function d(offsetDays) {
      return new Date(now.getFullYear(), now.getMonth(), now.getDate() + offsetDays);
    }
    var iban = PlausiEngine.makeIban;
    var typo = iban('DE', '500105175407324931').replace(/931$/, '913');   // swapped digits

    var headers = ['Payout ID', 'Customer', 'IBAN', 'Email', 'Currency', 'Amount',
      'Interest rate %', 'Start date', 'End date', 'Created at'];

    var rows = [
      ['P-1001', 'Schmidt Logistik GmbH', iban('DE', '370400440532013000'), 'finance@schmidt-logistik.de', 'EUR', 12500, 3.1, d(-30), d(335), d(-2)],
      ['P-1002', 'Müller & Söhne GmbH', iban('DE', '120300000000202051'), 'buchhaltung@mueller-soehne.de', 'EUR', 8200.5, 2.9, d(-25), d(340), d(-2)],
      ['P-1003', 'Hoffmann Bau AG', iban('DE', '100100100123456789'), 'kontakt@hoffmann-bau', 'EUR', 4300, 3.0, d(-20), d(345), d(-1)],
      ['P-1004', 'Mueller und Soehne', iban('DE', '200400000628751300'), 'info@mueller-soehne.de', 'EUR', 1900, 2.9, d(-18), d(347), d(-1)],
      ['P-1005', 'Keller Consulting', typo, 'office@keller-consulting.de', 'EUR', 1500, 3.0, d(-15), d(350), d(-1)],
      ['P-1006', 'Becker Software UG', '', 'billing@becker-software.de', 'EUR', 2750, 3.3, d(-14), d(351), d(-1)],
      ['P-1003', 'Wagner Medien', iban('DE', '700202700015820011'), 'ap@wagner-medien.de', 'USD', 990, 2.7, d(-12), d(353), d(-1)],
      ['P-1008', 'Schmidt Logistik', iban('DE', '370400440532013000'), 'finance@schmidt-logistik.de', 'EUR', 640, 3.1, d(-10), d(355), d(-1)],
      ['P-1009', 'Fischer Handel KG', iban('DE', '600501010002034304'), 'zahlungen@fischer-handel.de', 'EURO', 3100, 3.0, d(-9), d(356), d(-1)],
      ['P-1010', 'Weber Immobilien', iban('DE', '250500000123456785'), 'konto@weber-immo.de', 'EUR', -250, 2.8, d(-8), d(357), d(0)],
      ['P-1011', 'Schulz Partners Ltd', iban('GB', 'WEST12345698765432'), 'accounts@schulzpartners.co.uk', 'GBP', 3333.333, 3.4, d(-7), d(358), d(0)],
      ['P-1012', 'Koch Energie GmbH', iban('DE', '430609670000001234'), 'finanzen@koch-energie.de', 'EUR', '1.250,00', 3.0, d(-6), d(359), d(0)],
      ['P-1013', 'Richter Pharma AG', iban('DE', '300500000001234567'), 'treasury@richter-pharma.de', 'EUR', 7000, 12, d(-5), d(360), d(0)],
      ['P-1014', 'Klein Textil', iban('DE', '820510000003456789'), 'info@klein-textil.de', 'EUR', 2100, 2.5, d(-10), d(-40), d(0)],
      ['P-1015', 'Wolf Maschinenbau GmbH', iban('DE', '480501610043006329'), 'buchhaltung@wolf-mb.de', 'EUR', 5600, 3.2, d(-3), d(362), d(5)],
      ['P-1016', 'Neumann Digital', iban('DE', '100700000123456700'), 'hello@neumann-digital.de', 'EUR', 2400, 3.0, d(-2), d(363), d(0)]
    ];

    var S = 'Payouts';
    var rules = [
      { sheet: S, column: 'Payout ID', check: 'required', param: '', severity: 'error', description: 'Every payout needs an ID' },
      { sheet: S, column: 'Payout ID', check: 'unique', param: '', severity: 'error', description: 'Payout IDs must be unique' },
      { sheet: S, column: 'Customer', check: 'fuzzy_duplicate', param: '', severity: 'warning', description: 'Same customer under different spellings' },
      { sheet: S, column: 'IBAN', check: 'required', param: '', severity: 'error', description: 'IBAN is mandatory' },
      { sheet: S, column: 'IBAN', check: 'iban', param: '', severity: 'error', description: 'IBAN must be valid' },
      { sheet: S, column: 'Email', check: 'email', param: '', severity: 'warning', description: 'Contact e-mail looks valid' },
      { sheet: S, column: 'Currency', check: 'allowed', param: 'EUR, USD, GBP', severity: 'error', description: 'Supported currencies only' },
      { sheet: S, column: 'Amount', check: 'number', param: '', severity: 'error', description: 'Amount must be a real number' },
      { sheet: S, column: 'Amount', check: 'range', param: '0..', severity: 'error', description: 'No negative payouts' },
      { sheet: S, column: 'Amount', check: 'decimals', param: '2', severity: 'error', description: 'Amounts in cents, max 2 decimals' },
      { sheet: S, column: 'Interest rate %', check: 'range', param: '0..5', severity: 'error', description: 'Rate between 0 and 5 %' },
      { sheet: S, column: 'End date', check: 'compare', param: '>= Start date', severity: 'error', description: 'End date not before start date' },
      { sheet: S, column: 'Created at', check: 'not_future', param: '', severity: 'error', description: 'No records from the future' },
      { sheet: S, column: 'Created at', check: 'freshness', param: '7', severity: 'warning', description: 'Export updated within the last week' },
      { sheet: S, column: '', check: 'min_rows', param: '10', severity: 'warning', description: 'Export is not suspiciously short' }
    ];

    return { sheetName: S, headers: headers, rows: rows, rules: rules };
  }

  return { build: build };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = PlausiDemo;
