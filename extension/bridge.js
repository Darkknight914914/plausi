// Plausi extension — hands data from the right-click menus to the checker page.
(function () {
  'use strict';
  window.PLAUSI_XLSX_HINT = 'The Excel reader is missing from the extension, so only CSV files and web tables work. ' +
    'Add lib/xlsx.full.min.js as described in the extension README, then reload the extension.';

  var TYPES = {
    'text/csv': 'csv', 'text/tab-separated-values': 'tsv', 'text/plain': 'csv',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
    'application/vnd.ms-excel.sheet.macroenabled.12': 'xlsm',
    'application/vnd.ms-excel.sheet.binary.macroenabled.12': 'xlsb',
    'application/vnd.ms-excel': 'xls',
    'application/vnd.oasis.opendocument.spreadsheet': 'ods'
  };

  function prompt(text, buttonText, onClick) {
    var box = document.getElementById('extPrompt');
    if (!box) {
      box = document.createElement('div');
      box.id = 'extPrompt';
      box.className = 'note sample';
      var anchor = document.querySelector('.layout');
      anchor.parentNode.insertBefore(box, anchor);
    }
    box.innerHTML = '';
    var span = document.createElement('span');
    span.textContent = text;
    box.appendChild(span);
    if (buttonText) {
      var b = document.createElement('button');
      b.type = 'button'; b.className = 'btn primary'; b.textContent = buttonText;
      b.addEventListener('click', onClick);
      box.appendChild(b);
    }
    box.hidden = false;
  }
  function clearPrompt() { var b = document.getElementById('extPrompt'); if (b) b.hidden = true; }

  function fileNameFor(url, res) {
    var cd = res.headers.get('content-disposition') || '';
    var m = /filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i.exec(cd);
    var name = m ? decodeURIComponent(m[1] || m[2]) : decodeURIComponent((new URL(url).pathname.split('/').pop() || ''));
    if (!/\.[a-z0-9]{2,5}$/i.test(name)) {
      var type = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
      name = (name || 'download') + '.' + (TYPES[type] || 'xlsx');
    }
    return name;
  }

  function download(url) {
    prompt('Downloading ' + url + ' …');
    fetch(url, { credentials: 'include' }).then(function (res) {
      if (!res.ok) throw new Error('the server answered ' + res.status + (res.status === 401 || res.status === 403 ? ' (are you signed in to that site?)' : ''));
      return res.blob().then(function (blob) {
        var type = (res.headers.get('content-type') || '').toLowerCase();
        if (/text\/html/.test(type)) throw new Error('the link opens a web page, not a file. Open the page and use “Check tables on this page” instead');
        clearPrompt();
        window.PlausiApp.openFile(new File([blob], fileNameFor(url, res)));
      });
    }).catch(function (e) {
      prompt('Plausi could not download the file: ' + e.message + '. You can also download it yourself and drop it here.');
    });
  }

  function checkLink(url) {
    var origin;
    try { origin = new URL(url).origin; } catch (e) { return prompt('That link is not a web address Plausi can open.'); }
    if (!/^https?:/.test(origin)) return prompt('Plausi can only download files from http or https links.');
    var perm = { origins: [origin + '/*'] };
    chrome.permissions.contains(perm, function (has) {
      if (has) return download(url);
      prompt('Plausi needs your permission to download files from ' + new URL(url).host + '. The file is read here and not sent anywhere.',
        'Allow and check the file', function () {
          chrome.permissions.request(perm, function (granted) {
            if (granted) download(url);
            else prompt('Permission was not given. Download the file yourself and drop it on this page.');
          });
        });
    });
  }

  if (!(window.chrome && chrome.storage && chrome.storage.session) || location.hash !== '#pending') return;
  chrome.storage.session.get('pending', function (got) {
    var p = got && got.pending;
    if (!p) return;
    chrome.storage.session.remove('pending');
    history.replaceState(null, '', location.pathname);
    if (p.kind === 'error') return window.PlausiApp.showProblem(p.message);
    if (p.kind === 'link') return checkLink(p.url);
    if (p.kind === 'tables') {
      if (!p.tables.length) {
        return window.PlausiApp.showProblem('No tables found on “' + p.source + '”. Plausi reads tables and data grids; charts, pictures and PDFs can’t be checked. If the report has an Excel or CSV export, download that and drop it here.');
      }
      window.PlausiApp.loadSheets(p.source, 'web page', p.tables.map(function (t) {
        return { name: t.name, rows: t.rows, textInput: true };
      }));
    }
  });
})();
