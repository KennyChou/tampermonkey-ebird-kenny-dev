const fs = require('fs');
const path = require('path');

const raw = fs.readFileSync(
  path.join(__dirname, '../../../ebird-kenny.user.js'),
  'utf8'
);

// Strip Tampermonkey header so it runs as plain JS
const USERSCRIPT = raw.replace(/\/\/ ==UserScript==[\s\S]*?\/\/ ==\/UserScript==\n/, '');

// Minimal bird taxonomy used across all tests
const DEFAULT_SP_INFO = {
  mallar3: { c: '綠頭鴨', s: 'Anas platyrhynchos', o: 1820 },
  categr1: { c: '蒼鷺', s: 'Ardea cinerea', o: 1310 },
};

// Standard CORS headers for mock API responses (cross-origin to api.ebird.org)
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Content-Type': 'application/json',
};

/**
 * Inject GM shims, XLSX mock, and the userscript into the page.
 * Call AFTER page.goto() so location.pathname is already set.
 */
async function injectScript(page, opts = {}) {
  const { apiKey = 'test-api-key', spInfo = DEFAULT_SP_INFO } = opts;

  // Seed localStorage before the script reads it
  await page.evaluate(({ k, s }) => {
    if (k) localStorage.setItem('ebirdKey', k);
    if (s) localStorage.setItem('sp_info', JSON.stringify(s));
  }, { k: apiKey, s: spInfo });

  // GM shims + XLSX mock in one script tag
  await page.addScriptTag({
    content: `
      window.GM_xmlhttpRequest = function({ method, url, headers, onload, onerror }) {
        fetch(url, { method: method || 'GET', headers: headers || {} })
          .then(r => r.text().then(text => ({ status: r.status, responseText: text })))
          .then(onload)
          .catch(err => onerror && onerror(err));
      };

      window.XLSX = {
        utils: {
          json_to_sheet: rows => ({ __rows: rows }),
          book_new: ()  => ({ __sheets: [] }),
          book_append_sheet: (wb, ws) => wb.__sheets.push(ws),
        },
        writeFile(wb, filename) {
          window.__xlsx = { rows: (wb.__sheets[0] || {}).__rows || [], filename };
        },
      };
    `,
  });

  await page.addScriptTag({ content: USERSCRIPT });
}

/**
 * Fill a route to respond with JSON and permissive CORS headers.
 */
function jsonResponse(data) {
  return { status: 200, headers: CORS_HEADERS, body: JSON.stringify(data) };
}

module.exports = { injectScript, jsonResponse, DEFAULT_SP_INFO };
