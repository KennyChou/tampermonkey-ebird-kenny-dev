const { test, expect } = require('@playwright/test');
const { injectScript, jsonResponse } = require('./helpers/inject');

const PAGE_HTML = `<!DOCTYPE html><html><body>
<h2 id="my-reports-heading">我的行程報告</h2>
<ul id="my-reports-items">
  <li>
    <h4 class="Heading Heading--h4" id="TRIP001">春季賞鳥之旅</h4>
  </li>
  <li>
    <h4 class="Heading Heading--h4" id="TRIP002">秋季遷徙觀察</h4>
  </li>
</ul>
</body></html>`;

const TRIP001_CHECKLISTS = [
  { subId: 'S111111', loc: { locName: '香山濕地', lat: 24.8, lng: 120.9 } },
];

const CHECKLIST_S111111 = {
  subId: 'S111111',
  obsDt: '2024-03-10 07:00',
  locId: 'L12345',
  subnational1Code: 'TW-HSQ',
  countyCode: '',
  protocolId: 'P22',
  durationHrs: '2',
  allObsReported: true,
  effortDistanceKm: 3,
  numObservers: 3,
  obs: [{ speciesCode: 'mallar3', howManyAtmost: 12, obsAux: [], comments: '' }],
};

test.describe('/mytripreports — 行程下載', () => {
  test.beforeEach(async ({ page }) => {
    await page.route('https://ebird.org/mytripreports', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.route('**/ebird.org/tripreport-internal/v1/checklists/TRIP001', r =>
      r.fulfill(jsonResponse(TRIP001_CHECKLISTS))
    );
    await page.route('**/api.ebird.org/v2/product/checklist/view/S111111', r =>
      r.fulfill(jsonResponse(CHECKLIST_S111111))
    );
    // Block TRIP002 so we can test without extra routes
    await page.route('**/ebird.org/tripreport-internal/v1/checklists/TRIP002', r =>
      r.fulfill(jsonResponse([]))
    );

    await page.goto('https://ebird.org/mytripreports');
    await injectScript(page);
  });

  test('每個行程都注入了下載按鈕', async ({ page }) => {
    const buttons = page.locator('button.kenny-dl-trip');
    await expect(buttons).toHaveCount(2);
  });

  test('下載行程產生正確 XLSX', async ({ page }) => {
    await page.locator('button.kenny-dl-trip').first().click();

    await page.waitForFunction(() => window.__xlsx != null, { timeout: 8000 });
    const { rows, filename } = await page.evaluate(() => window.__xlsx);

    expect(filename).toBe('trip_TRIP001.xlsx');
    expect(rows).toHaveLength(1);
    expect(rows[0]['Common Name']).toBe('綠頭鴨');
    expect(rows[0]['Count']).toBe(12);
    expect(rows[0]['Location']).toBe('香山濕地');
    expect(rows[0]['Latitude']).toBe(24.8);
    expect(rows[0]['Longitude']).toBe(120.9);
  });

  test('空行程（無清單）下載後 XLSX 資料列為空', async ({ page }) => {
    await page.locator('button.kenny-dl-trip').nth(1).click();

    await page.waitForFunction(() => window.__xlsx != null, { timeout: 8000 });
    const { rows, filename } = await page.evaluate(() => window.__xlsx);

    expect(filename).toBe('trip_TRIP002.xlsx');
    expect(rows).toHaveLength(0);
  });

  test('未設定 sp_info 時不顯示下載按鈕', async ({ page }) => {
    await page.evaluate(() => localStorage.removeItem('sp_info'));
    await page.goto('https://ebird.org/mytripreports');
    await injectScript(page, { spInfo: null });

    await expect(page.locator('button.kenny-dl-trip')).toHaveCount(0);
  });
});
