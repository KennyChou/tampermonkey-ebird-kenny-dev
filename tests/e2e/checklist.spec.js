const { test, expect } = require('@playwright/test');
const { injectScript, jsonResponse } = require('./helpers/inject');

const PAGE_HTML = `<!DOCTYPE html><html><body>
<form id="myChecklistsForm"></form>
<div id="place-species-observed-results">
  <div class="ResultsStats ResultsStats--manageMyChecklists" id="result-S111111-0">
    <div class="Color-text-neutral-4 u-text-3">2024-01-15</div>
  </div>
  <div class="ResultsStats ResultsStats--manageMyChecklists" id="result-S222222-0">
    <div class="Color-text-neutral-4 u-text-3">2024-01-16</div>
  </div>
</div>
</body></html>`;

const CHECKLIST_S111111 = {
  subId: 'S111111',
  obsDt: '2024-01-15 08:30',
  locId: 'L12345',
  subnational1Code: 'TW-PIF',
  countyCode: 'TW-PIF-07',
  protocolId: 'P22',
  durationHrs: '1.5',
  allObsReported: true,
  effortDistanceKm: 2.5,
  numObservers: 2,
  obs: [{ speciesCode: 'mallar3', howManyAtmost: 5, obsAux: [], comments: '水邊' }],
};

const CHECKLIST_S222222 = {
  ...CHECKLIST_S111111,
  subId: 'S222222',
  obsDt: '2024-01-16 09:00',
  obs: [{ speciesCode: 'categr1', howManyAtmost: 1, obsAux: [], comments: '' }],
};

const HOTSPOT_L12345 = { locId: 'L12345', name: '香山濕地', latitude: 24.8, longitude: 120.9 };

test.describe('/mychecklists — 清單下載', () => {
  test.beforeEach(async ({ page }) => {
    await page.route('https://ebird.org/mychecklists', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.route('**/api.ebird.org/v2/product/checklist/view/S111111', r =>
      r.fulfill(jsonResponse(CHECKLIST_S111111))
    );
    await page.route('**/api.ebird.org/v2/product/checklist/view/S222222', r =>
      r.fulfill(jsonResponse(CHECKLIST_S222222))
    );
    await page.route('**/api.ebird.org/v2/ref/hotspot/info/L12345', r =>
      r.fulfill(jsonResponse(HOTSPOT_L12345))
    );

    await page.goto('https://ebird.org/mychecklists');
    await injectScript(page);
  });

  test('每筆清單都注入了勾選框', async ({ page }) => {
    await expect(page.locator('.kenny-checkbox')).toHaveCount(2);
  });

  test('勾選後顯示下載按鈕且計數正確', async ({ page }) => {
    const dlBtn = page.locator('button', { hasText: '下載選取的清單' });
    await expect(dlBtn).toBeHidden();

    await page.locator('.kenny-checkbox').first().check();
    await expect(dlBtn).toBeVisible();
    await expect(dlBtn).toContainText('(1)');

    await page.locator('.kenny-checkbox').nth(1).check();
    await expect(dlBtn).toContainText('(2)');
  });

  test('取消勾選後下載按鈕消失', async ({ page }) => {
    const cb = page.locator('.kenny-checkbox').first();
    await cb.check();
    await cb.uncheck();
    await expect(page.locator('button', { hasText: '下載選取的清單' })).toBeHidden();
  });

  test('下載單筆清單產生正確 XLSX 資料', async ({ page }) => {
    await page.locator('.kenny-checkbox').first().check();
    await page.click('button:has-text("下載選取的清單")');

    await page.waitForFunction(() => window.__xlsx != null, { timeout: 8000 });
    const { rows, filename } = await page.evaluate(() => window.__xlsx);

    expect(filename).toMatch(/^checklist_\d+\.xlsx$/);
    expect(rows).toHaveLength(1);
    expect(rows[0]['Common Name']).toBe('綠頭鴨');
    expect(rows[0]['Count']).toBe(5);
    expect(rows[0]['Location']).toBe('香山濕地');
    expect(rows[0]['Date']).toBe('2024-01-15');
    expect(rows[0]['Submission ID']).toBe('S111111');
  });

  test('下載兩筆清單資料列數正確', async ({ page }) => {
    await page.locator('.kenny-checkbox').first().check();
    await page.locator('.kenny-checkbox').nth(1).check();
    await page.click('button:has-text("下載選取的清單")');

    await page.waitForFunction(() => window.__xlsx != null, { timeout: 10000 });
    const { rows } = await page.evaluate(() => window.__xlsx);

    expect(rows).toHaveLength(2);
    const codes = rows.map(r => r['Submission ID']);
    expect(codes).toContain('S111111');
    expect(codes).toContain('S222222');
  });

  test('未設定 API Key 時不顯示勾選框', async ({ page }) => {
    await page.evaluate(() => localStorage.removeItem('ebirdKey'));
    // Re-run the script without a key
    await page.goto('https://ebird.org/mychecklists');
    await injectScript(page, { apiKey: null });

    await expect(page.locator('.kenny-checkbox')).toHaveCount(0);
  });
});
