const fs = require('fs');
const { test, expect } = require('@playwright/test');
const { injectScript } = require('./helpers/inject');

// Track data: longitude,latitude pairs (120.5,25.1), (120.6,25.2), (120.7,25.3)
const TRACK_COORDS = '120.5,25.1,120.6,25.2,120.7,25.3';

const PAGE_HTML = `<!DOCTYPE html><html><body>
<section aria-labelledby="primary-details">
  <h3 class="Heading Heading--h3 u-margin-none">
    <a href="#"><span>香山濕地鳥類保護區</span></a>
  </h3>
</section>
<time datetime="2024-03-15T07:00:00"></time>
<div id="tracks-map-mini" data-maptrack-data="${TRACK_COORDS}"></div>
</body></html>`;

// Checklist page without GPS data
const PAGE_HTML_NO_TRACK = `<!DOCTYPE html><html><body>
<section aria-labelledby="primary-details">
  <h3 class="Heading Heading--h3 u-margin-none">
    <a href="#"><span>無軌跡地點</span></a>
  </h3>
</section>
<div id="tracks-map-mini"></div>
</body></html>`;

test.describe('/checklist/* — GPS 軌跡下載', () => {
  test('顯示「下載軌跡」按鈕', async ({ page }) => {
    await page.route('https://ebird.org/checklist/S12345', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.goto('https://ebird.org/checklist/S12345');
    await injectScript(page, { apiKey: null, spInfo: null });

    await expect(page.locator('button#ebird-track-dl')).toBeVisible();
    await expect(page.locator('button#ebird-track-dl')).toHaveText('下載軌跡');
  });

  test('下載 KML 檔案，座標與檔名正確', async ({ page }) => {
    await page.route('https://ebird.org/checklist/S12345', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.goto('https://ebird.org/checklist/S12345');
    await injectScript(page, { apiKey: null, spInfo: null });

    const downloadPromise = page.waitForEvent('download');
    await page.click('button#ebird-track-dl');
    const download = await downloadPromise;

    // Filename ends with .kml (Chinese chars may vary by OS/browser encoding)
    expect(download.suggestedFilename()).toMatch(/\.kml$/);

    const kmlPath = await download.path();
    const content = fs.readFileSync(kmlPath, 'utf8');

    // Verify KML structure
    expect(content).toContain('<?xml version="1.0"');
    expect(content).toContain('<kml');
    expect(content).toContain('<LineString>');

    // All three coordinate pairs should appear
    expect(content).toContain('120.5,25.1,0');
    expect(content).toContain('120.6,25.2,0');
    expect(content).toContain('120.7,25.3,0');

    // Placemark name should include checklist ID and date
    expect(content).toContain('S12345');
    expect(content).toContain('2024-03-15');
  });

  test('KML 包含正確的圖層樣式（紅色線條）', async ({ page }) => {
    await page.route('https://ebird.org/checklist/S12345', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.goto('https://ebird.org/checklist/S12345');
    await injectScript(page, { apiKey: null, spInfo: null });

    const downloadPromise = page.waitForEvent('download');
    await page.click('button#ebird-track-dl');
    const download = await downloadPromise;

    const content = fs.readFileSync(await download.path(), 'utf8');
    expect(content).toContain('<color>ff0000ff</color>');
    expect(content).toContain('<width>4</width>');
  });

  test('無 GPS 資料時顯示 alert 訊息', async ({ page }) => {
    await page.route('https://ebird.org/checklist/S99999', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML_NO_TRACK })
    );
    await page.goto('https://ebird.org/checklist/S99999');
    await injectScript(page, { apiKey: null, spInfo: null });

    let alertMessage = '';
    page.on('dialog', async dialog => {
      alertMessage = dialog.message();
      await dialog.dismiss();
    });

    await page.click('button#ebird-track-dl');
    await page.waitForFunction(() => true); // let event loop settle

    // Wait a tick for the alert to fire
    await page.waitForTimeout(500);
    expect(alertMessage).toContain('找不到 GPS 軌跡資料');
  });
});
