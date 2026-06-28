const { test, expect } = require('@playwright/test');
const { injectScript } = require('./helpers/inject');

const PAGE_HTML = `<!DOCTYPE html><html><body>
<fieldset id="share-contacts-fieldset"><legend>Share</legend></fieldset>
<div id="share-contacts">
  <button class="u-inset-squish-sm" onclick="addShareContact('alice@example.com')">Alice</button>
  <button class="u-inset-squish-sm" onclick="addShareContact('bob@example.com')">Bob</button>
</div>
<input id="share-recipients" type="text" value="" />
</body></html>`;

test.describe('/checklist/* — 群組名片夾', () => {
  test.beforeEach(async ({ page }) => {
    await page.route('https://ebird.org/checklist/S12345', r =>
      r.fulfill({ contentType: 'text/html', body: PAGE_HTML })
    );
    await page.goto('https://ebird.org/checklist/S12345');
    // Clear any leftover groups from previous tests
    await page.evaluate(() => localStorage.removeItem('ebird-groups'));
    await injectScript(page, { apiKey: null, spInfo: null });
  });

  test('顯示群組通訊錄標題與「全部成員」連結', async ({ page }) => {
    await expect(page.locator('text=群組通訊錄')).toBeVisible();
    await expect(page.locator('text=全部成員')).toBeVisible();
    await expect(page.locator('text=新增群組')).toBeVisible();
  });

  test('點擊「全部成員」將所有聯絡人 ID 填入分享欄位', async ({ page }) => {
    await page.click('a:has-text("全部成員")');
    const val = await page.inputValue('#share-recipients');
    expect(val).toContain('alice@example.com');
    expect(val).toContain('bob@example.com');
  });

  test('可以新增群組並顯示在清單中', async ({ page }) => {
    await page.click('a:has-text("新增群組")');
    await expect(page.locator('text=編輯群組')).toBeVisible();

    await page.fill('input[type="text"]', '核心成員');
    await page.fill('textarea', 'alice@example.com,bob@example.com');
    await page.click('button:has-text("存檔")');

    await expect(page.locator('text=核心成員')).toBeVisible();
    await expect(page.locator('text=全部成員')).not.toBeVisible();
  });

  test('群組儲存至 localStorage', async ({ page }) => {
    await page.click('a:has-text("新增群組")');
    await page.fill('input[type="text"]', '觀鳥小組');
    await page.fill('textarea', 'alice@example.com');
    await page.click('button:has-text("存檔")');

    const stored = await page.evaluate(() =>
      JSON.parse(localStorage.getItem('ebird-groups') || '[]')
    );
    expect(stored).toHaveLength(1);
    expect(stored[0].name).toBe('觀鳥小組');
    expect(stored[0].ebirdIds).toBe('alice@example.com');
  });

  test('點擊群組將成員 ID 填入分享欄位', async ({ page }) => {
    await page.click('a:has-text("新增群組")');
    await page.fill('input[type="text"]', '測試群組');
    await page.fill('textarea', 'alice@example.com');
    await page.click('button:has-text("存檔")');

    await page.click('a:has-text("測試群組")');
    const val = await page.inputValue('#share-recipients');
    expect(val).toContain('alice@example.com');
  });

  test('可以編輯既有群組', async ({ page }) => {
    // Create
    await page.click('a:has-text("新增群組")');
    await page.fill('input[type="text"]', '舊名稱');
    await page.fill('textarea', 'alice@example.com');
    await page.click('button:has-text("存檔")');

    // Edit
    await page.click('a:has-text("編輯")');
    await expect(page.locator('text=編輯群組')).toBeVisible();
    await page.fill('input[type="text"]', '新名稱');
    await page.click('button:has-text("存檔")');

    await expect(page.locator('text=新名稱')).toBeVisible();
    await expect(page.locator('text=舊名稱')).not.toBeVisible();
  });

  test('可以刪除群組', async ({ page }) => {
    await page.click('a:has-text("新增群組")');
    await page.fill('input[type="text"]', '暫時群組');
    await page.fill('textarea', 'alice@example.com');
    await page.click('button:has-text("存檔")');

    await page.click('a:has-text("編輯")');
    await page.click('button:has-text("刪除")');

    await expect(page.locator('text=暫時群組')).not.toBeVisible();
    await expect(page.locator('text=全部成員')).toBeVisible();
  });

  test('在編輯器中點擊聯絡人可追加至 ID 欄位', async ({ page }) => {
    await page.click('a:has-text("新增群組")');
    await page.fill('input[type="text"]', '動態群組');

    // Click contact "Alice" from the contact list inside the editor
    await page.locator('div[style*="max-height"] a:has-text("Alice")').click();
    const ids = await page.inputValue('textarea');
    expect(ids).toContain('alice@example.com');

    // Click "Bob" too; should append without duplicate
    await page.locator('div[style*="max-height"] a:has-text("Bob")').click();
    const ids2 = await page.inputValue('textarea');
    expect(ids2).toContain('bob@example.com');
    expect(ids2.split(',').length).toBe(2);
  });

  test('分享時不重複加入既有 ID', async ({ page }) => {
    await page.evaluate(() => {
      localStorage.setItem('ebird-groups', JSON.stringify([
        { name: 'G1', ebirdIds: 'alice@example.com' },
      ]));
    });
    await page.reload();
    await injectScript(page, { apiKey: null, spInfo: null });

    // Fill alice manually first
    await page.fill('#share-recipients', 'alice@example.com');
    // Click group (also contains alice)
    await page.click('a:has-text("G1")');

    const val = await page.inputValue('#share-recipients');
    const ids = val.split(',').map(s => s.trim()).filter(Boolean);
    const aliceCount = ids.filter(id => id === 'alice@example.com').length;
    expect(aliceCount).toBe(1);
  });
});
