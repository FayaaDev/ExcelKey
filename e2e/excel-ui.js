// Start `bun run serve`, then run:
// ego-browser nodejs -e "globalThis.excelKeyRoot = '$PWD'; await import('$PWD/e2e/excel-ui.js')"
// Set globalThis.excelKeyUrl to use another local port.
const { default: assert } = await import('node:assert/strict');
const fs = await import('node:fs/promises');
const { tmpdir } = await import('node:os');
const { join } = await import('node:path');
assert(globalThis.excelKeyRoot, 'Set globalThis.excelKeyRoot to the project directory.');
const { default: ExcelJS } = await import(join(globalThis.excelKeyRoot, 'node_modules/exceljs/excel.js'));
const dir = await fs.mkdtemp(join(tmpdir(), 'excelkey-ui-'));
const task = await taskSpace('ExcelKey UI regression');
const page = task.page('p1');
try {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Private').addRows([['Name', 'Amount'], ['Alice', 42]]);
  const original = Buffer.from(await workbook.xlsx.writeBuffer());
  const input = join(dir, 'original.xlsx');
  await fs.writeFile(input, original);
  await page.goto(globalThis.excelKeyUrl || 'http://localhost:8787');
  await page.waitForFunction(() => !document.querySelector('#controls').disabled);
  assert(await page.evaluate(() => document.querySelector('#workflow').hidden && document.querySelector('#dock').hidden), 'Initial workflow and dock stay hidden.');
  if (await page.evaluate(() => document.documentElement.lang === 'ar')) await page.click('#language-toggle');
  await page.click('#language-toggle');
  assert(await page.evaluate(async () => {
    await document.fonts.load('700 16px "Thmanyah Sans Bold"');
    return document.documentElement.lang === 'ar' && document.documentElement.dir === 'rtl'
      && document.querySelector('h1').textContent.includes('إنشاء')
      && document.fonts.check('700 16px "Thmanyah Sans Bold"');
  }), 'Arabic translates the UI, enables RTL, and loads the supplied font.');
  await page.reload();
  await page.waitForFunction(() => document.documentElement.lang === 'ar' && !document.querySelector('#controls').disabled);
  await page.click('#language-toggle');
  assert.equal(await page.evaluate(() => document.documentElement.dir), 'ltr', 'Language preference persists and switches back to English.');
  await page.click('#menu-toggle');
  await page.click('button[data-theme=dark]');
  assert.equal(await page.evaluate(() => document.documentElement.dataset.theme), 'dark');
  await page.press('button[data-theme=dark]', 'Escape');
  assert(await page.evaluate(() => document.querySelector('#app-menu').hidden), 'Escape dismisses menu.');
  await page.click('#menu-toggle');
  await page.click('button[data-theme=light]');
  await page.press('button[data-theme=light]', 'Escape');
  await page.click('.operations label:first-child');
  await page.setInputFiles('#workbook', [input]);
  assert.equal(await page.evaluate(() => document.querySelector('#workbook').tabIndex), -1, 'Hidden file input leaves the tab order; Change file stays available.');
  await page.fill('#password', 'ui-check-password');
  await page.fill('#confirmation', 'mismatch');
  await page.click('#submit');
  assert.match(await page.evaluate(() => document.querySelector('#error').textContent), /Passwords do not match/);
  await page.click('#language-toggle');
  assert.match(await page.evaluate(() => document.querySelector('#error').textContent), /غير متطابقتين/);
  assert.equal(await page.evaluate(() => document.querySelector('#password').value), 'ui-check-password', 'Language switching preserves form input.');
  await page.click('#language-toggle');
  await page.fill('#confirmation', 'ui-check-password');
  await page.click('#submit');
  await page.waitForFunction(() => !document.querySelector('#result').hidden, undefined, { timeout: 60000 });
  await page.click('#language-toggle');
  assert.match(await page.evaluate(() => document.querySelector('#result-summary').textContent), /أُنشئ/);
  assert.match(await page.evaluate(() => document.querySelector('#download').textContent), /حمّل/);
  await page.click('#language-toggle');
  assert(await page.evaluate(() => !document.querySelector('#submit').disabled && document.querySelector('#password').value === ''), 'Controls recover and passwords clear.');
  const download = page.waitForEvent('download');
  await page.click('#download');
  const protectedFile = await download;
  assert.equal(protectedFile.suggestedFilename(), 'synthetic.protected.xlsx');
  const protectedPath = join(dir, 'synthetic.protected.xlsx');
  await protectedFile.saveAs(protectedPath);
  await page.click('.operations label:last-child');
  assert(await page.evaluate(() => document.querySelector('#confirmation-field').hidden && !document.querySelector('#confirmation').required && document.querySelector('#result').hidden), 'Restore removes confirmation and stale results.');
  await page.setInputFiles('#workbook', [protectedPath]);
  await page.fill('#password', 'wrong-password');
  await page.click('#submit');
  await page.waitForFunction(() => !document.querySelector('#error').hidden && !document.querySelector('#controls').disabled, undefined, { timeout: 60000 });
  await page.fill('#password', 'ui-check-password');
  await page.click('#submit');
  await page.waitForFunction(() => !document.querySelector('#result').hidden, undefined, { timeout: 60000 });
  const restoredDownload = page.waitForEvent('download');
  await page.click('#download');
  const restored = await restoredDownload;
  const restoredPath = join(dir, 'restored.xlsx');
  await restored.saveAs(restoredPath);
  assert.deepEqual(await fs.readFile(restoredPath), original, 'Restoration returns exact original bytes.');
  for (const width of [1440, 1080, 720, 390]) {
    await page.cdp('Emulation.setDeviceMetricsOverride', { width, height: 900, deviceScaleFactor: 1, mobile: false });
    for (const language of ['ar', 'en']) {
      await page.click('#language-toggle');
      const layout = await page.evaluate(() => {
        const dock = document.querySelector('.dock-inner').getBoundingClientRect();
        const buttons = [...document.querySelectorAll('.dock-actions button')].map(button => button.getBoundingClientRect());
        return { overflow: document.documentElement.scrollWidth > innerWidth, dockFits: dock.left >= 0 && dock.right <= innerWidth, buttonsFit: buttons.every(button => button.left >= dock.left && button.right <= dock.right), equalButtons: Math.abs(buttons[0].width - buttons[1].width) < 1 };
      });
      assert(!layout.overflow && layout.dockFits && layout.buttonsFit, `${language} layout fits at ${width}px.`);
      if (width <= 720) assert(layout.equalButtons, 'Mobile dock buttons have equal widths.');
    }
  }
  console.log('PASS: Arabic toggle, font, persistence, themes, menu, modes, errors, generation, exact restoration, and bilingual layouts at four responsive widths.');
} finally {
  await task.finish({ keep: [] });
  await fs.rm(dir, { recursive: true, force: true });
}
