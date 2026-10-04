// Start `bun run serve`, then run:
// ego-browser nodejs -e "globalThis.excelKeyRoot = '$PWD'; await import('$PWD/e2e/excel-ui.js')"
// Set excelKeyUrl for another port, excelKeySpace = 1 to resume the parent space,
// and excelKeyKeepSpace = true to leave it open for screenshot inspection.
const { default: assert } = await import('node:assert/strict');
const fs = await import('node:fs/promises');
const { tmpdir } = await import('node:os');
const { join } = await import('node:path');
const { pathToFileURL } = await import('node:url');
const { execFile } = await import('node:child_process');
const { promisify } = await import('node:util');
assert(globalThis.excelKeyRoot, 'Set globalThis.excelKeyRoot to the project directory.');
const { default: ExcelJS } = await import(join(globalThis.excelKeyRoot, 'node_modules/exceljs/excel.js'));
const dir = await fs.mkdtemp(join(tmpdir(), 'excelkey-ui-'));
const screenshots = await fs.mkdtemp('/var/folders/2q/1rqmfh096px8z_t_bpn660dr0000gn/T/opencode/excelkey-ui-');
const task = await taskSpace(globalThis.excelKeySpace ?? 'ExcelKey UI regression');
const page = task.page('p1');
console.log({ spaceId: task.spaceId, screenshots });
const waitIdle = () => page.waitForFunction(() => !document.querySelector('#controls').disabled
  && !document.querySelector('#submit').disabled && !document.querySelector('#change-file').disabled
  && !document.querySelector('#workbook-form').hasAttribute('aria-busy'), undefined, { timeout: 60000 });
async function submitSuccess() {
  await page.click('#submit');
  await page.waitForFunction(() => !document.querySelector('#result').hidden, undefined, { timeout: 60000 });
  await waitIdle();
  assert(await page.evaluate(() => document.querySelector('#error').hidden));
}
async function submitError(pattern) {
  await page.click('#submit');
  await page.waitForFunction(() => !document.querySelector('#error').hidden, undefined, { timeout: 60000 });
  await waitIdle();
  assert.match(await page.evaluate(() => document.querySelector('#error').textContent), pattern);
  assert(await page.evaluate(() => document.querySelector('#result').hidden
    && !document.querySelector('#download').hasAttribute('href')
    && document.querySelector('#recovery-backup').hidden
    && document.querySelector('#backup-phrase').textContent === ''), 'Errors clear results and recovery backup.');
}
async function saveResult(name) {
  const pending = page.waitForEvent('download');
  await page.click('#download');
  const download = await pending;
  const path = join(dir, name);
  await download.saveAs(path);
  return { path, name: download.suggestedFilename() };
}
const observation = () => page.evaluate(() => globalThis.excelKeyWebAuthn.calls);
const inject = fault => page.evaluate(value => { globalThis.excelKeyWebAuthn.fault = value; }, fault);
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
      && document.querySelector('.brand').textContent.trim() === 'تشفيرة الاكسل'
      && document.title === 'تشفيرة الاكسل — إنشاء تشفيرة ببيانات اصطناعية'
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
  assert.equal(await page.evaluate(() => document.documentElement.dataset.theme), 'light');
  await page.cdp('WebAuthn.disable');
  await page.cdp('WebAuthn.enable');
  await page.cdp('WebAuthn.addVirtualAuthenticator', { options: {
    protocol: 'ctap2', ctap2Version: 'ctap2_1', transport: 'internal',
    hasResidentKey: true, hasUserVerification: true, hasPrf: true,
    isUserVerified: true, automaticPresenceSimulation: true,
  } });
  // Native calls are observed, never replaced on successful passkey paths.
  await page.evaluate(() => {
    const saved = [
      [navigator.credentials, 'create'], [navigator.credentials, 'get'],
      [PublicKeyCredential, 'getClientCapabilities'],
    ].map(([object, key]) => ({ object, key, descriptor: Object.getOwnPropertyDescriptor(object, key), original: object[key] }));
    const state = globalThis.excelKeyWebAuthn = { calls: [], fault: null };
    state.restore = () => {
      for (const { object, key, descriptor, original } of saved) {
        if (descriptor && !descriptor.configurable) object[key] = original;
        else if (descriptor) Object.defineProperty(object, key, descriptor);
        else delete object[key];
      }
      delete globalThis.excelKeyWebAuthn;
    };
    for (const { object, key, original } of saved) {
      object[key] = async function (options) {
        if (key === 'getClientCapabilities') {
          if (state.fault === 'capability-false') return { 'extension:prf': false };
          if (state.fault === 'capability-unknown') return {};
          return original ? original.call(this) : {};
        }
        const call = { method: key, allowCredentials: (options.publicKey.allowCredentials ?? []).map(item => Array.from(new Uint8Array(item.id))) };
        state.calls.push(call);
        if (state.fault === `${key}-cancel`) throw new DOMException('Injected cancellation', 'NotAllowedError');
        if (key === 'create' && state.fault === 'enrollment-prf-false') return { getClientExtensionResults: () => ({ prf: { enabled: false } }) };
        if (key === 'get' && state.fault === 'no-prf') return { getClientExtensionResults: () => ({}) };
        if (key === 'get' && state.fault === 'wrong-key') return { getClientExtensionResults: () => ({ prf: { results: { first: new Uint8Array(32).fill(173).buffer } } }) };
        // Bypass installed passkey-manager wrappers so CDP's native authenticator handles tests.
        const credential = await CredentialsContainer.prototype[key].call(this, options);
        call.id = Array.from(new Uint8Array(credential.rawId));
        const prf = credential.getClientExtensionResults().prf;
        call.prfEnabled = prf?.enabled;
        call.prfBytes = prf?.results?.first?.byteLength;
        return credential;
      };
    }
  });
  await page.click('.operations label:first-child');
  await page.setInputFiles('#workbook', [input]);
  assert.equal(await page.evaluate(() => document.querySelector('#workbook').tabIndex), -1, 'Hidden file input leaves the tab order; Change file stays available.');
  await page.selectOption('#protect-method', 'create');
  assert(await page.evaluate(() => document.querySelector('#password').disabled && document.querySelector('#password-field').hidden));
  if (!await page.evaluate(() => document.querySelector('#backup-enabled').checked)) await page.click('#backup-enabled');
  await submitSuccess();
  const enrolled = await observation();
  assert.deepEqual(enrolled.map(call => call.method), ['create', 'get'], 'Native enrollment precedes the PRF assertion.');
  assert.equal(enrolled[0].prfEnabled, true);
  assert.equal(enrolled[1].prfBytes, 32);
  assert.equal(enrolled[1].allowCredentials.length, 1, 'Enrollment verification pins one credential.');
  assert.deepEqual(enrolled[1].allowCredentials[0], enrolled[0].id);
  assert.deepEqual(enrolled[1].id, enrolled[0].id);
  assert(await page.evaluate(() => !document.querySelector('#recovery-backup').hidden));
  const phrase = await page.evaluate(() => document.querySelector('#backup-phrase').textContent.trim());
  assert.equal(phrase.split(/\s+/).length, 24, 'Capture the actual 24-word backup.');
  await page.click('#language-toggle');
  assert.match(await page.evaluate(() => document.querySelector('#result-summary').textContent), /أُنشئ/);
  assert.match(await page.evaluate(() => document.querySelector('#download').textContent), /حمّل/);
  await page.click('#language-toggle');
  const protectedFile = await saveResult('synthetic.protected.xlsx');
  assert.equal(protectedFile.name, 'synthetic.protected.xlsx');
  await page.click('#hide-backup');
  assert(await page.evaluate(() => document.querySelector('#recovery-backup').hidden && document.querySelector('#backup-phrase').textContent === ''), 'Hide clears the backup text.');
  await page.click('.operations label:last-child');
  await page.selectOption('#restore-method', 'passkey');
  assert(await page.evaluate(() => document.querySelector('#protect-method-field').hidden && document.querySelector('#result').hidden), 'Restore switches controls and clears stale results.');
  await page.setInputFiles('#workbook', [protectedFile.path]);
  await submitSuccess();
  const returning = await observation();
  assert.equal(returning.length, 3);
  assert.equal(returning[2].method, 'get');
  assert.deepEqual(returning[2].allowCredentials, [], 'Returning assertion uses the discoverable picker.');
  assert.deepEqual(returning[2].id, enrolled[0].id, 'Restoration uses the enrolled passkey.');
  assert.equal(returning[2].prfBytes, 32);
  assert.deepEqual(await fs.readFile((await saveResult('passkey-restored.xlsx')).path), original, 'Native passkey restores exact bytes.');
  await page.selectOption('#restore-method', 'recovery');
  await page.fill('#recovery-phrase', 'invalid phrase');
  await submitError(/Invalid recovery phrase/);
  await page.fill('#recovery-phrase', phrase);
  await page.click('#language-toggle');
  assert.equal(await page.evaluate(() => document.querySelector('#recovery-phrase').value), phrase, 'Language switching preserves input.');
  await page.click('#language-toggle');
  await submitSuccess();
  assert.deepEqual(await observation(), returning, 'Phrase restoration never calls WebAuthn.');
  assert.equal(await page.evaluate(() => document.querySelector('#recovery-phrase').value), '', 'Recovery input clears after use.');
  assert.deepEqual(await fs.readFile((await saveResult('recovery-restored.xlsx')).path), original, 'Recovery phrase restores exact bytes.');
  await page.selectOption('#restore-method', 'passkey');
  for (const [fault, pattern] of [
    ['get-cancel', /cancelled or timed out/], ['capability-false', /does not support PRF/],
    ['no-prf', /does not support PRF/], ['wrong-key', /Incorrect passkey or recovery phrase/],
  ]) {
    const before = (await observation()).length;
    await inject(fault);
    try {
      await submitError(pattern);
      if (fault === 'capability-false') assert.equal((await observation()).length, before, 'Explicit PRF false prevents WebAuthn.');
      if (fault === 'get-cancel') {
        await page.click('#language-toggle');
        assert.match(await page.evaluate(() => document.querySelector('#error').textContent), /أُلغي/);
        await page.click('#language-toggle');
      }
    } finally { await inject(null); }
  }
  await page.click('.operations label:first-child');
  await page.setInputFiles('#workbook', [input]);
  await page.selectOption('#protect-method', 'create');
  for (const [fault, pattern] of [
    ['capability-false', /does not support PRF/], ['create-cancel', /cancelled or timed out/],
    ['enrollment-prf-false', /does not support PRF/],
  ]) {
    const before = (await observation()).length;
    await inject(fault);
    try {
      await submitError(pattern);
      assert.deepEqual((await observation()).slice(before).map(call => call.method),
        fault === 'capability-false' ? [] : ['create'], 'Failed enrollment never proceeds to assertion.');
    } finally { await inject(null); }
  }
  await page.selectOption('#protect-method', 'existing');
  const beforeExisting = (await observation()).length;
  await inject('capability-unknown');
  try { await submitSuccess(); } finally { await inject(null); }
  const existing = (await observation()).slice(beforeExisting);
  assert.equal(existing.length, 1);
  assert.equal(existing[0].method, 'get');
  assert.equal(existing[0].prfBytes, 32, 'Unknown capability still succeeds with native PRF.');
  assert.deepEqual(existing[0].allowCredentials, []);
  assert.deepEqual(existing[0].id, enrolled[0].id);
  assert(await page.evaluate(() => document.querySelector('#recovery-backup').hidden && document.querySelector('#backup-phrase').textContent === ''), 'Existing-key protection does not reveal a backup.');
  const existingFile = await saveResult('existing.protected.xlsx');
  await page.click('.operations label:last-child');
  await page.selectOption('#restore-method', 'passkey');
  await page.setInputFiles('#workbook', [existingFile.path]);
  await submitSuccess();
  assert.deepEqual(await fs.readFile((await saveResult('existing-restored.xlsx')).path), original);

  // Build password v2 through the core only; the UI never offers password protection.
  const legacyPath = join(dir, 'legacy-v2.xlsx');
  await promisify(execFile)('bun', ['-e', `
    const { transformWorkbook } = await import(${JSON.stringify(pathToFileURL(join(globalThis.excelKeyRoot, 'src/excel.ts')).href)});
    const { hasV2Protection } = await import(${JSON.stringify(pathToFileURL(join(globalThis.excelKeyRoot, 'src/excel-ooxml.ts')).href)});
    const original = new Uint8Array(await Bun.file(${JSON.stringify(input)}).arrayBuffer());
    const result = await transformWorkbook(original, 'ui-check-password', 'protect', { filename: 'original.xlsx' });
    if (!hasV2Protection(result.data)) throw new Error('Expected a legacy v2 fixture');
    await Bun.write(${JSON.stringify(legacyPath)}, result.data);
  `], { cwd: globalThis.excelKeyRoot, timeout: 60000 });
  await page.selectOption('#restore-method', 'password');
  await page.setInputFiles('#workbook', [legacyPath]);
  const beforeLegacy = await observation();
  await page.fill('#password', 'wrong-password');
  await submitError(/Incorrect password/);
  await page.fill('#password', 'ui-check-password');
  await submitSuccess();
  assert.equal(await page.evaluate(() => document.querySelector('#password').value), '');
  assert.deepEqual(await observation(), beforeLegacy, 'Legacy password restoration never calls WebAuthn.');
  assert.deepEqual(await fs.readFile((await saveResult('legacy-restored.xlsx')).path), original, 'Legacy v2 restores exact bytes.');
  const beforeNoPassword = await observation();
  await page.click('.operations label:first-child');
  await page.selectOption('#protect-method', 'none');
  await page.setInputFiles('#workbook', [input]);
  assert(await page.evaluate(() => !document.querySelector('#no-password-help').hidden
    && document.querySelector('#backup-option').hidden && document.querySelector('#passkey-help').hidden
    && document.querySelector('#password').disabled && document.querySelector('#recovery-phrase').disabled));
  await submitSuccess();
  assert(await page.evaluate(() => document.querySelector('#recovery-backup').hidden));
  const noPasswordFile = await saveResult('no-password.xlsx');
  await page.click('.operations label:last-child');
  await page.selectOption('#restore-method', 'none');
  await page.setInputFiles('#workbook', [protectedFile.path]);
  await submitError(/requires a passkey or recovery phrase/);
  await page.setInputFiles('#workbook', [noPasswordFile.path]);
  await submitSuccess();
  assert.deepEqual(await fs.readFile((await saveResult('no-password-restored.xlsx')).path), original, 'No-password mode restores exact bytes.');
  assert.deepEqual(await observation(), beforeNoPassword, 'No-password operations never call WebAuthn.');
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
      if (width === 1440 || width === 390) {
        const path = join(screenshots, `${width === 390 ? 'mobile' : 'desktop'}-${language}.png`);
        await page.screenshot({ path });
        console.log({ screenshot: path });
      }
    }
  }
  console.log('PASS: native enrollment and pinned PRF, discoverable restore, recovery backup, injected errors, legacy v2, no-password round-trip, exact bytes, themes, Arabic, and four responsive widths.');
} finally {
  try { await page.evaluate(() => globalThis.excelKeyWebAuthn?.restore()); }
  finally { await fs.rm(dir, { recursive: true, force: true }); }
}
if (!globalThis.excelKeyKeepSpace) await task.finish({ keep: [] });
