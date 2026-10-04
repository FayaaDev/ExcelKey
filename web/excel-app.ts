import { EXCEL_EXTENSIONS, MAX_FILE_BYTES, transformWorkbook } from '../src/excel';
import { masterPrkFromPrfSecret } from '../src/identity';
import { decodeRecoveryBip39, encodeRecoveryBip39 } from '../src/recovery';
import { checkSupport, deploymentRpId, enrollPasskey, getPrfSecret, prfBrowserSupport } from './webauthn';
import { translations } from './excel-i18n';

const form = document.querySelector<HTMLFormElement>('#workbook-form')!;
const controls = document.querySelector<HTMLFieldSetElement>('#controls')!;
const fileInput = document.querySelector<HTMLInputElement>('#workbook')!;
const password = document.querySelector<HTMLInputElement>('#password')!;
const showPassword = document.querySelector<HTMLInputElement>('#show-password')!;
const protectMethod = document.querySelector<HTMLSelectElement>('#protect-method')!;
const restoreMethod = document.querySelector<HTMLSelectElement>('#restore-method')!;
const recoveryPhrase = document.querySelector<HTMLTextAreaElement>('#recovery-phrase')!;
const backupEnabled = document.querySelector<HTMLInputElement>('#backup-enabled')!;
const recoveryBackup = document.querySelector<HTMLElement>('#recovery-backup')!;
const backupPhrase = document.querySelector<HTMLElement>('#backup-phrase')!;
const hideBackup = document.querySelector<HTMLButtonElement>('#hide-backup')!;
const submit = document.querySelector<HTMLButtonElement>('#submit')!;
const status = document.querySelector<HTMLElement>('#status')!;
const error = document.querySelector<HTMLElement>('#error')!;
const result = document.querySelector<HTMLElement>('#result')!;
const download = document.querySelector<HTMLAnchorElement>('#download')!;
const workflow = document.querySelector<HTMLElement>('#workflow')!;
const dock = document.querySelector<HTMLElement>('#dock')!;
const changeFile = document.querySelector<HTMLButtonElement>('#change-file')!;
const fileCard = document.querySelector<HTMLElement>('#file-card')!;
const menuToggle = document.querySelector<HTMLButtonElement>('#menu-toggle')!;
const menu = document.querySelector<HTMLElement>('#app-menu')!;
let downloadUrl: string | undefined;
let busy = false;
let activeRoot: Uint8Array | undefined;
let pageRevision = 0;
const cryptoAvailable = window.isSecureContext && !!globalThis.crypto?.subtle;
try { if (localStorage.getItem('excelkey-passkey-used') === 'true') protectMethod.value = 'existing'; } catch { /* A preference, never a secret. */ }
const languageToggle = document.querySelector<HTMLButtonElement>('#language-toggle')!;
let language = 'en';
try { language = localStorage.getItem('excelkey-language') === 'ar' ? 'ar' : 'en'; } catch { /* Storage can be unavailable. */ }

function setText(element: HTMLElement, english: string, arabic?: string): void {
  element.dataset.i18n = english;
  if (arabic) element.dataset.i18nAr = arabic;
  else delete element.dataset.i18nAr;
  element.textContent = language === 'ar' ? arabic ?? translations[english] ?? english : english;
}

function applyLanguage(): void {
  document.documentElement.lang = language;
  document.documentElement.dir = language === 'ar' ? 'rtl' : 'ltr';
  document.querySelectorAll<HTMLElement>('[data-i18n]').forEach(element => {
    setText(element, element.dataset.i18n!, element.dataset.i18nAr);
  });
  document.querySelectorAll<HTMLElement>('[data-i18n-aria-label]').forEach(element => {
    const label = element.dataset.i18nAriaLabel!;
    element.setAttribute('aria-label', language === 'ar' ? translations[label] ?? label : label);
  });
  const description = document.querySelector<HTMLMetaElement>('meta[name=description]')!;
  description.dataset.english ??= description.content;
  description.content = language === 'ar' ? translations[description.dataset.english]! : description.dataset.english;
  languageToggle.textContent = language === 'ar' ? 'English' : 'العربية';
  languageToggle.lang = language === 'ar' ? 'en' : 'ar';
  languageToggle.setAttribute('aria-label', language === 'ar' ? translations['Switch to English']! : 'Switch to Arabic');
}

languageToggle.addEventListener('click', () => {
  language = language === 'ar' ? 'en' : 'ar';
  try { localStorage.setItem('excelkey-language', language); } catch { /* Keep the choice for this page. */ }
  applyLanguage();
});
applyLanguage();

function closeMenu(): void {
  menu.hidden = true;
  menuToggle.setAttribute('aria-expanded', 'false');
}

menuToggle.addEventListener('click', () => {
  menu.hidden = !menu.hidden;
  menuToggle.setAttribute('aria-expanded', String(!menu.hidden));
});
document.addEventListener('click', event => {
  if (!(event.target as Element).closest('.navigation')) closeMenu();
});
document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && !menu.hidden) {
    closeMenu();
    menuToggle.focus();
  }
});
menu.querySelectorAll('a').forEach(link => link.addEventListener('click', () => {
  if (link.getAttribute('href') === '#workbook-details') {
    document.querySelector<HTMLDetailsElement>('#workbook-details')!.open = true;
  }
  closeMenu();
}));

const systemTheme = matchMedia('(prefers-color-scheme: dark)');
let appearance = 'auto';
try { appearance = localStorage.getItem('excelkey-appearance') ?? 'auto'; } catch { /* Storage can be unavailable. */ }
function applyAppearance(): void {
  if (!['light', 'dark', 'auto'].includes(appearance)) appearance = 'auto';
  const dark = appearance === 'dark' || (appearance === 'auto' && systemTheme.matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  document.querySelector<HTMLMetaElement>('meta[name=theme-color]')!.content = dark ? '#0c0c0e' : '#ffffff';
  menu.querySelectorAll<HTMLButtonElement>('[data-theme]').forEach(button => {
    button.setAttribute('aria-pressed', String(button.dataset.theme === appearance));
  });
}
menu.querySelectorAll<HTMLButtonElement>('[data-theme]').forEach(button => button.addEventListener('click', () => {
  appearance = button.dataset.theme!;
  try { localStorage.setItem('excelkey-appearance', appearance); } catch { /* Keep the choice for this page. */ }
  applyAppearance();
}));
systemTheme.addEventListener('change', applyAppearance);
applyAppearance();

changeFile.addEventListener('click', () => fileInput.click());

function updateFile(): void {
  const file = fileInput.files?.[0];
  fileCard.hidden = !file;
  fileInput.classList.toggle('visually-hidden', !!file);
  fileInput.tabIndex = file ? -1 : 0;
  setText(changeFile, file ? 'Change file' : 'Choose file');
  if (file) {
    document.querySelector<HTMLElement>('#file-name')!.textContent = file.name;
    document.querySelector<HTMLElement>('#file-name')!.title = file.name;
    const size = (file.size / 1024).toLocaleString('en', { maximumFractionDigits: 1 });
    setText(document.querySelector<HTMLElement>('#file-meta')!, `${size} KiB · Local workbook`, `${size} كيبيبايت · تشفيرة محلية`);
  }
}

function clearResult(): void {
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = undefined;
  download.removeAttribute('href');
  result.hidden = true;
  error.hidden = true;
  setText(status, '');
}

function mode(): 'protect' | 'restore' {
  return document.querySelector<HTMLInputElement>('input[name=mode]:checked')?.value === 'restore' ? 'restore' : 'protect';
}

function clearBackup(): void {
  backupPhrase.textContent = '';
  recoveryBackup.hidden = true;
}

function clearSecrets(): void {
  activeRoot?.fill(0);
  activeRoot = undefined;
  password.value = recoveryPhrase.value = '';
  showPassword.checked = false;
  password.type = 'password';
  clearBackup();
}

function updateMethod(): void {
  const protect = mode() === 'protect';
  const legacy = !protect && restoreMethod.value === 'password';
  const recovery = !protect && restoreMethod.value === 'recovery';
  const noPassword = (protect ? protectMethod : restoreMethod).value === 'none';
  document.querySelector<HTMLElement>('#passkey-help')!.hidden = !protect || noPassword;
  document.querySelector<HTMLElement>('#no-password-help')!.hidden = !noPassword;
  document.querySelector<HTMLElement>('#protect-method-field')!.hidden = !protect;
  document.querySelector<HTMLElement>('#restore-method-field')!.hidden = protect;
  protectMethod.disabled = !protect;
  restoreMethod.disabled = protect;
  document.querySelector<HTMLElement>('#password-field')!.hidden = !legacy;
  document.querySelector<HTMLElement>('#recovery-field')!.hidden = !recovery;
  password.disabled = showPassword.disabled = !legacy;
  password.required = legacy;
  recoveryPhrase.disabled = !recovery;
  recoveryPhrase.required = recovery;
  const setup = protect && protectMethod.value === 'create';
  document.querySelector<HTMLElement>('#backup-option')!.hidden = !setup;
  backupEnabled.disabled = !setup;
  let action = 'Restore with passkey';
  if (protect) action = setup ? 'Create passkey' : 'Protect with passkey';
  else if (legacy) action = 'Restore with password';
  else if (recovery) action = 'Restore with recovery phrase';
  if (noPassword) action = protect ? 'Generate without password' : 'Restore without password';
  setText(submit, action);
  setText(document.querySelector<HTMLElement>('#dock-title')!, protect ? 'Generate synthetic workbook' : 'Restore original workbook');
}

hideBackup.addEventListener('click', () => {
  clearBackup();
  download.focus({ preventScroll: true });
});

form.addEventListener('change', event => {
  if (busy) return;
  if ((event.target as HTMLInputElement).name === 'mode' || event.target === protectMethod || event.target === restoreMethod) {
    workflow.hidden = dock.hidden = false;
    updateMethod();
    submit.disabled = !cryptoAvailable;
    clearSecrets();
    clearResult();
  } else if (event.target === fileInput) {
    clearSecrets();
    clearResult();
    updateFile();
  }
});

showPassword.addEventListener('change', () => {
  password.type = showPassword.checked ? 'text' : 'password';
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (busy) return;
  clearResult();
  clearBackup();
  const operation = mode();
  const file = fileInput.files?.[0];
  const revision = pageRevision;
  const setup = operation === 'protect' && protectMethod.value === 'create';
  const noPassword = (operation === 'protect' ? protectMethod : restoreMethod).value === 'none';
  const showBackup = setup && backupEnabled.checked;
  let passkeyPending = false;
  try {
    if (!cryptoAvailable) throw new Error('Open ExcelKey over HTTPS or localhost to use encryption.');
    if (!file || !EXCEL_EXTENSIONS.test(file.name)) throw new Error('Choose an Excel workbook.');
    if (file.size > MAX_FILE_BYTES) throw new Error('Choose a workbook smaller than 64 MiB.');
    busy = true;
    controls.disabled = true;
    submit.disabled = changeFile.disabled = true;
    form.setAttribute('aria-busy', 'true');
    const legacy = operation === 'restore' && restoreMethod.value === 'password';
    if (legacy) {
      if (!password.value) throw new Error('Enter a password.');
    } else if (operation === 'restore' && restoreMethod.value === 'recovery') {
      try { activeRoot = decodeRecoveryBip39(recoveryPhrase.value.toLowerCase()); }
      catch { throw new Error('Invalid recovery phrase. Enter all 24 English words in order and check their spelling.'); }
      recoveryPhrase.value = '';
    } else if (!noPassword) {
      passkeyPending = true;
      if (!checkSupport().webauthn) throw new Error('This browser does not support passkeys. Use a supported browser, or restore with a recovery phrase or legacy password.');
      if (await prfBrowserSupport() === false) throw new Error('This browser or passkey provider does not support PRF encryption. Try another provider or browser, or restore with your recovery phrase.');
      if (revision !== pageRevision) return;
      setText(status, setup ? 'Create a passkey, then verify it to protect this workbook.' : 'Choose your original passkey and verify to continue.');
      const credentialId = setup ? await enrollPasskey('ExcelKey') : undefined;
      if (revision !== pageRevision) return;
      const secret = await getPrfSecret(credentialId);
      try {
        if (revision !== pageRevision) return;
        activeRoot = masterPrkFromPrfSecret(secret);
      } finally { secret.fill(0); }
      passkeyPending = false;
      protectMethod.value = 'existing';
      try { localStorage.setItem('excelkey-passkey-used', 'true'); } catch { /* Only remember the preferred action. */ }
    }
    setText(status, operation === 'protect' ? 'Generating synthetic workbook… Keep this tab open.' : 'Restoring workbook… Keep this tab open.');
    const input = new Uint8Array(await file.arrayBuffer());
    if (revision !== pageRevision) return;
    const transformed = await transformWorkbook(input, noPassword ? null : legacy ? password.value : { masterPrk: activeRoot!, namespace: deploymentRpId() }, operation, { filename: file.name });
    if (revision !== pageRevision) return;
    if (showBackup) {
      backupPhrase.textContent = encodeRecoveryBip39(activeRoot!);
      recoveryBackup.hidden = false;
    }
    downloadUrl = URL.createObjectURL(new Blob([new Uint8Array(transformed.data)], {
      type: 'application/octet-stream',
    }));
    download.href = downloadUrl;
    const name = (transformed.filename ?? file.name).replace(/^.*[/\\]/, '');
    const extension = /\.[^.]+$/.exec(name)?.[0] ?? '.xlsx';
    const base = name.slice(0, name.length - extension.length).replace(/\.(protected|restored)$/i, '');
    download.download = operation === 'protect' ? 'synthetic.protected.xlsx' : `${base}.restored${extension}`;
    setText(download, operation === 'protect' ? 'Download synthetic workbook' : 'Download restored workbook');
    setText(document.querySelector<HTMLElement>('#result-summary')!, operation === 'protect'
      ? `Generated a synthetic workbook across ${transformed.sheets.toLocaleString()} worksheets. Original headers in row 1 and dates remain visible; the encrypted original is embedded inside.`
      : `Restored ${transformed.cells.toLocaleString()} cells across ${transformed.sheets.toLocaleString()} worksheets.`, operation === 'protect'
      ? `أُنشئت تشفيرة ببيانات اصطناعية في ${transformed.sheets.toLocaleString('en')} أوراق عمل. تبقى العناوين الأصلية في الصف 1 والتواريخ ظاهرة، وتُضمّن النسخة الأصلية المشفّرة داخلها.`
      : `استُعيدت ${transformed.cells.toLocaleString('en')} خلايا في ${transformed.sheets.toLocaleString('en')} أوراق عمل.`);
    result.hidden = false;
    setText(status, showBackup ? 'Your workbook is ready. Save your recovery phrase below before leaving.' : 'Your workbook is ready to download.');
    download.focus({ preventScroll: true });
    download.scrollIntoView({ block: 'center' });
  } catch (cause) {
    clearSecrets();
    if (revision !== pageRevision) return;
    setText(status, '');
    let message = cause instanceof Error ? cause.message : '';
    if (passkeyPending && cause instanceof Error && ['NotAllowedError', 'AbortError'].includes(cause.name)) {
      message = 'Passkey request was cancelled or timed out. Try again and choose the original passkey.';
    } else if (passkeyPending && /PRF|authenticator/i.test(message) && !translations[message]) {
      message = 'This browser or passkey provider does not support PRF encryption. Try another provider or browser, or restore with your recovery phrase.';
    } else if (passkeyPending && !translations[message]) {
      message = 'Could not use this passkey. Try again with a supported browser and passkey provider.';
    } else if (message.startsWith('Could not read this Excel file:')) {
      const reason = message.slice('Could not read this Excel file: '.length);
      message = translations[reason] ? reason : 'Could not read this Excel file. Check that it is a valid, unencrypted workbook.';
      if (/password|encrypt/i.test(reason)) {
        message = 'This source workbook requires an Excel password. Remove its password in Excel and save a copy before protecting it here.';
      }
    } else if (message === 'This workbook requires a password.') {
      message = 'This workbook requires a password. Select Legacy password to restore it.';
    } else if (message === 'This workbook requires a passkey or recovery phrase.') {
      message = 'This workbook requires a passkey or recovery phrase. Change the restore method.';
    }
    if (!translations[message]) message = 'Could not process this workbook. Check the file and try again.';
    setText(error, message);
    error.hidden = false;
  } finally {
    activeRoot?.fill(0);
    activeRoot = undefined;
    password.value = recoveryPhrase.value = '';
    showPassword.checked = false;
    password.type = 'password';
    busy = false;
    controls.disabled = !cryptoAvailable;
    submit.disabled = changeFile.disabled = !cryptoAvailable;
    updateMethod();
    form.removeAttribute('aria-busy');
  }
});

updateMethod();
if (cryptoAvailable) {
  controls.disabled = false;
  setText(status, '');
} else {
  setText(status, 'Open ExcelKey over HTTPS or localhost to use encryption.');
}

window.addEventListener('pagehide', () => {
  pageRevision++;
  clearSecrets();
  clearResult();
});

if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {
  // Online processing remains available when offline caching is unavailable.
});
