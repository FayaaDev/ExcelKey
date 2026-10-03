import { EXCEL_EXTENSIONS, MAX_FILE_BYTES, transformWorkbook } from '../src/excel';
import { translations } from './excel-i18n';

const form = document.querySelector<HTMLFormElement>('#workbook-form')!;
const controls = document.querySelector<HTMLFieldSetElement>('#controls')!;
const fileInput = document.querySelector<HTMLInputElement>('#workbook')!;
const password = document.querySelector<HTMLInputElement>('#password')!;
const confirmation = document.querySelector<HTMLInputElement>('#confirmation')!;
const showPassword = document.querySelector<HTMLInputElement>('#show-password')!;
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
  return new FormData(form).get('mode') === 'restore' ? 'restore' : 'protect';
}

form.addEventListener('change', event => {
  if (busy) return;
  if ((event.target as HTMLInputElement).name === 'mode') {
    const protect = mode() === 'protect';
    workflow.hidden = dock.hidden = false;
    document.querySelector<HTMLElement>('#confirmation-field')!.hidden = !protect;
    confirmation.required = protect;
    confirmation.disabled = !protect;
    password.autocomplete = protect ? 'new-password' : 'current-password';
    setText(document.querySelector<HTMLElement>('#password-label')!, protect ? 'Create a password' : 'Workbook password');
    setText(submit, protect ? 'Generate synthetic workbook' : 'Restore workbook');
    setText(document.querySelector<HTMLElement>('#dock-title')!, protect ? 'Generate synthetic workbook' : 'Restore original workbook');
    submit.disabled = controls.disabled;
    password.value = confirmation.value = '';
    clearResult();
  } else if (event.target === fileInput) {
    clearResult();
    updateFile();
  }
});

showPassword.addEventListener('change', () => {
  password.type = confirmation.type = showPassword.checked ? 'text' : 'password';
});

form.addEventListener('submit', async event => {
  event.preventDefault();
  if (busy) return;
  clearResult();
  const operation = mode();
  const file = fileInput.files?.[0];
  try {
    if (!file || !EXCEL_EXTENSIONS.test(file.name)) throw new Error('Choose an Excel workbook.');
    if (file.size > MAX_FILE_BYTES) throw new Error('Choose a workbook smaller than 64 MiB.');
    if (!password.value) throw new Error('Enter a password.');
    if (operation === 'protect' && password.value !== confirmation.value) {
      confirmation.focus();
      throw new Error('Passwords do not match. Enter the same password in both fields.');
    }
    busy = true;
    controls.disabled = true;
    submit.disabled = changeFile.disabled = true;
    form.setAttribute('aria-busy', 'true');
    setText(status, operation === 'protect' ? 'Generating synthetic workbook… Keep this tab open.' : 'Restoring workbook… Keep this tab open.');
    const transformed = await transformWorkbook(new Uint8Array(await file.arrayBuffer()), password.value, operation, { filename: file.name });
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
    setText(status, 'Your workbook is ready to download.');
    download.focus({ preventScroll: true });
    download.scrollIntoView({ block: 'center' });
  } catch (cause) {
    setText(status, '');
    const message = cause instanceof Error ? cause.message : 'Could not process this workbook. Check the file and try again.';
    setText(error, message, translations[message] ?? translations['Could not process this workbook. Check the file and try again.']);
    error.hidden = false;
  } finally {
    if (busy) {
      password.value = confirmation.value = '';
      showPassword.checked = false;
      password.type = confirmation.type = 'password';
    }
    busy = false;
    controls.disabled = false;
    submit.disabled = changeFile.disabled = false;
    form.removeAttribute('aria-busy');
  }
});

if (window.isSecureContext && crypto.subtle) {
  controls.disabled = false;
  setText(status, '');
} else {
  setText(status, 'Open ExcelKey over HTTPS or localhost to use password encryption.');
}

window.addEventListener('pagehide', () => {
  password.value = confirmation.value = '';
});

if ('serviceWorker' in navigator) navigator.serviceWorker.register('/sw.js').catch(() => {
  // Online processing remains available when offline caching is unavailable.
});
