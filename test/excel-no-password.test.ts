import { expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { read, write, utils } from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { transformWorkbook, type WorkbookKey } from '../src/excel';
import { transformWorkbook as transformV1 } from '../src/excel-v1';
import { hasV2Protection, hasWorkbookProtection, transformOoxml } from '../src/excel-ooxml';

const passkey = { masterPrk: new Uint8Array(32).fill(7), namespace: 'filekey.app' };
const book = utils.book_new();
utils.book_append_sheet(book, utils.aoa_to_sheet([['Name', 'Date'], ['private-name', '2026-10-04']]), 'Private');
const original = new Uint8Array(write(book, { type: 'array', bookType: 'xlsx' }));
const protect = () => transformWorkbook(original, null, 'protect', { filename: 'original.xlsx' });
type Entries = Record<string, Uint8Array>;
function pathOf(parts: Entries): string { return Object.keys(parts).find(path => path.startsWith('excelkey/') && path.endsWith('.json'))!; }
function envelope(parts: Entries, mutate: (value: Record<string, unknown>) => void): void {
  const path = pathOf(parts), value = JSON.parse(strFromU8(parts[path]!));
  mutate(value); parts[path] = strToU8(JSON.stringify(value));
}
function decode(value: string): Uint8Array<ArrayBuffer> { return Uint8Array.from(atob(value), c => c.charCodeAt(0)); }

test('explicit no-password mode restores exact XLSX and XLS bytes and filenames', async () => {
  for (const bookType of ['xlsx', 'xls'] as const) {
    const input = new Uint8Array(write(book, { type: 'array', bookType }));
    const filename = `الأصل.${bookType}`;
    const result = await transformWorkbook(input, null, 'protect', { filename });
    expect(result.filename).toBe('synthetic.xlsx');
    expect(result.cells).toBe(1); expect(result.sheets).toBe(1);
    expect(hasWorkbookProtection(result.data)).toBe(true);
    expect(hasV2Protection(result.data)).toBe(false);
    const visible = read(result.data, { type: 'array' }).Sheets['Worksheet 1']!;
    expect(visible.A1.v).toBe('Name'); expect(visible.A2.v).not.toBe('private-name');
    expect(visible.B2.v).toBe('2026-10-04');
    const restored = await transformWorkbook(result.data, null, 'restore');
    expect(restored.data).toEqual(input); expect(restored.filename).toBe(filename);
    expect(restored.cells).toBe(result.cells); expect(restored.sheets).toBe(result.sheets);
  }
});

test('v4 embeds a random AES256 key and nonce with a version 4 manifest and path-bound AAD', async () => {
  const first = await protect(), second = await protect();
  const parts = unzipSync(first.data), path = pathOf(parts);
  const a = JSON.parse(strFromU8(parts[path]!));
  const other = unzipSync(second.data), b = JSON.parse(strFromU8(other[pathOf(other)]!));
  expect(Object.keys(a).sort()).toEqual(['ciphertext', 'key', 'nonce', 'version']);
  expect(a.version).toBe(4); expect(decode(a.key).length).toBe(32); expect(decode(a.nonce).length).toBe(12);
  expect(a.key).not.toBe(b.key); expect(a.nonce).not.toBe(b.nonce);
  expect(strFromU8(parts['_rels/.rels']!)).toContain('urn:excelkey:protected-workbook:v4');
  const key = await crypto.subtle.importKey('raw', decode(a.key), 'AES-GCM', false, ['decrypt']);
  const plain = new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: decode(a.nonce), additionalData: strToU8(JSON.stringify(['urn:excelkey:protected-workbook:v4', path])) }, key, decode(a.ciphertext)));
  const length = new DataView(plain.buffer).getUint32(0);
  expect(JSON.parse(strFromU8(plain.subarray(4, 4 + length))).version).toBe(4);
  expect(plain.slice(4 + length)).toEqual(original);
});

test('every credential mode rejects mismatched formats with the required mode message', async () => {
  const credentials: WorkbookKey[] = ['password', passkey, null];
  const messages = ['This workbook requires a password.', 'This workbook requires a passkey or recovery phrase.', 'This workbook requires no password. Select No password to restore it.'];
  for (const [index, credential] of credentials.entries()) {
    const result = await transformWorkbook(original, credential, 'protect');
    for (const attempted of credentials) {
      if (attempted !== credential) await expect(transformWorkbook(result.data, attempted, 'restore')).rejects.toThrow(messages[index]!);
      await expect(transformWorkbook(result.data, attempted, 'protect')).rejects.toThrow('already protected');
    }
  }
  const legacy = new ExcelJS.Workbook();
  legacy.addWorksheet('Private').addRows([['Name'], ['private-name']]);
  const v1 = await transformV1(new Uint8Array(await legacy.xlsx.writeBuffer()), 'password', 'protect');
  await expect(transformWorkbook(v1.data, null, 'restore')).rejects.toThrow('This workbook requires a password.');
  await expect(transformWorkbook(v1.data, null, 'protect')).rejects.toThrow('already protected');
});

test('v4 rejects damaged envelopes, invalid key lengths and changed visible packages', async () => {
  const result = await protect();
  const mutations: ((parts: Entries) => void)[] = [
    ...['key', 'nonce', 'ciphertext'].flatMap(field => [
      (parts: Entries) => envelope(parts, e => { delete e[field]; }),
      (parts: Entries) => envelope(parts, e => { e[field] = null; }),
      (parts: Entries) => envelope(parts, e => { e[field] = '!invalid!'; }),
      (parts: Entries) => envelope(parts, e => { e[field] = ''; }),
      (parts: Entries) => envelope(parts, e => {
        const bytes = decode(e[field] as string); bytes[0] = bytes[0]! ^ 1;
        e[field] = btoa(String.fromCharCode(...bytes));
      }),
    ]),
    ...[16, 24, 31, 33].map(length => (parts: Entries) => envelope(parts, e => { e.key = btoa('x'.repeat(length)); })),
    parts => envelope(parts, e => { e.version = 3; }),
    parts => envelope(parts, e => { e.extra = true; }),
    parts => { parts[pathOf(parts)] = strToU8('null'); },
    parts => { parts[pathOf(parts)] = strToU8('{'); },
    parts => { delete parts[pathOf(parts)]; },
    parts => { parts['new.bin'] = new Uint8Array([1]); },
    parts => { delete parts['xl/styles.xml']; },
    parts => { parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('Name', 'Edited')); },
    parts => { parts['_rels/.rels'] = strToU8(strFromU8(parts['_rels/.rels']!).replace('ExcelKeyV4', 'Changed')); },
    parts => {
      const path = pathOf(parts), next = 'excelkey/moved.json';
      parts[next] = parts[path]!; delete parts[path];
      for (const name of ['_rels/.rels', '[Content_Types].xml']) parts[name] = strToU8(strFromU8(parts[name]!).replace(path, next));
    },
  ];
  for (const mutate of mutations) {
    const parts = unzipSync(result.data); mutate(parts);
    await expect(transformWorkbook(zipSync(parts), null, 'restore')).rejects.toThrow('damaged or modified');
  }
});

test('relabeling secret-based v2/v3 envelopes as v4 cannot bypass authentication', async () => {
  for (const credential of ['password', passkey]) {
    const result = await transformWorkbook(original, credential, 'protect');
    const parts = unzipSync(result.data);
    parts['_rels/.rels'] = strToU8(strFromU8(parts['_rels/.rels']!).replace(/protected-workbook:v[23]/, 'protected-workbook:v4'));
    envelope(parts, e => { e.version = 4; e.key = btoa('x'.repeat(32)); delete e.salt; delete e.namespace; });
    await expect(transformWorkbook(zipSync(parts), null, 'restore')).rejects.toThrow('damaged or modified');
  }
});

test('empty and undefined passwords remain invalid in both core entry points', async () => {
  const result = await protect();
  for (const transform of [transformWorkbook, transformOoxml]) {
    for (const invalid of ['', undefined]) {
      for (const mode of ['protect', 'restore'] as const) {
        await expect(transform(mode === 'protect' ? original : result.data, invalid as WorkbookKey, mode)).rejects.toThrow('Enter a password.');
      }
    }
  }
});
