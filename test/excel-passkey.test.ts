import { expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { read, write, utils } from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { transformWorkbook, type WorkbookKey } from '../src/excel';
import { transformWorkbook as transformV1 } from '../src/excel-v1';
import { hasV2Protection, hasWorkbookProtection } from '../src/excel-ooxml';
import { masterPrkFromPrfSecret } from '../src/identity';
import { encodeRecoveryBip39, decodeRecoveryBip39 } from '../src/recovery';

const key = { masterPrk: masterPrkFromPrfSecret(new Uint8Array(32).fill(7)), namespace: 'filekey.app' };
const book = utils.book_new();
utils.book_append_sheet(book, utils.aoa_to_sheet([['Name', 'Date'], ['private-name', '2026-10-04']]), 'Private');
const original = new Uint8Array(write(book, { type: 'array', bookType: 'xlsx' }));
const protect = () => transformWorkbook(original, key, 'protect', { filename: 'original.xlsx' });
type Entries = Record<string, Uint8Array>;
function pathOf(parts: Entries): string { return Object.keys(parts).find(path => path.startsWith('excelkey/') && path.endsWith('.json'))!; }
function envelope(parts: Entries, mutate: (value: Record<string, unknown>) => void): void {
  const path = pathOf(parts), value = JSON.parse(strFromU8(parts[path]!));
  mutate(value); parts[path] = strToU8(JSON.stringify(value));
}

test('PRF-root and recovery phrase restore exact originals through the synthetic pipeline', async () => {
  const rootBefore = key.masterPrk.slice();
  for (const bookType of ['xlsx', 'xls', 'xlsb'] as const) {
    const input = new Uint8Array(write(book, { type: 'array', bookType }));
    const filename = `original.${bookType}`;
    const result = await transformWorkbook(input, key, 'protect', { filename });
    expect(result.filename).toBe('synthetic.xlsx');
    expect(result.cells).toBe(1); expect(result.sheets).toBe(1);
    expect(hasWorkbookProtection(result.data)).toBe(true);
    expect(hasV2Protection(result.data)).toBe(false);
    const visible = read(result.data, { type: 'array' }).Sheets['Worksheet 1']!;
    expect(visible.A1.v).toBe('Name'); expect(visible.A2.v).not.toBe('private-name');
    expect(visible.B2.v).toBe('2026-10-04');
    const restored = await transformWorkbook(result.data, key, 'restore');
    expect(restored.data).toEqual(input); expect(restored.filename).toBe(filename);
    const recovery = { masterPrk: decodeRecoveryBip39(encodeRecoveryBip39(key.masterPrk)), namespace: key.namespace };
    expect((await transformWorkbook(result.data, recovery, 'restore')).data).toEqual(input);
  }
  expect(key.masterPrk).toEqual(rootBefore);
});

test('v3 randomizes salt and nonce; wrong root, namespace and credential kinds fail closed', async () => {
  const first = await protect(), second = await protect();
  const unpack = (data: Uint8Array) => { const parts = unzipSync(data); return JSON.parse(strFromU8(parts[pathOf(parts)]!)); };
  const a = unpack(first.data), b = unpack(second.data);
  expect(a.version).toBe(3); expect(a.namespace).toBe(key.namespace);
  expect(a.salt).not.toBe(b.salt); expect(a.nonce).not.toBe(b.nonce);
  for (const wrong of [
    { ...key, masterPrk: masterPrkFromPrfSecret(new Uint8Array(32).fill(8)) },
    { ...key, masterPrk: decodeRecoveryBip39(encodeRecoveryBip39(new Uint8Array(32).fill(9))) },
    { ...key, namespace: 'other.example' },
  ]) await expect(transformWorkbook(first.data, wrong, 'restore')).rejects.toThrow();
  await expect(transformWorkbook(first.data, 'password', 'restore')).rejects.toThrow('requires a passkey');
  for (const credential of [key, 'password']) {
    await expect(transformWorkbook(first.data, credential, 'protect')).rejects.toThrow('already protected');
  }
  for (const invalid of [
    { ...key, masterPrk: new Uint8Array(31) },
    { ...key, namespace: 'https://filekey.app' },
    { ...key, namespace: '' },
    { ...key, namespace: 123 },
  ]) await expect(transformWorkbook(first.data, invalid as WorkbookKey, 'restore')).rejects.toThrow();
});

test('v3 rejects envelope, relationship and package tampering', async () => {
  const protectedFile = await protect();
  const mutations: ((parts: Entries) => void)[] = [
    parts => envelope(parts, e => { e.version = 2; }),
    parts => envelope(parts, e => { e.version = 4; }),
    parts => envelope(parts, e => { e.namespace = 'other.example'; }),
    parts => envelope(parts, e => { e.namespace = null; }),
    parts => envelope(parts, e => { delete e.namespace; }),
    parts => envelope(parts, e => { e.extra = true; }),
    ...['salt', 'nonce', 'ciphertext'].flatMap(field => [
      (parts: Entries) => envelope(parts, e => { e[field] = '!invalid!'; }),
      (parts: Entries) => envelope(parts, e => { e[field] = ''; }),
      (parts: Entries) => envelope(parts, e => {
        const bytes = Uint8Array.from(atob(e[field] as string), c => c.charCodeAt(0));
        bytes[0] = bytes[0]! ^ 1; e[field] = btoa(String.fromCharCode(...bytes));
      }),
    ]),
    parts => { parts[pathOf(parts)] = strToU8('null'); },
    parts => { parts[pathOf(parts)] = strToU8('{'); },
    parts => { delete parts[pathOf(parts)]; },
    parts => { parts['new.bin'] = new Uint8Array([1]); },
    parts => { delete parts['xl/styles.xml']; },
    parts => { parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('Name', 'Edited')); },
    parts => { parts['_rels/.rels'] = strToU8(strFromU8(parts['_rels/.rels']!).replace('ExcelKeyV3', 'Changed')); },
    parts => {
      const path = pathOf(parts), next = 'excelkey/moved.json';
      parts[next] = parts[path]!; delete parts[path];
      for (const name of ['_rels/.rels', '[Content_Types].xml']) parts[name] = strToU8(strFromU8(parts[name]!).replace(path, next));
    },
  ];
  for (const mutate of mutations) {
    const parts = unzipSync(protectedFile.data); mutate(parts);
    await expect(transformWorkbook(zipSync(parts), key, 'restore')).rejects.toThrow();
  }
  // Even supplying the tampered namespace cannot make the original ciphertext authenticate.
  const changed = unzipSync(protectedFile.data);
  envelope(changed, e => { e.namespace = 'other.example'; });
  await expect(transformWorkbook(zipSync(changed), { ...key, namespace: 'other.example' }, 'restore')).rejects.toThrow();
});

test('unsupported, ambiguous and dangling protection is never converted into synthetic output', async () => {
  const protectedFile = await protect();
  for (const change of [
    (xml: string) => xml.replace('protected-workbook:v3', 'protected-workbook:v99'),
    (xml: string) => xml.replace('</Relationships>', '<Relationship Id="duplicate" Type="urn:excelkey:protected-workbook:v2" Target="/other.json"/></Relationships>'),
    (xml: string) => xml.replace('Type="urn:excelkey:protected-workbook:v3"', 'TargetMode="External" Type="urn:excelkey:protected-workbook:v3"'),
  ]) {
    const parts = unzipSync(protectedFile.data);
    parts['_rels/.rels'] = strToU8(change(strFromU8(parts['_rels/.rels']!)));
    for (const mode of ['protect', 'restore'] as const) await expect(transformWorkbook(zipSync(parts), key, mode)).rejects.toThrow();
  }
  const dangling = unzipSync(protectedFile.data); delete dangling[pathOf(dangling)];
  expect(hasWorkbookProtection(zipSync(dangling))).toBe(true);
  await expect(transformWorkbook(zipSync(dangling), key, 'protect')).rejects.toThrow('already protected');
});

test('password writer stays v2 and password reader restores exact bytes', async () => {
  const result = await transformWorkbook(original, 'legacy-password', 'protect', { filename: 'legacy.xlsx' });
  expect(hasV2Protection(result.data)).toBe(true);
  const parts = unzipSync(result.data);
  expect(Object.keys(JSON.parse(strFromU8(parts[pathOf(parts)]!))).sort()).toEqual(['ciphertext', 'nonce', 'salt', 'version']);
  const restored = await transformWorkbook(result.data, 'legacy-password', 'restore');
  expect(restored.data).toEqual(original); expect(restored.filename).toBe('legacy.xlsx');
  await expect(transformWorkbook(result.data, key, 'restore')).rejects.toThrow('requires a password');
  await expect(transformWorkbook(result.data, key, 'protect')).rejects.toThrow('already protected');
});

test('legacy v1 accepts its password and rejects root credentials', async () => {
  const legacy = new ExcelJS.Workbook();
  legacy.addWorksheet('Private').addRows([['Name'], ['private-name']]);
  const result = await transformV1(new Uint8Array(await legacy.xlsx.writeBuffer()), 'legacy-password', 'protect');
  await expect(transformWorkbook(result.data, key, 'restore')).rejects.toThrow('requires a password');
  await expect(transformWorkbook(result.data, key, 'protect')).rejects.toThrow('already protected');
  const restored = await transformWorkbook(result.data, 'legacy-password', 'restore');
  expect(read(restored.data, { type: 'array' }).Sheets.Private!.A2.v).toBe('private-name');
});
