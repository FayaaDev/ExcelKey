import { beforeAll, describe, expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { MAX_FILE_BYTES, transformWorkbook } from '../src/excel-v1';

const password = 'test password — not persisted';
const write = async (book: ExcelJS.Workbook) => new Uint8Array(await book.xlsx.writeBuffer());
async function read(data: Uint8Array) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(new Uint8Array(data).buffer);
  return book;
}
function fixture() {
  const book = new ExcelJS.Workbook();
  book.keywords = 'original keywords';
  const first = book.addWorksheet('Accounts');
  first.getRow(1).values = ['Readable header', 'Second header'];
  first.getCell('A2').value = 'secret text 8dfe';
  first.getCell('A2').note = 'secret note 83ef';
  first.getCell('B2').value = 123.45;
  first.getCell('C2').value = '001234';
  first.getCell('D2').value = new Date('2026-10-03T12:34:56.000Z');
  first.getCell('E2').value = false;
  first.getCell('F2').value = { formula: 'B2*2', result: 246.9 };
  first.getCell('G2').value = { richText: [{ text: 'rich secret 94ed', font: { bold: true } }] };
  first.getCell('H2').value = { text: 'private link', hyperlink: 'https://example.test/private-7adf', tooltip: 'Private' };
  first.getCell('I2').value = { error: '#DIV/0!' };
  first.getCell('J2').value = { formula: '"cached secret 30ba"', result: 'cached secret 30ba' };
  first.getCell('A3').value = '';
  first.getCell('B4').note = 'note without value';
  first.getCell('B4').font = { italic: true }; // Ensure ExcelJS writes a cell record for the note.
  first.getCell('C5').font = { bold: true }; // Styled blank, not populated.
  first.getCell('A2').font = { bold: true, color: { argb: 'FF112233' } };
  first.getCell('B2').numFmt = '#,##0.00';
  first.getColumn(1).width = 28;
  first.getRow(2).height = 25;
  first.views = [{ state: 'frozen', ySplit: 1 }];
  first.mergeCells('K2:L2');
  first.getCell('K2').value = 'merged secret';
  const second = book.addWorksheet('More data');
  second.getCell('A1').value = 'Other readable header';
  second.getCell('A2').value = true;
  second.getCell('B2').value = 0;
  return book;
}

describe('Excel workbook protection', () => {
  let original: Uint8Array, encrypted: Uint8Array;
  beforeAll(async () => {
    original = await write(fixture());
    encrypted = (await transformWorkbook(original, password, 'protect')).data;
  });

  test('two-sheet round trip preserves types, headers, notes, merged cells, blanks and layout; ZIP contains no data plaintext', async () => {
    const protectedBook = await read(encrypted);
    expect(protectedBook.getWorksheet('Accounts')!.getCell('A1').value).toBe('Readable header');
    expect(protectedBook.getWorksheet('More data')!.getCell('A1').value).toBe('Other readable header');
    expect(protectedBook.getWorksheet('Accounts')!.getCell('C5').value).toBeNull();
    expect(protectedBook.getWorksheet('Accounts')!.getCell('A2').note).toBeUndefined();
    const xml = Object.values(unzipSync(encrypted)).map(bytes => strFromU8(bytes)).join('\n');
    for (const secret of ['secret text 8dfe', 'secret note 83ef', '001234', 'rich secret 94ed', 'private-7adf', 'cached secret 30ba', 'merged secret', 'note without value']) {
      expect(xml).not.toContain(secret);
    }
    const result = await transformWorkbook(encrypted, password, 'restore');
    expect(result.sheets).toBe(2);
    expect(result.cells).toBe(15);
    const restored = await read(result.data), source = await read(original);
    expect(restored.worksheets.length).toBe(2);
    expect(restored.keywords).toBe('original keywords');
    for (const sheet of source.worksheets) {
      const target = restored.getWorksheet(sheet.name)!;
      sheet.eachRow({ includeEmpty: true }, row => row.eachCell({ includeEmpty: true }, cell => {
        expect(target.getCell(cell.address).value).toEqual(cell.value);
        expect(target.getCell(cell.address).note).toEqual(cell.note);
        expect(target.getCell(cell.address).style).toEqual(cell.style);
      }));
      expect(target.model.merges).toEqual(sheet.model.merges);
      expect(target.views).toEqual(sheet.views);
      expect(target.getColumn(1).width).toBe(sheet.getColumn(1).width);
      expect(target.getRow(2).height).toBe(sheet.getRow(2).height);
    }
  });

  test('wrong password and double protection fail', async () => {
    await expect(transformWorkbook(encrypted, 'wrong', 'restore')).rejects.toThrow('Incorrect password');
    await expect(transformWorkbook(encrypted, password, 'protect')).rejects.toThrow('already protected');
  });

  test('accepts an Excel calculation chain, drops it on rewrite, and restores formulas with recalculation', async () => {
    const book = new ExcelJS.Workbook(), sheet = book.addWorksheet('Data');
    sheet.getCell('A1').value = 'Header';
    sheet.getCell('A2').value = { formula: '1+1', result: 2 };
    sheet.getCell('B2').value = { text: 'Link', hyperlink: 'https://example.test/' };
    const entries = unzipSync(await write(book));
    entries['xl/calcChain.xml'] = strToU8('<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A2" i="1"/></calcChain>');
    entries['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(entries['xl/_rels/workbook.xml.rels']!).replace('</Relationships>',
      '<Relationship Target="calcChain.xml" Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"/></Relationships>'));
    entries['[Content_Types].xml'] = strToU8(strFromU8(entries['[Content_Types].xml']!).replace('</Types>',
      '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
    const protectedResult = await transformWorkbook(zipSync(entries), password, 'protect');
    const restoredResult = await transformWorkbook(protectedResult.data, password, 'restore');
    for (const result of [protectedResult, restoredResult]) {
      const output = unzipSync(result.data);
      expect(output['xl/calcChain.xml']).toBeUndefined();
      expect(strFromU8(output['xl/_rels/workbook.xml.rels']!)).not.toContain('calcChain');
      expect(strFromU8(output['[Content_Types].xml']!)).not.toContain('calcChain');
    }
    const output = unzipSync(restoredResult.data);
    expect(strFromU8(output['xl/worksheets/sheet1.xml']!)).toContain('<f>1+1</f>');
    expect(strFromU8(output['xl/workbook.xml']!)).toContain('fullCalcOnLoad="1"');
    const restored = await read(restoredResult.data);
    expect(restored.getWorksheet('Data')!.getCell('A2').value).toEqual({ formula: '1+1', result: 2 });
    expect(restored.getWorksheet('Data')!.getCell('B2').value).toEqual({ text: 'Link', hyperlink: 'https://example.test/' });
  });

  test.each([
    ['', '"', false],
    ['', "'", true],
    ['x:', '"', true],
    ['x:', "'", false],
  ] as const)('rejects lossy hyperlinks in raw XML (prefix=%s, quote=%s, reversed=%s)', async (prefix, quote, reversed) => {
    const book = new ExcelJS.Workbook();
    book.addWorksheet('Data').getCell('A1').value = 'Header';
    const entries = unzipSync(await write(book));
    const attributes = (pairs: [string, string][]) => (reversed ? pairs.reverse() : pairs)
      .map(([key, value]) => `${key} = ${quote}${value}${quote}`).join(' ');
    for (const internal of [true, false]) {
      const link = internal
        ? attributes([['ref', 'A2'], ['display', 'a > b'], ['location', 'Data!A1']])
        : attributes([['ref', 'A&#50;'], ['display', 'a > b'], ['r:id', 'rIdLink']]);
      entries['xl/worksheets/sheet1.xml'] = strToU8(
        `<${prefix}worksheet xmlns${prefix ? ':x' : ''}="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
        + `<${prefix}sheetData><${prefix}row r="2"><${prefix}c ${attributes([['r', 'A2'], ['s', '0']])}>`
        + (internal ? '' : `<${prefix}f>1+1</${prefix}f>`)
        + `<${prefix}v>2</${prefix}v></${prefix}c></${prefix}row></${prefix}sheetData>`
        + `<${prefix}hyperlinks><${prefix}hyperlink ${link}/></${prefix}hyperlinks></${prefix}worksheet>`);
      entries['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/" TargetMode="External"/></Relationships>');
      await expect(transformWorkbook(zipSync(entries), password, 'protect')).rejects.toThrow(
        internal ? 'Internal location hyperlinks' : 'Formula cells with hyperlinks');
    }
  });

  test.each([
    ['changed ciphertext', (b: ExcelJS.Workbook) => { const c = b.getWorksheet('Accounts')!.getCell('A2'); const text = String(c.value); c.value = text.slice(0, -8) + 'AAAAAAAA'; }],
    ['deleted ciphertext', (b: ExcelJS.Workbook) => { b.getWorksheet('Accounts')!.getCell('A2').value = null; }],
    ['replaced with plaintext', (b: ExcelJS.Workbook) => { b.getWorksheet('Accounts')!.getCell('A2').value = 'replacement'; }],
    ['moved ciphertext', (b: ExcelJS.Workbook) => { const s = b.getWorksheet('Accounts')!; s.getCell('B2').value = s.getCell('A2').value; }],
    ['deleted worksheet', (b: ExcelJS.Workbook) => { b.removeWorksheet('More data'); }],
    ['renamed worksheet', (b: ExcelJS.Workbook) => { b.getWorksheet('More data')!.name = 'Renamed'; }],
    ['extra data', (b: ExcelJS.Workbook) => { b.getWorksheet('Accounts')!.getCell('A6').value = 123; }],
    ['missing metadata', (b: ExcelJS.Workbook) => { b.removeWorksheet('__ExcelKey_v1'); }],
    ['modified metadata', (b: ExcelJS.Workbook) => { b.getWorksheet('__ExcelKey_v1')!.getCell('A3').value = 'AAAA'; }],
    ['renamed metadata', (b: ExcelJS.Workbook) => { b.getWorksheet('__ExcelKey_v1')!.name = 'Hidden elsewhere'; }],
    ['metadata marker removed', (b: ExcelJS.Workbook) => { b.getWorksheet('__ExcelKey_v1')!.getCell('A1').value = 'not metadata'; }],
    ['workbook sentinel removed', (b: ExcelJS.Workbook) => { b.keywords = ''; }],
    ['note added to ciphertext', (b: ExcelJS.Workbook) => { b.getWorksheet('Accounts')!.getCell('A2').note = 'replacement note'; }],
  ] as const)('%s fails without output', async (_name, mutate) => {
    const book = await read(encrypted);
    mutate(book);
    await expect(transformWorkbook(await write(book), password, 'restore')).rejects.toThrow();
  });

  test('renaming metadata cannot bypass double-protection detection', async () => {
    const book = await read(encrypted);
    book.getWorksheet('__ExcelKey_v1')!.name = 'Renamed';
    book.keywords = '';
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('already protected');
  });

  test('zero-data workbooks still authenticate the password and missing metadata', async () => {
    const book = new ExcelJS.Workbook();
    book.addWorksheet('Headers only').getCell('A1').value = 'Header';
    const result = await transformWorkbook(await write(book), password, 'protect');
    expect(result.cells).toBe(0);
    await expect(transformWorkbook(result.data, 'wrong', 'restore')).rejects.toThrow('Incorrect password');
    expect((await transformWorkbook(result.data, password, 'restore')).cells).toBe(0);
    const changed = await read(result.data);
    changed.removeWorksheet('__ExcelKey_v1');
    await expect(transformWorkbook(await write(changed), password, 'restore')).rejects.toThrow('metadata');
  });

  test('randomized workbook salts and unique nonces', async () => {
    const other = await transformWorkbook(original, password, 'protect');
    const left = await read(encrypted), right = await read(other.data);
    expect(left.getWorksheet('__ExcelKey_v1')!.getCell('A2').value).not.toBe(right.getWorksheet('__ExcelKey_v1')!.getCell('A2').value);
    const nonces: string[] = [];
    for (const sheet of left.worksheets.filter(s => s.name !== '__ExcelKey_v1')) sheet.eachRow(row => row.eachCell(cell => {
      if (cell.isMerged && cell.master !== cell) return;
      if (typeof cell.value === 'string' && cell.value.startsWith('ExcelKey:cell:v1:')) nonces.push(cell.value.slice('ExcelKey:cell:v1:'.length, 'ExcelKey:cell:v1:'.length + 16));
    }));
    nonces.push(String(left.getWorksheet('__ExcelKey_v1')!.getCell('A3').value).slice(0, 16));
    expect(new Set(nonces).size).toBe(nonces.length);
  });

  test('invalid ZIP, compressed/inflated limits, unsupported embedded parts and empty password', async () => {
    await expect(transformWorkbook(new Uint8Array([1, 2, 3]), password, 'protect')).rejects.toThrow('Invalid .xlsx');
    await expect(transformWorkbook(new Uint8Array(MAX_FILE_BYTES + 1), password, 'protect')).rejects.toThrow('20 MB');
    await expect(transformWorkbook(original, '', 'protect')).rejects.toThrow('password');
    const entries = unzipSync(original);
    entries['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<cache>secret</cache>');
    await expect(transformWorkbook(zipSync(entries), password, 'protect')).rejects.toThrow('Unsupported workbook parts');
    // Patch the central-directory inflated size; rejection must happen before allocation.
    const large = original.slice();
    const view = new DataView(large.buffer);
    for (let i = 0; i < large.length - 24; i++) {
      if (view.getUint32(i, true) === 0x02014b50) { view.setUint32(i + 24, 101 * 1024 * 1024, true); break; }
    }
    await expect(transformWorkbook(large, password, 'protect')).rejects.toThrow('100 MB');
  });

  test('rejects header/data merges, header formulas, shared/array formulas and oversized ciphertext', async () => {
    const book = new ExcelJS.Workbook(), sheet = book.addWorksheet('Data');
    sheet.mergeCells('A1:A2'); sheet.getCell('A1').value = 'crossing';
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('Merged cells');
    sheet.unMergeCells('A1:A2'); sheet.getCell('A1').value = { formula: 'A2', result: 42 };
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('header row 1');
    sheet.getCell('A1').value = 'Header';
    sheet.fillFormula('A2:A3', 'ROW()', [2, 3]);
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('shared');
    const encodedEntries = unzipSync(await write(book));
    encodedEntries['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(encodedEntries['xl/worksheets/sheet1.xml']!).replaceAll('t="shared"', 't="&#x73;hared"'));
    await expect(transformWorkbook(zipSync(encodedEntries), password, 'protect')).rejects.toThrow('shared');
    const arrayEntries = unzipSync(await write(book));
    arrayEntries['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(arrayEntries['xl/worksheets/sheet1.xml']!).replaceAll('t="shared"', 't="array"'));
    await expect(transformWorkbook(zipSync(arrayEntries), password, 'protect')).rejects.toThrow('Array');
    sheet.getCell('A2').value = 'x'.repeat(25_000); sheet.getCell('A3').value = null;
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('32,767');
  });

  test('rejects notes that ExcelJS cannot load rather than silently dropping them', async () => {
    const book = new ExcelJS.Workbook(), sheet = book.addWorksheet('Data');
    sheet.getCell('A1').value = 'Header';
    sheet.getCell('A2').note = 'orphan note';
    await expect(transformWorkbook(await write(book), password, 'protect')).rejects.toThrow('Notes without');
  });

  test('chunked metadata and header notes survive; date and boolean formula results restore', async () => {
    const book = new ExcelJS.Workbook(), sheet = book.addWorksheet('Data');
    sheet.getCell('A1').value = 'Header'; sheet.getCell('A1').note = 'readable header note';
    for (let row = 2; row < 3502; row++) sheet.getCell(row, 1).value = row;
    sheet.getCell('B2').value = { formula: 'TRUE()', result: true };
    sheet.getCell('C2').value = { formula: 'DATE(2026,10,3)', result: new Date('2026-10-03T00:00:00.000Z') };
    sheet.getCell('C2').numFmt = 'yyyy-mm-dd';
    const source = await write(book), result = await transformWorkbook(source, password, 'protect');
    const protectedBook = await read(result.data);
    const header = JSON.parse(String(protectedBook.getWorksheet('__ExcelKey_v1')!.getCell('A2').value));
    expect(header.chunks).toBeGreaterThan(1);
    expect(protectedBook.getWorksheet('Data')!.getCell('A1').note).toBe('readable header note');
    const restored = await read((await transformWorkbook(result.data, password, 'restore')).data);
    expect(restored.getWorksheet('Data')!.getCell('A3501').value).toBe(3501);
    const baseline = await read(source);
    for (const address of ['B2', 'C2']) expect(restored.getWorksheet('Data')!.getCell(address).value).toEqual(baseline.getWorksheet('Data')!.getCell(address).value);
  });

  test('authenticated manifest is still schema-validated (duplicate addresses)', async () => {
    const book = await read(encrypted), meta = book.getWorksheet('__ExcelKey_v1')!;
    const header = JSON.parse(String(meta.getCell('A2').value));
    const salt = Uint8Array.from(atob(header.salt), c => c.charCodeAt(0));
    const material = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveKey']);
    const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 600_000, hash: 'SHA-256' }, material,
      { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
    const packed = Uint8Array.from(atob(String(meta.getCell('A3').value)), c => c.charCodeAt(0));
    const aad = new TextEncoder().encode(JSON.stringify(['ExcelKey:workbook:v1', header.salt, 'manifest']));
    const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: packed.slice(0, 12), additionalData: aad }, key, packed.slice(12));
    const manifest = JSON.parse(new TextDecoder().decode(plain));
    manifest.sheets[0].cells.push(manifest.sheets[0].cells[0]);
    const nonce = crypto.getRandomValues(new Uint8Array(12));
    const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce, additionalData: aad }, key, new TextEncoder().encode(JSON.stringify(manifest))));
    const changed = new Uint8Array(12 + cipher.length); changed.set(nonce); changed.set(cipher, 12);
    meta.getCell('A3').value = btoa(String.fromCharCode(...changed));
    await expect(transformWorkbook(await write(book), password, 'restore')).rejects.toThrow('damaged or modified');
  });
});
