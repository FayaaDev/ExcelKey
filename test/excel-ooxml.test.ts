import { describe, expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { Zip, ZipDeflate, unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { hasV2Protection, MAX_FILE_BYTES, transformOoxml } from '../src/excel-ooxml';

const PASSWORD = 'correct horse battery staple';
const TYPE = 'urn:excelkey:protected-workbook:v2';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
type Entries = Record<string, Uint8Array>;
function text(entries: Entries, name: string): string { return strFromU8(entries[name]!); }
function change(entries: Entries, name: string, fn: (xml: string) => string): void { entries[name] = strToU8(fn(text(entries, name))); }
function row1(xml: string): string { return /<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/.exec(xml)![0]; }
function metaPath(entries: Entries): string { return Object.keys(entries).find(name => name.startsWith('excelkey/') && name.endsWith('.json'))!; }
function streamedZip(entries: Entries): Uint8Array<ArrayBuffer> {
  const chunks: Uint8Array[] = [];
  const zip = new Zip((error, chunk) => { if (error) throw error; chunks.push(chunk); });
  for (const [name, bytes] of Object.entries(entries)) {
    const part = new ZipDeflate(name); zip.add(part); part.push(bytes, true);
  }
  zip.end();
  const data = new Uint8Array(chunks.reduce((size, chunk) => size + chunk.length, 0));
  let offset = 0;
  for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
  return data;
}

async function fixture(): Promise<Uint8Array> {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('__ExcelKey_v1');
  sheet.addRow(['ExcelKey:workbook:v1', 'Header 😀', { formula: '1+1', result: 2 }]);
  sheet.addRow(['secret-data', 'x'.repeat(32767), { formula: 'SUM(1,2)', result: 3 }]);
  sheet.getCell('A1').note = 'header-note';
  sheet.getCell('A2').note = 'secret-note';
  sheet.getCell('D1').value = { text: 'header-link', hyperlink: 'https://example.com/header' };
  sheet.getCell('D2').value = { text: 'private-link', hyperlink: 'https://example.com/secret-link' };
  sheet.getCell('B2').dataValidation = { type: 'whole', operator: 'greaterThan', formulae: [0] };
  sheet.addConditionalFormatting({ ref: 'A2:A3', rules: [{ type: 'expression', formulae: ['A2>0'], priority: 1, style: { font: { bold: true } } }] });
  sheet.addTable({ name: 'DataTable', ref: 'A5', headerRow: true, columns: [{ name: 'First' }, { name: 'Second' }], rows: [['table-secret', 42]] });
  const hidden = book.addWorksheet('Hidden', { state: 'veryHidden' });
  hidden.addRow(['Hidden header']); hidden.addRow(['hidden-secret']);
  book.definedNames.add("'__ExcelKey_v1'!$A$2", 'ExcelKeyLegitimateName');
  const entries = unzipSync(new Uint8Array(await book.xlsx.writeBuffer()));
  // Add unknown feature records and arbitrary binary feature parts. Nothing is whitelisted.
  change(entries, 'xl/worksheets/sheet1.xml', xml => xml.replace('</worksheet>', '<extLst><ext uri="custom"><opaque xmlns="urn:custom">feature</opaque></ext></extLst></worksheet>'));
  change(entries, 'xl/worksheets/sheet1.xml', xml => xml.replace('<f>SUM(1,2)</f>', '<f t="array" ref="C2:C3">SUM(1,2)</f>'));
  change(entries, 'xl/workbook.xml', xml => xml.replace('</workbook>', '<extLst><ext uri="workbook-feature"/></extLst></workbook>'));
  entries['xl/vbaProject.bin'] = new Uint8Array([0, 255, 1, 2, 3]);
  entries['xl/media/image1.png'] = new Uint8Array([137, 80, 78, 71]);
  entries['xl/charts/chart1.xml'] = strToU8('<chart xmlns="urn:feature"><data>cached feature data</data></chart>');
  entries['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<records xmlns="urn:feature"/>');
  entries['customXml/item1.xml'] = strToU8('<arbitrary xmlns="urn:arbitrary"/>');
  entries['excelkey/legitimate.bin'] = strToU8('legitimate reserved-looking path');
  change(entries, '[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));
  return zipSync(entries);
}

describe('raw OOXML v2', () => {
  test('preserves features, headers, opaque parts and restores the exact original ZIP', async () => {
    const original = await fixture();
    expect(hasV2Protection(original)).toBe(false);
    const protectedFile = await transformOoxml(original, PASSWORD, 'protect', { filename: 'features.xlsm' });
    expect(protectedFile.cells).toBe(9);
    expect(protectedFile.sheets).toBe(2);
    expect(protectedFile.filename).toBe('features.xlsm');
    expect(hasV2Protection(protectedFile.data)).toBe(true);
    const before = unzipSync(original), after = unzipSync(protectedFile.data);
    for (const name of ['xl/workbook.xml', 'xl/styles.xml', 'xl/tables/table1.xml', 'xl/vbaProject.bin', 'xl/media/image1.png', 'xl/charts/chart1.xml', 'xl/pivotCache/pivotCacheRecords1.xml', 'customXml/item1.xml', 'excelkey/legitimate.bin']) expect(after[name]).toEqual(before[name]);
    expect(row1(text(after, 'xl/worksheets/sheet1.xml'))).toBe(row1(text(before, 'xl/worksheets/sheet1.xml')));
    const sheet = text(after, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('dataValidations'); expect(sheet).toContain('conditionalFormatting'); expect(sheet).toContain('urn:custom');
    expect(sheet).not.toContain('SUM(1,2)');
    expect(sheet).toContain('<f>1+1</f>');
    expect(sheet).toContain('inlineStr');
    const visible = Object.entries(after).filter(([name]) => name !== metaPath(after)).map(([, bytes]) => strFromU8(bytes)).join('\n');
    for (const secret of ['secret-data', 'hidden-secret', 'secret-note', 'private-link', 'https://example.com/secret-link', 'x'.repeat(100)]) expect(visible).not.toContain(secret);
    expect(visible).toContain('header-note'); expect(visible).toContain('https://example.com/header');
    expect(text(after, 'xl/sharedStrings.xml')).toContain('Header 😀');
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(protectedFile.data.buffer as ArrayBuffer);
    expect(loaded.getWorksheet('__ExcelKey_v1')!.getCell('A1').value).toBe('ExcelKey:workbook:v1');
    expect(String(loaded.getWorksheet('Hidden')!.getCell('A2').value).length).toBeLessThan(100);
    const restored = await transformOoxml(protectedFile.data, PASSWORD, 'restore');
    expect(restored.data).toEqual(original); expect(restored.filename).toBe('features.xlsm');
    expect(restored.cells).toBe(9); expect(restored.sheets).toBe(2);
  });

  test('authenticates every package part, structural changes, metadata and original filename', async () => {
    const original = await fixture();
    const protectedFile = await transformOoxml(original, PASSWORD, 'protect');
    await expect(transformOoxml(protectedFile.data, 'wrong', 'restore')).rejects.toThrow('Incorrect password');
    for (const mutate of [
      (entries: Entries) => { entries['xl/vbaProject.bin'] = new Uint8Array([8]); },
      (entries: Entries) => { entries['new.bin'] = new Uint8Array([9]); },
      (entries: Entries) => { delete entries['xl/styles.xml']; },
      (entries: Entries) => change(entries, 'xl/sharedStrings.xml', xml => xml.replace('Header', 'Edited')),
      (entries: Entries) => change(entries, '_rels/.rels', xml => xml.replace('Id="ExcelKeyV2"', 'Id="changed"')),
      (entries: Entries) => { const path = metaPath(entries); const envelope = JSON.parse(text(entries, path)); envelope.nonce = btoa('000000000000'); entries[path] = strToU8(JSON.stringify(envelope)); },
      (entries: Entries) => { delete entries[metaPath(entries)]; },
      (entries: Entries) => change(entries, '_rels/.rels', xml => xml.replace(new RegExp(`<Relationship[^>]*Type="${TYPE}"[^>]*/>`), '')),
    ]) {
      const entries = unzipSync(protectedFile.data); mutate(entries);
      await expect(transformOoxml(zipSync(entries), PASSWORD, 'restore')).rejects.toThrow();
    }
    const dangling = unzipSync(protectedFile.data); delete dangling[metaPath(dangling)];
    expect(hasV2Protection(zipSync(dangling))).toBe(true);
    await expect(transformOoxml(protectedFile.data, PASSWORD, 'protect')).rejects.toThrow('already protected');
    await expect(transformOoxml(original, PASSWORD, 'restore')).rejects.toThrow('Incorrect password');
  });

  test('separate binary original and authenticated original filename restore exactly', async () => {
    const original = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 4, 5, 6]);
    const protectedFile = await transformOoxml(await fixture(), PASSWORD, 'protect', { original, filename: 'legacy.xls' });
    expect(protectedFile.filename).toBe('legacy.xlsx');
    const restored = await transformOoxml(protectedFile.data, PASSWORD, 'restore');
    expect(restored.data).toEqual(original); expect(restored.filename).toBe('legacy.xls');
  });

  test('handles streaming ZIP data descriptors, UTF-16 XML, empty string tables and arbitrary part names', async () => {
    const entries = unzipSync(await fixture());
    const utf16 = (value: string): Uint8Array<ArrayBuffer> => {
      const bytes = new Uint8Array(2 + value.length * 2), view = new DataView(bytes.buffer);
      view.setUint16(0, 0xfeff, true);
      for (let i = 0; i < value.length; i++) view.setUint16(2 + i * 2, value.charCodeAt(i), true);
      return bytes;
    };
    const sourceSheet = text(entries, 'xl/worksheets/sheet1.xml').replace('encoding="UTF-8"', 'encoding="UTF-16"');
    entries['xl/worksheets/sheet1.xml'] = utf16(sourceSheet);
    Object.defineProperty(entries, '__proto__', { value: new Uint8Array([1, 2, 3]), enumerable: true });
    change(entries, '[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/__proto__" ContentType="application/octet-stream"/></Types>'));
    const original = streamedZip(entries);
    const protectedFile = await transformOoxml(original, PASSWORD, 'protect');
    const after = unzipSync(protectedFile.data);
    expect(new TextDecoder('utf-16le').decode(after['xl/worksheets/sheet1.xml'])).toContain(row1(sourceSheet));
    // unzipSync uses a normal object; a literal __proto__ entry becomes its prototype.
    expect(Object.getPrototypeOf(after)).toEqual(new Uint8Array([1, 2, 3]));
    expect((await transformOoxml(protectedFile.data, PASSWORD, 'restore')).data).toEqual(original);
    Object.defineProperty(after, '__proto__', { value: new Uint8Array([4, 5, 6]), enumerable: true });
    await expect(transformOoxml(streamedZip(after), PASSWORD, 'restore')).rejects.toThrow('Incorrect password');

    const empty = unzipSync(await fixture());
    empty['xl/sharedStrings.xml'] = strToU8('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>');
    for (const path of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']) change(empty, path, xml => xml.replace(/<c([^>]*?) t="s"([^>]*)><v>\d+<\/v><\/c>/g, '<c$1 t="inlineStr"$2><is><t>text</t></is></c>'));
    const emptyOriginal = zipSync(empty);
    const emptyProtected = await transformOoxml(emptyOriginal, PASSWORD, 'protect');
    expect(text(unzipSync(emptyProtected.data), 'xl/sharedStrings.xml')).toContain('</sst>');
    expect((await transformOoxml(emptyProtected.data, PASSWORD, 'restore')).data).toEqual(emptyOriginal);
  });

  test('resolves Strict prefixed unusual paths; handles CRLF, astral text, inferred rows, shared and header-spanning formulas/merges', async () => {
    const ss = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    const office = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
    const header = '<s:row r="1"><s:c r="A1" t="s"><s:v>&#48;</s:v></s:c><s:c r="B1"><s:f t="shared" si="0" ref="B1:B3">1+1</s:f><s:v>2</s:v></s:c></s:row>';
    const entries: Entries = {
      '[Content_Types].xml': strToU8(`<ct:Types xmlns:ct="${CT}"/>`),
      '_rels/.rels': strToU8(`<p:Relationships xmlns:p="${PR}"><p:Relationship Id="w" Type="${office}/officeDocument" Target="/odd/book.xml"/></p:Relationships>`),
      'odd/book.xml': strToU8(`<s:workbook xmlns:s="${ss}" xmlns:q="${office}"><s:sheets><s:sheet name="Odd" sheetId="1" q:id="sheet"/></s:sheets></s:workbook>`),
      'odd/_rels/book.xml.rels': strToU8(`<Relationships xmlns="${PR}"><Relationship Id="sheet" Type="${office}/worksheet" Target="../data/sheet%20one.xml"/><Relationship Id="strings" Type="${office}/sharedStrings" Target="/data/strings.xml"/></Relationships>`),
      'data/sheet one.xml': strToU8(`<?xml version="1.0"?><s:worksheet xmlns:s="${ss}"><!-- 😀 <c> -->\r\n<s:sheetData>${header}<s:row><s:c t="inlineStr" odd="&gt;"><s:is><s:t>private😀&amp;value</s:t></s:is></s:c><s:c r="B2"><s:f t="shared" si="0"/><s:v>2</s:v></s:c></s:row><s:row r="3"><s:c r="A3"><s:f t="array" ref="A3:A4">SUM(A1)</s:f><s:v>2</s:v></s:c></s:row></s:sheetData><s:mergeCells><s:mergeCell ref="A1:A2"/></s:mergeCells><s:hyperlinks><s:hyperlink ref="A1:A3" location="header"/><s:hyperlink ref="B2" location="private-location"/></s:hyperlinks></s:worksheet>`),
      'data/strings.xml': strToU8(`<s:sst xmlns:s="${ss}" count="2" uniqueCount="2"><s:si><s:r><s:t>Header 😀</s:t></s:r></s:si><s:si><s:t>unused-secret</s:t></s:si></s:sst>`),
    };
    const original = zipSync(entries);
    const protectedFile = await transformOoxml(original, PASSWORD, 'protect');
    expect(protectedFile.cells).toBe(3);
    const after = unzipSync(protectedFile.data), worksheet = text(after, 'data/sheet one.xml');
    expect(worksheet).toContain(header); expect(worksheet).toContain('A1:A2'); expect(worksheet).toContain('ref="A1:A1"');
    expect(worksheet).not.toContain('private'); expect(worksheet).not.toContain('SUM(A1)');
    expect(text(after, 'data/strings.xml')).not.toContain('unused-secret');
    expect((await transformOoxml(protectedFile.data, PASSWORD, 'restore')).data).toEqual(original);
  });

  test('rejects malformed XML, DTDs, duplicate entries and resource exhaustion, not feature names', async () => {
    expect(MAX_FILE_BYTES).toBe(64 * 1024 * 1024);
    await expect(transformOoxml(new Uint8Array(MAX_FILE_BYTES + 1), PASSWORD, 'protect')).rejects.toThrow('64 MiB');
    await expect(transformOoxml(new Uint8Array([1, 2]), PASSWORD, 'protect')).rejects.toThrow();
    const base = unzipSync(await fixture());
    for (const value of ['<!DOCTYPE x [<!ENTITY foo "secret">]><x>&foo;</x>', '<x><y></x>', '<x>&unknown;</x>']) {
      await expect(transformOoxml(zipSync({ ...base, 'customXml/item1.xml': strToU8(value) }), PASSWORD, 'protect')).rejects.toThrow();
    }
    await expect(transformOoxml(zipSync({ ...base, 'xl/workbook.xml': strToU8('<notAWorkbook/>') }), PASSWORD, 'protect')).rejects.toThrow('Invalid workbook root');
    const typedXml = { ...base, 'customXml/unusual.data': strToU8('<!DOCTYPE x><x/>') };
    change(typedXml, '[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/customXml/unusual.data" ContentType="application/xml"/></Types>'));
    await expect(transformOoxml(zipSync(typedXml), PASSWORD, 'protect')).rejects.toThrow('DTDs');
    const many: Entries = { ...base };
    for (let i = 0; i < 10001; i++) many[`part${i}.bin`] = new Uint8Array();
    await expect(transformOoxml(zipSync(many), PASSWORD, 'protect')).rejects.toThrow('10,000');
    // Forge central-directory sizes: reject before fflate allocates/inflates payloads.
    const oversized = zipSync(base);
    const view = new DataView(oversized.buffer);
    for (let i = 0; i < oversized.length - 46; i++) if (view.getUint32(i, true) === 0x02014b50) { view.setUint32(i + 24, 257 * 1024 * 1024, true); break; }
    await expect(transformOoxml(oversized, PASSWORD, 'protect')).rejects.toThrow('256 MiB');
    const understated = zipSync({ ...base, 'bomb.bin': new Uint8Array(1024 * 1024) });
    const smallView = new DataView(understated.buffer);
    for (let i = 0; i < understated.length - 46; i++) {
      if (smallView.getUint32(i, true) === 0x02014b50 && strFromU8(understated.subarray(i + 46, i + 46 + smallView.getUint16(i + 28, true))) === 'bomb.bin') smallView.setUint32(i + 24, 1, true);
    }
    await expect(transformOoxml(understated, PASSWORD, 'protect')).rejects.toThrow('Invalid ZIP entry');
    const duplicate = zipSync({ a: strToU8('one'), b: strToU8('two') });
    const dv = new DataView(duplicate.buffer);
    for (let i = 0; i < duplicate.length - 31; i++) {
      const sig = dv.getUint32(i, true);
      if (sig === 0x04034b50 && duplicate[i + 30] === 98) duplicate[i + 30] = 97;
      if (sig === 0x02014b50 && duplicate[i + 46] === 98) duplicate[i + 46] = 97;
    }
    await expect(transformOoxml(duplicate, PASSWORD, 'protect')).rejects.toThrow('Duplicate');
  });
});
