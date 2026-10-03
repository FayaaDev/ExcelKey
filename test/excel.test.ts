import { expect, test } from 'bun:test';
import ExcelJS from 'exceljs';
import { read, write, utils } from 'xlsx';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { transformWorkbook } from '../src/excel';
import { transformWorkbook as transformV1 } from '../src/excel-v1';
import { syntheticWorkbook } from '../src/excel-synthetic';

const PASSWORD = 'compatibility-test-password';

test('public API strips feature data and restores original bytes', async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addRow(['Header']);
  sheet.addRow(['original-data']);
  sheet.getCell('A2').dataValidation = { type: 'list', formulae: ['"One,Two"'] };
  sheet.addConditionalFormatting({ ref: 'A2', rules: [{ type: 'expression', formulae: ['A2="One"'], priority: 1, style: { font: { bold: true } } }] });
  workbook.definedNames.add('Data!$A$2', 'NamedValue');
  const sourceParts = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()));
  sourceParts['xl/charts/chart1.xml'] = strToU8('<chart>private-chart-cache</chart>');
  sourceParts['xl/pivotCache/pivotCacheRecords1.xml'] = strToU8('<records>private-pivot-cache</records>');
  sourceParts['xl/vbaProject.bin'] = strToU8('private-macro');
  sourceParts['customXml/item1.xml'] = strToU8('<data>private-feature</data>');
  const original = zipSync(sourceParts);
  const protectedFile = await transformWorkbook(original, PASSWORD, 'protect', { filename: 'features.xlsx' });
  expect(protectedFile.cells).toBe(1);
  const parts = unzipSync(protectedFile.data);
  expect(strFromU8(parts['xl/worksheets/sheet1.xml']!)).not.toContain('dataValidation');
  expect(strFromU8(parts['xl/worksheets/sheet1.xml']!)).not.toContain('conditionalFormatting');
  expect(strFromU8(parts['xl/workbook.xml']!)).not.toContain('NamedValue');
  for (const name of ['xl/charts/chart1.xml', 'xl/pivotCache/pivotCacheRecords1.xml', 'xl/vbaProject.bin', 'customXml/item1.xml']) expect(parts[name]).toBeUndefined();
  const restored = await transformWorkbook(protectedFile.data, PASSWORD, 'restore');
  expect(restored.data).toEqual(original);
  expect(restored.filename).toBe('features.xlsx');
});

test('preserves 1900 date serials, plain date columns, percentages and empty cells', () => {
  const book = utils.book_new();
  const sheet = utils.aoa_to_sheet([['Date', 'Timestamp', 'Percent', 'Amount', 'Empty', 'Text'], [60, 45292.5, 0.35, 42, '', 'private-value']]);
  sheet.A2.z = 'm/d/yy'; sheet.B2.z = 'yyyy-mm-dd hh:mm:ss'; sheet.C2.z = '0.00%';
  sheet.A3 = { t: 'n', v: 20240101 }; sheet.B3 = { t: 's', v: 'October 3, 2026 2:30 PM' }; sheet['!ref'] = 'A1:F3';
  utils.book_append_sheet(book, sheet, 'Data');
  const input = new Uint8Array(write(book, { type: 'array', bookType: 'xlsx' }));
  const first = syntheticWorkbook(input, PASSWORD), second = syntheticWorkbook(input, PASSWORD);
  const generated = read(first.data, { type: 'array', cellNF: true });
  const data = generated.Sheets['Worksheet 1']!;
  expect(generated.Workbook?.WBProps?.date1904).toBe(false);
  expect(data.A2.v).toBe(60); expect(data.A2.z).toBe('m/d/yy');
  expect(data.B2.v).toBe(45292.5); expect(data.B2.z).toBe('yyyy-mm-dd hh:mm:ss');
  expect(data.A3.v).toBe(20240101); expect(data.B3.v).toBe('October 3, 2026 2:30 PM');
  expect(data.C2.t).toBe('n'); expect(data.C2.v).toBeGreaterThanOrEqual(0.01); expect(data.C2.v).toBeLessThanOrEqual(1); expect(data.C2.z).toBe('0.00%');
  expect(data.D2.v).not.toBe(42); expect(data.E2?.v ?? '').toBe(''); expect(data.C3).toBeUndefined();
  expect(data.F2.v).not.toBe(read(second.data, { type: 'array' }).Sheets['Worksheet 1']!.F2.v);
});

test('requires cached date formula results instead of changing dates silently', async () => {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet('Data');
  sheet.addRows([['Date'], [{ formula: 'TODAY()' }]]);
  sheet.getCell('A2').numFmt = 'yyyy-mm-dd';
  const input = new Uint8Array(await book.xlsx.writeBuffer());
  await expect(transformWorkbook(input, PASSWORD, 'protect')).rejects.toThrow('Recalculate and save');
});

test('imports binary Excel formats and restores the exact original format', async () => {
  const book = utils.book_new();
  utils.book_append_sheet(book, utils.aoa_to_sheet([['Name', 'Value'], ['private-data', 42]]), 'Data');
  for (const bookType of ['xls', 'xlsb'] as const) {
    const original = new Uint8Array(write(book, { type: 'array', bookType }));
    const protectedFile = await transformWorkbook(original, PASSWORD, 'protect', { filename: `binary.${bookType}` });
    expect(protectedFile.filename).toBe('synthetic.xlsx');
    expect(protectedFile.cells).toBe(2);
    const visible = read(protectedFile.data, { type: 'array' });
    expect(visible.Sheets['Worksheet 1']!.A1.v).toBe('Name');
    expect(visible.Sheets['Worksheet 1']!.A2.v).not.toBe('private-data');
    await expect(transformWorkbook(protectedFile.data, 'wrong', 'restore')).rejects.toThrow('Incorrect password');
    const restored = await transformWorkbook(protectedFile.data, PASSWORD, 'restore');
    expect(restored.data).toEqual(original);
    expect(restored.filename).toBe(`binary.${bookType}`);
  }
});

test('generates typed, consistent synthetic data, preserves dates, and excludes source payloads', async () => {
  const book = new ExcelJS.Workbook();
  book.creator = 'private-author';
  book.properties.date1904 = true;
  const sheet = book.addWorksheet('private-sheet');
  const headers = ['Full name', 'Email', 'Customer ID', 'Amount', 'Ratio', 'Active', 'Date', 'Signup date', 'Category', 'Formula', 'Uncached', 'Birth date'];
  sheet.addRow(headers);
  sheet.addRow(['private-name', 'private-email@private.test', 'private-id', 987654321, 99999.1234, true, new Date('2024-02-29T12:34:56Z'), '03/04/2024', 'private-category', { formula: '987654321+1', result: 987654322 }, { formula: '987654321+2' }, { formula: 'DATE(2024,2,29)', result: 43889 }]);
  sheet.addRow(['private-name', 'private-email@private.test', 'private-id', 987654321, 99999.1234, true, new Date('2024-02-29T12:34:56Z'), '03/04/2024', 'private-category']);
  sheet.getCell('G2').numFmt = sheet.getCell('G3').numFmt = 'yyyy-mm-dd';
  sheet.getCell('L2').numFmt = 'yyyy-mm-dd';
  sheet.getCell('D2').numFmt = '0.00"private-format"';
  sheet.getCell('A2').note = 'private-note';
  sheet.getCell('B2').value = { text: 'private-email@private.test', hyperlink: 'https://private.test/private-link' };
  const hidden = book.addWorksheet('private-hidden', { state: 'veryHidden' });
  hidden.addRows([['Customer ID', 'Full name', 'Category'], ['private-id', 'private-name', 'private-category']]);
  const original = new Uint8Array(await book.xlsx.writeBuffer());
  const before = read(original, { type: 'array', cellNF: true });
  const generated = await transformWorkbook(original, PASSWORD, 'protect', { filename: 'private-file.xlsm' });
  const visible = read(generated.data, { type: 'array', cellNF: true });
  expect(visible.SheetNames).toEqual(['Worksheet 1', 'Worksheet 2']);
  const data = visible.Sheets['Worksheet 1']!;
  expect(utils.sheet_to_json(data, { header: 1 })[0]).toEqual(headers);
  expect(data.A2.v).toMatch(/^[A-Za-z]+ [A-Za-z]+(?: \d+)?$/);
  expect(data.B2.v).toMatch(/@example\.com$/);
  for (const column of ['A', 'B', 'C', 'D', 'E', 'F', 'I']) expect(data[`${column}2`].v).toBe(data[`${column}3`].v);
  expect(visible.Sheets['Worksheet 2']!.A2.v).toBe(data.C2.v);
  expect(visible.Sheets['Worksheet 2']!.B2.v).toBe(data.A2.v);
  expect(visible.Sheets['Worksheet 2']!.C2.v).toBe(data.I2.v);
  expect(data.D2.t).toBe('n'); expect(data.D2.v).toBeGreaterThanOrEqual(1); expect(data.D2.v).toBeLessThanOrEqual(1000);
  expect(data.E2.t).toBe('n'); expect(data.E2.v).toBeGreaterThanOrEqual(0.01); expect(data.E2.v).toBeLessThanOrEqual(1000);
  expect(data.F2.t).toBe('b');
  expect(data.G2.v).toBe(before.Sheets['private-sheet']!.G2.v); expect(data.G2.z).toBe('yyyy-mm-dd');
  expect(data.H2.v).toBe('03/04/2024');
  expect(data.L2.v).toBe(before.Sheets['private-sheet']!.L2.v); expect(data.L2.f).toBeUndefined();
  expect(visible.Workbook?.WBProps?.date1904).toBe(true);
  expect(data.J2.t).toBe('n'); expect(data.J2.f).toBeUndefined(); expect(data.K2.v).toBeDefined();
  expect(data.J3).toBeUndefined();
  const parts = unzipSync(generated.data);
  const cleartext = Object.entries(parts).filter(([name]) => !name.startsWith('excelkey/')).map(([, bytes]) => strFromU8(bytes)).join('\n');
  for (const secret of ['private-', '987654321', '99999.1234', 'DATE(2024,2,29)']) expect(cleartext).not.toContain(secret);
  expect((await transformWorkbook(generated.data, PASSWORD, 'restore')).data).toEqual(original);
  await expect(transformWorkbook(generated.data, 'wrong', 'restore')).rejects.toThrow('Incorrect password');
  const modified = generated.data.slice(); modified[100] = modified[100]! ^ 1;
  await expect(transformWorkbook(modified, PASSWORD, 'restore')).rejects.toThrow();
});

test('restores existing v1 workbooks through the public API', async () => {
  const book = new ExcelJS.Workbook();
  book.addWorksheet('Data').addRows([['Header'], ['legacy-data']]);
  const original = new Uint8Array(await book.xlsx.writeBuffer());
  const protectedFile = await transformV1(original, PASSWORD, 'protect');
  await expect(transformWorkbook(protectedFile.data, PASSWORD, 'protect')).rejects.toThrow('already protected');
  const restored = await transformWorkbook(protectedFile.data, PASSWORD, 'restore');
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(restored.data.buffer as ArrayBuffer);
  expect(loaded.getWorksheet('Data')!.getCell('A2').value).toBe('legacy-data');
});

test('rejects empty files and rejects restoration without ExcelKey metadata', async () => {
  await expect(transformWorkbook(new Uint8Array(), PASSWORD, 'protect')).rejects.toThrow('valid Excel');
  await expect(transformWorkbook(new Uint8Array([1]), PASSWORD, 'restore')).rejects.toThrow('protected workbook');
});
