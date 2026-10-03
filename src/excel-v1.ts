import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8 } from 'fflate';
import { SaxesParser } from 'saxes';

export const MAX_FILE_BYTES = 20 * 1024 * 1024;
const MAX_INFLATED_BYTES = 100 * 1024 * 1024;
const MARKER = 'ExcelKey:workbook:v1';
const CELL_MARKER = 'ExcelKey:cell:v1:';
const META_NAME = '__ExcelKey_v1';
const CHUNK_SIZE = 30_000;
const CORRUPT = 'Incorrect password, or protected workbook is damaged or modified.';
const encoder = new TextEncoder();

type SheetManifest = { id: number; name: string; cells: string[]; merges: string[] };
type Manifest = { version: 1; keywords: string | null; sheets: SheetManifest[] };

function fail(message: string): never { throw new Error(message); }
function base64(bytes: Uint8Array): string {
  let text = '';
  for (let i = 0; i < bytes.length; i += 8192) text += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(text);
}
function unbase64(text: string): Uint8Array<ArrayBuffer> {
  if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(text)) fail(CORRUPT);
  return Uint8Array.from(atob(text), c => c.charCodeAt(0));
}

// Whitelist ordinary OOXML parts: ExcelJS cannot safely preserve embedded caches.
function checkedZip(input: Uint8Array): Record<string, Uint8Array> {
  if (input.length > MAX_FILE_BYTES) fail('Workbook exceeds the 20 MB compressed limit.');
  let total = 0;
  const names = new Set<string>();
  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(input, { filter(file) {
      if (names.has(file.name)) fail('Duplicate ZIP entries are unsupported.');
      names.add(file.name);
      total += file.originalSize;
      if (total > MAX_INFLATED_BYTES || names.size > 10_000) fail('Workbook exceeds the 100 MB inflated or 10,000 ZIP-entry limit.');
      if (file.name.endsWith('/') && file.originalSize === 0) return false;
      if (!/^(?:\[Content_Types\]\.xml|_rels\/\.rels|docProps\/(?:app|core)\.xml|xl\/(?:workbook|styles|sharedStrings|calcChain)\.xml|xl\/_rels\/workbook\.xml\.rels|xl\/theme\/theme\d+\.xml|xl\/worksheets\/sheet\d+\.xml|xl\/worksheets\/_rels\/sheet\d+\.xml\.rels|xl\/comments\d+\.xml|xl\/drawings\/vmlDrawing\d+\.vml)$/.test(file.name)) {
        fail('Unsupported workbook parts (macros, charts, images, tables, pivots, external links, or embedded data).');
      }
      return true;
    } });
  } catch (error) {
    if (error instanceof Error && /^(Workbook exceeds|Unsupported workbook|Duplicate ZIP)/.test(error.message)) throw error;
    fail('Invalid .xlsx ZIP workbook.');
  }
  if (!entries['[Content_Types].xml'] || !entries['xl/workbook.xml']) fail('Invalid .xlsx ZIP workbook.');
  for (const [name, bytes] of Object.entries(entries)) {
    if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
    const xml = strFromU8(bytes);
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) fail('XML entities are unsupported.');
    if (/<(?:\w+:)?(?:dataValidations|conditionalFormatting|definedName|oleObjects|controls|customSheetViews)(?:\s|>)/.test(xml)
      || (name !== 'xl/styles.xml' && /<(?:\w+:)?extLst(?:\s|>)/.test(xml))) {
      fail('Data validation, conditional formatting, defined names, and extended workbook features are unsupported.');
    }
    if (/<(?:\w+:)?f\b[^>]*\bt\s*=\s*["'](?:array|shared|dataTable)["']/.test(xml)) {
      fail('Array, shared, and data-table formulas are unsupported.');
    }
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) {
      // Parse XML so quotes, prefixes and character references cannot bypass this check.
      const parser = new SaxesParser({ xmlns: false });
      const formulas = new Set<string>(), hyperlinks = new Set<string>();
      let address: string | undefined;
      parser.on('error', () => fail('Invalid worksheet XML.'));
      parser.on('opentag', tag => {
        const local = tag.name.split(':').pop();
        if (local === 'c') address = tag.attributes.r;
        if (local === 'f' && address) formulas.add(address);
        if (local === 'hyperlink') {
          if (tag.attributes.location !== undefined) fail('Internal location hyperlinks are unsupported.');
          if (tag.attributes.ref) hyperlinks.add(tag.attributes.ref);
        }
      });
      parser.on('closetag', tag => {
        if (tag.name.split(':').pop() === 'c') address = undefined;
      });
      parser.write(xml).close();
      if ([...hyperlinks].some(ref => formulas.has(ref))) fail('Formula cells with hyperlinks are unsupported.');
    }
    if (name.endsWith('.rels')) {
      const types = [...xml.matchAll(/\bType\s*=\s*["']([^"']+)["']/g)].map(match => match[1]!);
      if (types.some(type => !/\/(?:officeDocument|metadata\/core-properties|extended-properties|worksheet|styles|sharedStrings|theme|hyperlink|comments|vmlDrawing|calcChain)$/.test(type))) {
        fail('Unsupported workbook relationships.');
      }
    }
  }
  return entries;
}

function merges(sheet: ExcelJS.Worksheet): string[] { return [...sheet.model.merges].sort(); }
function clearNote(cell: ExcelJS.Cell): void {
  // ExcelJS's note setter creates a comment even for undefined; it has no removal API.
  delete (cell as unknown as { _comment?: unknown })._comment;
}
function dataCells(sheet: ExcelJS.Worksheet): ExcelJS.Cell[] {
  const cells: ExcelJS.Cell[] = [];
  sheet.eachRow({ includeEmpty: true }, row => {
    row.eachCell({ includeEmpty: true }, cell => {
      if (cell.isMerged && cell.master !== cell) {
        if (cell.note != null) fail('Notes on merged slave cells are unsupported.');
        return;
      }
      if (row.number === 1) {
        if (cell.type === ExcelJS.ValueType.Formula) fail('Formulas in header row 1 are unsupported.');
        return;
      }
      if (cell.value != null || cell.note != null) cells.push(cell);
    });
  });
  return cells.sort((a, b) => a.address.localeCompare(b.address));
}

function exactKeys(value: unknown, keys: string[]): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
    && Object.keys(value).sort().join(',') === [...keys].sort().join(',');
}
function parseManifest(value: unknown): Manifest {
  if (!exactKeys(value, ['version', 'keywords', 'sheets']) || value.version !== 1
    || (value.keywords !== null && typeof value.keywords !== 'string')
    || !Array.isArray(value.sheets) || !value.sheets.length || value.sheets.length > 10_000) fail(CORRUPT);
  const ids = new Set<number>(), names = new Set<string>();
  for (const sheet of value.sheets) {
    if (!exactKeys(sheet, ['id', 'name', 'cells', 'merges'])
      || !Number.isSafeInteger(sheet.id) || (sheet.id as number) < 1
      || typeof sheet.name !== 'string' || !sheet.name || sheet.name.length > 31
      || !Array.isArray(sheet.cells) || !Array.isArray(sheet.merges)) fail(CORRUPT);
    if (ids.has(sheet.id as number) || names.has(sheet.name)) fail(CORRUPT);
    ids.add(sheet.id as number); names.add(sheet.name);
    const addresses = new Set<string>();
    for (const address of sheet.cells) {
      if (typeof address !== 'string' || !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(address)) fail(CORRUPT);
      const [, col, row] = /^([A-Z]+)(\d+)$/.exec(address)!;
      const column = [...col!].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
      if (column > 16384 || Number(row) < 2 || Number(row) > 1048576 || addresses.has(address)) fail(CORRUPT);
      addresses.add(address);
    }
    if (sheet.merges.some(range => typeof range !== 'string' || !/^[A-Z]{1,3}[1-9]\d{0,6}:[A-Z]{1,3}[1-9]\d{0,6}$/.test(range))
      || new Set(sheet.merges).size !== sheet.merges.length) fail(CORRUPT);
  }
  return value as Manifest;
}

function serialize(value: unknown): string {
  return JSON.stringify(value, function (key, item) {
    const original = this[key];
    return original instanceof Date ? { $excelKeyDate: original.toISOString() } : item;
  });
}
function deserialize(text: string): unknown {
  return JSON.parse(text, (_key, value) => {
    if (exactKeys(value, ['$excelKeyDate'])) {
      if (typeof value.$excelKeyDate !== 'string' || !Number.isFinite(Date.parse(value.$excelKeyDate))) fail(CORRUPT);
      return new Date(value.$excelKeyDate);
    }
    return value;
  });
}

/** Encrypt populated cells below row 1, or atomically restore an authenticated workbook. */
export async function transformWorkbook(
  input: Uint8Array, password: string, mode: 'protect' | 'restore',
): Promise<{ data: Uint8Array; cells: number; sheets: number }> {
  if (mode !== 'protect' && mode !== 'restore') fail('Choose protect or restore.');
  if (!password) fail('Enter a password.');
  const entries = checkedZip(input);
  const workbook = new ExcelJS.Workbook();
  try {
    // Repack bounded extracted bytes; never let JSZip inflate the untrusted original.
    await workbook.xlsx.load(zipSync(entries, { level: 0 }).buffer as ArrayBuffer);
  } catch { fail('Invalid or unsupported .xlsx workbook.'); }
  if (!workbook.worksheets.length) fail('Workbook has no worksheets.');
  const originalNotes = Object.entries(entries).filter(([name]) => /^xl\/comments\d+\.xml$/.test(name))
    .reduce((count, [, bytes]) => count + [...strFromU8(bytes).matchAll(/<(?:\w+:)?comment\b/g)].length, 0);
  let loadedNotes = 0;
  let markerFound = workbook.keywords?.includes('ExcelKey:') ?? false;
  const metadataSheets: ExcelJS.Worksheet[] = [];
  for (const sheet of workbook.worksheets) {
    if (sheet.name === META_NAME) markerFound = true;
    if (sheet.getCell('A1').value === MARKER) metadataSheets.push(sheet);
    sheet.eachRow({ includeEmpty: true }, row => row.eachCell({ includeEmpty: true }, cell => {
      const value = cell.value;
      if (value && typeof value === 'object' && ('sharedFormula' in value
        || ('shareType' in value && value.shareType != null))) {
        fail('Array, shared, and data-table formulas are unsupported.');
      }
      if (cell.note != null) {
        loadedNotes++;
        if (cell.isMerged && cell.master !== cell) fail('Notes on merged slave cells are unsupported.');
      }
      if (typeof cell.value === 'string' && cell.value.startsWith('ExcelKey:')) markerFound = true;
    }));
    for (const range of merges(sheet)) {
      const [start, end] = range.split(':');
      if (Number(start!.replace(/[A-Z]/g, '')) === 1 && Number(end!.replace(/[A-Z]/g, '')) > 1) {
        fail('Merged cells spanning header row 1 and data rows are unsupported.');
      }
    }
  }
  if (originalNotes !== loadedNotes) fail('Notes without a supported worksheet cell are unsupported.');
  if (mode === 'protect' && markerFound) fail('Workbook is already protected or contains reserved ExcelKey markers.');
  if (mode === 'restore' && (metadataSheets.length !== 1 || workbook.keywords !== MARKER)) fail('ExcelKey metadata is missing or damaged.');

  let salt: Uint8Array<ArrayBuffer>;
  let manifestCipher = '';
  const metadata = metadataSheets[0];
  if (mode === 'protect') salt = crypto.getRandomValues(new Uint8Array(16));
  else {
    try {
      if (metadata!.name !== META_NAME) fail(CORRUPT);
      const header = JSON.parse(String(metadata!.getCell('A2').value));
      if (!exactKeys(header, ['version', 'salt', 'chunks']) || header.version !== 1 || typeof header.salt !== 'string'
        || !Number.isSafeInteger(header.chunks) || (header.chunks as number) < 1 || (header.chunks as number) > 5000) fail(CORRUPT);
      salt = unbase64(header.salt);
      if (salt.length !== 16 || metadata!.actualRowCount !== (header.chunks as number) + 2 || metadata!.actualColumnCount !== 1) fail(CORRUPT);
      for (let i = 0; i < (header.chunks as number); i++) {
        const chunk = metadata!.getCell(i + 3, 1).value;
        if (typeof chunk !== 'string' || !chunk.length || chunk.length > CHUNK_SIZE) fail(CORRUPT);
        manifestCipher += chunk;
      }
    } catch { fail(CORRUPT); }
  }
  const material = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 600_000, hash: 'SHA-256' }, material,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  const noncePrefix = crypto.getRandomValues(new Uint8Array(8));
  let counter = 0;
  const aad = (context: unknown[]) => encoder.encode(JSON.stringify([MARKER, base64(salt), ...context]));
  async function seal(value: unknown, context: unknown[]): Promise<string> {
    if (counter > 0xffffffff) fail('Too many encrypted cells.');
    const nonce = new Uint8Array(12);
    nonce.set(noncePrefix); new DataView(nonce.buffer).setUint32(8, counter++);
    const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce, additionalData: aad(context) }, key, encoder.encode(serialize(value))));
    const packed = new Uint8Array(nonce.length + cipher.length);
    packed.set(nonce); packed.set(cipher, nonce.length);
    return base64(packed);
  }
  async function open(text: string, context: unknown[]): Promise<unknown> {
    try {
      const packed = unbase64(text);
      if (packed.length < 28) fail(CORRUPT);
      const plain = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: packed.slice(0, 12), additionalData: aad(context) }, key, packed.slice(12));
      return deserialize(new TextDecoder('utf-8', { fatal: true }).decode(plain));
    } catch { fail(CORRUPT); }
  }
  let count = 0;
  if (mode === 'protect') {
    const manifest: Manifest = { version: 1, keywords: workbook.keywords ?? null, sheets: [] };
    for (const sheet of workbook.worksheets) {
      const cells = dataCells(sheet);
      manifest.sheets.push({ id: sheet.id, name: sheet.name, cells: cells.map(cell => cell.address), merges: merges(sheet) });
      for (const cell of cells) {
        const cipher = CELL_MARKER + await seal({ value: cell.value ?? null, note: cell.note ?? null }, ['cell', sheet.id, sheet.name, cell.address]);
        if (cipher.length > 32767) fail(`Encrypted cell ${sheet.name}!${cell.address} exceeds Excel's 32,767-character limit.`);
        cell.value = cipher;
        clearNote(cell);
        count++;
      }
    }
    manifestCipher = await seal(manifest, ['manifest']);
    const chunks = Math.ceil(manifestCipher.length / CHUNK_SIZE);
    if (chunks > 5000) fail('Encrypted manifest exceeds the metadata limit.');
    const meta = workbook.addWorksheet(META_NAME, { state: 'veryHidden' });
    meta.getCell('A1').value = MARKER;
    meta.getCell('A2').value = JSON.stringify({ version: 1, salt: base64(salt), chunks });
    for (let i = 0; i < chunks; i++) meta.getCell(i + 3, 1).value = manifestCipher.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    workbook.keywords = MARKER;
  } else {
    const manifest = parseManifest(await open(manifestCipher, ['manifest']));
    const sheets = workbook.worksheets.filter(sheet => sheet !== metadata);
    if (sheets.length !== manifest.sheets.length) fail(CORRUPT);
    for (const [index, expected] of manifest.sheets.entries()) {
      const sheet = sheets[index]!;
      if (sheet.id !== expected.id || sheet.name !== expected.name || JSON.stringify(merges(sheet)) !== JSON.stringify(expected.merges)) fail(CORRUPT);
      const cells = dataCells(sheet);
      if (cells.length !== expected.cells.length || cells.some((cell, i) => cell.address !== expected.cells[i])) fail(CORRUPT);
      for (const cell of cells) {
        if (typeof cell.value !== 'string' || !cell.value.startsWith(CELL_MARKER) || cell.note != null) fail(CORRUPT);
        const payload = await open(cell.value.slice(CELL_MARKER.length), ['cell', sheet.id, sheet.name, cell.address]);
        if (!exactKeys(payload, ['value', 'note'])) fail(CORRUPT);
        cell.value = payload.value as ExcelJS.CellValue;
        if (payload.note !== null) cell.note = payload.note as string | ExcelJS.Comment;
        count++;
      }
    }
    workbook.removeWorksheet(metadata!.id);
    workbook.keywords = manifest.keywords ?? '';
    workbook.calcProperties.fullCalcOnLoad = true;
  }
  const data = new Uint8Array(await workbook.xlsx.writeBuffer());
  checkedZip(data);
  return { data, cells: count, sheets: workbook.worksheets.length - (mode === 'protect' ? 1 : 0) };
}
