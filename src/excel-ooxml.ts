import { Unzip, UnzipInflate, Zip, ZipDeflate, unzipSync } from 'fflate';
import { SaxesParser } from 'saxes';
import type { SaxesTagNS } from 'saxes';

export const MAX_FILE_BYTES = 64 * 1024 * 1024;
const MAX_INFLATED_BYTES = 256 * 1024 * 1024;
const TYPE = 'urn:excelkey:protected-workbook:v2';
const CONTENT_TYPE = 'application/vnd.excelkey.protected-workbook+json';
const CORRUPT = 'Incorrect password, or protected workbook is damaged or modified.';
const SS = new Set(['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main']);
const REL = new Set(['http://schemas.openxmlformats.org/package/2006/relationships', 'http://purl.oclc.org/ooxml/package/relationships']);
const OFFICE = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/', 'http://purl.oclc.org/ooxml/officeDocument/relationships/'];
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const encoder = new TextEncoder();
type Entries = Record<string, Uint8Array>;
type Node = { tag: SaxesTagNS; start: number; openEnd: number; end: number; parent?: Node; text?: string };
type Edit = { start: number; end: number; text: string };
type Relationship = { id: string; type: string; target: string; external: boolean; node: Node };
type Manifest = { version: 2; digest: string; cells: number; sheets: number; filename?: string };

function fail(message: string): never { throw new Error(message); }
function attr(node: Node, name: string): string | undefined { return node.tag.attributes[name]?.value; }
function escape(text: string): string { return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;').replace(/\r/g, '&#13;').replace(/\n/g, '&#10;').replace(/\t/g, '&#9;'); }
function opening(node: Node, changes: Record<string, string>): string {
  const attrs = Object.fromEntries(Object.values(node.tag.attributes).map(a => [a.name, a.value]));
  Object.assign(attrs, changes);
  return `<${node.tag.name}${Object.entries(attrs).map(([k, v]) => ` ${k}="${escape(v)}"`).join('')}>`;
}
function element(node: Node, local: string): string { return node.tag.prefix ? `${node.tag.prefix}:${local}` : local; }

function xml(bytes: Uint8Array): { text: string; encode(text: string): Uint8Array } {
  const little = bytes[0] === 255 && bytes[1] === 254 || bytes[0] === 60 && bytes[1] === 0;
  const big = bytes[0] === 254 && bytes[1] === 255 || bytes[0] === 0 && bytes[1] === 60;
  const encoding = little ? 'utf-16le' : big ? 'utf-16be' : 'utf-8';
  let text: string;
  try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); } catch { fail('Invalid workbook XML encoding.'); }
  return { text, encode(value) {
    if (!little && !big) return encoder.encode(value);
    const out = new Uint8Array(2 + value.length * 2);
    const view = new DataView(out.buffer);
    view.setUint16(0, 0xfeff, little);
    for (let i = 0; i < value.length; i++) view.setUint16(2 + i * 2, value.charCodeAt(i), little);
    return out;
  } };
}

// Saxes positions are UTF-16 string offsets, including astral characters and CRLF.
function scan(text: string, close: (node: Node) => void, open?: (node: Node) => void): void {
  const parser = new SaxesParser({ xmlns: true });
  const stack: Node[] = [];
  let start = 0;
  parser.on('doctype', () => fail('XML DTDs and entities are unsupported.'));
  parser.on('error', () => fail('Invalid workbook XML.'));
  parser.on('opentagstart', () => { start = text.lastIndexOf('<', parser.position - 1); });
  parser.on('opentag', tag => {
    const node: Node = { tag, start, openEnd: parser.position, end: parser.position, parent: stack.at(-1) };
    stack.push(node); open?.(node);
  });
  const textValue = (value: string) => { const node = stack.at(-1); if (node) node.text = (node.text ?? '') + value; };
  parser.on('text', textValue);
  parser.on('cdata', textValue);
  parser.on('closetag', () => { const node = stack.pop()!; node.end = parser.position; close(node); });
  parser.write(text).close();
}
function edited(text: string, edits: Edit[]): string {
  edits.sort((a, b) => a.start - b.start);
  let end = 0;
  const chunks: string[] = [];
  for (const edit of edits) {
    if (edit.start < end) fail('Overlapping workbook XML edits.');
    chunks.push(text.slice(end, edit.start), edit.text); end = edit.end;
  }
  chunks.push(text.slice(end)); return chunks.join('');
}
function replacePart(entries: Entries, path: string, edits: (text: string) => Edit[]): void {
  const doc = xml(entries[path] ?? fail(`Missing workbook part: ${path}`));
  const changes = edits(doc.text);
  if (changes.length) entries[path] = doc.encode(edited(doc.text, changes));
}
function checkedZip(input: Uint8Array): Entries {
  if (input.length > MAX_FILE_BYTES) fail('Workbook exceeds the 64 MiB compressed limit.');
  let total = 0;
  const sizes = new Map<string, number>();
  const entries: Entries = Object.create(null);
  const completed = new Set<string>();
  try {
    // Preflight central-directory limits before allocating any entry buffers.
    unzipSync(input, { filter(file) {
      if (sizes.has(file.name)) fail('Duplicate ZIP entries are unsupported.');
      if (!file.name || file.name.startsWith('/') || file.name.includes('\\') || file.name.split('/').some(p => p === '..' || p === '.')) fail('Invalid ZIP entry path.');
      sizes.set(file.name, file.originalSize); total += file.originalSize;
      if (total > MAX_INFLATED_BYTES || sizes.size > 10_000) fail('Workbook exceeds the 256 MiB inflated or 10,000 ZIP-entry limit.');
      return false;
    } });
    const archive = new Unzip(file => {
      const size = sizes.get(file.name);
      if (size === undefined || Object.hasOwn(entries, file.name)) fail('Invalid ZIP entry directory.');
      const bytes = new Uint8Array(size);
      entries[file.name] = bytes;
      let offset = 0;
      file.ondata = (error, chunk, final) => {
        if (error) throw error;
        if (offset + chunk.length > size) fail('Invalid ZIP entry size.');
        bytes.set(chunk, offset); offset += chunk.length;
        if (final && offset !== size) fail('Invalid ZIP entry size.');
        if (final) completed.add(file.name);
      };
      file.start();
    });
    archive.register(UnzipInflate);
    // Bound actual inflation too: forged advertised sizes cannot hide a ZIP bomb.
    for (let i = 0; i < input.length; i += 8192) archive.push(input.subarray(i, i + 8192), i + 8192 >= input.length);
    if (completed.size !== sizes.size) fail('Invalid ZIP entry directory.');
  } catch (error) {
    if (error instanceof Error && /^(Workbook exceeds|Duplicate ZIP|Invalid ZIP entry)/.test(error.message)) throw error;
    fail('Invalid OOXML ZIP workbook.');
  }
  let actual = 0;
  const xmlParts = new Set<string>();
  const xmlExtensions = new Set(['xml', 'rels', 'vml']);
  if (entries['[Content_Types].xml']) scan(xml(entries['[Content_Types].xml']).text, node => {
    if (!node.parent && (node.tag.local !== 'Types' || node.tag.uri !== CT)) fail('Invalid package content types.');
    if (node.tag.uri !== CT) return;
    const type = attr(node, 'ContentType');
    if (!type || !/(?:\+xml|\/xml)$/i.test(type)) return;
    if (node.tag.local === 'Override') {
      const name = attr(node, 'PartName');
      if (name) xmlParts.add(resolve('', name));
    }
    if (node.tag.local === 'Default') {
      const extension = attr(node, 'Extension');
      if (extension) xmlExtensions.add(extension.toLowerCase());
    }
  });
  for (const [path, bytes] of Object.entries(entries)) {
    actual += bytes.length;
    if (xmlParts.has(path) || xmlExtensions.has(path.split('.').at(-1)!.toLowerCase())) scan(xml(bytes).text, () => {});
  }
  if (actual > MAX_INFLATED_BYTES || !entries['[Content_Types].xml'] || !entries['_rels/.rels']) fail('Invalid or oversized OOXML package.');
  return entries;
}
function pack(entries: Entries): Uint8Array {
  const chunks: Uint8Array[] = [];
  let size = 0;
  // Streaming ZIP also preserves literal names such as __proto__, unlike zipSync.
  const archive = new Zip((error, chunk) => {
    if (error) throw error;
    size += chunk.length;
    if (size > MAX_FILE_BYTES) fail('Protected workbook exceeds the 64 MiB compressed limit.');
    chunks.push(chunk);
  });
  for (const [name, bytes] of Object.entries(entries)) {
    const file = new ZipDeflate(name, { level: 6 });
    archive.add(file); file.push(bytes, true);
  }
  archive.end();
  const data = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
  return data;
}
function relationships(entries: Entries, path: string): Relationship[] {
  if (!entries[path]) return [];
  const result: Relationship[] = [];
  const ids = new Set<string>();
  scan(xml(entries[path]).text, node => {
    if (!node.parent && (node.tag.local !== 'Relationships' || !REL.has(node.tag.uri))) fail('Invalid package relationships.');
    if (node.tag.local !== 'Relationship' || !REL.has(node.tag.uri) || node.parent?.tag.local !== 'Relationships') return;
    const id = attr(node, 'Id'), type = attr(node, 'Type'), target = attr(node, 'Target');
    if (!id || !type || !target || ids.has(id)) fail('Invalid package relationship.');
    ids.add(id); result.push({ id, type, target, external: attr(node, 'TargetMode') === 'External', node });
  });
  return result;
}
function isOffice(rel: Relationship, kind: string): boolean { return OFFICE.some(prefix => rel.type === prefix + kind); }
function resolve(source: string, target: string): string {
  try {
    const url = new URL(target, `https://package.invalid/${source}`);
    if (url.origin !== 'https://package.invalid' || url.search || url.hash) fail('Invalid package target.');
    const path = decodeURIComponent(url.pathname).slice(1);
    if (!path || path.includes('\\') || path.split('/').some(p => p === '..' || p === '.')) fail('Invalid package target.');
    return path;
  } catch { fail('Invalid package target.'); }
}
function relPath(path: string): string { const slash = path.lastIndexOf('/'); return `${path.slice(0, slash + 1)}_rels/${path.slice(slash + 1)}.rels`; }
function metadata(entries: Entries): string | undefined {
  const matches = relationships(entries, '_rels/.rels').filter(r => r.type === TYPE);
  if (!matches.length) return undefined;
  if (matches.length !== 1 || matches[0]!.external) fail(CORRUPT);
  return resolve('', matches[0]!.target);
}
/** Detect the v2 root relationship, including a dangling metadata reference. */
export function hasV2Protection(input: Uint8Array): boolean { return metadata(checkedZip(input)) !== undefined; }

function base64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  for (let i = 0; i < bytes.length; i += 8192) chunks.push(String.fromCharCode(...bytes.subarray(i, i + 8192)));
  return btoa(chunks.join(''));
}
function unbase64(text: unknown): Uint8Array<ArrayBuffer> {
  if (typeof text !== 'string' || text.length % 4 || !/^[A-Za-z0-9+/]*={0,2}$/.test(text)) fail(CORRUPT);
  try { return Uint8Array.from(atob(text), c => c.charCodeAt(0)); } catch { fail(CORRUPT); }
}
async function hash(bytes: Uint8Array): Promise<string> { return base64(new Uint8Array(await crypto.subtle.digest('SHA-256', new Uint8Array(bytes)))); }
async function digest(entries: Entries, excluded: string): Promise<string> {
  const list: [string, number, string][] = [];
  for (const path of Object.keys(entries).sort()) if (path !== excluded) list.push([path, entries[path]!.length, await hash(entries[path]!)]);
  return hash(encoder.encode(JSON.stringify(list)));
}
async function keyFor(password: string, salt: Uint8Array<ArrayBuffer>): Promise<CryptoKey> {
  const material = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 600_000, hash: 'SHA-256' }, material, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
function insert(entries: Entries, path: string, make: (root: Node) => string): void {
  replacePart(entries, path, text => {
    let root: Node | undefined;
    scan(text, node => { if (!node.parent) root = node; });
    if (!root) fail('Invalid package XML.');
    if (root.tag.isSelfClosing) return [{ start: root.start, end: root.end, text: opening(root, {}) + make(root) + `</${root.tag.name}>` }];
    const end = text.lastIndexOf('</', root.end);
    return [{ start: end, end, text: make(root) }];
  });
}
function rowOf(ref: string | undefined): number | undefined { const match = /^\$?[A-Za-z]+\$?(\d+)$/.exec(ref ?? ''); return match ? Number(match[1]) : undefined; }
function headerRange(ref: string): string | undefined {
  const [first, last = first] = ref.split(':');
  if (rowOf(first) !== 1) return undefined;
  return rowOf(last) === 1 ? ref : `${first}:${last!.replace(/\d+$/, '1')}`;
}

function protectSheets(entries: Entries): { cells: number; sheets: number } {
  const roots = relationships(entries, '_rels/.rels').filter(r => isOffice(r, 'officeDocument') && !r.external);
  if (roots.length !== 1) fail('Missing or ambiguous workbook relationship.');
  const workbook = resolve('', roots[0]!.target);
  const rels = relationships(entries, relPath(workbook));
  const sheets: string[] = [];
  scan(xml(entries[workbook] ?? fail('Missing workbook part.')).text, node => {
    if (!node.parent && (node.tag.local !== 'workbook' || !SS.has(node.tag.uri))) fail('Invalid workbook root.');
    if (node.tag.local !== 'sheet' || !SS.has(node.tag.uri) || node.parent?.tag.local !== 'sheets') return;
    const id = Object.values(node.tag.attributes).find(a => a.local === 'id' && OFFICE.some(prefix => a.uri + '/' === prefix))?.value;
    const rel = rels.find(r => r.id === id);
    if (!rel || rel.external) fail('Missing sheet relationship.');
    if (isOffice(rel, 'worksheet')) sheets.push(resolve(workbook, rel.target));
  });
  if (new Set(sheets).size !== sheets.length) fail('Duplicate worksheet target.');
  const headerStrings = new Set<number>();
  let headerStringCount = 0, cells = 0;
  for (const path of sheets) {
    replacePart(entries, path, text => {
      const changes: Edit[] = [];
      const populated = new Set<Node>();
      const headerValues = new Map<Node, string>();
      const rowNumbers = new Map<Node, number>();
      let nextRow = 0;
      scan(text, node => {
        const parent = node.parent;
        if (SS.has(node.tag.uri) && ['v', 'f', 'is'].includes(node.tag.local) && parent?.tag.local === 'c' && SS.has(parent.tag.uri)) populated.add(parent);
        if (node.tag.local === 'c' && SS.has(node.tag.uri) && parent?.tag.local === 'row') {
          const row = rowOf(attr(node, 'r')) ?? rowNumbers.get(parent);
          if (!row || row < 1) fail('Invalid worksheet cell address.');
          if (row === 1) {
            if (attr(node, 't') === 's') {
              // Read the already namespace-resolved direct value without reparsing a fragment.
              const value = headerValues.get(node);
              const index = Number(value);
              if (value === undefined || !Number.isSafeInteger(index) || index < 0) fail('Invalid shared string index.');
              headerStrings.add(index); headerStringCount++;
            }
          } else if (populated.has(node)) {
            const is = element(node, 'is'), t = element(node, 't');
            changes.push({ start: node.start, end: node.end, text: opening(node, { t: 'inlineStr' }) + `<${is}><${t}>${base64(crypto.getRandomValues(new Uint8Array(12)))}</${t}></${is}></${node.tag.name}>` });
            cells++;
          }
          populated.delete(node); headerValues.delete(node);
        }
        if (node.tag.local === 'v' && SS.has(node.tag.uri) && parent?.tag.local === 'c') headerValues.set(parent, (node.text ?? '').trim());
        if (node.tag.local === 'row') rowNumbers.delete(node);
        if (node.tag.local === 'hyperlink' && SS.has(node.tag.uri)) {
          const ref = attr(node, 'ref');
          const keep = ref && headerRange(ref);
          if (!keep) changes.push({ start: node.start, end: node.end, text: '' });
          else if (keep !== ref) changes.push({ start: node.start, end: node.end, text: opening(node, { ref: keep }).replace(/>$/, '/>') });
        }
      }, node => {
        if (!node.parent && (node.tag.local !== 'worksheet' || !SS.has(node.tag.uri))) fail('Invalid worksheet root.');
        if (node.tag.local === 'row' && SS.has(node.tag.uri) && node.parent?.tag.local === 'sheetData') {
          const row = Number(attr(node, 'r') ?? nextRow + 1);
          if (!Number.isSafeInteger(row) || row < 1) fail('Invalid worksheet row.');
          nextRow = row; rowNumbers.set(node, row);
        }
      });
      return changes;
    });
    const sheetRels = relationships(entries, relPath(path));
    const usedLinks = new Set<string>();
    scan(xml(entries[path]!).text, node => {
      if (node.tag.local === 'hyperlink' && SS.has(node.tag.uri)) for (const a of Object.values(node.tag.attributes)) if (a.local === 'id') usedLinks.add(a.value);
    });
    const remove = sheetRels.filter(r => isOffice(r, 'hyperlink') && !usedLinks.has(r.id));
    if (remove.length) replacePart(entries, relPath(path), () => remove.map(r => ({ start: r.node.start, end: r.node.end, text: '' })));
    for (const rel of sheetRels) {
      if (rel.external) continue;
      if (isOffice(rel, 'comments') || /\/threadedComment$/.test(rel.type)) replacePart(entries, resolve(path, rel.target), text => {
        const changes: Edit[] = [];
        scan(text, node => {
          if ((node.tag.local === 'comment' && SS.has(node.tag.uri) || node.tag.local === 'threadedComment' && node.tag.uri === 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments')
            && rowOf(attr(node, 'ref')) !== 1) changes.push({ start: node.start, end: node.end, text: '' });
        });
        return changes;
      });
      if (isOffice(rel, 'vmlDrawing')) replacePart(entries, resolve(path, rel.target), text => {
        const changes: Edit[] = [];
        const dataShapes = new Set<Node>();
        scan(text, node => {
          if (node.tag.local === 'Row' && node.tag.uri === 'urn:schemas-microsoft-com:office:excel') {
            let shape = node.parent;
            while (shape && !(shape.tag.local === 'shape' && shape.tag.uri === 'urn:schemas-microsoft-com:vml')) shape = shape.parent;
            if (shape && Number(node.text) > 0) dataShapes.add(shape);
          }
          if (dataShapes.has(node)) changes.push({ start: node.start, end: node.end, text: '' });
        });
        return changes;
      });
    }
  }
  const shared = rels.filter(r => isOffice(r, 'sharedStrings') && !r.external);
  if (shared.length > 1 || headerStrings.size && !shared.length) fail('Invalid shared string relationship.');
  if (shared.length) replacePart(entries, resolve(workbook, shared[0]!.target), text => {
    const changes: Edit[] = [];
    let index = 0, root: Node | undefined;
    scan(text, node => {
      if (node.tag.local === 'si' && SS.has(node.tag.uri) && node.parent?.tag.local === 'sst') {
        if (!headerStrings.has(index)) changes.push({ start: node.start, end: node.end, text: opening(node, {}) + `<${element(node, 't')}></${element(node, 't')}></${node.tag.name}>` });
        index++;
      }
      if (!node.parent) root = node;
    });
    if (!root || root.tag.local !== 'sst' || !SS.has(root.tag.uri) || [...headerStrings].some(i => i >= index)) fail('Invalid shared string table.');
    changes.push({ start: root.start, end: root.openEnd, text: opening(root, { count: String(headerStringCount), uniqueCount: String(index) }) + (root.tag.isSelfClosing ? `</${root.tag.name}>` : '') });
    return changes;
  });
  return { cells, sheets: sheets.length };
}

/** Encrypt the original snapshot; protect worksheet values while retaining opaque feature parts. */
export async function transformOoxml(
  input: Uint8Array, password: string, mode: 'protect' | 'restore',
  options: { original?: Uint8Array; filename?: string; syntheticCounts?: { cells: number; sheets: number } } = {},
): Promise<{ data: Uint8Array; cells: number; sheets: number; filename?: string }> {
  if (mode !== 'protect' && mode !== 'restore') fail('Choose protect or restore.');
  if (!password) fail('Enter a password.');
  const entries = checkedZip(input);
  const existing = metadata(entries);
  if (mode === 'restore') {
    try {
      if (!existing || !entries[existing]) fail(CORRUPT);
      const envelope = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(entries[existing]));
      if (envelope.version !== 2) fail(CORRUPT);
      const salt = unbase64(envelope.salt), nonce = unbase64(envelope.nonce), ciphertext = unbase64(envelope.ciphertext);
      if (salt.length !== 16 || nonce.length !== 12 || ciphertext.length < 20) fail(CORRUPT);
      const plain = new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: nonce, additionalData: encoder.encode(JSON.stringify([TYPE, existing])) }, await keyFor(password, salt), ciphertext));
      const length = new DataView(plain.buffer).getUint32(0);
      if (length > 65536 || length + 4 > plain.length) fail(CORRUPT);
      const manifest: Manifest = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(plain.subarray(4, 4 + length)));
      if (manifest.version !== 2 || !Number.isSafeInteger(manifest.cells) || manifest.cells < 0 || !Number.isSafeInteger(manifest.sheets) || manifest.sheets < 0
        || manifest.filename !== undefined && typeof manifest.filename !== 'string' || manifest.digest !== await digest(entries, existing)) fail(CORRUPT);
      const data = plain.slice(4 + length);
      if (data.length > MAX_FILE_BYTES) fail(CORRUPT);
      return { data, cells: manifest.cells, sheets: manifest.sheets, filename: manifest.filename };
    } catch { fail(CORRUPT); }
  }
  if (existing) fail('Workbook is already protected.');
  const original = options.original ?? input;
  if (original.length > MAX_FILE_BYTES) fail('Workbook exceeds the 64 MiB compressed limit.');
  if (options.filename !== undefined && (typeof options.filename !== 'string' || encoder.encode(options.filename).length > 4096)) fail('Invalid original filename.');
  const counts = options.syntheticCounts ?? protectSheets(entries);
  let path: string;
  const overrides = new Set<string>();
  scan(xml(entries['[Content_Types].xml']!).text, node => { if (node.tag.local === 'Override' && node.tag.uri === CT) overrides.add(attr(node, 'PartName') ?? ''); });
  do { path = `excelkey/${base64(crypto.getRandomValues(new Uint8Array(12))).replace(/[+/]/g, '_')}.json`; } while (entries[path] || overrides.has('/' + path));
  const rootRels = relationships(entries, '_rels/.rels');
  let id = 'ExcelKeyV2';
  while (rootRels.some(r => r.id === id)) id += '_';
  insert(entries, '_rels/.rels', root => `<${element(root, 'Relationship')} Id="${id}" Type="${TYPE}" Target="/${path}"/>`);
  insert(entries, '[Content_Types].xml', root => `<${element(root, 'Override')} PartName="/${path}" ContentType="${CONTENT_TYPE}"/>`);
  const manifest: Manifest = { version: 2, digest: await digest(entries, path), ...counts, ...(options.filename === undefined ? {} : { filename: options.filename }) };
  const header = encoder.encode(JSON.stringify(manifest));
  const plain = new Uint8Array(4 + header.length + original.length);
  new DataView(plain.buffer).setUint32(0, header.length); plain.set(header, 4); plain.set(original, 4 + header.length);
  const salt = crypto.getRandomValues(new Uint8Array(16)), nonce = crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce, additionalData: encoder.encode(JSON.stringify([TYPE, path])) }, await keyFor(password, salt), plain));
  entries[path] = encoder.encode(JSON.stringify({ version: 2, salt: base64(salt), nonce: base64(nonce), ciphertext: base64(ciphertext) }));
  const data = pack(entries);
  // Expanded output must satisfy the same limits used on the next restore.
  checkedZip(data);
  return { data, ...counts, filename: options.filename?.replace(/\.(xls|xlsb)$/i, '.xlsx') };
}
