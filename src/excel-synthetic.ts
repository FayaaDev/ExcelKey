import { read, write, utils, SSF } from 'xlsx';
import type { CellObject, WorkSheet } from 'xlsx';
type Value = NonNullable<CellObject['v']>;

const FIRST = ['Alex', 'Maya', 'Noah', 'Leila', 'Oliver', 'Nora', 'Ethan', 'Sara', 'Liam', 'Zoe', 'Adam', 'Ava'];
const LAST = ['Morgan', 'Reed', 'Bennett', 'Parker', 'Hayes', 'Brooks', 'Foster', 'Ellis', 'Taylor', 'Stone', 'Rivera', 'Lane'];
const FORMATS = new Set(Object.values(SSF.get_table()));
function random(limit: number): number {
  const ceiling = Math.floor(0x100000000 / limit) * limit;
  let value: number;
  do { value = crypto.getRandomValues(new Uint32Array(1))[0]!; } while (value >= ceiling);
  return value % limit;
}
function field(header: string): string {
  const text = header.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().replace(/[_-]/g, ' ').trim();
  // ponytail: header heuristics cannot identify every domain field; add explicit column overrides if needed.
  if (/\b(email|e mail)\b/.test(text)) return 'email';
  if (/\b(id|identifier|code|key)\b/.test(text)) return `id:${text}`;
  if (/\b(first|given)\s*name\b/.test(text)) return 'first-name';
  if (/\b(last|family|sur)\s*name\b/.test(text)) return 'last-name';
  if (/\b(name|customer|employee|person|contact|owner)\b/.test(text) && !/\b(company|product|category|department|team|id|code)\b/.test(text)) return 'name';
  if (/\b(date|time|timestamp|dob|birthday)\b|\b(created|updated|modified|joined|registered)\s+(at|on)\b/.test(text)) return 'date';
  return `text:${text}`;
}
function dateText(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  return /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/.test(value);
}

/** Copy only intentional data into a fresh package; never carry source feature parts. */
export function syntheticWorkbook(input: Uint8Array, password: string): { data: Uint8Array; cells: number; sheets: number } {
  const source = read(input, { type: 'array', cellDates: false, cellNF: true, sheetStubs: true, password });
  if (!source.SheetNames.length) throw new Error('Workbook has no worksheets.');
  const output = utils.book_new();
  // SheetJS otherwise copies its global format table, including source literals.
  (output as typeof output & { SSF: Record<number, string> }).SSF = { 0: 'General' };
  output.Workbook = { WBProps: { date1904: !!source.Workbook?.WBProps?.date1904 } };
  const replacements = new Map<string, Value>();
  const used = new Set<string>();
  let cells = 0;
  function mapped(scope: string, cell: CellObject, make: () => Value): Value {
    const key = JSON.stringify([scope, cell.t, cell.v]);
    if (!replacements.has(key)) {
      let value = make();
      if (typeof value === 'string') {
        const base = value;
        let suffix = 2;
        while (used.has(value) || value === cell.v) value = `${base} ${suffix++}`;
        used.add(value);
      }
      replacements.set(key, value);
    }
    return replacements.get(key)!;
  }
  for (const [index, name] of source.SheetNames.entries()) {
    const sheet = source.Sheets[name]!;
    const clean: WorkSheet = {};
    const headers = new Map<number, string>();
    for (const address of Object.keys(sheet)) {
      if (address.startsWith('!')) continue;
      const position = utils.decode_cell(address);
      if (position.r === 0) headers.set(position.c, String(sheet[address].v ?? ''));
    }
    for (const address of Object.keys(sheet)) {
      if (address.startsWith('!')) continue;
      const cell: CellObject = sheet[address];
      const position = utils.decode_cell(address);
      const header = headers.get(position.c) ?? `column ${position.c + 1}`;
      const kind = field(header);
      const numberFormat = typeof cell.z === 'string' ? cell.z : undefined;
      if (cell.t === 'z' || cell.v === undefined || cell.v === null) {
        if (cell.f) {
          if (kind === 'date' || numberFormat && SSF.is_date(numberFormat)) throw new Error('A date formula has no cached value. Recalculate and save the original workbook before generating synthetic data.');
          clean[address] = { t: 'n', v: random(1000) + 1 }; cells++;
        }
        continue;
      }
      if (cell.v === '') { clean[address] = { t: 's', v: '' }; continue; }
      const nativeDate = cell.t === 'd' || cell.t === 'n' && !!numberFormat && SSF.is_date(numberFormat);
      const date = nativeDate || kind === 'date' || dateText(cell.v);
      let value = cell.v, type = cell.t;
      let format = numberFormat && (FORMATS.has(numberFormat) || date && /^(?:[ymdhs]+|AM\/PM|A\/P|[\s0:\/.,-]|\[[hms]+\])+$/i.test(numberFormat)) ? numberFormat : undefined;
      if (nativeDate && cell.t === 'n' && !format) format = 'yyyy-mm-dd hh:mm:ss';
      if (position.r > 0 && !date) {
        cells++;
        if (kind.startsWith('id:')) {
          type = 's'; format = undefined;
          value = mapped(kind, cell, () => `ID-${crypto.randomUUID()}`);
        } else if (cell.t === 'n') {
          const percent = numberFormat?.includes('%');
          const integer = Number.isInteger(cell.v) && !percent;
          format = percent ? '0.00%' : integer ? '0' : '0.00';
          value = mapped(`number:${header.toLowerCase()}:${format}`, cell, () => {
            let number = integer ? random(1000) + 1 : (random(percent ? 100 : 100000) + 1) / 100;
            if (number === cell.v) {
              if (integer) number = number % 1000 + 1;
              else if (number === (percent ? 1 : 1000)) number = 0.01;
              else number += 0.01;
            }
            return number;
          });
        } else if (cell.t === 'b') {
          value = mapped(`boolean:${header.toLowerCase()}`, cell, () => !!random(2));
        } else {
          type = 's'; format = undefined;
          value = mapped(kind, cell, () => {
            if (kind === 'email' || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(cell.v))) return `user-${crypto.randomUUID()}@example.com`;
            if (kind === 'first-name') return FIRST[random(FIRST.length)]!;
            if (kind === 'last-name') return LAST[random(LAST.length)]!;
            if (kind === 'name') return `${FIRST[random(FIRST.length)]} ${LAST[random(LAST.length)]}`;
            return `Label-${crypto.randomUUID()}`;
          });
        }
      }
      // No cached display text, formulas, rich text, links, comments or source styles.
      clean[address] = { t: type, v: value, ...(format ? { z: format } : {}) };
    }
    if (sheet['!ref']) clean['!ref'] = sheet['!ref'];
    if (sheet['!merges']) clean['!merges'] = sheet['!merges'].map(range => ({ s: { ...range.s }, e: { ...range.e } }));
    utils.book_append_sheet(output, clean, `Worksheet ${index + 1}`);
  }
  return { data: new Uint8Array(write(output, { type: 'array', bookType: 'xlsx', compression: true })), cells, sheets: source.SheetNames.length };
}
