import { strFromU8, unzipSync } from 'fflate';
import { hasV2Protection, MAX_FILE_BYTES, transformOoxml } from './excel-ooxml';
import { transformWorkbook as transformV1 } from './excel-v1';
import { syntheticWorkbook } from './excel-synthetic';

export { MAX_FILE_BYTES };
export const EXCEL_EXTENSIONS = /\.(xlsx|xlsm|xltx|xltm|xls|xlsb|xlt|xml)$/i;

function isZip(input: Uint8Array): boolean {
  return input[0] === 0x50 && input[1] === 0x4b;
}

function packageInfo(input: Uint8Array): { legacy: boolean } {
  const parts = unzipSync(input, { filter(part) {
    return ['docProps/core.xml', 'xl/workbook.xml'].includes(part.name)
      && part.originalSize <= 1024 * 1024;
  } });
  const core = strFromU8(parts['docProps/core.xml'] ?? new Uint8Array());
  const workbook = strFromU8(parts['xl/workbook.xml'] ?? new Uint8Array());
  return {
    legacy: /<[^>]*keywords[^>]*>ExcelKey:workbook:v1<\//.test(core)
      && /<sheet\b[^>]*name="__ExcelKey_v1"[^>]*state="veryHidden"/.test(workbook),
  };
}

/** Generate a clean synthetic XLSX view with an encrypted, exact-original snapshot. */
export async function transformWorkbook(
  input: Uint8Array, password: string, mode: 'protect' | 'restore',
  options: { filename?: string } = {},
): Promise<{ data: Uint8Array; cells: number; sheets: number; filename?: string }> {
  if (mode !== 'protect' && mode !== 'restore') throw new Error('Choose protect or restore.');
  if (!password) throw new Error('Enter a password.');
  if (input.length > MAX_FILE_BYTES) throw new Error('Workbook exceeds the 64 MiB limit.');
  if (!input.length) throw new Error('Choose a valid Excel workbook.');

  if (isZip(input)) {
    // Validate and bound ZIP inflation before any secondary parser sees the package.
    const v2 = hasV2Protection(input);
    if (v2) return transformOoxml(input, password, mode, options);
    const info = packageInfo(input);
    if (info.legacy) {
      if (mode === 'protect') throw new Error('Workbook is already protected.');
      return transformV1(input, password, 'restore');
    }
    if (mode === 'restore') throw new Error('ExcelKey metadata is missing or damaged.');
  } else if (mode === 'restore') {
    throw new Error('Select the protected workbook downloaded from ExcelKey.');
  }

  let view: ReturnType<typeof syntheticWorkbook>;
  try {
    view = syntheticWorkbook(input, password);
  } catch (cause) {
    const reason = cause instanceof Error ? cause.message : 'Could not read the workbook.';
    throw new Error(`Could not read this Excel file: ${reason}`);
  }
  const transformed = await transformOoxml(view.data, password, 'protect', { original: input, filename: options.filename, syntheticCounts: { cells: view.cells, sheets: view.sheets } });
  return { ...transformed, filename: 'synthetic.xlsx' };
}
