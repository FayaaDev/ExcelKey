# ExcelKey

Generate a clean synthetic Excel workbook locally in your browser, with an encrypted original embedded inside. **Original headers in row 1 and dates remain visible on every worksheet.** The same password restores the exact original file. This uses reversible encryption, not hashing.

## Run

Requires Bun:

```sh
cd reference
bun install
bun run build
bun run serve
```

Open `http://localhost:8787`. Deploy over HTTPS outside localhost; Web Crypto requires a secure context.

## Generate and restore

1. Choose **Generate synthetic workbook**, select an Excel workbook, and enter and confirm a password.
2. Download `synthetic.protected.xlsx`. Generation always outputs `.xlsx` with this neutral filename, regardless of the original filename or format. The original workbook is not overwritten.
3. Choose **Restore values**, select the generated workbook, and enter the original password.
4. Download the restored original using its original filename base and extension with a `.restored` suffix (for example, `name.restored.xlsm`).

The mode/API identifiers remain `protect` and `restore`.

Keep a copy of the original and keep your password safe. There is no password recovery. Passwords are not saved or uploaded. Processing happens entirely in the browser.

## Scope and limits

- All worksheets are processed, including hidden sheets. Row 1 is always the header row.
- The clean visible workbook uses neutral worksheet names. Original headers in row 1, dates, layout, basic number and date formats, and empty cells are retained.
- Other text, names, emails and IDs are replaced consistently. Generic text maps to synthetic labels; name detection is a header-based heuristic. Numbers use independent synthetic ranges rather than preserving original numeric values.
- Integers use 1–1,000; decimals use 0.01–1,000; percentages use 0.01–1. Matching field headers share replacement mappings across sheets. Mappings are fresh for each run.
- Excel date/time formats and date-column headers identify dates. Recognizable ISO date strings are retained too. Basic date formats and the 1900/1904 date system are preserved; custom formats containing literal text are normalized to avoid leaking that text. Date formulas without cached results must be recalculated and saved before generation.
- Original headers, dates, repetition and relationships intentionally remain visible. Synthetic output is not a guarantee of anonymity.
- Original charts, pivots, macros, metadata, images, hyperlinks, comments and formulas are not carried into the visible output. Date formulas are frozen to their cached date values. Original features are retained in the encrypted original, rather than rejected solely because they exist.
- Accepted inputs: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb`, `.xlt` and XML workbooks. Generation always produces `.xlsx`; v2 restoration returns the exact original bytes and format, including features omitted from the synthetic view.
- Current limits: 64 MiB compressed input/output, 256 MiB inflated OOXML content and 10,000 ZIP entries. The embedded encrypted original increases output size.
- Keep the generated workbook intact. Editing or resaving it in Excel can prevent restoration. Tamper detection rejects modified protected content instead of returning partial output; restore the original before making changes.

## Security format

The workbook transformations use browser Web Crypto: PBKDF2-HMAC-SHA-256 with 600,000 iterations and a random 16-byte workbook salt, followed by AES-256-GCM. The v2 format embeds the encrypted original and an authenticated manifest that verifies the generated workbook's integrity before releasing the original.

Missing protection metadata, wrong passwords and tampering fail restoration. ZIP size/part checks run before workbook parsing. Password strength matters because anyone holding a generated workbook can try password guesses offline.

### Restore compatibility

Restoration supports both legacy v1 protected workbooks and v2 workbooks. New generation uses v2. V2 restores the embedded original byte-for-byte; v1 restores its authenticated cell values through the legacy path and does not promise byte-identical recovery or an embedded original filename. V1's 20 MiB compressed and 100 MiB inflated limits and feature restrictions apply only to that legacy restore path, not current generation.

## Development

```sh
cd reference
bun test
bun run build
```

- `reference/src/excel.ts`: workbook/password transformations.
- `reference/web/excel-app.ts`: browser workflow.
- `reference/web/index.html`: accessible form and states.
- `reference/test/excel.test.ts`: round-trip, malformed-input and tampering tests.

The original generic ExcelKey cryptography remains available under `reference/src/`; its documentation is in `reference/FILEKEY.md`. The main UI now uses ExcelKey's password workflow, not ExcelKey passkeys or `.filekey` files. Licensed under GPL-3.0-or-later; see `LICENSE`.
