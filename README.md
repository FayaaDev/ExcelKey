# ExcelKey

Create a clean synthetic Excel workbook entirely in the browser. The generated workbook keeps a reversible, encrypted copy of the exact original, which can be restored with the same passkey, its recovery phrase, or the legacy password used for older files. No account, upload, or server-side processing is required.

## Run locally

Requires [Bun](https://bun.sh/):

```sh
bun install
bun run build
bun run serve
```

Open `http://localhost:8787`. Use HTTPS outside localhost: Web Crypto and passkeys require a secure context.

## Generate and restore

1. Select **Generate**, choose an Excel workbook, then choose a protection method:
   - **Create passkey** creates an ExcelKey passkey and can display a 24-word recovery phrase.
   - **Use existing passkey** uses the original ExcelKey passkey.
   - **No password** makes a non-confidential workbook that anyone holding it can restore.
2. Download `synthetic.protected.xlsx`. Generation always creates an `.xlsx` file with this neutral name; it never overwrites the source file.
3. To restore, select **Restore**, choose the protected workbook, and select the matching method: passkey, recovery phrase, legacy password, or no password.
4. Download the original filename with a `.restored` suffix, such as `name.restored.xlsm`.

Keep the original and, when used, its passkey or recovery phrase. ExcelKey cannot recover a lost passkey, recovery phrase, or legacy password.

## What the synthetic workbook contains

- All worksheets are processed, including hidden sheets. Row 1 is treated as the header row.
- Neutral worksheet names, original headers, dates, layout, basic number/date formats, merged cells, and empty cells remain visible.
- Text, names, emails, IDs, numeric values, and booleans are replaced consistently within a run. Name detection is header-based; generic text becomes synthetic labels.
- Integers use 1-1,000, decimals use 0.01-1,000, and percentages use 0.01-1. Values sharing a field header map consistently across worksheets.
- Dates are identified by Excel date/time formats, date-like headers, and recognizable ISO date strings. The 1900/1904 date system is retained. A date formula without a cached value must be recalculated and saved before generation.
- Original headers, dates, repeated values, and relationships intentionally remain visible. Synthetic output is not a guarantee of anonymity.
- Charts, pivots, macros, metadata, images, hyperlinks, comments, styles beyond basic formats, and formulas are not carried into the visible workbook. The encrypted original retains them byte-for-byte.

Accepted inputs are `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xls`, `.xlsb`, `.xlt`, and XML workbooks. Source workbooks protected with an Excel password must be unlocked and saved before generation.

## Limits and safety

- Input and protected-workbook compressed size limit: 64 MiB.
- OOXML validation limits: 256 MiB inflated content and 10,000 ZIP entries.
- Keep generated workbooks intact. Editing or resaving one can make it impossible to restore. Modified protected content is rejected rather than partly restored.
- Files and encryption secrets remain in the browser. Passwords, recovery phrases, and passkey-derived keys are not uploaded or saved by ExcelKey.

## Protection formats

New workbooks support three formats:

- **Passkey / recovery phrase:** a WebAuthn PRF passkey derives a root key; the 24-word recovery phrase encodes the same root key. Either restores the original.
- **No password:** a random AES-256-GCM key is embedded in the workbook. It detects accidental modification but provides no confidentiality.
- **Legacy password:** older workbooks use PBKDF2-HMAC-SHA-256 with 600,000 iterations and AES-256-GCM. This option exists only for restoration.

Protected workbooks authenticate both the encrypted original and the visible synthetic package. Anyone with a protected workbook can make offline guesses against a legacy password, so use a strong one. Passkey protection requires a browser and authenticator that support the WebAuthn PRF extension.

## Development

```sh
bun test
bun run build
bun run serve
```

- `src/excel.ts`: workbook detection and generate/restore entry point.
- `src/excel-synthetic.ts`: visible synthetic workbook generation.
- `src/excel-ooxml.ts`: protected-workbook packaging, encryption, and integrity checks.
- `web/index.html` and `web/excel-app.ts`: workbook browser workflow.
- `test/`: unit and round-trip coverage.
- `e2e/`: deployed-app and browser interoperability checks; see `e2e/README.md`.

The repository also contains the reusable ExcelKey passkey encryption core in `src/` and its generic file UI in `web/app.ts`. Licensed under GPL-3.0-or-later; see `LICENSE`.
