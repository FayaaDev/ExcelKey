# ExcelKey Implementation Plan

**Goal:** Password-protect populated Excel cells except row 1, and restore through the browser UI.

**Architecture:** Isolate workbook/password processing in `reference/src/excel.ts`; keep existing generic ExcelKey crypto intact. A new `reference/web/excel-app.ts` binds a focused form in `reference/web/index.html` to that module. ExcelJS handles workbook values and styles; Web Crypto handles cryptography.

**Tech stack:** TypeScript, Bun, ExcelJS, Web Crypto, existing static HTML/CSS.

## Tasks

- [x] Add ExcelJS and implement `transformWorkbook(input: Uint8Array, password: string, mode: 'protect' | 'restore'): Promise<{data: Uint8Array, cells: number, sheets: number}>` in `reference/src/excel.ts`.
- [x] Add `reference/test/excel.test.ts`: round-trip typed cells/formulas, preserve row-1 headers on all sheets, skip blanks, reject wrong passwords, deleted/tampered ciphertext/metadata, malformed files and double protection. Run `bun test test/excel.test.ts`.
- [x] Implement file/password/confirmation/operation controls, busy/error/success/download states in `reference/web/index.html` and `reference/web/excel-app.ts`. Validate file extension, size, password presence and confirmation before reading. Clear passwords after operations; revoke superseded download URLs.
- [x] Include the new browser entry in `reference/package.json` build and service-worker shell; update manifest identity and README for ExcelKey.
- [x] Run `bun test`, `bunx tsc --noEmit`, and `bun run build:web`. Inspect desktop/mobile UI and exercise real upload/protect/download/restore, wrong-password and confirmation errors.

Verification: 113 tests pass. Build and strict targeted TypeScript check pass. Repository-wide typecheck reports pre-existing errors in `web/build-blog.ts`. Browser-downloaded round trip preserves values, dates, formula, number format and headers on two worksheets. Wrong passwords and mismatched confirmation show errors without output. No horizontal overflow at desktop/mobile sizes.

No commits: workspace is not a Git repository and commits were not requested.
