# Excel workbook compatibility

User amendment: remove feature-based rejection; accept Excel files regardless of their workbook contents, with row 1 readable and password restoration.

## Implementation

- Preserve original OOXML ZIP entries rather than round-tripping unknown features through ExcelJS. Replace populated data cells with short opaque tokens; leave row 1 unchanged.
- Store an authenticated, encrypted copy of the original file in a custom package part. Restoration returns the original bytes, so all workbook content survives exactly. Authenticate the protected package to reject tampering.
- Allow all ordinary package parts, relationships, formulas, validation, conditional formatting, named ranges, extensions, charts, images, tables, pivots and VBA. This protects worksheet cell values; opaque workbook features may independently retain copies of values.
- Support `.xlsx`, `.xlsm`, `.xltx`, `.xltm` directly. Parse `.xls` and `.xlsb` into a protected `.xlsx` view, retain the encrypted original bytes and restore their original format.
- Retain the existing v1 restore path for previously protected files.
- Update file picker, output naming, UI copy, README and build inputs.

## Verification

Test feature-rich OOXML entry preservation and exact original-byte restoration, shared strings without data-value leakage, headers including formulas, shared/array formulas, macros and opaque parts, notes/hyperlinks, long cells, wrong passwords, tampering, metadata deletion, existing v1 restoration, binary import/restore, build and browser upload/download.

Retain malformed-input, authentication and resource-exhaustion checks. No commits requested.
