# ExcelKey password protection

Approved in conversation: protect every populated cell below row 1 on every worksheet; row 1 remains readable. Restore requires the original password in the UI. This is reversible encryption, not hashing.

Support `.xlsx` workbooks, processed locally. Use ExcelJS for workbook parsing/writing and Web Crypto PBKDF2-SHA-256 plus AES-256-GCM for password-based protection. Keep original cell types and formulas in authenticated encrypted payloads, with random workbook salt and unique nonces. Authenticate a manifest of protected sheets/cell addresses so deletion, replacement, or missing metadata cannot silently produce partial restoration. No output on incorrect password or corrupt input. Preserve ordinary styles and layout; disclose that advanced Excel features may not survive workbook rewriting.

UI: existing white/blue, self-hosted Inter visual language, single task form with Protect/Restore choices, `.xlsx` file input, password, protection-only confirmation, submit button, accessible status/error messages, and explicit download link. Preserve headers; do not display workbook data in the UI. No passkey required for this workflow. Do not persist passwords.

Verify values/types/formulas, header exclusion on multiple sheets, empty cells, wrong passwords, corrupted/removed cells, duplicate protection, malformed inputs, production build, and desktop/mobile browser flow.
