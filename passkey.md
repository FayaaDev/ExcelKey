Implementation plan
1. Replace password entry with passkey actions
- First use: Create passkey.
- Later: Protect with passkey or Restore with passkey.
- Reuse web/webauthn.ts to trigger Chrome’s native picker and require device verification.
- Show clear messages for cancellation or unsupported PRF.
2. Connect passkeys to workbook encryption
- Add a new, versioned workbook envelope.
- Reuse existing PRF-derived key handling.
- Preserve synthetic workbook output, integrity checks, and exact-original restoration.
- Retain password readers for existing protected files.
3. Add recovery backup
- Offer existing recovery-phrase functionality when setting up the passkey.
- Allow workbook restoration using that phrase if the passkey becomes unavailable.
- Explain that creating a replacement passkey does not unlock files protected with the original.
4. Verify
- Exact-byte restoration with the original passkey and recovery phrase.
- Wrong-passkey rejection and tamper detection.
- Existing password-file compatibility.
- Chrome enrollment, restoration, cancellation, and unsupported-provider behavior.
Recommended default: offer a recovery backup during setup, and keep encryption entirely local. No accounts or server storage needed.
