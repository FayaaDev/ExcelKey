// Shared error type for the crypto core. Kept in its own dependency-free leaf module so any
// module — including low-level ones like bytes.ts — can throw a coded ExcelKeyError without
// creating an import cycle. Re-exported from namespace.ts for backward compatibility.
export class ExcelKeyError extends Error {
  constructor(
    message: string,
    readonly code: string,
  ) {
    super(message);
    this.name = "ExcelKeyError";
  }
}
