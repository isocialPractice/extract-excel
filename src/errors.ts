/**
 * Custom error types for extract-excel.
 *
 * The draft documentation repeatedly references "a custom error" for specific
 * failure conditions (ambiguous sheet, malformed action arguments, etc.). Using
 * a dedicated error class lets the CLI present clean, actionable messages while
 * still allowing API consumers to `instanceof`-check programmatically.
 */

/** Stable, machine-readable error codes surfaced by the tool. */
export type ExtractErrorCode =
  | 'AMBIGUOUS_SHEET'
  | 'SHEET_NOT_FOUND'
  | 'WORKBOOK_NOT_FOUND'
  | 'UNREADABLE_WORKBOOK'
  | 'INVALID_CELL'
  | 'INVALID_RANGE'
  | 'TITLE_NOT_FOUND'
  | 'MALFORMED_ACTION'
  | 'NO_WORKBOOK'
  | 'INVALID_ARGUMENT'
  | 'TEST_FAILED'
  | 'PYTHON_NOT_FOUND';

/** Base error for all anticipated, user-facing failures. */
export class ExtractError extends Error {
  readonly code: ExtractErrorCode;
  /** Optional hint rendered under the message to guide the user. */
  readonly hint?: string;

  constructor(code: ExtractErrorCode, message: string, hint?: string) {
    super(message);
    this.name = 'ExtractError';
    this.code = code;
    this.hint = hint;
    // Restore prototype chain for instanceof across compiled targets.
    Object.setPrototypeOf(this, new.target.prototype);
  }
}

/** Thrown when a workbook has multiple sheets and none was selected. */
export class AmbiguousSheetError extends ExtractError {
  constructor(workbook: string, sheets: string[]) {
    super(
      'AMBIGUOUS_SHEET',
      `Workbook "${workbook}" contains ${sheets.length} sheets; a sheet must be selected.`,
      `Pass --sheet <name> immediately after the workbook. Available: ${sheets
        .map((s) => `"${s}"`)
        .join(', ')}`,
    );
    this.name = 'AmbiguousSheetError';
  }
}
