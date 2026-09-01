export type SheetStructure = 'financial_model' | 'data_table' | 'report' | 'unknown';

/**
 * One entry per column, read from that column's first data row (TASKS.md #64) —
 * column-level granularity, not a true per-cell snapshot (matches the existing
 * `numberFormats` precedent, extended to font/fill). Feeds `ColumnMeta.format`
 * in `contextAdapter.ts`, which the backend broadcasts into `SheetContext.formats`
 * for the shadow workbook to capture and restore on revert.
 */
export interface ColumnFormatSnapshot {
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  fontColor?: string;
  fillColor?: string;
}

export interface SheetContext {
  name: string;
  usedRange: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: string[][];
  numberFormats: string[][];
  /** Absent for a minimal/fallback read (`buildMinimalWorkbookContext`) — treat as unknown, not "no formatting". */
  columnFormats?: ColumnFormatSnapshot[];
  structure: SheetStructure;
  headers: string[];
  /**
   * 0-based index into `values` where `headers` was actually found. Sheets
   * exported from other tools often have title/preamble rows above the real
   * header row — code must key off this instead of assuming headers are row 0.
   */
  headerRowIndex: number;
  formulaSummary: string;
  isHidden: boolean;
}

export interface NamedRangeInfo {
  name: string;
  formula: string;
  type: string;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  range: string;
  hasHeaders: boolean;
  columnNames: string[];
}

/**
 * A conditional-format rule already present on the live sheet, read back via
 * Office.js (TASKS.md #38) — not limited to rules Cellix itself applied. The
 * `id` is what lets a follow-up request ("change the threshold to 15%")
 * target this specific rule (`MODIFY_CONDITIONAL_FORMAT`) instead of
 * stacking a duplicate `CONDITIONAL_FORMAT` on top.
 */
export interface ConditionalFormatRuleInfo {
  id: string;
  sheetName: string;
  range: string;
  ruleKind: 'cellValue' | 'formula' | 'topBottom' | 'colorScale' | 'other';
  /** Human-readable summary for prompt context, e.g. "greaterThan 1000" or "top 5 items". */
  summary: string;
}

export interface DeepWorkbookContext {
  activeSheetName: string;
  selectedRange?: string;
  sheets: SheetContext[];
  namedRanges: NamedRangeInfo[];
  tables: TableInfo[];
  conditionalFormats: ConditionalFormatRuleInfo[];
  prompt_context: string;
}
