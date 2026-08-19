export type SheetStructure = 'financial_model' | 'data_table' | 'report' | 'unknown';

export interface SheetContext {
  name: string;
  usedRange: string;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  formulas: string[][];
  numberFormats: string[][];
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

export interface DeepWorkbookContext {
  activeSheetName: string;
  selectedRange?: string;
  sheets: SheetContext[];
  namedRanges: NamedRangeInfo[];
  tables: TableInfo[];
  prompt_context: string;
}
