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
