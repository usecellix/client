export interface ClarificationPayload {
  question: string;
  suggestions?: string[];
  ambiguityScore: number;
}

export interface ConversationHistoryEntry {
  role: 'user' | 'assistant';
  content: string;
}

export interface ColumnMeta {
  index: number;
  header?: string;
  sampleValues: (string | number | null)[];
  detectedType: 'date' | 'number' | 'currency' | 'text' | 'boolean' | 'unknown';
  numberFormat?: string;
}

export type SheetStructure = 'financial_model' | 'data_table' | 'report' | 'unknown';

export interface SheetCompressionMeta {
  originalRowCount: number;
  compressedRowCount: number;
  truncated: boolean;
  onDemandFetchEnabled: boolean;
  includedRowIndices?: number[];
}

export interface SheetSnapshot {
  sheetName: string;
  usedRange: string;
  rowCount: number;
  colCount: number;
  headers: string[];
  /** 0-based row index within the original sheet where `headers` was detected. */
  headerRowIndex?: number;
  sampleData: (string | number | null)[][];
  columnMeta: ColumnMeta[];
  structure?: SheetStructure;
  formulaSummary?: string;
  compressionMeta?: SheetCompressionMeta;
}

export interface NamedRangeInfo {
  name: string;
  formula: string;
  type?: string;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  range?: string;
  hasHeaders?: boolean;
  columnNames: string[];
}

export interface WorkbookContext {
  sheets: SheetSnapshot[];
  activeSheet: string;
  selectedRange?: string;
  selectedValues?: (string | number | null)[][];
  namedRanges?: NamedRangeInfo[];
  tables?: TableInfo[];
  prompt_context?: string;
}

export interface DiffRow {
  rowIndex: number;
  type: 'added' | 'removed' | 'modified' | 'unchanged';
  label: string;
}

export interface CompareResult {
  sheetA: string;
  sheetB: string;
  summary: string;
  differences: DiffRow[];
  addedInB: string[];
  removedInB: string[];
  modifiedCells: { address: string; valueA: string; valueB: string }[];
}
