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

export interface SheetSnapshot {
  sheetName: string;
  usedRange: string;
  rowCount: number;
  colCount: number;
  headers: string[];
  sampleData: (string | number | null)[][];
  columnMeta: ColumnMeta[];
}

export interface WorkbookContext {
  sheets: SheetSnapshot[];
  activeSheet: string;
  selectedRange?: string;
  selectedValues?: (string | number | null)[][];
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
