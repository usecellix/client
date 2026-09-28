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
  /**
   * Bold/italic/fontColor/fillColor read from the column's first data row
   * (TASKS.md #64) — column-level granularity, same as `numberFormat` above,
   * not a genuine per-cell snapshot. Feeds revert's format restoration.
   */
  format?: {
    bold?: boolean;
    italic?: boolean;
    fontColor?: string;
    fillColor?: string;
  };
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
  /**
   * Read from Office.js's real `worksheet.visibility` (DeepWorkbookContext's
   * own SheetContext already carries this — TASKS.md #257 found it was being
   * read correctly but silently dropped on the way into this lighter-weight
   * shape, so "how many sheets" could list a hidden sheet's name with no way
   * to say it was hidden). Optional so older cached/minimal contexts that
   * never populated it are "unknown", not "definitely visible".
   */
  isHidden?: boolean;
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

export interface ConditionalFormatRuleInfo {
  id: string;
  sheetName: string;
  range: string;
  ruleKind: 'cellValue' | 'formula' | 'topBottom' | 'colorScale' | 'other';
  summary: string;
}

export interface WorkbookContext {
  sheets: SheetSnapshot[];
  activeSheet: string;
  selectedRange?: string;
  selectedValues?: (string | number | null)[][];
  namedRanges?: NamedRangeInfo[];
  tables?: TableInfo[];
  conditionalFormats?: ConditionalFormatRuleInfo[];
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
