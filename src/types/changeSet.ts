export type ChangeSetStatus = 'previewed' | 'applied' | 'reverted';

export type DomainDocumentType =
  | 'gstr2b'
  | 'form26as'
  | 'tally'
  | 'bank_statement'
  | 'workbook';

export interface SourceRef {
  documentType: DomainDocumentType;
  documentId: string;
  rowOrLine: string | number;
}

export interface DomainException {
  code: string;
  severity: 'flag' | 'block';
  message: string;
  affectedRows: number[];
}

export interface CellChange {
  cell: string;
  sheet: string;
  before: unknown;
  after: unknown;
  formula?: string;
  isHardcoded: boolean;
  sourceRefs?: SourceRef[];
  exceptionFlags?: DomainException[];
}

export interface ChangeSetSummary {
  changeSetId: string;
  conversationId: string;
  traceId: string;
  timestamp: string;
  prompt: string;
  changes: CellChange[];
  status: ChangeSetStatus;
  appliedAt?: string;
  revertedAt?: string;
  provenanceConfidence?: number;
}

export function formatCellValue(value: unknown): string {
  if (value === null || value === undefined || value === '') return '(empty)';
  return String(value);
}

export function hasExceptionFlags(change: CellChange): boolean {
  return Boolean(change.exceptionFlags && change.exceptionFlags.length > 0);
}

export function formatSourceRefLabel(ref: SourceRef): string {
  if (ref.documentType === 'workbook') {
    return String(ref.rowOrLine);
  }
  return `${ref.documentType} · ${ref.rowOrLine}`;
}
