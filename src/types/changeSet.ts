export type ChangeSetStatus = 'previewed' | 'applied' | 'reverted';

export interface CellChange {
  cell: string;
  sheet: string;
  before: unknown;
  after: unknown;
  formula?: string;
  isHardcoded: boolean;
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
}

export function formatCellValue(value: unknown): string {
  if (value === null || value === undefined || value === '') return '(empty)';
  return String(value);
}
