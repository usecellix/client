import { WorkbookContext } from '@/types/cellix.types';

const MAX_ROWS = 50;

export interface CompressedSheetPayload {
  sheetData: any[][];
  originalRowCount: number;
  compressedRowCount: number;
  truncated: boolean;
}

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeNumericString(value: string): string | number {
  const trimmed = value.trim();
  if (!trimmed) return '';

  const withoutCurrency = trimmed.replace(/[\u20B9\u0024\u20AC\u00A3,\s]/g, '');
  const isWrappedNegative = /^\(.+\)$/.test(withoutCurrency);
  const numericCandidate = isWrappedNegative
    ? `-${withoutCurrency.slice(1, -1)}`
    : withoutCurrency;

  if (/^-?\d+(\.\d+)?%?$/.test(numericCandidate)) {
    const hasPercent = numericCandidate.endsWith('%');
    const parsed = Number(hasPercent ? numericCandidate.slice(0, -1) : numericCandidate);
    if (Number.isFinite(parsed)) {
      return hasPercent ? parsed / 100 : parsed;
    }
  }

  return trimmed;
}

function normalizeCell(value: unknown): any {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return normalizeNumericString(value);
  if (value instanceof Date) return value.toISOString();
  return value;
}

function normalizeRow(row: unknown): any[] {
  if (!Array.isArray(row)) return [];
  return row.map(normalizeCell);
}

function getLastNonEmptyColumn(rows: any[][]): number {
  return rows.reduce((maxColumn, row) => {
    for (let index = row.length - 1; index >= 0; index -= 1) {
      if (!isBlankCell(row[index])) {
        return Math.max(maxColumn, index + 1);
      }
    }
    return maxColumn;
  }, 0);
}

function compactRows(rows: any[][]): any[][] {
  const columnCount = getLastNonEmptyColumn(rows);
  return rows.map((row) =>
    Array.from({ length: columnCount }, (_, index) => (index < row.length ? row[index] : '')),
  );
}

export function compressSheetData(sheetData: any[][], maxRows = MAX_ROWS): CompressedSheetPayload {
  const normalizedRows = sheetData
    .map(normalizeRow)
    .filter((row) => row.some((cell) => !isBlankCell(cell)));

  const sheetDataSlice = compactRows(normalizedRows.slice(0, maxRows));

  return {
    sheetData: sheetDataSlice,
    originalRowCount: sheetData.length,
    compressedRowCount: sheetDataSlice.length,
    truncated: normalizedRows.length > maxRows,
  };
}

export function prepareStreamRequestPayload(prompt: string, sheetData: any[][]): {
  prompt: string;
  sheetData: any[][];
} {
  const compressed = compressSheetData(sheetData);

  if ((import.meta as any)?.env?.DEV) {
    // eslint-disable-next-line no-console
    console.debug('[Cellix] Prepared compressed sheet payload', {
      originalRowCount: compressed.originalRowCount,
      compressedRowCount: compressed.compressedRowCount,
      truncated: compressed.truncated,
    });
  }

  return {
    prompt,
    sheetData: compressed.sheetData,
  };
}

export interface ConversationHistoryMessage {
  role: 'user' | 'assistant';
  content: string;
  timestamp: string;
  type?: 'question' | 'answer' | 'command' | 'clarification';
}

export interface SheetLayoutPayload {
  headerRow: number;
  nextDataRow: number;
  dataRowCount: number;
  columnCount: number;
  headers: string[];
  isEmpty: boolean;
}

/** Compute sheet layout locally (not sent to API — backend validates a strict DTO). */
export function computeSheetLayout(sheetData: unknown[][]): SheetLayoutPayload {
  let lastNonEmptyRow = -1;
  let columnCount = 0;

  sheetData.forEach((row, rowIndex) => {
    if (!Array.isArray(row)) return;
    if (row.some((cell) => !isBlankCell(cell))) lastNonEmptyRow = rowIndex;
    row.forEach((cell, colIndex) => {
      if (!isBlankCell(cell)) columnCount = Math.max(columnCount, colIndex + 1);
    });
  });

  const headers = (sheetData[0] ?? []).map((cell) =>
    cell === null || cell === undefined ? '' : String(cell),
  );
  const hasHeader = headers.some((h) => h.trim() !== '');
  const isEmpty = lastNonEmptyRow < 0 && !hasHeader;
  const nextDataRow =
    lastNonEmptyRow >= 0 ? lastNonEmptyRow + 1 : hasHeader ? 1 : 0;

  return {
    headerRow: 0,
    nextDataRow,
    dataRowCount: Math.max(0, lastNonEmptyRow > 0 ? lastNonEmptyRow : 0),
    columnCount: Math.max(columnCount, 1),
    headers,
    isEmpty,
  };
}

/** @deprecated Use WorkbookContext from cellix.types */
export interface WorkbookContextPayload {
  activeSheet: string;
  sheets: string[];
}

export function prepareConversationRequestPayload(
  message: string,
  sheetData: any[][],
  options?: {
    conversationId?: string | null;
    previousMessages?: ConversationHistoryMessage[];
    workbookContext?: WorkbookContext | WorkbookContextPayload;
    previewEnabled?: boolean;
  },
): {
  conversationId?: string;
  message: string;
  sheetData: any[][];
  workbookContext?: WorkbookContext | WorkbookContextPayload;
  previewEnabled?: boolean;
  conversationHistory: Array<{ role: 'user' | 'assistant'; content: string }>;
  context: {
    previousMessages: ConversationHistoryMessage[];
  };
} {
  const compressed = compressSheetData(sheetData);

  if ((import.meta as any)?.env?.DEV) {
    // eslint-disable-next-line no-console
    console.debug('[Cellix] Prepared compressed conversation payload', {
      originalRowCount: compressed.originalRowCount,
      compressedRowCount: compressed.compressedRowCount,
      truncated: compressed.truncated,
      conversationId: options?.conversationId,
      previousMessages: options?.previousMessages?.length ?? 0,
      workbookContext: options?.workbookContext,
    });
  }

  const previousMessages = options?.previousMessages ?? [];
  const conversationHistory = previousMessages.map((entry) => ({
    role: entry.role,
    content: entry.content,
  }));

  return {
    ...(options?.conversationId ? { conversationId: options.conversationId } : {}),
    message,
    sheetData: compressed.sheetData,
    ...(options?.workbookContext ? { workbookContext: options.workbookContext } : {}),
    ...(options?.previewEnabled !== undefined ? { previewEnabled: options.previewEnabled } : {}),
    conversationHistory,
    context: {
      previousMessages,
    },
  };
}
