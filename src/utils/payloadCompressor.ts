import { WorkbookContext } from '@/types/cellix.types';
import { AssistantMode } from '@/types/mode';

/** Header row + first/last data rows sent to the API (metadata-first compression). */
export const HEADER_ROWS = 1;
export const FIRST_DATA_ROWS = 5;
export const LAST_DATA_ROWS = 5;

export function isFindLookupQuery(message: string): boolean {
  return /\b(find|search|locate|look up|lookup|show me|where is|get me|fetch|pull up|bring up)\b/i.test(
    message,
  );
}

export interface CompressedSheetPayload {
  sheetData: any[][];
  originalRowCount: number;
  compressedRowCount: number;
  truncated: boolean;
  onDemandFetchEnabled: boolean;
  includedRowIndices: number[];
}

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeNumericString(value: string): string | number {
  const trimmed = value.trim();
  if (!trimmed) return '';

  const withoutDrCr = trimmed.replace(/\s*(Dr|Cr)\s*$/i, '');
  const withoutCurrency = withoutDrCr.replace(/[\u20B9\u0024\u20AC\u00A3,\s]/g, '');
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

function pickSparseRowIndices(
  totalRows: number,
  firstDataRows = FIRST_DATA_ROWS,
  lastDataRows = LAST_DATA_ROWS,
): number[] {
  if (totalRows <= 0) return [];
  if (totalRows <= HEADER_ROWS + firstDataRows + lastDataRows) {
    return Array.from({ length: totalRows }, (_, index) => index);
  }

  const indices = new Set<number>([0]);
  for (let row = 1; row <= firstDataRows; row += 1) {
    indices.add(row);
  }
  for (let offset = 0; offset < lastDataRows; offset += 1) {
    indices.add(totalRows - 1 - offset);
  }
  return Array.from(indices).sort((a, b) => a - b);
}

export function compressSheetData(
  sheetData: any[][],
  options?: {
    firstDataRows?: number;
    lastDataRows?: number;
  },
): CompressedSheetPayload {
  const firstDataRows = options?.firstDataRows ?? FIRST_DATA_ROWS;
  const lastDataRows = options?.lastDataRows ?? LAST_DATA_ROWS;

  const normalizedRows = sheetData
    .map(normalizeRow)
    .filter((row) => row.some((cell) => !isBlankCell(cell)));

  const originalRowCount = sheetData.length;
  const includedRowIndices = pickSparseRowIndices(
    normalizedRows.length,
    firstDataRows,
    lastDataRows,
  );
  const sparseRows = includedRowIndices.map((index) => normalizedRows[index] ?? []);
  const sheetDataSlice = compactRows(sparseRows);
  const truncated = normalizedRows.length > sheetDataSlice.length;

  return {
    sheetData: sheetDataSlice,
    originalRowCount,
    compressedRowCount: sheetDataSlice.length,
    truncated,
    onDemandFetchEnabled: truncated,
    includedRowIndices,
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
      onDemandFetchEnabled: compressed.onDemandFetchEnabled,
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

function attachCompressionMeta(
  workbookContext: WorkbookContext | WorkbookContextPayload | undefined,
  compressed: CompressedSheetPayload,
  activeSheetName?: string,
): WorkbookContext | WorkbookContextPayload | undefined {
  if (!workbookContext || !('sheets' in workbookContext)) {
    return workbookContext;
  }

  if (!Array.isArray(workbookContext.sheets) || workbookContext.sheets.length === 0) {
    return workbookContext;
  }

  const firstSheet = workbookContext.sheets[0];
  if (typeof firstSheet === 'string') {
    return workbookContext;
  }

  const activeSheet = activeSheetName ?? workbookContext.activeSheet;
  const sheets = workbookContext.sheets.map((sheet) => {
    if (typeof sheet === 'string') return sheet;
    if (sheet.sheetName !== activeSheet) return sheet;

    return {
      ...sheet,
      rowCount: Math.max(sheet.rowCount, compressed.originalRowCount),
      compressionMeta: {
        originalRowCount: compressed.originalRowCount,
        compressedRowCount: compressed.compressedRowCount,
        truncated: compressed.truncated,
        onDemandFetchEnabled: compressed.onDemandFetchEnabled,
        includedRowIndices: compressed.includedRowIndices,
      },
    };
  });

  return {
    ...workbookContext,
    sheets,
  } as WorkbookContext;
}

export function prepareConversationRequestPayload(
  message: string,
  sheetData: any[][],
  options?: {
    conversationId?: string | null;
    workbookId?: string | null;
    previousMessages?: ConversationHistoryMessage[];
    workbookContext?: WorkbookContext | WorkbookContextPayload;
    promptContext?: string;
    previewEnabled?: boolean;
    refinementChangeSetId?: string | null;
    mode?: AssistantMode;
  },
): {
  conversationId?: string;
  workbookId?: string;
  message: string;
  sheetData: any[][];
  sheetCompression?: Omit<CompressedSheetPayload, 'sheetData'>;
  workbookContext?: WorkbookContext | WorkbookContextPayload;
  promptContext?: string;
  previewEnabled?: boolean;
  refinementChangeSetId?: string;
  mode?: AssistantMode;
  conversationHistory: Array<{ role: 'user' | 'assistant'; content: string }>;
  context: {
    previousMessages: ConversationHistoryMessage[];
  };
} {
  const isQuickEdit = Boolean(options?.refinementChangeSetId);
  // Ask mode and find/lookup queries need the full sheet (no aggressive
  // truncation) so cross-sheet search can scan everything via on-demand fetch.
  const skipCompression = isFindLookupQuery(message) || options?.mode === 'ask';
  const compressed =
    isQuickEdit || skipCompression
      ? {
          sheetData,
          originalRowCount: sheetData.length,
          compressedRowCount: sheetData.length,
          truncated: false,
          onDemandFetchEnabled: true,
          includedRowIndices: sheetData.map((_, index) => index),
        }
      : compressSheetData(sheetData);
  const workbookContext = attachCompressionMeta(
    options?.workbookContext,
    compressed,
    options?.workbookContext && 'activeSheet' in options.workbookContext
      ? options.workbookContext.activeSheet
      : undefined,
  );

  if ((import.meta as any)?.env?.DEV) {
    // eslint-disable-next-line no-console
    console.debug('[Cellix] Prepared compressed conversation payload', {
      originalRowCount: compressed.originalRowCount,
      compressedRowCount: compressed.compressedRowCount,
      truncated: compressed.truncated,
      onDemandFetchEnabled: compressed.onDemandFetchEnabled,
      conversationId: options?.conversationId,
      previousMessages: options?.previousMessages?.length ?? 0,
      workbookContext,
    });
  }

  const previousMessages = options?.previousMessages ?? [];
  const conversationHistory = previousMessages.map((entry) => ({
    role: entry.role,
    content: entry.content,
  }));

  const promptContext =
    options?.promptContext ??
    (workbookContext && 'prompt_context' in workbookContext
      ? workbookContext.prompt_context
      : undefined);

  return {
    ...(options?.conversationId ? { conversationId: options.conversationId } : {}),
    ...(options?.workbookId ? { workbookId: options.workbookId } : {}),
    message,
    sheetData: compressed.sheetData,
    sheetCompression: {
      originalRowCount: compressed.originalRowCount,
      compressedRowCount: compressed.compressedRowCount,
      truncated: compressed.truncated,
      onDemandFetchEnabled: compressed.onDemandFetchEnabled,
      includedRowIndices: compressed.includedRowIndices,
    },
    ...(workbookContext ? { workbookContext } : {}),
    ...(promptContext ? { promptContext } : {}),
    ...(options?.previewEnabled !== undefined ? { previewEnabled: options.previewEnabled } : {}),
    ...(options?.refinementChangeSetId
      ? { refinementChangeSetId: options.refinementChangeSetId }
      : {}),
    ...(options?.mode ? { mode: options.mode } : {}),
    conversationHistory,
    context: {
      previousMessages,
    },
  };
}
