import { SheetAction } from '@/types/sheet-actions';

export const HEADER_ROW = 0;

export interface SheetLayout {
  headerRow: number;
  /** 0-indexed row where ADD_ROW will append (after last data row). */
  nextDataRow: number;
  dataRowCount: number;
  columnCount: number;
  /** First row values treated as column headers. */
  headers: string[];
  /** True when the sheet has no headers and no data rows. */
  isEmpty: boolean;
}

export interface SanitizeResult {
  actions: SheetAction[];
  blocked: SheetAction[];
  warnings: string[];
  /** True when every action was blocked or converted — caller should ask the user. */
  requiresClarification: boolean;
}

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

export function computeSheetLayout(sheetData: unknown[][]): SheetLayout {
  let lastNonEmptyRow = -1;
  let columnCount = 0;

  sheetData.forEach((row, rowIndex) => {
    if (!Array.isArray(row)) return;
    const hasContent = row.some((cell) => !isBlankCell(cell));
    if (hasContent) lastNonEmptyRow = rowIndex;
    row.forEach((cell, colIndex) => {
      if (!isBlankCell(cell)) columnCount = Math.max(columnCount, colIndex + 1);
    });
  });

  const headers = (sheetData[HEADER_ROW] ?? []).map((cell) =>
    cell === null || cell === undefined ? '' : String(cell),
  );

  const hasHeader = headers.some((h) => h.trim() !== '');
  const isEmpty = lastNonEmptyRow < 0 && !hasHeader;
  const nextDataRow =
    lastNonEmptyRow >= HEADER_ROW ? lastNonEmptyRow + 1 : hasHeader ? HEADER_ROW + 1 : HEADER_ROW;

  return {
    headerRow: HEADER_ROW,
    nextDataRow,
    dataRowCount: Math.max(0, lastNonEmptyRow > HEADER_ROW ? lastNonEmptyRow : 0),
    columnCount: Math.max(columnCount, 1),
    headers,
    isEmpty,
  };
}

function isHeaderMutation(action: SheetAction, sheetIsEmpty = false): boolean {
  if (action.type === 'ADD_ROW' || action.type === 'WRITE_TABLE') return false;
  if (sheetIsEmpty) {
    const allowedOnHeader = new Set(['SET_CELL', 'SET_FORMULA', 'FORMAT_RANGE', 'MERGE_CELLS']);
    if (action.row === HEADER_ROW && allowedOnHeader.has(action.type)) return false;
  }
  if (action.type === 'DELETE_ROW') return action.row === HEADER_ROW;
  return action.row === HEADER_ROW;
}

function collectHeaderRowSetCells(actions: SheetAction[]): SheetAction[] {
  return actions.filter(
    (a) =>
      (a.type === 'SET_CELL' || a.type === 'SET_FORMULA' || a.type === 'CLEAR_CELL') &&
      a.row === HEADER_ROW,
  );
}

/** Merge SET_CELL / SET_FORMULA on header row into a single ADD_ROW. */
function convertHeaderWritesToAddRow(headerActions: SheetAction[]): SheetAction | null {
  if (!headerActions.length) return null;

  const maxCol = Math.max(...headerActions.map((a) => a.col ?? 0), 0);
  const rowData: unknown[] = Array.from({ length: maxCol + 1 }, () => '');

  for (const action of headerActions) {
    if (action.col === undefined) continue;
    if (action.type === 'SET_FORMULA') rowData[action.col] = action.formula ?? '';
    else if (action.type === 'SET_CELL') rowData[action.col] = action.value ?? '';
    else if (action.type === 'CLEAR_CELL') rowData[action.col] = '';
  }

  return { type: 'ADD_ROW', data: rowData as SheetAction['data'] };
}

function guardCellMutation(action: SheetAction, sheetIsEmpty = false): boolean {
  if (action.type === 'ADD_ROW' || action.type === 'WRITE_TABLE') return false;
  if (action.row === undefined) return false;
  if (sheetIsEmpty && action.row === HEADER_ROW) {
    return !['SET_CELL', 'SET_FORMULA', 'FORMAT_RANGE', 'MERGE_CELLS'].includes(action.type);
  }
  return action.row <= HEADER_ROW;
}

export function sanitizeActions(
  actions: SheetAction[],
  layout?: SheetLayout,
): SanitizeResult {
  const warnings: string[] = [];
  const blocked: SheetAction[] = [];
  const sheetIsEmpty = layout?.isEmpty ?? false;

  const headerWrites = collectHeaderRowSetCells(actions);
  const withoutHeaderWrites = actions.filter((a) => !headerWrites.includes(a));

  let normalized: SheetAction[] = [...withoutHeaderWrites];

  if (headerWrites.length > 0 && !sheetIsEmpty) {
    const addRow = convertHeaderWritesToAddRow(headerWrites);
    if (addRow) {
      normalized.unshift(addRow);
      warnings.push(
        'Converted header-row cell writes to ADD_ROW so data appends after existing rows.',
      );
    }
  } else if (headerWrites.length > 0 && sheetIsEmpty) {
    normalized = [...headerWrites, ...withoutHeaderWrites];
  }

  const safe: SheetAction[] = [];

  for (const action of normalized) {
    if (isHeaderMutation(action, sheetIsEmpty) || guardCellMutation(action, sheetIsEmpty)) {
      blocked.push(action);
      continue;
    }
    safe.push(action);
  }

  if (blocked.length > 0) {
    warnings.push(`Blocked ${blocked.length} action(s) that would modify row ${HEADER_ROW} (headers).`);
  }

  const requiresClarification =
    safe.length === 0 && (blocked.length > 0 || headerWrites.length > 0);

  if (requiresClarification && layout) {
    warnings.push(
      `Sheet has headers in row ${layout.headerRow + 1}; next available data row is row ${layout.nextDataRow + 1}.`,
    );
  }

  return { actions: safe, blocked, warnings, requiresClarification };
}

export const CLARIFY_ROW_PLACEMENT = {
  question:
    'Where should the new row go? I will not overwrite your header row (row 1).',
  options: [
    'After the last row (recommended)',
    'At a specific row number',
    'Replace an existing data row',
  ],
} as const;

export function buildActionRulesPrompt(layout?: SheetLayout): string {
  const nextRowHuman = layout ? layout.nextDataRow + 1 : 'last row + 1';
  const headerList = layout?.headers.filter(Boolean).join(', ') || '(see preview)';

  return `
SPREADSHEET ACTION RULES (0-indexed rows in JSON; row 0 = Excel row 1):
- Row ${HEADER_ROW} (Excel row 1) is the HEADER row: [${headerList}]. NEVER use SET_CELL, CLEAR_CELL, SET_FORMULA, or DELETE_ROW on row ${HEADER_ROW} unless the user explicitly asks to rename headers.
- To ADD a new data row, ALWAYS use: {"type":"ADD_ROW","data":["col1","col2",...]} — do NOT set row index; it appends automatically at row ${nextRowHuman}.
- Do NOT write dummy/data values into row ${HEADER_ROW} when the user asks to "add a row".
- If placement, values, or target row are unclear, respond with a clarifying QUESTION only (no actions JSON). Ask: where to insert, what values, how many rows.
- Prefer ADD_ROW over multiple SET_CELL actions when adding a full row.

Example — add dummy row (correct):
\`\`\`json
{"actions":[{"type":"ADD_ROW","data":["Sample A","Sample B","100"]}],"explanation":"Added a new data row after existing rows"}
\`\`\`
`.trim();
}
