import { SheetAction } from '@/types/sheet-actions';
import { sheetKeyOf, SheetGuardStates } from '@/engine/sheetGuardState';

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

/** Cosmetic header paints — never treat as "overwrite headers with data". */
const HEADER_COSMETIC_TYPES = new Set<SheetAction['type']>([
  'FORMAT_RANGE',
  'HIGHLIGHT_CELL',
  'MERGE_CELLS',
]);

function isHeaderMutation(action: SheetAction, sheetIsEmpty = false): boolean {
  if (action.type === 'ADD_ROW' || action.type === 'WRITE_TABLE') return false;
  // Spec 24: bold/fill/highlight the header row must not be blocked as a "new row" write.
  if (HEADER_COSMETIC_TYPES.has(action.type)) return false;
  if (sheetIsEmpty) {
    const allowedOnHeader = new Set(['SET_CELL', 'SET_FORMULA', 'FORMAT_RANGE', 'MERGE_CELLS']);
    if (action.row === HEADER_ROW && allowedOnHeader.has(action.type)) return false;
  }
  if (action.type === 'DELETE_ROW') return action.row === HEADER_ROW;
  if (isConfirmedRangeClear(action)) return false;
  return action.row === HEADER_ROW;
}

function collectHeaderRowSetCells(actions: SheetAction[]): SheetAction[] {
  return actions.filter(
    (a) =>
      (a.type === 'SET_CELL' || a.type === 'SET_FORMULA' || a.type === 'CLEAR_CELL') &&
      a.row === HEADER_ROW,
  );
}

/**
 * Merge SET_CELL / SET_FORMULA on header row into a single ADD_ROW.
 *
 * `sheetName` is not optional bookkeeping — it is the whole reason this merge
 * is safe. A synthesized action that loses its target falls through
 * `resolveWorksheet`'s missing-name branch onto `getActiveWorksheet()`, which
 * silently writes it to whatever tab happens to be open. Callers must only
 * ever pass actions that share one sheet, and that sheet's name must come back
 * out on the merged action. See TASKS.md #137.
 */
function convertHeaderWritesToAddRow(
  headerActions: SheetAction[],
  sheetName: string | undefined,
): SheetAction | null {
  if (!headerActions.length) return null;

  const maxCol = Math.max(...headerActions.map((a) => a.col ?? 0), 0);
  const rowData: unknown[] = Array.from({ length: maxCol + 1 }, () => '');

  for (const action of headerActions) {
    if (action.col === undefined) continue;
    if (action.type === 'SET_FORMULA') rowData[action.col] = action.formula ?? '';
    else if (action.type === 'SET_CELL') rowData[action.col] = action.value ?? '';
    else if (action.type === 'CLEAR_CELL') rowData[action.col] = '';
  }

  const merged: SheetAction = { type: 'ADD_ROW', data: rowData as SheetAction['data'] };
  if (sheetName) merged.sheetName = sheetName;
  return merged;
}

/**
 * A whole-range CLEAR the server already confirmed as deliberate is not an
 * accidental header clobber.
 *
 * Client twin of the server fix in TASKS.md #179, and the consequence here was
 * worse than a silent drop: the blocked clear left `safe.length === 0`, which
 * made `useConversation` raise CLARIFY_ROW_PLACEMENT — so plain "clear the
 * sheet" was answered with "Where should the new row go?", a question about
 * ADDING a row asked because a request to REMOVE one had been thrown away.
 *
 * Lives in one place because `sanitizeActions` ORs `isHeaderMutation` with
 * `guardCellMutation`: exempting only one of them changes nothing at all.
 * TASKS.md #205.
 */
function isConfirmedRangeClear(action: SheetAction): boolean {
  const clearRangeTypes = new Set(['CLEAR_CONTENT', 'CLEAR_ALL', 'CLEAR_FORMAT']);
  return clearRangeTypes.has(action.type) && action.explicitOverwriteConfirmed === true;
}

function guardCellMutation(action: SheetAction, sheetIsEmpty = false): boolean {
  if (action.type === 'ADD_ROW' || action.type === 'WRITE_TABLE') return false;
  // Cosmetic format on header/body is always allowed (fill, bold, highlight).
  if (HEADER_COSMETIC_TYPES.has(action.type)) return false;
  if (isConfirmedRangeClear(action)) return false;
  if (action.row === undefined) return false;
  if (sheetIsEmpty && action.row === HEADER_ROW) {
    return !['SET_CELL', 'SET_FORMULA', 'FORMAT_RANGE', 'MERGE_CELLS'].includes(action.type);
  }
  return action.row <= HEADER_ROW;
}

export interface SanitizeOptions {
  /**
   * Live per-sheet facts from `probeSheetGuardStates`. When a sheet is present
   * here it is authoritative: `hasHeaderRow: false` means row 1 is free and a
   * header write into it is legitimate, not an overwrite.
   *
   * When absent (probe failed, or a caller with no Excel context), the sheet
   * falls back to `layout` — the pre-existing single-sheet behaviour.
   */
  sheetStates?: SheetGuardStates;
}

/**
 * Decide, for one sheet, whether row 1 is free to write.
 *
 * `layout` describes the *active* sheet only, so it may only speak for the
 * group that has no explicit sheet name. Letting it answer for named sheets is
 * how a batch got graded against the wrong sheet's headers.
 */
function resolveSheetIsEmpty(
  sheetKey: string,
  layout: SheetLayout | undefined,
  options: SanitizeOptions | undefined,
): boolean {
  const probed = options?.sheetStates?.get(sheetKey);
  if (probed) return !probed.hasHeaderRow;
  if (sheetKey === '') return layout?.isEmpty ?? false;
  // A named sheet with no probe result: the active sheet's layout says nothing
  // about it, so fall back to the conservative legacy answer.
  return layout?.isEmpty ?? false;
}

/**
 * Split a batch into per-sheet groups, preserving each action's original index
 * so the sanitized output keeps the batch's ordering.
 */
function groupBySheet(actions: SheetAction[]): Map<string, SheetAction[]> {
  const groups = new Map<string, SheetAction[]>();
  for (const action of actions) {
    const key = sheetKeyOf(action);
    const group = groups.get(key);
    if (group) group.push(action);
    else groups.set(key, [action]);
  }
  return groups;
}

export function sanitizeActions(
  actions: SheetAction[],
  layout?: SheetLayout,
  options?: SanitizeOptions,
): SanitizeResult {
  const warnings: string[] = [];
  const blocked: SheetAction[] = [];
  const safe: SheetAction[] = [];

  let convertedGroups = 0;
  let totalHeaderWrites = 0;

  // Header-row logic is per sheet. Row 1 of "January" and row 1 of "Main" are
  // unrelated cells; pooling them let a 121-cell, 13-sheet batch collapse into
  // one row on one tab. See TASKS.md #137.
  for (const [sheetKey, groupActions] of groupBySheet(actions)) {
    const sheetIsEmpty = resolveSheetIsEmpty(sheetKey, layout, options);
    const sheetName = groupActions.find((a) => a.sheetName)?.sheetName;

    const headerWrites = collectHeaderRowSetCells(groupActions);
    const withoutHeaderWrites = groupActions.filter((a) => !headerWrites.includes(a));
    totalHeaderWrites += headerWrites.length;

    let normalized: SheetAction[] = [...withoutHeaderWrites];

    if (headerWrites.length > 0 && !sheetIsEmpty) {
      const addRow = convertHeaderWritesToAddRow(headerWrites, sheetName);
      if (addRow) {
        normalized.unshift(addRow);
        convertedGroups += 1;
      }
    } else if (headerWrites.length > 0 && sheetIsEmpty) {
      normalized = [...headerWrites, ...withoutHeaderWrites];
    }

    for (const action of normalized) {
      if (isHeaderMutation(action, sheetIsEmpty) || guardCellMutation(action, sheetIsEmpty)) {
        blocked.push(action);
        continue;
      }
      safe.push(action);
    }
  }

  if (convertedGroups > 0) {
    warnings.push(
      'Converted header-row cell writes to ADD_ROW so data appends after existing rows.',
    );
  }

  if (blocked.length > 0) {
    warnings.push(`Blocked ${blocked.length} action(s) that would modify row ${HEADER_ROW} (headers).`);
  }

  const requiresClarification =
    safe.length === 0 && (blocked.length > 0 || totalHeaderWrites > 0);

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

/** True when sanitize blocked value-writes (not cosmetic header fill). Safe to ask row-placement. */
export function blockedActionsAreDataWrites(blocked: SheetAction[]): boolean {
  if (blocked.length === 0) return false;
  // Only actions that PLACE content qualify. The card this gates asks "Where
  // should the new row go?", which is meaningless for a removal — a blocked
  // CLEAR_*/DELETE_ROW has no new row to position, so it must fall through to
  // an honest failure message instead. TASKS.md #205.
  return blocked.some(
    (a) =>
      a.type === 'SET_CELL' ||
      a.type === 'SET_FORMULA' ||
      a.type === 'ADD_ROW' ||
      a.type === 'INSERT_ROW' ||
      a.type === 'WRITE_TABLE' ||
      a.type === 'BATCH_SET',
  );
}

/** User free-text rejecting the "new row" clarification and asking for header format instead. */
export function isHeaderFormatCorrectionMessage(message: string): boolean {
  const lower = message.toLowerCase();
  const rejectsNewRow =
    /\bnot\s+(a\s+)?new\s+row\b|\bno\s+new\s+row\b|\bexisting\s+row\b|\bnot\s+insert/i.test(
      lower,
    );
  const wantsHeaderFormat =
    /\b(header|headers)\b/.test(lower) &&
    /\b(bg|background|fill|color|colour|highlight|bold|green|red)\b/.test(lower);
  return (rejectsNewRow && wantsHeaderFormat) || wantsHeaderFormat;
}

export function buildActionRulesPrompt(layout?: SheetLayout): string {
  const nextRowHuman = layout ? layout.nextDataRow + 1 : 'last row + 1';
  const headerList = layout?.headers.filter(Boolean).join(', ') || '(see preview)';

  return `
SPREADSHEET ACTION RULES (0-indexed rows in JSON; row 0 = Excel row 1):
- Row ${HEADER_ROW} (Excel row 1) is the HEADER row: [${headerList}].
- NEVER use SET_CELL, CLEAR_CELL, SET_FORMULA, or DELETE_ROW on row ${HEADER_ROW} unless the user explicitly asks to rename headers.
- FORMAT_RANGE and HIGHLIGHT_CELL on row ${HEADER_ROW} ARE allowed when the user wants to style/bold/fill the headers (not add data).
- To ADD a new data row, ALWAYS use: {"type":"ADD_ROW","data":["col1","col2",...]} — do NOT set row index; it appends automatically at row ${nextRowHuman}.
- Do NOT write dummy/data values into row ${HEADER_ROW} when the user asks to "add a row".
- If placement, values, or target row are unclear for adding data, respond with a clarifying QUESTION only (no actions JSON). Ask: where to insert, what values, how many rows — never when the user only asked to format/color headers.
- Prefer ADD_ROW over multiple SET_CELL actions when adding a full row.

Example — add dummy row (correct):
\`\`\`json
{"actions":[{"type":"ADD_ROW","data":["Sample A","Sample B","100"]}],"explanation":"Added a new data row after existing rows"}
\`\`\`
`.trim();
}
