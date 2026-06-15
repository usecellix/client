import { CellValue, RichAction } from '@/action.types';
import { SheetAction } from '@/types/sheet-actions';
import { columnLetterToIndex, parseCellAddress, parseRangeAddress } from './addressUtils';
import { convertLegacyToRich } from './legacyConverter';

/** Set to true to route unconverted actions through the legacy engine (migration fallback). */
export const USE_LEGACY_ACTION_ENGINE =
  typeof localStorage !== 'undefined' &&
  localStorage.getItem('CELLIX_USE_LEGACY_ENGINE') === 'true';

function isRichAction(action: SheetAction): boolean {
  const richOnly = new Set([
    'BATCH_SET',
    'CREATE_TABLE',
    'DEFINE_NAMED_RANGE',
    'AUTOFIT_COLUMNS',
    'CLARIFY',
    'CHECKPOINT',
    'ADD_SHEET',
    'DELETE_SHEET',
    'SORT_RANGE',
  ]);
  if (richOnly.has(action.type)) return true;

  const record = action as unknown as Record<string, unknown>;
  if (action.type === 'ADD_ROW' && typeof record.afterRow === 'number' && Array.isArray(record.values)) {
    return true;
  }
  if (
    action.type === 'DELETE_ROW' &&
    (Array.isArray(record.rowNumbers) ||
      (Array.isArray(record.rows) &&
        (record.rows as unknown[]).every((n) => typeof n === 'number')))
  ) {
    return true;
  }
  if (action.type === 'INSERT_COLUMN' && typeof record.beforeColumn === 'string') return true;
  if (action.type === 'DELETE_COLUMN' && Array.isArray(record.columns)) return true;
  if (action.type === 'SET_CELL' && typeof record.address === 'string') return true;
  if (action.type === 'SET_FORMULA' && typeof record.address === 'string') return true;
  if (action.type === 'FORMAT_RANGE' && typeof record.range === 'string' && record.sheetName) return true;
  if (action.type === 'FILL_DOWN' && typeof record.sourceRange === 'string') return true;
  if (action.type === 'RENAME_SHEET' && typeof record.oldName === 'string') return true;
  if (action.type === 'COPY_SHEET' && typeof record.sourceName === 'string') return true;
  if (action.type === 'DELETE_SHEET' && typeof record.sheetName === 'string') return true;
  if (action.type === 'CREATE_SHEET' && (record.sheetName || record.name)) return true;
  return false;
}

export function toRichAction(action: SheetAction): RichAction | null {
  if (!isRichAction(action)) return null;
  const r = action as unknown as Record<string, unknown>;

  switch (action.type) {
    case 'ADD_ROW':
      return {
        type: 'ADD_ROW',
        sheetName: String(r.sheetName ?? ''),
        afterRow: Number(r.afterRow),
        values: (r.values as CellValue[]) ?? [],
        copyFormatFromRow: r.copyFormatFromRow as number | undefined,
      };
    case 'SET_CELL':
      return {
        type: 'SET_CELL',
        sheetName: String(r.sheetName ?? ''),
        address: String(r.address),
        value: r.value as string | number | boolean | null,
        format: r.format as RichAction extends { type: 'SET_CELL' } ? never : never,
      } as RichAction;
    case 'SET_FORMULA':
      return {
        type: 'SET_FORMULA',
        sheetName: String(r.sheetName ?? ''),
        address: String(r.address),
        formula: String(r.formula ?? ''),
        format: r.format as RichAction extends { type: 'SET_FORMULA' } ? never : never,
      } as RichAction;
    case 'FORMAT_RANGE':
      return {
        type: 'FORMAT_RANGE',
        sheetName: String(r.sheetName ?? ''),
        range: String(r.range),
        format: r.format as RichAction extends { type: 'FORMAT_RANGE' } ? never : never,
      } as RichAction;
    case 'FILL_DOWN':
      return {
        type: 'FILL_DOWN',
        sheetName: String(r.sheetName ?? ''),
        sourceRange: String(r.sourceRange),
        targetRange: String(r.targetRange),
      } as RichAction;
    case 'BATCH_SET':
      return {
        type: 'BATCH_SET',
        sheetName: String(r.sheetName ?? ''),
        operations: r.operations as RichAction extends { type: 'BATCH_SET' } ? never : never,
      } as RichAction;
    case 'DELETE_ROW': {
      const rowNums = Array.isArray(r.rowNumbers)
        ? (r.rowNumbers as number[])
        : Array.isArray(r.rows) && (r.rows as unknown[]).every((n) => typeof n === 'number')
          ? (r.rows as number[])
          : [];
      return {
        type: 'DELETE_ROW',
        sheetName: String(r.sheetName ?? ''),
        rows: rowNums,
      };
    }
    case 'INSERT_COLUMN':
      return {
        type: 'INSERT_COLUMN',
        sheetName: String(r.sheetName ?? ''),
        beforeColumn: String(r.beforeColumn),
        count: Number(r.count ?? 1),
        copyFormatFromColumn: r.copyFormatFromColumn as string | undefined,
      } as RichAction;
    case 'DELETE_COLUMN':
      return {
        type: 'DELETE_COLUMN',
        sheetName: String(r.sheetName ?? ''),
        columns: r.columns as string[],
      } as RichAction;
    case 'ADD_SHEET':
      return {
        type: 'ADD_SHEET',
        name: String(r.name ?? r.sheetName ?? 'Sheet'),
        position: r.position as number | undefined,
        copyFrom: r.copyFrom as string | undefined,
      } as RichAction;
    case 'DELETE_SHEET':
      return {
        type: 'DELETE_SHEET',
        sheetName: String(r.sheetName ?? ''),
      } as RichAction;
    case 'CREATE_SHEET':
      return {
        type: 'ADD_SHEET',
        name: String(r.sheetName ?? r.name ?? 'New Sheet'),
        position: r.position as number | undefined,
        copyFrom: r.copyFrom as string | undefined,
      } as RichAction;
    case 'RENAME_SHEET':
      return {
        type: 'RENAME_SHEET',
        oldName: String(r.oldName ?? r.sheetName ?? ''),
        newName: String(r.newName ?? r.newSheetName ?? ''),
      } as RichAction;
    case 'COPY_SHEET':
      return {
        type: 'COPY_SHEET',
        sourceName: String(r.sourceName ?? r.sheetName ?? ''),
        newName: String(r.newName ?? r.newSheetName ?? ''),
        position: r.position as number | undefined,
      } as RichAction;
    case 'CREATE_TABLE':
      return {
        type: 'CREATE_TABLE',
        sheetName: String(r.sheetName ?? ''),
        range: String(r.range),
        tableName: String(r.tableName ?? ''),
        hasHeaders: Boolean(r.hasHeaders),
        style: r.style as string | undefined,
      } as RichAction;
    case 'DEFINE_NAMED_RANGE':
      return {
        type: 'DEFINE_NAMED_RANGE',
        name: String(r.name ?? ''),
        formula: String(r.formula ?? ''),
        comment: r.comment as string | undefined,
      } as RichAction;
    case 'AUTOFIT_COLUMNS':
      return {
        type: 'AUTOFIT_COLUMNS',
        sheetName: String(r.sheetName ?? ''),
        columns: r.columns as string[] | undefined,
      } as RichAction;
    case 'CLARIFY':
      return {
        type: 'CLARIFY',
        question: String(r.question ?? ''),
        options: r.options as string[] | undefined,
      } as RichAction;
    case 'CHECKPOINT':
      return {
        type: 'CHECKPOINT',
        message: String(r.message ?? ''),
      } as RichAction;
    case 'SORT_RANGE':
      return {
        type: 'SORT_RANGE',
        sheetName: String(r.sheetName ?? ''),
        range: String(r.range ?? ''),
        key: Number(r.key ?? r.col ?? 0),
        ascending: r.ascending !== false,
        hasHeaders: r.hasHeaders !== false,
        columnName: r.columnName as string | undefined,
      } as RichAction;
    default:
      return null;
  }
}

function normalizeToRich(action: SheetAction): RichAction | RichAction[] | null {
  const direct = toRichAction(action);
  if (direct) return direct;
  return convertLegacyToRich(action);
}

/** Convert rich address-based actions to legacy row/col SheetAction when possible. */
export function richToLegacyAction(action: SheetAction): SheetAction | SheetAction[] | null {
  const rich = toRichAction(action);
  if (!rich) return null;

  switch (rich.type) {
    case 'ADD_ROW': {
      const insertAt = rich.afterRow;
      return {
        type: 'INSERT_ROW',
        sheetName: rich.sheetName,
        row: insertAt,
        position: 'below',
        data: rich.values,
      };
    }
    case 'SET_CELL': {
      const cell = parseCellAddress(rich.address);
      if (!cell) return null;
      return {
        type: 'SET_CELL',
        sheetName: rich.sheetName,
        row: cell.row,
        col: cell.col,
        value: rich.value,
        format: rich.format as SheetAction['format'],
      };
    }
    case 'SET_FORMULA': {
      const cell = parseCellAddress(rich.address);
      if (!cell) return null;
      return {
        type: 'SET_FORMULA',
        sheetName: rich.sheetName,
        row: cell.row,
        col: cell.col,
        formula: rich.formula,
        format: rich.format as SheetAction['format'],
      };
    }
    case 'FORMAT_RANGE': {
      const bounds = parseRangeAddress(rich.range);
      if (!bounds) return null;
      return {
        type: 'FORMAT_RANGE',
        sheetName: rich.sheetName,
        row: bounds.row,
        col: bounds.col,
        rowCount: bounds.rowCount,
        colCount: bounds.colCount,
        format: rich.format as SheetAction['format'],
      };
    }
    case 'DELETE_ROW':
      return rich.rows.map((row) => ({
        type: 'DELETE_ROW' as const,
        sheetName: rich.sheetName,
        row: row - 1,
      }));
    case 'INSERT_COLUMN':
      return {
        type: 'INSERT_COLUMN',
        sheetName: rich.sheetName,
        col: columnLetterToIndex(rich.beforeColumn),
        count: rich.count,
        position: 'left',
      };
    case 'DELETE_COLUMN':
      return rich.columns.map((col) => ({
        type: 'DELETE_COLUMN' as const,
        sheetName: rich.sheetName,
        col: columnLetterToIndex(col),
      }));
    case 'ADD_SHEET':
      return {
        type: 'CREATE_SHEET',
        sheetName: rich.name,
        position: rich.position !== undefined ? 'before' : undefined,
      };
    case 'RENAME_SHEET':
      return {
        type: 'RENAME_SHEET',
        sheetName: rich.oldName,
        newSheetName: rich.newName,
      };
    case 'COPY_SHEET':
      return {
        type: 'COPY_SHEET',
        sheetName: rich.sourceName,
        newSheetName: rich.newName,
      };
    default:
      return null;
  }
}

export function partitionActions(actions: SheetAction[]): {
  rich: RichAction[];
  legacy: SheetAction[];
} {
  const rich: RichAction[] = [];
  const legacy: SheetAction[] = [];

  for (const action of actions) {
    const converted = normalizeToRich(action);
    if (converted) {
      if (Array.isArray(converted)) {
        rich.push(...converted);
      } else {
        rich.push(converted);
      }
    } else if (USE_LEGACY_ACTION_ENGINE) {
      legacy.push(action);
    } else {
      console.warn(
        `[Cellix] Action ${action.type} could not be converted to rich format — skipped (set CELLIX_USE_LEGACY_ENGINE=true for legacy fallback)`,
      );
    }
  }

  return { rich, legacy };
}
