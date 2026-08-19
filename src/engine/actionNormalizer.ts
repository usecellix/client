import { CellValue, RichAction, CreateChartAction, UpdateChartAction, AggregateTableAction, FormatSpec } from '@/action.types';
import { SheetAction } from '@/types/sheet-actions';
import { columnLetterToIndex, parseCellAddress, parseRangeAddress } from './addressUtils';
import { convertLegacyToRich } from './legacyConverter';

function isRichAction(action: SheetAction): boolean {
  const richOnly = new Set([
    'BATCH_SET',
    'CREATE_TABLE',
    'CREATE_CHART',
    'DEFINE_NAMED_RANGE',
    'AUTOFIT_COLUMNS',
    'CLARIFY',
    'CHECKPOINT',
    'ADD_SHEET',
    'DELETE_SHEET',
    'SORT_RANGE',
    'COPY_FILTERED_RANGE',
    'FORMAT_MATCHING_ROWS',
    'SET_MATCHING_ROWS',
    'MOVE_RANGE',
    'CREATE_CHART',
    'UPDATE_CHART',
    'AGGREGATE_TABLE',
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
  if (
    action.type === 'INSERT_COLUMN' &&
    typeof record.columnName === 'string' &&
    (record.position === 'afterLastColumn' ||
      typeof record.afterColumn === 'string' ||
      (record.position &&
        typeof record.position === 'object' &&
        typeof (record.position as { afterColumn?: string }).afterColumn === 'string'))
  ) {
    return true;
  }
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
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
      } as RichAction;
    case 'SET_FORMULA':
      return {
        type: 'SET_FORMULA',
        sheetName: String(r.sheetName ?? ''),
        address: String(r.address),
        formula: String(r.formula ?? ''),
        format: r.format as RichAction extends { type: 'SET_FORMULA' } ? never : never,
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
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
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
      } as RichAction;
    case 'BATCH_SET':
      return {
        type: 'BATCH_SET',
        sheetName: String(r.sheetName ?? ''),
        operations: r.operations as RichAction extends { type: 'BATCH_SET' } ? never : never,
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
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
        beforeColumn: r.beforeColumn as string | undefined,
        count: Number(r.count ?? 1),
        copyFormatFromColumn: r.copyFormatFromColumn as string | undefined,
        columnName: r.columnName as string | undefined,
        position:
          r.position === 'afterLastColumn'
            ? 'afterLastColumn'
            : r.position &&
                typeof r.position === 'object' &&
                typeof (r.position as { afterColumn?: string }).afterColumn === 'string'
              ? { afterColumn: (r.position as { afterColumn: string }).afterColumn }
              : undefined,
        afterColumn: r.afterColumn as string | undefined,
        formula: r.formula as string | undefined,
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
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
        tableName: String(r.tableName ?? r.name ?? '').trim(),
        hasHeaders: r.hasHeaders === undefined ? true : Boolean(r.hasHeaders),
        style: r.style as string | undefined,
      } as RichAction;
    case 'CREATE_CHART':
      return {
        type: 'CREATE_CHART',
        sheetName: String(r.sheetName ?? ''),
        sourceSheetName: String(r.sourceSheetName ?? r.sheetName ?? ''),
        sourceRange: String(r.sourceRange ?? r.range ?? ''),
        chartType: String(r.chartType ?? 'ColumnClustered'),
        title: r.title as string | undefined,
        startCell: String(r.startCell ?? r.destCell ?? 'A1'),
        endCell: String(r.endCell ?? 'H16'),
        destCell: r.destCell as string | undefined,
        colorScheme: r.colorScheme as CreateChartAction['colorScheme'],
        chartId: r.chartId as string | undefined,
      } as RichAction;
    case 'UPDATE_CHART':
      return {
        type: 'UPDATE_CHART',
        sheetName: String(r.sheetName ?? ''),
        chartId: String(r.chartId ?? ''),
        chartType: r.chartType as string | undefined,
        colorScheme: r.colorScheme as UpdateChartAction['colorScheme'],
      } as RichAction;
    case 'AGGREGATE_TABLE':
      return {
        type: 'AGGREGATE_TABLE',
        sourceSheet: String(r.sourceSheet ?? r.sheetName ?? ''),
        sourceRange: String(r.sourceRange ?? r.range ?? ''),
        groupByColumn: String(r.groupByColumn ?? ''),
        groupByTransform: r.groupByTransform as AggregateTableAction['groupByTransform'],
        aggregations: (Array.isArray(r.aggregations)
          ? r.aggregations
          : []) as AggregateTableAction['aggregations'],
        sortBy: r.sortBy as AggregateTableAction['sortBy'],
        topN: typeof r.topN === 'number' ? r.topN : undefined,
        destSheet: String(r.destSheet ?? ''),
        destStartCell: String(r.destStartCell ?? r.startCell ?? 'A1'),
        hasHeaders: r.hasHeaders !== false,
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
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
    case 'COPY_FILTERED_RANGE': {
      const filter =
        r.filter && typeof r.filter === 'object'
          ? (r.filter as {
              column: string;
              operator:
                | 'equals'
                | 'contains'
                | 'greaterThan'
                | 'lessThan'
                | 'notEquals'
                | 'lengthEquals'
                | 'lengthNotEquals'
                | 'matchesRegex'
                | 'notMatchesRegex';
              value: string | number;
            })
          : undefined;
      return {
        type: 'COPY_FILTERED_RANGE',
        sourceSheet: String(r.sourceSheet ?? r.sheetName ?? ''),
        sourceRange: String(r.sourceRange ?? r.range ?? ''),
        hasHeaders: r.hasHeaders !== false,
        destSheet: String(r.destSheet ?? ''),
        destStartCell: String(r.destStartCell ?? r.startCell ?? 'A1'),
        filter,
        mode: r.mode === 'move' ? 'move' : 'copy',
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
      };
    }
    case 'FORMAT_MATCHING_ROWS': {
      const filter =
        r.filter && typeof r.filter === 'object'
          ? (r.filter as {
              column: string;
              operator:
                | 'equals'
                | 'contains'
                | 'greaterThan'
                | 'lessThan'
                | 'notEquals'
                | 'lengthEquals'
                | 'lengthNotEquals'
                | 'matchesRegex'
                | 'notMatchesRegex';
              value: string | number;
            })
          : undefined;
      if (!filter?.column) return null;
      const format =
        r.format && typeof r.format === 'object'
          ? (r.format as FormatSpec)
          : { fillColor: '#FFC7CE' };
      return {
        type: 'FORMAT_MATCHING_ROWS',
        sheetName: String(r.sheetName ?? ''),
        range: String(r.range ?? ''),
        hasHeaders: r.hasHeaders !== false,
        filter,
        format: format.clearFill
          ? { clearFill: true }
          : {
              ...format,
              fillColor: format.fillColor ?? '#FFC7CE',
            },
      };
    }
    case 'SET_MATCHING_ROWS': {
      const filter =
        r.filter && typeof r.filter === 'object'
          ? (r.filter as {
              column: string;
              operator:
                | 'equals'
                | 'contains'
                | 'greaterThan'
                | 'lessThan'
                | 'notEquals'
                | 'lengthEquals'
                | 'lengthNotEquals'
                | 'matchesRegex'
                | 'notMatchesRegex';
              value: string | number;
            })
          : undefined;
      if (filter && !filter.column) return null;
      const targetColumn = String(r.targetColumn ?? r.columnName ?? '').trim();
      if (!targetColumn) return null;
      if (
        r.value === undefined ||
        (typeof r.value !== 'string' &&
          typeof r.value !== 'number' &&
          typeof r.value !== 'boolean')
      ) {
        return null;
      }
      return {
        type: 'SET_MATCHING_ROWS',
        sheetName: String(r.sheetName ?? ''),
        range: String(r.range ?? ''),
        hasHeaders: r.hasHeaders !== false,
        ...(filter?.column ? { filter } : {}),
        targetColumn,
        value: r.value,
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
      };
    }
    case 'MOVE_RANGE':
      return {
        type: 'MOVE_RANGE',
        sourceSheet: String(r.sourceSheet ?? r.sheetName ?? ''),
        sourceRange: String(r.sourceRange ?? r.range ?? ''),
        destSheet: String(r.destSheet ?? ''),
        destStartCell: String(r.destStartCell ?? r.startCell ?? 'A1'),
        explicitOverwriteConfirmed: r.explicitOverwriteConfirmed === true,
      };
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
    case 'INSERT_COLUMN': {
      if (rich.columnName) {
        return {
          type: 'INSERT_COLUMN',
          sheetName: rich.sheetName,
          columnName: rich.columnName,
          position:
            rich.position === 'afterLastColumn'
              ? 'afterLastColumn'
              : undefined,
          afterColumn:
            typeof rich.position === 'object'
              ? rich.position.afterColumn
              : rich.afterColumn,
          formula: rich.formula,
          explicitOverwriteConfirmed: rich.explicitOverwriteConfirmed,
        };
      }
      if (!rich.beforeColumn) return null;
      return {
        type: 'INSERT_COLUMN',
        sheetName: rich.sheetName,
        col: columnLetterToIndex(rich.beforeColumn),
        count: rich.count,
        position: 'left',
      };
    }
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
  unsupported: SheetAction[];
} {
  const rich: RichAction[] = [];
  const unsupported: SheetAction[] = [];

  for (const action of actions) {
    const converted = normalizeToRich(action);
    if (converted) {
      if (Array.isArray(converted)) {
        rich.push(...converted);
      } else {
        rich.push(converted);
      }
    } else {
      unsupported.push(action);
    }
  }

  return { rich, unsupported };
}
