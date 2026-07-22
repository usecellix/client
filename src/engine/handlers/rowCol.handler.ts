import {
  AddRowAction,
  DeleteRowAction,
  InsertColumnAction,
  DeleteColumnAction,
  InsertColumnPosition,
} from '@/action.types';
import { columnIndexToLetter, columnLetterToIndex, parseRangeAddress } from '../addressUtils';
import {
  OverwriteGuardError,
  rangeHasExistingData,
} from '../overwriteGuard';
import { stampAppliedBounds } from '../selectRanges';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

export async function handleAddRow(action: AddRowAction, ctx: Excel.RequestContext): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const insertRow = action.afterRow + 1;
  const insertRange = sheet.getRange(`${insertRow}:${insertRow}`);
  insertRange.insert('Down');

  const colCount = Math.max(action.values.length, 1);
  const newRowRange = sheet.getRangeByIndexes(insertRow - 1, 0, 1, colCount);
  newRowRange.values = [action.values];

  if (action.copyFormatFromRow) {
    const sourceRow = sheet.getRange(`${action.copyFormatFromRow}:${action.copyFormatFromRow}`);
    const targetRow = sheet.getRange(`${insertRow}:${insertRow}`);
    targetRow.copyFrom(sourceRow, Excel.RangeCopyType.formats, false, false);
  }

  stampAppliedBounds(action, {
    sheetName: action.sheetName,
    row: insertRow - 1,
    col: 0,
    rowCount: 1,
    colCount,
  });

  await ctx.sync();
}

export async function handleDeleteRow(
  action: DeleteRowAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const sorted = [...action.rows].sort((a, b) => b - a);
  for (const row of sorted) {
    sheet.getRange(`${row}:${row}`).delete('Up');
  }
  if (action.rows.length) {
    const min = Math.min(...action.rows) - 1;
    const max = Math.max(...action.rows) - 1;
    stampAppliedBounds(action as never, {
      sheetName: action.sheetName,
      row: Math.max(min, 0),
      col: 0,
      rowCount: Math.max(max - min + 1, 1),
      colCount: 1,
    });
  }
  await ctx.sync();
}

function isSemanticInsertColumn(action: InsertColumnAction): boolean {
  if (!action.columnName?.trim()) return false;
  if (action.position === 'afterLastColumn') return true;
  if (action.afterColumn?.trim()) return true;
  if (
    action.position &&
    typeof action.position === 'object' &&
    typeof (action.position as { afterColumn?: string }).afterColumn === 'string'
  ) {
    return true;
  }
  return false;
}

function resolveSemanticPosition(
  action: InsertColumnAction,
): InsertColumnPosition {
  if (action.position === 'afterLastColumn') return 'afterLastColumn';
  if (
    action.position &&
    typeof action.position === 'object' &&
    typeof action.position.afterColumn === 'string'
  ) {
    return { afterColumn: action.position.afterColumn };
  }
  if (action.afterColumn?.trim()) {
    return { afterColumn: action.afterColumn.trim() };
  }
  return 'afterLastColumn';
}

function formulaForExcelRow(template: string, excelRow: number): string {
  if (template.includes('{row}')) {
    return template.replace(/\{row\}/g, String(excelRow));
  }
  return template;
}

type UsedBounds = {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
  values: unknown[][];
};

/**
 * Resolve live used-range as 0-based bounds. Prefer parsing address (reliable);
 * fall back to A1 origin using rowCount/columnCount only — never trust unloaded row/column props.
 */
async function resolveUsedBounds(
  sheet: Excel.Worksheet,
  ctx: Excel.RequestContext,
): Promise<UsedBounds | null> {
  const used = sheet.getUsedRange();
  if (!used) return null;
  used.load(['address', 'rowCount', 'columnCount', 'values']);
  await ctx.sync();

  const parsed = typeof used.address === 'string' ? parseRangeAddress(used.address) : null;
  if (parsed) {
    return {
      row: parsed.row,
      col: parsed.col,
      rowCount: used.rowCount,
      colCount: used.columnCount,
      values: (used.values ?? []) as unknown[][],
    };
  }

  // Address unparsable — assume table starts at A1 (standard for these sheets).
  return {
    row: 0,
    col: 0,
    rowCount: used.rowCount,
    colCount: used.columnCount,
    values: (used.values ?? []) as unknown[][],
  };
}

/**
 * Semantic INSERT_COLUMN: place a named column using the live used range.
 * afterLastColumn writes into the first empty column after usedRange (no shift).
 * { afterColumn } inserts after that header and shifts subsequent columns right.
 */
async function handleSemanticInsertColumn(
  action: InsertColumnAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const columnName = String(action.columnName).trim();
  const position = resolveSemanticPosition(action);

  const used = await resolveUsedBounds(sheet, ctx);

  let targetCol: number;
  let headerRowIndex = 0;
  let dataRowCount = 0;

  if (!used || used.rowCount === 0 || used.colCount === 0) {
    targetCol = 0;
    headerRowIndex = 0;
    dataRowCount = 0;
  } else {
    headerRowIndex = used.row;
    dataRowCount = Math.max(used.rowCount - 1, 0);

    if (position === 'afterLastColumn') {
      targetCol = used.col + used.colCount;
    } else {
      const afterName = position.afterColumn.trim().toLowerCase();
      const headerValues = (used.values[0] ?? []) as unknown[];
      const headerOffset = headerValues.findIndex(
        (cell) => String(cell ?? '').trim().toLowerCase() === afterName,
      );
      if (headerOffset < 0) {
        throw new Error(
          `INSERT_COLUMN: column "${position.afterColumn}" not found in the sheet header row`,
        );
      }
      targetCol = used.col + headerOffset + 1;
      sheet.getRangeByIndexes(0, targetCol, 1, 1).getEntireColumn().insert('Right');
      await ctx.sync();
    }
  }

  const probeRows = Math.max(dataRowCount + 1, 1);
  const probe = sheet.getRangeByIndexes(headerRowIndex, targetCol, probeRows, 1);
  probe.load('values');
  await ctx.sync();
  if (rangeHasExistingData(probe.values as unknown[][]) && !action.explicitOverwriteConfirmed) {
    const letter = columnIndexToLetter(targetCol);
    throw new OverwriteGuardError({
      message:
        `Write blocked: target column ${letter} already contains data. ` +
        `INSERT_COLUMN expected an empty column for "${columnName}".`,
      targetRange: `${letter}:${letter}`,
      sampleExistingValues: (probe.values as unknown[][]).slice(0, 3),
    });
  }

  // Header on the same row as existing headers (never below the table).
  const headerLetter = columnIndexToLetter(targetCol);
  const headerExcelRow = headerRowIndex + 1;
  sheet.getRange(`${headerLetter}${headerExcelRow}`).values = [[columnName]];

  if (action.formula && dataRowCount > 0) {
    const firstDataRow = headerRowIndex + 1;
    const firstExcelRow = firstDataRow + 1;

    if (action.formula.includes('{row}')) {
      const formulas = Array.from({ length: dataRowCount }, (_, i) => [
        formulaForExcelRow(action.formula!, firstExcelRow + i),
      ]);
      sheet.getRangeByIndexes(firstDataRow, targetCol, dataRowCount, 1).formulas = formulas;
    } else {
      const firstCell = sheet.getRangeByIndexes(firstDataRow, targetCol, 1, 1);
      firstCell.formulas = [[formulaForExcelRow(action.formula, firstExcelRow)]];
      if (dataRowCount > 1) {
        const dest = sheet.getRangeByIndexes(firstDataRow, targetCol, dataRowCount, 1);
        dest.copyFrom(firstCell, Excel.RangeCopyType.formulas, false, false);
      }
    }
  }

  stampAppliedBounds(action, {
    sheetName: action.sheetName,
    row: headerRowIndex,
    col: targetCol,
    rowCount: Math.max(dataRowCount + 1, 1),
    colCount: 1,
  });

  await ctx.sync();
}

export async function handleInsertColumn(
  action: InsertColumnAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  if (isSemanticInsertColumn(action)) {
    return handleSemanticInsertColumn(action, ctx);
  }

  const sheet = resolveWorksheet(ctx, action.sheetName);
  const col = action.beforeColumn;
  if (!col) {
    throw new Error(
      'INSERT_COLUMN requires beforeColumn (legacy) or columnName + position (semantic)',
    );
  }
  const count = action.count ?? 1;
  for (let i = 0; i < count; i += 1) {
    sheet.getRange(`${col}:${col}`).insert('Right');
  }
  stampAppliedBounds(action, {
    sheetName: action.sheetName,
    row: 0,
    col: columnLetterToIndex(col),
    rowCount: 1,
    colCount: count,
  });
  await ctx.sync();
}

export async function handleDeleteColumn(
  action: DeleteColumnAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const sorted = [...action.columns].sort().reverse();
  for (const col of sorted) {
    sheet.getRange(`${col}:${col}`).delete('Left');
  }
  await ctx.sync();
}
