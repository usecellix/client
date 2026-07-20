import {
  CopyFilteredRangeAction,
  FormatMatchingRowsAction,
  MoveRangeAction,
} from '@/action.types';
import {
  isLocalRangeAddress,
  parseCellAddress,
  stripSheetPrefix,
} from '../addressUtils';
import {
  applyFilterOperator,
  buildOutputRows,
  filterDataRows,
  findMatchingRowOffsets,
  type RangeFilterSpec,
} from '../rangeFilter';
import { applyRichFormat } from './format.handler';
import { resolveWorksheet } from './resolveWorksheet';

/* global Excel */

async function resolveOrCreateSheet(
  ctx: Excel.RequestContext,
  sheetName: string,
): Promise<Excel.Worksheet> {
  const sheets = ctx.workbook.worksheets;
  const existing = sheets.getItemOrNullObject(sheetName);
  existing.load('isNullObject');
  await ctx.sync();
  if (!existing.isNullObject) {
    return existing;
  }
  const created = sheets.add(sheetName);
  await ctx.sync();
  return created;
}

function resolveSourceRangeAddress(address: string): string {
  const local = stripSheetPrefix(address);
  if (!isLocalRangeAddress(local)) {
    throw new Error(`Invalid source range "${address}"`);
  }
  return local;
}

async function writeOutputRows(
  destSheet: Excel.Worksheet,
  destStartCell: string,
  outputRows: unknown[][],
  ctx: Excel.RequestContext,
): Promise<void> {
  if (outputRows.length === 0) return;
  const start = parseCellAddress(destStartCell);
  if (!start) {
    throw new Error(`Invalid destStartCell "${destStartCell}"`);
  }
  const colCount = Math.max(...outputRows.map((row) => row.length), 1);
  const padded = outputRows.map((row) => {
    const next = [...row];
    while (next.length < colCount) next.push(null);
    return next;
  });
  const destRange = destSheet.getRangeByIndexes(
    start.row,
    start.col,
    padded.length,
    colCount,
  );
  destRange.values = padded as (string | number | boolean)[][];
  await ctx.sync();
}

async function clearMatchedSourceRows(
  sourceSheet: Excel.Worksheet,
  sourceRangeAddress: string,
  hasHeaders: boolean,
  filter: RangeFilterSpec,
  ctx: Excel.RequestContext,
): Promise<void> {
  const range = sourceSheet.getRange(sourceRangeAddress);
  range.load(['values', 'rowIndex', 'columnIndex', 'columnCount']);
  await ctx.sync();

  const values = (range.values ?? []) as unknown[][];
  if (values.length === 0) return;

  const headerRow = hasHeaders ? values[0] : null;
  if (!headerRow) {
    throw new Error('Move with filter requires hasHeaders: true');
  }

  const colIndex = headerRow.findIndex(
    (cell) => String(cell ?? '').trim().toLowerCase() === filter.column.trim().toLowerCase(),
  );
  if (colIndex === -1) {
    throw new Error(`Column "${filter.column}" not found in source range`);
  }

  const dataStart = hasHeaders ? 1 : 0;
  const absoluteRowsToDelete: number[] = [];
  for (let i = dataStart; i < values.length; i += 1) {
    if (applyFilterOperator(values[i]?.[colIndex], filter)) {
      absoluteRowsToDelete.push(range.rowIndex + i);
    }
  }

  // Delete bottom-to-top so indices stay stable
  absoluteRowsToDelete.sort((a, b) => b - a);
  for (const rowIndex of absoluteRowsToDelete) {
    sourceSheet.getRangeByIndexes(rowIndex, 0, 1, 1).getEntireRow().delete('Up');
  }
  if (absoluteRowsToDelete.length > 0) {
    await ctx.sync();
  }
}

export async function handleCopyFilteredRange(
  action: CopyFilteredRangeAction,
  ctx: Excel.RequestContext,
): Promise<{ rowsCopied: number }> {
  const sourceSheet = resolveWorksheet(ctx, action.sourceSheet);
  const sourceRangeAddress = resolveSourceRangeAddress(action.sourceRange);
  const sourceRange = sourceSheet.getRange(sourceRangeAddress);
  sourceRange.load('values');
  await ctx.sync();

  const rows = (sourceRange.values ?? []) as unknown[][];
  const { headerRow, filteredRows } = filterDataRows(
    rows,
    action.hasHeaders,
    action.filter,
  );
  const outputRows = buildOutputRows(headerRow, filteredRows);

  const destSheet = await resolveOrCreateSheet(ctx, action.destSheet);
  await writeOutputRows(destSheet, action.destStartCell, outputRows, ctx);

  if (action.mode === 'move' && action.filter) {
    await clearMatchedSourceRows(
      sourceSheet,
      sourceRangeAddress,
      action.hasHeaders,
      action.filter,
      ctx,
    );
  }

  return { rowsCopied: filteredRows.length };
}

export async function handleMoveRange(
  action: MoveRangeAction,
  ctx: Excel.RequestContext,
): Promise<{ rowsMoved: number }> {
  const sourceSheet = resolveWorksheet(ctx, action.sourceSheet);
  const sourceRangeAddress = resolveSourceRangeAddress(action.sourceRange);
  const sourceRange = sourceSheet.getRange(sourceRangeAddress);
  sourceRange.load(['values', 'rowCount', 'columnCount']);
  await ctx.sync();

  const rows = (sourceRange.values ?? []) as unknown[][];
  const destSheet = await resolveOrCreateSheet(ctx, action.destSheet);
  await writeOutputRows(destSheet, action.destStartCell, rows, ctx);

  sourceRange.clear(Excel.ClearApplyTo.contents);
  await ctx.sync();

  return { rowsMoved: rows.length };
}

export async function handleFormatMatchingRows(
  action: FormatMatchingRowsAction,
  ctx: Excel.RequestContext,
): Promise<{ rowsFormatted: number }> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const rangeAddress = resolveSourceRangeAddress(action.range);
  const range = sheet.getRange(rangeAddress);
  range.load(['values', 'rowIndex', 'columnIndex', 'columnCount']);
  await ctx.sync();

  const rows = (range.values ?? []) as unknown[][];
  const offsets = findMatchingRowOffsets(rows, action.hasHeaders !== false, action.filter);

  for (const offset of offsets) {
    const rowRange = sheet.getRangeByIndexes(
      range.rowIndex + offset,
      range.columnIndex,
      1,
      range.columnCount,
    );
    applyRichFormat(rowRange, action.format);
  }

  if (offsets.length > 0) {
    await ctx.sync();
  }

  return { rowsFormatted: offsets.length };
}

/** Re-export for unit tests */
export { applyFilterOperator, filterDataRows, buildOutputRows, findMatchingRowOffsets };
