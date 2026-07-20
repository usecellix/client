import { AggregateTableAction } from '@/action.types';
import {
  isLocalRangeAddress,
  parseCellAddress,
  stripSheetPrefix,
} from '../addressUtils';
import { buildAggregateTable } from '../aggregateTable';

/* global Excel */

async function resolveOrCreateSheet(
  ctx: Excel.RequestContext,
  sheetName: string,
): Promise<Excel.Worksheet> {
  const sheets = ctx.workbook.worksheets;
  const existing = sheets.getItemOrNullObject(sheetName);
  existing.load('isNullObject');
  await ctx.sync();
  if (!existing.isNullObject) return existing;
  const created = sheets.add(sheetName);
  await ctx.sync();
  return created;
}

export async function handleAggregateTable(
  action: AggregateTableAction,
  ctx: Excel.RequestContext,
): Promise<{ rowsWritten: number }> {
  if (
    !action.sourceSheet?.trim() ||
    !action.sourceRange?.trim() ||
    !action.destSheet?.trim() ||
    !action.destStartCell?.trim() ||
    !action.groupByColumn?.trim() ||
    !Array.isArray(action.aggregations) ||
    action.aggregations.length === 0
  ) {
    throw new Error(
      'AGGREGATE_TABLE requires sourceSheet, sourceRange, groupByColumn, aggregations, destSheet, destStartCell',
    );
  }

  const sourceRangeAddr = stripSheetPrefix(action.sourceRange);
  if (!isLocalRangeAddress(sourceRangeAddr)) {
    throw new Error(`Invalid source range "${action.sourceRange}"`);
  }

  const sourceSheet = ctx.workbook.worksheets.getItem(action.sourceSheet);
  const sourceRange = sourceSheet.getRange(sourceRangeAddr);
  sourceRange.load('values');
  await ctx.sync();

  const outputRows = buildAggregateTable({
    rows: (sourceRange.values ?? []) as unknown[][],
    hasHeaders: action.hasHeaders !== false,
    groupByColumn: action.groupByColumn,
    aggregations: action.aggregations,
    sortBy: action.sortBy,
    topN: action.topN,
  });

  if (outputRows.length === 0) return { rowsWritten: 0 };

  const destSheet = await resolveOrCreateSheet(ctx, action.destSheet);
  const start = parseCellAddress(action.destStartCell);
  if (!start) {
    throw new Error(`Invalid destStartCell "${action.destStartCell}"`);
  }

  const colCount = Math.max(...outputRows.map((r) => r.length), 1);
  const padded = outputRows.map((row) => {
    const next = [...row];
    while (next.length < colCount) next.push(null);
    return next;
  });

  destSheet.getRangeByIndexes(start.row, start.col, padded.length, colCount).values =
    padded as (string | number | boolean)[][];
  await ctx.sync();

  return { rowsWritten: Math.max(0, outputRows.length - 1) };
}
