import {
  ConditionalFormatAction,
  ConditionalFormatOperator,
  CopyFilteredRangeAction,
  DeleteConditionalFormatAction,
  FormatMatchingRowsAction,
  MoveRangeAction,
  SetMatchingRowsAction,
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
  resolveFilterColumnIndex,
  type RangeFilterSpec,
} from '../rangeFilter';
import { applyConditionalRangeFormat, applyRichFormat } from './format.handler';
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
    (cell) =>
      String(cell ?? '').trim().toLowerCase() === String(filter.column).trim().toLowerCase(),
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

  // Preserve header + matched-row formatting when we filtered by header column.
  // (Values are written above; this copies formats only.)
  if (action.filter && action.hasHeaders && headerRow && outputRows.length > 0) {
    const start = parseCellAddress(action.destStartCell);
    if (start) {
      const colCount = Math.max(...outputRows.map((row) => row.length), 1);
      const formatsCopyType = (Excel as any)?.RangeCopyType?.formats;
      if (formatsCopyType) {
        // Copy header formatting. getRangeByIndexes is a Worksheet method — the
        // Range-relative equivalent is getCell() widened with getResizedRange().
        const sourceHeaderRange = sourceRange.getCell(0, 0).getResizedRange(0, colCount - 1);
        const destHeaderRange = destSheet.getRangeByIndexes(start.row, start.col, 1, colCount);
        destHeaderRange.copyFrom(sourceHeaderRange, formatsCopyType);

        // Copy formatting for each matched data row (in the same order we wrote values).
        const matchedOffsets = findMatchingRowOffsets(rows, action.hasHeaders, action.filter);
        for (let i = 0; i < matchedOffsets.length; i++) {
          const sourceRowIndex = matchedOffsets[i]!;
          const destRowIndex = i + 1; // outputRows: [header, ...matchedDataRows]
          const sourceRowRange = sourceRange
            .getCell(sourceRowIndex, 0)
            .getResizedRange(0, colCount - 1);
          const destRowRange = destSheet.getRangeByIndexes(start.row + destRowIndex, start.col, 1, colCount);
          destRowRange.copyFrom(sourceRowRange, formatsCopyType);
        }
        await ctx.sync();
      }
    }
  }

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

export async function handleSetMatchingRows(
  action: SetMatchingRowsAction,
  ctx: Excel.RequestContext,
): Promise<{ rowsUpdated: number }> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const rangeAddress = resolveSourceRangeAddress(action.range);
  const range = sheet.getRange(rangeAddress);
  range.load(['values', 'rowIndex', 'columnIndex', 'columnCount']);
  await ctx.sync();

  const rows = (range.values ?? []) as unknown[][];
  if (rows.length === 0) return { rowsUpdated: 0 };

  const hasHeaders = action.hasHeaders !== false;
  const headerRow = hasHeaders ? rows[0] : null;
  if (!headerRow) {
    throw new Error('SET_MATCHING_ROWS requires hasHeaders: true to resolve targetColumn');
  }

  const targetColOffset = resolveFilterColumnIndex(headerRow, action.targetColumn);
  const absoluteTargetCol = range.columnIndex + targetColOffset;
  const offsets = action.filter
    ? findMatchingRowOffsets(rows, hasHeaders, action.filter)
    : Array.from({ length: Math.max(rows.length - (hasHeaders ? 1 : 0), 0) }, (_, i) =>
        hasHeaders ? i + 1 : i,
      );

  for (const offset of offsets) {
    const cell = sheet.getRangeByIndexes(
      range.rowIndex + offset,
      absoluteTargetCol,
      1,
      1,
    );
    cell.values = [[action.value]];
  }

  if (offsets.length > 0) {
    await ctx.sync();
  }

  return { rowsUpdated: offsets.length };
}

const CONDITIONAL_FORMAT_OPERATOR_MAP: Record<
  ConditionalFormatOperator,
  Excel.ConditionalCellValueRule['operator']
> = {
  greaterThan: 'GreaterThan',
  greaterThanOrEqual: 'GreaterThanOrEqual',
  lessThan: 'LessThan',
  lessThanOrEqual: 'LessThanOrEqual',
  equalTo: 'EqualTo',
  notEqualTo: 'NotEqualTo',
  between: 'Between',
  notBetween: 'NotBetween',
};

/** Office.js's ConditionalCellValueRule.formula1/2 take a formula string — quote text values. */
function toConditionalFormatFormula(value: number | string): string {
  return typeof value === 'number' ? String(value) : `"${String(value).replace(/"/g, '""')}"`;
}

const CONDITIONAL_FORMAT_RULE_KIND_TO_OFFICE_TYPE: Record<
  ConditionalFormatAction['rule']['kind'],
  Excel.ConditionalFormatType | 'CellValue' | 'Custom' | 'TopBottom' | 'ColorScale'
> = {
  cellValue: 'CellValue',
  formula: 'Custom',
  topBottom: 'TopBottom',
  colorScale: 'ColorScale',
};

/**
 * Applies a live, re-evaluating Excel conditional-format rule — never a
 * one-shot computed fill (PRD M7). When `action.existingRuleId` is set
 * (TASKS.md #38), resolves the already-existing rule via
 * `conditionalFormats.getItem(id)` and mutates it in place instead of
 * `.add(...)`-ing a duplicate — `rule.kind` must match the existing rule's
 * own kind; Office.js throws if the wrong sub-object (`cellValue`/`custom`/
 * `topBottom`/`colorScale`) is written for the retrieved rule's actual type.
 *
 * On a plain create, reads back the real Excel-assigned rule `.id` and
 * returns it — this is the apply-time id capture TASKS.md #40 needs to make
 * revert possible: the backend can't know this id at preview time (Office.js
 * only assigns it once `.add()` actually runs), so `RichActionEngine`
 * collects it here and reports it to `POST /audit/apply/:changeSetId` on
 * accept, which patches it into the change set's `structuralOps`.
 */
export async function handleConditionalFormat(
  action: ConditionalFormatAction,
  ctx: Excel.RequestContext,
): Promise<{ createdConditionalFormatId?: string } | void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);

  const conditionalFormat = action.existingRuleId
    ? sheet.getUsedRange().conditionalFormats.getItem(action.existingRuleId)
    : sheet
        .getRange(resolveSourceRangeAddress(action.range))
        .conditionalFormats.add(
          CONDITIONAL_FORMAT_RULE_KIND_TO_OFFICE_TYPE[action.rule.kind] as Excel.ConditionalFormatType,
        );

  if (action.rule.kind === 'formula') {
    conditionalFormat.custom.rule.formula = action.rule.formula;
    applyConditionalRangeFormat(conditionalFormat.custom.format, action.rule.format);
  } else if (action.rule.kind === 'topBottom') {
    const type: Excel.ConditionalTopBottomRule['type'] = action.rule.isPercent
      ? action.rule.side === 'top'
        ? 'TopPercent'
        : 'BottomPercent'
      : action.rule.side === 'top'
        ? 'TopItems'
        : 'BottomItems';
    conditionalFormat.topBottom.rule = { type, rank: action.rule.rank };
    applyConditionalRangeFormat(conditionalFormat.topBottom.format, action.rule.format);
  } else if (action.rule.kind === 'colorScale') {
    const { colors } = action.rule;
    const criteria: Excel.ConditionalColorScaleCriteria = {
      minimum: { type: 'LowestValue', color: colors[0] },
      maximum: { type: 'HighestValue', color: colors[colors.length - 1]! },
    };
    if (colors.length === 3) {
      criteria.midpoint = { formula: '50', type: 'Percentile', color: colors[1] };
    }
    conditionalFormat.colorScale.criteria = criteria;
  } else {
    conditionalFormat.cellValue.rule = {
      operator: CONDITIONAL_FORMAT_OPERATOR_MAP[action.rule.operator],
      formula1: toConditionalFormatFormula(action.rule.value),
      ...(action.rule.value2 !== undefined
        ? { formula2: toConditionalFormatFormula(action.rule.value2) }
        : {}),
    };
    applyConditionalRangeFormat(conditionalFormat.cellValue.format, action.rule.format);
  }

  if (!action.existingRuleId) {
    conditionalFormat.load('id');
  }
  await ctx.sync();

  if (!action.existingRuleId) {
    return { createdConditionalFormatId: conditionalFormat.id };
  }
}

/**
 * Revert-only inverse of a CONDITIONAL_FORMAT create (TASKS.md #40) — deletes
 * a specific rule by its real Excel-assigned id, resolved via the sheet's
 * used range (the same range CONDITIONAL_FORMAT rules were originally read
 * back from in `workbookReader.ts`).
 */
export async function handleDeleteConditionalFormat(
  action: DeleteConditionalFormatAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = resolveWorksheet(ctx, action.sheetName);
  const conditionalFormat = sheet.getUsedRange().conditionalFormats.getItem(action.ruleId);
  conditionalFormat.delete();
  await ctx.sync();
}

/** Re-export for unit tests */
export { applyFilterOperator, filterDataRows, buildOutputRows, findMatchingRowOffsets };
