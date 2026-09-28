import { CreateChartAction, DeleteChartAction, UpdateChartAction } from '@/action.types';
import { stripSheetPrefix } from '../addressUtils';

/* global Excel */

const CHART_COLOR_HEX: Record<string, string> = {
  blue: '#4472C4',
  grey: '#7F7F7F',
  blueGrey: '#5B9BD5',
  green: '#70AD47',
  red: '#C00000',
  orange: '#ED7D31',
  purple: '#7030A0',
  yellow: '#FFC000',
};

async function applyChartColorScheme(
  chart: Excel.Chart,
  ctx: Excel.RequestContext,
  colorScheme: string | undefined,
): Promise<void> {
  if (!colorScheme || colorScheme === 'default') return;
  const fill = CHART_COLOR_HEX[colorScheme];
  if (!fill) return;
  chart.series.load('items');
  await ctx.sync();
  for (const series of chart.series.items) {
    series.format.fill.setSolidColor(fill);
  }
}

/** Map friendly aliases + Excel enum names to ChartType. */
export function resolveChartType(value: string): Excel.ChartType {
  const normalized = value.replace(/[\s_-]/g, '').toLowerCase();

  const aliases: Record<string, Excel.ChartType> = {
    bar: Excel.ChartType.barClustered,
    barhorizontal: Excel.ChartType.barClustered,
    horizontalbar: Excel.ChartType.barClustered,
    barclustered: Excel.ChartType.barClustered,
    column: Excel.ChartType.columnClustered,
    columnclustered: Excel.ChartType.columnClustered,
    pie: Excel.ChartType.pie,
    line: Excel.ChartType.line,
    doughnut: Excel.ChartType.doughnut,
  };
  if (aliases[normalized]) {
    return aliases[normalized];
  }

  const chartTypes = Excel.ChartType as unknown as Record<string, Excel.ChartType>;
  const match = Object.entries(chartTypes).find(
    ([key, enumValue]) =>
      key.replace(/[\s_-]/g, '').toLowerCase() === normalized ||
      String(enumValue).replace(/[\s_-]/g, '').toLowerCase() === normalized,
  );
  return match?.[1] ?? Excel.ChartType.columnClustered;
}

/**
 * True when a first column is a label column: every filled cell below its
 * header is non-numeric text.
 */
export function isLabelColumn(columnValues: unknown[][]): boolean {
  const body = columnValues.slice(1).map((row) => row[0]).filter((v) => v !== '' && v != null);
  return body.length > 0 && body.every((v) => typeof v === 'string' && Number.isNaN(Number(v)));
}

/**
 * Excel guesses whether a source's first column is the category axis. With a
 * filled header cell above it, it can guess wrong: a live dashboard charted
 * "Month" as a SERIES (legend: Month, Total, Paid, Pending) over an axis of
 * 1…12 instead of January…December. TASKS.md #327. When the first column is
 * labels and Excel made a series of it, delete that series and use the column
 * as the category names. A correct guess is left alone; a host without the
 * API keeps Excel's chart unchanged.
 */
async function useTextFirstColumnAsCategories(
  chart: Excel.Chart,
  dataRange: Excel.Range,
  ctx: Excel.RequestContext,
): Promise<void> {
  try {
    const firstColumn = dataRange.getColumn(0);
    firstColumn.load('values');
    chart.series.load('items/name');
    await ctx.sync();

    const values = (firstColumn.values as unknown[][]) ?? [];
    if (!isLabelColumn(values)) return;
    const header = String(values[0]?.[0] ?? '').trim();
    const misread = chart.series.items.find((series) => series.name?.trim() === header && header !== '');
    if (!misread) return;

    misread.delete();
    chart.axes.categoryAxis.setCategoryNames(
      firstColumn.getOffsetRange(1, 0).getResizedRange(-1, 0),
    );
    await ctx.sync();
  } catch (error) {
    console.warn('[Cellix] Could not correct chart categories; keeping Excel\'s layout:', error);
  }
}

export async function handleCreateChart(
  action: CreateChartAction,
  ctx: Excel.RequestContext,
): Promise<{ chartId: string }> {
  if (!action.sheetName.trim() || !action.sourceSheetName.trim() || !action.sourceRange.trim()) {
    throw new Error('CREATE_CHART requires sheetName, sourceSheetName, and sourceRange');
  }

  // Verify both sheets exist BEFORE queuing any dependent lookups (getRange/charts.add)
  // on the shared batch RequestContext. An unresolved getItem() for a missing sheet
  // does not throw immediately — it poisons the next ctx.sync(), which can fail the
  // ENTIRE batch (all other actions already applied in this run) with a generic
  // "The requested resource doesn't exist." Same class of bug worksheet.handler.ts's
  // WORKSHEET_ACTION_TYPES gate already fixed for worksheet-level actions — charts had
  // no equivalent guard.
  const targetSheetCheck = ctx.workbook.worksheets.getItemOrNullObject(action.sheetName);
  const sourceSheetCheck = ctx.workbook.worksheets.getItemOrNullObject(action.sourceSheetName);
  targetSheetCheck.load('isNullObject');
  sourceSheetCheck.load('isNullObject');
  await ctx.sync();

  if (targetSheetCheck.isNullObject) {
    throw new Error(`CREATE_CHART target sheet "${action.sheetName}" does not exist`);
  }
  if (sourceSheetCheck.isNullObject) {
    throw new Error(`CREATE_CHART source sheet "${action.sourceSheetName}" does not exist`);
  }

  const targetSheet = ctx.workbook.worksheets.getItem(action.sheetName);

  // A retried card must not stack a second copy of the chart it already made —
  // TASKS.md #319: a step that landed its chart and was then blocked later on
  // would add another chart on every retry. Same sheet + same title is the same
  // chart. An untitled chart has no identity to match on, so it is always created.
  const title = action.title?.trim();
  if (title) {
    const existing = targetSheet.charts;
    existing.load('items/name,items/title/text');
    await ctx.sync();
    const same = existing.items.find((item) => item.title?.text?.trim() === title);
    if (same) return { chartId: same.name };
  }

  const sourceSheet = ctx.workbook.worksheets.getItem(action.sourceSheetName);
  const dataRange = sourceSheet.getRange(stripSheetPrefix(action.sourceRange));

  const chart = targetSheet.charts.add(
    resolveChartType(action.chartType),
    dataRange,
    Excel.ChartSeriesBy.columns,
  );

  if (action.title?.trim()) {
    chart.title.text = action.title.trim();
    chart.title.visible = true;
  }

  await useTextFirstColumnAsCategories(chart, dataRange, ctx);

  const start = action.startCell ?? action.destCell ?? 'A1';
  const end = action.endCell ?? 'H16';
  chart.setPosition(start, end);

  await applyChartColorScheme(chart, ctx, action.colorScheme);

  // Capture Office.js-assigned name for follow-up UPDATE_CHART
  if (action.chartId?.trim()) {
    chart.name = action.chartId.trim();
  }
  chart.load('name');
  await ctx.sync();

  return { chartId: chart.name };
}

export async function handleUpdateChart(
  action: UpdateChartAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  if (!action.sheetName.trim() || !action.chartId.trim()) {
    throw new Error('UPDATE_CHART requires sheetName and chartId');
  }

  const sheetCheck = ctx.workbook.worksheets.getItemOrNullObject(action.sheetName);
  sheetCheck.load('isNullObject');
  await ctx.sync();
  if (sheetCheck.isNullObject) {
    throw new Error(`UPDATE_CHART sheet "${action.sheetName}" does not exist`);
  }

  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const chartCheck = sheet.charts.getItemOrNullObject(action.chartId);
  chartCheck.load('isNullObject');
  await ctx.sync();
  if (chartCheck.isNullObject) {
    throw new Error(`UPDATE_CHART chart "${action.chartId}" does not exist on "${action.sheetName}"`);
  }

  const chart = sheet.charts.getItem(action.chartId);

  if (action.chartType) {
    chart.chartType = resolveChartType(action.chartType);
  }

  await applyChartColorScheme(chart, ctx, action.colorScheme);

  await ctx.sync();
}

/** Revert-only inverse of CREATE_CHART (TASKS.md #15) — mirrors handleDeleteConditionalFormat's shape. */
export async function handleDeleteChart(
  action: DeleteChartAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheetCheck = ctx.workbook.worksheets.getItemOrNullObject(action.sheetName);
  sheetCheck.load('isNullObject');
  await ctx.sync();
  if (sheetCheck.isNullObject) {
    throw new Error(`DELETE_CHART sheet "${action.sheetName}" does not exist`);
  }

  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const chartCheck = sheet.charts.getItemOrNullObject(action.chartId);
  chartCheck.load('isNullObject');
  await ctx.sync();
  if (chartCheck.isNullObject) {
    throw new Error(`DELETE_CHART chart "${action.chartId}" does not exist on "${action.sheetName}"`);
  }

  const chart = sheet.charts.getItem(action.chartId);
  chart.delete();
  await ctx.sync();
}
