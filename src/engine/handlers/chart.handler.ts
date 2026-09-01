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
