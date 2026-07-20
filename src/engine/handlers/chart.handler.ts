import { CreateChartAction, UpdateChartAction } from '@/action.types';
import { stripSheetPrefix } from '../addressUtils';

/* global Excel */

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

  const sheet = ctx.workbook.worksheets.getItem(action.sheetName);
  const chart = sheet.charts.getItem(action.chartId);

  if (action.chartType) {
    chart.chartType = resolveChartType(action.chartType);
  }

  if (action.colorScheme && action.colorScheme !== 'default') {
    const colorMap: Record<string, string> = {
      blue: '#4472C4',
      grey: '#7F7F7F',
      blueGrey: '#5B9BD5',
    };
    const fill = colorMap[action.colorScheme];
    if (fill) {
      chart.series.load('items');
      await ctx.sync();
      for (const series of chart.series.items) {
        series.format.fill.setSolidColor(fill);
      }
    }
  }

  await ctx.sync();
}
