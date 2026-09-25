import { afterEach, describe, expect, it, vi } from 'vitest';
import { handleCreateChart, isLabelColumn } from './chart.handler';

/**
 * TASKS.md #327 — a live dashboard chart showed "Month" as a series and an
 * axis of 1…12. When Excel turns a label column into a series, the handler
 * deletes that series and uses the column as category names.
 */
function stub(seriesNames: string[], firstColumn: unknown[][]) {
  const series = seriesNames.map((name) => ({ name, delete: vi.fn() }));
  const categoryRange = {};
  const column = {
    load: vi.fn(),
    values: firstColumn,
    getOffsetRange: vi.fn(() => ({ getResizedRange: vi.fn(() => categoryRange) })),
  };
  const setCategoryNames = vi.fn();
  const chart = {
    title: {},
    name: 'Chart1',
    load: vi.fn(),
    setPosition: vi.fn(),
    series: { load: vi.fn(), items: series },
    axes: { categoryAxis: { setCategoryNames } },
  };
  const sheet = {
    charts: { add: vi.fn(() => chart), load: vi.fn(), items: [] },
    getRange: vi.fn(() => ({ getColumn: () => column })),
  };
  const ctx = {
    workbook: {
      worksheets: {
        getItem: vi.fn(() => sheet),
        getItemOrNullObject: vi.fn(() => ({ load: vi.fn(), isNullObject: false })),
      },
    },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal('Excel', {
    ChartType: { columnClustered: 'ColumnClustered' },
    ChartSeriesBy: { columns: 'Columns' },
  });
  return { ctx: ctx as unknown as Excel.RequestContext, series, setCategoryNames, categoryRange };
}

const action = {
  type: 'CREATE_CHART',
  sheetName: 'Main',
  sourceSheetName: 'Main',
  sourceRange: 'A8:D20',
  chartType: 'ColumnClustered',
} as never;

const MONTH_COLUMN = [['Month'], ['January'], ['February'], ['March']];

describe('CREATE_CHART — a label column is the category axis, not a series (TASKS.md #327)', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('deletes the "Month" series Excel made and uses the months as category names', async () => {
    const { ctx, series, setCategoryNames, categoryRange } = stub(
      ['Month', 'Total Amount', 'Paid', 'Pending'],
      MONTH_COLUMN,
    );
    await handleCreateChart(action, ctx);
    expect(series[0].delete).toHaveBeenCalled();
    expect(series.slice(1).every((s) => s.delete.mock.calls.length === 0)).toBe(true);
    expect(setCategoryNames).toHaveBeenCalledWith(categoryRange);
  });

  it('leaves a chart alone when Excel already used the labels as categories', async () => {
    const { ctx, series, setCategoryNames } = stub(['Total Amount', 'Paid', 'Pending'], MONTH_COLUMN);
    await handleCreateChart(action, ctx);
    expect(series.every((s) => s.delete.mock.calls.length === 0)).toBe(true);
    expect(setCategoryNames).not.toHaveBeenCalled();
  });

  it('only treats text as labels — a numeric first column is data', () => {
    expect(isLabelColumn(MONTH_COLUMN)).toBe(true);
    expect(isLabelColumn([['Year'], [2024], [2025]])).toBe(false);
    expect(isLabelColumn([['Year'], ['2024'], ['2025']])).toBe(false);
    expect(isLabelColumn([['Month']])).toBe(false);
  });
});
