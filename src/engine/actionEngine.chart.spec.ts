import { describe, expect, it, vi, beforeEach } from 'vitest';
import { RichActionEngine } from './actionEngine';
import type { RichAction } from '@/action.types';

/**
 * TASKS.md #15 — RichActionEngine.applyActions must collect the real
 * Office.js-assigned chart name a CREATE_CHART create just got (reported by
 * handleCreateChart's return value) and surface it on its own return value,
 * keyed by sheetName+sourceRange — this is what App.tsx threads into
 * POST /audit/apply/:changeSetId so the backend can patch it into
 * structuralOps before the change set is marked applied. Mirrors #40's
 * actionEngine.conditionalFormat.spec.ts exactly.
 */
describe('RichActionEngine.applyActions — createdChartIds (TASKS.md #15)', () => {
  function stubExcel() {
    const chart = {
      title: {},
      series: { load: vi.fn(), items: [] },
      name: '',
      setPosition: vi.fn(),
      load: vi.fn(),
    };
    const add = vi.fn(() => chart);
    const targetSheet = { charts: { add, getItem: vi.fn(() => chart) }, name: 'Dashboard' };
    const sourceSheet = { getRange: vi.fn(() => ({ load: vi.fn(), values: [] })), name: 'Sheet1' };
    const getItem = vi.fn((name: string) => (name === 'Dashboard' ? targetSheet : sourceSheet));
    // getItemOrNullObject existence-check pattern (see chart.handler.ts) — both
    // sheets used in these tests are real, so isNullObject is always false.
    const getItemOrNullObject = vi.fn(() => ({ isNullObject: false, load: vi.fn() }));
    const run = vi.fn(async (fn: (ctx: unknown) => Promise<void>) => {
      const ctx = {
        workbook: {
          worksheets: { getItem, getItemOrNullObject, getActiveWorksheet: vi.fn(() => targetSheet) },
        },
        sync: vi.fn(async () => {
          // Office.js assigns a real chart name on the first sync after .add() —
          // simulate that here since chart.name starts out empty.
          if (!chart.name) chart.name = 'Chart_real_id';
        }),
      };
      await fn(ctx);
    });
    vi.stubGlobal('Excel', {
      run,
      ChartType: { columnClustered: 'ColumnClustered', barClustered: 'BarClustered' },
      ChartSeriesBy: { columns: 'Columns' },
      ClearApplyTo: { contents: 'Contents' },
    });
    return { add, chart };
  }

  beforeEach(() => {
    vi.unstubAllGlobals();
  });

  it('reports the created chart name keyed by sheetName+sourceRange for a CREATE_CHART create', async () => {
    stubExcel();
    const engine = new RichActionEngine();
    const actions: RichAction[] = [
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A1:B10',
        chartType: 'column',
      } as RichAction,
    ];

    const result = await engine.applyActions(actions);

    expect(result.applied).toBe(1);
    expect(result.createdChartIds).toEqual([
      { sheetName: 'Dashboard', sourceRange: 'A1:B10', chartId: 'Chart_real_id' },
    ]);
  });

  it('omits createdChartIds entirely when no CREATE_CHART action ran', async () => {
    stubExcel();
    const engine = new RichActionEngine();
    const actions: RichAction[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 'x' } as RichAction,
    ];

    const result = await engine.applyActions(actions);
    expect(result.createdChartIds).toBeUndefined();
  });

  // TASKS.md #319 — a retried card that already landed its chart must not add
  // a second one.
  it('reuses an existing chart with the same title on the target sheet instead of adding another', async () => {
    const { add } = stubExcel();
    const existing = { name: 'Chart_existing', title: { text: 'Monthly Totals' } };
    // Give the target sheet's chart collection a loadable item list.
    const engine = new RichActionEngine();
    const run = (globalThis as unknown as { Excel: { run: ReturnType<typeof vi.fn> } }).Excel.run;
    const original = run.getMockImplementation()!;
    run.mockImplementation(async (fn: (ctx: unknown) => Promise<void>) =>
      original(async (ctx: unknown) => {
        const sheets = (ctx as { workbook: { worksheets: { getItem: (n: string) => { charts: Record<string, unknown> } } } })
          .workbook.worksheets;
        const charts = sheets.getItem('Dashboard').charts;
        charts.load = vi.fn();
        charts.items = [existing];
        await fn(ctx);
      }),
    );

    const result = await engine.applyActions([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Sheet1',
        sourceRange: 'A4:D16',
        chartType: 'column',
        title: 'Monthly Totals',
      } as RichAction,
    ]);

    expect(add).not.toHaveBeenCalled();
    expect(result.createdChartIds).toEqual([
      { sheetName: 'Dashboard', sourceRange: 'A4:D16', chartId: 'Chart_existing' },
    ]);
  });
});
