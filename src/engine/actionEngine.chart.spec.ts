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
    const run = vi.fn(async (fn: (ctx: unknown) => Promise<void>) => {
      const ctx = {
        workbook: {
          worksheets: { getItem, getActiveWorksheet: vi.fn(() => targetSheet) },
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
});
