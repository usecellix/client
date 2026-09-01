import { describe, expect, it, vi, beforeEach } from 'vitest';
import { RichActionEngine } from './actionEngine';
import type { RichAction } from '@/action.types';

/**
 * TASKS.md #40 — RichActionEngine.applyActions must collect the real
 * Excel-assigned id a CONDITIONAL_FORMAT create just got (reported by
 * handleConditionalFormat's return value) and surface it on its own return
 * value, keyed by sheetName+range — this is what App.tsx threads into
 * POST /audit/apply/:changeSetId so the backend can patch it into
 * structuralOps before the change set is marked applied.
 */
describe('RichActionEngine.applyActions — createdConditionalFormatIds (TASKS.md #40)', () => {
  function stubExcel() {
    const conditionalFormat = { cellValue: { format: { font: {}, fill: {} } } as { rule?: unknown; format: unknown }, id: 'cf-real-id', load: vi.fn() };
    const add = vi.fn(() => conditionalFormat);
    const range = { conditionalFormats: { add }, load: vi.fn() };
    const getRange = vi.fn(() => range);
    const sheet = { getRange, name: 'Sheet1' };
    const getItem = vi.fn(() => sheet);
    const run = vi.fn(async (fn: (ctx: unknown) => Promise<void>) => {
      const ctx = {
        workbook: {
          worksheets: { getItem, getActiveWorksheet: vi.fn(() => sheet) },
        },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
    });
    vi.stubGlobal('Excel', { run, ClearApplyTo: { contents: 'Contents' } });
    return { add, conditionalFormat };
  }

  beforeEach(() => {
    vi.unstubAllGlobals();
  });

  it('reports the created id keyed by sheetName+range for a CONDITIONAL_FORMAT create', async () => {
    stubExcel();
    const engine = new RichActionEngine();
    const actions: RichAction[] = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B50',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
      },
    ];

    const result = await engine.applyActions(actions);

    expect(result.applied).toBe(1);
    expect(result.createdConditionalFormatIds).toEqual([
      { sheetName: 'Sheet1', range: 'B2:B50', ruleId: 'cf-real-id' },
    ]);
  });

  it('omits createdConditionalFormatIds entirely when no CONDITIONAL_FORMAT action ran', async () => {
    stubExcel();
    const engine = new RichActionEngine();
    const actions: RichAction[] = [
      { type: 'SET_CELL', sheetName: 'Sheet1', address: 'A1', value: 'x' } as RichAction,
    ];

    const result = await engine.applyActions(actions);
    expect(result.createdConditionalFormatIds).toBeUndefined();
  });

  it('does not report an id for a MODIFY (existingRuleId set) — handleConditionalFormat returns nothing to collect', async () => {
    const conditionalFormat = { cellValue: { format: { font: {}, fill: {} } } as { rule?: unknown; format: unknown }, load: vi.fn() };
    const getItemOnCollection = vi.fn(() => conditionalFormat);
    const usedRange = { conditionalFormats: { getItem: getItemOnCollection } };
    const sheet = { getUsedRange: vi.fn(() => usedRange), getRange: vi.fn(), name: 'Sheet1' };
    const getItem = vi.fn(() => sheet);
    const run = vi.fn(async (fn: (ctx: unknown) => Promise<void>) => {
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn(() => sheet) } },
        sync: vi.fn(async () => undefined),
      };
      await fn(ctx);
    });
    vi.stubGlobal('Excel', { run, ClearApplyTo: { contents: 'Contents' } });

    const engine = new RichActionEngine();
    const actions: RichAction[] = [
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Sheet1',
        range: 'B2:B50',
        existingRuleId: 'cf-existing',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
      },
    ];

    const result = await engine.applyActions(actions);
    expect(result.applied).toBe(1);
    expect(result.createdConditionalFormatIds).toBeUndefined();
  });
});
