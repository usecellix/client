import { describe, expect, it, vi } from 'vitest';
import { extractConditionalFormats, summarizeConditionalFormat } from './workbookReader';

/**
 * TASKS.md #38 — reading back existing conditional-format rules via Office.js
 * so a follow-up request can target/modify an existing rule (by id) instead
 * of stacking a duplicate CONDITIONAL_FORMAT on top. These tests mock the
 * Office.js call chain directly, the same convention
 * `range.handler.conditionalFormat.spec.ts` already uses for the dispatch
 * side — a real add-in smoke test against live Excel is a separate,
 * unverified caveat noted in TASKS.md, same class as #34/#65.
 */
describe('extractConditionalFormats', () => {
  function makeCf(overrides: Record<string, unknown> & { type: string }) {
    const base: Record<string, unknown> = {
      id: 'cf-1',
      getRangeOrNullObject: vi.fn(() => ({
        address: 'B2:B50',
        isNullObject: false,
        load: vi.fn(),
      })),
    };
    return { ...base, ...overrides } as unknown as Excel.ConditionalFormat;
  }

  function makeCtx(items: Excel.ConditionalFormat[]) {
    const cfs = { items, load: vi.fn() };
    const usedRange = { conditionalFormats: cfs } as unknown as Excel.Range;
    const ctx = { sync: vi.fn(async () => undefined) } as unknown as Excel.RequestContext;
    return { ctx, usedRange, cfs };
  }

  it('reads a cellValue rule and produces a compact summary', async () => {
    const cf = makeCf({
      type: 'CellValue',
      cellValue: { rule: { operator: 'GreaterThan', formula1: '1000' }, load: vi.fn() } as never,
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Purchase Register', usedRange);

    expect(result).toEqual([
      {
        id: 'cf-1',
        sheetName: 'Purchase Register',
        range: 'B2:B50',
        ruleKind: 'cellValue',
        summary: 'GreaterThan 1000',
      },
    ]);
  });

  it('reads a between-operator cellValue rule with both formulas in the summary', async () => {
    const cf = makeCf({
      type: 'CellValue',
      cellValue: {
        rule: { operator: 'Between', formula1: '100', formula2: '500' },
        load: vi.fn(),
      } as never,
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Sheet1', usedRange);
    expect(result[0]?.summary).toBe('Between 100..500');
  });

  it('reads a formula (Custom) rule verbatim', async () => {
    const cf = makeCf({
      type: 'Custom',
      custom: { rule: { formula: '=$C2<$B2*0.9' }, load: vi.fn() } as never,
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Regional Revenue', usedRange);
    expect(result[0]).toEqual(
      expect.objectContaining({ ruleKind: 'formula', summary: '=$C2<$B2*0.9' }),
    );
  });

  it('reads a topBottom rule', async () => {
    const cf = makeCf({
      type: 'TopBottom',
      topBottom: { rule: { type: 'TopItems', rank: 5 }, load: vi.fn() } as never,
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Suppliers', usedRange);
    expect(result[0]).toEqual(expect.objectContaining({ ruleKind: 'topBottom', summary: 'TopItems 5' }));
  });

  it('reads a colorScale rule and joins the stop colors', async () => {
    const cf = makeCf({
      type: 'ColorScale',
      colorScale: {
        criteria: {
          minimum: { color: '#F8696B' },
          midpoint: { color: '#FFEB84' },
          maximum: { color: '#63BE7B' },
        },
        load: vi.fn(),
      } as never,
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Scores', usedRange);
    expect(result[0]).toEqual(
      expect.objectContaining({
        ruleKind: 'colorScale',
        summary: 'scale #F8696B -> #FFEB84 -> #63BE7B',
      }),
    );
  });

  it('skips a multi-range rule whose getRangeOrNullObject resolves to a null object', async () => {
    const cf = makeCf({
      type: 'CellValue',
      cellValue: { rule: { operator: 'GreaterThan', formula1: '1000' }, load: vi.fn() } as never,
      getRangeOrNullObject: vi.fn(() => ({ address: undefined, isNullObject: true, load: vi.fn() })),
    });
    const { ctx, usedRange } = makeCtx([cf]);

    const result = await extractConditionalFormats(ctx, 'Sheet1', usedRange);
    expect(result).toEqual([]);
  });

  it('returns an empty array and never throws when the read fails entirely', async () => {
    const usedRange = {
      conditionalFormats: {
        load: () => {
          throw new Error('boom');
        },
      },
    } as unknown as Excel.Range;
    const ctx = { sync: vi.fn(async () => undefined) } as unknown as Excel.RequestContext;

    const result = await extractConditionalFormats(ctx, 'Sheet1', usedRange);
    expect(result).toEqual([]);
  });

  it('returns an empty array immediately when there are no rules on the range', async () => {
    const { ctx, usedRange } = makeCtx([]);
    const result = await extractConditionalFormats(ctx, 'Sheet1', usedRange);
    expect(result).toEqual([]);
  });
});

describe('summarizeConditionalFormat', () => {
  it('falls back to "other" and never throws when a sub-object access itself throws', () => {
    const cf = {
      type: 'DataBar',
      get cellValue(): never {
        throw new Error('not a CellValue rule');
      },
    } as unknown as Excel.ConditionalFormat;

    expect(summarizeConditionalFormat(cf)).toEqual({ ruleKind: 'other', summary: 'DataBar rule' });
  });
});
