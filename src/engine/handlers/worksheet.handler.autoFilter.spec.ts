import { describe, expect, it, vi } from 'vitest';
import { handleWorksheetAction } from './worksheet.handler';
import type { RichAction } from '@/action.types';

// `/* global Excel */` is a type-only ambient declaration (office-js's real
// runtime injects the value inside Excel itself) — no other handler spec has
// exercised a branch that reads an Excel.* enum at runtime before, so nothing
// stubs it yet. Values match office.d.ts's real string literals exactly
// (Excel.FilterOn.custom = "Custom", .values = "Values") so a real Office.js
// payload comparison would still match this test's expectations.
(globalThis as unknown as { Excel: typeof Excel }).Excel = {
  FilterOn: { custom: 'Custom', values: 'Values' },
} as unknown as typeof Excel;

/**
 * TASKS.md #221 — "Show only rows where the taxable amount is above 1 lakh"
 * (guide T2.2) used to produce a bare AUTO_FILTER with no criteria: the
 * dropdown arrows appeared, but every row stayed visible, because `filter`
 * was silently dropped on the way from the backend to this handler (three
 * separate drop points — legacyConverter.ts, the server normalizer, and this
 * handler itself never read it in the first place).
 */
describe('handleWorksheetAction — AUTO_FILTER with a real condition (#221)', () => {
  const HEADERS = ['Invoice No', 'Invoice Date', 'Supplier Name', 'GSTIN', 'Taxable Amount'];

  function makeCtx(values: unknown[][]) {
    const apply = vi.fn();
    const range = { load: vi.fn(), values };
    const getRange = vi.fn(() => range);
    const sheet = { autoFilter: { apply }, getRange };
    const getItem = vi.fn(() => sheet);
    return {
      ctx: {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn(() => sheet) } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext,
      apply,
      getRange,
    };
  }

  it('applies a real greaterThan criterion on the resolved column', async () => {
    const { ctx, apply, getRange } = makeCtx([HEADERS, ['INV/001', 45383, 'ABC', 'X', 118500]]);
    const action = {
      type: 'AUTO_FILTER',
      sheetName: 'Purchase Register',
      range: 'A1:E31',
      filter: { column: 'Taxable Amount', operator: 'greaterThan', value: 100000 },
    } as unknown as RichAction;

    const handled = await handleWorksheetAction(action, ctx);

    expect(handled).toBe(true);
    expect(getRange).toHaveBeenCalledWith('A1:E31');
    // Column 4 (0-based) = "Taxable Amount"; a real Custom criterion, not a bare apply().
    expect(apply).toHaveBeenCalledWith('A1:E31', 4, {
      filterOn: Excel.FilterOn.custom,
      criterion1: '>100000',
    });
  });

  it('resolves the column by header name regardless of position', async () => {
    const { ctx, apply } = makeCtx([HEADERS, []]);
    const action = {
      type: 'AUTO_FILTER',
      sheetName: 'S',
      range: 'A1:E2',
      filter: { column: 'GSTIN', operator: 'equals', value: '' },
    } as unknown as RichAction;

    await handleWorksheetAction(action, ctx);

    expect(apply).toHaveBeenCalledWith('A1:E2', 3, {
      filterOn: Excel.FilterOn.values,
      values: [''],
    });
  });

  it('maps every RangeFilterSpec operator Excel AutoFilter can express', async () => {
    const cases: Array<[string, unknown, object]> = [
      ['equals', 'ABC Traders', { filterOn: Excel.FilterOn.values, values: ['ABC Traders'] }],
      ['notEquals', 0, { filterOn: Excel.FilterOn.custom, criterion1: '<>0' }],
      ['contains', 'Credit', { filterOn: Excel.FilterOn.custom, criterion1: '*Credit*' }],
      ['greaterThan', 100000, { filterOn: Excel.FilterOn.custom, criterion1: '>100000' }],
      ['lessThan', 50000, { filterOn: Excel.FilterOn.custom, criterion1: '<50000' }],
    ];

    for (const [operator, value, expected] of cases) {
      const { ctx, apply } = makeCtx([HEADERS, []]);
      const action = {
        type: 'AUTO_FILTER',
        sheetName: 'S',
        range: 'A1:E2',
        filter: { column: 'Taxable Amount', operator, value },
      } as unknown as RichAction;

      await handleWorksheetAction(action, ctx);

      expect(apply).toHaveBeenCalledWith('A1:E2', 4, expected);
    }
  });

  it('falls back to plain dropdown arrows when the operator has no AutoFilter equivalent', async () => {
    const { ctx, apply } = makeCtx([HEADERS, []]);
    const action = {
      type: 'AUTO_FILTER',
      sheetName: 'S',
      range: 'A1:E2',
      filter: { column: 'GSTIN', operator: 'lengthEquals', value: 15 },
    } as unknown as RichAction;

    await handleWorksheetAction(action, ctx);

    // Called with ONLY the range — no columnIndex/criteria args.
    expect(apply).toHaveBeenCalledWith('A1:E2');
  });

  it('falls back to plain dropdowns when no filter is given at all (guide T2.1)', async () => {
    const { ctx, apply, getRange } = makeCtx([HEADERS, []]);
    const action = {
      type: 'AUTO_FILTER',
      sheetName: 'S',
      range: 'A1:E31',
    } as unknown as RichAction;

    await handleWorksheetAction(action, ctx);

    expect(apply).toHaveBeenCalledWith('A1:E31');
    // No column resolution needed — the range was never read for headers.
    expect(getRange).not.toHaveBeenCalled();
  });

  it('falls back to plain dropdowns when the named column cannot be resolved', async () => {
    const { ctx, apply } = makeCtx([HEADERS, []]);
    const action = {
      type: 'AUTO_FILTER',
      sheetName: 'S',
      range: 'A1:E2',
      filter: { column: 'Nonexistent Column', operator: 'equals', value: 'x' },
    } as unknown as RichAction;

    await handleWorksheetAction(action, ctx);

    expect(apply).toHaveBeenCalledWith('A1:E2');
  });
});
