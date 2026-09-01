import { describe, expect, it, vi } from 'vitest';
import { handleConditionalFormat, handleDeleteConditionalFormat } from './range.handler';
import type { ConditionalFormatAction, DeleteConditionalFormatAction } from '@/action.types';

/**
 * TASKS.md #34 — CONDITIONAL_FORMAT must dispatch through Office.js's own
 * Range.conditionalFormats.add('CellValue') API so the rule re-evaluates live
 * in Excel, never a computed one-shot fill. These tests assert the Office.js
 * call shape directly (operator/formula1/formula2/format), since a real
 * re-evaluation check needs live Excel and can't run in this environment.
 */
describe('handleConditionalFormat', () => {
  function makeConditionalFormat() {
    const cellValueFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
    const cellValue: { rule?: unknown; format: typeof cellValueFormat } = { format: cellValueFormat };
    return { cellValue, cellValueFormat };
  }

  function makeCtx() {
    const { cellValue, cellValueFormat } = makeConditionalFormat();
    const add = vi.fn(() => ({ cellValue, id: 'cf-mock-id', load: vi.fn() }));
    const range = { conditionalFormats: { add } };
    const getRange = vi.fn(() => range);
    const getItem = vi.fn(() => ({ getRange }));
    const ctx = {
      workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
      sync: vi.fn(async () => undefined),
    } as unknown as Excel.RequestContext;
    return { ctx, getItem, getRange, add, cellValue, cellValueFormat };
  }

  it('adds a CellValue rule with the correct operator/formula and resolves the target sheet+range', async () => {
    const { ctx, getItem, getRange, add, cellValue, cellValueFormat } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: 'J2:J51',
      rule: {
        kind: 'cellValue',
        operator: 'greaterThan',
        value: 1000,
        format: { fillColor: '#FFC7CE' },
      },
    };

    await handleConditionalFormat(action, ctx);

    expect(getItem).toHaveBeenCalledWith('Purchase Register');
    expect(getRange).toHaveBeenCalledWith('J2:J51');
    expect(add).toHaveBeenCalledWith('CellValue');
    expect(cellValue.rule).toEqual({ operator: 'GreaterThan', formula1: '1000' });
    expect(cellValueFormat.fill.color).toBe('#FFC7CE');
  });

  it('strips a sheet-prefixed range before calling getRange', async () => {
    const { ctx, getRange } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: "'Purchase Register'!J2:J51",
      rule: { kind: 'cellValue', operator: 'lessThan', value: 0, format: { fillColor: '#FFC7CE' } },
    };

    await handleConditionalFormat(action, ctx);
    expect(getRange).toHaveBeenCalledWith('J2:J51');
  });

  it('sets formula2 only when the operator is between/notBetween', async () => {
    const { ctx, cellValue } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: 'J2:J51',
      rule: {
        kind: 'cellValue',
        operator: 'between',
        value: 100,
        value2: 500,
        format: { fillColor: '#FFF2CC' },
      },
    };

    await handleConditionalFormat(action, ctx);
    expect(cellValue.rule).toEqual({ operator: 'Between', formula1: '100', formula2: '500' });
  });

  it('quotes a text rule value as a formula literal', async () => {
    const { ctx, cellValue } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: 'K2:K51',
      rule: { kind: 'cellValue', operator: 'equalTo', value: 'Overdue', format: { fillColor: '#FFC7CE' } },
    };

    await handleConditionalFormat(action, ctx);
    expect(cellValue.rule).toEqual({ operator: 'EqualTo', formula1: '"Overdue"' });
  });

  it('applies bold/fontColor alongside fillColor via the conditional-range format object', async () => {
    const { ctx, cellValueFormat } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: 'J2:J51',
      rule: {
        kind: 'cellValue',
        operator: 'greaterThan',
        value: 1000,
        format: { fillColor: '#FFC7CE', bold: true, fontColor: '#9C0006' },
      },
    };

    await handleConditionalFormat(action, ctx);
    expect(cellValueFormat.fill.color).toBe('#FFC7CE');
    expect((cellValueFormat.font as Record<string, unknown>).bold).toBe(true);
    expect((cellValueFormat.font as Record<string, unknown>).color).toBe('#9C0006');
  });

  describe('formula-kind rule (TASKS.md #35 — cross-column comparison)', () => {
    function makeCustomCtx() {
      const customFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
      const custom: { rule: { formula?: string }; format: typeof customFormat } = {
        rule: {},
        format: customFormat,
      };
      const add = vi.fn(() => ({ custom, id: 'cf-mock-id', load: vi.fn() }));
      const range = { conditionalFormats: { add } };
      const getRange = vi.fn(() => range);
      const getItem = vi.fn(() => ({ getRange }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;
      return { ctx, getItem, getRange, add, custom, customFormat };
    }

    it('adds a Custom rule and sets the formula verbatim, reproducing VISION.md\'s own example', async () => {
      const { ctx, add, custom, customFormat } = makeCustomCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: {
          kind: 'formula',
          formula: '=$C2<$B2*0.9',
          format: { fillColor: '#FFC7CE' },
        },
      };

      await handleConditionalFormat(action, ctx);

      expect(add).toHaveBeenCalledWith('Custom');
      expect(custom.rule.formula).toBe('=$C2<$B2*0.9');
      expect(customFormat.fill.color).toBe('#FFC7CE');
    });

    it('never uses the CellValue path for a formula rule', async () => {
      const cellValue = { rule: undefined, format: { font: {}, fill: {} } };
      const add = vi.fn(() => ({
        cellValue,
        custom: { rule: {}, format: { font: {}, fill: {} } },
        id: 'cf-mock-id',
        load: vi.fn(),
      }));
      const range = { conditionalFormats: { add } };
      const getItem = vi.fn(() => ({ getRange: vi.fn(() => range) }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;

      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: { kind: 'formula', formula: '=$C2<$B2*0.9', format: { fillColor: '#FFC7CE' } },
      };

      await handleConditionalFormat(action, ctx);
      expect(add).toHaveBeenCalledWith('Custom');
      expect(add).not.toHaveBeenCalledWith('CellValue');
    });
  });

  describe('topBottom-kind rule (TASKS.md #36 — rank-based highlight)', () => {
    function makeTopBottomCtx() {
      const topBottomFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
      const topBottom: { rule?: { type?: string; rank?: number }; format: typeof topBottomFormat } = {
        format: topBottomFormat,
      };
      const add = vi.fn(() => ({ topBottom, id: 'cf-mock-id', load: vi.fn() }));
      const range = { conditionalFormats: { add } };
      const getRange = vi.fn(() => range);
      const getItem = vi.fn(() => ({ getRange }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;
      return { ctx, getItem, getRange, add, topBottom, topBottomFormat };
    }

    it('adds a TopBottom rule with TopItems for a top-N-by-count request', async () => {
      const { ctx, add, topBottom, topBottomFormat } = makeTopBottomCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 5, format: { fillColor: '#C6EFCE' } },
      };

      await handleConditionalFormat(action, ctx);

      expect(add).toHaveBeenCalledWith('TopBottom');
      expect(topBottom.rule).toEqual({ type: 'TopItems', rank: 5 });
      expect(topBottomFormat.fill.color).toBe('#C6EFCE');
    });

    it('maps side/isPercent to BottomPercent for a bottom-N% request', async () => {
      const { ctx, topBottom } = makeTopBottomCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'topBottom', side: 'bottom', rank: 10, isPercent: true, format: { fillColor: '#FFC7CE' } },
      };

      await handleConditionalFormat(action, ctx);
      expect(topBottom.rule).toEqual({ type: 'BottomPercent', rank: 10 });
    });

    it('maps top+isPercent to TopPercent', async () => {
      const { ctx, topBottom } = makeTopBottomCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'topBottom', side: 'top', rank: 20, isPercent: true, format: { fillColor: '#C6EFCE' } },
      };

      await handleConditionalFormat(action, ctx);
      expect(topBottom.rule).toEqual({ type: 'TopPercent', rank: 20 });
    });

    it('never uses the CellValue or Custom path for a topBottom rule', async () => {
      const cellValue = { rule: undefined, format: { font: {}, fill: {} } };
      const custom = { rule: {}, format: { font: {}, fill: {} } };
      const topBottom = { rule: undefined, format: { font: {}, fill: {} } };
      const add = vi.fn(() => ({ cellValue, custom, topBottom, id: 'cf-mock-id', load: vi.fn() }));
      const range = { conditionalFormats: { add } };
      const getItem = vi.fn(() => ({ getRange: vi.fn(() => range) }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;

      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 5, format: { fillColor: '#C6EFCE' } },
      };

      await handleConditionalFormat(action, ctx);
      expect(add).toHaveBeenCalledWith('TopBottom');
      expect(add).not.toHaveBeenCalledWith('CellValue');
      expect(add).not.toHaveBeenCalledWith('Custom');
    });
  });

  describe('colorScale-kind rule (TASKS.md #37 — gradient/heat-map highlight)', () => {
    function makeColorScaleCtx() {
      const colorScale: { criteria?: Excel.ConditionalColorScaleCriteria } = {};
      const add = vi.fn(() => ({ colorScale, id: 'cf-mock-id', load: vi.fn() }));
      const range = { conditionalFormats: { add } };
      const getRange = vi.fn(() => range);
      const getItem = vi.fn(() => ({ getRange }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;
      return { ctx, getItem, getRange, add, colorScale };
    }

    it('adds a ColorScale rule with a minimum/maximum for a 2-color scale', async () => {
      const { ctx, add, colorScale } = makeColorScaleCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      };

      await handleConditionalFormat(action, ctx);

      expect(add).toHaveBeenCalledWith('ColorScale');
      expect(colorScale.criteria).toEqual({
        minimum: { type: 'LowestValue', color: '#F8696B' },
        maximum: { type: 'HighestValue', color: '#63BE7B' },
      });
    });

    it('adds a midpoint for a 3-color scale', async () => {
      const { ctx, colorScale } = makeColorScaleCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
      };

      await handleConditionalFormat(action, ctx);

      expect(colorScale.criteria).toEqual({
        minimum: { type: 'LowestValue', color: '#F8696B' },
        midpoint: { type: 'Percentile', formula: '50', color: '#FFEB84' },
        maximum: { type: 'HighestValue', color: '#63BE7B' },
      });
    });

    it('never uses the CellValue/Custom/TopBottom path for a colorScale rule', async () => {
      const cellValue = { rule: undefined, format: { font: {}, fill: {} } };
      const custom = { rule: {}, format: { font: {}, fill: {} } };
      const topBottom = { rule: undefined, format: { font: {}, fill: {} } };
      const colorScale = {};
      const add = vi.fn(() => ({ cellValue, custom, topBottom, colorScale, id: 'cf-mock-id', load: vi.fn() }));
      const range = { conditionalFormats: { add } };
      const getItem = vi.fn(() => ({ getRange: vi.fn(() => range) }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;

      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      };

      await handleConditionalFormat(action, ctx);
      expect(add).toHaveBeenCalledWith('ColorScale');
      expect(add).not.toHaveBeenCalledWith('CellValue');
      expect(add).not.toHaveBeenCalledWith('Custom');
      expect(add).not.toHaveBeenCalledWith('TopBottom');
    });
  });

  describe('existingRuleId — modify an existing rule in place (TASKS.md #38)', () => {
    function makeModifyCtx() {
      const cellValueFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
      const cellValue: { rule?: unknown; format: typeof cellValueFormat } = { format: cellValueFormat };
      const getItemOnCollection = vi.fn(() => ({ cellValue }));
      const conditionalFormats = { getItem: getItemOnCollection, add: vi.fn() };
      const usedRange = { conditionalFormats };
      const getUsedRange = vi.fn(() => usedRange);
      const getRange = vi.fn();
      const getItem = vi.fn(() => ({ getUsedRange, getRange }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;
      return { ctx, getUsedRange, getRange, getItemOnCollection, conditionalFormats, cellValue, cellValueFormat };
    }

    it('resolves the existing rule via getUsedRange().conditionalFormats.getItem(id) and mutates it, never calling add()', async () => {
      const { ctx, getUsedRange, getRange, getItemOnCollection, conditionalFormats, cellValue, cellValueFormat } =
        makeModifyCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        existingRuleId: 'cf-existing-1',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
      };

      await handleConditionalFormat(action, ctx);

      expect(getUsedRange).toHaveBeenCalledTimes(1);
      expect(getItemOnCollection).toHaveBeenCalledWith('cf-existing-1');
      expect(getRange).not.toHaveBeenCalled();
      expect(conditionalFormats.add).not.toHaveBeenCalled();
      expect(cellValue.rule).toEqual({ operator: 'GreaterThan', formula1: '1500' });
      expect(cellValueFormat.fill.color).toBe('#FFC7CE');
    });

    it('mutates a topBottom rule retrieved by id the same way a create would', async () => {
      const topBottomFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
      const topBottom: { rule?: unknown; format: typeof topBottomFormat } = { format: topBottomFormat };
      const getItemOnCollection = vi.fn(() => ({ topBottom }));
      const usedRange = { conditionalFormats: { getItem: getItemOnCollection, add: vi.fn() } };
      const getItem = vi.fn(() => ({ getUsedRange: vi.fn(() => usedRange), getRange: vi.fn() }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;

      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        existingRuleId: 'cf-existing-2',
        rule: { kind: 'topBottom', side: 'top', rank: 10, format: { fillColor: '#C6EFCE' } },
      };

      await handleConditionalFormat(action, ctx);
      expect(getItemOnCollection).toHaveBeenCalledWith('cf-existing-2');
      expect(topBottom.rule).toEqual({ type: 'TopItems', rank: 10 });
    });
  });

  it('syncs the context after applying the rule', async () => {
    const { ctx } = makeCtx();
    const action: ConditionalFormatAction = {
      type: 'CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      range: 'J2:J51',
      rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
    };

    await handleConditionalFormat(action, ctx);
    expect(ctx.sync).toHaveBeenCalledTimes(1);
  });

  describe('apply-time id capture (TASKS.md #40)', () => {
    it('loads and returns the real Excel-assigned id on a plain create', async () => {
      const { ctx, add } = makeCtx();
      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
      };

      const result = await handleConditionalFormat(action, ctx);

      expect(result).toEqual({ createdConditionalFormatId: 'cf-mock-id' });
      const created = add.mock.results[0]?.value as { load: ReturnType<typeof vi.fn> };
      expect(created.load).toHaveBeenCalledWith('id');
    });

    it('does not load or return an id when modifying via existingRuleId', async () => {
      const cellValueFormat = { font: {} as Record<string, unknown>, fill: {} as Record<string, unknown> };
      const cellValue: { rule?: unknown; format: typeof cellValueFormat } = { format: cellValueFormat };
      const load = vi.fn();
      const getItemOnCollection = vi.fn(() => ({ cellValue, load, id: 'cf-existing-1' }));
      const usedRange = { conditionalFormats: { getItem: getItemOnCollection } };
      const getItem = vi.fn(() => ({ getUsedRange: vi.fn(() => usedRange), getRange: vi.fn() }));
      const ctx = {
        workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
        sync: vi.fn(async () => undefined),
      } as unknown as Excel.RequestContext;

      const action: ConditionalFormatAction = {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        existingRuleId: 'cf-existing-1',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
      };

      const result = await handleConditionalFormat(action, ctx);

      expect(result).toBeUndefined();
      expect(load).not.toHaveBeenCalled();
    });
  });
});

describe('handleDeleteConditionalFormat (TASKS.md #40 — revert-only)', () => {
  it('resolves the rule via getUsedRange().conditionalFormats.getItem(ruleId) and deletes it', async () => {
    const deleteFn = vi.fn();
    const getItemOnCollection = vi.fn(() => ({ delete: deleteFn }));
    const usedRange = { conditionalFormats: { getItem: getItemOnCollection } };
    const getUsedRange = vi.fn(() => usedRange);
    const getItem = vi.fn(() => ({ getUsedRange }));
    const ctx = {
      workbook: { worksheets: { getItem, getActiveWorksheet: vi.fn() } },
      sync: vi.fn(async () => undefined),
    } as unknown as Excel.RequestContext;

    const action: DeleteConditionalFormatAction = {
      type: 'DELETE_CONDITIONAL_FORMAT',
      sheetName: 'Purchase Register',
      ruleId: 'cf-to-delete',
    };

    await handleDeleteConditionalFormat(action, ctx);

    expect(getItem).toHaveBeenCalledWith('Purchase Register');
    expect(getItemOnCollection).toHaveBeenCalledWith('cf-to-delete');
    expect(deleteFn).toHaveBeenCalledTimes(1);
    expect(ctx.sync).toHaveBeenCalledTimes(1);
  });
});
