import { describe, expect, it } from 'vitest';
import { partitionActions } from './actionNormalizer';
import type { SheetAction } from '@/types/sheet-actions';

describe('actionNormalizer', () => {
  it('routes DELETE_SHEET to the rich engine', () => {
    const { rich, unsupported } = partitionActions([
      { type: 'DELETE_SHEET', sheetName: 'Cellix' },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Cellix' }]);
  });

  it('routes CREATE_SHEET to ADD_SHEET in the rich engine', () => {
    const { rich } = partitionActions([{ type: 'CREATE_SHEET', sheetName: 'Summary' }]);
    expect(rich).toEqual([{ type: 'ADD_SHEET', name: 'Summary' }]);
  });

  it('normalizes legacy CREATE_TABLE name and defaults headers to true', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CREATE_TABLE',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        name: 'PurchaseTable',
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CREATE_TABLE',
        sheetName: 'Purchase Register',
        range: 'A1:L51',
        tableName: 'PurchaseTable',
        hasHeaders: true,
      },
    ]);
  });

  it('routes CONDITIONAL_FORMAT to the rich engine (TASKS.md #34)', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1000, format: { fillColor: '#FFC7CE' } },
      },
    ]);
  });

  it("routes a formula-kind CONDITIONAL_FORMAT to the rich engine (TASKS.md #35 — VISION.md's own example)", () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: { kind: 'formula', formula: '=$C2<$B2*0.9', format: { fillColor: '#FFC7CE' } },
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: { kind: 'formula', formula: '=$C2<$B2*0.9', format: { fillColor: '#FFC7CE' } },
      },
    ]);
  });

  it('treats a formula-kind CONDITIONAL_FORMAT with an empty formula as unsupported', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Regional Revenue',
        range: 'A2:D9',
        rule: { kind: 'formula', formula: '  ', format: { fillColor: '#FFC7CE' } },
      },
    ]);
    expect(rich).toHaveLength(0);
    expect(unsupported).toHaveLength(1);
  });

  it("routes a topBottom-kind CONDITIONAL_FORMAT to the rich engine (TASKS.md #36 — 'top 5 suppliers by total')", () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 5, format: { fillColor: '#C6EFCE' } },
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 5, format: { fillColor: '#C6EFCE' } },
      },
    ]);
  });

  it('preserves isPercent on a topBottom-kind CONDITIONAL_FORMAT', () => {
    const { rich } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'topBottom', side: 'bottom', rank: 10, isPercent: true, format: { fillColor: '#FFC7CE' } },
      },
    ]);

    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'topBottom', side: 'bottom', rank: 10, isPercent: true, format: { fillColor: '#FFC7CE' } },
      },
    ]);
  });

  it('treats a topBottom-kind CONDITIONAL_FORMAT with an invalid side or non-positive rank as unsupported', () => {
    const badSide = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        // 'middle' isn't a real ConditionalFormatTopBottomRule.side — simulates
        // malformed LLM/wire output, which the normalizer must reject at runtime.
        rule: { kind: 'topBottom', side: 'middle', rank: 5, format: { fillColor: '#C6EFCE' } },
      } as unknown as SheetAction,
    ]);
    expect(badSide.rich).toHaveLength(0);
    expect(badSide.unsupported).toHaveLength(1);

    const badRank = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Suppliers',
        range: 'C2:C40',
        rule: { kind: 'topBottom', side: 'top', rank: 0, format: { fillColor: '#C6EFCE' } },
      },
    ]);
    expect(badRank.rich).toHaveLength(0);
    expect(badRank.unsupported).toHaveLength(1);
  });

  it("routes a colorScale-kind CONDITIONAL_FORMAT to the rich engine (TASKS.md #37 — 'color-scale the Total Amount column')", () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
      },
    ]);
  });

  it('routes a 2-color colorScale rule to the rich engine', () => {
    const { rich } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      },
    ]);

    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B', '#63BE7B'] },
      },
    ]);
  });

  it('treats a colorScale-kind CONDITIONAL_FORMAT with the wrong number of colors as unsupported', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Scores',
        range: 'B2:B50',
        rule: { kind: 'colorScale', colors: ['#F8696B'] },
      } as unknown as SheetAction,
    ]);
    expect(rich).toHaveLength(0);
    expect(unsupported).toHaveLength(1);
  });

  it('preserves existingRuleId on a CONDITIONAL_FORMAT action, so it modifies rather than duplicates (TASKS.md #38)', () => {
    const { rich } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        existingRuleId: 'cf-abc123',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
      },
    ]);

    expect(rich).toEqual([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
        existingRuleId: 'cf-abc123',
      },
    ]);
  });

  it('omits existingRuleId entirely (not just falsy) when absent — a plain create', () => {
    const { rich } = partitionActions([
      {
        type: 'CONDITIONAL_FORMAT',
        sheetName: 'Purchase Register',
        range: 'J2:J51',
        rule: { kind: 'cellValue', operator: 'greaterThan', value: 1500, format: { fillColor: '#FFC7CE' } },
      },
    ]);

    expect(rich[0]).not.toHaveProperty('existingRuleId');
  });

  it('treats a CONDITIONAL_FORMAT with no rule as unsupported rather than guessing', () => {
    const { rich, unsupported } = partitionActions([
      { type: 'CONDITIONAL_FORMAT', sheetName: 'Purchase Register', range: 'J2:J51' },
    ]);
    expect(rich).toHaveLength(0);
    expect(unsupported).toHaveLength(1);
  });

  it('normalizes CREATE_CHART with safe placement defaults', () => {
    const { rich, unsupported } = partitionActions([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Purchase Register',
        sourceRange: 'A1:B10',
        chartType: 'ColumnClustered',
        title: 'Purchases by Month',
      },
    ]);

    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      {
        type: 'CREATE_CHART',
        sheetName: 'Dashboard',
        sourceSheetName: 'Purchase Register',
        sourceRange: 'A1:B10',
        chartType: 'ColumnClustered',
        title: 'Purchases by Month',
        startCell: 'A1',
        endCell: 'H16',
      },
    ]);
  });

  it('routes DELETE_CHART to the rich engine (TASKS.md #15 — revert-only inverse of CREATE_CHART)', () => {
    const { rich, unsupported } = partitionActions([
      { type: 'DELETE_CHART', sheetName: 'Dashboard', chartId: 'Chart_real_1' },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      { type: 'DELETE_CHART', sheetName: 'Dashboard', chartId: 'Chart_real_1' },
    ]);
  });

  it('routes DELETE_CONDITIONAL_FORMAT to the rich engine (TASKS.md #67 — revert-only inverse of CONDITIONAL_FORMAT, previously silently unsupported)', () => {
    const { rich, unsupported } = partitionActions([
      { type: 'DELETE_CONDITIONAL_FORMAT', sheetName: 'Purchase Register', ruleId: 'cf-real-1' },
    ]);
    expect(unsupported).toHaveLength(0);
    expect(rich).toEqual([
      { type: 'DELETE_CONDITIONAL_FORMAT', sheetName: 'Purchase Register', ruleId: 'cf-real-1' },
    ]);
  });
});
