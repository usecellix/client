import { describe, expect, it, vi } from 'vitest';
import { extractColumnFormats } from './workbookReader';

/**
 * TASKS.md #64 — reading bold/italic/fontColor/fillColor from a column's first
 * data row (not the header row — headers are almost always bold, so reading
 * the header row would make revert *introduce* bold on data cells rather than
 * restore them). Mocks the Office.js call chain directly, same convention as
 * `workbookReader.conditionalFormat.spec.ts`'s `extractConditionalFormats`
 * tests — a real add-in smoke test against live Excel remains an unverified
 * caveat, same class as #34/#65/#15/#67.
 */
describe('extractColumnFormats', () => {
  function makeCtx(cellProps: unknown[][], numberFormats: string[][]) {
    const sampleRow = {
      numberFormat: numberFormats,
      load: vi.fn(),
      getCellProperties: vi.fn(() => ({ value: cellProps })),
    };
    const sheet = { getRangeByIndexes: vi.fn(() => sampleRow) } as unknown as Excel.Worksheet;
    const usedRange = { rowIndex: 0, columnIndex: 0 } as unknown as Excel.Range;
    const ctx = { sync: vi.fn(async () => undefined) } as unknown as Excel.RequestContext;
    return { ctx, sheet, usedRange, sampleRow };
  }

  it('reads bold/fillColor/numberFormat from the first DATA row, not the header row', async () => {
    const { ctx, sheet, usedRange, sampleRow } = makeCtx(
      [[{ format: { font: { bold: false, italic: false, color: '#000000' }, fill: { color: '' } } }]],
      [['0.00%']],
    );

    const result = await extractColumnFormats(ctx, sheet, usedRange, 0, 2, 1);

    // headerRowIndex=0, rowCount=2 -> sample row is index 1 (the data row), not 0.
    expect(sheet.getRangeByIndexes).toHaveBeenCalledWith(1, 0, 1, 1);
    expect(sampleRow.getCellProperties).toHaveBeenCalledWith({
      format: { font: { bold: true, italic: true, color: true }, fill: { color: true } },
    });
    expect(result).toEqual([
      { numberFormat: '0.00%', bold: false, italic: false, fontColor: '#000000', fillColor: undefined },
    ]);
  });

  it('falls back to the header row when the sheet has no data rows', async () => {
    const { ctx, sheet, usedRange } = makeCtx(
      [[{ format: { font: { bold: true }, fill: { color: '#FFFF00' } } }]],
      [['General']],
    );

    await extractColumnFormats(ctx, sheet, usedRange, 0, 1, 1);

    expect(sheet.getRangeByIndexes).toHaveBeenCalledWith(0, 0, 1, 1);
  });

  it('treats an empty-string color as "not captured" rather than a literal empty color', async () => {
    const { ctx, sheet, usedRange } = makeCtx(
      [[{ format: { font: { color: '' }, fill: { color: '' } } }]],
      [['General']],
    );

    const result = await extractColumnFormats(ctx, sheet, usedRange, 0, 2, 1);
    expect(result[0]?.fontColor).toBeUndefined();
    expect(result[0]?.fillColor).toBeUndefined();
  });

  it('returns one empty entry per column and never throws when the read fails entirely', async () => {
    const sheet = {
      getRangeByIndexes: vi.fn(() => {
        throw new Error('boom');
      }),
    } as unknown as Excel.Worksheet;
    const usedRange = { rowIndex: 0, columnIndex: 0 } as unknown as Excel.Range;
    const ctx = { sync: vi.fn(async () => undefined) } as unknown as Excel.RequestContext;

    const result = await extractColumnFormats(ctx, sheet, usedRange, 0, 3, 2);
    expect(result).toEqual([{}, {}]);
  });

  it('returns an empty array immediately when the sheet has no columns', async () => {
    const sheet = { getRangeByIndexes: vi.fn() } as unknown as Excel.Worksheet;
    const usedRange = { rowIndex: 0, columnIndex: 0 } as unknown as Excel.Range;
    const ctx = { sync: vi.fn(async () => undefined) } as unknown as Excel.RequestContext;

    const result = await extractColumnFormats(ctx, sheet, usedRange, 0, 0, 0);
    expect(result).toEqual([]);
    expect(sheet.getRangeByIndexes).not.toHaveBeenCalled();
  });
});
