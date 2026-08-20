import { describe, expect, it } from 'vitest';
import { compressSheetData, prepareConversationRequestPayload } from '@/utils/payloadCompressor';

describe('compressSheetData', () => {
  it('keeps header plus first and last data rows for large sheets', () => {
    const rows = Array.from({ length: 1000 }, (_, index) => [
      index === 0 ? 'Name' : `Row ${index}`,
      index,
    ]);

    const compressed = compressSheetData(rows);

    expect(compressed.originalRowCount).toBe(1000);
    expect(compressed.truncated).toBe(true);
    expect(compressed.onDemandFetchEnabled).toBe(true);
    expect(compressed.compressedRowCount).toBeLessThan(20);
    expect(compressed.sheetData[0][0]).toBe('Name');
    expect(compressed.sheetData[1][0]).toBe('Row 1');
    expect(compressed.sheetData[compressed.sheetData.length - 1]?.[0]).toBe('Row 999');
  });

  it('does not truncate small sheets', () => {
    const rows = [
      ['A', 'B'],
      ['1', '2'],
      ['3', '4'],
    ];

    const compressed = compressSheetData(rows);

    expect(compressed.truncated).toBe(false);
    expect(compressed.compressedRowCount).toBe(3);
    expect(compressed.includedRowIndices).toEqual([0, 1, 2]);
  });
});

describe('prepareConversationRequestPayload', () => {
  it('sends the full sheet for find/search queries', () => {
    const rows = Array.from({ length: 100 }, (_, index) => [
      index === 0 ? 'CGST' : `${index}`,
      index === 25 ? '1868.41 Dr' : `${index * 10}`,
    ]);

    const payload = prepareConversationRequestPayload('Find CGST value 1868', rows);

    expect(payload.sheetCompression?.truncated).toBe(false);
    expect(payload.sheetData.length).toBe(100);
    expect(payload.sheetData[25]?.[1]).toBe('1868.41 Dr');
  });

  it('includes workbookId on the payload when provided (TASKS.md #23)', () => {
    const payload = prepareConversationRequestPayload('Add a total column', [['A'], ['1']], {
      workbookId: 'wb_abc123',
    });

    expect(payload.workbookId).toBe('wb_abc123');
  });

  it('omits workbookId entirely when not provided (backward compatible)', () => {
    const payload = prepareConversationRequestPayload('Add a total column', [['A'], ['1']]);

    expect(payload).not.toHaveProperty('workbookId');
  });
});
