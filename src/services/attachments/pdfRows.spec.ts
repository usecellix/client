import { describe, expect, it } from 'vitest';
import { buildPdfRows, countPdfTextCharacters, PdfTextItem } from './pdfRows';

const item = (str: string, x: number, y: number, width: number, height = 7): PdfTextItem => ({
  str,
  x,
  y,
  width,
  height,
});

describe('buildPdfRows', () => {
  it('groups items into rows by height, top of the page first', () => {
    const rows = buildPdfRows([
      {
        page: 1,
        items: [
          // Deliberately out of order: a PDF lists text in drawing order, not reading order.
          item('S93580561', 318.3, 417, 34.1),
          item('UPIOUT/1234', 129.5, 392, 60),
          item('11-APR-2026', 22.4, 417, 42.1),
          item('12-APR-2026', 22.4, 392, 42.1),
        ],
      },
    ]);

    expect(rows).toEqual([
      {
        ref: 'p1 l1',
        page: 1,
        y: 417,
        cells: [
          { t: '11-APR-2026', x0: 22.4, x1: 64.5 },
          { t: 'S93580561', x0: 318.3, x1: 352.4 },
        ],
      },
      {
        ref: 'p1 l2',
        page: 1,
        y: 392,
        cells: [
          { t: '12-APR-2026', x0: 22.4, x1: 64.5 },
          { t: 'UPIOUT/1234', x0: 129.5, x1: 189.5 },
        ],
      },
    ]);
  });

  it('keeps lines a few points apart as separate rows', () => {
    // A three-line header: "Tran" above, "Date" on the main line, "Type" below.
    const rows = buildPdfRows([
      { page: 1, items: [item('Tran', 265.5, 459, 16.3), item('Date', 32.8, 454.7, 15), item('Type', 265.5, 450.3, 16.3)] },
    ]);
    expect(rows.map((row) => row.cells[0].t)).toEqual(['Tran', 'Date', 'Type']);
  });

  it('treats items within two points vertically as one row', () => {
    const rows = buildPdfRows([{ page: 1, items: [item('Debit', 10, 100, 20), item('Credit', 80, 101.5, 24)] }]);
    expect(rows).toHaveLength(1);
    expect(rows[0].cells.map((cell) => cell.t)).toEqual(['Debit', 'Credit']);
  });

  it('numbers lines per page', () => {
    const rows = buildPdfRows([
      { page: 1, items: [item('a', 10, 100, 5), item('b', 10, 80, 5)] },
      { page: 2, items: [item('c', 10, 100, 5)] },
    ]);
    expect(rows.map((row) => row.ref)).toEqual(['p1 l1', 'p1 l2', 'p2 l1']);
  });

  it('joins a cell that the PDF emits one character at a time', () => {
    const letters = 'Balance'.split('').map((ch, i) => item(ch, 100 + i * 4, 50, 4));
    const rows = buildPdfRows([{ page: 1, items: letters }]);
    expect(rows[0].cells).toEqual([{ t: 'Balance', x0: 100, x1: 128 }]);
  });

  it('joins words separated by a space item, with a single space', () => {
    const rows = buildPdfRows([
      { page: 1, items: [item('Opening', 129.5, 50, 28), item(' ', 157.5, 50, 2), item('Balance', 159.5, 50, 22)] },
    ]);
    expect(rows[0].cells).toEqual([{ t: 'Opening Balance', x0: 129.5, x1: 181.5 }]);
  });

  it('does not join cells separated by a column gutter, even with a space item between them', () => {
    const rows = buildPdfRows([
      { page: 1, items: [item('11-APR-2026', 22.4, 50, 42.1), item(' ', 64.5, 50, 2), item('11-APR-2026', 79.1, 50, 42.1)] },
    ]);
    expect(rows[0].cells).toHaveLength(2);
  });

  it('never joins two finished amounts, however close they sit', () => {
    // A deposit and a balance wide enough to touch must stay two columns.
    const rows = buildPdfRows([
      { page: 1, items: [item('250000.00', 473.2, 50, 31.9), item('1258687.44', 505.6, 50, 43.1)] },
    ]);
    expect(rows[0].cells.map((cell) => cell.t)).toEqual(['250000.00', '1258687.44']);
  });

  it('drops rows that hold only whitespace', () => {
    const rows = buildPdfRows([{ page: 1, items: [item('   ', 10, 100, 9), item('real', 10, 80, 12)] }]);
    expect(rows).toHaveLength(1);
    expect(rows[0].ref).toBe('p1 l1');
  });
});

describe('countPdfTextCharacters', () => {
  it('counts text and ignores whitespace', () => {
    expect(
      countPdfTextCharacters([
        { page: 1, items: [item('ab c', 0, 0, 1), item('   ', 0, 0, 1)] },
        { page: 2, items: [item('de', 0, 0, 1)] },
      ]),
    ).toBe(5);
  });
});
