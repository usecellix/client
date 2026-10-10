import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { AttachmentError } from './attachmentError';
import { readGrid, XlsxLike } from './gridDecoder';

const xlsx = XLSX as unknown as XlsxLike;
const encode = (text: string) => new TextEncoder().encode(text);

/** A workbook written to real file bytes, so the test goes through the same reader as an attached file. */
function workbookBytes(sheets: Record<string, XLSX.WorkSheet>, bookType: 'xlsx' | 'biff8'): Uint8Array {
  const workbook = XLSX.utils.book_new();
  for (const [name, sheet] of Object.entries(sheets)) XLSX.utils.book_append_sheet(workbook, sheet, name);
  return new Uint8Array(XLSX.write(workbook, { type: 'array', bookType }) as ArrayBuffer);
}

function statementSheet(): XLSX.WorkSheet {
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Sample Bank Ltd'],
    [],
    ['Date', 'Narration', 'Chq No', 'Withdrawal', 'Deposit', 'Balance'],
    [46124, 'NEFT IN SYNTH CORP', 'N0001', null, 15000, 125000],
    ['13/04/2026', 'CHQ PAID RENT', '004521', 20000.5, null, 104999.5],
  ]);
  // A4 is a true date cell: a serial number with a date format.
  sheet.A4.z = 'dd-mmm-yyyy';
  return sheet;
}

const cells = (row: { cells: { t: string }[] }) => row.cells.map((cell) => cell.t);

describe.each([
  ['xlsx', 'xlsx'],
  ['xls', 'biff8'],
] as const)('readGrid — .%s file', (source, bookType) => {
  const table = readGrid(xlsx, workbookBytes({ Statement: statementSheet() }, bookType), `statement.${source}`, source);

  it('is a grid table that keeps the sheet row numbers', () => {
    expect(table).toMatchObject({ source, fileName: `statement.${source}`, layout: 'grid' });
    // Row 2 is blank and is left out, so the refs skip it rather than renumbering.
    expect(table.rows.map((row) => row.ref)).toEqual(['row 1', 'row 3', 'row 4', 'row 5']);
  });

  it('turns a true date cell into an ISO date', () => {
    expect(cells(table.rows[2])[0]).toBe('2026-04-12');
  });

  it('leaves a date typed as text exactly as written', () => {
    expect(cells(table.rows[3])[0]).toBe('13/04/2026');
  });

  it('writes numbers plainly and keeps blanks blank', () => {
    expect(cells(table.rows[2])).toEqual(['2026-04-12', 'NEFT IN SYNTH CORP', 'N0001', '', '15000', '125000']);
    expect(cells(table.rows[3])).toEqual(['13/04/2026', 'CHQ PAID RENT', '004521', '20000.5', '', '104999.5']);
  });
});

describe('readGrid — number cells', () => {
  const read = (sheet: XLSX.WorkSheet) =>
    readGrid(xlsx, workbookBytes({ S: sheet }, 'xlsx'), 's.xlsx', 'xlsx').rows.map(cells);

  it('removes binary floating-point noise', () => {
    expect(read(XLSX.utils.aoa_to_sheet([[0.1 + 0.2, 20000.499999999996]]))).toEqual([['0.3', '20000.5']]);
  });

  it('keeps the zeros of a reference stored as a number with a zero-padded format', () => {
    const sheet = XLSX.utils.aoa_to_sheet([[4521]]);
    sheet.A1.z = '000000';
    expect(read(sheet)).toEqual([['004521']]);
  });

  it('does not let a thousands format leak separators into the value', () => {
    const sheet = XLSX.utils.aoa_to_sheet([[125000]]);
    sheet.A1.z = '#,##0';
    expect(read(sheet)).toEqual([['125000']]);
  });
});

describe('readGrid — CSV', () => {
  it('does not interpret anything: dates, amounts and references stay as written', () => {
    const csv = [
      'Date,Narration,Ref,Debit,Credit,Balance',
      '04/05/2026,"UPI, TO SHOP",004521,"1,250.00",,"98,750.00"',
      '05/05/2026,=CMD|x,1E5,,500.00,"99,250.00"',
    ].join('\n');
    const table = readGrid(xlsx, encode(csv), 'statement.csv', 'csv');

    expect(table.source).toBe('csv');
    expect(table.rows.map(cells)).toEqual([
      ['Date', 'Narration', 'Ref', 'Debit', 'Credit', 'Balance'],
      // 04/05 is neither turned into a date nor reordered; 004521 keeps its zeros.
      ['04/05/2026', 'UPI, TO SHOP', '004521', '1,250.00', '', '98,750.00'],
      // A formula-looking cell is text, and 1E5 is not read as 100000.
      ['05/05/2026', '=CMD|x', '1E5', '', '500.00', '99,250.00'],
    ]);
  });

  it('reads a file that starts with a byte-order mark', () => {
    const table = readGrid(xlsx, encode('﻿Date,Amount\n01/04/2026,10.00\n'), 's.csv', 'csv');
    expect(cells(table.rows[0])).toEqual(['Date', 'Amount']);
  });
});

describe('readGrid — an HTML table saved with an .xls extension', () => {
  it('reads the table, again without interpreting values', () => {
    const html =
      '<html><body><table>' +
      '<tr><td>Date</td><td>Narration</td><td>Debit</td><td>Balance</td></tr>' +
      '<tr><td>04/05/2026</td><td>ATM WDL</td><td>2,000.00</td><td>8,000.00</td></tr>' +
      '</table></body></html>';
    const table = readGrid(xlsx, encode(html), 'statement.xls', 'xls');
    expect(table.rows.map(cells)).toEqual([
      ['Date', 'Narration', 'Debit', 'Balance'],
      ['04/05/2026', 'ATM WDL', '2,000.00', '8,000.00'],
    ]);
  });
});

describe('readGrid — choosing the sheet', () => {
  it('reads the sheet with the most rows, skipping a cover sheet', () => {
    const table = readGrid(
      xlsx,
      workbookBytes(
        { Cover: XLSX.utils.aoa_to_sheet([['Generated by Sample Bank']]), Statement: statementSheet() },
        'xlsx',
      ),
      's.xlsx',
      'xlsx',
    );
    expect(table.rows).toHaveLength(4);
  });

  it('refuses a workbook with nothing in it', () => {
    const bytes = workbookBytes({ Empty: XLSX.utils.aoa_to_sheet([]) }, 'xlsx');
    try {
      readGrid(xlsx, bytes, 'empty.xlsx', 'xlsx');
      throw new Error('expected readGrid to throw');
    } catch (error) {
      expect(error).toBeInstanceOf(AttachmentError);
      expect((error as AttachmentError).code).toBe('empty');
    }
  });
});
