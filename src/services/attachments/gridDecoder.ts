import type { CellObject, WorkBook, WorkSheet } from 'xlsx';
import { RawTable, RawTableCell, RawTableRow, RawTableSource } from '@/types/rawTable';
import { AttachmentError } from './attachmentError';

/** The part of SheetJS this decoder uses, so tests and the task pane can each supply it. */
export interface XlsxLike {
  read(data: Uint8Array | string, options: Record<string, unknown>): WorkBook;
  utils: {
    decode_range(range: string): { s: { r: number; c: number }; e: { r: number; c: number } };
    encode_cell(address: { r: number; c: number }): string;
  };
  SSF: {
    is_date(format: string): boolean;
    parse_date_code(serial: number): { y: number; m: number; d: number } | null;
  };
}

/** Wider or longer than this is not a bank statement; it stops a huge workbook being walked cell by cell. */
const MAX_COLUMNS = 80;
const MAX_ROWS = 200_000;

const pad2 = (n: number) => String(n).padStart(2, '0');

/**
 * One cell as text, with no interpretation beyond what the file itself states.
 *
 * A true date cell becomes ISO `yyyy-mm-dd`, which is unambiguous. Text that
 * merely looks like a date ("04/05/2026") is passed through untouched for the
 * server to read in context: guessing day-first or month-first one cell at a
 * time is how a statement gets silently misdated.
 */
function cellText(xlsx: XlsxLike, cell: CellObject | undefined): string {
  if (!cell || cell.v === undefined || cell.v === null) return '';

  if (cell.t === 'n' && typeof cell.v === 'number') {
    const format = typeof cell.z === 'string' ? cell.z : '';
    if (format && xlsx.SSF.is_date(format)) {
      const parts = xlsx.SSF.parse_date_code(cell.v);
      if (parts && parts.y > 0) return `${parts.y}-${pad2(parts.m)}-${pad2(parts.d)}`;
    }
    if (Number.isInteger(cell.v)) {
      // A reference stored as a number with a "000000" format keeps its zeros in the display text.
      const shown = typeof cell.w === 'string' ? cell.w.trim() : '';
      return /^\d+$/.test(shown) ? shown : String(cell.v);
    }
    // Trims binary noise such as 20000.499999999996 without losing real decimals.
    return String(Number(cell.v.toFixed(4)));
  }
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.t === 'e') return '';
  return String(cell.v).trim();
}

function sheetRows(xlsx: XlsxLike, sheet: WorkSheet): RawTableRow[] {
  const ref = sheet['!ref'];
  if (typeof ref !== 'string') return [];
  const range = xlsx.utils.decode_range(ref);
  const lastColumn = Math.min(range.e.c, range.s.c + MAX_COLUMNS - 1);
  const lastRow = Math.min(range.e.r, range.s.r + MAX_ROWS - 1);

  const rows: RawTableRow[] = [];
  for (let r = range.s.r; r <= lastRow; r++) {
    const cells: RawTableCell[] = [];
    // Always from column A, so a cell's index is its column whatever the used range starts at.
    for (let c = 0; c <= lastColumn; c++) {
      cells.push({ t: cellText(xlsx, sheet[xlsx.utils.encode_cell({ r, c })] as CellObject | undefined) });
    }
    while (cells.length && !cells[cells.length - 1].t) cells.pop();
    if (cells.length) rows.push({ ref: `row ${r + 1}`, cells });
  }
  return rows;
}

/** The sheet with the most filled rows. A bank export has one sheet; this skips an empty cover sheet. */
export function workbookToRawTable(
  xlsx: XlsxLike,
  workbook: WorkBook,
  fileName: string,
  source: RawTableSource,
): RawTable {
  let best: RawTableRow[] = [];
  for (const name of workbook.SheetNames) {
    const rows = sheetRows(xlsx, workbook.Sheets[name]);
    if (rows.length > best.length) best = rows;
  }
  if (!best.length) {
    throw new AttachmentError('empty', 'This file has no data in it.');
  }
  return { source, fileName, layout: 'grid', rows: best };
}

/**
 * `raw: true` stops SheetJS parsing values in text formats (CSV, and the HTML
 * tables some banks save with an .xls extension), so dates and amounts arrive
 * exactly as written. Typed cells in real Excel files are unaffected by it.
 */
const READ_OPTIONS = { cellDates: false, cellNF: true, cellText: true, raw: true };

export function readGrid(
  xlsx: XlsxLike,
  data: Uint8Array,
  fileName: string,
  source: RawTableSource,
): RawTable {
  let workbook: WorkBook;
  try {
    workbook =
      source === 'csv'
        ? xlsx.read(new TextDecoder('utf-8').decode(data), { ...READ_OPTIONS, type: 'string' })
        : xlsx.read(data, { ...READ_OPTIONS, type: 'array' });
  } catch {
    throw new AttachmentError('read_failed', 'This file could not be opened. It may be damaged or password protected.');
  }
  return workbookToRawTable(xlsx, workbook, fileName, source);
}

let xlsxPromise: Promise<XlsxLike> | null = null;

/** Loaded on first use so the spreadsheet reader is not in the task pane's startup bundle. */
function loadXlsx(): Promise<XlsxLike> {
  if (!xlsxPromise) {
    xlsxPromise = import('xlsx').then((module) => module as unknown as XlsxLike);
    xlsxPromise.catch(() => {
      xlsxPromise = null;
    });
  }
  return xlsxPromise;
}

export async function decodeGrid(
  data: Uint8Array,
  fileName: string,
  source: RawTableSource,
): Promise<RawTable> {
  let xlsx: XlsxLike;
  try {
    xlsx = await loadXlsx();
  } catch {
    throw new AttachmentError('read_failed', 'The spreadsheet reader could not be loaded. Check your connection and try again.');
  }
  return readGrid(xlsx, data, fileName, source);
}
