import { RawTableCell, RawTableRow } from '@/types/rawTable';

/**
 * Rebuilds the rows of a PDF page from positioned text items.
 *
 * A PDF has no rows or columns, only pieces of text at coordinates, and a plain
 * text dump scrambles a statement (the dates come out as one block, detached
 * from their amounts). Grouping by height gives the rows back; the left and
 * right edge of every piece is kept so the server can work out the columns.
 *
 * Pure: takes plain data, so it runs the same in the task pane and in tests.
 */

export interface PdfTextItem {
  str: string;
  /** Left edge in page points. */
  x: number;
  /** Baseline height in page points, larger is higher up the page. */
  y: number;
  width: number;
  /** Font height in points. */
  height: number;
}

export interface PdfPageText {
  /** 1-based. */
  page: number;
  items: PdfTextItem[];
}

/** Items within this many points of each other vertically are on one line. */
const SAME_LINE_TOLERANCE = 2;
const FALLBACK_FONT_HEIGHT = 8;
const COMPLETE_NUMBER_RE = /^[-+(]?[\d,]+\.\d{2}\)?$/;

const round1 = (value: number): number => Math.round(value * 10) / 10;

interface Line {
  y: number;
  items: PdfTextItem[];
}

function groupIntoLines(items: PdfTextItem[]): Line[] {
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: Line[] = [];
  for (const item of sorted) {
    const line = lines[lines.length - 1];
    // Compared with the line's first item, not its latest, so a line cannot drift downwards.
    if (line && Math.abs(line.y - item.y) <= SAME_LINE_TOLERANCE) line.items.push(item);
    else lines.push({ y: item.y, items: [item] });
  }
  return lines;
}

/**
 * Joins the items of one line into cells. Some PDFs emit a cell as one item,
 * others emit a word or a single character at a time, so pieces that touch are
 * merged and pieces separated by a column gutter are not.
 */
function lineToCells(line: Line): RawTableCell[] {
  const cells: Array<{ t: string; x0: number; x1: number }> = [];
  let pendingSpace = false;

  for (const item of [...line.items].sort((a, b) => a.x - b.x)) {
    if (!item.str.trim()) {
      pendingSpace = pendingSpace || item.str.length > 0;
      continue;
    }
    const text = item.str.trim();
    const x0 = item.x;
    const x1 = item.x + Math.max(0, item.width);
    const height = item.height > 0 ? item.height : FALLBACK_FONT_HEIGHT;
    const previous = cells[cells.length - 1];

    if (previous) {
      const gap = x0 - previous.x1;
      const touching = gap <= height * 0.3;
      const spaced = pendingSpace && gap <= height * 0.7;
      // Two finished amounts side by side are two columns however close they sit.
      const twoNumbers = COMPLETE_NUMBER_RE.test(previous.t) && COMPLETE_NUMBER_RE.test(text);
      if ((touching || spaced) && !twoNumbers) {
        const separator = pendingSpace || gap > height * 0.12 ? ' ' : '';
        previous.t = `${previous.t}${separator}${text}`;
        previous.x1 = Math.max(previous.x1, x1);
        pendingSpace = false;
        continue;
      }
    }
    cells.push({ t: text, x0, x1 });
    pendingSpace = false;
  }

  return cells.map((cell) => ({ t: cell.t, x0: round1(cell.x0), x1: round1(cell.x1) }));
}

export function buildPdfRows(pages: PdfPageText[]): RawTableRow[] {
  const rows: RawTableRow[] = [];
  for (const page of pages) {
    let lineNumber = 0;
    for (const line of groupIntoLines(page.items)) {
      const cells = lineToCells(line);
      if (!cells.length) continue;
      lineNumber += 1;
      rows.push({ ref: `p${page.page} l${lineNumber}`, page: page.page, y: round1(line.y), cells });
    }
  }
  return rows;
}

/** Characters of real text across the document, used to tell a scanned PDF from a text one. */
export function countPdfTextCharacters(pages: PdfPageText[]): number {
  let count = 0;
  for (const page of pages) {
    for (const item of page.items) count += item.str.replace(/\s+/g, '').length;
  }
  return count;
}
