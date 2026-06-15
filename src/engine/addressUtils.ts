export function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function columnIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

/** Strip sheet qualifier — `'Purchases'!A1:M339` → `A1:M339`. */
export function stripSheetPrefix(address: string): string {
  const trimmed = address.trim().replace(/\$/g, '');
  const bang = trimmed.lastIndexOf('!');
  return bang >= 0 ? trimmed.slice(bang + 1) : trimmed;
}

const LOCAL_RANGE = /^[A-Za-z]+\d+(?::[A-Za-z]+\d+)?$/;

export function isLocalRangeAddress(address: string): boolean {
  return LOCAL_RANGE.test(stripSheetPrefix(address));
}

export function parseCellAddress(address: string): { row: number; col: number } | null {
  const local = stripSheetPrefix(address);
  const match = /^([A-Za-z]+)(\d+)$/.exec(local.trim());
  if (!match) return null;
  return {
    col: columnLetterToIndex(match[1]),
    row: Number.parseInt(match[2], 10) - 1,
  };
}

export function parseRangeAddress(range: string): {
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
} | null {
  const local = stripSheetPrefix(range);
  const parts = local.split(':');
  const start = parseCellAddress(parts[0]);
  if (!start) return null;
  if (parts.length === 1) {
    return { row: start.row, col: start.col, rowCount: 1, colCount: 1 };
  }
  const end = parseCellAddress(parts[1]);
  if (!end) return null;
  return {
    row: start.row,
    col: start.col,
    rowCount: end.row - start.row + 1,
    colCount: end.col - start.col + 1,
  };
}
