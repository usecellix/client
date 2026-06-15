import { CompareResult } from '@/types/cellix.types';
import { deepToApiWorkbookContext } from '@/context/contextAdapter';
import {
  buildWorkbookContext as buildDeepWorkbookContext,
  readActiveSheetMinimal,
} from '@/context/workbookReader';
import { getCompareEndpoint } from '@/lib/apiConfig';

/* global Excel */

function colIndexToLetter(col: number): string {
  let index = col + 1;
  let letter = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

export async function buildWorkbookContext(
  selectedSheetNames: string[],
): Promise<{
  context: ReturnType<typeof deepToApiWorkbookContext>;
  activeSheetData: unknown[][];
  promptContext: string;
}> {
  try {
    const deep = await buildDeepWorkbookContext(selectedSheetNames);
    const context = deepToApiWorkbookContext(deep);
    const activeSheet = deep.sheets.find((s) => s.name === deep.activeSheetName);

    return {
      context,
      activeSheetData: activeSheet?.values ?? [],
      promptContext: deep.prompt_context,
    };
  } catch (err) {
    console.warn('[Cellix] Full workbook read failed, trying minimal active-sheet read:', err);
    return buildMinimalWorkbookContext();
  }
}

async function buildMinimalWorkbookContext(): Promise<{
  context: ReturnType<typeof deepToApiWorkbookContext>;
  activeSheetData: unknown[][];
  promptContext: string;
}> {
  const minimal = await readActiveSheetMinimal();
  const headers = (minimal.values[0] ?? []).map((cell) => String(cell ?? '').trim());

  const deep = {
    activeSheetName: minimal.name,
    selectedRange: minimal.usedRange,
    sheets: [
      {
        name: minimal.name,
        usedRange: minimal.usedRange,
        rowCount: minimal.rowCount,
        columnCount: minimal.columnCount,
        values: minimal.values,
        formulas: [],
        numberFormats: [],
        structure: 'unknown' as const,
        headers,
        formulaSummary: '',
        isHidden: false,
      },
    ],
    namedRanges: [],
    tables: [],
    prompt_context: `Active sheet: ${minimal.name}\nRows: ${minimal.rowCount}\nHeaders: ${headers.join(' | ')}`,
  };

  const context = deepToApiWorkbookContext(deep);
  return {
    context,
    activeSheetData: minimal.values,
    promptContext: deep.prompt_context,
  };
}

function isBlankCell(value: unknown): boolean {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeCell(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

export function compareSheetValuesLocally(
  sheetA: string,
  sheetB: string,
  valuesA: unknown[][],
  valuesB: unknown[][],
): CompareResult {
  const maxRows = Math.max(valuesA.length, valuesB.length);
  const maxCols = Math.max(
    valuesA.reduce((m, r) => Math.max(m, r.length), 0),
    valuesB.reduce((m, r) => Math.max(m, r.length), 0),
  );

  const modifiedCells: CompareResult['modifiedCells'] = [];
  const addedInB: string[] = [];
  const removedInB: string[] = [];

  for (let row = 0; row < maxRows; row += 1) {
    const rowA = valuesA[row];
    const rowB = valuesB[row];
    const hasA = rowA?.some((cell) => !isBlankCell(cell));
    const hasB = rowB?.some((cell) => !isBlankCell(cell));

    if (!hasA && hasB) {
      addedInB.push(`Row ${row + 1}`);
      continue;
    }
    if (hasA && !hasB) {
      removedInB.push(`Row ${row + 1}`);
      continue;
    }

    for (let col = 0; col < maxCols; col += 1) {
      const a = normalizeCell(rowA?.[col]);
      const b = normalizeCell(rowB?.[col]);
      if (a !== b) {
        modifiedCells.push({
          address: `${colIndexToLetter(col)}${row + 1}`,
          valueA: a || '(empty)',
          valueB: b || '(empty)',
        });
      }
    }
  }

  return {
    sheetA,
    sheetB,
    summary: `${modifiedCells.length} cell difference(s), ${addedInB.length} row(s) added in B, ${removedInB.length} row(s) removed in B`,
    differences: [],
    addedInB,
    removedInB,
    modifiedCells,
  };
}

async function readFullSheetValues(sheetName: string): Promise<unknown[][]> {
  const deep = await buildDeepWorkbookContext([sheetName]);
  const sheet = deep.sheets.find((s) => s.name === sheetName);
  return sheet?.values ?? [];
}

export async function compareSheets(
  sheetA: string,
  sheetB: string,
): Promise<CompareResult> {
  const { context } = await buildWorkbookContext([sheetA, sheetB]);

  try {
    const endpoint = getCompareEndpoint();
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (endpoint.includes('.ngrok-free.app')) {
      headers['ngrok-skip-browser-warning'] = 'true';
    }

    const response = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify({ sheetA, sheetB, context }),
    });

    if (response.ok) {
      const body = (await response.json()) as CompareResult | { data?: CompareResult };
      return ('data' in body && body.data ? body.data : body) as CompareResult;
    }
  } catch (error) {
    console.warn('[Cellix] Compare API unavailable, using local diff:', error);
  }

  const [valuesA, valuesB] = await Promise.all([
    readFullSheetValues(sheetA),
    readFullSheetValues(sheetB),
  ]);

  return compareSheetValuesLocally(sheetA, sheetB, valuesA, valuesB);
}
