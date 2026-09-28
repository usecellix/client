import { describe, expect, it } from 'vitest';
import { deepToApiWorkbookContext } from './contextAdapter';
import { DeepWorkbookContext, SheetContext } from './context.types';

function makeSheet(overrides: Partial<SheetContext>): SheetContext {
  return {
    name: 'Sheet1',
    usedRange: 'A1:A1',
    rowCount: 1,
    columnCount: 1,
    values: [['x']],
    formulas: [['']],
    numberFormats: [['General']],
    structure: 'unknown',
    headers: [],
    headerRowIndex: 0,
    formulaSummary: '',
    isHidden: false,
    ...overrides,
  };
}

/**
 * TASKS.md #257 — `isHidden` was read correctly client-side (workbookReader.ts,
 * via Office.js's real `worksheet.visibility`) but silently dropped converting
 * into the lighter-weight `SheetSnapshot` this function builds, which is what
 * actually gets sent to the backend. A live "how many sheets" question then had
 * no way to know which of the listed sheets were hidden.
 */
describe('deepToApiWorkbookContext — isHidden propagation (#257)', () => {
  it('carries isHidden through for both visible and hidden sheets', () => {
    const deep: DeepWorkbookContext = {
      activeSheetName: 'Purchase Register',
      sheets: [
        makeSheet({ name: 'Purchase Register', isHidden: false }),
        makeSheet({ name: 'Working', isHidden: true }),
      ],
      namedRanges: [],
      tables: [],
      conditionalFormats: [],
      prompt_context: '',
    };

    const api = deepToApiWorkbookContext(deep);

    expect(api.sheets.find((s) => s.sheetName === 'Purchase Register')?.isHidden).toBe(false);
    expect(api.sheets.find((s) => s.sheetName === 'Working')?.isHidden).toBe(true);
  });
});
