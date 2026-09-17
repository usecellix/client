import { describe, expect, it } from 'vitest';
import {
  buildDeleteSheetActions,
  detectDeleteSheetIntent,
  detectEmptySheetCreateIntent,
  detectSheetDataGenerationIntent,
  extractDeleteSheetNames,
  tryLocalCreateEmptySheetActions,
  tryLocalDeleteSheetActions,
  tryLocalSheetActions,
} from './localSheetActions';
import { WorkbookContext } from '@/types/cellix.types';

const context: WorkbookContext = {
  activeSheet: 'Invoices',
  sheets: [
    { sheetName: 'Invoices' },
    { sheetName: 'Cellix' },
    { sheetName: 'Archive' },
  ] as WorkbookContext['sheets'],
};

/**
 * TASKS.md #250 — "Copy the X sheet and name it Y" (the guide's own T1.1
 * phrasing) always paid the full Tier 3 planner/executor/verifier LLM
 * round-trip (5-40s observed live) for what is a fully deterministic
 * operation. This local fast-lane resolves it instantly, client-side only.
 */
describe('tryLocalCopySheetActions (#250)', () => {
  it('resolves "Copy the X sheet and name it Y" to a single ADD_SHEET{copyFrom}', () => {
    const plan = tryLocalSheetActions(
      'Copy the Invoices sheet and name it March Copy',
      context,
      'action',
    );
    expect(plan).not.toBeNull();
    expect(plan?.actions).toEqual([
      { type: 'ADD_SHEET', name: 'March Copy', copyFrom: 'Invoices' },
    ]);
  });

  it('resolves "call it" phrasing too', () => {
    const plan = tryLocalSheetActions('Copy the Archive tab and call it Archive Backup', context, 'action');
    expect(plan?.actions).toEqual([
      { type: 'ADD_SHEET', name: 'Archive Backup', copyFrom: 'Archive' },
    ]);
  });

  it('dedupes the destination name against existing sheets', () => {
    const plan = tryLocalSheetActions('Copy the Invoices sheet and name it Cellix', context, 'action');
    expect(plan?.actions[0]).toMatchObject({ type: 'ADD_SHEET', copyFrom: 'Invoices' });
    expect((plan?.actions[0] as { name: string }).name).not.toBe('Cellix');
  });

  it('falls through to the backend when the source sheet does not resolve to a real sheet', () => {
    const plan = tryLocalSheetActions(
      'Copy the Nonexistent sheet and name it Copy 1',
      context,
      'action',
    );
    expect(plan).toBeNull();
  });

  it('does not fire for a filtered/partial copy — that still needs backend planning', () => {
    const plan = tryLocalSheetActions(
      'Copy rows where Status is Pending from Invoices to a new sheet called Pending',
      context,
      'action',
    );
    expect(plan).toBeNull();
  });
});

describe('localSheetActions', () => {
  it('detects delete sheet intent', () => {
    expect(detectDeleteSheetIntent('delete sheet Cellix')).toBe(true);
    expect(detectDeleteSheetIntent('remove the sheets Cellix and Archive')).toBe(true);
    expect(detectDeleteSheetIntent('create sheet Cellix')).toBe(false);
  });

  it('extracts sheet name from @[mention] tags', () => {
    expect(
      extractDeleteSheetNames('Delete the sheet Azhar @[Azhar]', ['Invoices', 'Azhar']),
    ).toEqual(['Azhar']);
  });

  it('uses @[mention] when sheet is not in available list', () => {
    expect(extractDeleteSheetNames('delete sheet Azhar @[Azhar]', ['Invoices'])).toEqual(['Azhar']);
  });

  it('handles delete with mention via tryLocalDeleteSheetActions', () => {
    const plan = tryLocalDeleteSheetActions(
      'Delete the sheet Azhar @[Azhar]',
      context,
      'action',
    );
    expect(plan?.actions).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Azhar' }]);
  });

  it('extracts a single sheet name', () => {
    expect(extractDeleteSheetNames('delete sheet Cellix', ['Invoices', 'Cellix'])).toEqual([
      'Cellix',
    ]);
  });

  it('extracts multiple sheet names', () => {
    expect(
      extractDeleteSheetNames('delete sheets Cellix and Archive', [
        'Invoices',
        'Cellix',
        'Archive',
      ]),
    ).toEqual(['Cellix', 'Archive']);
  });

  it('builds delete actions for multiple sheets', () => {
    expect(buildDeleteSheetActions(['Cellix', 'Archive'])).toEqual([
      { type: 'DELETE_SHEET', sheetName: 'Cellix' },
      { type: 'DELETE_SHEET', sheetName: 'Archive' },
    ]);
  });

  it('returns local plan without LLM for delete requests', () => {
    const plan = tryLocalDeleteSheetActions('delete sheets Cellix and Archive', context, 'action');
    expect(plan?.actions).toHaveLength(2);
    expect(plan?.explanation).toContain('Cellix');
  });

  it('detects empty sheet create vs data generation', () => {
    expect(detectEmptySheetCreateIntent('create an empty sheet named Cellix')).toBe(true);
    expect(detectSheetDataGenerationIntent('create sheet Cellix with 5 dummy rows')).toBe(true);
    expect(detectEmptySheetCreateIntent('create sheet Cellix with 5 dummy rows')).toBe(false);
  });

  it('returns local plan for empty sheet create', () => {
    const plan = tryLocalCreateEmptySheetActions(
      'create an empty sheet named Reports',
      context,
      'action',
    );
    expect(plan?.actions).toEqual([{ type: 'ADD_SHEET', name: 'Reports' }]);
    expect(plan?.explanation).toContain('Reports');
  });

  it('returns null for create with data so LLM is used', () => {
    expect(
      tryLocalSheetActions('create a new sheet named Summary with dummy data', context, 'action'),
    ).toBeNull();
  });

  it('returns null for add sheet and fill', () => {
    expect(tryLocalSheetActions('add a sheet and fill it', context, 'action')).toBeNull();
  });

  it('returns null for sort intent', () => {
    expect(tryLocalSheetActions('sort the invoices sheet', context, 'action')).toBeNull();
  });

  it('returns null for conditional delete', () => {
    expect(tryLocalSheetActions('delete all blank sheets', context, 'action')).toBeNull();
  });

  it('returns null for create sheet with chart analysis follow-up (LLM)', () => {
    expect(
      tryLocalSheetActions(
        'create a sheet called Purchase Register and give a chart analysis of the purchase register',
        context,
        'action',
      ),
    ).toBeNull();
  });

  it('returns null for create sheet with total row (LLM)', () => {
    expect(
      tryLocalSheetActions('create a sheet called Q2 and add a total row', context, 'action'),
    ).toBeNull();
  });

  it('handles remove tab phrasing for delete', () => {
    const plan = tryLocalSheetActions('remove the tab called Archive', context, 'action');
    expect(plan?.actions[0]).toEqual({ type: 'DELETE_SHEET', sheetName: 'Archive' });
  });

  it('deletes all sheets except the named keep sheet', () => {
    const plan = tryLocalDeleteSheetActions(
      'Delete all the sheets except purchase register',
      {
        activeSheet: 'Purchase Register',
        sheets: [
          { sheetName: 'Purchase Register' },
          { sheetName: 'Sales' },
          { sheetName: 'Invoices' },
        ] as WorkbookContext['sheets'],
      },
      'action',
    );
    expect(plan?.actions).toEqual([
      { type: 'DELETE_SHEET', sheetName: 'Sales' },
      { type: 'DELETE_SHEET', sheetName: 'Invoices' },
    ]);
    expect(plan?.explanation).not.toContain('Purchase Register');
  });

  it('still deletes a sheet when asked to delete it by name (not except-phrase)', () => {
    const plan = tryLocalDeleteSheetActions(
      'Delete sheet Purchase Register',
      {
        activeSheet: 'Purchase Register',
        sheets: [
          { sheetName: 'Purchase Register' },
          { sheetName: 'Sales' },
        ] as WorkbookContext['sheets'],
      },
      'action',
    );
    expect(plan?.actions).toEqual([{ type: 'DELETE_SHEET', sheetName: 'Purchase Register' }]);
  });

  it('handles blank tab create phrasing', () => {
    const plan = tryLocalSheetActions('create a blank tab', context, 'action');
    expect(plan?.actions[0]?.type).toBe('ADD_SHEET');
  });

  it('detects extended data generation intents', () => {
    expect(detectSheetDataGenerationIntent('create a sheet with headers')).toBe(true);
    expect(detectSheetDataGenerationIntent('add a total row below')).toBe(true);
    expect(detectSheetDataGenerationIntent('fill with sample rows')).toBe(true);
  });

  // TASKS.md #208 — the sheet has to be the delete verb's object, not just
  // mentioned as the place where rows/columns/duplicates get deleted.
  describe('delete requests that only mention a sheet as a location (#208)', () => {
    const guideContext: WorkbookContext = {
      activeSheet: 'Purchase Register',
      sheets: [
        { sheetName: 'Purchase Register' },
        { sheetName: 'Summary' },
        { sheetName: 'Rows Data' },
      ] as WorkbookContext['sheets'],
    };

    it.each([
      'Delete blank rows in the Purchase Register sheet',
      'Delete column C from this sheet',
      'Remove duplicates from the Summary sheet',
      'Remove the Narration column from the Summary sheet',
      'Delete all rows where column A is blank in this sheet',
      'Delete everything in the Summary sheet',
      'Remove the header from this sheet',
      'Delete the formatting on the Summary tab',
    ])('does not propose DELETE_SHEET for %j', (message) => {
      expect(tryLocalSheetActions(message, guideContext, 'action')).toBeNull();
    });

    it.each([
      ['Delete the Summary sheet', ['Summary']],
      ['Delete sheet Summary', ['Summary']],
      ['Remove this sheet', ['Purchase Register']],
      ['remove the tab called Summary', ['Summary']],
      ['Delete the Rows Data sheet', ['Rows Data']],
      ['Delete all the sheets except Summary', ['Purchase Register', 'Rows Data']],
    ])('still deletes the sheet for %j', (message, expected) => {
      const plan = tryLocalSheetActions(message, guideContext, 'action');
      expect(plan?.actions).toEqual(
        expected.map((sheetName) => ({ type: 'DELETE_SHEET', sheetName })),
      );
    });
  });

  // TASKS.md #181 — "clear all the data" / "clear the sheet" means make it a
  // plain workbook, not just wipe cell values and leave a stranded chart
  // floating over the empty grid.
  // TASKS.md #209 — a scoped clear must not wipe the whole sheet (and its
  // charts). Only an unqualified "clear the sheet" stays local.
  describe('scoped clears do not become whole-sheet clears (#209)', () => {
    it.each([
      'Clear all data in column C',
      'Clear all the content in the Narration column',
      'Clear all cells with errors',
      'Clear all the data where GSTIN is blank',
      'Clear the data in rows 5 to 10',
      'Clear all data in the Summary sheet',
    ])('sends %j to the backend instead of clearing the sheet', (message) => {
      expect(tryLocalSheetActions(message, context, 'action')).toBeNull();
    });

    // "clear the data" / "clear the entire sheet" never matched this lane's
    // trigger (it needs "all", and only "this entire sheet") — they already
    // went to the backend before #209 and still do.
    it.each(['Clear this sheet', 'clear all the data', 'clear all the cells', 'clear this entire sheet'])(
      'still clears the whole sheet for %j',
      (message) => {
        const plan = tryLocalSheetActions(message, context, 'action');
        expect(plan?.actions).toEqual([
          { type: 'CLEAR_RANGE', range: 'A1:XFD1048576', mode: 'contents', clearCharts: true },
        ]);
      },
    );
  });

  it('asks to clear charts too on a whole-sheet clear', () => {
    const plan = tryLocalSheetActions('clear all the data', context, 'action');
    expect(plan?.actions).toEqual([
      { type: 'CLEAR_RANGE', range: 'A1:XFD1048576', mode: 'contents', clearCharts: true },
    ]);
  });
});
