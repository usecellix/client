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

  it('handles blank tab create phrasing', () => {
    const plan = tryLocalSheetActions('create a blank tab', context, 'action');
    expect(plan?.actions[0]?.type).toBe('ADD_SHEET');
  });

  it('detects extended data generation intents', () => {
    expect(detectSheetDataGenerationIntent('create a sheet with headers')).toBe(true);
    expect(detectSheetDataGenerationIntent('add a total row below')).toBe(true);
    expect(detectSheetDataGenerationIntent('fill with sample rows')).toBe(true);
  });
});
