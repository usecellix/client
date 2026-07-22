import { describe, expect, it } from 'vitest';
import {
  detectCompoundSheetFollowUp,
  extractSheetNameFromPrompt,
  sanitizeExcelSheetName,
} from './sheetName.util';

describe('sheetName.util', () => {
  it('stops unquoted sheet name before compound and-clause', () => {
    expect(
      extractSheetNameFromPrompt(
        'create a sheet called Purchase Register and give a chart analysis of the purchase register',
      ),
    ).toBe('Purchase Register');
  });

  it('sanitizes invalid characters and truncates to 31 chars', () => {
    const long = 'A'.repeat(40);
    expect(sanitizeExcelSheetName(long).length).toBe(31);
    expect(sanitizeExcelSheetName('Bad/Name?')).toBe('Bad Name');
  });

  it('detects compound follow-up after sheet create', () => {
    expect(
      detectCompoundSheetFollowUp(
        'create a sheet called Q2 and give a chart analysis',
      ),
    ).toBe(true);
    expect(detectCompoundSheetFollowUp('create an empty sheet named Reports')).toBe(false);
  });
});
