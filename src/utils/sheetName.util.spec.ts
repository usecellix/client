import { describe, expect, it } from 'vitest';
import {
  detectCompoundSheetFollowUp,
  extractSheetNameFromPrompt,
  sanitizeExcelSheetName,
  stripIllegalSheetNameChars,
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

  it('sanitizeExcelSheetName strips every Excel-illegal character: \\ / ? * [ ] :', () => {
    expect(sanitizeExcelSheetName('A\\B/C?D*E[F]G:H')).toBe('A B C D E F G H');
  });

  it('stripIllegalSheetNameChars strips characters without truncating or falling back', () => {
    expect(stripIllegalSheetNameChars('Missed vs GSTR-2B / GSTR-2A')).toBe(
      'Missed vs GSTR-2B GSTR-2A',
    );
    // No 31-char truncation — callers that need to reserve room for a suffix
    // truncate the stripped result themselves.
    const long = 'A'.repeat(40);
    expect(stripIllegalSheetNameChars(long)).toBe(long);
    // No fallback for an all-illegal/empty input — returns empty string.
    expect(stripIllegalSheetNameChars('///')).toBe('');
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
