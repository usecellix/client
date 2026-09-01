import { describe, expect, it } from 'vitest';
import {
  coerceValueForNumberFormat,
  isGeneralOrEmptyFormat,
  resolvePreservedNumberFormat,
} from '../services/formatGuard';

describe('formatGuard date format preservation', () => {
  it('treats General and empty as unset formats', () => {
    expect(isGeneralOrEmptyFormat(undefined)).toBe(true);
    expect(isGeneralOrEmptyFormat('')).toBe(true);
    expect(isGeneralOrEmptyFormat('General')).toBe(true);
    expect(isGeneralOrEmptyFormat('d-mmm-yyyy')).toBe(false);
    expect(isGeneralOrEmptyFormat('dd/mm/yyyy')).toBe(false);
  });

  it('keeps the sheet cell format over locale defaults', () => {
    expect(resolvePreservedNumberFormat('d-mmm-yyyy', 'General')).toBe('d-mmm-yyyy');
    expect(resolvePreservedNumberFormat('General', 'dd/mm/yyyy')).toBe('dd/mm/yyyy');
    expect(resolvePreservedNumberFormat('General', 'General', 'yyyy-mm-dd')).toBe('yyyy-mm-dd');
    expect(resolvePreservedNumberFormat('m/d/yyyy', 'd-mmm-yyyy')).toBe('m/d/yyyy');
  });

  it('coerces ISO date strings when the column is a date format', () => {
    const coerced = coerceValueForNumberFormat('2026-01-15', 'd-mmm-yyyy');
    expect(coerced).toBeInstanceOf(Date);
    expect((coerced as Date).getFullYear()).toBe(2026);
    expect((coerced as Date).getMonth()).toBe(0);
    expect((coerced as Date).getDate()).toBe(15);
  });

  it('leaves non-date columns unchanged', () => {
    expect(coerceValueForNumberFormat('2026-01-15', 'General')).toBe('2026-01-15');
    expect(coerceValueForNumberFormat('Paid', '@')).toBe('Paid');
  });
});
