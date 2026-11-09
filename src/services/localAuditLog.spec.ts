import { beforeEach, describe, expect, it } from 'vitest';
import { auditCsv, auditEntries, clearAuditLog, describeAction, recordApplied, recordReverted } from './localAuditLog';
import { detectInspection } from '@/inspection/inspectionLanes';
import type { SheetAction } from '@/types/sheet-actions';

describe('P3.10 local audit trail', () => {
  beforeEach(() => clearAuditLog());

  it('one entry per accepted action — clean text then change case', () => {
    recordApplied('cs-1', 'Clean spaces and hidden characters in 5 cells.', [
      { type: 'BATCH_SET', sheetName: 'T_Clean', operations: [{ address: 'A2', value: 'ABC Traders' }, { address: 'A3', value: 'XYZ Pvt Ltd' }] } as SheetAction,
    ]);
    recordApplied('cs-2', 'Change 5 cells to proper case.', [
      { type: 'BATCH_SET', sheetName: 'T_Case', operations: [{ address: 'A2', value: 'Abc Pvt Ltd' }] } as SheetAction,
    ]);
    expect(auditEntries().map((e) => [e.changeSetId, e.sheet, e.range, e.cells, e.status])).toEqual([
      ['cs-1', 'T_Clean', 'A2, A3', 2, 'Applied'],
      ['cs-2', 'T_Case', 'A2', 1, 'Applied'],
    ]);
  });

  it('the export never contains cell values', () => {
    recordApplied('cs-1', 'Clean text', [
      { type: 'BATCH_SET', sheetName: 'T_Clean', operations: [{ address: 'A2', value: 'ZZ-SENTINEL-123' }] } as SheetAction,
      { type: 'WRITE_TABLE', sheetName: 'Out', headers: ['H'], rows: [['ZZ-SENTINEL-123']] } as SheetAction,
    ]);
    const csv = auditCsv();
    expect(csv).not.toMatch(/ZZ-SENTINEL-123/);
    expect(csv.split('\r\n')[0]).toBe('Time (UTC),Change set,What was asked / done,Action,Sheet,Range,Cells,Status');
  });

  it('marks reverted change sets', () => {
    recordApplied('cs-9', 'x', [{ type: 'FORMAT_RANGE', sheetName: 'S', range: 'A1:B3' } as SheetAction]);
    recordReverted('cs-9');
    expect(auditEntries()[0]).toMatchObject({ status: 'Reverted', cells: 6 });
  });

  it('describes ranges without values', () => {
    expect(describeAction({ type: 'FORMAT_RANGE', sheetName: 'S', range: 'C2:D6' } as SheetAction)).toEqual({ sheet: 'S', range: 'C2:D6', cells: 10 });
  });

  it('routes the prompt', () => {
    expect(detectInspection('Export the audit trail')).toBe('audit-export');
  });
});
