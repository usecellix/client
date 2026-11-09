import { beforeEach, describe, expect, it } from 'vitest';
import { buildTablePdf, pdfSafe } from './simplePdf';
import { auditPdf, clearAuditLog, recordApplied } from './localAuditLog';
import type { SheetAction } from '@/types/sheet-actions';

describe('audit trail PDF export', () => {
  beforeEach(() => clearAuditLog());

  it('is a well-formed PDF: header, xref offsets that point at each object, EOF', () => {
    const pdf = buildTablePdf({ title: 'T', columns: [{ header: 'A', width: 100 }], rows: [['x'], ['y']] });
    expect(pdf.startsWith('%PDF-1.4')).toBe(true);
    expect(pdf.trimEnd().endsWith('%%EOF')).toBe(true);
    const xrefAt = Number(/startxref\n(\d+)/.exec(pdf)![1]);
    expect(pdf.slice(xrefAt, xrefAt + 4)).toBe('xref');
    const offsets = [...pdf.slice(xrefAt).matchAll(/^(\d{10}) 00000 n $/gm)].map((m) => Number(m[1]));
    offsets.forEach((off, i) => expect(pdf.slice(off, off + `${i + 1} 0 obj`.length)).toBe(`${i + 1} 0 obj`));
  });

  it('pages a long trail', () => {
    const rows = Array.from({ length: 120 }, (_, i) => [`row ${i}`]);
    const pdf = buildTablePdf({ title: 'T', columns: [{ header: 'A', width: 100 }], rows });
    expect(Number(/\/Count (\d+)/.exec(pdf)![1])).toBeGreaterThan(2);
  });

  it('escapes PDF syntax and keeps output ASCII', () => {
    expect(pdfSafe('₹ (A) \\ – “x”')).toBe('Rs. \\(A\\) \\\\ - "x"');
    expect(/[^\x00-\x7F]/.test(buildTablePdf({ title: '₹ lakh · test', columns: [{ header: 'A', width: 50 }], rows: [['é']] }))).toBe(false);
  });

  it('audit PDF has one row per action and no cell values', () => {
    recordApplied('cs-1234567890', 'Clean text in 2 cells', [
      { type: 'BATCH_SET', sheetName: 'T_Clean', operations: [{ address: 'A2', value: 'ZZ-SENTINEL-789' }, { address: 'A3', value: 'x' }] } as SheetAction,
    ]);
    const pdf = auditPdf();
    expect(pdf).toMatch(/CELLIX - Audit trail of accepted actions/);
    expect(pdf).toMatch(/T_Clean/);
    expect(pdf).toMatch(/1 entries \\\(1 applied, 0 reverted\\\)/);
    expect(pdf).not.toMatch(/ZZ-SENTINEL-789/);
  });
});
