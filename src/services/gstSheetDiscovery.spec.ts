import { describe, expect, it } from 'vitest';
import {
  buildMissingSheetMessage,
  discoverGstSheets,
  inferColumnMapping,
} from '@/services/gstSheetDiscovery';

describe('gstSheetDiscovery', () => {
  it('discovers PR and GSTR-2B from headers', () => {
    const result = discoverGstSheets(
      [
        {
          name: 'Purchases',
          headers: ['Supplier GSTIN', 'Invoice No', 'Date', 'Taxable Amount', 'CGST', 'SGST'],
        },
        {
          name: 'GSTR2B',
          headers: ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value'],
        },
      ],
      'GSTR2B',
    );
    expect(result.purchaseRegister?.name).toBe('Purchases');
    expect(result.portal?.name).toBe('GSTR2B');
    expect(result.missing).toEqual([]);
  });

  it('reports missing portal sheet', () => {
    const result = discoverGstSheets(
      [
        {
          name: 'PR',
          headers: ['GSTIN', 'Invoice No', 'Taxable Amount'],
        },
      ],
      'GSTR2B',
    );
    expect(result.missing).toContain('GSTR2B');
    expect(buildMissingSheetMessage(result.missing, 'PR vs 2B')).toMatch(/GSTR-2B/);
  });

  it('infers column mapping', () => {
    const m = inferColumnMapping(['Supplier GSTIN', 'Invoice No', 'Taxable Amount', 'IGST']);
    expect(m.gstin).toBe(0);
    expect(m.invoiceNo).toBe(1);
    expect(m.taxableAmt).toBe(2);
    expect(m.igst).toBe(3);
  });
});
