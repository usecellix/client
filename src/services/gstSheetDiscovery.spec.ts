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
    expect(result.missing).toEqual(expect.arrayContaining(['GSTR2B', 'GSTR2A']));
    expect(buildMissingSheetMessage(result.missing, 'PR vs 2B')).toMatch(/GSTR-2B/);
  });

  it('discovers both GSTR-2B and GSTR-2A when present', () => {
    const result = discoverGstSheets(
      [
        {
          name: 'PR',
          headers: ['Supplier GSTIN', 'Invoice No', 'Date', 'Taxable Amount', 'CGST', 'SGST'],
        },
        {
          name: 'GSTR2B',
          headers: ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value'],
        },
        {
          name: 'GSTR2A',
          headers: ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value'],
        },
      ],
      'GSTR2B',
    );
    expect(result.gstr2b?.name).toBe('GSTR2B');
    expect(result.gstr2a?.name).toBe('GSTR2A');
    expect(result.missing).toEqual([]);
  });

  it('treats GSTR-2A alone as enough for purchase portal', () => {
    const result = discoverGstSheets(
      [
        {
          name: 'PR',
          headers: ['Supplier GSTIN', 'Invoice No', 'Date', 'Taxable Amount', 'CGST'],
        },
        {
          name: 'GSTR2A',
          headers: ['GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Taxable Value'],
        },
      ],
      'GSTR2B',
    );
    expect(result.gstr2a?.name).toBe('GSTR2A');
    expect(result.portal?.name).toBe('GSTR2A');
    expect(result.missing).toEqual([]);
  });

  it('infers column mapping', () => {
    const m = inferColumnMapping(['Supplier GSTIN', 'Invoice No', 'Taxable Amount', 'IGST']);
    expect(m.gstin).toBe(0);
    expect(m.invoiceNo).toBe(1);
    expect(m.taxableAmt).toBe(2);
    expect(m.igst).toBe(3);
  });
});
