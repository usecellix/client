import { describe, expect, it, vi } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import React from 'react';
import { SourcePreview } from '@/components/SourcePreview/SourcePreview';
import { hasExceptionFlags, formatSourceRefLabel } from '@/types/changeSet';

describe('SourcePreview', () => {
  it('renders exception markers distinctly from clean citations', () => {
    const html = renderToStaticMarkup(
      React.createElement(SourcePreview, {
        sourceRefs: [
          {
            documentType: 'workbook',
            documentId: 'wb',
            rowOrLine: 'Sheet2!C4:C40',
          },
        ],
        exceptionFlags: [
          {
            code: 'GST_NAME_FUZZY_MATCH',
            severity: 'flag',
            message: 'Vendor name matched fuzzily',
            affectedRows: [2],
          },
        ],
        onJumpToSource: vi.fn(),
      }),
    );

    expect(html).toContain('source-exceptions');
    expect(html).toContain('cellix-source-exception-flag');
    expect(html).toContain('GST_NAME_FUZZY_MATCH');
    expect(html).toContain('Sheet2!C4:C40');
  });

  it('marks CellChange with exceptionFlags as flagged', () => {
    expect(
      hasExceptionFlags({
        cell: 'A1',
        sheet: 'Sheet1',
        before: 1,
        after: 2,
        isHardcoded: false,
        exceptionFlags: [
          {
            code: 'GST_NAME_FUZZY_MATCH',
            severity: 'flag',
            message: 'flag',
            affectedRows: [1],
          },
        ],
      }),
    ).toBe(true);

    expect(
      hasExceptionFlags({
        cell: 'A1',
        sheet: 'Sheet1',
        before: 1,
        after: 2,
        isHardcoded: false,
      }),
    ).toBe(false);
  });

  it('formats workbook source labels for jump targets', () => {
    expect(
      formatSourceRefLabel({
        documentType: 'workbook',
        documentId: 'wb',
        rowOrLine: 'Sheet2!C4:C40',
      }),
    ).toBe('Sheet2!C4:C40');
  });
});
