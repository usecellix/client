import { describe, expect, it, vi } from 'vitest';
import { AttachmentError } from './attachmentError';
import { extractPdfPages, PdfJsLike, pdfPagesToRawTable } from './pdfDecoder';
import { PdfPageText } from './pdfRows';

interface FakePage {
  items: unknown[];
}

/** Stands in for pdf.js: opens `pages` only when given `password` (if one is set). */
function fakePdfJs(pages: FakePage[], password?: string) {
  const destroy = vi.fn().mockResolvedValue(undefined);
  const getDocument = vi.fn((source: { data: Uint8Array; password?: string }) => {
    const reject = (code: number) =>
      Promise.reject(Object.assign(new Error('password'), { name: 'PasswordException', code }));
    let promise;
    if (password && !source.password) promise = reject(1);
    else if (password && source.password !== password) promise = reject(2);
    else {
      promise = Promise.resolve({
        numPages: pages.length,
        getPage: async (n: number) => ({ getTextContent: async () => ({ items: pages[n - 1].items }) }),
      });
    }
    return { promise, destroy };
  });
  return { pdfjs: { getDocument } as unknown as PdfJsLike, getDocument, destroy };
}

const textItem = (str: string, x: number, y: number, width = 20, height = 7) => ({
  str,
  transform: [1, 0, 0, 1, x, y],
  width,
  height,
});

async function codeOf(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    expect(error).toBeInstanceOf(AttachmentError);
    return (error as AttachmentError).code;
  }
  throw new Error('expected a rejection');
}

describe('extractPdfPages', () => {
  it('reads each page with the position of every piece of text', async () => {
    const { pdfjs } = fakePdfJs([{ items: [textItem('Date', 32.8, 454.7, 15)] }, { items: [textItem('x', 1, 2, 3)] }]);
    expect(await extractPdfPages(pdfjs, new Uint8Array([1]))).toEqual([
      { page: 1, items: [{ str: 'Date', x: 32.8, y: 454.7, width: 15, height: 7 }] },
      { page: 2, items: [{ str: 'x', x: 1, y: 2, width: 3, height: 7 }] },
    ]);
  });

  it('skips marked-content items, which carry no text', async () => {
    const { pdfjs } = fakePdfJs([{ items: [{ type: 'beginMarkedContent' }, textItem('kept', 1, 1)] }]);
    const pages = await extractPdfPages(pdfjs, new Uint8Array([1]));
    expect(pages[0].items.map((item) => item.str)).toEqual(['kept']);
  });

  it('asks for a password when the PDF needs one', async () => {
    const { pdfjs } = fakePdfJs([{ items: [] }], 'secret');
    expect(await codeOf(extractPdfPages(pdfjs, new Uint8Array([1])))).toBe('password_required');
  });

  it('says so when the password is wrong', async () => {
    const { pdfjs } = fakePdfJs([{ items: [] }], 'secret');
    expect(await codeOf(extractPdfPages(pdfjs, new Uint8Array([1]), 'guess'))).toBe('password_incorrect');
  });

  it('opens the PDF with the right password and passes it only to pdf.js', async () => {
    const { pdfjs, getDocument } = fakePdfJs([{ items: [textItem('ok', 1, 1)] }], 'secret');
    const pages = await extractPdfPages(pdfjs, new Uint8Array([1]), 'secret');
    expect(pages).toHaveLength(1);
    expect(getDocument).toHaveBeenCalledWith({ data: new Uint8Array([1]), password: 'secret' });
  });

  it('reports a file that is not a readable PDF', async () => {
    const pdfjs = {
      getDocument: () => ({
        promise: Promise.reject(Object.assign(new Error('bad'), { name: 'InvalidPDFException' })),
        destroy: async () => undefined,
      }),
    } as unknown as PdfJsLike;
    expect(await codeOf(extractPdfPages(pdfjs, new Uint8Array([1])))).toBe('read_failed');
  });

  it('releases the document whether or not it opened', async () => {
    const opened = fakePdfJs([{ items: [] }]);
    await extractPdfPages(opened.pdfjs, new Uint8Array([1]));
    expect(opened.destroy).toHaveBeenCalledTimes(1);

    const locked = fakePdfJs([{ items: [] }], 'secret');
    await codeOf(extractPdfPages(locked.pdfjs, new Uint8Array([1])));
    expect(locked.destroy).toHaveBeenCalledTimes(1);
  });
});

describe('pdfPagesToRawTable', () => {
  const textPage = (page: number): PdfPageText => ({
    page,
    items: [{ str: 'Statement of Account for the period', x: 10, y: 700, width: 200, height: 9 }],
  });

  it('builds a positioned raw table', () => {
    expect(pdfPagesToRawTable([textPage(1)], 'statement.pdf')).toEqual({
      source: 'pdf',
      fileName: 'statement.pdf',
      layout: 'positioned',
      rows: [
        {
          ref: 'p1 l1',
          page: 1,
          y: 700,
          cells: [{ t: 'Statement of Account for the period', x0: 10, x1: 210 }],
        },
      ],
    });
  });

  it('refuses a scanned PDF, which has pages but no text', () => {
    const scanned: PdfPageText[] = [
      { page: 1, items: [] },
      { page: 2, items: [{ str: '1', x: 0, y: 0, width: 1, height: 1 }] },
    ];
    expect(() => pdfPagesToRawTable(scanned, 'scan.pdf')).toThrowError(
      expect.objectContaining({ code: 'scanned_pdf' }),
    );
  });

  it('refuses a PDF with no pages', () => {
    expect(() => pdfPagesToRawTable([], 'empty.pdf')).toThrowError(expect.objectContaining({ code: 'scanned_pdf' }));
  });
});
