import { RawTable } from '@/types/rawTable';
import { AttachmentError } from './attachmentError';
import { buildPdfRows, countPdfTextCharacters, PdfPageText } from './pdfRows';

/**
 * The part of pdf.js this decoder uses. Declared here so `extractPdfPages` can
 * be handed the library by its caller: the task pane loads it lazily with a
 * web worker, tests load it directly in Node.
 */
export interface PdfJsLike {
  getDocument(source: { data: Uint8Array; password?: string }): {
    promise: Promise<{
      numPages: number;
      getPage(pageNumber: number): Promise<{
        getTextContent(): Promise<{ items: unknown[] }>;
      }>;
    }>;
    destroy(): Promise<void>;
  };
}

interface PdfJsTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
}

function isTextItem(item: unknown): item is PdfJsTextItem {
  const candidate = item as Partial<PdfJsTextItem> | null;
  return (
    typeof candidate?.str === 'string' &&
    Array.isArray(candidate.transform) &&
    candidate.transform.length >= 6
  );
}

// pdf.js PasswordResponses: 1 = a password is needed, 2 = the one given is wrong.
const NEED_PASSWORD = 1;

/** Fewer characters per page than this means the pages are images, not text. */
const MIN_TEXT_CHARACTERS_PER_PAGE = 20;

/**
 * Reads the text of every page with its position. The password is used here
 * and nowhere else; it is never stored or sent.
 */
export async function extractPdfPages(
  pdfjs: PdfJsLike,
  data: Uint8Array,
  password?: string,
): Promise<PdfPageText[]> {
  const task = pdfjs.getDocument({ data, ...(password ? { password } : {}) });
  let doc: Awaited<typeof task.promise>;
  try {
    doc = await task.promise;
  } catch (error) {
    await task.destroy().catch(() => undefined);
    const failure = error as { name?: string; code?: number } | null;
    if (failure?.name === 'PasswordException') {
      throw failure.code === NEED_PASSWORD
        ? new AttachmentError('password_required', 'This PDF is password protected. Enter its password to read it.')
        : new AttachmentError('password_incorrect', 'That password did not open the PDF. Check it and try again.');
    }
    throw new AttachmentError('read_failed', 'This PDF could not be opened. It may be damaged.');
  }

  try {
    const pages: PdfPageText[] = [];
    for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber++) {
      const page = await doc.getPage(pageNumber);
      const content = await page.getTextContent();
      pages.push({
        page: pageNumber,
        items: content.items.filter(isTextItem).map((item) => ({
          str: item.str,
          x: item.transform[4],
          y: item.transform[5],
          width: item.width,
          height: item.height,
        })),
      });
    }
    return pages;
  } finally {
    await task.destroy().catch(() => undefined);
  }
}

export function pdfPagesToRawTable(pages: PdfPageText[], fileName: string): RawTable {
  if (countPdfTextCharacters(pages) < Math.max(1, pages.length) * MIN_TEXT_CHARACTERS_PER_PAGE) {
    throw new AttachmentError(
      'scanned_pdf',
      'This PDF is a scan or a photo, so it has no text to read. Download the statement from your bank as a PDF or Excel file instead.',
    );
  }
  return { source: 'pdf', fileName, layout: 'positioned', rows: buildPdfRows(pages) };
}

let pdfJsPromise: Promise<PdfJsLike> | null = null;

/**
 * Loads pdf.js on first use, so its size is paid only by someone who attaches a
 * PDF. The legacy build is used because the task pane's webview on older
 * Office hosts lacks the newest JavaScript features the default build needs.
 */
function loadPdfJs(): Promise<PdfJsLike> {
  if (!pdfJsPromise) {
    pdfJsPromise = (async () => {
      const [pdfjs, worker] = await Promise.all([
        import('pdfjs-dist/legacy/build/pdf.mjs'),
        import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url'),
      ]);
      pdfjs.GlobalWorkerOptions.workerSrc = worker.default;
      return pdfjs as unknown as PdfJsLike;
    })();
    // A failed load (offline, blocked worker) must not be cached as the answer.
    pdfJsPromise.catch(() => {
      pdfJsPromise = null;
    });
  }
  return pdfJsPromise;
}

export async function decodePdf(
  data: Uint8Array,
  fileName: string,
  password?: string,
): Promise<RawTable> {
  let pdfjs: PdfJsLike;
  try {
    pdfjs = await loadPdfJs();
  } catch {
    throw new AttachmentError('read_failed', 'The PDF reader could not be loaded. Check your connection and try again.');
  }
  return pdfPagesToRawTable(await extractPdfPages(pdfjs, data, password), fileName);
}
