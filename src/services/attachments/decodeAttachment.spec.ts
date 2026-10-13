import { describe, expect, it, vi } from 'vitest';
import { ATTACHMENT_ACCEPT, AttachmentError, attachmentExtension } from './attachmentError';
import { AttachmentFile, decodeAttachment } from './decodeAttachment';

const file = (name: string, content: string | number): AttachmentFile => {
  const bytes = typeof content === 'string' ? new TextEncoder().encode(content) : new Uint8Array(0);
  return {
    name,
    size: typeof content === 'number' ? content : bytes.byteLength,
    arrayBuffer: vi.fn(async () => bytes.buffer as ArrayBuffer),
  };
};

async function codeOf(promise: Promise<unknown>): Promise<string> {
  try {
    await promise;
  } catch (error) {
    expect(error).toBeInstanceOf(AttachmentError);
    return (error as AttachmentError).code;
  }
  throw new Error('expected a rejection');
}

describe('decodeAttachment', () => {
  it.each(['notes.docx', 'notes.doc', 'notes.txt', 'photo.png', 'no-extension'])(
    'refuses %s without reading it',
    async (name) => {
      const attached = file(name, 'content');
      expect(await codeOf(decodeAttachment(attached))).toBe('unsupported_type');
      expect(attached.arrayBuffer).not.toHaveBeenCalled();
    },
  );

  it('refuses a file over 20 MB without reading it', async () => {
    const attached = file('statement.pdf', 20 * 1024 * 1024 + 1);
    expect(await codeOf(decodeAttachment(attached))).toBe('too_large');
    expect(attached.arrayBuffer).not.toHaveBeenCalled();
  });

  it('refuses an empty file', async () => {
    expect(await codeOf(decodeAttachment(file('statement.csv', 0)))).toBe('empty');
  });

  it('reports a file the browser cannot read', async () => {
    const attached: AttachmentFile = {
      name: 'statement.csv',
      size: 10,
      arrayBuffer: async () => {
        throw new Error('NotReadableError');
      },
    };
    expect(await codeOf(decodeAttachment(attached))).toBe('read_failed');
  });

  it('decodes a CSV into a grid table', async () => {
    const table = await decodeAttachment(file('Statement.CSV', 'Date,Debit,Balance\n01/04/2026,10.00,90.00\n'));
    expect(table).toEqual({
      source: 'csv',
      fileName: 'Statement.CSV',
      layout: 'grid',
      rows: [
        { ref: 'row 1', cells: [{ t: 'Date' }, { t: 'Debit' }, { t: 'Balance' }] },
        { ref: 'row 2', cells: [{ t: '01/04/2026' }, { t: '10.00' }, { t: '90.00' }] },
      ],
    });
  });
});

describe('supported types', () => {
  it('offers exactly the types that can be decoded', () => {
    expect(ATTACHMENT_ACCEPT).toBe('.pdf,.xlsx,.xls,.csv');
  });

  it('reads the extension without regard to case', () => {
    expect(attachmentExtension('A.B.PDF')).toBe('pdf');
    expect(attachmentExtension('statement.Xlsx')).toBe('xlsx');
    expect(attachmentExtension('statement.pdf.exe')).toBeNull();
  });
});
