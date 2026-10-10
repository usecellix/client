import { RawTable } from '@/types/rawTable';
import {
  AttachmentError,
  attachmentExtension,
  MAX_ATTACHMENT_BYTES,
  SUPPORTED_ATTACHMENT_EXTENSIONS,
} from './attachmentError';
import { decodeGrid } from './gridDecoder';
import { decodePdf } from './pdfDecoder';

/** The slice of the browser `File` this needs, so tests can pass a plain object. */
export interface AttachmentFile {
  name: string;
  size: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}

/**
 * Turns an attached file into a raw table, entirely in the task pane.
 *
 * Nothing leaves the machine here. The file is read into memory, decoded and
 * dropped; only the decoded table is later sent to the server. `password` is
 * handed to the PDF reader for this one call and kept nowhere.
 *
 * Throws `AttachmentError`, whose message is written for the user.
 */
export async function decodeAttachment(file: AttachmentFile, password?: string): Promise<RawTable> {
  const extension = attachmentExtension(file.name);
  if (!extension) {
    throw new AttachmentError(
      'unsupported_type',
      `This file type is not supported. Attach a ${SUPPORTED_ATTACHMENT_EXTENSIONS.map((e) => e.toUpperCase()).join(', ')} file.`,
    );
  }
  if (file.size > MAX_ATTACHMENT_BYTES) {
    throw new AttachmentError('too_large', 'This file is larger than 20 MB, which is too big to be a statement.');
  }
  if (file.size === 0) {
    throw new AttachmentError('empty', 'This file is empty.');
  }

  let data: Uint8Array;
  try {
    // A fresh read each call: the PDF reader takes ownership of the bytes it is given.
    data = new Uint8Array(await file.arrayBuffer());
  } catch {
    throw new AttachmentError('read_failed', 'This file could not be read from your computer.');
  }

  return extension === 'pdf'
    ? decodePdf(data, file.name, password)
    : decodeGrid(data, file.name, extension);
}
