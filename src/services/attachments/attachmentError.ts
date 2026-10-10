export type AttachmentErrorCode =
  | 'unsupported_type'
  | 'too_large'
  | 'password_required'
  | 'password_incorrect'
  | 'scanned_pdf'
  | 'empty'
  | 'read_failed';

/** A file that could not be decoded, with a message written for the user. */
export class AttachmentError extends Error {
  constructor(
    readonly code: AttachmentErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'AttachmentError';
  }
}

export const SUPPORTED_ATTACHMENT_EXTENSIONS = ['pdf', 'xlsx', 'xls', 'csv'] as const;
export type SupportedAttachmentExtension = (typeof SUPPORTED_ATTACHMENT_EXTENSIONS)[number];

/** The picker's `accept` list. Kept next to the decoders so the two cannot drift apart. */
export const ATTACHMENT_ACCEPT = SUPPORTED_ATTACHMENT_EXTENSIONS.map((ext) => `.${ext}`).join(',');

/** 20 MB. A year of statements is well under 2 MB; this only stops a wrong file being read into memory. */
export const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;

export function attachmentExtension(fileName: string): SupportedAttachmentExtension | null {
  const ext = fileName.toLowerCase().split('.').pop() ?? '';
  return (SUPPORTED_ATTACHMENT_EXTENSIONS as readonly string[]).includes(ext)
    ? (ext as SupportedAttachmentExtension)
    : null;
}
