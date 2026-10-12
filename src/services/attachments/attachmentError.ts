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

/**
 * What a file being dragged is, judged from its MIME type: a browser hides the
 * name until the file is dropped, but tells the type. A blank or unfamiliar type
 * is 'unknown' rather than refused, because Windows reports a CSV as Excel or as
 * nothing; the name is checked again on drop (`attachmentExtension`).
 */
export type DraggedFileKind = 'pdf' | 'sheet' | 'unsupported' | 'unknown';

const SHEET_MIME_TYPES = [
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'text/csv',
  'application/csv',
];

export function draggedFileKind(mimeType: string): DraggedFileKind {
  const mime = mimeType.toLowerCase();
  if (!mime) return 'unknown';
  if (mime === 'application/pdf') return 'pdf';
  if (SHEET_MIME_TYPES.includes(mime)) return 'sheet';
  // Anything with a definite, different type (Word, an image, a zip...) is not a statement.
  return 'unsupported';
}
