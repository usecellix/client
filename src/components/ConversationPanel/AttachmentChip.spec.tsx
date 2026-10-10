// @vitest-environment jsdom
import React from 'react';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import AttachmentChip from '@/components/ConversationPanel/AttachmentChip';
import { AttachmentError } from '@/services/attachments/attachmentError';
import { decodeAttachment } from '@/services/attachments/decodeAttachment';
import { RawTable } from '@/types/rawTable';

vi.mock('@/services/attachments/decodeAttachment', () => ({ decodeAttachment: vi.fn() }));
const decode = vi.mocked(decodeAttachment);

const file = new File([new Uint8Array(487_489)], 'statement.pdf', { type: 'application/pdf' });
const table: RawTable = { source: 'pdf', fileName: 'statement.pdf', layout: 'positioned', rows: [] };
const READY = 'Ready. Press Enter to import.';

type ChipProps = React.ComponentProps<typeof AttachmentChip>;

/** A decode that stays pending until the test settles it. */
function deferredDecode() {
  let resolve!: (value: RawTable) => void;
  decode.mockReturnValueOnce(
    new Promise<RawTable>((res) => {
      resolve = res;
    }),
  );
  return { resolve };
}

const passwordInput = () => screen.findByLabelText('Password for statement.pdf') as Promise<HTMLInputElement>;
const status = () => screen.getByRole('status').textContent;
const waitUntilReady = () => waitFor(() => expect(status()).toBe(READY));

function renderChip(props: Partial<ChipProps> = {}) {
  const onImport = vi.fn<(rawTable: RawTable) => boolean>(() => true);
  const all: ChipProps = { file, onRemove: vi.fn(), onImport, ...props };
  const view = render(<AttachmentChip {...all} />);
  /** Re-renders with Send pressed `count` times so far. */
  const pressSend = (count = 1, extra: Partial<ChipProps> = {}) =>
    view.rerender(<AttachmentChip {...all} {...extra} importRequest={count} />);
  return { ...view, onImport: all.onImport as typeof onImport, pressSend, all };
}

describe('AttachmentChip', () => {
  beforeEach(() => {
    decode.mockReset();
  });

  afterEach(() => {
    cleanup();
  });

  it('reads a file as soon as it is attached, then waits for Send', async () => {
    decode.mockResolvedValue(table);
    const onReadyChange = vi.fn();
    const { onImport } = renderChip({ onReadyChange });

    await waitUntilReady();
    expect(decode).toHaveBeenCalledTimes(1);
    expect(decode).toHaveBeenCalledWith(file, undefined);
    // Nothing is imported until the user sends, so they can type a question first.
    expect(onImport).not.toHaveBeenCalled();
    // The composer is told in an effect, which can land just after the capsule repaints.
    await waitFor(() => expect(onReadyChange).toHaveBeenLastCalledWith(true));
  });

  it('imports when Send is pressed', async () => {
    decode.mockResolvedValue(table);
    const onReadyChange = vi.fn();
    const { onImport, pressSend } = renderChip({ onReadyChange });
    await waitUntilReady();

    pressSend();
    await waitFor(() => expect(onImport).toHaveBeenCalledWith(table));
    expect(onImport).toHaveBeenCalledTimes(1);
    expect(onReadyChange).toHaveBeenLastCalledWith(false);
    // Read once, not again on Send.
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('shows that it is being read', () => {
    deferredDecode();
    renderChip();
    expect(screen.getByText('statement.pdf')).toBeTruthy();
    expect(status()).toBe('Reading…');
  });

  it('shows type and size, and reads nothing, when there is nothing to import into', () => {
    render(<AttachmentChip file={file} onRemove={vi.fn()} />);
    expect(status()).toBe('PDF · 476 KB');
    expect(decode).not.toHaveBeenCalled();
  });

  it('is a capsule holding only an icon, the name and Remove, with the details in its tooltip', () => {
    deferredDecode();
    const { container } = renderChip();
    const chip = container.querySelector('.cellix-attachment-chip') as HTMLElement;
    expect(chip.getAttribute('title')).toBe('statement.pdf (PDF · 476 KB) — Reading…');
    expect(chip.querySelectorAll('button')).toHaveLength(1);
    // The status is for screen readers only; it takes no room in the capsule.
    expect(screen.getByRole('status').className).toBe('cellix-sr-only');
  });

  it('shows the file-type icon once read: a red document for a PDF, no lock', async () => {
    decode.mockResolvedValue(table);
    renderChip();
    await waitUntilReady();
    expect(document.querySelector('.cellix-attachment-icon.is-pdf .lucide-file-text')).not.toBeNull();
    expect(document.querySelector('.cellix-attachment-icon [class*="lucide-lock"]')).toBeNull();
  });

  it('shows a sheet icon for a spreadsheet', async () => {
    decode.mockResolvedValue({ ...table, source: 'csv', layout: 'grid' });
    render(<AttachmentChip file={new File(['a,b'], 'statement.csv')} onRemove={vi.fn()} onImport={vi.fn(() => true)} />);
    await waitUntilReady();
    expect(document.querySelector('.cellix-attachment-icon.is-sheet')).not.toBeNull();
    expect(document.querySelector('.cellix-attachment-icon.is-pdf')).toBeNull();
  });

  it('asks for the password of a protected PDF, with the cursor already in the field', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'This PDF is password protected.'));
    const { onImport } = renderChip();

    const input = await passwordInput();
    expect(input.type).toBe('password');
    expect(document.activeElement).toBe(input);
    expect(status()).toBe('Password protected');
    expect(screen.getByRole('button', { name: 'Unlock' })).toBeTruthy();
    expect(onImport).not.toHaveBeenCalled();
  });

  it('only unlocks the PDF on Enter in the password field: nothing is imported yet', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    decode.mockResolvedValueOnce(table);
    const onReadyChange = vi.fn();
    const { onImport } = renderChip({ onReadyChange });

    const input = await passwordInput();
    fireEvent.change(input, { target: { value: 'secret' } });
    fireEvent.submit(input.closest('form') as HTMLFormElement);
    await waitUntilReady();

    expect(decode).toHaveBeenLastCalledWith(file, 'secret');
    expect(onImport).not.toHaveBeenCalled();
    expect(screen.queryByLabelText('Password for statement.pdf')).toBeNull();
    // The composer is told in an effect, which can land just after the capsule repaints.
    await waitFor(() => expect(onReadyChange).toHaveBeenLastCalledWith(true));
  });

  it('is not released by a Send pressed before it was unlocked', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    decode.mockResolvedValueOnce(table);
    const { onImport, pressSend } = renderChip();
    const input = await passwordInput();

    // Send is pressed while the file is still locked (to send a chat message, say).
    pressSend(1);

    fireEvent.change(input, { target: { value: 'secret' } });
    fireEvent.submit(input.closest('form') as HTMLFormElement);
    await waitUntilReady();
    expect(onImport).not.toHaveBeenCalled();

    pressSend(2);
    await waitFor(() => expect(onImport).toHaveBeenCalledTimes(1));
  });

  it('says so and clears the field when the password is wrong', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    decode.mockRejectedValueOnce(new AttachmentError('password_incorrect', 'That password did not open the PDF.'));
    const { onImport } = renderChip();

    const input = await passwordInput();
    fireEvent.change(input, { target: { value: 'guess' } });
    fireEvent.submit(input.closest('form') as HTMLFormElement);

    expect((await screen.findByRole('alert')).textContent).toBe('That password did not open the PDF.');
    const retry = await passwordInput();
    expect(retry.value).toBe('');
    expect(retry.placeholder).toBe('Wrong password, try again');
    expect(retry.getAttribute('aria-invalid')).toBe('true');
    expect(onImport).not.toHaveBeenCalled();

    // Typing again clears the warning.
    fireEvent.change(retry, { target: { value: 's' } });
    expect(screen.queryByRole('alert')).toBeNull();
  });

  it('does not submit an empty password', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    renderChip();

    const input = await passwordInput();
    fireEvent.submit(input.closest('form') as HTMLFormElement);
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('shows why a file could not be read, keeps the chip, and can retry', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('scanned_pdf', 'This PDF is a scan or a photo.'));
    decode.mockResolvedValueOnce(table);
    const onReadyChange = vi.fn();
    renderChip({ onReadyChange });

    expect((await screen.findByRole('alert')).textContent).toBe('This PDF is a scan or a photo.');
    expect(screen.getByText('statement.pdf')).toBeTruthy();
    // A file that could not be read is never offered for sending.
    expect(onReadyChange).not.toHaveBeenCalledWith(true);

    fireEvent.click(screen.getByRole('button', { name: 'Retry' }));
    await waitUntilReady();
  });

  it('gives a plain message for an unexpected failure', async () => {
    decode.mockRejectedValue(new Error('TypeError: x is undefined'));
    renderChip();
    expect((await screen.findByRole('alert')).textContent).toBe('This file could not be read.');
  });

  it('waits while the assistant is busy, then imports when it is free', async () => {
    decode.mockResolvedValue(table);
    const { onImport, pressSend } = renderChip();
    await waitUntilReady();

    pressSend(1, { disabled: true });
    await waitFor(() => expect(status()).toBe('Importing…'));
    expect(onImport).not.toHaveBeenCalled();

    pressSend(1, { disabled: false });
    await waitFor(() => expect(onImport).toHaveBeenCalledWith(table));
  });

  it('keeps the file and offers it again when the import cannot start yet', async () => {
    decode.mockResolvedValue(table);
    const onImport = vi.fn<(rawTable: RawTable) => boolean>().mockReturnValueOnce(false).mockReturnValue(true);
    const { pressSend } = renderChip({ onImport });
    await waitUntilReady();

    pressSend(1);
    await waitFor(() => expect(onImport).toHaveBeenCalledTimes(1));

    // Another import is running: the assistant turns busy, then free again.
    pressSend(1, { disabled: true });
    pressSend(1, { disabled: false });
    await waitFor(() => expect(onImport).toHaveBeenCalledTimes(2));
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('tells the composer it is no longer waiting when it is removed', async () => {
    decode.mockResolvedValue(table);
    const onReadyChange = vi.fn();
    const { unmount } = renderChip({ onReadyChange });
    await waitUntilReady();
    unmount();
    expect(onReadyChange).toHaveBeenLastCalledWith(false);
  });

  it('can be removed while it is being read, and then imports nothing', async () => {
    const pending = deferredDecode();
    const onRemove = vi.fn();
    const { unmount, onImport } = renderChip({ onRemove });

    fireEvent.click(screen.getByRole('button', { name: 'Remove statement.pdf' }));
    expect(onRemove).toHaveBeenCalledTimes(1);

    unmount();
    pending.resolve(table);
    await Promise.resolve();
    expect(onImport).not.toHaveBeenCalled();
  });
});
