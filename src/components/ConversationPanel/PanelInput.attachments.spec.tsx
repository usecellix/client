// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { ImportAttachmentHandler, PanelInput } from '@/components/ConversationPanel/ConversationPanel';
import { AttachmentError } from '@/services/attachments/attachmentError';
import { decodeAttachment } from '@/services/attachments/decodeAttachment';
import { RawTable } from '@/types/rawTable';

vi.mock('@/services/attachments/decodeAttachment', () => ({ decodeAttachment: vi.fn() }));
const decode = vi.mocked(decodeAttachment);

const pdf = new File(['%PDF-'], 'statement.pdf', { type: 'application/pdf' });
const table: RawTable = { source: 'pdf', fileName: 'statement.pdf', layout: 'positioned', rows: [] };
const READY = 'Ready. Press Enter to import.';
const QUESTION = 'Which are the 10 largest debits?';

/** How the composer behaves around an attached statement. ATTACHMENT_EXTRACTION_PLAN.md §6.5. */
describe('PanelInput — attachments', () => {
  let onSend: ReturnType<typeof vi.fn>;
  let onImportAttachment: ReturnType<typeof vi.fn<ImportAttachmentHandler>>;

  function renderComposer() {
    const view = render(
      <PanelInput onSend={onSend} mode="action" onModeChange={vi.fn()} onImportAttachment={onImportAttachment} />,
    );
    const attach = (...files: File[]) =>
      fireEvent.change(view.container.querySelector('input[type="file"]') as HTMLInputElement, {
        target: { files },
      });
    const textarea = view.container.querySelector('textarea') as HTMLTextAreaElement;
    const sendButton = () => view.container.querySelector('.cellix-send-btn') as HTMLButtonElement;
    const type = (text: string) => fireEvent.change(textarea, { target: { value: text } });
    const pressEnter = () => fireEvent.keyDown(textarea, { key: 'Enter' });
    return { ...view, attach, textarea, sendButton, type, pressEnter };
  }

  const waitUntilReady = () => waitFor(() => expect(screen.getByRole('status').textContent).toBe(READY));

  async function unlockWith(password: string) {
    const input = (await screen.findByLabelText('Password for statement.pdf')) as HTMLInputElement;
    fireEvent.change(input, { target: { value: password } });
    fireEvent.submit(input.closest('form') as HTMLFormElement);
    await waitUntilReady();
  }

  beforeEach(() => {
    decode.mockReset();
    onSend = vi.fn();
    onImportAttachment = vi.fn<ImportAttachmentHandler>();
  });

  afterEach(() => cleanup());

  it('holds an attached file in the box until the user sends', async () => {
    decode.mockResolvedValue(table);
    const { attach, textarea, sendButton } = renderComposer();

    attach(pdf);
    expect(sendButton().disabled).toBe(true);
    await waitUntilReady();

    expect(onImportAttachment).not.toHaveBeenCalled();
    expect(screen.getByText('statement.pdf')).toBeTruthy();
    // The message box has the focus and says what can be done; Send is live with no text.
    expect(document.activeElement).toBe(textarea);
    expect(textarea.placeholder).toBe('Ask about this statement, or press Enter to import it');
    expect(sendButton().disabled).toBe(false);
  });

  it('imports the file on Enter when nothing was typed', async () => {
    decode.mockResolvedValue(table);
    const { attach, pressEnter } = renderComposer();
    attach(pdf);
    await waitUntilReady();

    pressEnter();
    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledWith(table, 'statement.pdf', undefined));
    expect(onImportAttachment).toHaveBeenCalledTimes(1);
    expect(onSend).not.toHaveBeenCalled();
    // The capsule goes once the import has started.
    await waitFor(() => expect(screen.queryByText('statement.pdf')).toBeNull());
  });

  it('sends a typed question with the file, and clears the box', async () => {
    decode.mockResolvedValue(table);
    const { attach, type, pressEnter, textarea } = renderComposer();
    attach(pdf);
    await waitUntilReady();

    type(QUESTION);
    pressEnter();

    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledWith(table, 'statement.pdf', QUESTION));
    // The question travels with the import; it is not sent as a separate chat message.
    expect(onSend).not.toHaveBeenCalled();
    expect(textarea.value).toBe('');
  });

  it('does the same from the Send button', async () => {
    decode.mockResolvedValue(table);
    const { attach, type, sendButton } = renderComposer();
    attach(pdf);
    await waitUntilReady();

    type(QUESTION);
    fireEvent.click(sendButton());
    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledWith(table, 'statement.pdf', QUESTION));
  });

  it('gives the question back when the import fails', async () => {
    decode.mockResolvedValue(table);
    onImportAttachment.mockResolvedValue(false);
    const { attach, type, pressEnter, textarea } = renderComposer();
    attach(pdf);
    await waitUntilReady();

    type(QUESTION);
    pressEnter();
    await waitFor(() => expect(textarea.value).toBe(QUESTION));
  });

  it('leaves the box empty when the import succeeds', async () => {
    decode.mockResolvedValue(table);
    onImportAttachment.mockResolvedValue(true);
    const { attach, type, pressEnter, textarea } = renderComposer();
    attach(pdf);
    await waitUntilReady();

    type(QUESTION);
    pressEnter();
    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledTimes(1));
    await Promise.resolve();
    expect(textarea.value).toBe('');
  });

  it('only unlocks a protected PDF on Enter in the password field', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    decode.mockResolvedValueOnce(table);
    const { attach, textarea, sendButton } = renderComposer();

    attach(pdf);
    await unlockWith('secret');

    expect(onImportAttachment).not.toHaveBeenCalled();
    expect(screen.getByText('statement.pdf')).toBeTruthy();
    expect(document.activeElement).toBe(textarea);
    expect(sendButton().disabled).toBe(false);
  });

  it('sends an unlocked PDF with its question on the main Enter', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    decode.mockResolvedValueOnce(table);
    const { attach, type, pressEnter } = renderComposer();
    attach(pdf);
    await unlockWith('secret');

    type(QUESTION);
    pressEnter();
    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledWith(table, 'statement.pdf', QUESTION));
  });

  it('still sends a plain chat message while a file is waiting for its password', async () => {
    decode.mockRejectedValueOnce(new AttachmentError('password_required', 'Needs a password.'));
    const { attach, type, pressEnter } = renderComposer();
    attach(pdf);
    await screen.findByLabelText('Password for statement.pdf');

    type('hello');
    pressEnter();

    expect(onSend).toHaveBeenCalledWith('hello');
    expect(onImportAttachment).not.toHaveBeenCalled();
    // The locked file is still attached.
    expect(screen.getByText('statement.pdf')).toBeTruthy();
  });

  it('gives the question to the first file only when two are attached', async () => {
    decode.mockResolvedValue(table);
    // The first import never finishes here, so the second file has to wait its turn.
    onImportAttachment.mockReturnValue(new Promise<boolean>(() => undefined));
    const { attach, type, pressEnter } = renderComposer();
    attach(pdf, new File(['a,b'], 'other.csv'));
    await waitFor(() => expect(screen.getAllByRole('status').map((el) => el.textContent)).toEqual([READY, READY]));

    type(QUESTION);
    pressEnter();

    await waitFor(() => expect(onImportAttachment).toHaveBeenCalledTimes(1));
    expect(onImportAttachment).toHaveBeenCalledWith(table, 'statement.pdf', QUESTION);
    // The second file is still in the box, waiting.
    expect(screen.getByText('other.csv')).toBeTruthy();
  });

  it('refuses a file type it cannot read, with a message', () => {
    const { attach } = renderComposer();
    attach(new File(['x'], 'notes.docx'));
    expect(screen.getByRole('alert').textContent).toBe(
      'notes.docx cannot be attached. Use a PDF, XLSX, XLS or CSV file.',
    );
    expect(decode).not.toHaveBeenCalled();
  });

  describe('files dropped on the pane', () => {
    const noop = vi.fn();
    const renderWithDrop = (droppedFiles: { id: number; files: File[] } | null, disabled = false) =>
      render(
        <PanelInput
          onSend={onSend}
          mode="action"
          onModeChange={noop}
          onImportAttachment={onImportAttachment}
          droppedFiles={droppedFiles}
          disabled={disabled}
        />,
      );

    it('attaches a dropped statement exactly as the paperclip would', async () => {
      decode.mockResolvedValue(table);
      renderWithDrop({ id: 1, files: [pdf] });

      await waitUntilReady();
      expect(decode).toHaveBeenCalledTimes(1);
      expect(onImportAttachment).not.toHaveBeenCalled();
    });

    it('refuses a dropped file of another type with the same message as the picker', () => {
      const docx = new File(['x'], 'Account_Statement.docx');
      renderWithDrop({ id: 1, files: [docx] });

      expect(screen.getByRole('alert').textContent).toBe(
        'Account_Statement.docx cannot be attached. Use a PDF, XLSX, XLS or CSV file.',
      );
      expect(decode).not.toHaveBeenCalled();
    });

    it('takes the supported file and names the refused one when both are dropped', async () => {
      decode.mockResolvedValue(table);
      renderWithDrop({ id: 1, files: [pdf, new File(['x'], 'notes.txt')] });

      await waitUntilReady();
      expect(screen.getByRole('alert').textContent).toContain('notes.txt cannot be attached');
    });

    it('ignores a drop while the box is disabled', () => {
      renderWithDrop({ id: 1, files: [pdf] }, true);

      expect(decode).not.toHaveBeenCalled();
      expect(screen.queryByRole('status')).toBeNull();
    });

    it('handles each drop once, even when the composer re-renders', async () => {
      decode.mockResolvedValue(table);
      const drop = { id: 1, files: [pdf] };
      const view = renderWithDrop(drop);
      await waitUntilReady();

      view.rerender(
        <PanelInput
          onSend={onSend}
          mode="action"
          onModeChange={noop}
          onImportAttachment={onImportAttachment}
          droppedFiles={drop}
        />,
      );
      expect(decode).toHaveBeenCalledTimes(1);
    });
  });
});
