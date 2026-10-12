// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { cleanup, createEvent, fireEvent, render, screen, waitFor } from '@testing-library/react';
import ConversationPanel from '@/components/ConversationPanel/ConversationPanel';
import { decodeAttachment } from '@/services/attachments/decodeAttachment';
import { RawTable } from '@/types/rawTable';

vi.mock('@/services/attachments/decodeAttachment', () => ({ decodeAttachment: vi.fn() }));
vi.mock('@/components/ConversationPanel/PanelHeader', () => ({ default: () => null }));
vi.mock('@/auth/auth-client', () => ({ useSession: () => ({ data: null }) }));
const decode = vi.mocked(decodeAttachment);

const table: RawTable = { source: 'pdf', fileName: 'statement.pdf', layout: 'positioned', rows: [] };
const pdf = new File(['%PDF-'], 'statement.pdf', { type: 'application/pdf' });

function renderPanel() {
  const props = {
    sessions: [],
    activeSessionId: 's1',
    turns: [],
    activeTurnId: null,
    isWaitingForResponse: false,
    isWaitingClarification: false,
    activeClarification: null,
    previewEnabled: true,
    mode: 'action',
    onModeChange: vi.fn(),
    onImportAttachment: vi.fn(),
    onSend: vi.fn(),
  };
  // Only what the drop path touches is supplied; the rest is not rendered.
  const view = render(<ConversationPanel {...({ ...props } as unknown as React.ComponentProps<typeof ConversationPanel>)} />);
  const pane = view.container.querySelector('.cellix-panel') as HTMLElement;
  return { ...view, pane };
}

/** A drag event the way a webview builds it for files from the desktop. */
const filesTransfer = (files: File[]) => ({
  types: ['Files'],
  files,
  items: files.map((file) => ({ kind: 'file', type: file.type })),
  dropEffect: 'none',
});

describe('ConversationPanel — dropping a file', () => {
  // jsdom has no scrollTo; the panel scrolls its thread when turns change.
  beforeAll(() => {
    Element.prototype.scrollTo = vi.fn() as unknown as Element['scrollTo'];
  });

  afterEach(() => {
    cleanup();
    decode.mockReset();
  });

  it('shows the drop overlay while a file is dragged over the pane, and hides it on leave', () => {
    const { pane } = renderPanel();
    expect(screen.queryByText('Drop the PDF to attach it')).toBeNull();

    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([pdf]) });
    expect(screen.getByText('Drop the PDF to attach it')).toBeTruthy();

    fireEvent.dragLeave(pane, { dataTransfer: filesTransfer([pdf]) });
    expect(screen.queryByText('Drop the PDF to attach it')).toBeNull();
  });

  it('does not react to a drag that carries no file, such as selected text', () => {
    const { pane } = renderPanel();
    fireEvent.dragEnter(pane, { dataTransfer: { types: ['text/plain'], files: [] } });
    expect(screen.queryByText('Drop the PDF to attach it')).toBeNull();
  });

  it('keeps the overlay while the pointer crosses child elements', () => {
    const { pane } = renderPanel();
    const child = pane.querySelector('.cellix-input-area') as HTMLElement;

    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([pdf]) });
    fireEvent.dragEnter(child, { dataTransfer: filesTransfer([pdf]) });
    fireEvent.dragLeave(pane, { dataTransfer: filesTransfer([pdf]) });
    expect(screen.getByText('Drop the PDF to attach it')).toBeTruthy();

    fireEvent.dragLeave(child, { dataTransfer: filesTransfer([pdf]) });
    expect(screen.queryByText('Drop the PDF to attach it')).toBeNull();
  });

  it('cancels dragover for files, so the webview does not open the dropped file', () => {
    const { pane } = renderPanel();
    const over = createEvent.dragOver(pane, { dataTransfer: filesTransfer([pdf]) });
    fireEvent(pane, over);
    expect(over.defaultPrevented).toBe(true);
  });

  it('attaches a file dropped anywhere on the pane and hides the overlay', async () => {
    decode.mockResolvedValue(table);
    const { pane } = renderPanel();

    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([pdf]) });
    const drop = createEvent.drop(pane, { dataTransfer: filesTransfer([pdf]) });
    fireEvent(pane, drop);

    expect(drop.defaultPrevented).toBe(true);
    expect(screen.queryByText('Drop the PDF to attach it')).toBeNull();
    await waitFor(() => expect(screen.getByRole('status').textContent).toBe('Ready. Press Enter to import.'));
    expect(decode).toHaveBeenCalledTimes(1);
  });

  it('refuses a dropped file of the wrong type with a message', () => {
    const { pane } = renderPanel();
    const docx = new File(['x'], 'Account_Statement.docx');

    fireEvent.drop(pane, { dataTransfer: filesTransfer([docx]) });
    expect(screen.getByRole('alert').textContent).toContain('Account_Statement.docx cannot be attached');
    expect(decode).not.toHaveBeenCalled();
  });

  it('shows a PDF icon and says so while a PDF is dragged over', () => {
    const { pane, container } = renderPanel();
    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([pdf]) });

    expect(screen.getByText('Drop the PDF to attach it')).toBeTruthy();
    expect(container.querySelector('.cellix-drop-overlay-icons .is-pdf')).not.toBeNull();
    expect(container.querySelector('.cellix-drop-overlay-icons .is-sheet')).toBeNull();
  });

  it('shows a sheet icon for a spreadsheet', () => {
    const { pane, container } = renderPanel();
    const xlsx = new File(['x'], 'a.xlsx', {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([xlsx]) });

    expect(container.querySelector('.cellix-drop-overlay-icons .is-sheet')).not.toBeNull();
    expect(container.querySelector('.cellix-drop-overlay-icons .is-pdf')).toBeNull();
  });

  it('warns, and shows the not-allowed cursor, when only an unsupported type is dragged', () => {
    const { pane } = renderPanel();
    const docx = new File(['x'], 'a.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([docx]) });
    expect(screen.getByText("This file type can't be attached")).toBeTruthy();

    const dataTransfer = filesTransfer([docx]);
    fireEvent(pane, createEvent.dragOver(pane, { dataTransfer }));
    expect(dataTransfer.dropEffect).toBe('none');
  });

  it('lets a drag of unknown type through, and sorts it out by name on drop', () => {
    const { pane } = renderPanel();
    const csv = new File(['d'], 'a.csv', { type: '' });
    fireEvent.dragEnter(pane, { dataTransfer: filesTransfer([csv]) });
    expect(screen.getByText('Drop to attach')).toBeTruthy();

    const dataTransfer = filesTransfer([csv]);
    fireEvent(pane, createEvent.dragOver(pane, { dataTransfer }));
    expect(dataTransfer.dropEffect).toBe('copy');
  });
});
