// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import { GstReconCollisionCard } from '@/components/ConversationPanel/GstReconCollisionCard';
import { GstReconCollisionBlock } from '@/types/conversationTurn';

function makeBlock(overrides: Partial<GstReconCollisionBlock> = {}): GstReconCollisionBlock {
  return {
    id: 'gst_recon_collision_collision_1',
    type: 'gst_recon_collision',
    collisionId: 'collision_1',
    sheetName: 'Missed vs GSTR-2B',
    ...overrides,
  };
}

describe('GstReconCollisionCard', () => {
  afterEach(() => {
    cleanup();
  });

  it('renders the sheet name and three buttons instead of a typed-reply prompt', () => {
    render(
      <GstReconCollisionCard
        block={makeBlock()}
        onOverwrite={vi.fn()}
        onCreateNew={vi.fn()}
        onReject={vi.fn()}
      />,
    );

    expect(screen.getByText(/Missed vs GSTR-2B/)).toBeTruthy();
    expect(screen.getByTestId('gst-recon-collision-overwrite')).toBeTruthy();
    expect(screen.getByTestId('gst-recon-collision-create-new')).toBeTruthy();
    expect(screen.getByTestId('gst-recon-collision-reject')).toBeTruthy();
  });

  it('clicking Overwrite invokes onOverwrite directly — no text is produced or parsed', () => {
    const onOverwrite = vi.fn();
    const onCreateNew = vi.fn();
    const onReject = vi.fn();
    render(
      <GstReconCollisionCard
        block={makeBlock()}
        onOverwrite={onOverwrite}
        onCreateNew={onCreateNew}
        onReject={onReject}
      />,
    );

    fireEvent.click(screen.getByTestId('gst-recon-collision-overwrite'));

    expect(onOverwrite).toHaveBeenCalledTimes(1);
    expect(onOverwrite).toHaveBeenCalledWith();
    expect(onCreateNew).not.toHaveBeenCalled();
    expect(onReject).not.toHaveBeenCalled();
  });

  it('clicking Create new invokes onCreateNew directly — no text is produced or parsed', () => {
    const onOverwrite = vi.fn();
    const onCreateNew = vi.fn();
    const onReject = vi.fn();
    render(
      <GstReconCollisionCard
        block={makeBlock()}
        onOverwrite={onOverwrite}
        onCreateNew={onCreateNew}
        onReject={onReject}
      />,
    );

    fireEvent.click(screen.getByTestId('gst-recon-collision-create-new'));

    expect(onCreateNew).toHaveBeenCalledTimes(1);
    expect(onOverwrite).not.toHaveBeenCalled();
    expect(onReject).not.toHaveBeenCalled();
  });

  it('clicking Reject invokes onReject directly — discards without writing anything', () => {
    const onOverwrite = vi.fn();
    const onCreateNew = vi.fn();
    const onReject = vi.fn();
    render(
      <GstReconCollisionCard
        block={makeBlock()}
        onOverwrite={onOverwrite}
        onCreateNew={onCreateNew}
        onReject={onReject}
      />,
    );

    fireEvent.click(screen.getByTestId('gst-recon-collision-reject'));

    expect(onReject).toHaveBeenCalledTimes(1);
    expect(onOverwrite).not.toHaveBeenCalled();
    expect(onCreateNew).not.toHaveBeenCalled();
  });

  it('disables all three buttons while a click is resolving', () => {
    render(
      <GstReconCollisionCard
        block={makeBlock({ resolving: true })}
        onOverwrite={vi.fn()}
        onCreateNew={vi.fn()}
        onReject={vi.fn()}
      />,
    );

    expect(
      (screen.getByTestId('gst-recon-collision-overwrite') as HTMLButtonElement).disabled,
    ).toBe(true);
    expect(
      (screen.getByTestId('gst-recon-collision-create-new') as HTMLButtonElement).disabled,
    ).toBe(true);
    expect(
      (screen.getByTestId('gst-recon-collision-reject') as HTMLButtonElement).disabled,
    ).toBe(true);
  });
});
