// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useConversation } from '@/hooks/useConversation';
import type { WorkbookContext } from '@/types/cellix.types';
import type { ActionBlock } from '@/types/conversationTurn';

/**
 * TASKS.md #346 — a step that failed only because a sheet it writes to was
 * never created is finished in place, not by re-running the whole request.
 *
 * Live: step 4 of 4 wrote month formulas and a Main dashboard; no step had
 * created Main. The month writes landed, the Main writes threw ItemNotFound,
 * and the only advice was "re-run the request".
 */

const liveSheets = vi.hoisted(() => ({ names: ['Sheet1', 'January'] as string[] }));
vi.mock('@/utils/missingSheetRetry', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/utils/missingSheetRetry')>()),
  listWorksheetNames: vi.fn(async () => liveSheets.names),
}));

function sseBlock(type: string, data: unknown): string {
  return `event: ${type}\ndata: ${JSON.stringify(data)}\n\n`;
}

function makeSseResponse(blocks: string[]): Response {
  const encoder = new TextEncoder();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const block of blocks) controller.enqueue(encoder.encode(block));
      controller.close();
    },
  });
  return { ok: true, status: 200, body: stream, headers: new Headers() } as unknown as Response;
}

const WORKBOOK_CONTEXT = { activeSheet: 'Sheet1', sheets: ['Sheet1'] } as unknown as WorkbookContext;

const STEP_ACTIONS = [
  { type: 'SET_FORMULA', sheetName: 'January', range: 'F2', formula: '=E2-D2' },
  { type: 'BATCH_SET', sheetName: 'Main', operations: [] },
  { type: 'FORMAT_RANGE', sheetName: 'Main', range: 'A1:E1' },
];

const LAST_STEP = [
  sseBlock('actions', {
    actions: STEP_ACTIONS,
    explanation: 'Build the Main dashboard',
    changeSetId: 'cs_step_4',
    stepIndex: 4,
    stepTotal: 4,
    stepLabel: 'Build the Main dashboard',
    stepwise: true,
    runId: 'run_main',
  }),
  sseBlock('wave_ready', { runId: 'run_main', waveIndex: 3, waveTotal: 4, hasMore: true }),
];

const RUN_END = [
  sseBlock('conversation_end', { summary: 'All steps applied.' }),
  sseBlock('wave_ready', { runId: 'run_main', waveIndex: 3, waveTotal: 4, hasMore: false }),
];

const NOT_FOUND = "The requested resource doesn't exist.";

describe('useConversation — finishing a step that failed for want of a sheet', () => {
  beforeEach(() => {
    localStorage.clear();
    liveSheets.names = ['Sheet1', 'January'];
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  async function failStepOnMissingMain(
    workbookKey: string,
    applyError = `BATCH_SET: ${NOT_FOUND}; FORMAT_RANGE: ${NOT_FOUND}`,
  ) {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(makeSseResponse(LAST_STEP))
      .mockResolvedValueOnce(makeSseResponse(RUN_END));
    vi.stubGlobal('fetch', fetchMock);
    const onActions = vi
      .fn()
      .mockRejectedValueOnce(new Error(applyError))
      .mockResolvedValue(undefined);
    const onClearPreview = vi.fn();
    const { result } = renderHook(() =>
      useConversation({ workbookKey, onActions, onClearPreview }),
    );

    await act(async () => {
      await result.current.sendMessage('Create a table with sample headers and rows', [], WORKBOOK_CONTEXT, 'Sheet1 is empty.', {
        mode: 'action',
      });
    });
    await waitFor(() => expect(result.current.turns).toHaveLength(1));

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;
    await act(async () => {
      await result.current.acceptActions(turn.id, block.id).catch(() => undefined);
    });
    return { result, fetchMock, onActions, onClearPreview, turnId: turn.id };
  }

  it('names the missing sheet and offers to finish the step instead of saying "re-run"', async () => {
    const { result } = await failStepOnMissingMain('missing-sheet-offer');

    const turn = result.current.turns[0];
    expect(turn.missingSheetRetry?.sheets).toEqual(['Main']);
    expect(turn.error).toContain("'Main'");
    expect(turn.error).not.toMatch(/re-?run/i);
  });

  it('creates Main, re-applies only its actions, and carries the build on', async () => {
    const { result, fetchMock, onActions, onClearPreview, turnId } =
      await failStepOnMissingMain('missing-sheet-retry');

    await act(async () => {
      await result.current.retryMissingSheets(turnId);
    });

    // The January formula already landed on the failed attempt — not replayed.
    const retried = onActions.mock.calls[1][0] as Array<{ type: string; sheetName?: string }>;
    expect(retried.map((a) => `${a.type}:${a.sheetName}`)).toEqual([
      'ADD_SHEET:Main',
      'BATCH_SET:Main',
      'FORMAT_RANGE:Main',
    ]);
    // The pending preview would otherwise apply the card's full list.
    expect(onClearPreview).toHaveBeenCalled();

    const turn = result.current.turns[0];
    const block = turn.blocks.find((b): b is ActionBlock => b.type === 'actions')!;
    expect(block.proposalStatus).toBe('accepted');
    expect(turn.missingSheetRetry).toBeUndefined();

    // The run continues from this step — never re-planned from the start.
    expect(fetchMock).toHaveBeenCalledTimes(2);
    const [url, init] = fetchMock.mock.calls[1];
    expect(String(url)).toContain('/conversation/continue');
    expect(JSON.parse((init as RequestInit).body as string)).toMatchObject({
      runId: 'run_main',
      decision: 'accepted',
    });
  });

  it('still applies the failed part when the user created the sheet by hand meanwhile', async () => {
    const { result, onActions, turnId } = await failStepOnMissingMain('missing-sheet-manual');
    liveSheets.names = ['Sheet1', 'January', 'Main'];

    await act(async () => {
      await result.current.retryMissingSheets(turnId);
    });

    const retried = onActions.mock.calls[1][0] as Array<{ type: string; sheetName?: string }>;
    expect(retried.filter((a) => a.type !== 'ADD_SHEET').map((a) => a.type)).toEqual([
      'BATCH_SET',
      'FORMAT_RANGE',
    ]);
    expect(retried.some((a) => a.sheetName === 'January')).toBe(false);
  });

  it('makes no offer when something other than a missing sheet also failed', async () => {
    const { result } = await failStepOnMissingMain(
      'missing-sheet-mixed',
      `BATCH_SET: ${NOT_FOUND}; SET_FORMULA: Invalid argument`,
    );

    const turn = result.current.turns[0];
    expect(turn.missingSheetRetry).toBeUndefined();
    expect(turn.error).toBeTruthy();
  });
});
