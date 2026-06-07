import { useCallback, useRef, useState } from 'react';
import { getConversationEndpoint } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import {
  prepareConversationRequestPayload,
  ConversationHistoryMessage,
  computeSheetLayout,
  SheetLayoutPayload,
  WorkbookContextPayload,
} from '@/utils/payloadCompressor';
import {
  sanitizeActions,
  CLARIFY_ROW_PLACEMENT,
} from '@/utils/actionGuard';
import { parseSseEventBlock } from '@/utils/sseParser';
import { TIMING, createGate, delay, waitWithMin } from '@/utils/revealQueue';
import { buildThoughtSummary } from '@/utils/thoughtSummary';
import { buildClientStatusMessage, isSimpleCreateTask } from '@/utils/statusMessage';
import {
  ActionBlock,
  AnswerBlock,
  ConversationTurn,
  StepPhase,
  ThinkingBlock,
  TurnBlock,
  truncateTabLabel,
} from '@/types/conversationTurn';

interface UseConversationReturn {
  turns: ConversationTurn[];
  activeTurnId: string | null;
  isWaitingForResponse: boolean;
  conversationId: string | null;
  sendMessage: (message: string, sheetData: unknown[][], workbookContext?: WorkbookContextPayload) => Promise<void>;
  answerQuestion: (answer: string, sheetData: unknown[][], workbookContext?: WorkbookContextPayload) => Promise<void>;
  acceptActions: (turnId: string, blockId: string) => Promise<void>;
  rejectActions: (turnId: string, blockId: string) => void;
  endConversation: () => void;
  clearConversation: () => void;
  toggleThinking: (turnId: string, blockId: string) => void;
  markAnswerComplete: (turnId: string, blockId: string) => void;
}

interface UseConversationOptions {
  onActions?: (actions: SheetAction[], explanation: string) => void | Promise<void>;
  onPreviewActions?: (actions: SheetAction[]) => void | Promise<void>;
  onClearPreview?: () => void | Promise<void>;
  autoApplyActions?: boolean;
}

interface PendingResponse {
  type: 'answer' | 'question';
  answer?: string;
  question?: string;
  options?: string[];
}

interface PendingActions {
  id: string;
  actions: SheetAction[];
  explanation: string;
}

interface TurnRuntime {
  analyzingGate: ReturnType<typeof createGate>;
  responseGate: ReturnType<typeof createGate>;
  pendingResponse: PendingResponse | null;
  pendingActions: PendingActions | null;
  aborted: boolean;
}

const THINKING_ID = 'thinking_main';
const STATUS_ID = 'status_active';
const STEP_READING_ID = 'step_reading';
const STEP_ANALYZING_ID = 'step_analyzing';
const ANSWER_ID = 'answer_main';

const READING_LABEL = 'Reading your worksheet…';
const ANALYZING_LABEL = 'Analyzing your spreadsheet…';

interface TimelineOptions {
  sheetIsEmpty: boolean;
  userMessage: string;
}

function finalizeSteps(blocks: TurnBlock[], userMessage: string): TurnBlock[] {
  const summary = buildThoughtSummary(userMessage, 'final');

  return blocks
    .filter((b) => b.type !== 'step' && b.type !== 'status')
    .map((block) => {
      if (block.type === 'thinking') {
        return {
          ...block,
          content: block.content.trim() ? block.content : summary,
          loading: false,
          expanded: false,
          visible: true,
        };
      }
      return block;
    });
}

function withoutStatus(blocks: TurnBlock[]): TurnBlock[] {
  return blocks.filter((b) => b.type !== 'status');
}

function upsertStep(
  blocks: TurnBlock[],
  id: string,
  label: string,
  phase: StepPhase,
): TurnBlock[] {
  const exists = blocks.some((b) => b.id === id);
  if (!exists) {
    return [...blocks, { id, type: 'step', label, phase }];
  }
  return blocks.map((b) =>
    b.id === id && b.type === 'step' ? { ...b, label, phase } : b,
  );
}

function upsertThinking(
  blocks: TurnBlock[],
  content: string,
  opts: { loading?: boolean; visible?: boolean; expanded?: boolean } = {},
): TurnBlock[] {
  const rest = blocks.filter((b) => b.type !== 'thinking');
  const existing = blocks.find((b): b is ThinkingBlock => b.type === 'thinking');

  const expanded =
    opts.expanded !== undefined
      ? opts.expanded
      : (existing?.type === 'thinking' ? existing.expanded : true);
  const loading =
    opts.loading !== undefined
      ? opts.loading
      : (existing?.type === 'thinking' ? existing.loading : true);
  const visible =
    opts.visible !== undefined
      ? opts.visible
      : (existing?.type === 'thinking' ? existing.visible : true);

  return [
    ...rest,
    {
      id: THINKING_ID,
      type: 'thinking',
      content,
      expanded,
      loading,
      visible,
    },
  ];
}

function upsertStatus(
  blocks: TurnBlock[],
  label: string,
  pulsing: boolean,
  visible = true,
): TurnBlock[] {
  const rest = blocks.filter((b) => b.type !== 'status');
  return [...rest, { id: STATUS_ID, type: 'status', label, pulsing, visible }];
}

function createRuntime(): TurnRuntime {
  return {
    analyzingGate: createGate(),
    responseGate: createGate(),
    pendingResponse: null,
    pendingActions: null,
    aborted: false,
  };
}

function createActionBlock(pending: PendingActions): ActionBlock {
  return {
    id: pending.id,
    type: 'actions',
    actions: pending.actions,
    explanation: pending.explanation,
    proposalStatus: 'pending',
  };
}

export const useConversation = (options: UseConversationOptions = {}): UseConversationReturn => {
  const { onActions, onPreviewActions, onClearPreview, autoApplyActions = false } = options;

  const [turns, setTurns] = useState<ConversationTurn[]>([]);
  const [activeTurnId, setActiveTurnId] = useState<string | null>(null);
  const [isWaitingForResponse, setIsWaitingForResponse] = useState(false);
  const [conversationId, setConversationId] = useState<string | null>(null);

  const abortControllerRef = useRef<AbortController | null>(null);
  const conversationIdRef = useRef<string | null>(null);
  const turnsRef = useRef<ConversationTurn[]>([]);
  const historyRef = useRef<ConversationHistoryMessage[]>([]);
  const runtimeRef = useRef<Map<string, TurnRuntime>>(new Map());
  const revealScheduledRef = useRef<Set<string>>(new Set());
  const sheetLayoutRef = useRef<SheetLayoutPayload | null>(null);

  const syncTurns = useCallback((next: ConversationTurn[]) => {
    turnsRef.current = next;
    setTurns(next);
  }, []);

  const updateTurn = useCallback(
    (turnId: string, updater: (turn: ConversationTurn) => ConversationTurn) => {
      syncTurns(
        turnsRef.current.map((turn) => (turn.id === turnId ? updater(turn) : turn)),
      );
    },
    [syncTurns],
  );

  const syncConversationId = useCallback((id?: string) => {
    if (!id) return;
    conversationIdRef.current = id;
    setConversationId(id);
  }, []);

  const pushHistory = useCallback((entry: ConversationHistoryMessage) => {
    historyRef.current = [...historyRef.current, entry];
  }, []);

  const getUserFacingErrorMessage = useCallback(async (response: Response): Promise<string> => {
    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      try {
        const json = await response.json();
        if (json?.error?.message) return String(json.error.message);
        if (json?.message) return String(json.message);
      } catch {
        // fall through
      }
    }
    try {
      const text = await response.text();
      if (text.trim()) return text.trim();
    } catch {
      // ignore
    }
    return `Request failed (HTTP ${response.status})`;
  }, []);

  const revealFinalResponse = useCallback(
    (turnId: string, response: PendingResponse) => {
      if (revealScheduledRef.current.has(turnId)) return;
      revealScheduledRef.current.add(turnId);
      const pendingActions = runtimeRef.current.get(turnId)?.pendingActions;

      updateTurn(turnId, (turn) => ({
        ...turn,
        phase: response.type === 'question' ? 'awaiting_input' : 'complete',
        blocks: finalizeSteps(
          upsertThinking(turn.blocks, buildThoughtSummary(turn.userMessage, 'final'), {
            loading: false,
            expanded: false,
            visible: true,
          }),
          turn.userMessage,
        ).concat(
          response.type === 'answer'
            ? [
                {
                  id: ANSWER_ID,
                  type: 'answer',
                  content: response.answer ?? '',
                  revealState: 'typing',
                } satisfies AnswerBlock,
              ]
            : [
                {
                  id: `question_${Date.now()}`,
                  type: 'question',
                  question: response.question ?? '',
                  options: response.options,
                  revealState: 'visible',
                },
              ],
        ).concat(pendingActions ? [createActionBlock(pendingActions)] : []),
      }));

      setIsWaitingForResponse(false);
    },
    [updateTurn],
  );

  const runVisualTimeline = useCallback(
    async (turnId: string, runtime: TurnRuntime, opts: TimelineOptions) => {
      const isAborted = () => runtime.aborted;
      const simple = isSimpleCreateTask(opts.userMessage, opts.sheetIsEmpty);
      const statusLabel = buildClientStatusMessage(opts.userMessage, opts.sheetIsEmpty);

      if (simple) {
        await delay(150);
        if (isAborted()) return;
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStep(turn.blocks, STEP_READING_ID, statusLabel, 'running'),
        }));
        runtime.analyzingGate.open();
        await waitWithMin(runtime.responseGate, 300);
        if (isAborted()) return;
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStep(turn.blocks, STEP_READING_ID, statusLabel, 'done'),
        }));
        return;
      }

      await delay(TIMING.readingReveal);
      if (isAborted()) return;

      const readingLabel = opts.sheetIsEmpty
        ? buildClientStatusMessage(opts.userMessage, true)
        : READING_LABEL;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(turn.blocks, STEP_READING_ID, readingLabel, 'revealed'),
      }));

      await delay(TIMING.readingSpinner);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(turn.blocks, STEP_READING_ID, readingLabel, 'running'),
      }));

      await delay(TIMING.pauseAfterReadingStepBeforeThinking);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertThinking(turn.blocks, buildThoughtSummary(turn.userMessage, 'reading'), {
          loading: true,
          visible: true,
          expanded: true,
        }),
      }));

      await delay(TIMING.statusReveal);
      if (isAborted()) return;

      if (!opts.sheetIsEmpty) {
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStatus(turn.blocks, statusLabel, true, true),
        }));
      }

      await waitWithMin(runtime.analyzingGate, TIMING.readingMinRun);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => {
        const prev = turn.blocks.find((b): b is ThinkingBlock => b.type === 'thinking');
        return {
          ...turn,
          blocks: withoutStatus(
            upsertThinking(turn.blocks, prev?.content ?? buildThoughtSummary(turn.userMessage, 'reading'), {
              visible: false,
              loading: false,
              expanded: false,
            }),
          ),
        };
      });

      if (opts.sheetIsEmpty) {
        runtime.analyzingGate.open();
        await waitWithMin(runtime.responseGate, TIMING.readingMinRun);
        return;
      }

      await delay(TIMING.gapBeforeAnalyzingStepRow);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(
          upsertStep(turn.blocks, STEP_READING_ID, readingLabel, 'done'),
          STEP_ANALYZING_ID,
          ANALYZING_LABEL,
          'revealed',
        ),
      }));

      await delay(TIMING.analyzingSpinner);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(turn.blocks, STEP_ANALYZING_ID, ANALYZING_LABEL, 'running'),
      }));

      await delay(TIMING.pauseAfterAnalyzingStepBeforeThinking);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertThinking(
          turn.blocks,
          buildThoughtSummary(turn.userMessage, 'analyzing'),
          { loading: true, visible: true, expanded: true },
        ),
      }));

      await waitWithMin(runtime.responseGate, TIMING.analyzingMinRun);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(turn.blocks, STEP_ANALYZING_ID, ANALYZING_LABEL, 'done'),
      }));

      await delay(TIMING.pauseBeforeComposing);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertThinking(
          upsertStatus(turn.blocks, 'Composing response…', true, true),
          buildThoughtSummary(turn.userMessage, 'composing'),
          { loading: true, visible: true, expanded: true },
        ),
      }));

      await delay(
        runtime.pendingResponse?.type === 'question' ? TIMING.questionReveal : TIMING.answerReveal,
      );
      if (isAborted()) return;

      if (!runtime.pendingResponse) {
        await runtime.responseGate.wait();
      }

      const response = runtime.pendingResponse;
      if (response && !revealScheduledRef.current.has(turnId)) {
        revealFinalResponse(turnId, response);
      }
    },
    [revealFinalResponse, updateTurn],
  );

  const signalResponse = useCallback(
    (turnId: string, response: PendingResponse) => {
      const runtime = runtimeRef.current.get(turnId);
      if (!runtime) return;
      runtime.pendingResponse = response;
      runtime.analyzingGate.open();
      runtime.responseGate.open();
    },
    [],
  );

  const processStream = useCallback(
    async (response: Response, turnId: string) => {
      const reader = response.body?.getReader();
      if (!reader) throw new Error('Response body is not readable');

      const decoder = new TextDecoder();
      let buffer = '';
      let streamBuffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        const parts = buffer.split(/\r?\n\r?\n/);
        buffer = parts.pop() || '';

        for (const part of parts) {
          const event = parseSseEventBlock(part);
          if (!event) continue;

          if (event.data && 'conversationId' in event.data && event.data.conversationId) {
            syncConversationId(event.data.conversationId);
          }

          if (event.type === 'status' && /analyz/i.test(event.data.message)) {
            runtimeRef.current.get(turnId)?.analyzingGate.open();
            updateTurn(turnId, (turn) => ({
              ...turn,
              blocks: upsertThinking(turn.blocks, buildThoughtSummary(turn.userMessage, 'analyzing'), {
                loading: true,
              }),
            }));
            continue;
          }

          if (event.type === 'chunk') {
            streamBuffer += event.data.text;
            runtimeRef.current.get(turnId)?.analyzingGate.open();
            continue;
          }

          if (event.type === 'answer') {
            pushHistory({
              role: 'assistant',
              content: event.data.answer,
              timestamp: new Date().toISOString(),
              type: 'answer',
            });
            signalResponse(turnId, { type: 'answer', answer: event.data.answer });
            continue;
          }

          if (event.type === 'question') {
            pushHistory({
              role: 'assistant',
              content: event.data.question,
              timestamp: new Date().toISOString(),
              type: 'question',
            });
            signalResponse(turnId, {
              type: 'question',
              question: event.data.question,
              options: event.data.options,
            });
            continue;
          }

          if (event.type === 'actions') {
            const sanitized = sanitizeActions(
              event.data.actions,
              sheetLayoutRef.current ?? undefined,
            );

            if (sanitized.requiresClarification && sanitized.actions.length === 0) {
              pushHistory({
                role: 'assistant',
                content: CLARIFY_ROW_PLACEMENT.question,
                timestamp: new Date().toISOString(),
                type: 'question',
              });
              signalResponse(turnId, {
                type: 'question',
                question: CLARIFY_ROW_PLACEMENT.question,
                options: [...CLARIFY_ROW_PLACEMENT.options],
              });
              continue;
            }

            const explanation =
              sanitized.warnings.length > 0
                ? `${event.data.explanation} (${sanitized.warnings.join(' ')})`
                : event.data.explanation;

            const pendingActions = {
              id: `actions_${Date.now()}`,
              actions: sanitized.actions,
              explanation,
            };
            const runtime = runtimeRef.current.get(turnId);
            if (runtime) {
              runtime.pendingActions = pendingActions;
            }

            if (!autoApplyActions) {
              await onPreviewActions?.(sanitized.actions);
            }

            if (revealScheduledRef.current.has(turnId)) {
              updateTurn(turnId, (turn) => {
                const withoutOldPending = turn.blocks.filter(
                  (b) => !(b.type === 'actions' && b.proposalStatus === 'pending'),
                );
                return {
                  ...turn,
                  blocks: [...withoutOldPending, createActionBlock(pendingActions)],
                };
              });
            }

            if (autoApplyActions && onActions) {
              await onClearPreview?.();
              await onActions(sanitized.actions, explanation);
              updateTurn(turnId, (turn) => ({
                ...turn,
                blocks: turn.blocks.map((b) =>
                  b.id === pendingActions.id && b.type === 'actions'
                    ? { ...b, proposalStatus: 'accepted' }
                    : b,
                ),
              }));
            }
            continue;
          }

          if (event.type === 'error') {
            updateTurn(turnId, (turn) => ({
              ...turn,
              phase: 'error',
              error: event.data.message,
              blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
            }));
            runtimeRef.current.get(turnId)!.aborted = true;
            setIsWaitingForResponse(false);
            continue;
          }

          if (event.type === 'conversation_end' || event.type === 'done') {
            if (streamBuffer && !runtimeRef.current.get(turnId)?.pendingResponse) {
              signalResponse(turnId, { type: 'answer', answer: streamBuffer });
            }
          }
        }
      }

      if (streamBuffer && !runtimeRef.current.get(turnId)?.pendingResponse) {
        signalResponse(turnId, { type: 'answer', answer: streamBuffer });
      }
    },
    [
      autoApplyActions,
      onActions,
      onClearPreview,
      onPreviewActions,
      pushHistory,
      signalResponse,
      syncConversationId,
      updateTurn,
    ],
  );

  const sendMessage = useCallback(
    async (message: string, sheetData: unknown[][], workbookContext?: WorkbookContextPayload) => {
      const trimmed = message.trim();
      if (!trimmed) return;

      const turnId = `turn_${Date.now()}`;
      const timestamp = new Date();
      const runtime = createRuntime();
      runtimeRef.current.set(turnId, runtime);
      revealScheduledRef.current.delete(turnId);

      const newTurn: ConversationTurn = {
        id: turnId,
        userMessage: trimmed,
        timestamp,
        tabLabel: truncateTabLabel(trimmed),
        phase: 'processing',
        blocks: [],
      };

      pushHistory({
        role: 'user',
        content: trimmed,
        timestamp: timestamp.toISOString(),
        type: 'command',
      });

      syncTurns([...turnsRef.current, newTurn]);
      setActiveTurnId(turnId);
      setIsWaitingForResponse(true);

      abortControllerRef.current?.abort();
      abortControllerRef.current = new AbortController();

      const sheetLayout = computeSheetLayout(sheetData);
      sheetLayoutRef.current = sheetLayout;

      const timelinePromise = runVisualTimeline(turnId, runtime, {
        sheetIsEmpty: sheetLayout.isEmpty,
        userMessage: trimmed,
      });

      const requestPayload = prepareConversationRequestPayload(trimmed, sheetData, {
        conversationId: conversationIdRef.current,
        previousMessages: historyRef.current.slice(0, -1),
        workbookContext,
      });

      try {
        const endpoint = getConversationEndpoint();
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (endpoint.includes('.ngrok-free.app')) {
          headers['ngrok-skip-browser-warning'] = 'true';
        }

        const response = await fetch(endpoint, {
          method: 'POST',
          headers,
          body: JSON.stringify(requestPayload),
          signal: abortControllerRef.current.signal,
        });

        if (!response.ok) {
          const errorMessage = await getUserFacingErrorMessage(response);
          throw new Error(errorMessage);
        }

        await Promise.all([processStream(response, turnId), timelinePromise]);
      } catch (error: unknown) {
        runtime.aborted = true;

        if (error instanceof Error && error.name === 'AbortError') {
          updateTurn(turnId, (turn) => ({
            ...turn,
            phase: 'complete',
            blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
          }));
          return;
        }

        const endpoint = getConversationEndpoint();
        const messageText =
          error instanceof Error && error.message === 'Failed to fetch'
            ? `Failed to reach ${endpoint}. Ensure the backend is running and the Vite /api proxy is configured.`
            : error instanceof Error
              ? error.message
              : 'Stream connection failed';

        updateTurn(turnId, (turn) => ({
          ...turn,
          phase: 'error',
          error: messageText,
          blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
        }));
      } finally {
        if (!revealScheduledRef.current.has(turnId)) {
          setIsWaitingForResponse(false);
        }
        abortControllerRef.current = null;
        runtimeRef.current.delete(turnId);
      }
    },
    [
      getUserFacingErrorMessage,
      processStream,
      pushHistory,
      runVisualTimeline,
      syncTurns,
      updateTurn,
    ],
  );

  const answerQuestion = useCallback(
    async (answer: string, sheetData: unknown[][], workbookContext?: WorkbookContextPayload) => {
      await sendMessage(answer, sheetData, workbookContext);
    },
    [sendMessage],
  );

  const acceptActions = useCallback(
    async (turnId: string, blockId: string) => {
      const turn = turnsRef.current.find((t) => t.id === turnId);
      const block = turn?.blocks.find(
        (b): b is ActionBlock => b.id === blockId && b.type === 'actions',
      );
      if (!block || block.proposalStatus !== 'pending') return;

      if (onActions) {
        await onActions(block.actions, block.explanation);
      }

      updateTurn(turnId, (t) => ({
        ...t,
        blocks: t.blocks.map((b) =>
          b.id === blockId && b.type === 'actions'
            ? { ...b, proposalStatus: 'accepted' }
            : b,
        ),
      }));
    },
    [onActions, updateTurn],
  );

  const rejectActions = useCallback(
    async (turnId: string, blockId: string) => {
      await onClearPreview?.();
      updateTurn(turnId, (t) => ({
        ...t,
        blocks: t.blocks.map((b) =>
          b.id === blockId && b.type === 'actions'
            ? { ...b, proposalStatus: 'rejected' }
            : b,
        ),
      }));
    },
    [onClearPreview, updateTurn],
  );

  const toggleThinking = useCallback(
    (turnId: string, blockId: string) => {
      updateTurn(turnId, (t) => ({
        ...t,
        blocks: t.blocks.map((b) =>
          b.id === blockId && b.type === 'thinking'
            ? { ...b, expanded: !b.expanded }
            : b,
        ),
      }));
    },
    [updateTurn],
  );

  const markAnswerComplete = useCallback(
    (turnId: string, blockId: string) => {
      updateTurn(turnId, (t) => ({
        ...t,
        blocks: t.blocks.map((b) =>
          b.id === blockId && b.type === 'answer'
            ? { ...b, revealState: 'complete' }
            : b,
        ),
      }));
    },
    [updateTurn],
  );

  const endConversation = useCallback(() => {
    abortControllerRef.current?.abort();
    void onClearPreview?.();
    if (activeTurnId) {
      const runtime = runtimeRef.current.get(activeTurnId);
      if (runtime) runtime.aborted = true;
      updateTurn(activeTurnId, (turn) => ({
        ...turn,
        phase: 'complete',
        blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
      }));
    }
    setIsWaitingForResponse(false);
  }, [activeTurnId, onClearPreview, updateTurn]);

  const clearConversation = useCallback(() => {
    abortControllerRef.current?.abort();
    void onClearPreview?.();
    runtimeRef.current.forEach((r) => {
      r.aborted = true;
    });
    runtimeRef.current.clear();
    conversationIdRef.current = null;
    historyRef.current = [];
    turnsRef.current = [];
    revealScheduledRef.current.clear();
    setTurns([]);
    setActiveTurnId(null);
    setConversationId(null);
    setIsWaitingForResponse(false);
  }, [onClearPreview]);

  return {
    turns,
    activeTurnId,
    isWaitingForResponse,
    conversationId,
    sendMessage,
    answerQuestion,
    acceptActions,
    rejectActions,
    endConversation,
    clearConversation,
    toggleThinking,
    markAnswerComplete,
  };
};
