import { useCallback, useEffect, useRef, useState } from 'react';
import { buildWorkbookContext } from '@/services/sheetContextBuilder';
import {
  getContextForSend,
  markPendingWorkbookContextStale,
  setPendingWorkbookContext,
} from '@/utils/pendingWorkbookContext';
import { getConversationEndpoint } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import {
  prepareConversationRequestPayload,
  ConversationHistoryMessage,
  computeSheetLayout,
  SheetLayoutPayload,
} from '@/utils/payloadCompressor';
import { WorkbookContext } from '@/types/cellix.types';
import {
  sanitizeActions,
  CLARIFY_ROW_PLACEMENT,
} from '@/utils/actionGuard';
import { parseSseEventBlock } from '@/utils/sseParser';
import { handleToolRequest } from '@/services/toolRequestHandler';
import { navigateToCell } from '@/services/rangeFetchService';
import { shouldAcceptIncomingClarification } from '@/utils/clarification.util';
import { TIMING, createGate, delay, waitWithMin } from '@/utils/revealQueue';
import { buildThoughtSummary } from '@/utils/thoughtSummary';
import { buildClientStatusMessage, isSimpleCreateTask } from '@/utils/statusMessage';
import { tryLocalSheetActions, LocalSheetActionPlan } from '@/utils/localSheetActions';
import { shouldPreviewActions } from '@/utils/previewPolicy';
import { ClarificationPayload } from '@/types/cellix.types';
import { CellChange } from '@/types/changeSet';
import { AssistantMode, DEFAULT_ASSISTANT_MODE } from '@/types/mode';
import {
  ActionBlock,
  AnswerBlock,
  ConversationTurn,
  MatchResult,
  PlanBlock,
  StepPhase,
  ThinkingBlock,
  TurnBlock,
  truncateTabLabel,
} from '@/types/conversationTurn';
import { ChatSession, createChatSession } from '@/types/chatSession';
import {
  loadChatSessions,
  saveChatSessions,
} from '@/utils/chatSessionStorage';
import {
  mergeSessionFromStored,
  StoredConversation,
} from '@/utils/rehydrateConversation';
import { getConversationByIdEndpoint } from '@/lib/apiConfig';

/* global Excel, Office */

interface UseConversationReturn {
  sessions: ChatSession[];
  activeSessionId: string | null;
  turns: ConversationTurn[];
  activeTurnId: string | null;
  isWaitingForResponse: boolean;
  isWaitingClarification: boolean;
  activeClarification: ClarificationPayload | null;
  conversationId: string | null;
  sendMessage: (
    message: string,
    sheetData: unknown[][],
    workbookContext?: import('@/types/cellix.types').WorkbookContext,
    promptContext?: string,
    options?: SendMessageOptions,
  ) => Promise<void>;
  answerQuestion: (
    answer: string,
    sheetData: unknown[][],
    workbookContext?: import('@/types/cellix.types').WorkbookContext,
    promptContext?: string,
    options?: SendMessageOptions,
  ) => Promise<void>;
  answerClarification: (
    answer: string,
    sheetData: unknown[][],
    workbookContext?: import('@/types/cellix.types').WorkbookContext,
    promptContext?: string,
    options?: SendMessageOptions,
  ) => Promise<void>;
  dismissClarification: () => void;
  acceptActions: (turnId: string, blockId: string) => Promise<void>;
  rejectActions: (turnId: string, blockId: string) => void;
  endConversation: () => void;
  newChat: () => void;
  clearConversation: () => void;
  selectSession: (sessionId: string) => void;
  closeSession: (sessionId: string) => void;
  selectTurn: (turnId: string) => void;
  closeTurn: (turnId: string) => void;
  toggleThinking: (turnId: string, blockId: string) => void;
  markAnswerComplete: (turnId: string, blockId: string) => void;
}

interface UseConversationOptions {
  workbookKey?: string;
  onActions?: (
    actions: SheetAction[],
    explanation: string,
    meta?: PreviewActionsMeta,
  ) => void | Promise<void>;
  onPreviewActions?: (
    actions: SheetAction[],
    explanation: string,
    meta?: PreviewActionsMeta,
  ) => void | Promise<void>;
  onClearPreview?: () => void | Promise<void>;
  onChangeSetApplied?: (changeSetId: string) => void;
  autoApplyActions?: boolean;
  previewEnabled?: boolean;
  isChangeSetApplied?: (changeSetId?: string) => boolean;
}

export interface SendMessageOptions {
  refinementChangeSetId?: string;
  mode?: AssistantMode;
}

interface PendingResponse {
  type: 'answer' | 'question' | 'clarification';
  answer?: string;
  matches?: MatchResult[];
  question?: string;
  options?: string[];
}

interface PendingActions {
  id: string;
  actions: SheetAction[];
  explanation: string;
  changeSetId?: string;
  changes?: CellChange[];
}

export interface PreviewActionsMeta {
  changeSetId?: string;
  changes?: CellChange[];
}

interface TurnRuntime {
  analyzingGate: ReturnType<typeof createGate>;
  responseGate: ReturnType<typeof createGate>;
  pendingResponse: PendingResponse | null;
  pendingActions: PendingActions | null;
  pendingPlan: PlanBlock | null;
  aborted: boolean;
  mode: AssistantMode;
}

const THINKING_ID = 'thinking_main';
const STATUS_ID = 'status_active';
const STEP_READING_ID = 'step_reading';
const STEP_ANALYZING_ID = 'step_analyzing';
const ANSWER_ID_PREFIX = 'answer_';

function answerBlockId(turnId: string): string {
  return `${ANSWER_ID_PREFIX}${turnId}`;
}

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

function createRuntime(mode: AssistantMode = DEFAULT_ASSISTANT_MODE): TurnRuntime {
  return {
    analyzingGate: createGate(),
    responseGate: createGate(),
    pendingResponse: null,
    pendingActions: null,
    pendingPlan: null,
    aborted: false,
    mode,
  };
}

function createActionBlock(
  pending: PendingActions,
  isChangeSetApplied?: (changeSetId?: string) => boolean,
): ActionBlock {
  const alreadyApplied = isChangeSetApplied?.(pending.changeSetId) ?? false;
  return {
    id: pending.id,
    type: 'actions',
    actions: pending.actions,
    explanation: pending.explanation,
    proposalStatus: alreadyApplied ? 'accepted' : 'pending',
    changeSetId: pending.changeSetId,
    changes: pending.changes,
  };
}

export const useConversation = (options: UseConversationOptions = {}): UseConversationReturn => {
  const {
    workbookKey = 'workbook',
    onActions,
    onPreviewActions,
    onClearPreview,
    onChangeSetApplied,
    autoApplyActions = false,
    previewEnabled = true,
    isChangeSetApplied,
  } = options;

  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<string | null>(null);
  const [activeTurnId, setActiveTurnId] = useState<string | null>(null);
  const [isWaitingForResponse, setIsWaitingForResponse] = useState(false);
  const [activeClarification, setActiveClarification] = useState<ClarificationPayload | null>(null);
  const [isOfficeReady, setIsOfficeReady] = useState(false);

  const preCompressionInProgressRef = useRef(false);

  useEffect(() => {
    if (typeof Office === 'undefined') return;
    Office.onReady(() => setIsOfficeReady(true));
  }, []);

  useEffect(() => {
    if (!isOfficeReady || typeof Excel === 'undefined') return;

    let selectionHandler: { remove: () => void } | null = null;
    const changedHandlers: Array<{ remove: () => void }> = [];

    Excel.run(async (ctx) => {
      const selectionResult = ctx.workbook.onSelectionChanged.add(async () => {
        if (preCompressionInProgressRef.current) return;
        preCompressionInProgressRef.current = true;
        try {
          // Selection change invalidates then warm-rebuilds the pending snapshot.
          markPendingWorkbookContextStale();
          const { context, promptContext, activeSheetData } = await buildWorkbookContext([]);
          const toonString = promptContext ?? '';
          setPendingWorkbookContext({
            toon: toonString,
            workbookContext: context,
            activeSheetData,
          });
        } catch {
          markPendingWorkbookContextStale();
        } finally {
          preCompressionInProgressRef.current = false;
        }
      });
      selectionHandler = selectionResult as unknown as { remove: () => void };

      ctx.workbook.worksheets.load('items');
      await ctx.sync();
      for (const sheet of ctx.workbook.worksheets.items) {
        const changed = sheet.onChanged.add(async () => {
          // External or in-sheet edits invalidate the prebuild (rebuild on next send).
          markPendingWorkbookContextStale();
        });
        changedHandlers.push(changed as unknown as { remove: () => void });
      }
      await ctx.sync();
    }).catch(() => {
      // Office.js not available in test env — ignore
    });

    return () => {
      void selectionHandler?.remove?.();
      for (const handler of changedHandlers) {
        void handler.remove?.();
      }
    };
  }, [isOfficeReady]);

  useEffect(() => {
    activeClarificationRef.current = activeClarification;
  }, [activeClarification]);
  const [conversationId, setConversationId] = useState<string | null>(null);

  const abortControllerRef = useRef<AbortController | null>(null);
  const conversationIdRef = useRef<string | null>(null);
  const sessionsRef = useRef<ChatSession[]>([]);
  const activeSessionIdRef = useRef<string | null>(null);
  const historyRef = useRef<ConversationHistoryMessage[]>([]);
  const runtimeRef = useRef<Map<string, TurnRuntime>>(new Map());
  const revealScheduledRef = useRef<Set<string>>(new Set());
  const sheetLayoutRef = useRef<SheetLayoutPayload | null>(null);
  const activeClarificationRef = useRef<ClarificationPayload | null>(null);
  const persistTimerRef = useRef<number | null>(null);
  const hydratedRef = useRef(false);

  const activeSession = sessions.find((s) => s.id === activeSessionId) ?? null;
  const turns = activeSession?.turns ?? [];
  const awaitingInput = turns.some((turn) => turn.phase === 'awaiting_input');

  const applySessionContext = useCallback((session: ChatSession | null) => {
    conversationIdRef.current = session?.conversationId ?? null;
    historyRef.current = session?.history ?? [];
    setConversationId(session?.conversationId ?? null);
  }, []);

  const schedulePersist = useCallback(() => {
    if (!workbookKey) return;
    if (persistTimerRef.current) {
      window.clearTimeout(persistTimerRef.current);
    }
    persistTimerRef.current = window.setTimeout(() => {
      const existing = loadChatSessions(workbookKey);
      saveChatSessions(workbookKey, {
        activeSessionId: activeSessionIdRef.current,
        sessions: sessionsRef.current,
        assistantMode: existing?.assistantMode,
      });
    }, 400);
  }, [workbookKey]);

  const syncSessions = useCallback(
    (next: ChatSession[]) => {
      sessionsRef.current = next;
      setSessions(next);
      schedulePersist();
    },
    [schedulePersist],
  );

  const updateSession = useCallback(
    (sessionId: string, updater: (session: ChatSession) => ChatSession) => {
      syncSessions(
        sessionsRef.current.map((session) =>
          session.id === sessionId ? updater(session) : session,
        ),
      );
    },
    [syncSessions],
  );

  const getActiveSession = useCallback((): ChatSession | null => {
    const id = activeSessionIdRef.current;
    if (!id) return null;
    return sessionsRef.current.find((session) => session.id === id) ?? null;
  }, []);

  const updateTurn = useCallback(
    (turnId: string, updater: (turn: ConversationTurn) => ConversationTurn) => {
      const sessionId = activeSessionIdRef.current;
      if (!sessionId) return;
      updateSession(sessionId, (session) => ({
        ...session,
        updatedAt: new Date().toISOString(),
        turns: session.turns.map((turn) => (turn.id === turnId ? updater(turn) : turn)),
      }));
    },
    [updateSession],
  );

  const syncConversationId = useCallback(
    (id?: string) => {
      if (!id) return;
      const sessionId = activeSessionIdRef.current;
      if (!sessionId) return;
      conversationIdRef.current = id;
      setConversationId(id);
      updateSession(sessionId, (session) => ({
        ...session,
        conversationId: id,
        updatedAt: new Date().toISOString(),
      }));
    },
    [updateSession],
  );

  const pushHistory = useCallback(
    (entry: ConversationHistoryMessage) => {
      const sessionId = activeSessionIdRef.current;
      if (!sessionId) return;
      const nextHistory = [...historyRef.current, entry];
      historyRef.current = nextHistory;
      updateSession(sessionId, (session) => ({
        ...session,
        history: nextHistory,
        updatedAt: new Date().toISOString(),
      }));
    },
    [updateSession],
  );

  const ensureActiveSession = useCallback((): ChatSession => {
    let session = getActiveSession();
    if (session) return session;

    const created = createChatSession();
    activeSessionIdRef.current = created.id;
    setActiveSessionId(created.id);
    applySessionContext(created);
    syncSessions([...sessionsRef.current, created]);
    return created;
  }, [applySessionContext, getActiveSession, syncSessions]);

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
      const runtimeForReveal = runtimeRef.current.get(turnId);
      const pendingActions = runtimeForReveal?.pendingActions;
      const pendingPlan = runtimeForReveal?.pendingPlan;

      updateTurn(turnId, (turn) => ({
        ...turn,
        phase:
          response.type === 'question' || response.type === 'clarification'
            ? 'awaiting_input'
            : 'complete',
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
                  id: answerBlockId(turnId),
                  type: 'answer',
                  content: response.answer ?? '',
                  revealState: 'typing',
                  matches: response.matches,
                } satisfies AnswerBlock,
              ]
            : response.type === 'question' || response.type === 'clarification'
              ? [
                  {
                    id: `question_${Date.now()}`,
                    type: 'question',
                    question: response.question ?? '',
                    options: response.options,
                    revealState: 'visible',
                  },
                ]
              : [],
        )
          .concat(
            pendingActions &&
            !turn.blocks.some((b) => b.type === 'actions' && b.id === pendingActions.id)
              ? [createActionBlock(pendingActions, isChangeSetApplied)]
              : [],
          )
          .concat(pendingPlan ? [pendingPlan] : []),
      }));

      setIsWaitingForResponse(false);
    },
    [updateTurn, isChangeSetApplied],
  );

  const dispatchLocalSheetActions = useCallback(
    async (turnId: string, plan: LocalSheetActionPlan, mode: AssistantMode) => {
      const runtime = runtimeRef.current.get(turnId);
      if (!runtime) return;

      const pendingActions: PendingActions = {
        id: `actions_${Date.now()}`,
        actions: plan.actions,
        explanation: plan.explanation,
      };
      runtime.pendingActions = pendingActions;

      pushHistory({
        role: 'assistant',
        content: plan.explanation,
        timestamp: new Date().toISOString(),
        type: 'answer',
      });

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStatus(turn.blocks, 'Preparing changes for review…', true, true),
      }));

      await delay(250);

      if (runtime.aborted) return;

      revealFinalResponse(turnId, { type: 'answer', answer: plan.explanation });

      const isActionMode = mode === 'action';

      if (isActionMode) {
        if (shouldPreviewActions(plan.actions, autoApplyActions)) {
          await onPreviewActions?.(plan.actions, plan.explanation, {});
        }
        updateTurn(turnId, (turn) => {
          const withoutOldPending = turn.blocks.filter(
            (b) => !(b.type === 'actions' && b.proposalStatus === 'pending'),
          );
          const alreadyHasBlock = withoutOldPending.some(
            (b) => b.type === 'actions' && b.id === pendingActions.id,
          );
          if (alreadyHasBlock) return turn;
          return {
            ...turn,
            phase: 'complete',
            blocks: [...withoutOldPending, createActionBlock(pendingActions, isChangeSetApplied)],
          };
        });
      }
    },
    [
      autoApplyActions,
      isChangeSetApplied,
      onActions,
      onClearPreview,
      onPreviewActions,
      pushHistory,
      revealFinalResponse,
      updateTurn,
    ],
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
        if (!runtime.pendingResponse) {
          await runtime.responseGate.wait();
        }
        const simpleResponse = runtime.pendingResponse;
        if (simpleResponse && !revealScheduledRef.current.has(turnId)) {
          revealFinalResponse(turnId, simpleResponse);
        }
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
        if (isAborted()) return;
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStep(turn.blocks, STEP_READING_ID, readingLabel, 'done'),
        }));
        if (!runtime.pendingResponse) {
          await runtime.responseGate.wait();
        }
        const emptyResponse = runtime.pendingResponse;
        if (emptyResponse && !revealScheduledRef.current.has(turnId)) {
          revealFinalResponse(turnId, emptyResponse);
        }
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
            signalResponse(turnId, {
              type: 'answer',
              answer: event.data.answer,
              matches: event.data.matches,
            });
            continue;
          }

          if (event.type === 'select_cell') {
            const { sheetName, row, col } = event.data;
            void navigateToCell(sheetName, row, col).catch((error) => {
              console.warn('[Cellix] Failed to select find target cell:', error);
            });
            continue;
          }

          if (event.type === 'plan') {
            const planBlock: PlanBlock = {
              id: `plan_${Date.now()}`,
              type: 'plan',
              summary: event.data.summary,
              steps: event.data.steps,
              affectedSheets: event.data.affectedSheets,
              estimatedRows: event.data.estimatedRows,
              safestApproach: event.data.safestApproach,
              prompt: event.data.prompt,
            };
            const runtime = runtimeRef.current.get(turnId);
            if (runtime) runtime.pendingPlan = planBlock;
            if (revealScheduledRef.current.has(turnId)) {
              updateTurn(turnId, (turn) => ({
                ...turn,
                blocks: [...turn.blocks.filter((b) => b.type !== 'plan' && b.type !== 'plan_only'), planBlock],
              }));
            }
            continue;
          }

          if (event.type === 'plan_only') {
            const planBlock: PlanBlock = {
              id: `plan_only_${Date.now()}`,
              type: 'plan_only',
              summary: event.data.summary,
              steps: event.data.steps,
              affectedSheets: event.data.affectedSheets,
              estimatedRows: event.data.estimatedRows,
              safestApproach: event.data.safestApproach,
              prompt: event.data.prompt,
              proposedActions: event.data.proposedActions,
              tier: event.data.tier,
            };
            const runtime = runtimeRef.current.get(turnId);
            if (runtime) runtime.pendingPlan = planBlock;
            if (revealScheduledRef.current.has(turnId)) {
              updateTurn(turnId, (turn) => ({
                ...turn,
                blocks: [...turn.blocks.filter((b) => b.type !== 'plan' && b.type !== 'plan_only'), planBlock],
              }));
            }
            continue;
          }

          if (event.type === 'matches') {
            const runtime = runtimeRef.current.get(turnId);
            if (runtime?.pendingResponse?.type === 'answer') {
              runtime.pendingResponse = {
                ...runtime.pendingResponse,
                matches: event.data.matches,
              };
            }
            if (revealScheduledRef.current.has(turnId)) {
              updateTurn(turnId, (turn) => ({
                ...turn,
                blocks: turn.blocks.map((block) =>
                  block.type === 'answer' && block.id === answerBlockId(turnId)
                    ? { ...block, matches: event.data.matches }
                    : block,
                ),
              }));
            }
            continue;
          }

          if (event.type === 'question') {
            if (!shouldAcceptIncomingClarification(activeClarificationRef.current)) {
              console.warn('[Cellix] Ignoring question event while clarification is pending');
              continue;
            }
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

          if (event.type === 'clarification') {
            if (!shouldAcceptIncomingClarification(activeClarificationRef.current)) {
              console.warn('[Cellix] Ignoring duplicate clarification while one is pending');
              continue;
            }
            pushHistory({
              role: 'assistant',
              content: `[Clarification needed]: ${event.data.question}`,
              timestamp: new Date().toISOString(),
              type: 'clarification',
            });
            setActiveClarification(event.data);
            signalResponse(turnId, {
              type: 'clarification',
              question: event.data.question,
              options: event.data.suggestions,
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

            const pendingActions: PendingActions = {
              id: `actions_${Date.now()}`,
              actions: sanitized.actions,
              explanation,
              changeSetId: event.data.changeSetId,
              changes: event.data.changes,
            };
            const runtime = runtimeRef.current.get(turnId);
            if (runtime) {
              runtime.pendingActions = pendingActions;
            }

            // Ask / Plan modes are read-only: never preview, apply, or audit
            // write actions even if the backend emits them.
            const isActionMode = (runtime?.mode ?? DEFAULT_ASSISTANT_MODE) === 'action';
            const usePreview = shouldPreviewActions(sanitized.actions, autoApplyActions);

            if (isActionMode && usePreview) {
              await onPreviewActions?.(sanitized.actions, explanation, {
                changeSetId: event.data.changeSetId,
                changes: event.data.changes,
              });
            }

            if (isActionMode && usePreview) {
              updateTurn(turnId, (turn) => {
                const withoutOldPending = turn.blocks.filter(
                  (b) => !(b.type === 'actions' && b.proposalStatus === 'pending'),
                );
                const alreadyHasBlock = withoutOldPending.some(
                  (b) => b.type === 'actions' && b.id === pendingActions.id,
                );
                if (alreadyHasBlock) return turn;
                return {
                  ...turn,
                  blocks: [
                    ...withoutOldPending,
                    createActionBlock(pendingActions, isChangeSetApplied),
                  ],
                };
              });
            }

            if (isActionMode && !usePreview && onActions) {
              await onClearPreview?.();
              await onActions(sanitized.actions, explanation, {
                changeSetId: event.data.changeSetId,
                changes: event.data.changes,
              });
              setActiveClarification(null);
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

          if (event.type === 'tool_request') {
            runtimeRef.current.get(turnId)?.analyzingGate.open();
            updateTurn(turnId, (turn) => ({
              ...turn,
              blocks: upsertThinking(
                turn.blocks,
                `Reading ${event.data.range} from ${event.data.sheet}…`,
                { loading: true },
              ),
            }));
            await handleToolRequest(event.data);
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
      isChangeSetApplied,
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
    async (
      message: string,
      sheetData: unknown[][],
      workbookContext?: WorkbookContext,
      promptContext?: string,
      sendOptions?: SendMessageOptions,
    ) => {
      const trimmed = message.trim();
      if (!trimmed) return;

      await onClearPreview?.();

      const session = ensureActiveSession();
      setActiveClarification(null);

      updateSession(session.id, (current) => ({
        ...current,
        turns: current.turns.map((turn) =>
          turn.phase === 'awaiting_input' ? { ...turn, phase: 'complete' } : turn,
        ),
      }));

      const turnId = `turn_${Date.now()}`;
      const timestamp = new Date();
      const mode = sendOptions?.mode ?? DEFAULT_ASSISTANT_MODE;
      const runtime = createRuntime(mode);
      runtimeRef.current.set(turnId, runtime);
      revealScheduledRef.current.delete(turnId);

      const localSheetPlan =
        mode === 'action' ? tryLocalSheetActions(trimmed, workbookContext, mode) : null;

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

      const nextTitle =
        session.turns.length === 0 ? truncateTabLabel(trimmed, 24) : session.title;

      updateSession(session.id, (current) => ({
        ...current,
        title: nextTitle,
        updatedAt: timestamp.toISOString(),
        turns: [...current.turns, newTurn],
      }));

      setActiveTurnId(turnId);
      setIsWaitingForResponse(true);

      abortControllerRef.current?.abort();
      abortControllerRef.current = new AbortController();

      const sheetLayout = computeSheetLayout(sheetData);
      sheetLayoutRef.current = sheetLayout;

      if (localSheetPlan) {
        try {
          await dispatchLocalSheetActions(turnId, localSheetPlan, mode);
        } catch (error: unknown) {
          const messageText =
            error instanceof Error ? error.message : 'Failed to prepare sheet deletion';
          updateTurn(turnId, (turn) => ({
            ...turn,
            phase: 'error',
            error: messageText,
            blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
          }));
        } finally {
          setIsWaitingForResponse(false);
          abortControllerRef.current = null;
          runtimeRef.current.delete(turnId);
        }
        return;
      }

      const timelinePromise = runVisualTimeline(turnId, runtime, {
        sheetIsEmpty: sheetLayout.isEmpty,
        userMessage: trimmed,
      });

      const buildPayloadContext = async () => {
        // Prefer caller-supplied context only when it already came from getContextForSend.
        // Otherwise reuse/refresh the shared pending snapshot (Spec 09 item 1).
        if (workbookContext && promptContext) {
          return { workbookContext, promptContext, sheetData };
        }
        const resolved = await getContextForSend();
        return {
          workbookContext: resolved.workbookContext,
          promptContext: resolved.promptContext,
          sheetData: resolved.sheetData.length ? resolved.sheetData : sheetData,
        };
      };

      const {
        workbookContext: resolvedWorkbookContext,
        promptContext: resolvedPromptContext,
        sheetData: resolvedSheetData,
      } = await buildPayloadContext();

      const requestPayload = prepareConversationRequestPayload(trimmed, resolvedSheetData, {
        conversationId: conversationIdRef.current,
        previousMessages: historyRef.current.slice(0, -1),
        workbookContext: resolvedWorkbookContext,
        promptContext: resolvedPromptContext,
        previewEnabled,
        refinementChangeSetId: sendOptions?.refinementChangeSetId,
        mode,
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
          credentials: 'include',
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
      dispatchLocalSheetActions,
      ensureActiveSession,
      getActiveSession,
      getUserFacingErrorMessage,
      onClearPreview,
      previewEnabled,
      processStream,
      pushHistory,
      runVisualTimeline,
      updateSession,
      updateTurn,
    ],
  );

  const answerQuestion = useCallback(
    async (
      answer: string,
      sheetData: unknown[][],
      workbookContext?: WorkbookContext,
      promptContext?: string,
      options?: SendMessageOptions,
    ) => {
      await sendMessage(answer, sheetData, workbookContext, promptContext, options);
    },
    [sendMessage],
  );

  const answerClarification = useCallback(
    async (
      answer: string,
      sheetData: unknown[][],
      workbookContext?: WorkbookContext,
      promptContext?: string,
      options?: SendMessageOptions,
    ) => {
      const trimmed = answer.trim();
      if (!trimmed) return;
      setActiveClarification(null);
      await sendMessage(trimmed, sheetData, workbookContext, promptContext, options);
    },
    [sendMessage],
  );

  const dismissClarification = useCallback(() => {
    setActiveClarification(null);
    if (activeTurnId) {
      updateTurn(activeTurnId, (turn) => ({
        ...turn,
        phase: 'complete',
      }));
    }
  }, [activeTurnId, updateTurn]);

  const applyingActionsRef = useRef(false);

  const acceptActions = useCallback(
    async (turnId: string, blockId: string) => {
      if (applyingActionsRef.current) return;

      const turn = getActiveSession()?.turns.find((t) => t.id === turnId);
      const block = turn?.blocks.find(
        (b): b is ActionBlock => b.id === blockId && b.type === 'actions',
      );
      if (!block || block.proposalStatus !== 'pending') return;

      applyingActionsRef.current = true;

      // Mark accepted immediately so Accept / preview controls disappear.
      updateTurn(turnId, (t) => ({
        ...t,
        blocks: t.blocks.map((b) =>
          b.id === blockId && b.type === 'actions'
            ? { ...b, proposalStatus: 'accepted' }
            : b,
        ),
      }));

      try {
        if (onActions) {
          await onActions(block.actions, block.explanation, {
            changeSetId: block.changeSetId,
            changes: block.changes,
          });
        }

        if (block.changeSetId) {
          onChangeSetApplied?.(block.changeSetId);
        }

        setActiveClarification(null);
      } catch (error) {
        updateTurn(turnId, (t) => ({
          ...t,
          blocks: t.blocks.map((b) =>
            b.id === blockId && b.type === 'actions'
              ? { ...b, proposalStatus: 'pending' }
              : b,
          ),
        }));
        throw error;
      } finally {
        applyingActionsRef.current = false;
      }
    },
    [onActions, onChangeSetApplied, updateTurn],
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
    setActiveClarification(null);
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

  const newChat = useCallback(() => {
    abortControllerRef.current?.abort();
    void onClearPreview?.();
    if (activeTurnId) {
      const runtime = runtimeRef.current.get(activeTurnId);
      if (runtime) {
        runtime.aborted = true;
        runtimeRef.current.delete(activeTurnId);
      }
      revealScheduledRef.current.delete(activeTurnId);
    }

    const created = createChatSession();
    activeSessionIdRef.current = created.id;
    setActiveSessionId(created.id);
    setActiveTurnId(null);
    setIsWaitingForResponse(false);
    setActiveClarification(null);
    applySessionContext(created);
    syncSessions([...sessionsRef.current, created]);
  }, [activeTurnId, applySessionContext, onClearPreview, syncSessions]);

  const clearConversation = newChat;

  const selectSession = useCallback(
    (sessionId: string) => {
      const session = sessionsRef.current.find((entry) => entry.id === sessionId);
      if (!session) return;
      activeSessionIdRef.current = sessionId;
      setActiveSessionId(sessionId);
      applySessionContext(session);
      const lastTurn = session.turns[session.turns.length - 1];
      setActiveTurnId(lastTurn?.id ?? null);
      setActiveClarification(null);
      schedulePersist();
    },
    [applySessionContext, schedulePersist],
  );

  const closeSession = useCallback(
    (sessionId: string) => {
      const nextSessions = sessionsRef.current.filter((session) => session.id !== sessionId);
      const closingActive = activeSessionIdRef.current === sessionId;

      if (closingActive) {
        abortControllerRef.current?.abort();
        void onClearPreview?.();
        setActiveClarification(null);
        setIsWaitingForResponse(false);
        setActiveTurnId(null);

        const fallback = nextSessions[nextSessions.length - 1] ?? null;
        activeSessionIdRef.current = fallback?.id ?? null;
        setActiveSessionId(fallback?.id ?? null);
        applySessionContext(fallback);
        setActiveTurnId(fallback?.turns[fallback.turns.length - 1]?.id ?? null);
      }

      syncSessions(nextSessions);
    },
    [applySessionContext, onClearPreview, syncSessions],
  );

  const selectTurn = useCallback(
    (turnId: string) => {
      const session = getActiveSession();
      if (session?.turns.some((turn) => turn.id === turnId)) {
        setActiveTurnId(turnId);
      }
    },
    [getActiveSession],
  );

  const closeTurn = useCallback(
    (sessionId: string) => {
      closeSession(sessionId);
    },
    [closeSession],
  );

  const hydrateFromStorage = useCallback(async () => {
    if (hydratedRef.current || !workbookKey) return;
    hydratedRef.current = true;

    const stored = loadChatSessions(workbookKey);
    if (!stored?.sessions.length) return;

    sessionsRef.current = stored.sessions;
    activeSessionIdRef.current = stored.activeSessionId;
    setSessions(stored.sessions);
    setActiveSessionId(stored.activeSessionId);

    const active =
      stored.sessions.find((session) => session.id === stored.activeSessionId) ??
      stored.sessions[stored.sessions.length - 1];
    if (!active) return;

    applySessionContext(active);
    setActiveTurnId(active.turns[active.turns.length - 1]?.id ?? null);

    if (active.conversationId) {
      try {
        const endpoint = getConversationByIdEndpoint(active.conversationId);
        const headers: Record<string, string> = {};
        if (endpoint.includes('.ngrok-free.app')) {
          headers['ngrok-skip-browser-warning'] = 'true';
        }
        const response = await fetch(endpoint, { headers, credentials: 'include' });
        if (response.ok) {
          const json = (await response.json()) as StoredConversation & { data?: StoredConversation };
          const stored = json.data ?? json;
          const merged = mergeSessionFromStored(active.turns, active.history, stored);
          const conversationId = stored.conversationId ?? active.conversationId;
          updateSession(active.id, (session) => ({
            ...session,
            turns: merged.turns,
            history: merged.history,
            conversationId,
          }));
          applySessionContext({
            ...active,
            turns: merged.turns,
            history: merged.history,
            conversationId,
          });
        }
      } catch (error) {
        console.warn('[Cellix] Failed to hydrate conversation from server:', error);
      }
    }
  }, [applySessionContext, updateSession, workbookKey]);

  useEffect(() => {
    void hydrateFromStorage();
  }, [hydrateFromStorage]);

  return {
    sessions,
    activeSessionId,
    turns,
    activeTurnId,
    isWaitingForResponse,
    isWaitingClarification: awaitingInput || activeClarification !== null,
    activeClarification,
    conversationId,
    sendMessage,
    answerQuestion,
    answerClarification,
    dismissClarification,
    acceptActions,
    rejectActions,
    endConversation,
    newChat,
    clearConversation,
    selectSession,
    closeSession,
    selectTurn,
    closeTurn,
    toggleThinking,
    markAnswerComplete,
  };
};
