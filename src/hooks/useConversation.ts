import { useCallback, useEffect, useRef, useState } from 'react';
import { buildWorkbookContext } from '@/services/sheetContextBuilder';
import {
  getContextForSend,
  markPendingWorkbookContextStale,
  setPendingWorkbookContext,
} from '@/utils/pendingWorkbookContext';
import { getContinueRunEndpoint, getConversationEndpoint } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';
import { RichAction } from '@/action.types';
import {
  prepareConversationRequestPayload,
  ConversationHistoryMessage,
  computeSheetLayout,
  SheetLayoutPayload,
} from '@/utils/payloadCompressor';
import { WorkbookContext } from '@/types/cellix.types';
import {
  sanitizeActions,
  blockedActionsAreDataWrites,
  CLARIFY_ROW_PLACEMENT,
} from '@/utils/actionGuard';
import { probeSheetGuardStatesSafe, sheetsCreatedInBatch } from '@/engine/sheetGuardState';
import type { OutcomeVerification } from '@/services/outcomeVerifier';
import type { RepairRequest } from '@/services/repairRequest';
import type { SheetGuardStates } from '@/engine/sheetGuardState';
import { probeExcelCapabilities } from '@/services/capabilityProbe';
import { parseSseEventBlock, SseCreditsData } from '@/utils/sseParser';
import { handleToolRequest } from '@/services/toolRequestHandler';
import { navigateToCell } from '@/services/rangeFetchService';
import { shouldAcceptIncomingClarification } from '@/utils/clarification.util';
import { TIMING, createGate, delay, waitWithMin } from '@/utils/revealQueue';
import {
  buildThoughtSummary,
  isCompletenessWarningThought,
  stillWorkingMessage,
} from '@/utils/thoughtSummary';
import { buildClientStatusMessage, isLikelyChitchat, isSimpleCreateTask } from '@/utils/statusMessage';
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
  messagesToHistory,
  messagesToTurns,
  StoredConversation,
} from '@/utils/rehydrateConversation';
import {
  deleteConversation as deleteConversationOnServer,
  fetchConversationById,
  renameConversation as renameConversationOnServer,
} from '@/services/conversationHistoryService';
import { getConversationByIdEndpoint } from '@/lib/apiConfig';
import type {
  ResponseInternalDetails,
  UserFacingSummary,
} from '@/utils/userFacingResponse';
import { resolveActionBlockCopy } from '@/utils/userFacingResponse';
import { toUserFacingApplyError } from '@/utils/toUserFacingApplyError';
import { collectCascadeRejectIds, isWaveDependencySatisfied } from '@/utils/actionWaveGating';
import { stripSheetPrefix } from '@/engine/addressUtils';
import { guardAgainstOverwrite, isOverwriteGuardError } from '@/engine/overwriteGuard';

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
  /** Resolves true when the actions were applied, false when the accept was
   *  refused (block missing/not pending, or a staged wave whose dependency has
   *  not been accepted yet). Callers must not treat a refusal as applied. */
  acceptActions: (turnId: string, blockId: string) => Promise<boolean>;
  /** Accept this step and every remaining one in a staged build — TASKS.md #160. */
  acceptAllActions: (turnId: string, fromBlockId: string) => Promise<boolean>;
  rejectActions: (turnId: string, blockId: string) => void;
  endConversation: () => void;
  newChat: () => void;
  clearConversation: () => void;
  selectSession: (sessionId: string) => void;
  closeSession: (sessionId: string) => void;
  /** Rename an open tab, and its server conversation if it has one (TASKS.md #177). */
  renameSession: (sessionId: string, title: string) => void;
  /** Delete an open tab, and its server conversation if it has one (TASKS.md #177). */
  deleteSession: (sessionId: string) => Promise<void>;
  /**
   * Delete a conversation from server-backed history that isn't necessarily an
   * open tab (TASKS.md #177) — closes the tab too if it happens to be open.
   */
  deleteHistoryConversation: (conversationId: string) => Promise<void>;
  /**
   * Open a past conversation from server-backed history (TASKS.md #172).
   * Resolves false when the fetch failed, so the caller can surface that rather
   * than silently showing an empty thread.
   */
  openConversationFromHistory: (conversationId: string) => Promise<boolean>;
  /** True while a history conversation's full body is being fetched. */
  isLoadingHistoryConversation: boolean;
  selectTurn: (turnId: string) => void;
  closeTurn: (turnId: string) => void;
  toggleThinking: (turnId: string, blockId: string) => void;
  markAnswerComplete: (turnId: string, blockId: string) => void;
}

export interface UseConversationOptions {
  workbookKey?: string;
  /** Durable per-workbook identity (TASKS.md #22-23), distinct from workbookKey above. */
  workbookId?: string;
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
  /** Emitted once per completed debit (CREDIT_SYSTEM_SCHEMA.md §6) — lets the
   *  balance indicator update live instead of only on next mount/refetch. */
  onCredits?: (event: SseCreditsData) => void;
}

export interface SendMessageOptions {
  refinementChangeSetId?: string;
  mode?: AssistantMode;
  /**
   * When set, re-runs this request against the existing turn with this id
   * instead of appending a new one — used by regenerate/edit-and-resend so
   * the message stays anchored in place rather than duplicating in the
   * thread. The turn must already exist in the active session.
   */
  regenerateTurnId?: string;
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
  userFacingSummary?: UserFacingSummary;
  internalDetails?: ResponseInternalDetails;
  dependsOnChangeSetId?: string;
  irreversibleActionTypes?: string[];
  stepIndex?: number;
  stepTotal?: number;
  stepLabel?: string;
  /** Step-wise run this wave belongs to — TASKS.md #153. */
  runId?: string;
  stepwise?: boolean;
}

export interface PreviewActionsMeta {
  changeSetId?: string;
  changes?: CellChange[];
  userFacingSummary?: UserFacingSummary;
  internalDetails?: ResponseInternalDetails;
  irreversibleActionTypes?: string[];
  /**
   * Called after the post-apply read-back (TASKS.md #150) with what the
   * workbook actually looks like. `message` is null when everything matched —
   * a clean verification stays silent, a divergent one must not.
   */
  onOutcomeVerified?: (
    verification: OutcomeVerification,
    message: string | null,
    /**
     * A ready-to-send follow-up that repairs the erroring cells, or null when
     * there is nothing a formula fix can address. TASKS.md #168.
     */
    repair: RepairRequest | null,
  ) => void;
}

interface TurnRuntime {
  analyzingGate: ReturnType<typeof createGate>;
  responseGate: ReturnType<typeof createGate>;
  pendingResponse: PendingResponse | null;
  pendingActions: PendingActions | null;
  pendingPlan: PlanBlock | null;
  aborted: boolean;
  mode: AssistantMode;
  /**
   * True once a real backend `status`/`thinking` SSE event has been folded
   * into this turn's thought log. The scripted reading/analyzing/composing
   * narration below is generic filler shown while waiting for the backend —
   * once the backend has said something *real* about what it's doing, the
   * filler must stop overwriting it (that was overwriting a genuine process
   * description with a canned one right before the answer revealed).
   */
  hasLiveThinking: boolean;
  /**
   * `Date.now()` of the last time this turn's status/thinking line actually
   * changed — a real backend event OR the "still working" ticker below.
   * Drives the ticker's own escalation: it only writes a new tier once enough
   * time has passed since the LAST update of either kind, so it never fights
   * a real backend message that just arrived.
   */
  lastLiveUpdateAt: number;
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
        // Prefer agent live log over synthetic summary so Blocked / progress messages remain.
        const keepContent = block.content.trim().length > 0 ? block.content.trim() : summary;
        return {
          ...block,
          content: keepContent,
          loading: false,
          // Tap-to-expand only: never force it open here, even for notable
          // content (blocked/verification) — but if the user already tapped
          // it open mid-stream, respect that and don't force it shut either.
          visible: true,
        };
      }
      return block;
    });
}

function appendThinkingLog(
  blocks: TurnBlock[],
  message: string,
  opts: { loading?: boolean; expanded?: boolean } = {},
): TurnBlock[] {
  const text = message.trim();
  if (!text) return blocks;
  const existing = blocks.find((b): b is ThinkingBlock => b.type === 'thinking');
  const prev = existing?.content?.trim() ?? '';
  // Dedup consecutive repeats (status + thinking often carry the same line).
  if (prev.endsWith(text)) {
    return upsertThinking(blocks, prev, {
      loading: opts.loading ?? true,
      // Default collapsed: live agent chatter (column lists, range reads) is
      // noise until something goes wrong — the status line above already
      // shows the current step, so re-showing it expanded here is redundant.
      expanded: opts.expanded ?? false,
      visible: true,
    });
  }
  const next = prev ? `${prev}\n\n${text}` : text;
  // Cap growth so a long agent run does not blow the UI.
  const capped =
    next.length > 6000 ? `…\n\n${next.slice(next.length - 5800)}` : next;
  return upsertThinking(blocks, capped, {
    loading: opts.loading ?? true,
    expanded: opts.expanded ?? false,
    visible: true,
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
      : (existing?.type === 'thinking' ? existing.expanded : false);
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
    hasLiveThinking: false,
    lastLiveUpdateAt: Date.now(),
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
    userFacingSummary: pending.userFacingSummary,
    internalDetails: pending.internalDetails,
    dependsOnChangeSetId: pending.dependsOnChangeSetId,
    irreversibleActionTypes: pending.irreversibleActionTypes,
    stepIndex: pending.stepIndex,
    stepTotal: pending.stepTotal,
    stepLabel: pending.stepLabel,
    runId: pending.runId,
    stepwise: pending.stepwise,
  };
}

function normalizeLocalCellAddress(address: string): string {
  return stripSheetPrefix(address).trim().toUpperCase();
}

function extractOverwriteGuardWriteCells(action: SheetAction): string[] {
  switch (action.type) {
    case 'SET_CELL':
    case 'SET_FORMULA':
      return action.address ? [normalizeLocalCellAddress(action.address)] : [];
    case 'BATCH_SET':
      return (action.operations ?? [])
        .map((op) => op.address)
        .filter(Boolean)
        .map((addr) => normalizeLocalCellAddress(String(addr)));
    default:
      return [];
  }
}

async function preflightOverwriteBlockedActions(
  actions: SheetAction[],
  changes: CellChange[],
  /**
   * Live per-sheet facts already gathered by `probeSheetGuardStatesSafe` a few
   * lines above this call. Reused rather than re-probed — TASKS.md #162.
   */
  sheetStates?: SheetGuardStates,
): Promise<{
  safeActions: SheetAction[];
  safeChanges: CellChange[];
  blockedMessage: string | null;
}> {
  // If Office.js isn't available (tests / non-taskpane env), fail open.
  if (typeof Excel === 'undefined' || typeof Excel.run !== 'function') {
    return { safeActions: actions, safeChanges: changes, blockedMessage: null };
  }

  // Sheets this very batch creates do not exist yet, so nothing in them can be
  // overwritten — guarding them is meaningless by definition, and actively
  // harmful: `resolveWorksheet` uses the THROWING `worksheets.getItem()`, which
  // raises ItemNotFound for a sheet that is not there. That error is not an
  // OverwriteGuardError, so the old `else { throw error }` below rethrew it out
  // of the SSE handler and killed the whole turn *before* the action card was
  // ever created — the user saw a bare "The requested resource doesn't exist."
  // and no Accept button at all.
  //
  // This was latent until TASKS.md #141 merged the staged accept waves. The old
  // two-wave split accidentally guaranteed ordering: wave 1 (ADD_SHEETs, which
  // are not overwrite-guarded) reached the structural preview and created the
  // sheets, so wave 2's preflight always found them present. Merging the waves
  // removed that guarantee. See TASKS.md #147.
  const createdHere = sheetsCreatedInBatch(actions);

  /**
   * A sheet that does not exist RIGHT NOW cannot be overwritten, whoever
   * creates it and whenever.
   *
   * #147 skipped only sheets created in the SAME batch, which was enough while
   * a build was one batch. TASKS.md #160's staging broke that: the writes now
   * arrive in step 2 and the ADD_SHEETs live in step 1, so
   * `sheetsCreatedInBatch(step2)` is empty and every write hit the throwing
   * `worksheets.getItem()` — 16 failed Office.js round trips and 16 identical
   * "probe failed" warnings in a single observed run. It degraded safely (that
   * is #147 working) but the work and the noise were pure waste.
   *
   * The probe that ran moments earlier already knew these sheets were absent;
   * the information was simply not passed here. TASKS.md #162.
   */
  const knownAbsent = new Set<string>();
  for (const [key, state] of sheetStates ?? []) {
    if (!state.exists) knownAbsent.add(key);
  }

  return Excel.run(async (ctx) => {
    const activeWs = ctx.workbook.worksheets.getActiveWorksheet();
    activeWs.load('name');
    await ctx.sync();
    const activeSheetName = activeWs.name ?? '';

    const safeActions: SheetAction[] = [];
    const blocked: Array<{ action: SheetAction; message: string }> = [];

    for (const action of actions) {
      const targetSheet = String(action.sheetName ?? '').trim().toLowerCase();
      if (targetSheet && (createdHere.has(targetSheet) || knownAbsent.has(targetSheet))) {
        safeActions.push(action);
        continue;
      }
      try {
        // Dry-run: guardAgainstOverwrite only loads values & throws on occupancy.
        // These are still legacy-shaped (normalization happens at apply time), but
        // the guard reads only the address/range/value fields both shapes share —
        // and RichActionEngine re-runs the authoritative guard before writing.
        await guardAgainstOverwrite(action as unknown as RichAction, ctx);
        safeActions.push(action);
      } catch (error) {
        if (isOverwriteGuardError(error)) {
          blocked.push({ action, message: error.message });
          continue;
        }
        // Any other error here is an Office.js problem with the *probe*, not a
        // finding about the user's data. This pass is advisory UX only —
        // `RichActionEngine.dispatch` re-runs the authoritative guard immediately
        // before every real write (ARCHITECTURE.md AD-1) — so degrade to
        // "let Accept decide" rather than destroying the turn.
        console.warn('[Cellix] Overwrite preflight probe failed; deferring to apply-time guard:', error);
        safeActions.push(action);
      }
    }

    if (blocked.length === 0) {
      return { safeActions, safeChanges: changes, blockedMessage: null };
    }

    // Filter ChangeSet diff/highlight cells for the blocked actions only.
    // (We start conservative: only address exact cells for SET_* and BATCH_SET.)
    const blockedKeys = new Set<string>();
    for (const { action } of blocked) {
      const sheetName = String(action.sheetName ?? activeSheetName).trim();
      for (const cell of extractOverwriteGuardWriteCells(action)) {
        blockedKeys.add(`${sheetName}|${cell}`);
      }
    }

    const hasAnyBlockedCellAddresses = blockedKeys.size > 0;
    const safeChanges = hasAnyBlockedCellAddresses
      ? changes.filter(
          (ch) => !blockedKeys.has(`${ch.sheet}|${String(ch.cell).trim().toUpperCase()}`),
        )
      : changes;

    return {
      safeActions,
      safeChanges,
      blockedMessage: blocked[0]?.message ?? null,
    };
  });
}

export const useConversation = (options: UseConversationOptions = {}): UseConversationReturn => {
  const {
    workbookKey = 'workbook',
    workbookId,
    onActions,
    onPreviewActions,
    onClearPreview,
    onChangeSetApplied,
    autoApplyActions = false,
    previewEnabled = true,
    isChangeSetApplied,
    onCredits,
  } = options;

  const [sessions, setSessions] = useState<ChatSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<string | null>(null);
  const [isLoadingHistoryConversation, setIsLoadingHistoryConversation] = useState(false);
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
  // workbookId is minted asynchronously (client-side, TASKS.md #22) and passed
  // in as a prop that may change after mount without a re-render of sendMessage's
  // closure — a ref keeps sendMessage reading the latest value without pulling
  // workbookId into its (already large) dependency array.
  const workbookIdRef = useRef<string | undefined>(workbookId);
  useEffect(() => {
    workbookIdRef.current = workbookId;
  }, [workbookId]);
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
      // Ask / Plan modes are read-only: never surface an Accept-able action
      // block, even if pendingActions got populated (e.g. by a backend that
      // shouldn't have emitted write actions on a read-only turn in the first
      // place — this is the last line of defense, not the only one).
      const isActionModeForReveal = (runtimeForReveal?.mode ?? DEFAULT_ASSISTANT_MODE) === 'action';
      const pendingActions = isActionModeForReveal ? runtimeForReveal?.pendingActions : undefined;
      const pendingPlan = runtimeForReveal?.pendingPlan;

      updateTurn(turnId, (turn) => {
        const thinkingExisting = turn.blocks.find(
          (b): b is ThinkingBlock => b.type === 'thinking',
        );
        const preservedThought =
          thinkingExisting?.content?.trim() &&
          thinkingExisting.content.trim().length > 12 &&
          !/^thought process/i.test(thinkingExisting.content)
            ? thinkingExisting.content.trim()
            : buildThoughtSummary(turn.userMessage, 'final');

        const finalized = finalizeSteps(
          upsertThinking(turn.blocks, preservedThought, {
            loading: false,
            expanded: isCompletenessWarningThought(preservedThought),
            visible: true,
          }),
          turn.userMessage,
        );
        const actionBlocks = finalized.filter((b): b is ActionBlock => b.type === 'actions');
        const withoutActions: TurnBlock[] = finalized.filter((b) => b.type !== 'actions');
        const nextActionBlocks =
          pendingActions && !actionBlocks.some((b) => b.id === pendingActions.id)
            ? [...actionBlocks, createActionBlock(pendingActions, isChangeSetApplied)]
            : actionBlocks;

        return {
          ...turn,
          phase:
            response.type === 'question' || response.type === 'clarification'
              ? 'awaiting_input'
              : 'complete',
          blocks: withoutActions
            .concat(
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
            .concat(nextActionBlocks)
            .concat(pendingPlan ? [pendingPlan] : []),
        };
      });

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

  /**
   * Fire-and-forget background loop that keeps the thinking line moving
   * during a long, opaque wait — a single non-streaming Planner call for a
   * large build can run 30s-4min with nothing for the backend to report in
   * between. Polls rather than a single timer so it naturally stops as soon
   * as `responseGate` opens (a real response arrived) without needing manual
   * cleanup/cancellation plumbing — the loop just exits on its own next tick.
   *
   * Never overwrites a real backend update: it only writes when the time
   * since the LAST update of any kind (`lastLiveUpdateAt`, stamped by both
   * real events and this ticker's own writes) has crossed the next escalation
   * tier, so a real status event arriving resets the clock exactly as if the
   * ticker itself had just spoken.
   */
  const runStillWorkingTicker = useCallback(
    (turnId: string, runtime: TurnRuntime) => {
      const POLL_MS = 3000;
      const tick = async () => {
        while (!runtime.aborted && !runtime.responseGate.isOpen()) {
          await delay(POLL_MS);
          if (runtime.aborted || runtime.responseGate.isOpen()) break;
          const elapsed = Date.now() - runtime.lastLiveUpdateAt;
          const message = stillWorkingMessage(elapsed);
          if (!message) continue;
          runtime.lastLiveUpdateAt = Date.now();
          updateTurn(turnId, (turn) => ({
            ...turn,
            blocks: appendThinkingLog(upsertStatus(turn.blocks, message, true, true), message, {
              loading: true,
              expanded: false,
            }),
          }));
        }
      };
      void tick();
    },
    [updateTurn],
  );

  const runVisualTimeline = useCallback(
    async (turnId: string, runtime: TurnRuntime, opts: TimelineOptions) => {
      runStillWorkingTicker(turnId, runtime);
      const isAborted = () => runtime.aborted;

      // Greetings/small talk (backend's own CHITCHAT route — see
      // `isLikelyChitchat`'s docblock) skip workbook context entirely and
      // reply in well under a second via plain `chunk` streaming. Running the
      // full "Reading your worksheet… Analyzing your spreadsheet…"
      // choreography (with its hard minimum delays meant to pace a real
      // multi-second build) for a bare "hi" was pure wasted wait — nothing
      // about it was ever true for that message. Show a single lightweight
      // "Thinking…" step instead and resolve as soon as the real reply lands.
      if (isLikelyChitchat(opts.userMessage)) {
        await delay(150);
        if (isAborted()) return;
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStep(turn.blocks, STEP_READING_ID, 'Thinking…', 'running'),
        }));
        runtime.analyzingGate.open();
        await waitWithMin(runtime.responseGate, 200);
        if (isAborted()) return;
        updateTurn(turnId, (turn) => ({
          ...turn,
          blocks: upsertStep(turn.blocks, STEP_READING_ID, 'Thinking…', 'done'),
        }));
        if (!runtime.pendingResponse) {
          await runtime.responseGate.wait();
        }
        const chitchatResponse = runtime.pendingResponse;
        if (chitchatResponse && !revealScheduledRef.current.has(turnId)) {
          revealFinalResponse(turnId, chitchatResponse);
        }
        return;
      }

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
        // No `expanded` override — tap-to-expand only, so a fresh block
        // starts collapsed (upsertThinking's own default) and a block the
        // user already opened stays open across phase transitions.
        blocks: upsertThinking(turn.blocks, buildThoughtSummary(turn.userMessage, 'reading'), {
          loading: true,
          visible: true,
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

      updateTurn(turnId, (turn) => {
        // Once the backend has said something real about what it's doing,
        // keep that instead of overwriting it with generic filler text.
        const existing = turn.blocks.find((b): b is ThinkingBlock => b.type === 'thinking');
        const content =
          runtime.hasLiveThinking && existing?.content
            ? existing.content
            : buildThoughtSummary(turn.userMessage, 'analyzing');
        return {
          ...turn,
          blocks: upsertThinking(turn.blocks, content, {
            loading: true,
            visible: true,
          }),
        };
      });

      await waitWithMin(runtime.responseGate, TIMING.analyzingMinRun);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => ({
        ...turn,
        blocks: upsertStep(turn.blocks, STEP_ANALYZING_ID, ANALYZING_LABEL, 'done'),
      }));

      await delay(TIMING.pauseBeforeComposing);
      if (isAborted()) return;

      updateTurn(turnId, (turn) => {
        const existing = turn.blocks.find((b): b is ThinkingBlock => b.type === 'thinking');
        const content =
          runtime.hasLiveThinking && existing?.content
            ? existing.content
            : buildThoughtSummary(turn.userMessage, 'composing');
        return {
          ...turn,
          blocks: upsertThinking(
            upsertStatus(turn.blocks, 'Composing response…', true, true),
            content,
            { loading: true, visible: true },
          ),
        };
      });

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
    [revealFinalResponse, runStillWorkingTicker, updateTurn],
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
            const runtime = runtimeRef.current.get(turnId);
            runtime?.analyzingGate.open();
            if (runtime) {
              runtime.hasLiveThinking = true;
              runtime.lastLiveUpdateAt = Date.now();
            }
            updateTurn(turnId, (turn) => ({
              ...turn,
              blocks: appendThinkingLog(
                upsertStatus(turn.blocks, event.data.message, true, true),
                event.data.message,
              ),
            }));
            continue;
          }

          // Live agent status / thinking — show like Cursor/agent progress, not idle spinner.
          if (event.type === 'status' || event.type === 'thinking') {
            // Both carry { message }; there is no text variant on these events.
            const message = typeof event.data.message === 'string' ? event.data.message : '';
            if (!message.trim()) continue;
            const runtime = runtimeRef.current.get(turnId);
            runtime?.analyzingGate.open();
            if (runtime) {
              runtime.hasLiveThinking = true;
              runtime.lastLiveUpdateAt = Date.now();
            }
            updateTurn(turnId, (turn) => ({
              ...turn,
              blocks: appendThinkingLog(
                upsertStatus(turn.blocks, message, true, true),
                message,
              ),
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

          if (event.type === 'credits') {
            onCredits?.(event.data);
            continue;
          }

          if (event.type === 'wave_ready') {
            // Still needed for the ORIGINAL send flow: `runVisualTimeline` is
            // blocked on `runtime.responseGate.wait()`, and `signalResponse`
            // is what opens that gate so it can proceed to
            // `revealFinalResponse` (which sets phase itself). A continuation
            // never runs `runVisualTimeline` at all, so for that path this
            // call is inert — harmless, not load-bearing.
            const runtime = runtimeRef.current.get(turnId);
            if (runtime && !runtime.pendingResponse) {
              signalResponse(turnId, { type: 'answer', answer: '' });
            }

            // The explicit restore a continuation actually needs.
            // `continueStepwiseRun` set `phase: 'processing'` when it started
            // (so the wave's own live status/thinking updates could render,
            // see that function's own comment), and — unlike the send flow —
            // NOTHING else in the continuation path ever calls
            // `revealFinalResponse` to set it back. Whether or not another
            // wave follows, this stream is ending and whatever card it
            // produced needs to actually become visible.
            //
            // Live incident (Sept 9, 2026): a `hasMore: true` card generated
            // via `/continue` was correct in every server log but never
            // rendered — invisible, no error, nothing in the DOM. Root cause:
            // this branch used to rely SOLELY on `signalResponse` above, which
            // — for a continuation — has no `runVisualTimeline` waiting on it,
            // so it does nothing. `phase` stayed stuck at 'processing',
            // `isTurnPresentationComplete` returned false,
            // `previewActionsReady`/`showActionButtons` went false for the
            // active turn, and `ActionResponseCard` renders every PENDING
            // block as `null` whenever `showActionButtons` is false.
            updateTurn(turnId, (turn) => {
              const withoutStatusBlocks = withoutStatus(turn.blocks);
              // A completeness warning (TASKS.md #171/#187 — "proceeded under
              // an assumption", "could not fully plan N steps") landed in the
              // thinking log via `appendThinkingLog` as this wave streamed in,
              // same as it always does. But `revealFinalResponse` — the ONLY
              // place that decides whether the thinking block starts expanded
              // — never runs for a continuation (see this branch's own
              // comment). Without this, that warning is technically present
              // but collapsed behind a disclosure nothing ever opens: a live
              // report where a truncated plan got silently pruned from 10
              // subtasks to 3, and the user saw two "Applied" cards with zero
              // indication 70% of the request never got built.
              const thinking = withoutStatusBlocks.find(
                (b): b is ThinkingBlock => b.type === 'thinking',
              );
              const blocks =
                thinking && isCompletenessWarningThought(thinking.content)
                  ? withoutStatusBlocks.map((b) =>
                      b.id === thinking.id ? { ...b, expanded: true } : b,
                    )
                  : withoutStatusBlocks;
              return { ...turn, phase: 'complete', blocks };
            });
            setIsWaitingForResponse(false);
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
            // Probe the live sheets this batch targets. Without it this pass
            // graded every action against the *active* sheet's layout, while
            // the apply-time pass graded them against nothing at all — two
            // guards, two answers, and a batch that previewed as 189 changes
            // then applied as 69. See TASKS.md #137.
            const sheetStates = await probeSheetGuardStatesSafe(event.data.actions);
            const sanitized = sanitizeActions(
              event.data.actions,
              sheetLayoutRef.current ?? undefined,
              { sheetStates },
            );

            // Only the "where does the NEW ROW go?" card is for blocked data writes.
            // Never use it for pure header formatting (FORMAT_RANGE on row 0) — that
            // used to fire every time a valid header fill was wrongly blocked.
            if (
              sanitized.requiresClarification &&
              sanitized.actions.length === 0 &&
              blockedActionsAreDataWrites(sanitized.blocked)
            ) {
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

            if (sanitized.requiresClarification && sanitized.actions.length === 0) {
              // Cosmetic/other actions fully blocked — honest failure, not row-insert UI.
              const message =
                'I could not apply that change without altering protected cells. Rephrase the formatting request (e.g. "highlight the header row light green").';
              pushHistory({
                role: 'assistant',
                content: message,
                timestamp: new Date().toISOString(),
                type: 'answer',
              });
              signalResponse(turnId, {
                type: 'answer',
                answer: message,
              });
              continue;
            }

            const explanation =
              sanitized.warnings.length > 0
                ? `${event.data.explanation} (${sanitized.warnings.join(' ')})`
                : event.data.explanation;

            const runtime = runtimeRef.current.get(turnId);
            // Ask / Plan modes are read-only: never preview, apply, or audit
            // write actions even if the backend emits them.
            const isActionMode = (runtime?.mode ?? DEFAULT_ASSISTANT_MODE) === 'action';
            const usePreview = shouldPreviewActions(sanitized.actions, autoApplyActions);

            const basePendingId = `actions_${Date.now()}`;

            let actionsForPreview = sanitized.actions;
            let changesForPreview = event.data.changes;
            let blockedGuardMessage: string | null = null;

            if (isActionMode && usePreview && sanitized.actions.length > 0) {
              const preflight = await preflightOverwriteBlockedActions(
                sanitized.actions,
                event.data.changes ?? [],
                sheetStates,
              );
              actionsForPreview = preflight.safeActions;
              changesForPreview = preflight.safeChanges;
              blockedGuardMessage = preflight.blockedMessage;
            }

            const pendingActions: PendingActions = {
              id: basePendingId,
              actions: actionsForPreview,
              explanation,
              changeSetId: event.data.changeSetId,
              changes: changesForPreview,
              userFacingSummary: event.data.userFacingSummary,
              internalDetails: event.data.internalDetails,
              dependsOnChangeSetId: event.data.dependsOnChangeSetId,
              stepIndex: event.data.stepIndex,
              stepTotal: event.data.stepTotal,
              stepLabel: event.data.stepLabel,
              irreversibleActionTypes: event.data.irreversibleActionTypes,
              runId: event.data.runId,
              stepwise: event.data.stepwise,
            };

            if (runtime && isActionMode && pendingActions.actions.length > 0) {
              runtime.pendingActions = pendingActions;
            }

            const previewCopy = resolveActionBlockCopy({
              userFacingSummary: event.data.userFacingSummary,
              explanation,
              actions: pendingActions.actions,
              changes: pendingActions.changes,
            });

            // Staged accept waves: only the first (dependency-free) wave gets a
            // live soft-preview. previewManager holds one active preview at a
            // time — soft-previewing a later wave would reject/clear the
            // earlier wave's still-pending preview out from under it. Later
            // waves are added as pending blocks below and applied directly on
            // Accept, once their dependency is satisfied.
            if (
              isActionMode &&
              usePreview &&
              pendingActions.actions.length > 0 &&
              !pendingActions.dependsOnChangeSetId
            ) {
              await onPreviewActions?.(pendingActions.actions, previewCopy.headline, {
                changeSetId: event.data.changeSetId,
                changes: pendingActions.changes,
                userFacingSummary: event.data.userFacingSummary,
                internalDetails: event.data.internalDetails,
                irreversibleActionTypes: pendingActions.irreversibleActionTypes,
              });
            }

            if (isActionMode && usePreview) {
              updateTurn(turnId, (turn) => {
                // Stack multiple Accept packages for large multi-wave work — don't replace prior pending.
                const existingIds = new Set(
                  turn.blocks.filter((b) => b.type === 'actions').map((b) => b.id),
                );

                const hasSafeActions = pendingActions.actions.length > 0;
                const blockedBlockId = `${basePendingId}_blocked`;

                const blocksToAdd = [
                  hasSafeActions && !existingIds.has(pendingActions.id)
                    ? createActionBlock(pendingActions, isChangeSetApplied)
                    : null,
                  blockedGuardMessage && !existingIds.has(blockedBlockId)
                    ? ({
                        id: blockedBlockId,
                        type: 'actions',
                        actions: [],
                        explanation: blockedGuardMessage,
                        proposalStatus: 'rejected',
                        changes: [],
                      } satisfies ActionBlock)
                    : null,
                ].filter(Boolean) as TurnBlock[];

                return {
                  ...turn,
                  blocks: appendThinkingLog(
                    upsertStatus(
                      [...turn.blocks, ...blocksToAdd],
                      hasSafeActions
                        ? `Ready for review: ${pendingActions.actions.length} change(s) — Accept to apply`
                        : 'Could not prepare applyable changes for this step',
                      false,
                      true,
                    ),
                    hasSafeActions
                      ? `Preview ready (${pendingActions.actions.length} changes) — Accept when ready.`
                      : blockedGuardMessage ?? 'No applyable changes for this package.',
                    { loading: false },
                  ),
                };
              });
            }

            if (isActionMode && !usePreview && onActions) {
              await onClearPreview?.();
              await onActions(sanitized.actions, previewCopy.headline, {
                changeSetId: event.data.changeSetId,
                changes: event.data.changes,
                userFacingSummary: event.data.userFacingSummary,
                internalDetails: event.data.internalDetails,
              });
              setActiveClarification(null);
              updateTurn(turnId, (turn) => ({
                ...turn,
                blocks: turn.blocks.map((b) =>
                  b.id === basePendingId && b.type === 'actions'
                    ? { ...b, proposalStatus: 'accepted' }
                    : b,
                ),
              }));
            }
            continue;
          }

          if (event.type === 'tool_request') {
            const runtime = runtimeRef.current.get(turnId);
            runtime?.analyzingGate.open();
            if (runtime) {
              runtime.hasLiveThinking = true;
              runtime.lastLiveUpdateAt = Date.now();
            }
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
            // CREDIT_SYSTEM.md CD-4 — the gate check blocked dispatch before any
            // LLM call. Clear, non-alarming copy per CREDIT_SYSTEM.md §5, not the
            // generic apply-error mapping below.
            const isInsufficientCredit = event.data.code === 'INSUFFICIENT_CREDIT';
            const message = isInsufficientCredit
              ? `You're out of credits for this action${
                  typeof event.data.requiredCredits === 'number'
                    ? ` (needs ${event.data.requiredCredits}, have ${event.data.availableBalance ?? 0})`
                    : ''
                }. Add credits or upgrade your plan to continue.`
              : toUserFacingApplyError(event.data.message);
            updateTurn(turnId, (turn) => ({
              ...turn,
              phase: 'error',
              // Map here too: this path rendered raw engine/host strings like
              // Office.js "The requested resource doesn't exist." straight into
              // chat. The mapper passes clean short messages through unchanged.
              error: message,
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
      onCredits,
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

      const regenerateTurnId = sendOptions?.regenerateTurnId;
      const turnId = regenerateTurnId ?? `turn_${Date.now()}`;
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

      // Regenerate/edit-and-resend: replace the existing turn in place
      // (same id, same position) instead of appending a new one, so the
      // message doesn't duplicate itself further down the thread.
      updateSession(session.id, (current) => {
        const existingIndex = regenerateTurnId
          ? current.turns.findIndex((t) => t.id === regenerateTurnId)
          : -1;
        const turns =
          existingIndex !== -1
            ? current.turns.map((t, i) => (i === existingIndex ? newTurn : t))
            : [...current.turns, newTurn];
        const nextTitle =
          existingIndex === -1 && current.turns.length === 0
            ? truncateTabLabel(trimmed, 24)
            : current.title;
        return {
          ...current,
          title: nextTitle,
          updatedAt: timestamp.toISOString(),
          turns,
        };
      });

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

      // TASKS.md #152 — tell the server what this Excel can actually do, so a
      // formula family is chosen against a probed fact rather than an
      // assumption. Cached per session; never throws.
      const excelCapabilities = await probeExcelCapabilities();

      const requestPayload = prepareConversationRequestPayload(trimmed, resolvedSheetData, {
        conversationId: conversationIdRef.current,
        workbookId: workbookIdRef.current,
        previousMessages: historyRef.current.slice(0, -1),
        workbookContext: resolvedWorkbookContext,
        promptContext: resolvedPromptContext,
        previewEnabled,
        refinementChangeSetId: sendOptions?.refinementChangeSetId,
        mode,
        excelCapabilities,
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

  /**
   * Advances a step-wise Tier 3 run (TASKS.md #153, STEPWISE_EXECUTION.md §3).
   *
   * The backend has generated NOTHING past the card being decided, so this call
   * is what causes the next wave to exist at all. It streams back exactly like
   * the original request, and is pumped through the same `processStream` onto
   * the same turn, so the next wave's card appends to this conversation turn
   * rather than opening a new one.
   *
   * Never throws: a continuation that fails leaves the already-accepted waves
   * applied and the turn showing why it stopped, which is strictly better than
   * unwinding work the user already approved.
   */
  const continueStepwiseRun = useCallback(
    async (
      turnId: string,
      runId: string,
      decision: 'accepted' | 'rejected' | 'skipped',
    ): Promise<void> => {
      const runtime = runtimeRef.current.get(turnId) ?? createRuntime('action');
      runtimeRef.current.set(turnId, runtime);
      runtime.aborted = false;
      // Reset the clock the ticker measures against — otherwise a user who took
      // five minutes to click Accept would make it think five minutes of
      // silence had already passed and jump straight to the most escalated
      // "still going" tier the instant this continuation starts.
      runtime.lastLiveUpdateAt = Date.now();
      setIsWaitingForResponse(true);
      revealScheduledRef.current.delete(turnId);
      runStillWorkingTicker(turnId, runtime);

      // The turn re-enters "working" — the build is not finished, and showing
      // it as complete between waves would be the same false-completeness the
      // staged-accept work keeps guarding against.
      //
      // This is also what makes progress VISIBLE at all: `TurnRenderer` hides
      // every status/step/thinking block while `turn.phase === 'complete'`
      // (`hideProgress`), which it still was here — set by the PREVIOUS wave's
      // resolution and never reset. Without this, every `status`/`thinking`
      // SSE event the next wave sends was written into turn state correctly
      // but silently suppressed at render time: a live "no loading or
      // anything" report (Sept 8, 2026) traced to exactly this — the backend
      // was genuinely working, the UI just never showed it.
      //
      // An immediate status line fills the gap between the click and the
      // first real backend event (which can be seconds away for a large
      // wave) — the moment that most reads as "did anything happen?".
      updateTurn(turnId, (turn) => ({
        ...turn,
        phase: 'processing',
        blocks: upsertStatus(
          turn.blocks.filter((b) => b.type !== 'answer' || b.id !== answerBlockId(turnId)),
          decision === 'accepted' ? 'Preparing the next step…' : 'Continuing…',
          true,
          true,
        ),
      }));

      // Readback (observed post-apply sheet state, so the next wave plans
      // against reality rather than the shadow workbook's prediction) is NOT
      // sent yet. This hook's `getContextForSend().workbookContext.sheets` is
      // `SheetSnapshot[]` (types/cellix.types.ts — `sheetName`, `colCount`,
      // `headers`, `sampleData`) which shares no fields with the backend's
      // `SheetContext[]` (agents/types/agent.types.ts — `name`, `values`,
      // `formulas`, `numberFormats`). Sending it produced a live "Cannot read
      // properties of undefined (reading 'length')" crash: the backend merged
      // a malformed sheet object into its context and the next wave's Executor
      // read `.values.length` on it. The backend now rejects a mismatched
      // shape defensively too (`AgentRunStateService.applyReadback`), but
      // there is no reason to send data that can never be used until a real
      // translator between the two shapes exists — that is a distinct,
      // separately-scoped task, not something to bolt on here.
      const readback: unknown[] | undefined = undefined;

      try {
        const endpoint = getContinueRunEndpoint();
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (endpoint.includes('.ngrok-free.app')) {
          headers['ngrok-skip-browser-warning'] = 'true';
        }

        const response = await fetch(endpoint, {
          method: 'POST',
          headers,
          credentials: 'include',
          body: JSON.stringify({ runId, decision, readback }),
        });

        if (!response.ok) {
          throw new Error(await getUserFacingErrorMessage(response));
        }

        await processStream(response, turnId);
      } catch (error) {
        const message =
          error instanceof Error ? error.message : 'Could not continue this build';
        console.error('[Cellix] Stepwise continue failed:', message);
        updateTurn(turnId, (turn) => ({
          ...turn,
          error: `${message} — the steps you already accepted are still applied.`,
          blocks: finalizeSteps(withoutStatus(turn.blocks), turn.userMessage),
        }));
      } finally {
        setIsWaitingForResponse(false);
      }
    },
    [getUserFacingErrorMessage, processStream, runStillWorkingTicker, updateTurn],
  );

  const acceptActions = useCallback(
    async (turnId: string, blockId: string): Promise<boolean> => {
      if (applyingActionsRef.current) return false;

      const turn = getActiveSession()?.turns.find((t) => t.id === turnId);
      const block = turn?.blocks.find(
        (b): b is ActionBlock => b.id === blockId && b.type === 'actions',
      );
      if (!block || block.proposalStatus !== 'pending') return false;

      // Defense in depth: the Accept button is disabled while a dependency is
      // unmet (see TurnRenderer), but never apply a staged wave out of order
      // even if something else calls acceptActions directly.
      const siblingActionBlocks = (turn?.blocks.filter((b) => b.type === 'actions') ??
        []) as ActionBlock[];
      if (!isWaveDependencySatisfied(block, siblingActionBlocks)) {
        console.warn(
          '[Cellix] Refused to accept a staged wave before its dependency was accepted:',
          { blockId, dependsOnChangeSetId: block.dependsOnChangeSetId },
        );
        return false;
      }

      applyingActionsRef.current = true;

      let outcomeWarning: string | null = null;
      // A holder rather than a bare `let`: the assignment happens inside the
      // onOutcomeVerified callback, which TypeScript's control-flow analysis
      // cannot see, so a plain `let` narrows to `never` at the read below.
      const outcome: { repair: RepairRequest | null } = { repair: null };
      try {
        if (onActions) {
          await onActions(block.actions, block.explanation, {
            changeSetId: block.changeSetId,
            changes: block.changes,
            // TASKS.md #150: the read-back's verdict comes back here so the UI
            // can say so. A clean verification passes `null` and stays silent.
            onOutcomeVerified: (_verification, message, repair) => {
              outcomeWarning = message;
              // #150 could only report. The read-back now also carries a
              // concrete next step, so a run that wrote a broken formula
              // offers the fix instead of leaving it in the workbook.
              outcome.repair = repair;
            },
          });
        }

        // Spec 22 Bug 3: only mark Applied after the apply path succeeds.
        updateTurn(turnId, (t) => ({
          ...t,
          error: undefined,
          blocks: t.blocks.map((b) =>
            b.id === blockId && b.type === 'actions'
              ? { ...b, proposalStatus: 'accepted' }
              : b,
          ),
        }));

        if (block.changeSetId) {
          onChangeSetApplied?.(block.changeSetId);
        }

        // An apply that succeeded but did not produce the proposed workbook is
        // NOT a clean success. Surface it on the turn rather than letting
        // "Applied" stand alone — the §3.7 rule this whole class of bug keeps
        // re-teaching: never let incomplete work look finished.
        if (outcomeWarning) {
          updateTurn(turnId, (t) => ({ ...t, error: outcomeWarning ?? undefined }));
        }

        // TASKS.md #168 — a repairable failure offers the repair. Exposed on
        // the turn rather than sent automatically: the write already landed in
        // the user's workbook, so the follow-up that rewrites those cells is
        // their call, the same consent rule Accept itself follows.
        if (outcome.repair) {
          const repair = outcome.repair;
          console.warn(
            `[Cellix] ${repair.cellCount} cell(s) returned ${repair.errors.join('/')} — repair available`,
          );
          updateTurn(turnId, (t) => ({ ...t, repairSuggestion: repair }));
        }

        setActiveClarification(null);

        // TASKS.md #153 — this card was one wave of a PAUSED run: nothing past
        // it has been generated yet, so accepting it is what triggers the next
        // wave. Awaited rather than fired-and-forgotten so `acceptActions`
        // resolving means "the build actually advanced", which is what
        // Accept All's sequential gate depends on.
        if (block.stepwise && block.runId) {
          await continueStepwiseRun(turnId, block.runId, 'accepted');
        }

        return true;
      } catch (error) {
        const rawMessage =
          error instanceof Error ? error.message : 'Failed to apply changes';
        const messageText = toUserFacingApplyError(rawMessage);
        console.error('[Cellix] Accept apply failed:', rawMessage);
        updateTurn(turnId, (t) => ({
          ...t,
          error: messageText,
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
    [continueStepwiseRun, onActions, onChangeSetApplied, updateTurn],
  );

  /**
   * Accept this step and every remaining pending step of the same staged build,
   * in order — TASKS.md #160.
   *
   * Sequential and fail-closed on purpose. `acceptActions` already refuses a
   * step whose dependency has not been applied (the #80 gate), and it returns
   * `false` for a refusal — so a step that will not apply STOPS the run rather
   * than letting later steps write into a workbook that never got its earlier
   * ones. "Accept All" skips the human gate, never the checking.
   */
  const acceptAllActions = useCallback(
    async (turnId: string, fromBlockId: string): Promise<boolean> => {
      const turn = getActiveSession()?.turns.find((t) => t.id === turnId);
      if (!turn) return false;

      const pending = turn.blocks.filter(
        (b): b is ActionBlock => b.type === 'actions' && b.proposalStatus === 'pending',
      );
      const startAt = pending.findIndex((b) => b.id === fromBlockId);
      if (startAt === -1) return false;

      for (const block of pending.slice(startAt)) {
        const ok = await acceptActions(turnId, block.id);
        if (!ok) {
          console.warn('[Cellix] Accept All stopped: a step did not apply', { blockId: block.id });
          return false;
        }
      }
      return true;
    },
    [acceptActions, getActiveSession],
  );


  const rejectActions = useCallback(
    async (turnId: string, blockId: string) => {
      await onClearPreview?.();

      // Captured before the state update, because the block is what carries the
      // run correlation and the update below rewrites the block list.
      const rejectedBlock = getActiveSession()
        ?.turns.find((t) => t.id === turnId)
        ?.blocks.find((b): b is ActionBlock => b.id === blockId && b.type === 'actions');

      updateTurn(turnId, (t) => {
        const siblingActionBlocks = t.blocks.filter((b) => b.type === 'actions') as ActionBlock[];
        // A pending wave that depends on the one being rejected (directly or
        // transitively) targets sheets/ranges that wave would have created — its
        // actions can no longer succeed, so leaving it "pending" would be a dead
        // end the user can never resolve. Reject it too.
        const cascadeIds = collectCascadeRejectIds(siblingActionBlocks, blockId);

        return {
          ...t,
          blocks: t.blocks.map((b) =>
            b.type === 'actions' && cascadeIds.has(b.id)
              ? { ...b, proposalStatus: 'rejected' as const }
              : b,
          ),
        };
      });

      // TASKS.md #153 — rejecting one wave of a step-wise run does NOT abort the
      // build (STEPWISE_EXECUTION.md SD-4): the server cascade-skips whatever
      // depended on this wave and carries on with the independent remainder.
      // Without this call the run would simply stall, since nothing else asks
      // the backend to generate the next wave.
      if (rejectedBlock?.stepwise && rejectedBlock.runId) {
        await continueStepwiseRun(turnId, rejectedBlock.runId, 'rejected');
      }
    },
    [continueStepwiseRun, getActiveSession, onClearPreview, updateTurn],
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
      if (sessionId === activeSessionIdRef.current) return;

      // Abort any in-flight stream before switching, exactly as newChat/closeSession/
      // openConversationFromHistory already do — updateTurn/syncConversationId key off
      // activeSessionIdRef.current, so a late chunk would otherwise land in (or clobber
      // the conversationId of) the tab we're leaving instead of the one that sent it.
      abortControllerRef.current?.abort();
      if (activeTurnId) {
        const runtime = runtimeRef.current.get(activeTurnId);
        if (runtime) {
          runtime.aborted = true;
          runtimeRef.current.delete(activeTurnId);
        }
        revealScheduledRef.current.delete(activeTurnId);
      }
      void onClearPreview?.();
      setIsWaitingForResponse(false);

      activeSessionIdRef.current = sessionId;
      setActiveSessionId(sessionId);
      applySessionContext(session);
      const lastTurn = session.turns[session.turns.length - 1];
      setActiveTurnId(lastTurn?.id ?? null);
      setActiveClarification(null);
      schedulePersist();
    },
    [activeTurnId, applySessionContext, onClearPreview, schedulePersist],
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

  /**
   * Rename a session (TASKS.md #177) — updates the local tab immediately and,
   * when the session has a server `conversationId`, persists the title there
   * too so it survives a reload/history-panel view.
   *
   * The local update is optimistic and unconditional: a session with no
   * `conversationId` yet (nothing sent) is local-only and has nothing to sync,
   * and a server failure on an existing conversation is logged rather than
   * rolled back — losing a rename on a flaky connection is a much smaller
   * problem than losing it silently with no feedback at all, and the next
   * successful rename or reload will reconcile the two anyway.
   */
  const renameSession = useCallback(
    (sessionId: string, title: string): void => {
      const trimmed = title.trim();
      if (!trimmed) return;

      const target = sessionsRef.current.find((session) => session.id === sessionId);
      updateSession(sessionId, (session) => ({
        ...session,
        title: trimmed,
        updatedAt: new Date().toISOString(),
      }));

      if (target?.conversationId) {
        void renameConversationOnServer(target.conversationId, trimmed).catch((error) => {
          console.warn('[Cellix] Failed to persist chat rename:', error);
        });
      }
    },
    [updateSession],
  );

  /**
   * Delete a session (TASKS.md #177) — closes the local tab (reusing
   * `closeSession`'s active-session fallback logic exactly, so deleting the
   * active chat behaves the same as closing it) and, when it has a server
   * `conversationId`, deletes it there too. This is a hard delete with no
   * undo, matching #177's scope decision.
   *
   * Resolves after the server delete settles (or is skipped) so a caller can
   * show a spinner and know when it's safe to assume the row is gone — but
   * the local tab closes immediately regardless of server outcome, since the
   * user's "delete" intent applies to what they can see whether or not the
   * network cooperates.
   */
  const deleteSession = useCallback(
    async (sessionId: string): Promise<void> => {
      const target = sessionsRef.current.find((session) => session.id === sessionId);
      closeSession(sessionId);

      if (target?.conversationId) {
        try {
          await deleteConversationOnServer(target.conversationId);
        } catch (error) {
          console.warn('[Cellix] Failed to delete chat on the server:', error);
          throw error;
        }
      }
    },
    [closeSession],
  );

  /**
   * Delete a conversation that is server-side history but not (or no longer)
   * an open local tab (TASKS.md #177) — the common case when deleting from
   * the history menu rather than from an open tab. If it *is* also open,
   * closes that tab too so the two views can't disagree.
   */
  const deleteHistoryConversation = useCallback(
    async (conversationId: string): Promise<void> => {
      const openSession = sessionsRef.current.find(
        (session) => session.conversationId === conversationId,
      );
      if (openSession) {
        await deleteSession(openSession.id);
        return;
      }
      await deleteConversationOnServer(conversationId);
    },
    [deleteSession],
  );

  /**
   * Open a conversation from server-backed history (TASKS.md #172).
   *
   * If that conversation is already an open tab — the common case right after
   * sending a message — this just selects it rather than fetching and appending
   * a duplicate tab for the same thread.
   *
   * Otherwise it fetches the full body and rehydrates it through the *existing*
   * `messagesToTurns` path that reload-restore already uses, rather than adding
   * a second message-loading mechanism that could drift from it.
   */
  const openConversationFromHistory = useCallback(
    async (conversationId: string): Promise<boolean> => {
      const alreadyOpen = sessionsRef.current.find(
        (session) => session.conversationId === conversationId,
      );
      if (alreadyOpen) {
        selectSession(alreadyOpen.id);
        return true;
      }

      setIsLoadingHistoryConversation(true);
      try {
        const stored = await fetchConversationById(conversationId);
        const messages = stored.messages ?? [];
        const restored: ChatSession = {
          ...createChatSession(truncateTabLabel(stored.title || 'Chat', 24)),
          conversationId: stored.conversationId ?? conversationId,
          turns: messagesToTurns(messages),
          history: messagesToHistory(messages),
          updatedAt: stored.updatedAt ?? new Date().toISOString(),
        };

        // Abort anything streaming into the tab we're leaving, exactly as
        // newChat does — otherwise a late chunk lands in the restored thread.
        abortControllerRef.current?.abort();
        void onClearPreview?.();

        activeSessionIdRef.current = restored.id;
        setActiveSessionId(restored.id);
        setActiveClarification(null);
        setIsWaitingForResponse(false);
        applySessionContext(restored);
        setActiveTurnId(restored.turns[restored.turns.length - 1]?.id ?? null);
        syncSessions([...sessionsRef.current, restored]);
        return true;
      } catch (error) {
        console.warn('[Cellix] Failed to open conversation from history:', error);
        return false;
      } finally {
        setIsLoadingHistoryConversation(false);
      }
    },
    [applySessionContext, onClearPreview, selectSession, syncSessions],
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
    acceptAllActions,
    rejectActions,
    endConversation,
    newChat,
    clearConversation,
    selectSession,
    closeSession,
    renameSession,
    deleteSession,
    deleteHistoryConversation,
    openConversationFromHistory,
    isLoadingHistoryConversation,
    selectTurn,
    closeTurn,
    toggleThinking,
    markAnswerComplete,
  };
};
