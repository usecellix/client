/**
 * End-to-end orchestrator for chat-triggered GST recon (client-side).
 * Casual "Reconcile purchase/sales register" skips GSTIN prompts and offers
 * a missed-rows sheet. Full recon still collects GSTIN/period/client first.
 */

import {
  detectGstReconIntent,
  detectGstReconLayout,
  detectPurchasePortalSourcePreference,
  extractGstReconContext,
  GstPurchasePortalPreference,
  GstReconIntentType,
  GstReconLayout,
  isCasualMissedRowsRecon,
} from '@/utils/gstReconIntent';
import { resolveGstPeriod } from '@/utils/gstPeriod';
import {
  buildAmbiguousSheetMessage,
  buildMissingSheetMessage,
  discoverGstSheets,
  inferColumnMapping,
} from '@/services/gstSheetDiscovery';
import { readAllSheetHeaders, readSheetsFull, sheetExists } from '@/services/gstSheetReader';
import {
  buildGstReconAnswerText,
  buildGstReconUserFacingSummary,
  GstReconcileResponse,
  GstReconType,
  reportGstReconAuditOutcome,
  runGstReconcile,
} from '@/services/gstReconService';
import { loadChatSessions, saveChatSessions } from '@/utils/chatSessionStorage';
import { SheetAction } from '@/types/sheet-actions';
import { UserFacingSummary } from '@/utils/userFacingResponse';
import { GstReconMissedRow } from '@/types/conversationTurn';

/** Every books row (PR_ONLY) from the result, ready to render as a clickable jump-to-cell list. */
function extractMissedRows(result: GstReconcileResponse): GstReconMissedRow[] {
  return result.rows
    .filter((r) => r.status === 'PR_ONLY')
    .map((r) => ({
      vendorName: r.vendor_name || r.invoice_number || 'Unnamed row',
      gstin: r.gstin || '',
      reason: r.mismatch_reason || r.status,
      explanation: r.explanation || r.difference || '',
      sheetName: r.books_sheet_name ?? null,
      row: r.books_row ?? null,
    }));
}

export type GstReconChatOutcome =
  | {
      kind: 'message_only';
      answer: string;
    }
  | {
      kind: 'recon_ready';
      answer: string;
      actions: SheetAction[];
      explanation: string;
      userFacingSummary: UserFacingSummary;
      auditLogId?: string | null;
      /** The sheet these actions write to — the resolved name after Overwrite/Create-new, when applicable. */
      sheetName: string;
      /** Books rows that didn't cleanly match the portal — rendered as a clickable list so
       * each one can be jumped to directly in the books sheet, never the portal sheet. */
      missedRows?: GstReconMissedRow[];
    }
  | {
      /**
       * The reconciliation finished but its output sheet name collides with one
       * already in the workbook — a button choice (Overwrite / Create new), not
       * a typed reply, resolves it via `resolveGstReconSheetCollision`.
       */
      kind: 'sheet_collision';
      answer: string;
      collisionId: string;
      sheetName: string;
    }
  | {
      /**
       * A clarifying question, docked above the composer via the app's standard
       * `QuestionChoicesPanel` mechanism (same as every other mid-conversation
       * clarification) — answering it re-sends a wrapped reply through ordinary chat,
       * which `tryHandleGstReconChat` recognizes as a continuation via `pendingContext`-
       * style state (see `pendingPurchasePortalAsk`) and resolves without re-asking.
       */
      kind: 'question';
      answer: string;
      question: string;
      options: string[];
    };

/** Session-scoped partial context for multi-turn prompts. */
export interface GstReconPendingContext {
  intentType: GstReconIntentType;
  clientGstin?: string;
  period?: string;
  financialYear?: string;
  clientName?: string;
  awaiting: Array<'clientGstin' | 'period' | 'clientName'>;
}

let pendingContext: GstReconPendingContext | null = null;

/** Last recon audit id awaiting Accept/Reject (single active card). */
let pendingAuditLogId: string | null = null;

/** A completed reconciliation whose output sheet name collides with an existing sheet. */
interface GstReconSheetCollision {
  collisionId: string;
  result: GstReconcileResponse;
  prName: string;
  portalName: string;
  layout: GstReconLayout;
  targetSheetName: string;
}

let pendingSheetCollision: GstReconSheetCollision | null = null;
let sheetCollisionCounter = 0;

/**
 * Held while asking "GSTR-2B only, GSTR-2A only, or combined?" — remembers the ORIGINAL
 * triggering message (for its period/layout/etc. wording) so the next message can resume
 * the same run once answered. Single-slot, matching `pendingContext`'s own design (one
 * outstanding ask at a time is enough) — resolved from the NEXT `tryHandleGstReconChat`
 * call, same continuation shape as GSTIN/period/client prompts.
 */
interface PendingPurchasePortalAsk {
  originalMessage: string;
  workbookKey?: string;
}

let pendingPurchasePortalAsk: PendingPurchasePortalAsk | null = null;

const PURCHASE_PORTAL_ASK_QUESTION =
  'This workbook has both GSTR-2B and GSTR-2A sheets. Which should I reconcile against?';

const PURCHASE_PORTAL_ASK_OPTION_LABELS: Record<GstPurchasePortalPreference, string> = {
  gstr2b_only: 'GSTR-2B only',
  gstr2a_only: 'GSTR-2A only',
  combined: 'Both, combined',
};

const PURCHASE_PORTAL_ASK_OPTIONS: string[] = [
  PURCHASE_PORTAL_ASK_OPTION_LABELS.gstr2b_only,
  PURCHASE_PORTAL_ASK_OPTION_LABELS.gstr2a_only,
  PURCHASE_PORTAL_ASK_OPTION_LABELS.combined,
];

function persistPurchasePortalPreference(
  workbookKey: string | undefined,
  preference: GstPurchasePortalPreference,
): void {
  if (!workbookKey) return;
  const existing = loadChatSessions(workbookKey) ?? {
    activeSessionId: null,
    sessions: [],
  };
  saveChatSessions(workbookKey, { ...existing, gstPurchasePortalPreference: preference });
}

/**
 * The standard docked-question flow (`answerQuestion` in `useConversation.ts`) never sends
 * the raw answer alone — it wraps it as `Replying to your question "<Q>" about my earlier
 * request "<original>": <answer>` so the reply carries enough context to be recognized as a
 * GST-recon continuation (see `isGstReconPrompt`/`isCasualMissedRowsRecon`, both of which
 * match against the embedded original request text). But that same embedding means the
 * QUESTION text itself — which names both "GSTR-2B" and "GSTR-2A" — is present in the
 * payload too, so running `detectPurchasePortalSourcePreference` against the whole string
 * would see both mentioned and (correctly, by its own "don't guess" rule) return null even
 * when the actual answer was unambiguous. Extracting just the answer segment (after the
 * final ": ") avoids that false ambiguity.
 */
function extractPurchasePortalAnswer(message: string): GstPurchasePortalPreference | null {
  const wrapperMatch = message.match(/:\s*([^:]+)$/);
  const answerText = (wrapperMatch ? wrapperMatch[1] : message).trim();

  const byLabel = (Object.entries(PURCHASE_PORTAL_ASK_OPTION_LABELS) as Array<
    [GstPurchasePortalPreference, string]
  >).find(([, label]) => label.toLowerCase() === answerText.toLowerCase());
  if (byLabel) return byLabel[0];

  return detectPurchasePortalSourcePreference(answerText);
}

/**
 * A completed action batch whose target sheet name turned out to already exist at
 * write (Accept) time, even though it did not when the batch was built — e.g. the
 * name was free during chat-generation's collision check but got occupied before
 * the user clicked Accept. Held so Overwrite/Create-new can retry directly against
 * the already-computed actions, without re-running the reconciliation. Keyed
 * separately from `pendingSheetCollision` (a different prefix on the id) since it
 * carries no `result`/`prName`/`portalName` to rebuild a full chat summary from.
 */
interface GstReconWriteFailureCollision {
  collisionId: string;
  actions: SheetAction[];
  targetSheetName: string;
}

const pendingWriteFailureCollisions = new Map<string, GstReconWriteFailureCollision>();

const WRITE_FAILURE_COLLISION_PREFIX = 'write_collision_';

/**
 * Registers a write-time collision detected right before applying a GST-recon
 * action batch (the target sheet name now exists, though it did not when the
 * Accept card was built) and returns the id/name for a fresh collision card.
 * Never applies or modifies anything itself — purely bookkeeping.
 */
export function registerGstReconWriteFailureCollision(
  actions: SheetAction[],
  targetSheetName: string,
): { collisionId: string; sheetName: string } {
  const collisionId = `${WRITE_FAILURE_COLLISION_PREFIX}${++sheetCollisionCounter}`;
  pendingWriteFailureCollisions.set(collisionId, { collisionId, actions, targetSheetName });
  return { collisionId, sheetName: targetSheetName };
}

async function resolveGstReconWriteFailureCollision(
  collisionId: string,
  choice: 'overwrite' | 'new',
): Promise<GstReconChatOutcome> {
  const collision = pendingWriteFailureCollisions.get(collisionId);
  if (!collision) {
    return {
      kind: 'message_only',
      answer: 'That choice is no longer available — please run the reconciliation again.',
    };
  }
  pendingWriteFailureCollisions.delete(collisionId);
  const { actions, targetSheetName } = collision;

  if (choice === 'overwrite') {
    return {
      kind: 'recon_ready',
      answer: `Sheet "${targetSheetName}" already existed — overwriting it.`,
      actions: [{ type: 'DELETE_SHEET', sheetName: targetSheetName }, ...actions],
      explanation: 'Retry after a name conflict detected at write time.',
      userFacingSummary: { headline: 'Retry after a write-time name conflict.' },
      sheetName: targetSheetName,
    };
  }

  const headerIndex = await readAllSheetHeaders();
  const taken = new Set(headerIndex.map((h) => h.name));
  const finalSheetName = versionedSheetName(targetSheetName, taken);
  return {
    kind: 'recon_ready',
    answer: `Sheet "${targetSheetName}" already existed — creating "${finalSheetName}" instead.`,
    actions: rewriteActionSheetNames(actions, targetSheetName, finalSheetName),
    explanation: 'Retry after a name conflict detected at write time.',
    userFacingSummary: { headline: 'Retry after a write-time name conflict.' },
    sheetName: finalSheetName,
  };
}

export function getGstReconPendingContext(): GstReconPendingContext | null {
  return pendingContext;
}

export function clearGstReconPendingContext(): void {
  pendingContext = null;
}

export function getPendingGstReconAuditLogId(): string | null {
  return pendingAuditLogId;
}

export async function finalizeGstReconAudit(
  outcome: 'applied' | 'rejected',
): Promise<void> {
  const id = pendingAuditLogId;
  pendingAuditLogId = null;
  if (id) await reportGstReconAuditOutcome(id, outcome);
}

function intentToApiType(type: GstReconIntentType): GstReconType {
  if (type === 'SALES_VS_GSTR1') return 'SALES_VS_GSTR1';
  if (type === 'PR_VS_GSTR2A') return 'PR_VS_GSTR2A';
  if (type === 'IMS_VS_PR') return 'IMS_VS_PR';
  if (type === 'GSTR3B_VS_GSTR2B') return 'GSTR3B_VS_GSTR2B';
  return 'PR_VS_GSTR2B';
}

function preferredPortal(
  type: GstReconIntentType,
): 'GSTR2B' | 'GSTR2A' | 'GSTR1' | 'IMS' {
  if (type === 'PR_VS_GSTR2A') return 'GSTR2A';
  if (type === 'SALES_VS_GSTR1') return 'GSTR1';
  if (type === 'IMS_VS_PR') return 'IMS';
  return 'GSTR2B';
}

function mergeContext(
  intentType: GstReconIntentType,
  fromMessage: ReturnType<typeof extractGstReconContext>,
): GstReconPendingContext {
  const base = pendingContext?.intentType === intentType ? pendingContext : null;
  const next: GstReconPendingContext = {
    intentType,
    clientGstin: fromMessage.extractedGstin ?? base?.clientGstin,
    period: fromMessage.extractedPeriod ?? base?.period,
    financialYear: fromMessage.extractedFinancialYear ?? base?.financialYear,
    clientName: fromMessage.extractedClientName ?? base?.clientName,
    awaiting: [],
  };
  if (!next.clientGstin) next.awaiting.push('clientGstin');
  if (!next.period) next.awaiting.push('period');
  if (!next.clientName) next.awaiting.push('clientName');
  return next;
}

function missingContextPrompts(ctx: GstReconPendingContext): string[] {
  const prompts: string[] = [];
  if (ctx.awaiting.includes('clientGstin')) {
    prompts.push(
      `Which client GSTIN should I reconcile against? (This should be the client's own GSTIN — not your CA firm's.)`,
    );
  }
  if (ctx.awaiting.includes('period')) {
    prompts.push(`Which period should I reconcile — a specific month, or a quarter?`);
  }
  if (ctx.awaiting.includes('clientName')) {
    prompts.push(`Which client is this for?`);
  }
  return prompts;
}

/** Reassign every action's sheetName (and CREATE_SHEET's name) from `from` to `to`. */
function rewriteActionSheetNames(actions: SheetAction[], from: string, to: string): SheetAction[] {
  if (from === to) return actions;
  return actions.map((action) => {
    if (action.sheetName !== from && action.name !== from) return action;
    return {
      ...action,
      ...(action.sheetName === from ? { sheetName: to } : {}),
      ...(action.name === from ? { name: to } : {}),
    };
  });
}

/** First "<base> (N)" (Excel's 31-char sheet-name limit respected) not already taken. */
function versionedSheetName(base: string, taken: Set<string>): string {
  if (!taken.has(base)) return base;
  for (let n = 2; n < 1000; n++) {
    const suffix = ` (${n})`;
    const candidate = `${base.slice(0, Math.max(31 - suffix.length, 1))}${suffix}`;
    if (!taken.has(candidate)) return candidate;
  }
  return base;
}

/**
 * Resolve a pending sheet-name collision from a button click — `collisionId` must match
 * the currently held collision (guards against a stale/duplicate click on an old card).
 * Never parses typed text; the choice comes directly from which button was clicked.
 * Dispatches to the write-time (Accept-time) collision store when `collisionId`
 * belongs to it, so callers never need to know which flow originated the card.
 */
export async function resolveGstReconSheetCollision(
  collisionId: string,
  choice: 'overwrite' | 'new',
): Promise<GstReconChatOutcome> {
  if (collisionId.startsWith(WRITE_FAILURE_COLLISION_PREFIX)) {
    return resolveGstReconWriteFailureCollision(collisionId, choice);
  }

  if (!pendingSheetCollision || pendingSheetCollision.collisionId !== collisionId) {
    return {
      kind: 'message_only',
      answer: 'That choice is no longer available — please run the reconciliation again.',
    };
  }

  const collision = pendingSheetCollision;
  pendingSheetCollision = null;
  const { result, prName, portalName, layout, targetSheetName } = collision;
  let actions = result.actions ?? [];
  let effectiveResult = result;

  if (choice === 'overwrite') {
    // Delete + recreate rather than clearing the existing sheet's contents in
    // place: an in-place clear of row 0 trips the header guard (it protects
    // row 0 as "headers", which is exactly where our own title text lives),
    // and even if it didn't, OverwriteGuard would still block the WRITE_TABLE
    // that follows unless explicitly confirmed. A freshly (re)created sheet is
    // empty, so CREATE_SHEET/WRITE_TABLE/styling run through unmodified,
    // exactly as the no-collision path already does — no guard needs bypassing.
    actions = [{ type: 'DELETE_SHEET', sheetName: targetSheetName }, ...actions];
  } else {
    const headerIndex = await readAllSheetHeaders();
    const taken = new Set(headerIndex.map((h) => h.name));
    const finalSheetName = versionedSheetName(targetSheetName, taken);
    actions = rewriteActionSheetNames(actions, targetSheetName, finalSheetName);
    effectiveResult = { ...result, output_sheet_name: finalSheetName };
  }

  const userFacingSummary = buildGstReconUserFacingSummary(effectiveResult, prName, portalName, layout);
  const answer = buildGstReconAnswerText(effectiveResult, prName, portalName, layout);
  pendingAuditLogId = effectiveResult.audit_log_id ?? null;

  return {
    kind: 'recon_ready',
    answer,
    actions,
    explanation: userFacingSummary.headline,
    userFacingSummary,
    auditLogId: effectiveResult.audit_log_id,
    sheetName: effectiveResult.output_sheet_name,
    missedRows: extractMissedRows(effectiveResult),
  };
}

/**
 * Returns null when the message is not a GST recon prompt (and no pending context).
 */
export async function tryHandleGstReconChat(
  message: string,
  options?: {
    conversationId?: string | null;
    workbookKey?: string;
    /**
     * Internal only — set when resuming after the "GSTR-2B / GSTR-2A / combined?" question
     * is answered, so the resumed run uses the just-chosen preference directly instead of
     * re-deriving it from persisted storage. Persisting-then-reloading was fragile: if
     * `workbookKey` were ever unavailable (or the storage round-trip lagged), the resumed
     * call would find no stored preference and re-ask the identical question — the exact
     * "asks twice" bug reported after this ask shipped. Passing the answer straight through
     * removes storage from the resume path's correctness entirely; persistence still
     * happens, but purely as a side effect for FUTURE runs, never as what THIS run depends on.
     */
    forcedPurchasePortalPreference?: GstPurchasePortalPreference;
  },
): Promise<GstReconChatOutcome | null> {
  if (pendingPurchasePortalAsk) {
    const ask = pendingPurchasePortalAsk;
    const preference = extractPurchasePortalAnswer(message);
    if (preference) {
      pendingPurchasePortalAsk = null;
      persistPurchasePortalPreference(ask.workbookKey, preference);
      return tryHandleGstReconChat(ask.originalMessage, {
        ...options,
        workbookKey: ask.workbookKey,
        forcedPurchasePortalPreference: preference,
      });
    }
    // Doesn't parse as an answer to the pending ask — fall through to ordinary intent
    // handling rather than getting stuck re-asking forever if state ever desyncs from
    // what's on screen (e.g. the user typed something unrelated instead of using the
    // docked question card).
  }

  let intent = detectGstReconIntent(message);

  // Continue multi-turn context collection
  if (!intent && pendingContext) {
    const followUp = extractGstReconContext(message);
    // Also accept bare GSTIN / short period / client name replies
    const gstinOnly = message.match(
      /\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])\b/i,
    );
    if (gstinOnly && !followUp.extractedGstin) {
      followUp.extractedGstin = gstinOnly[1].toUpperCase();
    }
    if (!followUp.extractedPeriod && pendingContext.awaiting.includes('period')) {
      const trimmed = message.trim();
      if (trimmed.length > 2 && trimmed.length < 40 && !gstinOnly) {
        followUp.extractedPeriod = trimmed;
      }
    }
    if (!followUp.extractedClientName && pendingContext.awaiting.includes('clientName')) {
      const trimmed = message.trim();
      if (trimmed.length > 1 && trimmed.length < 80 && !gstinOnly) {
        followUp.extractedClientName = trimmed;
      }
    }
    intent = {
      matched: true,
      type: pendingContext.intentType,
      ...followUp,
      extractedGstin: followUp.extractedGstin ?? pendingContext.clientGstin,
      extractedPeriod: followUp.extractedPeriod ?? pendingContext.period,
      extractedFinancialYear:
        followUp.extractedFinancialYear ?? pendingContext.financialYear,
      extractedClientName: followUp.extractedClientName ?? pendingContext.clientName,
    };
  }

  if (!intent) return null;

  if (intent.type === 'GSTR3B_VS_GSTR2B') {
    pendingContext = null;
    return {
      kind: 'message_only',
      answer:
        'GSTR-3B vs GSTR-2B is a **summary** comparison (ITC claimed vs available). ' +
        'For now, use chat purchase/2B recon for invoice matching. ' +
        'To compare 3B totals, provide IGST/CGST/SGST figures from both returns in a follow-up, or paste both as sheets and ask for invoice-level PR vs 2B recon.',
    };
  }

  const casual = isCasualMissedRowsRecon(message, intent);
  let ctx: GstReconPendingContext | null = null;
  if (!casual) {
    ctx = mergeContext(intent.type, {
      extractedGstin: intent.extractedGstin,
      extractedPeriod: intent.extractedPeriod,
      extractedFinancialYear: intent.extractedFinancialYear,
      extractedClientName: intent.extractedClientName,
    });
    pendingContext = ctx;
    if (ctx.awaiting.length) {
      const prompts = missingContextPrompts(ctx);
      return {
        kind: 'message_only',
        answer:
          `Before I run the reconciliation, I need a few details:\n\n` +
          prompts.map((p, i) => `${i + 1}. ${p}`).join('\n'),
      };
    }
  } else {
    pendingContext = null;
  }

  // Extract a period from the raw prompt if present (any confidently-recognized phrasing —
  // "April 2024", "Apr-24", "04/2024", "Q1 FY25", "last month", etc.). Falls back to the
  // conversationally-collected period string for full-recon multi-turn context. Ambiguous or
  // unrecognized phrasing resolves to null, which means no date filtering — never guess.
  const resolvedPeriod =
    resolveGstPeriod(message) ?? (ctx?.period ? resolveGstPeriod(ctx.period) : null);

  if (typeof Excel === 'undefined') {
    return {
      kind: 'message_only',
      answer:
        'GST reconciliation needs an open Excel workbook. Open your file with Purchase Register and GSTR sheets, then ask again.',
    };
  }

  const headerIndex = await readAllSheetHeaders();
  if (!headerIndex.length) {
    return {
      kind: 'message_only',
      answer: 'This workbook has no sheets I can read. Add your purchase and GSTR data, then try again.',
    };
  }

  const portalPref = preferredPortal(intent.type);
  const discovery = discoverGstSheets(headerIndex, portalPref);

  if (discovery.ambiguous.length) {
    return {
      kind: 'message_only',
      answer: buildAmbiguousSheetMessage(discovery.ambiguous),
    };
  }

  if (discovery.missing.length || !discovery.purchaseRegister || !discovery.portal) {
    const label =
      intent.type === 'PR_VS_GSTR2A'
        ? 'Purchase Register vs GSTR-2A'
        : intent.type === 'SALES_VS_GSTR1'
          ? 'Sales Register vs GSTR-1'
          : intent.type === 'IMS_VS_PR'
            ? 'IMS vs Purchase Register'
            : 'Purchase Register vs GSTR-2B';
    return {
      kind: 'message_only',
      answer: buildMissingSheetMessage(discovery.missing, label),
    };
  }

  const prName = discovery.purchaseRegister.name;
  const portalName = discovery.portal.name;
  const isSales = intent.type === 'SALES_VS_GSTR1';
  const eligibleForPortalChoice = casual && !isSales && intent.type !== 'IMS_VS_PR';
  const bothPortalsPresent = Boolean(
    discovery.gstr2b && discovery.gstr2a && discovery.gstr2b.name !== discovery.gstr2a.name,
  );

  let purchasePortalPreference: GstPurchasePortalPreference | null = null;
  if (eligibleForPortalChoice && bothPortalsPresent) {
    const explicitPreference = detectPurchasePortalSourcePreference(message);
    purchasePortalPreference =
      options?.forcedPurchasePortalPreference ??
      explicitPreference ??
      (options?.workbookKey
        ? loadChatSessions(options.workbookKey)?.gstPurchasePortalPreference ?? null
        : null);

    if (!purchasePortalPreference) {
      pendingPurchasePortalAsk = { originalMessage: message, workbookKey: options?.workbookKey };
      pendingContext = null;
      return {
        kind: 'question',
        answer: PURCHASE_PORTAL_ASK_QUESTION,
        question: PURCHASE_PORTAL_ASK_QUESTION,
        options: PURCHASE_PORTAL_ASK_OPTIONS,
      };
    }

    if (explicitPreference) {
      persistPurchasePortalPreference(options?.workbookKey, explicitPreference);
    }
  }

  const useDualPurchasePortals =
    eligibleForPortalChoice &&
    (!bothPortalsPresent || purchasePortalPreference === 'combined');

  // A deliberate single-source preference can name the OTHER portal than the one
  // `preferredPortal(intent.type)` defaulted to (e.g. intent defaults to GSTR-2B but the
  // user's remembered preference is 2A-only) — resolve the effective single-source sheet
  // from the preference itself rather than trusting `portalName` in that case.
  const effectivePortalName =
    eligibleForPortalChoice && bothPortalsPresent && !useDualPurchasePortals
      ? purchasePortalPreference === 'gstr2a_only'
        ? (discovery.gstr2a?.name ?? portalName)
        : (discovery.gstr2b?.name ?? portalName)
      : portalName;

  const toRead = [prName];
  if (useDualPurchasePortals) {
    if (discovery.gstr2b) toRead.push(discovery.gstr2b.name);
    if (discovery.gstr2a && discovery.gstr2a.name !== discovery.gstr2b?.name) {
      toRead.push(discovery.gstr2a.name);
    }
    if (!discovery.gstr2b && !discovery.gstr2a) toRead.push(effectivePortalName);
  } else {
    toRead.push(effectivePortalName);
    if (discovery.ims && discovery.ims.name !== effectivePortalName) {
      toRead.push(discovery.ims.name);
    }
  }

  const grids = await readSheetsFull([...new Set(toRead)]);
  const byName = new Map(grids.map((g) => [g.name, g]));
  const prGrid = byName.get(prName);

  if (!prGrid?.values?.length || prGrid.values.length < 2) {
    return {
      kind: 'message_only',
      answer: `Sheet "${prName}" looks empty. Add invoice rows with headers, then ask again.`,
    };
  }

  const gridOk = (name: string) => {
    const g = byName.get(name);
    return Boolean(g?.values && g.values.length >= 2);
  };

  if (useDualPurchasePortals) {
    const twoBOk = Boolean(discovery.gstr2b && gridOk(discovery.gstr2b.name));
    const twoAOk = Boolean(discovery.gstr2a && gridOk(discovery.gstr2a.name));
    if (!twoBOk && !twoAOk) {
      return {
        kind: 'message_only',
        answer:
          'GSTR-2B / GSTR-2A sheets look empty. Paste the portal download(s) into this workbook, then ask again.',
      };
    }
  } else {
    const portalGrid = byName.get(effectivePortalName);
    if (!portalGrid?.values?.length || portalGrid.values.length < 2) {
      return {
        kind: 'message_only',
        answer: `Sheet "${effectivePortalName}" looks empty. Paste the GSTR portal download into that sheet, then ask again.`,
      };
    }
  }

  const portalFileType =
    eligibleForPortalChoice && bothPortalsPresent && !useDualPurchasePortals
      ? purchasePortalPreference === 'gstr2a_only'
        ? 'GSTR2A'
        : 'GSTR2B'
      : intent.type === 'PR_VS_GSTR2A'
        ? 'GSTR2A'
        : intent.type === 'IMS_VS_PR'
          ? 'IMS'
          : isSales
            ? 'GSTR1'
            : 'GSTR2B';

  const sheetPayload = (
    name: string,
    fileType: string,
  ): { sheet_name: string; data: unknown[][]; column_mapping: ReturnType<typeof inferColumnMapping>; file_type: string; headers_row: number } | null => {
    const g = byName.get(name);
    if (!g?.values?.length) return null;
    return {
      sheet_name: name,
      data: g.values,
      column_mapping: inferColumnMapping(g.headers),
      file_type: fileType,
      headers_row: 1,
    };
  };

  let portal_file = sheetPayload(effectivePortalName, portalFileType);
  let portal_file_2a: ReturnType<typeof sheetPayload> = null;
  if (useDualPurchasePortals) {
    const twoB = discovery.gstr2b ? sheetPayload(discovery.gstr2b.name, 'GSTR2B') : null;
    const twoA = discovery.gstr2a ? sheetPayload(discovery.gstr2a.name, 'GSTR2A') : null;
    if (twoB && twoA && twoB.sheet_name !== twoA.sheet_name) {
      portal_file = twoB;
      portal_file_2a = twoA;
    } else if (twoB) {
      portal_file = twoB;
    } else if (twoA) {
      portal_file = twoA;
    }
  }

  if (!portal_file) {
    return {
      kind: 'message_only',
      answer: 'I could not read a GSTR portal sheet. Add GSTR-2B, GSTR-2A, or GSTR-1 and try again.',
    };
  }

  const layout = casual ? detectGstReconLayout(message) : 'categorized';

  try {
    const result = await runGstReconcile({
      reconciliation_type: intentToApiType(intent.type),
      client_gstin: ctx?.clientGstin,
      period: ctx?.period,
      financial_year: ctx?.financialYear,
      client_name: ctx?.clientName,
      missed_books_only: casual,
      layout: casual ? layout : undefined,
      period_start: resolvedPeriod?.start,
      period_end: resolvedPeriod?.end,
      period_label: resolvedPeriod?.label,
      purchase_register: {
        sheet_name: prName,
        data: prGrid.values,
        column_mapping: inferColumnMapping(prGrid.headers),
        headers_row: 1,
      },
      books_register: {
        sheet_name: prName,
        data: prGrid.values,
        column_mapping: inferColumnMapping(prGrid.headers),
        headers_row: 1,
      },
      portal_file,
      portal_file_2a: portal_file_2a ?? undefined,
      ims_data:
        !casual &&
        !isSales &&
        discovery.ims &&
        byName.get(discovery.ims.name)?.values?.length
          ? {
              sheet_name: discovery.ims.name,
              data: byName.get(discovery.ims.name)!.values,
              column_mapping: inferColumnMapping(byName.get(discovery.ims.name)!.headers),
              headers_row: 1,
            }
          : null,
      settings: {
        amount_tolerance_abs: 1,
        amount_tolerance_pct: 0.5,
        invoice_fuzzy_threshold: 85,
        date_tolerance_days: 3,
        detect_rcm: !casual && !isSales,
        use_ims_data: !casual && Boolean(!isSales && discovery.ims),
      },
      conversation_id: options?.conversationId ?? undefined,
    });

    pendingContext = null;

    if (result.period_zero_message) {
      pendingAuditLogId = null;
      return { kind: 'message_only', answer: result.period_zero_message };
    }

    const missed = result.summary.pr_only;
    const gstinMismatchCount = result.summary.gstin_mismatch_count ?? 0;
    // portal_flat's whole sheet IS the portal-only rows — a books-side-only issue count
    // would wrongly report "no missed rows" when portal-only rows are exactly what exist.
    const hasIssues =
      layout === 'portal_flat'
        ? (result.summary.portal_only ?? 0) > 0
        : missed > 0 || gstinMismatchCount > 0;
    const userFacingSummary = buildGstReconUserFacingSummary(result, prName, effectivePortalName, layout);
    const answer = buildGstReconAnswerText(result, prName, effectivePortalName, layout);

    if (casual && !hasIssues) {
      pendingAuditLogId = null;
      return { kind: 'message_only', answer };
    }

    // `headerIndex` was captured earlier in this function and could in principle be
    // stale or incomplete by now — never trust it alone for a decision this
    // consequential. `headerIndex.some(...)` short-circuits the common case (no
    // extra round trip needed when it already found the name); the fresh, minimal,
    // single-sheet `sheetExists` check only runs when headerIndex said "not found",
    // so a false negative there can never cause an undetected collision.
    const collisionDetected =
      headerIndex.some((h) => h.name === result.output_sheet_name) ||
      (await sheetExists(result.output_sheet_name));
    if (collisionDetected) {
      const collisionId = `collision_${++sheetCollisionCounter}`;
      pendingSheetCollision = {
        collisionId,
        result,
        prName,
        portalName: effectivePortalName,
        layout,
        targetSheetName: result.output_sheet_name,
      };
      pendingAuditLogId = null;
      return {
        kind: 'sheet_collision',
        answer,
        collisionId,
        sheetName: result.output_sheet_name,
      };
    }

    pendingAuditLogId = result.audit_log_id ?? null;
    return {
      kind: 'recon_ready',
      answer,
      actions: result.actions ?? [],
      explanation: userFacingSummary.headline,
      userFacingSummary,
      auditLogId: result.audit_log_id,
      sheetName: result.output_sheet_name,
      missedRows: extractMissedRows(result),
    };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    return {
      kind: 'message_only',
      answer: `I couldn't complete the reconciliation:\n\n${msg}`,
    };
  }
}
