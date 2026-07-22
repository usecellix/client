import { getApiBaseUrl } from '@/lib/apiConfig';
import { SheetAction } from '@/types/sheet-actions';

export type FrontendTelemetryLevel = 'error' | 'warn' | 'info' | 'action';
export type FrontendTelemetryCategory =
  | 'console'
  | 'preview'
  | 'accept'
  | 'reject'
  | 'apply'
  | 'sse'
  | 'navigation'
  | 'other';

export interface FrontendTelemetryEvent {
  ts?: string;
  level: FrontendTelemetryLevel;
  category: FrontendTelemetryCategory;
  event: string;
  message: string;
  conversationId?: string;
  changeSetId?: string;
  sessionId?: string;
  workbookKey?: string;
  userAgent?: string;
  pageUrl?: string;
  details?: Record<string, unknown>;
}

const FLUSH_MS = 1500;
const MAX_BATCH = 40;
const MAX_QUEUE = 200;

let sessionId =
  typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : `sess_${Date.now()}`;
let workbookKey: string | undefined;
let conversationId: string | undefined;
let queue: FrontendTelemetryEvent[] = [];
let flushTimer: ReturnType<typeof setTimeout> | null = null;
let consoleInstalled = false;
let flushing = false;

function getEndpoint(): string {
  return `${getApiBaseUrl()}/telemetry/frontend`;
}

function summarizeActions(actions: SheetAction[] | undefined): Record<string, unknown> {
  if (!actions?.length) return { actionCount: 0 };
  return {
    actionCount: actions.length,
    types: actions.map((a) => a.type),
    first: actions[0]
      ? {
          type: actions[0].type,
          sheetName: actions[0].sheetName,
          range: actions[0].range,
          row: actions[0].row,
          col: actions[0].col,
          key: actions[0].key,
          ascending: actions[0].ascending,
        }
      : undefined,
  };
}

function serializeError(error: unknown): Record<string, unknown> {
  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
      stack: error.stack?.slice(0, 4000),
    };
  }
  return { message: String(error) };
}

function enqueue(event: FrontendTelemetryEvent): void {
  queue.push({
    ...event,
    ts: event.ts ?? new Date().toISOString(),
    sessionId: event.sessionId ?? sessionId,
    workbookKey: event.workbookKey ?? workbookKey,
    conversationId: event.conversationId ?? conversationId,
    pageUrl: event.pageUrl ?? (typeof location !== 'undefined' ? location.href : undefined),
    userAgent:
      event.userAgent ?? (typeof navigator !== 'undefined' ? navigator.userAgent : undefined),
  });

  if (queue.length > MAX_QUEUE) {
    queue = queue.slice(-MAX_QUEUE);
  }

  if (!flushTimer) {
    flushTimer = setTimeout(() => {
      flushTimer = null;
      void flush();
    }, FLUSH_MS);
  }
}

async function flush(): Promise<void> {
  if (flushing || queue.length === 0) return;
  flushing = true;
  const batch = queue.splice(0, MAX_BATCH);
  try {
    await fetch(getEndpoint(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ events: batch }),
      keepalive: true,
    });
  } catch {
    // Re-queue failed batch (best effort; drop if queue already full).
    queue = [...batch, ...queue].slice(0, MAX_QUEUE);
  } finally {
    flushing = false;
    if (queue.length > 0 && !flushTimer) {
      flushTimer = setTimeout(() => {
        flushTimer = null;
        void flush();
      }, FLUSH_MS);
    }
  }
}

export const frontendTelemetry = {
  setContext( partial: {
    workbookKey?: string;
    conversationId?: string | null;
    sessionId?: string;
  }): void {
    if (partial.workbookKey) workbookKey = partial.workbookKey;
    if (partial.conversationId !== undefined) {
      conversationId = partial.conversationId ?? undefined;
    }
    if (partial.sessionId) sessionId = partial.sessionId;
  },

  log(event: FrontendTelemetryEvent): void {
    enqueue(event);
  },

  logAction(
    category: FrontendTelemetryCategory,
    event: string,
    message: string,
    details?: Record<string, unknown>,
  ): void {
    enqueue({
      level: 'action',
      category,
      event,
      message,
      details,
    });
  },

  logPreviewStart(actions: SheetAction[], explanation: string, meta?: { changeSetId?: string }): void {
    enqueue({
      level: 'action',
      category: 'preview',
      event: 'preview.start',
      message: explanation || 'Preview started',
      changeSetId: meta?.changeSetId,
      details: summarizeActions(actions),
    });
  },

  logPreviewFail(error: unknown, actions: SheetAction[]): void {
    enqueue({
      level: 'error',
      category: 'preview',
      event: 'preview.fail',
      message: error instanceof Error ? error.message : String(error),
      details: { ...summarizeActions(actions), error: serializeError(error) },
    });
  },

  logAcceptClick(actions: SheetAction[], meta?: { changeSetId?: string; source?: string }): void {
    enqueue({
      level: 'action',
      category: 'accept',
      event: 'accept.click',
      message: `Accept clicked (${actions.length} action(s))`,
      changeSetId: meta?.changeSetId,
      details: { ...summarizeActions(actions), source: meta?.source ?? 'ui' },
    });
  },

  logAcceptSuccess(actions: SheetAction[], meta?: { changeSetId?: string; explanation?: string }): void {
    enqueue({
      level: 'action',
      category: 'accept',
      event: 'accept.success',
      message: meta?.explanation || `Applied ${actions.length} action(s)`,
      changeSetId: meta?.changeSetId,
      details: summarizeActions(actions),
    });
  },

  logAcceptFail(error: unknown, actions: SheetAction[], meta?: { changeSetId?: string }): void {
    enqueue({
      level: 'error',
      category: 'accept',
      event: 'accept.fail',
      message: error instanceof Error ? error.message : String(error),
      changeSetId: meta?.changeSetId,
      details: { ...summarizeActions(actions), error: serializeError(error) },
    });
  },

  logReject(meta?: { changeSetId?: string; source?: string }): void {
    enqueue({
      level: 'action',
      category: 'reject',
      event: 'reject.click',
      message: 'Reject clicked — preview cleared',
      changeSetId: meta?.changeSetId,
      details: { source: meta?.source ?? 'ui' },
    });
  },

  /** Capture console.error / console.warn and window error events once. */
  installConsoleCapture(): void {
    if (consoleInstalled || typeof window === 'undefined') return;
    consoleInstalled = true;

    const originalError = console.error.bind(console);
    const originalWarn = console.warn.bind(console);

    console.error = (...args: unknown[]) => {
      originalError(...args);
      enqueue({
        level: 'error',
        category: 'console',
        event: 'console.error',
        message: args.map(stringifyArg).join(' ').slice(0, 2000),
        details: { args: args.map(safeArg).slice(0, 8) },
      });
    };

    console.warn = (...args: unknown[]) => {
      originalWarn(...args);
      const joined = args.map(stringifyArg).join(' ');
      // Avoid feedback loops from telemetry itself.
      if (joined.includes('[Cellix telemetry]')) {
        return;
      }
      enqueue({
        level: 'warn',
        category: 'console',
        event: 'console.warn',
        message: joined.slice(0, 2000),
        details: { args: args.map(safeArg).slice(0, 8) },
      });
    };

    window.addEventListener('error', (ev) => {
      enqueue({
        level: 'error',
        category: 'console',
        event: 'window.error',
        message: ev.message || 'Unhandled error',
        details: {
          filename: ev.filename,
          lineno: ev.lineno,
          colno: ev.colno,
          error: serializeError(ev.error),
        },
      });
    });

    window.addEventListener('unhandledrejection', (ev) => {
      enqueue({
        level: 'error',
        category: 'console',
        event: 'unhandledrejection',
        message: 'Unhandled promise rejection',
        details: { reason: serializeError(ev.reason) },
      });
    });

    window.addEventListener('pagehide', () => {
      void flush();
    });
  },

  flush,
};

function stringifyArg(value: unknown): string {
  if (typeof value === 'string') return value;
  if (value instanceof Error) return `${value.name}: ${value.message}`;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function safeArg(value: unknown): unknown {
  if (value instanceof Error) return serializeError(value);
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return String(value);
  }
}
