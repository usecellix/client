import { WorkbookContext } from '@/types/cellix.types';
import { buildWorkbookContext } from '@/services/sheetContextBuilder';

export type PendingWorkbookSnapshot = {
  toon: string;
  hash: string;
  workbookContext: WorkbookContext;
  activeSheetData: unknown[][];
  stale: boolean;
};

let pending: PendingWorkbookSnapshot | null = null;

/** Simple non-crypto hash for stale/fresh identity (matches prior useConversation helper). */
export function simpleHash(value: string): string {
  let hash = 0;
  for (let i = 0; i < value.length; i++) {
    hash = (hash << 5) - hash + value.charCodeAt(i);
    hash |= 0;
  }
  return String(hash);
}

export function getPendingWorkbookSnapshot(): PendingWorkbookSnapshot | null {
  return pending;
}

/** Mark prebuild invalid — next getContextForSend() will rebuild via Office.js. */
export function markPendingWorkbookContextStale(): void {
  if (pending) {
    pending = { ...pending, stale: true };
  }
}

/** Store a fresh prebuild (selection-change warm path). */
export function setPendingWorkbookContext(snapshot: {
  toon: string;
  workbookContext: WorkbookContext;
  activeSheetData?: unknown[][];
}): void {
  pending = {
    toon: snapshot.toon,
    hash: simpleHash(snapshot.toon),
    workbookContext: snapshot.workbookContext,
    activeSheetData: snapshot.activeSheetData ?? [],
    stale: false,
  };
}

export function clearPendingWorkbookContext(): void {
  pending = null;
}

/**
 * Spec 09 item 1: reuse pendingToon prebuild across sends until selection/edit/apply
 * marks it stale. Zero Office.js when a fresh prebuild exists.
 */
export async function getContextForSend(): Promise<{
  workbookContext: WorkbookContext;
  promptContext: string;
  sheetData: unknown[][];
  reusedPending: boolean;
}> {
  if (pending && !pending.stale) {
    return {
      workbookContext: pending.workbookContext,
      promptContext: pending.toon,
      sheetData: pending.activeSheetData,
      reusedPending: true,
    };
  }

  const built = await buildWorkbookContext([]);
  const toon = built.promptContext ?? '';
  setPendingWorkbookContext({
    toon,
    workbookContext: built.context,
    activeSheetData: built.activeSheetData,
  });

  return {
    workbookContext: built.context,
    promptContext: toon,
    sheetData: built.activeSheetData,
    reusedPending: false,
  };
}

/** Test-only reset. */
export function __resetPendingWorkbookContextForTests(): void {
  pending = null;
}
