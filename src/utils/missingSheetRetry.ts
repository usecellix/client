import type { SheetAction } from '@/types/sheet-actions';

/* global Excel */

/**
 * Finish a step that failed only because a sheet it writes to does not exist —
 * TASKS.md #346.
 *
 * Live: step 4 of a 12-month ledger wrote to Main, but no step had created
 * Main. Accept failed with ItemNotFound and the only advice was "re-run the
 * request", which rebuilds all four steps (and costs credits) to fix one
 * missing sheet. Accepting the same card again could not help either: the
 * engine applies each action separately and keeps going past a failure, so the
 * month-sheet writes in that step had ALREADY landed, and replaying them trips
 * the overwrite guard on the cells they just wrote.
 *
 * So the retry is exactly what is left, and nothing else: create each missing
 * sheet, then re-apply only the actions that touch one. No model call.
 */

/** One `ACTION_TYPE: message` segment of a joined engine error. */
const ERROR_SEGMENT_SPLIT = /;\s+(?=[A-Z][A-Z0-9_]+:\s)/;
const ERROR_SEGMENT_RE = /^([A-Z][A-Z0-9_]+):\s/;
const ITEM_NOT_FOUND_RE = /requested resource doesn'?t exist|itemnotfound/i;

/** Fields that name a worksheet an action reads or writes. */
function sheetsTouchedBy(action: SheetAction): string[] {
  if (action.type === 'ADD_SHEET' || action.type === 'CREATE_SHEET') return [];
  return [action.sheetName, action.sourceSheet, action.destSheet, action.sourceSheetName]
    .map((name) => (typeof name === 'string' ? name.trim() : ''))
    .filter(Boolean);
}

/** Whether the action reads or writes any of `sheets` (case-insensitive). */
export function actionTouchesSheet(action: SheetAction, sheets: string[]): boolean {
  const wanted = new Set(sheets.map((sheet) => sheet.trim().toLowerCase()));
  return sheetsTouchedBy(action).some((sheet) => wanted.has(sheet.toLowerCase()));
}

function createdSheetName(action: SheetAction): string {
  if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') return '';
  return String(action.name ?? action.sheetName ?? '').trim();
}

/**
 * The action types that failed, when EVERY failure in the joined engine error
 * was ItemNotFound. Null when anything else also failed — a retry that only
 * creates sheets would leave that other failure unfixed while reporting the
 * step applied.
 */
export function parseItemNotFoundFailures(rawError: string): string[] | null {
  const segments = (rawError ?? '').trim().split(ERROR_SEGMENT_SPLIT).filter(Boolean);
  if (segments.length === 0) return null;
  const types: string[] = [];
  for (const segment of segments) {
    const type = ERROR_SEGMENT_RE.exec(segment)?.[1];
    if (!type || !ITEM_NOT_FOUND_RE.test(segment)) return null;
    types.push(type);
  }
  return types;
}

export interface MissingSheetRetryPlan {
  /** Sheets the step writes to that are not in the workbook, as the step names them. */
  missingSheets: string[];
  /** Creates for the missing sheets, then every action that touches one, in step order. */
  actions: SheetAction[];
}

/**
 * Build the retry, or null when it would not fix the failure.
 *
 * `existingSheetNames` is the workbook as it is NOW. Actions that touch only
 * existing sheets are left out: they are the ones that already applied.
 */
export function buildMissingSheetRetry(
  stepActions: SheetAction[],
  existingSheetNames: string[],
  rawError: string,
): MissingSheetRetryPlan | null {
  const failedTypes = parseItemNotFoundFailures(rawError);
  if (!failedTypes) return null;
  return planMissingSheetRetry(stepActions, existingSheetNames, failedTypes);
}

/**
 * The retry for the workbook as it is now, without an error to check against —
 * what the click itself uses, since the sheets missing at failure time may not
 * be the ones missing when the user gets round to clicking.
 */
export function planMissingSheetRetry(
  stepActions: SheetAction[],
  existingSheetNames: string[],
  failedTypes?: string[],
): MissingSheetRetryPlan | null {
  const existing = new Set(existingSheetNames.map((name) => name.trim().toLowerCase()));
  const missing = new Map<string, string>();
  for (const action of stepActions) {
    for (const sheet of sheetsTouchedBy(action)) {
      const key = sheet.toLowerCase();
      if (!existing.has(key) && !missing.has(key)) missing.set(key, sheet);
    }
  }
  if (missing.size === 0) return null;

  const retryWrites = stepActions.filter((action) => actionTouchesSheet(action, [...missing.values()]));

  // Every action type that failed must be one this retry re-applies —
  // otherwise the ItemNotFound came from something other than the missing
  // sheet (a table, a chart) and creating the sheet would not fix it.
  const retriedTypes = new Set<string>(retryWrites.map((action) => action.type));
  if (failedTypes && !failedTypes.every((type) => retriedTypes.has(type))) return null;

  // The step's own create for a missing sheet is reused rather than
  // duplicated; any sheet it never created gets one.
  const ownCreates = stepActions.filter((action) =>
    missing.has(createdSheetName(action).toLowerCase()),
  );
  const coveredByOwnCreate = new Set(ownCreates.map((a) => createdSheetName(a).toLowerCase()));
  const addedCreates: SheetAction[] = [...missing.entries()]
    .filter(([key]) => !coveredByOwnCreate.has(key))
    .map(([, name]) => ({ type: 'ADD_SHEET', name, sheetName: name }));

  return {
    missingSheets: [...missing.values()],
    actions: [...ownCreates, ...addedCreates, ...retryWrites],
  };
}

/** Every worksheet name in the workbook right now. */
export async function listWorksheetNames(): Promise<string[]> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();
    return sheets.items.map((sheet) => sheet.name);
  });
}

/** "'Main'" / "'Main' and 'Totals'" / "'A', 'B' and 'C'". */
export function describeSheetList(sheets: string[]): string {
  const quoted = sheets.map((sheet) => `'${sheet}'`);
  if (quoted.length <= 1) return quoted.join('');
  return `${quoted.slice(0, -1).join(', ')} and ${quoted[quoted.length - 1]}`;
}

/** The message shown when the retry is on offer — names the sheet, not "re-run". */
export function describeMissingSheetFailure(missingSheets: string[]): string {
  const list = describeSheetList(missingSheets);
  const noun = missingSheets.length === 1 ? 'sheet' : 'sheets';
  return (
    `This step writes to the ${noun} ${list}, which ${missingSheets.length === 1 ? "doesn't" : "don't"} ` +
    `exist yet. The rest of the step is already applied. Create the missing ${noun} to finish ` +
    `just what's left. The earlier steps won't be redone.`
  );
}
