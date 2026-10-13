/**
 * Create a sheet before anything in the same card touches it — TASKS.md #352.
 *
 * `RichActionEngine` applies a card strictly in array order and keeps going
 * past a failure. Live (Sept 28, "Create and fill Lists"), the card arrived as
 * `[SET_CELL Lists!A1, ADD_SHEET Lists, BATCH_SET Lists, …]`: on a fresh
 * workbook the title write hit ItemNotFound, the create and the rest landed,
 * and Accept reported a failure over a sheet that was missing only its title.
 * `previewManager.accept()` assumes creates were hoisted server-side, but a
 * stepwise wave had not been.
 *
 * So each create moves to just before the first action that touches its sheet
 * — never further than that, and never past something that could change what
 * the name refers to:
 *  - a DELETE/RENAME/COPY of that name in between (delete-then-recreate
 *    must stay in that order, or the create reuses the old sheet and the
 *    delete then removes it; a copy INTO that name is itself the create);
 *  - a create that copies another sheet stays where it is, since its source
 *    may itself be created in between.
 */

const SHEET_OPS = new Set(['DELETE_SHEET', 'RENAME_SHEET', 'COPY_SHEET']);

type Named = { type: string } & Record<string, unknown>;

function createdName(action: Named): string {
  if (action.type !== 'ADD_SHEET' && action.type !== 'CREATE_SHEET') return '';
  return String(action.name ?? action.sheetName ?? '').trim().toLowerCase();
}

/** Every sheet name an action refers to, lower-cased. */
function namesIn(action: Named): string[] {
  return ['sheetName', 'sourceSheet', 'destSheet', 'sourceSheetName', 'oldName', 'newName', 'newSheetName', 'sourceName']
    .map((field) => action[field])
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .map((value) => value.trim().toLowerCase());
}

export function hoistSheetCreates<T extends { type: string }>(actions: T[]): T[] {
  const ordered = [...actions];
  for (let i = 0; i < ordered.length; i++) {
    const create = ordered[i] as unknown as Named;
    const name = createdName(create);
    if (!name || create.copyFrom) continue;

    const firstUse = ordered.findIndex(
      (action, j) => j < i && !createdName(action as unknown as Named) && namesIn(action as unknown as Named).includes(name),
    );
    if (firstUse < 0) continue;

    const blocked = ordered
      .slice(firstUse, i)
      .some((action) => SHEET_OPS.has(action.type) && namesIn(action as unknown as Named).includes(name));
    if (blocked) continue;

    const [moved] = ordered.splice(i, 1);
    ordered.splice(firstUse, 0, moved);
  }
  return ordered;
}
