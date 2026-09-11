import { AddSheetAction, DeleteSheetAction, RenameSheetAction, CopySheetAction } from '@/action.types';
import { sanitizeExcelSheetName } from '@/utils/sheetName.util';

/* global Excel */

export async function handleDeleteSheet(
  action: DeleteSheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  ctx.workbook.worksheets.getItem(action.sheetName).delete();
  await ctx.sync();
}

/**
 * What actually happened to the sheet the caller asked for, versus what was
 * asked. `sanitizeExcelSheetName` can truncate/strip characters, and Excel's
 * own `worksheets.add` silently renames on any residual collision — neither
 * is visible to a caller reading only `void`. This is the ground truth a
 * `StructuralIntentChecker`-style comparison (or a manual read) needs to catch
 * the `Main` -> `Main 2` class of failure (COMPETITIVE_STUDY_SHORTCUT.md:71).
 */
export interface SheetCreationOutcome {
  requestedName: string;
  actualName: string;
  /** True when a sheet by (sanitized) name already existed and was reused instead of created. */
  reusedExisting: boolean;
}

export async function handleAddSheet(
  action: AddSheetAction,
  ctx: Excel.RequestContext,
): Promise<SheetCreationOutcome> {
  const sheets = ctx.workbook.worksheets;
  const name = sanitizeExcelSheetName(action.name);
  const existing = sheets.getItemOrNullObject(name);
  existing.load('isNullObject');
  await ctx.sync();
  if (!existing.isNullObject) {
    existing.activate();
    await ctx.sync();
    return { requestedName: action.name, actualName: name, reusedExisting: true };
  }

  let created: Excel.Worksheet;

  if (action.copyFrom) {
    const source = sheets.getItem(action.copyFrom);
    const copy = source.copy();
    copy.name = name;
    if (action.position !== undefined) {
      copy.position = action.position;
    }
    created = copy;
  } else {
    created = sheets.add(name);
    if (action.position !== undefined) {
      created.position = action.position;
    }
  }

  created.activate();
  created.load('name');
  await ctx.sync();

  return { requestedName: action.name, actualName: created.name, reusedExisting: false };
}

export async function handleRenameSheet(
  action: RenameSheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const sheet = ctx.workbook.worksheets.getItem(action.oldName);
  sheet.name = action.newName;
  await ctx.sync();
}

export async function handleCopySheet(
  action: CopySheetAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const source = ctx.workbook.worksheets.getItem(action.sourceName);
  const copy = source.copy();
  copy.name = action.newName;
  if (action.position !== undefined) {
    copy.position = action.position;
  }
  await ctx.sync();
}
