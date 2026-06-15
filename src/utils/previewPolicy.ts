import { SheetAction } from '@/types/sheet-actions';

/** Sheet-level mutations that must never auto-apply without explicit Accept. */
const CONFIRMATION_REQUIRED_TYPES = new Set<SheetAction['type']>([
  'DELETE_SHEET',
  'ADD_SHEET',
  'CREATE_SHEET',
  'COPY_SHEET',
  'RENAME_SHEET',
]);

export function requiresExplicitAccept(actions: SheetAction[]): boolean {
  return actions.some((action) => CONFIRMATION_REQUIRED_TYPES.has(action.type));
}

export function shouldPreviewActions(
  actions: SheetAction[],
  autoApplyActions: boolean,
): boolean {
  return !autoApplyActions || requiresExplicitAccept(actions);
}
