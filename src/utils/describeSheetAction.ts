import { SheetAction } from '@/types/sheet-actions';

/** Turns 'RENAME_SHEET' into 'Rename sheet' for the generic fallback. */
function humanizeType(type: string): string {
  const words = type.toLowerCase().split('_');
  return words[0].charAt(0).toUpperCase() + words[0].slice(1) + ' ' + words.slice(1).join(' ');
}

function sortColumnLabel(action: SheetAction): string {
  if (action.sortBy?.column) return action.sortBy.column;
  if (action.columnName) return action.columnName;
  if (typeof action.key === 'number') return `column ${action.key + 1}`;
  return 'the first column';
}

/** Plain-language, one-line description of a single sheet action — for the user-facing details panel, never internal jargon (tiers, models, raw payload keys). */
export function describeSheetAction(action: SheetAction): string {
  const sheet = action.sheetName ? ` on '${action.sheetName}'` : '';
  const range = action.range ?? action.address ?? action.startCell;

  switch (action.type) {
    case 'SORT_RANGE': {
      const dir = (action.sortBy?.direction ?? (action.ascending === false ? 'desc' : 'asc')) === 'desc'
        ? 'descending'
        : 'ascending';
      return `Sort ${range ?? 'the range'}${sheet} by ${sortColumnLabel(action)}, ${dir}.`;
    }
    case 'SET_CELL':
      return `Set ${action.address ?? 'a cell'}${sheet} to ${JSON.stringify(action.value)}.`;
    case 'SET_FORMULA':
      return `Set formula ${action.formula ?? ''} in ${action.address ?? 'a cell'}${sheet}.`;
    case 'CLEAR_CELL':
      return `Clear ${action.address ?? 'a cell'}${sheet}.`;
    case 'SET_RANGE_VALUES':
      return `Write values into ${range ?? 'a range'}${sheet}.`;
    case 'WRITE_TABLE':
      return `Write a table into ${range ?? 'a range'}${sheet}.`;
    case 'CREATE_TABLE':
      return `Turn ${range ?? 'the range'}${sheet} into a table${action.tableName ? ` named '${action.tableName}'` : ''}.`;
    case 'DELETE_TABLE':
      return `Delete table${action.tableName ? ` '${action.tableName}'` : ''}${sheet}.`;
    case 'FORMAT_RANGE':
      return `Format ${range ?? 'the range'}${sheet}.`;
    case 'FORMAT_MATCHING_ROWS':
      return `Format rows matching a condition in ${range ?? 'the range'}${sheet}.`;
    case 'SET_MATCHING_ROWS':
      return `Update column ${action.targetColumn ?? ''} for rows matching a condition${sheet}.`;
    case 'CONDITIONAL_FORMAT':
      return `Add a conditional formatting rule to ${range ?? 'the range'}${sheet}.`;
    case 'DELETE_CONDITIONAL_FORMAT':
      return `Remove a conditional formatting rule${sheet}.`;
    case 'INSERT_ROW':
    case 'ADD_ROW':
      return `Insert ${action.count ?? 1} row${(action.count ?? 1) === 1 ? '' : 's'}${sheet}.`;
    case 'DELETE_ROW':
      return `Delete ${action.rowNumbers?.length ?? action.count ?? 1} row${(action.rowNumbers?.length ?? action.count ?? 1) === 1 ? '' : 's'}${sheet}.`;
    case 'INSERT_COLUMN':
      return `Insert a column${action.afterColumn ? ` after '${action.afterColumn}'` : ''}${sheet}.`;
    case 'DELETE_COLUMN':
      return `Delete column${action.columns?.length ? ` '${action.columns.join(', ')}'` : ''}${sheet}.`;
    case 'HIDE_ROW':
    case 'HIDE_COLUMN':
      return `Hide ${action.type === 'HIDE_ROW' ? 'row(s)' : 'column(s)'}${sheet}.`;
    case 'UNHIDE_ROW':
    case 'SHOW_ROW':
    case 'UNHIDE_COLUMN':
    case 'SHOW_COLUMN':
      return `Unhide ${action.type.includes('ROW') ? 'row(s)' : 'column(s)'}${sheet}.`;
    case 'MERGE_CELLS':
      return `Merge ${range ?? 'the selected cells'}${sheet}.`;
    case 'UNMERGE_CELLS':
      return `Unmerge ${range ?? 'the selected cells'}${sheet}.`;
    case 'CLEAR_RANGE':
    case 'CLEAR_CONTENT':
    case 'CLEAR_FORMAT':
    case 'CLEAR_ALL':
      return `Clear ${range ?? 'the range'}${sheet}.`;
    case 'MOVE_RANGE':
      return `Move ${action.sourceRange ?? 'a range'} to ${action.targetRange ?? action.destCell ?? 'a new location'}${sheet}.`;
    case 'COPY_FILTERED_RANGE':
      return `Copy matching rows from ${action.sourceSheet ?? 'this sheet'} to ${action.destSheet ?? 'another sheet'}.`;
    case 'AGGREGATE_TABLE':
      return `Summarize ${range ?? 'the data'}${sheet}${action.groupByColumn ? ` grouped by ${action.groupByColumn}` : ''}.`;
    case 'AUTO_FILTER':
      return `Add filters to ${range ?? 'the range'}${sheet}.`;
    case 'FREEZE_PANES':
      return `Freeze panes${sheet}.`;
    case 'UNFREEZE_PANES':
      return `Unfreeze panes${sheet}.`;
    case 'AUTOFIT_COLUMNS':
      return `Autofit column widths${sheet}.`;
    case 'SET_ROW_HEIGHT':
      return `Set row height${sheet}.`;
    case 'SET_COLUMN_WIDTH':
      return `Set column width${sheet}.`;
    case 'CREATE_SHEET':
    case 'ADD_SHEET':
      return `Create a new sheet${action.newSheetName ?? action.name ? ` named '${action.newSheetName ?? action.name}'` : ''}.`;
    case 'DELETE_SHEET':
      return `Delete sheet '${action.sheetName ?? ''}'.`;
    case 'RENAME_SHEET':
      return `Rename sheet '${action.oldName ?? action.sheetName ?? ''}' to '${action.newName ?? ''}'.`;
    case 'COPY_SHEET':
      return `Copy sheet '${action.sourceName ?? action.sheetName ?? ''}' to '${action.newSheetName ?? action.newName ?? ''}'.`;
    case 'HIDE_SHEET':
      return `Hide sheet '${action.sheetName ?? ''}'.`;
    case 'SHOW_SHEET':
      return `Show sheet '${action.sheetName ?? ''}'.`;
    case 'SET_SHEET_COLOR':
      return `Set the tab color for '${action.sheetName ?? ''}'.`;
    case 'CREATE_CHART':
      return `Create a ${action.chartType ?? ''} chart${action.title ? ` titled '${action.title}'` : ''}${sheet}.`;
    case 'UPDATE_CHART':
      return `Update a chart${sheet}.`;
    case 'DELETE_CHART':
      return `Delete a chart${sheet}.`;
    case 'DEFINE_NAMED_RANGE':
      return `Define a named range${action.name ? ` '${action.name}'` : ''} for ${range ?? 'a range'}.`;
    case 'ADD_COMMENT':
      return `Add a comment to ${action.address ?? 'a cell'}${sheet}.`;
    case 'DELETE_COMMENT':
      return `Delete a comment from ${action.address ?? 'a cell'}${sheet}.`;
    case 'PROTECT_SHEET':
      return `Protect sheet${sheet}.`;
    case 'UNPROTECT_SHEET':
      return `Unprotect sheet${sheet}.`;
    case 'FILL_DOWN':
      return `Fill down ${range ?? 'the range'}${sheet}.`;
    case 'FILL_RIGHT':
      return `Fill right ${range ?? 'the range'}${sheet}.`;
    case 'BATCH_SET':
      return `Update ${action.operations?.length ?? 'multiple'} cell${(action.operations?.length ?? 0) === 1 ? '' : 's'}${sheet}.`;
    case 'SET_ZOOM':
      return `Set zoom to ${action.zoomPercent ?? ''}%${sheet}.`;
    case 'HIGHLIGHT_CELL':
      return `Highlight ${action.address ?? 'a cell'}${sheet}.`;
    default:
      return `${humanizeType(action.type)}${sheet}.`;
  }
}

/** Caps the list so a huge batch doesn't flood the details panel. */
/**
 * Default limit is generous because the only caller renders this *inside* the
 * collapsed "Show details" disclosure — the place a user goes precisely to see
 * everything. It used to cap at 8, so a 190-action build's details read
 * "+182 more actions." and the full list existed nowhere in the UI. The card
 * body stays short by summarizing (see `rollUpActionsForUser` server-side);
 * that is what keeps the card small, not truncating the details. TASKS.md #140.
 */
export function describeSheetActions(actions: SheetAction[], limit = 250): string[] {
  const lines = actions.slice(0, limit).map(describeSheetAction);
  const remaining = actions.length - limit;
  if (remaining > 0) {
    lines.push(`+${remaining} more action${remaining === 1 ? '' : 's'}.`);
  }
  return lines;
}
