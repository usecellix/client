import { DataValidationAction } from '@/action.types';
import { resolveWorksheet } from '../sheetResolve';

/* global Excel */

/**
 * DATA_VALIDATION — dropdown lists and value rules. TASKS.md #166.
 *
 * The capability Shortcut used and Cellix had no way to express: Payment
 * Status and Bank Account as real dropdowns rather than free text, backed by a
 * hidden Lists sheet. Without it a "tracker" is just typed strings, and every
 * downstream `SUMIF(...,"Paid",...)` is one typo away from silently reading
 * zero — which is the failure mode that makes an unvalidated tracker worse
 * than no tracker.
 *
 * Two Office.js details this encodes so callers never have to:
 *
 *  - A **range reference** source (`=Lists!$B$3:$B$9`) keeps working when the
 *    source sheet is hidden, which is the whole reason the hidden-Lists-sheet
 *    pattern is viable. Office.js wants that string to start with `=`; a
 *    planner writing the plain `Lists!$B$3:$B$9` is right about the intent and
 *    wrong about the syntax, so it is normalised here rather than being made
 *    the model's problem.
 *  - A **literal list** goes in as a single comma-joined string, not an array.
 *    Passing the array is the natural guess and silently produces a validation
 *    rule that matches nothing.
 */
export async function handleDataValidation(
  action: DataValidationAction,
  ctx: Excel.RequestContext,
): Promise<void> {
  const spec = action.validation;
  if (!spec) {
    throw new Error('DATA_VALIDATION requires a validation spec');
  }

  const range = String(action.range ?? '').trim();
  if (!range) {
    throw new Error('DATA_VALIDATION requires a range');
  }

  const sheet = resolveWorksheet(ctx, action.sheetName);
  const target = sheet.getRange(range);

  // Clear any existing rule first: Office.js rejects a second rule on a range
  // that already carries one, and re-running a build must not fail on its own
  // previous output.
  target.dataValidation.clear();

  const rule = buildRule(spec);
  if (!rule) {
    throw new Error(`DATA_VALIDATION: unsupported kind "${String(spec.kind)}"`);
  }

  target.dataValidation.rule = rule as Excel.DataValidationRule;
  target.dataValidation.ignoreBlanks = spec.ignoreBlanks !== false;

  if (spec.promptMessage) {
    target.dataValidation.prompt = {
      message: spec.promptMessage,
      showPrompt: true,
      title: spec.promptTitle ?? '',
    };
  }

  // A list rule with no error alert accepts anything typed over it, which
  // defeats the point of the dropdown — so lists default to a hard stop.
  const wantsAlert = Boolean(spec.errorMessage) || spec.kind === 'list';
  if (wantsAlert) {
    target.dataValidation.errorAlert = {
      message: spec.errorMessage ?? 'Please choose a value from the list.',
      showAlert: true,
      style: (spec.errorStyle ?? 'stop') as Excel.DataValidationAlertStyle,
      title: spec.errorTitle ?? 'Invalid entry',
    };
  }

  await ctx.sync();
}

function buildRule(spec: NonNullable<DataValidationAction['validation']>): unknown | null {
  if (spec.kind === 'list') {
    return { list: { inCellDropDown: true, source: resolveListSource(spec.listSource) } };
  }

  const operator = spec.operator ?? 'between';
  const bounds = {
    formula1: spec.formula1 ?? '',
    formula2: spec.formula2,
    operator,
  };

  switch (spec.kind) {
    case 'decimal':
      return { decimal: bounds };
    case 'wholeNumber':
      return { wholeNumber: bounds };
    case 'date':
      return { date: bounds };
    case 'textLength':
      return { textLength: bounds };
    default:
      return null;
  }
}

/**
 * Literal values become one comma-joined string; a range reference is passed
 * through with a leading `=` guaranteed.
 */
export function resolveListSource(source: string[] | string | undefined): string {
  if (Array.isArray(source)) {
    return source.map((entry) => String(entry).trim()).filter(Boolean).join(',');
  }

  const text = String(source ?? '').trim();
  if (text === '') return '';

  // Anything naming a sheet or an absolute cell is a reference, not a literal.
  const looksLikeReference = text.startsWith('=') || /!\$?[A-Za-z]+\$?\d+/.test(text);
  if (!looksLikeReference) return text;

  return text.startsWith('=') ? text : `=${text}`;
}
