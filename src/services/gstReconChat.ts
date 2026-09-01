/**
 * End-to-end orchestrator for chat-triggered GST recon (client-side).
 * Called from useConversation when intent is detected — no form UI.
 */

import {
  detectGstReconIntent,
  GstReconIntentType,
} from '@/utils/gstReconIntent';
import {
  buildAmbiguousSheetMessage,
  buildMissingSheetMessage,
  discoverGstSheets,
  inferColumnMapping,
} from '@/services/gstSheetDiscovery';
import { readAllSheetHeaders, readSheetsFull } from '@/services/gstSheetReader';
import {
  buildGstReconAnswerText,
  buildGstReconUserFacingSummary,
  GstReconType,
  runGstReconcile,
} from '@/services/gstReconService';
import { SheetAction } from '@/types/sheet-actions';
import { UserFacingSummary } from '@/utils/userFacingResponse';

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
    };

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

/**
 * Returns null when the message is not a GST recon prompt.
 */
export async function tryHandleGstReconChat(
  message: string,
  options?: { conversationId?: string | null },
): Promise<GstReconChatOutcome | null> {
  const intent = detectGstReconIntent(message);
  if (!intent) return null;

  if (intent.type === 'GSTR3B_VS_GSTR2B') {
    return {
      kind: 'message_only',
      answer:
        'GSTR-3B vs GSTR-2B is a **summary** comparison (ITC claimed vs available). ' +
        'For now, use chat purchase/2B recon for invoice matching. ' +
        'To compare 3B totals, provide IGST/CGST/SGST figures from both returns in a follow-up, or paste both as sheets and ask for invoice-level PR vs 2B recon.',
    };
  }

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
  const toRead = [prName, portalName];
  if (discovery.ims && discovery.ims.name !== portalName) {
    toRead.push(discovery.ims.name);
  }

  const grids = await readSheetsFull(toRead);
  const byName = new Map(grids.map((g) => [g.name, g]));
  const prGrid = byName.get(prName);
  const portalGrid = byName.get(portalName);

  if (!prGrid?.values?.length || prGrid.values.length < 2) {
    return {
      kind: 'message_only',
      answer: `Sheet "${prName}" looks empty. Add purchase invoice rows with headers, then ask again.`,
    };
  }
  if (!portalGrid?.values?.length || portalGrid.values.length < 2) {
    return {
      kind: 'message_only',
      answer: `Sheet "${portalName}" looks empty. Paste the GSTR portal download into that sheet, then ask again.`,
    };
  }

  const result = await runGstReconcile({
    reconciliation_type: intentToApiType(intent.type),
    purchase_register: {
      sheet_name: prName,
      data: prGrid.values,
      column_mapping: inferColumnMapping(prGrid.headers),
      headers_row: 1,
    },
    portal_file: {
      sheet_name: portalName,
      data: portalGrid.values,
      column_mapping: inferColumnMapping(portalGrid.headers),
      file_type:
        intent.type === 'PR_VS_GSTR2A'
          ? 'GSTR2A'
          : intent.type === 'IMS_VS_PR'
            ? 'IMS'
            : 'GSTR2B',
      headers_row: 1,
    },
    ims_data:
      discovery.ims && byName.get(discovery.ims.name)?.values?.length
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
      detect_rcm: true,
      use_ims_data: Boolean(discovery.ims),
    },
    conversation_id: options?.conversationId ?? undefined,
  });

  const userFacingSummary = buildGstReconUserFacingSummary(result, prName, portalName);
  const answer = buildGstReconAnswerText(result, prName, portalName);

  return {
    kind: 'recon_ready',
    answer,
    actions: result.actions ?? [],
    explanation: userFacingSummary.headline,
    userFacingSummary,
  };
}
