import { SheetAction } from '@/types/sheet-actions';

export type CheckpointTrigger = 'auto' | 'manual';
export type CheckpointStatus = 'active' | 'restored';

export interface CheckpointSummary {
  checkpointId: string;
  workbookId: string;
  conversationId: string;
  label: string;
  trigger: CheckpointTrigger;
  anchorChangeSetId: string;
  createdAt: string;
  status: CheckpointStatus;
  restoredAt?: string;
}

export interface RestoreResult {
  checkpoint: CheckpointSummary;
  revertedChangeSetIds: string[];
  inverseActions: SheetAction[];
}
