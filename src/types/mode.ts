export type AssistantMode = 'ask' | 'action' | 'plan';

export const ASSISTANT_MODES: AssistantMode[] = ['ask', 'action', 'plan'];

export const DEFAULT_ASSISTANT_MODE: AssistantMode = 'action';

export interface AssistantModeMeta {
  id: AssistantMode;
  label: string;
  hint: string;
}

export const ASSISTANT_MODE_META: Record<AssistantMode, AssistantModeMeta> = {
  ask: {
    id: 'ask',
    label: 'Ask',
    hint: 'Read-only — find, search, explain. Never edits your sheets.',
  },
  action: {
    id: 'action',
    label: 'Action',
    hint: 'Execute changes with a preview and confirmation before applying.',
  },
  plan: {
    id: 'plan',
    label: 'Plan',
    hint: 'Design a step-by-step plan before any changes are made.',
  },
};

export function isAssistantMode(value: unknown): value is AssistantMode {
  return value === 'ask' || value === 'action' || value === 'plan';
}
