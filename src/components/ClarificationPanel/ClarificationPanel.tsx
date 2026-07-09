import React from 'react';
import { ClarificationPayload } from '@/types/cellix.types';
import QuestionChoicesPanel from '@/components/ConversationPanel/QuestionChoicesPanel';

interface ClarificationPanelProps {
  payload: ClarificationPayload;
  onAnswer: (answer: string) => void;
  onDismiss?: () => void;
  inline?: boolean;
}

export const ClarificationPanel: React.FC<ClarificationPanelProps> = ({
  payload,
  onAnswer,
  inline = false,
}) => {
  return (
    <div className={inline ? 'cellix-clarification-inline' : undefined}>
      <QuestionChoicesPanel
        question={payload.question}
        options={payload.suggestions}
        onSelect={onAnswer}
      />
    </div>
  );
};

export default ClarificationPanel;
