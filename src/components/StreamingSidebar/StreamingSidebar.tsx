import React, { useEffect, useRef } from 'react';
import { useSseStream, SheetAction } from '../../hooks/useSseStream';

interface StreamingSidebarProps {
  prompt: string;
  sheetData: any[][];
  onComplete?: (actions: SheetAction[], explanation: string) => void;
  onError?: (error: string) => void;
  onCancel?: () => void;
}

const StreamingSidebar: React.FC<StreamingSidebarProps> = ({
  prompt,
  sheetData,
  onComplete,
  onError,
  onCancel,
}) => {
  const {
    isConnected,
    statusMessages,
    aiText,
    actions,
    explanation,
    error,
    isComplete,
    startStream,
    stopStream,
    reset,
  } = useSseStream();

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const completionHandledRef = useRef(false);

  // Auto-scroll to bottom
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [statusMessages, aiText]);

  // Start stream when prompt changes
  useEffect(() => {
    if (prompt && sheetData) {
      completionHandledRef.current = false;
      startStream(prompt, sheetData);
    }

    return () => {
      stopStream();
    };
  }, [prompt, sheetData, startStream, stopStream]);

  // Handle completion
  useEffect(() => {
    if (isComplete && !error && !completionHandledRef.current) {
      completionHandledRef.current = true;
      onComplete?.(actions, explanation || aiText);
    }
  }, [isComplete, error, actions, explanation, aiText, onComplete]);

  // Handle errors
  useEffect(() => {
    if (error) {
      onError?.(error);
    }
  }, [error, onError]);

  const handleCancel = () => {
    stopStream();
    onCancel?.();
  };

  const handleReset = () => {
    reset();
  };

  return (
    <div className="flex flex-col h-full font-sans bg-background">
      {/* Header */}
      <div 
        className="px-5 py-4 border-b border-border flex items-center justify-between"
        style={{
          background: 'linear-gradient(135deg, #4338CA 0%, #5B4FCF 100%)',
          color: 'white'
        }}
      >
        <div className="flex items-center gap-3">
          <div className={`w-2 h-2 rounded-full ${isConnected ? 'bg-green-400 animate-pulse' : 'bg-gray-400'}`} />
          <h1 className="m-0 text-base font-semibold text-white tracking-tight" style={{ fontSize: '16px', fontWeight: 600, lineHeight: '1.2' }}>
            AI Assistant
          </h1>
        </div>
        {isConnected && (
          <button
            onClick={handleCancel}
            className="px-3 py-1 text-xs font-medium text-white bg-white/20 rounded-md hover:bg-white/30 transition-colors"
          >
            Cancel
          </button>
        )}
      </div>
      
      {/* Messages Container */}
      <div 
        className="flex-1 overflow-y-auto px-5 py-5 flex flex-col scroll-smooth"
        style={{ 
          background: '#F8FAFC',
          padding: '20px'
        }}
      >
        {/* Status Messages */}
        {statusMessages.map((message, index) => (
          <div
            key={`status-${index}`}
            className="flex items-start gap-3 flex-row animate-fadeIn"
            style={{ marginBottom: '16px' }}
          >
            <div
              className="shrink-0 w-4 h-4 rounded-full flex items-center justify-center"
              style={{
                width: '16px',
                height: '16px',
                borderRadius: '50%',
                backgroundColor: '#DBEAFE',
                color: '#3B82F6',
                fontSize: '10px',
                fontWeight: 600
              }}
            >
              <svg width="8" height="8" viewBox="0 0 8 8" fill="currentColor">
                <circle cx="4" cy="4" r="2" />
              </svg>
            </div>
            <div
              className="max-w-[75%] wrap-break-word shadow-sm"
              style={{
                paddingLeft: '24px',
                fontSize: '14px',
                lineHeight: '1.6',
                color: '#374151'
              }}
            >
              <div className="flex items-center gap-2">
                <span>{message}</span>
                {index === statusMessages.length - 1 && isConnected && (
                  <div className="flex gap-1">
                    <span className="w-1.5 h-1.5 rounded-full bg-blue-400/60 animate-bounce" style={{ animationDelay: '0ms' }}></span>
                    <span className="w-1.5 h-1.5 rounded-full bg-blue-400/60 animate-bounce" style={{ animationDelay: '150ms' }}></span>
                    <span className="w-1.5 h-1.5 rounded-full bg-blue-400/60 animate-bounce" style={{ animationDelay: '300ms' }}></span>
                  </div>
                )}
              </div>
            </div>
          </div>
        ))}

        {/* AI Response Text */}
        {aiText && (
          <div
            className="flex items-start gap-3 flex-row animate-fadeIn"
            style={{ marginBottom: '16px' }}
          >
            <div
              className="shrink-0 w-4 h-4 rounded-full flex items-center justify-center"
              style={{
                width: '16px',
                height: '16px',
                borderRadius: '50%',
                backgroundColor: '#D1FAE5',
                color: '#10B981',
                fontSize: '10px',
                fontWeight: 600
              }}
            >
              AI
            </div>
            <div
              className="max-w-[75%] wrap-break-word shadow-sm"
              style={{
                paddingLeft: '24px',
                fontSize: '14px',
                lineHeight: '1.6',
                color: '#1F2937',
                whiteSpace: 'pre-wrap'
              }}
            >
              {aiText}
              {isConnected && (
                <span className="inline-block w-2 h-4 bg-gray-400 animate-pulse ml-1" />
              )}
            </div>
          </div>
        )}

        {/* Actions Applied */}
        {actions.length > 0 && (
          <div
            className="flex items-start gap-3 flex-row animate-fadeIn"
            style={{ marginBottom: '16px' }}
          >
            <div
              className="shrink-0 w-4 h-4 rounded-full flex items-center justify-center"
              style={{
                width: '16px',
                height: '16px',
                borderRadius: '50%',
                backgroundColor: '#FEF3C7',
                color: '#F59E0B',
                fontSize: '10px',
                fontWeight: 600
              }}
            >
              ✓
            </div>
            <div
              className="max-w-[75%] wrap-break-word shadow-sm"
              style={{
                paddingLeft: '24px',
                fontSize: '14px',
                lineHeight: '1.6',
                color: '#1F2937'
              }}
            >
              <div className="mb-2 font-medium text-green-600">
                Applied {actions.length} change{actions.length > 1 ? 's' : ''}:
              </div>
              <div className="text-sm text-gray-600 mb-2">
                {explanation}
              </div>
              <div className="space-y-1">
                {actions.map((action, index) => (
                  <div key={index} className="text-xs text-gray-500">
                    • {action.type} {action.row !== undefined && action.col !== undefined && `(row ${action.row}, col ${action.col})`}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* Error State */}
        {error && (
          <div
            className="flex items-start gap-3 flex-row animate-fadeIn"
            style={{ marginBottom: '16px' }}
          >
            <div
              className="shrink-0 w-4 h-4 rounded-full flex items-center justify-center"
              style={{
                width: '16px',
                height: '16px',
                borderRadius: '50%',
                backgroundColor: '#FEE2E2',
                color: '#EF4444',
                fontSize: '10px',
                fontWeight: 600
              }}
            >
              !
            </div>
            <div
              className="max-w-[75%] wrap-break-word shadow-sm"
              style={{
                paddingLeft: '24px',
                fontSize: '14px',
                lineHeight: '1.6',
                color: '#DC2626'
              }}
            >
              <div className="font-medium mb-1">Error</div>
              <div className="text-sm">{error}</div>
              <button
                onClick={handleReset}
                className="mt-2 px-3 py-1 text-xs font-medium text-white bg-red-500 rounded-md hover:bg-red-600 transition-colors"
              >
                Try Again
              </button>
            </div>
          </div>
        )}

        {/* Completion State */}
        {isComplete && !error && (
          <div
            className="flex items-start gap-3 flex-row animate-fadeIn"
            style={{ marginBottom: '16px' }}
          >
            <div
              className="shrink-0 w-4 h-4 rounded-full flex items-center justify-center"
              style={{
                width: '16px',
                height: '16px',
                borderRadius: '50%',
                backgroundColor: '#D1FAE5',
                color: '#10B981',
                fontSize: '10px',
                fontWeight: 600
              }}
            >
              ✓
            </div>
            <div
              className="max-w-[75%] wrap-break-word shadow-sm"
              style={{
                paddingLeft: '24px',
                fontSize: '14px',
                lineHeight: '1.6',
                color: '#059669'
              }}
            >
              <div className="font-medium">Complete!</div>
            </div>
          </div>
        )}
        
        <div ref={messagesEndRef} className="h-2" />
      </div>
    </div>
  );
};

export default StreamingSidebar;
