import { useState, useRef, KeyboardEvent } from 'react';
import { Send, Mic, AtSign, Grid3X3, Paperclip } from 'lucide-react';
import { cn } from '@/lib/utils';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';

interface ChatInputProps {
  onSend: (message: string) => void;
  disabled?: boolean;
  placeholder?: string;
}

const ChatInput = ({ onSend, disabled, placeholder = "Ask anything about your spreadsheet..." }: ChatInputProps) => {
  const [message, setMessage] = useState('');
  const [isFocused, setIsFocused] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleSend = () => {
    if (message.trim() && !disabled) {
      onSend(message.trim());
      setMessage('');
      if (textareaRef.current) {
        textareaRef.current.style.height = 'auto';
      }
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleInput = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setMessage(e.target.value);
    // Auto-resize textarea
    const textarea = e.target;
    textarea.style.height = 'auto';
    textarea.style.height = `${Math.min(textarea.scrollHeight, 120)}px`;
  };

  const insertAtCursor = (text: string) => {
    const textarea = textareaRef.current;
    if (textarea) {
      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const newMessage = message.substring(0, start) + text + message.substring(end);
      setMessage(newMessage);
      setTimeout(() => {
        textarea.focus();
        textarea.setSelectionRange(start + text.length, start + text.length);
      }, 0);
    }
  };

  return (
    <div className="bg-background" style={{ padding: 0 }}>
      {/* Input Container */}
      <div
        className={cn(
          "border bg-background transition-all duration-200",
          isFocused ? "border-primary" : "border-border"
        )}
        style={{
          borderRadius: '8px',
          border: isFocused ? '1px solid #4338CA' : '1px solid #E5E7EB',
          boxShadow: isFocused ? '0 0 0 3px rgba(67, 56, 202, 0.1)' : 'none',
          background: '#FFFFFF'
        }}
      >
        {/* Textarea */}
        <textarea
          ref={textareaRef}
          value={message}
          onChange={handleInput}
          onKeyDown={handleKeyDown}
          onFocus={() => setIsFocused(true)}
          onBlur={() => setIsFocused(false)}
          placeholder={disabled ? "Processing..." : placeholder}
          disabled={disabled}
          rows={1}
          className={cn(
            "w-full resize-none bg-transparent text-sm focus:outline-none disabled:opacity-50",
            "leading-relaxed",
            disabled && "animate-pulse"
          )}
          style={{ 
            maxHeight: '200px',
            minHeight: '44px',
            padding: '4px 4px 28px 4px',
            border: 'none',
            borderRadius: 0,
            fontSize: '14px',
            background: disabled ? '#F9FAFB' : 'white',
            color: disabled ? '#9CA3AF' : '#1E293B'
          }}
        />
        <style>{`
          textarea::placeholder {
            color: #9CA3AF;
            font-weight: 400;
          }
        `}</style>

        {/* Bottom Bar with Actions and Send */}
        <div className="flex items-center justify-between" style={{ padding: '4px' }}>
          <div className="flex items-center gap-1">
            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => {}}
                  className="transition-all active:scale-95"
                  style={{
                    width: '24px',
                    height: '24px',
                    borderRadius: '6px',
                    padding: 0,
                    background: '#F9FAFB',
                    color: '#4338CA',
                    border: 'none',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  aria-label="Voice input"
                >
                  <Mic className="h-4 w-4" />
                </button>
              </TooltipTrigger>
              <TooltipContent>Voice input</TooltipContent>
            </Tooltip>

            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => insertAtCursor('@')}
                  className="transition-all active:scale-95"
                  style={{
                    width: '24px',
                    height: '24px',
                    borderRadius: '6px',
                    padding: 0,
                    background: '#F9FAFB',
                    color: '#4338CA',
                    border: 'none',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  aria-label="Mention cell or range"
                >
                  <AtSign className="h-4 w-4" />
                </button>
              </TooltipTrigger>
              <TooltipContent>Mention cell or range</TooltipContent>
            </Tooltip>

            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => insertAtCursor('[A1:B10]')}
                  className="transition-all active:scale-95"
                  style={{
                    width: '24px',
                    height: '24px',
                    borderRadius: '6px',
                    padding: 0,
                    background: '#F9FAFB',
                    color: '#4338CA',
                    border: 'none',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  aria-label="Select cell range"
                >
                  <Grid3X3 className="h-4 w-4" />
                </button>
              </TooltipTrigger>
              <TooltipContent>Select cell range</TooltipContent>
            </Tooltip>

            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => {}}
                  className="transition-all active:scale-95"
                  style={{
                    width: '24px',
                    height: '24px',
                    borderRadius: '6px',
                    padding: 0,
                    background: '#F3F4F6',
                    color: '#1E293B',
                    border: 'none',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  aria-label="Attach file"
                >
                  <Paperclip className="h-4 w-4" />
                </button>
              </TooltipTrigger>
              <TooltipContent>Attach file</TooltipContent>
            </Tooltip>
          </div>

          {/* Send Button */}
          <button
            onClick={handleSend}
            disabled={!message.trim() || disabled}
            className="transition-all active:scale-95"
            style={{
              width: '24px',
              height: '24px',
              borderRadius: '6px',
              padding: 0,
              background: message.trim() ? '#4338CA' : '#F3F4F6',
              color: message.trim() ? '#FFFFFF' : '#9CA3AF',
              border: 'none',
              cursor: message.trim() && !disabled ? 'pointer' : 'not-allowed',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              opacity: disabled ? 0.6 : 1,
              boxShadow: message.trim() ? '0 1px 3px rgba(0, 0, 0, 0.1)' : 'none'
            }}
            onMouseEnter={(e) => {
              if (message.trim() && !disabled) {
                e.currentTarget.style.background = '#3730A3';
                e.currentTarget.style.transform = 'translateY(-1px)';
                e.currentTarget.style.boxShadow = '0 4px 8px rgba(67, 56, 202, 0.2)';
              }
            }}
            onMouseLeave={(e) => {
              if (message.trim() && !disabled) {
                e.currentTarget.style.background = '#4338CA';
                e.currentTarget.style.transform = 'translateY(0)';
                e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.1)';
              }
            }}
            aria-label="Send message"
          >
            <Send className="h-4 w-4" />
          </button>
        </div>
      </div>

      {/* Hint text */}
      <div className="flex justify-end" style={{ marginTop: '10px', padding: '0 4px' }}>
        <span className="flex items-center gap-1.5" style={{ fontSize: '12px', color: '#64748B' }}>
          <kbd style={{
            padding: '2px 8px',
            borderRadius: '4px',
            background: '#F5F7FA',
            color: '#1E293B',
            fontFamily: 'monospace',
            fontSize: '11px',
            border: '1px solid #E5E7EB',
            boxShadow: '0 1px 2px rgba(0, 0, 0, 0.05)'
          }}>Enter</kbd>
          <span>to send</span>
          <span style={{ color: '#9CA3AF' }}>•</span>
          <kbd style={{
            padding: '2px 8px',
            borderRadius: '4px',
            background: '#F5F7FA',
            color: '#1E293B',
            fontFamily: 'monospace',
            fontSize: '11px',
            border: '1px solid #E5E7EB',
            boxShadow: '0 1px 2px rgba(0, 0, 0, 0.05)'
          }}>Shift</kbd>
          <span>+</span>
          <kbd style={{
            padding: '2px 8px',
            borderRadius: '4px',
            background: '#F5F7FA',
            color: '#1E293B',
            fontFamily: 'monospace',
            fontSize: '11px',
            border: '1px solid #E5E7EB',
            boxShadow: '0 1px 2px rgba(0, 0, 0, 0.05)'
          }}>Enter</kbd>
          <span>for new line</span>
        </span>
      </div>
    </div>
  );
};

export default ChatInput;

