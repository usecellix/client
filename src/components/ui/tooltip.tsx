import React, { useState, useRef, useEffect, createContext, useContext } from 'react';
import { cn } from '@/lib/utils';

interface TooltipContextValue {
  isVisible: boolean;
  handleMouseEnter: () => void;
  handleMouseLeave: () => void;
}

const TooltipContext = createContext<TooltipContextValue | null>(null);

interface TooltipProps {
  children: React.ReactNode;
}

export const Tooltip: React.FC<TooltipProps> = ({ children }) => {
  const [isVisible, setIsVisible] = useState(false);
  const timeoutRef = useRef<NodeJS.Timeout>();

  const handleMouseEnter = () => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    timeoutRef.current = setTimeout(() => setIsVisible(true), 300);
  };

  const handleMouseLeave = () => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    setIsVisible(false);
  };

  useEffect(() => {
    return () => {
      if (timeoutRef.current) clearTimeout(timeoutRef.current);
    };
  }, []);

  return (
    <TooltipContext.Provider value={{ isVisible, handleMouseEnter, handleMouseLeave }}>
      <div className="relative inline-block" onMouseEnter={handleMouseEnter} onMouseLeave={handleMouseLeave}>
        {children}
      </div>
    </TooltipContext.Provider>
  );
};

export const TooltipTrigger = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement> & { asChild?: boolean }
>(({ asChild, children, ...props }, ref) => {
  const context = useContext(TooltipContext);
  
  if (!context) {
    throw new Error('TooltipTrigger must be used within a Tooltip component');
  }

  const { handleMouseEnter, handleMouseLeave } = context;

  if (asChild && React.isValidElement(children)) {
    return React.cloneElement(children, {
      ref,
      onMouseEnter: handleMouseEnter,
      onMouseLeave: handleMouseLeave,
      ...props,
    } as any);
  }
  return (
    <div
      ref={ref}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      {...props}
    >
      {children}
    </div>
  );
});

TooltipTrigger.displayName = 'TooltipTrigger';

export const TooltipContent: React.FC<React.HTMLAttributes<HTMLDivElement>> = ({
  children,
  className,
  ...props
}) => {
  const context = useContext(TooltipContext);
  
  if (!context) {
    throw new Error('TooltipContent must be used within a Tooltip component');
  }

  const { isVisible } = context;

  if (!isVisible) return null;

  return (
    <div
      className={cn(
        'absolute z-50 px-2 py-1 text-xs text-white bg-gray-900 rounded-md shadow-lg',
        'bottom-full left-1/2 transform -translate-x-1/2 mb-2',
        'whitespace-nowrap',
        className
      )}
      {...props}
    >
      {children}
      <div className="absolute top-full left-1/2 transform -translate-x-1/2 -mt-1">
        <div className="border-4 border-transparent border-t-gray-900" />
      </div>
    </div>
  );
};

