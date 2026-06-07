import React from 'react';
import { cn } from '@/lib/utils';

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  size?: 'icon' | 'default' | 'sm' | 'lg';
  variant?: 'default' | 'outline' | 'ghost';
  children: React.ReactNode;
}

export const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, size = 'default', variant = 'default', children, ...props }, ref) => {
    return (
      <button
        ref={ref}
        className={cn(
          'inline-flex items-center justify-center rounded-md font-medium transition-colors',
          'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-2',
          'disabled:pointer-events-none disabled:opacity-50',
          size === 'icon' && 'h-10 w-10',
          size === 'sm' && 'h-9 px-3 text-sm',
          size === 'default' && 'h-10 px-4 py-2',
          size === 'lg' && 'h-11 px-8',
          variant === 'default' && 'bg-primary text-primary-foreground hover:bg-primary/90',
          variant === 'outline' && 'border border-input bg-background hover:bg-accent',
          variant === 'ghost' && 'hover:bg-accent hover:text-accent-foreground',
          className
        )}
        {...props}
      >
        {children}
      </button>
    );
  }
);

Button.displayName = 'Button';

