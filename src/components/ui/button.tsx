import * as React from 'react';
import { Slot } from '@radix-ui/react-slot';
import { cva, type VariantProps } from 'class-variance-authority';
import { cn } from '@/lib/utils';

const buttonVariants = cva(
  'inline-flex items-center justify-center whitespace-nowrap rounded-lg text-sm font-medium transition-all duration-150 active:scale-[0.97] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent-500 focus-visible:ring-offset-1 disabled:pointer-events-none disabled:opacity-50',
  {
    variants: {
      variant: {
        default:
          'bg-accent-600 text-white shadow-sm hover:bg-accent-700',
        'primary-neutral':
          'bg-neutral-900 text-white shadow-sm hover:bg-neutral-800',
        destructive:
          'bg-red-600 text-white shadow-sm hover:bg-red-700',
        outline:
          'border border-neutral-200/80 bg-white shadow-sm hover:bg-neutral-50 hover:border-neutral-300 hover:text-neutral-900',
        secondary:
          'bg-neutral-100 text-neutral-900 shadow-sm hover:bg-neutral-200',
        ghost: 'hover:bg-neutral-100 hover:text-neutral-900',
        link: 'text-accent-600 underline-offset-4 hover:underline',
      },
      size: {
        default: 'h-9 px-4 py-2',
        sm: 'h-8 rounded-lg px-3 text-xs',
        lg: 'h-10 rounded-lg px-8',
        icon: 'h-9 w-9',
      },
    },
    defaultVariants: {
      variant: 'default',
      size: 'default',
    },
  }
);

export interface ButtonProps
  extends React.ButtonHTMLAttributes<HTMLButtonElement>,
    VariantProps<typeof buttonVariants> {
  asChild?: boolean;
}

const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant, size, asChild = false, ...props }, ref) => {
    const Comp = asChild ? Slot : 'button';
    return (
      <Comp
        className={cn(buttonVariants({ variant, size, className }))}
        ref={ref}
        {...props}
      />
    );
  }
);
Button.displayName = 'Button';

// eslint-disable-next-line react-refresh/only-export-components
export { Button, buttonVariants };
