import { Plus } from "lucide-react";

import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";

interface SelectedRangeBadgeProps {
  selectedRange: string;
  onInsert?: () => void;
  disabled?: boolean;
}

export function SelectedRangeBadge({
  selectedRange,
  onInsert,
  disabled,
}: SelectedRangeBadgeProps) {
  if (!selectedRange) return null;

  const displayRange = selectedRange.includes("!")
    ? selectedRange.split("!").pop() ?? selectedRange
    : selectedRange;
  const chipClasses =
    "h-6 px-1.5 border border-blue-200 dark:border-blue-700 bg-blue-50 dark:bg-blue-900/20 text-blue-600 dark:text-blue-400 rounded-sm flex items-center";

  const badge = (
    <div
      data-testid="selected-range"
      className={`${chipClasses} cursor-default flex-shrink-0`}
    >
      <span className="text-[10px] font-mono leading-none">{displayRange}</span>
    </div>
  );

  const insertButton = onInsert && (
    <button
      type="button"
      onClick={onInsert}
      disabled={disabled}
      className={`${chipClasses} justify-center transition-colors ${
        disabled
          ? "cursor-not-allowed opacity-40"
          : "hover:bg-blue-100 dark:hover:bg-blue-900/40"
      }`}
      aria-label="Insert selected range into message"
    >
      <Plus className="w-3 h-3" />
    </button>
  );

  return (
    <div className="flex items-center gap-1">
      <Tooltip>
        <TooltipTrigger asChild>{badge}</TooltipTrigger>
        <TooltipContent>
          <p>Current Excel selection will be included with your message</p>
        </TooltipContent>
      </Tooltip>
      {insertButton && (
        <Tooltip>
          <TooltipTrigger asChild>{insertButton}</TooltipTrigger>
          <TooltipContent>
            <p>Insert selection into message</p>
          </TooltipContent>
        </Tooltip>
      )}
    </div>
  );
}
