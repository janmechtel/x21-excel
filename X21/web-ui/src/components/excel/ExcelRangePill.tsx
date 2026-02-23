import { cn } from "@/lib/utils";

interface ExcelRangePillProps {
  range: string;
  workbookName?: string;
  onClick?: (range: string, workbookName?: string) => void;
  className?: string;
  variant?: "default" | "minimal";
}

const BASE_STYLES =
  "inline-flex items-center font-mono text-[11px] font-medium px-1.5 py-0.5 border transition-colors bg-blue-50 dark:bg-blue-900/30 border-blue-200 dark:border-blue-600 text-blue-600 dark:text-blue-400 align-middle leading-none whitespace-nowrap max-w-full min-w-0 truncate";
const HOVER_STYLES = "hover:bg-blue-100 dark:hover:bg-blue-900/50";
const INTERACTIVE_STYLES =
  "cursor-pointer focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-blue-400/60";
const NON_INTERACTIVE_STYLES = "cursor-default";

export function ExcelRangePill({
  range,
  workbookName,
  onClick,
  className,
  variant = "default",
}: ExcelRangePillProps) {
  const displayRange = workbookName ? `${workbookName}!${range}` : range;
  const isInteractive = Boolean(onClick);
  const titleText = isInteractive
    ? `Click to navigate to ${displayRange} in Excel`
    : displayRange;
  const handleClick: React.MouseEventHandler<HTMLButtonElement> = (event) => {
    if (!onClick) return;
    event.preventDefault();
    event.stopPropagation();
    onClick(range, workbookName);
  };

  const minimalStyles =
    "inline-flex items-center font-mono text-[10px] font-medium px-1 py-0.5 border border-blue-200 dark:border-blue-600 text-blue-600 dark:text-blue-400 bg-transparent align-middle leading-none whitespace-nowrap max-w-full min-w-0 truncate";

  if (variant === "minimal") {
    if (isInteractive) {
      return (
        <button
          type="button"
          className={cn(
            minimalStyles,
            INTERACTIVE_STYLES,
            "hover:bg-blue-50 dark:hover:bg-blue-900/30",
            className,
          )}
          onClick={handleClick}
          title={titleText}
          aria-label={titleText}
        >
          {displayRange}
        </button>
      );
    }

    return (
      <span
        className={cn(minimalStyles, NON_INTERACTIVE_STYLES, className)}
        title={titleText}
      >
        {displayRange}
      </span>
    );
  }

  if (isInteractive) {
    return (
      <button
        type="button"
        className={cn(BASE_STYLES, HOVER_STYLES, INTERACTIVE_STYLES, className)}
        onClick={handleClick}
        title={titleText}
        aria-label={titleText}
      >
        {displayRange}
      </button>
    );
  }

  return (
    <span
      className={cn(BASE_STYLES, NON_INTERACTIVE_STYLES, className)}
      title={titleText}
    >
      {displayRange}
    </span>
  );
}

ExcelRangePill.displayName = "ExcelRangePill";
