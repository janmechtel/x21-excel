import { cn } from "@/lib/utils";

interface ExcelSheetMentionPillProps {
  sheetName: string;
  onClick?: (sheetName: string) => void;
  className?: string;
}

const BASE_STYLES =
  "inline-flex items-center font-mono text-[10px] font-medium px-1 py-[1px] h-3 border border-blue-200 dark:border-blue-600 text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/30 leading-[18px]";

export function ExcelSheetMentionPill({
  sheetName,
  onClick,
  className,
}: ExcelSheetMentionPillProps) {
  const cleaned = sheetName.replace(/^'+|'+$/g, "").replace(/''/g, "'");

  return (
    <div
      className={cn(
        BASE_STYLES,
        onClick
          ? "cursor-pointer hover:bg-amber-100 dark:hover:bg-amber-900/50"
          : "cursor-default",
        className,
      )}
      title={`Click to navigate to sheet ${cleaned}`}
      onClick={(event) => {
        event.preventDefault();
        event.stopPropagation();
        onClick?.(cleaned);
      }}
    >
      {cleaned}
    </div>
  );
}

ExcelSheetMentionPill.displayName = "ExcelSheetMentionPill";
