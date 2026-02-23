import { Info } from "lucide-react";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";

export interface InfoTooltipProps {
  content: string;
  className?: string;
}

export function InfoTooltip({ content, className = "" }: InfoTooltipProps) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <Info
          className={`w-4 h-4 text-slate-400 hover:text-slate-600 dark:hover:text-slate-300 cursor-help flex-shrink-0 ${className}`}
        />
      </TooltipTrigger>
      <TooltipContent className="max-w-[250px] z-[10000]">
        <p className="text-xs">{content}</p>
      </TooltipContent>
    </Tooltip>
  );
}
