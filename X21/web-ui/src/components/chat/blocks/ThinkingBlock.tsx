import { Brain, ChevronDown, ChevronRight } from "lucide-react";

import { MarkdownViewer } from "@/components/markdown/MarkdownViewer";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/components/ui/collapsible";
import { ContentBlock } from "@/types/chat";
import { formatThinkingDuration } from "@/utils/chat";

import { IconContainer, WaveAnimation } from "./Common";

interface ThinkingBlockProps {
  block: ContentBlock;
  isExpanded: boolean;
  onToggle: (id: string) => void;
  onRangeClick: (range: string, workbookName?: string) => void;
}

export function ThinkingBlock({
  block,
  isExpanded,
  onToggle,
  onRangeClick,
}: ThinkingBlockProps) {
  const hasContent = block.content.trim().length > 0;
  const hasDuration = typeof block.startTime === "number";
  const duration = hasDuration
    ? formatThinkingDuration(block.startTime as number, block.endTime)
    : null;

  return (
    <div className="group py-0">
      <Collapsible
        open={isExpanded}
        onOpenChange={() => hasContent && onToggle(block.id)}
      >
        <div className="flex items-center justify-between py-0.5 px-1 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
          <div className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-400 flex-1">
            <IconContainer>
              <Brain className="w-3 h-3 text-purple-500 dark:text-purple-400" />
            </IconContainer>
            <span className="font-medium">
              {block.isComplete ? "Thought" : "Thinking"}
            </span>
            {!block.isComplete && (
              <WaveAnimation className="text-purple-500 dark:text-purple-400" />
            )}
            {block.isComplete && duration && (
              <span className="text-slate-400 dark:text-slate-500">
                for {duration}
              </span>
            )}
          </div>

          {hasContent && (
            <CollapsibleTrigger asChild>
              <button className="opacity-0 group-hover:opacity-100 transition-opacity p-0.5 hover:bg-slate-200 dark:hover:bg-slate-700 rounded">
                {isExpanded ? (
                  <ChevronDown className="w-2.5 h-2.5 text-slate-400" />
                ) : (
                  <ChevronRight className="w-2.5 h-2.5 text-slate-400" />
                )}
              </button>
            </CollapsibleTrigger>
          )}
        </div>

        {hasContent && (
          <CollapsibleContent className="ml-4 mt-0.5">
            <MarkdownViewer
              text={block.content}
              onRangeClick={onRangeClick}
              className="text-xs text-slate-600 dark:text-slate-400 italic p-1.5 bg-slate-50 dark:bg-slate-800/50 prose-p:my-1 prose-ul:my-1"
            />
          </CollapsibleContent>
        )}
      </Collapsible>
    </div>
  );
}
