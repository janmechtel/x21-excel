import { X } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { MarkdownViewer } from "@/components/markdown/MarkdownViewer";
import { ContentBlock } from "@/types/chat";

interface TextBlockProps {
  block: ContentBlock;
  onRangeClick: (range: string, workbookName?: string) => void;
}

export function TextBlock({ block, onRangeClick }: TextBlockProps) {
  const isError =
    block.content.includes("❌ Error") ||
    block.content.includes("--- STREAM ERROR ---");

  if (isError) {
    return (
      <Alert
        variant="destructive"
        className="border-red-200 bg-red-50 dark:border-red-800 dark:bg-red-950/50"
      >
        <X className="h-4 w-4" />
        <AlertDescription className="text-red-800 dark:text-red-200">
          <MarkdownViewer
            text={block.content
              .replace("❌ Error: ", "")
              .replace("--- STREAM ERROR ---\n", "")}
            onRangeClick={onRangeClick}
            className="text-xs prose-p:my-1 prose-ul:my-1"
          />
        </AlertDescription>
      </Alert>
    );
  }

  return (
    <MarkdownViewer
      text={block.content}
      onRangeClick={onRangeClick}
      className="text-sm text-slate-800 dark:text-slate-200 leading-relaxed prose-p:my-2 prose-ul:my-2"
    />
  );
}
