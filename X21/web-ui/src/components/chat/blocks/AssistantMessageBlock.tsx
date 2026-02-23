import { ThumbsDown, ThumbsUp } from "lucide-react";
import type { Dispatch, SetStateAction } from "react";

import { Button } from "@/components/ui/button";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { Textarea } from "@/components/ui/textarea";
import { ToolBlock } from "./ToolBlock";
import { ThinkingBlock } from "./ThinkingBlock";
import { TextBlock } from "./TextBlock";
import type { ChatMessage, ContentBlock, ToolDecisionData } from "@/types/chat";
import type { UiRequestResponse } from "@/types/uiRequest";
import { ClaudeContentTypes, ContentBlockTypes } from "@/types/chat";
import { UiRequestBlock } from "./UiRequestBlock";

interface AssistantMessageBlockProps {
  message: ChatMessage;
  expandedBlocks: Set<string>;
  approvedTools: Set<string>;
  rejectedTools: Set<string>;
  viewedTools: Set<string>;
  erroredTools: Map<string, string>;
  autoApproveEnabled: boolean;
  toolDecisions: Map<string, ToolDecisionData>;
  rejectingToolId: string | null;
  rejectMessage: string;
  showInlineComment: string | null;
  commentText: string;
  isToolPending: (toolId: string) => boolean;
  getKeyInfo: (toolName: string, content: string, blockId: string) => string;
  toggleBlockExpansion: (id: string) => void;
  handleExcelRangeNavigate: (range: string, workbookName?: string) => void;
  handleViewTool: (toolId: string, isAutoView?: boolean) => Promise<void>;
  handleApproveTools: (toolId: string) => Promise<void>;
  handleRejectTool: (toolId: string) => Promise<void>;
  handleApproveAll: (toolId: string) => Promise<void>;
  handleRejectSubmit: () => void;
  handleRejectCancel: () => void;
  setRejectMessage: Dispatch<SetStateAction<string>>;
  handleScoreMessage: (messageId: string, score: "up" | "down") => void;
  handleCommentSubmit: (messageId: string) => void;
  handleCommentCancel: () => void;
  setCommentText: Dispatch<SetStateAction<string>>;
  selectedRange: string;
  onOpenSettings: () => void;
  onUiRequestSubmit: (
    toolUseId: string,
    response: UiRequestResponse,
    summary?: string,
  ) => Promise<void>;
}

export function AssistantMessageBlock({
  message,
  expandedBlocks,
  approvedTools,
  rejectedTools,
  viewedTools,
  erroredTools,
  autoApproveEnabled,
  toolDecisions,
  rejectingToolId,
  rejectMessage,
  showInlineComment,
  commentText,
  isToolPending,
  getKeyInfo,
  toggleBlockExpansion,
  handleExcelRangeNavigate,
  handleViewTool,
  handleApproveTools,
  handleRejectTool,
  handleApproveAll,
  handleRejectSubmit,
  handleRejectCancel,
  setRejectMessage,
  handleScoreMessage,
  handleCommentSubmit,
  handleCommentCancel,
  setCommentText,
  selectedRange,
  onOpenSettings,
  onUiRequestSubmit,
}: AssistantMessageBlockProps) {
  const renderContentBlock = (block: ContentBlock) => {
    if (block.type === ClaudeContentTypes.TOOL_USE) {
      return (
        <ToolBlock
          key={block.id}
          block={block}
          message={message}
          isExpanded={expandedBlocks.has(block.id)}
          onToggle={toggleBlockExpansion}
          approvedTools={approvedTools}
          rejectedTools={rejectedTools}
          viewedTools={viewedTools}
          erroredTools={erroredTools}
          autoApproveEnabled={autoApproveEnabled}
          toolDecisions={toolDecisions}
          isToolPending={isToolPending}
          onViewTool={handleViewTool}
          onApproveTool={handleApproveTools}
          onRejectTool={handleRejectTool}
          onApproveAll={handleApproveAll}
          rejectingToolId={rejectingToolId}
          rejectMessage={rejectMessage}
          onRejectMessageChange={setRejectMessage}
          onRejectSubmit={handleRejectSubmit}
          onRejectCancel={handleRejectCancel}
          extractKeyInfo={getKeyInfo}
          onRangeClick={handleExcelRangeNavigate}
          onOpenSettings={onOpenSettings}
        />
      );
    }

    if (block.type === ContentBlockTypes.UI_REQUEST) {
      return (
        <UiRequestBlock
          key={block.id}
          block={block}
          selectedRange={selectedRange}
          onSubmit={onUiRequestSubmit}
        />
      );
    }

    if (block.type === ContentBlockTypes.THINKING) {
      return (
        <ThinkingBlock
          key={block.id}
          block={block}
          isExpanded={expandedBlocks.has(block.id)}
          onToggle={toggleBlockExpansion}
          onRangeClick={handleExcelRangeNavigate}
        />
      );
    }

    return (
      <TextBlock
        key={block.id}
        block={block}
        onRangeClick={handleExcelRangeNavigate}
      />
    );
  };

  const hasRenderableBlock = message.contentBlocks?.some((block) => {
    if (
      block.type === ClaudeContentTypes.TOOL_USE ||
      block.type === ContentBlockTypes.THINKING
    )
      return true;
    if (block.type === ContentBlockTypes.UI_REQUEST) return true;
    return !!block.content.trim();
  });

  if (
    !message.contentBlocks ||
    message.contentBlocks.length === 0 ||
    !hasRenderableBlock
  ) {
    return null;
  }

  const isMessageComplete = !message.isStreaming;
  const containsUiRequest = message.contentBlocks?.some(
    (block) => block.type === ContentBlockTypes.UI_REQUEST,
  );
  const isCancellationMessage = message.contentBlocks?.some(
    (block) =>
      block.content.includes("Request Cancelled") ||
      block.content.includes("Request was cancelled") ||
      block.content.includes("cancelled by the user"),
  );

  return (
    <div key={message.id} className="max-w-[97.5%]" data-role="assistant">
      <div className="space-y-2">
        {message.contentBlocks.map((block) => renderContentBlock(block))}

        {isMessageComplete && !isCancellationMessage && !containsUiRequest && (
          <div className="flex items-center gap-2 mt-3 pt-2 border-t border-slate-200 dark:border-slate-700">
            <span className="text-xs text-slate-500 dark:text-slate-400">
              Rate this response:
            </span>
            <div className="flex gap-1">
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    onClick={() => handleScoreMessage(message.id, "up")}
                    variant="ghost"
                    size="sm"
                    className={`h-6 w-6 p-0 transition-colors ${
                      message.score === "up"
                        ? "bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400 hover:bg-green-200 dark:hover:bg-green-900/50"
                        : "text-slate-400 dark:text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-700 hover:text-green-600 dark:hover:text-green-400"
                    }`}
                  >
                    <ThumbsUp className="w-3 h-3" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>
                  <p>Good response</p>
                </TooltipContent>
              </Tooltip>

              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    onClick={() => handleScoreMessage(message.id, "down")}
                    variant="ghost"
                    size="sm"
                    className={`h-6 w-6 p-0 transition-colors ${
                      message.score === "down"
                        ? "bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 hover:bg-red-200 dark:hover:bg-red-900/50"
                        : "text-slate-400 dark:text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-700 hover:text-red-600 dark:hover:text-red-400"
                    }`}
                  >
                    <ThumbsDown className="w-3 h-3" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent>
                  <p>Poor response</p>
                </TooltipContent>
              </Tooltip>
            </div>
          </div>
        )}

        {showInlineComment === message.id && (
          <div className="mt-3 pt-2 border-t border-slate-200 dark:border-slate-700">
            <div className="bg-slate-50 dark:bg-slate-800/50 p-3 border border-slate-200 dark:border-slate-700">
              <div className="text-xs text-slate-600 dark:text-slate-400 font-medium mb-2">
                Tell us more about what could be improved (optional)
              </div>
              <Textarea
                value={commentText}
                onChange={(e) => setCommentText(e.target.value)}
                placeholder="Share specific details about what went wrong or what could be better..."
                className="w-full mb-3 resize-none border border-slate-300 dark:border-slate-600 focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 min-h-[60px] px-3 py-2 bg-white dark:bg-slate-900/50 text-slate-900 dark:text-slate-100 placeholder:text-slate-500 dark:placeholder:text-slate-400 text-xs"
                onKeyDown={(e) => {
                  if (e.key === "Enter" && !e.shiftKey) {
                    e.preventDefault();
                    if (commentText.trim()) {
                      handleCommentSubmit(message.id);
                    }
                  }
                }}
              />
              <div className="flex gap-2 justify-end">
                <Button
                  onClick={handleCommentCancel}
                  variant="ghost"
                  size="sm"
                  className="h-6 px-2 text-xs text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700"
                >
                  Close
                </Button>
                <Button
                  onClick={() => handleCommentSubmit(message.id)}
                  variant="default"
                  size="sm"
                  className="h-6 px-2 text-xs bg-slate-700 hover:bg-slate-800 dark:bg-slate-600 dark:hover:bg-slate-500 text-white"
                  disabled={!commentText.trim()}
                >
                  Send Feedback
                </Button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
