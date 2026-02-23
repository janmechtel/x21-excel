import { Command, Paperclip, Pencil, RotateCcw } from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { Card, CardContent } from "@/components/ui/card";
import type { ChatMessage } from "@/types/chat";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import { ToolNames } from "@/types/chat";

interface UserMessageBlockProps {
  message: ChatMessage;
  slashCommandMeta?: SlashCommandDefinition | null;
  isStreaming: boolean;
  isViewingHistoryConversation: boolean;
  isRevertingOrApplying: boolean;
  activeStickyMessageId: string | null;
  expandedStickyMessages: Set<string>;
  expandedUserMessages: Set<string>;
  firstTool?: { toolUseId: string; toolName: string } | null;
  tools?: { toolUseId: string; toolName: string }[];
  isConversationEffectivelyReverted: boolean;
  shouldShowCollapseIndicator: boolean;
  toggleStickyMessageExpansion: (id: string) => void;
  toggleUserMessageExpansion: (id: string) => void;
  onEditPrompt: (text: string) => void;
  handleRevertFromTool: (toolUseId: string, toolName: string) => Promise<void>;
  handleApplyFromTool: (toolUseId: string, toolName: string) => Promise<void>;
}

export function UserMessageBlock({
  message,
  slashCommandMeta,
  isStreaming,
  isViewingHistoryConversation,
  isRevertingOrApplying,
  activeStickyMessageId,
  expandedStickyMessages,
  expandedUserMessages,
  firstTool,
  tools,
  isConversationEffectivelyReverted,
  shouldShowCollapseIndicator,
  toggleStickyMessageExpansion,
  toggleUserMessageExpansion,
  onEditPrompt,
  handleRevertFromTool,
  handleApplyFromTool,
}: UserMessageBlockProps) {
  const hasWriteFormatBatch =
    firstTool?.toolName === ToolNames.WRITE_FORMAT_BATCH ||
    (tools || []).some((t) => t.toolName === ToolNames.WRITE_FORMAT_BATCH);
  const hasMergeFiles =
    firstTool?.toolName === ToolNames.MERGE_FILES ||
    (tools || []).some((t) => t.toolName === ToolNames.MERGE_FILES);
  const hasAttachments = (message.attachedFiles?.length ?? 0) > 0;
  const trimmedContent = (message.content ?? "").trim();
  const hasMessageText =
    trimmedContent.length > 0 && trimmedContent !== "(attachment)";
  const isAttachmentOnly = hasAttachments && !hasMessageText;

  return (
    <div key={message.id} className="w-full" data-role="user">
      <Card className="bg-blue-50 dark:bg-blue-950/30 border border-blue-200 dark:border-blue-800/50 shadow-sm mr-2">
        <CardContent className="px-3 py-2 relative">
          {message.attachedFiles && (
            <div className="flex flex-wrap gap-0.5 mb-1">
              {message.attachedFiles.map((file, index) => (
                <div
                  key={index}
                  className="flex items-center gap-0.5 bg-blue-100 dark:bg-blue-900/30 px-1.5 py-0.5 text-[10px] border border-blue-200 dark:border-blue-700"
                >
                  <Paperclip className="w-2 h-2 text-blue-600 dark:text-blue-400 flex-shrink-0" />
                  <span
                    className="text-blue-800 dark:text-blue-200 truncate max-w-[60px]"
                    title={file.name}
                  >
                    {file.name}
                  </span>
                </div>
              ))}
            </div>
          )}

          {slashCommandMeta && (
            <div className="flex items-center gap-1 text-[10px] text-blue-700 dark:text-blue-300 font-semibold mb-1 uppercase tracking-wide">
              <Command className="w-3 h-3" />
              <span>/{slashCommandMeta.title}</span>
            </div>
          )}

          <div className="absolute top-2 right-2 flex gap-0.5 z-10">
            {!isStreaming && firstTool && (
              <>
                {!isConversationEffectivelyReverted && (
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <span>
                        <Button
                          onClick={() => {
                            if (isViewingHistoryConversation) return;
                            handleRevertFromTool(
                              firstTool.toolUseId,
                              firstTool.toolName,
                            );
                          }}
                          variant="ghost"
                          size="sm"
                          className="h-auto w-auto p-0.5 text-slate-400 dark:text-slate-500 hover:text-red-600 dark:hover:text-red-400 hover:bg-slate-200/50 dark:hover:bg-slate-700/50 transition-all duration-200 flex flex-col items-center gap-0 disabled:cursor-not-allowed"
                          disabled={
                            isRevertingOrApplying ||
                            isViewingHistoryConversation ||
                            hasWriteFormatBatch ||
                            hasMergeFiles
                          }
                        >
                          <RotateCcw className="w-1.5 h-1.5" />
                          <span className="text-[7px] leading-none">
                            Revert
                          </span>
                        </Button>
                      </span>
                    </TooltipTrigger>
                    <TooltipContent>
                      <p>
                        {isViewingHistoryConversation
                          ? "We don't support revert in history view yet"
                          : hasMergeFiles
                          ? "Revert is not available because merge files was used in this turn"
                          : hasWriteFormatBatch
                          ? "Revert is not available because batch formatting was used in this turn"
                          : "Revert changes made by this tool"}
                      </p>
                    </TooltipContent>
                  </Tooltip>
                )}

                {isConversationEffectivelyReverted && (
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <span>
                        <Button
                          onClick={() => {
                            if (isViewingHistoryConversation) return;
                            handleApplyFromTool(
                              firstTool.toolUseId,
                              firstTool.toolName,
                            );
                          }}
                          variant="ghost"
                          size="sm"
                          className="h-auto w-auto p-0.5 text-slate-400 dark:text-slate-500 hover:text-green-600 dark:hover:text-green-400 hover:bg-slate-200/50 dark:hover:bg-slate-700/50 transition-all duration-200 flex flex-col items-center gap-0 disabled:cursor-not-allowed"
                          disabled={
                            isRevertingOrApplying ||
                            isViewingHistoryConversation
                          }
                        >
                          <svg
                            className="w-1.5 h-1.5"
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M5 13l4 4L19 7"
                            />
                          </svg>
                          <span className="text-[7px] leading-none">Apply</span>
                        </Button>
                      </span>
                    </TooltipTrigger>
                    <TooltipContent>
                      <p>
                        {isViewingHistoryConversation
                          ? "We don't support apply in history view yet"
                          : "Apply changes from this tool again"}
                      </p>
                    </TooltipContent>
                  </Tooltip>
                )}
              </>
            )}
            <Tooltip>
              <TooltipTrigger asChild>
                <span>
                  <Button
                    onClick={() => {
                      if (isAttachmentOnly) return;
                      onEditPrompt(message.content || "");
                    }}
                    variant="ghost"
                    size="sm"
                    className="h-auto w-auto p-0.5 text-slate-400 dark:text-slate-500 hover:text-blue-600 dark:hover:text-blue-400 hover:bg-slate-200/50 dark:hover:bg-slate-700/50 transition-all duration-200 flex flex-col items-center gap-0 disabled:cursor-not-allowed"
                    title={
                      isAttachmentOnly
                        ? "Edit is not available for attachment-only prompts."
                        : "Retry this prompt"
                    }
                    disabled={isAttachmentOnly}
                  >
                    <Pencil className="w-1.5 h-1.5" />
                    <span className="text-[7px] leading-none">Edit</span>
                  </Button>
                </span>
              </TooltipTrigger>
              {isAttachmentOnly && (
                <TooltipContent>
                  <p>Edit is not available for attachment-only prompts.</p>
                </TooltipContent>
              )}
            </Tooltip>
          </div>

          <div className="flex items-start gap-2">
            <div className="text-slate-800 dark:text-slate-200 text-sm flex-1 pr-10">
              {activeStickyMessageId === message.id ? (
                <div
                  className="cursor-pointer hover:bg-slate-100/50 dark:hover:bg-slate-800/50 p-1 -m-1 transition-colors"
                  onClick={() => toggleStickyMessageExpansion(message.id)}
                  title="Click to expand/collapse"
                >
                  <div
                    className="whitespace-pre-wrap text-sm leading-relaxed"
                    style={
                      !expandedStickyMessages.has(message.id)
                        ? {
                            display: "-webkit-box",
                            WebkitLineClamp: 3,
                            WebkitBoxOrient: "vertical",
                            overflow: "hidden",
                          }
                        : {}
                    }
                  >
                    {message.content || ""}
                  </div>
                </div>
              ) : shouldShowCollapseIndicator ? (
                <div
                  className="cursor-pointer hover:bg-slate-100/50 dark:hover:bg-slate-800/50 p-1 -m-1 transition-colors"
                  onClick={() => toggleUserMessageExpansion(message.id)}
                  title="Click to expand/collapse"
                >
                  <div
                    className="whitespace-pre-wrap text-sm leading-relaxed"
                    style={
                      !expandedUserMessages.has(message.id)
                        ? {
                            display: "-webkit-box",
                            WebkitLineClamp: 3,
                            WebkitBoxOrient: "vertical",
                            overflow: "hidden",
                          }
                        : {}
                    }
                  >
                    {message.content || ""}
                  </div>
                </div>
              ) : (
                <div className="whitespace-pre-wrap text-sm leading-relaxed">
                  {message.content || ""}
                </div>
              )}
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
