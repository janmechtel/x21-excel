import { type Dispatch, type SetStateAction, useMemo } from "react";

import { UserMessageBlock } from "./blocks/UserMessageBlock";
import { AssistantMessageBlock } from "./blocks/AssistantMessageBlock";
import { EmptyState } from "./EmptyState";
import {
  findFirstToolFromMessage,
  groupMessagesIntoConversations,
  isConversationEffectivelyReverted,
  shouldShowCollapseIndicator,
} from "@/utils/chat";
import { findSlashCommandById } from "@/lib/slashCommands";
import type { ChatMessage, ToolDecisionData } from "@/types/chat";
import type { UiRequestResponse } from "@/types/uiRequest";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import { ClaudeContentTypes } from "@/types/chat";

interface ChatThreadProps {
  chatHistory: ChatMessage[];
  isStreaming: boolean;
  isViewingHistoryConversation: boolean;
  isRevertingOrApplying: boolean;
  activeStickyMessageId: string | null;
  expandedStickyMessages: Set<string>;
  expandedUserMessages: Set<string>;
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
  toggleStickyMessageExpansion: (id: string) => void;
  toggleUserMessageExpansion: (id: string) => void;
  toggleBlockExpansion: (id: string) => void;
  handleExcelRangeNavigate: (range: string, workbookName?: string) => void;
  handleViewTool: (toolId: string, isAutoView?: boolean) => Promise<void>;
  handleApproveTools: (toolId: string) => Promise<void>;
  handleRejectTool: (toolId: string) => Promise<void>;
  handleApproveAll: (toolId: string) => Promise<void>;
  handleRejectSubmit: () => void;
  handleRejectCancel: () => void;
  setRejectMessage: Dispatch<SetStateAction<string>>;
  handleRevertFromTool: (toolUseId: string, toolName: string) => Promise<void>;
  handleApplyFromTool: (toolUseId: string, toolName: string) => Promise<void>;
  onEditPrompt: (text: string) => void;
  handleScoreMessage: (messageId: string, score: "up" | "down") => void;
  handleCommentSubmit: (messageId: string) => void;
  handleCommentCancel: () => void;
  setCommentText: Dispatch<SetStateAction<string>>;
  revertedConversations: Set<string>;
  onCommandSelect?: (command: SlashCommandDefinition) => void;
  onOpenCommands?: () => void;
  selectedRange: string;
  onOpenSettings: () => void;
  onUiRequestSubmit: (
    toolUseId: string,
    response: UiRequestResponse,
    summary?: string,
  ) => Promise<void>;
}

export function ChatThread({
  chatHistory,
  isStreaming,
  isViewingHistoryConversation,
  isRevertingOrApplying,
  activeStickyMessageId,
  expandedStickyMessages,
  expandedUserMessages,
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
  toggleStickyMessageExpansion,
  toggleUserMessageExpansion,
  toggleBlockExpansion,
  handleExcelRangeNavigate,
  handleViewTool,
  handleApproveTools,
  handleRejectTool,
  handleApproveAll,
  handleRejectSubmit,
  handleRejectCancel,
  setRejectMessage,
  handleRevertFromTool,
  handleApplyFromTool,
  onEditPrompt,
  handleScoreMessage,
  handleCommentSubmit,
  handleCommentCancel,
  setCommentText,
  revertedConversations,
  onCommandSelect,
  onOpenCommands,
  selectedRange,
  onOpenSettings,
  onUiRequestSubmit,
}: ChatThreadProps) {
  const conversations = useMemo(
    () => groupMessagesIntoConversations(chatHistory),
    [chatHistory],
  );

  // Show empty state when there are no conversations
  const isEmpty = conversations.length === 0 && !isStreaming;

  const isConversationEffectivelyRevertedForMessage = (
    userMessage: ChatMessage,
  ) =>
    isConversationEffectivelyReverted(
      chatHistory,
      revertedConversations,
      userMessage,
    );

  const renderChatMessage = (message: ChatMessage) => {
    if (message.role === "user") {
      const slashCommandMeta = findSlashCommandById(message.slashCommandId);
      const firstTool = findFirstToolFromMessage(chatHistory, message);
      const tools = (() => {
        const userMessageIndex = chatHistory.findIndex(
          (msg) => msg.id === message.id,
        );
        const assistantMessage =
          userMessageIndex >= 0 ? chatHistory[userMessageIndex + 1] : null;
        if (
          !assistantMessage ||
          assistantMessage.role !== "assistant" ||
          !assistantMessage.contentBlocks
        ) {
          return [];
        }
        return assistantMessage.contentBlocks
          .filter(
            (block) =>
              block.type === ClaudeContentTypes.TOOL_USE &&
              block.toolUseId &&
              block.toolName,
          )
          .map((block: any) => ({
            toolUseId: block.toolUseId,
            toolName: block.toolName,
          }));
      })();

      return (
        <UserMessageBlock
          key={message.id}
          message={message}
          slashCommandMeta={slashCommandMeta}
          isStreaming={isStreaming}
          isViewingHistoryConversation={isViewingHistoryConversation}
          isRevertingOrApplying={isRevertingOrApplying}
          activeStickyMessageId={activeStickyMessageId}
          expandedStickyMessages={expandedStickyMessages}
          expandedUserMessages={expandedUserMessages}
          firstTool={firstTool}
          tools={tools}
          isConversationEffectivelyReverted={isConversationEffectivelyRevertedForMessage(
            message,
          )}
          shouldShowCollapseIndicator={shouldShowCollapseIndicator(
            message.content || "",
          )}
          toggleStickyMessageExpansion={toggleStickyMessageExpansion}
          toggleUserMessageExpansion={toggleUserMessageExpansion}
          onEditPrompt={onEditPrompt}
          handleRevertFromTool={handleRevertFromTool}
          handleApplyFromTool={handleApplyFromTool}
        />
      );
    }

    return (
      <AssistantMessageBlock
        key={message.id}
        message={message}
        expandedBlocks={expandedBlocks}
        approvedTools={approvedTools}
        rejectedTools={rejectedTools}
        viewedTools={viewedTools}
        erroredTools={erroredTools}
        autoApproveEnabled={autoApproveEnabled}
        toolDecisions={toolDecisions}
        rejectingToolId={rejectingToolId}
        rejectMessage={rejectMessage}
        showInlineComment={showInlineComment}
        commentText={commentText}
        isToolPending={isToolPending}
        getKeyInfo={getKeyInfo}
        toggleBlockExpansion={toggleBlockExpansion}
        handleExcelRangeNavigate={handleExcelRangeNavigate}
        handleViewTool={handleViewTool}
        handleApproveTools={handleApproveTools}
        handleRejectTool={handleRejectTool}
        handleApproveAll={handleApproveAll}
        handleRejectSubmit={handleRejectSubmit}
        handleRejectCancel={handleRejectCancel}
        setRejectMessage={setRejectMessage}
        handleScoreMessage={handleScoreMessage}
        handleCommentSubmit={handleCommentSubmit}
        handleCommentCancel={handleCommentCancel}
        setCommentText={setCommentText}
        selectedRange={selectedRange}
        onOpenSettings={onOpenSettings}
        onUiRequestSubmit={onUiRequestSubmit}
      />
    );
  };

  if (isEmpty && onCommandSelect && onOpenCommands) {
    return (
      <EmptyState
        onCommandSelect={onCommandSelect}
        onOpenCommands={onOpenCommands}
      />
    );
  }

  return (
    <>
      {conversations.map((conversation) => (
        <div
          key={conversation.userMessage.id}
          className="conversation-group"
          data-message-id={conversation.userMessage.id}
        >
          <div className="sticky -top-4 z-50 -mx-6 px-6 pt-4 pb-2 mb-0">
            <div className="max-w-4xl mx-auto">
              {renderChatMessage(conversation.userMessage)}
            </div>
          </div>
          <div className="assistant-responses space-y-3 mb-4 pt-2">
            {conversation.assistantMessages
              .map(renderChatMessage)
              .filter(Boolean)}
          </div>
        </div>
      ))}
    </>
  );
}
