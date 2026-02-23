import { type Dispatch, type SetStateAction, useCallback } from "react";

import {
  DenoServerConnectionError,
  webSocketChatService,
} from "@/services/webSocketChatService";
import type {
  ChatMessage,
  ContentBlock,
  SendMessageOptions,
} from "@/types/chat";
import { OperationStatusValues, ContentBlockTypes } from "@/types/chat";
import { findSlashCommandById } from "@/lib/slashCommands";

interface UseSendMessageParams {
  prompt: string;
  setPrompt: (val: string) => void;
  activeSlashCommandId: string | null;
  setActiveSlashCommandId: (val: string | null) => void;
  attachedFiles: any[];
  setAttachedFiles: (files: any[]) => void;
  activeTools: Set<string>;
  autoApproveEnabled: boolean;
  resetSlashCommandState: () => void;
  setChatHistory: React.Dispatch<React.SetStateAction<ChatMessage[]>>;
  setIsStreaming: (val: boolean) => void;
  setLoadingState: (val: import("@/types/chat").OperationStatus) => void;
  setRevertedConversations: Dispatch<SetStateAction<Set<string>>>;
  resetToolState: () => void;
  currentAssistantMessageRef: React.MutableRefObject<string | null>;
  blockIndexMap: React.MutableRefObject<Map<number, string>>;
  keyInfoCacheRef: React.MutableRefObject<Map<string, string>>;
  updateCurrentAssistantMessage: (
    updater: (msg: ChatMessage) => ChatMessage,
  ) => void;
  wsUrl: string;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
  isStreaming: boolean;
  createDenoServerErrorContent: () => string;
  approveAllEnabled: boolean;
  setApproveAllEnabled: (val: boolean) => void;
  allowOtherWorkbookReads: boolean;
  setRejectedTools: Dispatch<SetStateAction<Set<string>>>;
  toolGroups: Map<string, string[]>;
  setToolGroups: Dispatch<SetStateAction<Map<string, string[]>>>;
  toolGroupsRef: React.MutableRefObject<Map<string, string[]>>;
  toolGroupDecisions: Map<
    string,
    Map<string, { decision: "approved" | "rejected"; message?: string }>
  >;
  setToolGroupDecisions: Dispatch<
    SetStateAction<
      Map<
        string,
        Map<string, { decision: "approved" | "rejected"; message?: string }>
      >
    >
  >;
  setIsCancelling: (val: boolean) => void;
  setIsViewingHistoryConversation: (val: boolean) => void;
  resetHistoryState: () => void;
  setExpandedBlocks: (val: Set<string>) => void;
  setToolCounter: (val: number) => void;
}

export function useSendMessage(params: UseSendMessageParams) {
  const {
    prompt,
    setPrompt,
    activeSlashCommandId,
    setActiveSlashCommandId,
    attachedFiles,
    setAttachedFiles,
    activeTools,
    autoApproveEnabled,
    resetSlashCommandState,
    setChatHistory,
    setIsStreaming,
    setLoadingState,
    setRevertedConversations,
    resetToolState,
    currentAssistantMessageRef,
    blockIndexMap,
    keyInfoCacheRef,
    updateCurrentAssistantMessage,
    wsUrl,
    textareaRef,
    isStreaming,
    createDenoServerErrorContent,
    allowOtherWorkbookReads,
    approveAllEnabled,
    setApproveAllEnabled,
    setRejectedTools,
    toolGroups,
    setToolGroups,
    toolGroupsRef,
    toolGroupDecisions,
    setToolGroupDecisions,
    setIsCancelling,
    setIsViewingHistoryConversation,
    resetHistoryState,
    setExpandedBlocks,
    setToolCounter,
  } = params;

  const handleSend = useCallback(
    async (options?: SendMessageOptions) => {
      const rawPrompt = options?.overridePrompt ?? prompt;
      const trimmedPrompt = rawPrompt.trim();
      const attachmentSource = options?.attachmentsOverride ?? attachedFiles;
      const slashCommandIdToUse =
        options?.slashCommandId ?? activeSlashCommandId;
      const slashCommand = slashCommandIdToUse
        ? findSlashCommandById(slashCommandIdToUse)
        : undefined;

      if (
        (!slashCommand && !trimmedPrompt && attachmentSource.length === 0) ||
        isStreaming
      )
        return;

      const combinedPrompt = slashCommand
        ? trimmedPrompt
          ? `${slashCommand.prompt.trim()}\n\nUser input:\n${trimmedPrompt}`
          : slashCommand.prompt.trim()
        : rawPrompt;

      const filesToSend = [...attachmentSource];

      const userMessage: ChatMessage = {
        id: `user-${Date.now()}-${Math.random()}`,
        role: "user",
        content: rawPrompt,
        timestamp: Date.now(),
        slashCommandId: slashCommand?.id,
        attachedFiles: filesToSend.length > 0 ? filesToSend : undefined,
      };

      setChatHistory((prev) => [...prev, userMessage]);

      if (options?.nextPromptValue !== undefined) {
        setPrompt(options.nextPromptValue);
      } else {
        setPrompt("");
      }
      setAttachedFiles([]);
      resetSlashCommandState();
      setActiveSlashCommandId(null);

      if (textareaRef.current) {
        textareaRef.current.style.height = "28px";
      }

      setIsStreaming(true);
      setLoadingState(OperationStatusValues.GENERATING_LLM);
      setRevertedConversations(new Set());
      resetToolState();

      const assistantMessage: ChatMessage = {
        id: `assistant-${Date.now()}-${Math.random()}`,
        role: "assistant",
        timestamp: Date.now(),
        contentBlocks: [],
        isStreaming: true,
      };

      setChatHistory((prev) => [...prev, assistantMessage]);
      currentAssistantMessageRef.current = assistantMessage.id;
      blockIndexMap.current.clear();
      keyInfoCacheRef.current.clear();

      try {
        await webSocketChatService.sendMessage({
          prompt: combinedPrompt,
          activeTools: Array.from(activeTools),
          autoApproveEnabled,
          allowOtherWorkbookReads,
          documentsBase64: filesToSend.length > 0 ? filesToSend : undefined,
        });
      } catch (error) {
        setIsStreaming(false);

        let errorContent = "";
        if (error instanceof DenoServerConnectionError) {
          errorContent = createDenoServerErrorContent();
        } else {
          const errorMessage =
            error instanceof Error ? error.message : String(error);
          if (
            errorMessage.includes("WebSocket") ||
            errorMessage.includes("connect") ||
            errorMessage.includes("ECONNREFUSED")
          ) {
            errorContent = `🔌 **Connection Error**\n\nFailed to send message due to connection issues.\n\n**Trying to connect to:** ${wsUrl}\n\nPlease check your connection and try again.\n\n**Error details:** ${errorMessage}`;
          } else {
            errorContent = `❌ **Error Sending Message**\n\nAn unexpected error occurred while sending your message.\n\n**Error details:** ${errorMessage}`;
          }
        }

        const errorBlock: ContentBlock = {
          id: `error-${Date.now()}-${Math.random()}`,
          type: "text",
          content: errorContent,
          isComplete: true,
        };

        updateCurrentAssistantMessage((msg) => ({
          ...msg,
          isStreaming: false,
          contentBlocks: [errorBlock],
        }));

        currentAssistantMessageRef.current = null;
      }
    },
    [
      prompt,
      activeSlashCommandId,
      attachedFiles,
      isStreaming,
      setChatHistory,
      setPrompt,
      setAttachedFiles,
      resetSlashCommandState,
      setActiveSlashCommandId,
      setIsStreaming,
      setLoadingState,
      setRevertedConversations,
      resetToolState,
      currentAssistantMessageRef,
      blockIndexMap,
      keyInfoCacheRef,
      activeTools,
      autoApproveEnabled,
      updateCurrentAssistantMessage,
      wsUrl,
      textareaRef,
      createDenoServerErrorContent,
    ],
  );

  const handleCancel = useCallback(async () => {
    setIsStreaming(false);
    setLoadingState(OperationStatusValues.IDLE);
    setIsCancelling(true);

    try {
      const activeGroups = Array.from(toolGroups.entries());

      if (activeGroups.length > 0) {
        for (const [groupId, toolIds] of activeGroups) {
          const groupDecisions = toolGroupDecisions.get(groupId) || new Map();

          for (const toolId of toolIds) {
            if (groupDecisions.has(toolId)) {
              continue;
            }
            try {
              groupDecisions.set(toolId, {
                decision: "rejected",
                message: "Cancelled by user",
              });
              setRejectedTools((prev) => new Set([...prev, toolId]));
            } catch (error) {
              console.error(`Failed to process tool ${toolId}:`, error);
            }
          }

          setToolGroupDecisions((prev) =>
            new Map(prev).set(groupId, groupDecisions),
          );
          try {
            const toolResponses = toolIds.map((toolId) => ({
              toolId,
              decision: "rejected" as const,
              userMessage: "Cancelled by user",
            }));
            await webSocketChatService.sendToolPermissionResponse(
              groupId,
              toolResponses,
            );
          } catch (error) {
            console.error(
              `Failed to send group rejection for ${groupId}:`,
              error,
            );
          }
        }

        setToolGroups(new Map());
        toolGroupsRef.current = new Map();
        setToolGroupDecisions(new Map());
      }

      await webSocketChatService.cancelCurrentRequest();

      const currentTime = Date.now();
      setChatHistory((prevMessages: ChatMessage[]) =>
        prevMessages.map((msg: ChatMessage) => {
          if (msg.role === "assistant" && msg.isStreaming) {
            // Close any incomplete thinking blocks
            const updatedBlocks =
              msg.contentBlocks?.map((block: ContentBlock) => {
                if (
                  block.type === ContentBlockTypes.THINKING &&
                  !block.isComplete
                ) {
                  return { ...block, isComplete: true, endTime: currentTime };
                }
                return block;
              }) || [];

            // Add cancellation message
            const cancellationBlock: ContentBlock = {
              id: `cancelled-${Date.now()}-${Math.random()}`,
              type: "text",
              content:
                "⏹️ **Request Cancelled**\n\nThe request was cancelled by the user.",
              isComplete: true,
            };

            return {
              ...msg,
              isStreaming: false,
              contentBlocks: [...updatedBlocks, cancellationBlock],
            };
          }
          return msg;
        }),
      );
    } catch (error) {
      console.error("Cancel request error:", error);
    }

    if (approveAllEnabled) {
      setApproveAllEnabled(false);
      webSocketChatService.setApproveAll(false);
    }

    currentAssistantMessageRef.current = null;

    // Return focus to textarea after canceling
    setTimeout(() => {
      textareaRef.current?.focus();
    }, 0);
  }, [
    setIsStreaming,
    setLoadingState,
    setIsCancelling,
    toolGroups,
    toolGroupDecisions,
    setRejectedTools,
    setToolGroupDecisions,
    setToolGroups,
    toolGroupsRef,
    setChatHistory,
    approveAllEnabled,
    setApproveAllEnabled,
    currentAssistantMessageRef,
    textareaRef,
  ]);

  const handleRestartState = useCallback(async () => {
    try {
      setIsCancelling(true);
      await webSocketChatService.cancelCurrentRequest();
      await new Promise((resolve) => setTimeout(resolve, 300));
    } catch {
      // ignore
    }

    try {
      setIsStreaming(false);
      setLoadingState(OperationStatusValues.IDLE);
      setIsCancelling(false);
      setChatHistory([]);
      setIsViewingHistoryConversation(false);
      resetHistoryState();
      setExpandedBlocks(new Set());
      setToolCounter(0);
      setRevertedConversations(new Set());
      resetToolState();

      const success = await webSocketChatService.restartState();
      if (!success) {
        console.error(
          "Failed to send restart state request - WebSocket not connected",
        );
      }
      setTimeout(() => {
        textareaRef.current?.focus();
      }, 0);
    } catch (error) {
      console.error("Error restarting state:", error);
    }
    setTimeout(() => {
      textareaRef.current?.focus();
    }, 0);
  }, [
    setIsStreaming,
    setLoadingState,
    setIsCancelling,
    setChatHistory,
    setIsViewingHistoryConversation,
    resetHistoryState,
    setExpandedBlocks,
    setToolCounter,
    setRevertedConversations,
    resetToolState,
    textareaRef,
  ]);

  return { handleSend, handleCancel, handleRestartState };
}
