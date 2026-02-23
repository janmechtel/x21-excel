import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { ChevronDown } from "lucide-react";
import { webViewBridge } from "./services/webViewBridge";
import {
  deleteChangelogEntry,
  fetchChangelogSummaries,
  updateChangelogEntry,
} from "@/services/changelogApi";
import { useFocusOnMount } from "./hooks/useFocusOnMount";
import { useSlashCommandState } from "./hooks/useSlashCommandState";
import { useAttachments } from "./hooks/useAttachments";
import { useToolState } from "./hooks/useToolState";
import { useWebSocketStream } from "./hooks/useWebSocketStream";
import { useChatViewport } from "./hooks/useChatViewport";
import { useConversationHistory } from "./hooks/useConversationHistory";
import { useFeedback } from "./hooks/useFeedback";
import { useSheetMentionState } from "./hooks/useSheetMentionState";
import { useSheetMentionInsertion } from "./hooks/useSheetMentionInsertion";
import { useSendMessage } from "./hooks/useSendMessage";
import { useToolActions } from "./hooks/useToolActions";
import { useSelectedRange } from "./hooks/useSelectedRange";
import { useInsertSelectedRange } from "./hooks/useInsertSelectedRange";
import { useAvailableTools } from "./hooks/useAvailableTools";
import posthog from "./posthog";
import { initPosthogClientLogging } from "./services/posthogClientLogging";
import { getErrorDisplayMessage } from "@/utils/errorMessages";
import { TooltipProvider } from "@/components/ui/tooltip";
import { Button } from "@/components/ui/button";
import { ProtectedRoute } from "./components/auth";
import { useAuth } from "./contexts/AuthContext";
import { ExcelContextProvider } from "@/contexts/ExcelContext";
import { webSocketChatService } from "./services/webSocketChatService";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import { ChatThread } from "@/components/chat/ChatThread";
import { ChatInputArea } from "@/components/chat/ChatInputArea";
import { NewChatDialog } from "@/components/chat/NewChatDialog";
import { AppHeader } from "@/components/chat/AppHeader";
import { SettingsConfigDialog } from "@/components/chat/SettingsConfigDialog";
import { ConversationHistoryPanel } from "./components/history/ConversationHistoryPanel";
import { ActivityPanel } from "@/components/changelog/ActivityPanel";
import { DragDropLayer } from "@/components/chat/DragDropLayer";
import { ToolTogglePanel } from "@/components/chat/ToolTogglePanel";
import { StatusIndicator } from "@/components/chat/StatusIndicator";
import type {
  ActivityLog,
  ActivitySummary,
  ChangeSummaryPayload,
  ChatMessage,
  ToolDecisionData,
} from "@/types/chat";
import {
  OperationStatusValues,
  ContentBlockTypes,
  ToolNames,
} from "@/types/chat";
import type { UiRequestResponse } from "@/types/uiRequest";
import { createPromptKeyDownHandler } from "@/utils/promptHandlers";
import { useSheetMentionClose } from "./hooks/useSheetMentionClose";
import { useSlashCommandClose } from "./hooks/useSlashCommandClose";
import { usePromptHistoryNavigation } from "./hooks/usePromptHistoryNavigation";

/**
 * Main chat application shell.
 * @returns App layout for the chat UI.
 */
function App() {
  const { user, loading } = useAuth();
  const pendingPosthogUserIdRef = useRef<string | null>(null);
  useEffect(() => {
    initPosthogClientLogging();
    // TEMP: log test warnings/errors to verify PostHog wiring (remove after validation).
    // console.warn("WEB TEMP TEST: web warning log");
  }, []);

  const [prompt, setPrompt] = useState("");
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);
  const [expandedBlocks, setExpandedBlocks] = useState<Set<string>>(new Set());
  const [isCancelling, setIsCancelling] = useState(false);
  const [showNewChatDialog, setShowNewChatDialog] = useState(false);
  const [, setHasShownConnectionError] = useState(false);
  const [toolCounter, setToolCounter] = useState(0); // Add tool counter
  const [activeTools, setActiveTools] = useState<Set<string>>(new Set()); // Will be initialized when availableTools is loaded
  const [showTools, setShowTools] = useState(false); // Toggle tools overlay visibility
  const [toolHighlightIndex, setToolHighlightIndex] = useState(0);
  const responseEndRef = useRef<HTMLDivElement>(null);
  const chatContainerRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const [expandedStickyMessages, setExpandedStickyMessages] = useState<
    Set<string>
  >(new Set());
  const [allowOtherWorkbookReads, setAllowOtherWorkbookReads] = useState(false);
  const [showConfigDialog, setShowConfigDialog] = useState(false);
  const [expandedUserMessages, setExpandedUserMessages] = useState<Set<string>>(
    new Set(),
  );
  const {
    slashCommandState,
    slashCommandMatches,
    slashCommandHighlightIndex,
    setSlashCommandHighlightIndex,
    activeSlashCommandId,
    setActiveSlashCommandId,
    activeSlashCommand,
    highlightedSlashCommand,
    isSlashCommandActive,
    promptPlaceholder,
    updateSlashCommandTrigger,
    resetSlashCommandState,
    clearActiveSlashCommand,
    openSlashCommandPalette,
  } = useSlashCommandState();
  const {
    sheetMentionState,
    mentionMatches,
    mentionHighlightIndex,
    setMentionHighlightIndex,
    mentionQuery,
    isMentionActive,
    highlightedSheetName,
    sheetNames,
    currentWorkbookName,
    openWorkbooks,
    updateMentionTrigger,
    resetMentionState,
    openSheetMentionPalette,
    otherWorkbookSheets,
    isRequestingOtherWorkbooks,
  } = useSheetMentionState();

  const { handleSheetMentionSelect } = useSheetMentionInsertion({
    prompt,
    setPrompt,
    sheetMentionState,
    resetMentionState,
    textareaRef,
  });
  const slashCommandQuery = slashCommandState?.query ?? "";
  const slashCommandCount = slashCommandMatches.length;
  const openSlashCommandPaletteHandler = () => openSlashCommandPalette(prompt);
  const showCommands = slashCommandState !== null;
  const isSheetMentionActive = isMentionActive && !isSlashCommandActive;
  const openSheetMentionPaletteHandler = (cursorPosition?: number) => {
    resetSlashCommandState();
    clearActiveSlashCommand();
    const caret =
      cursorPosition ?? textareaRef.current?.selectionStart ?? prompt.length;
    openSheetMentionPalette(prompt, caret);
  };
  const closeSheetMentionPalette = useSheetMentionClose({
    prompt,
    setPrompt,
    textareaRef,
    resetMentionState,
    sheetMentionState,
  });
  const closeSlashCommandPalette = useSlashCommandClose({
    prompt,
    setPrompt,
    textareaRef,
    resetSlashCommandState,
    slashCommandState,
  });

  // Track which conversation turns have been reverted by user message ID
  const [revertedConversations, setRevertedConversations] = useState<
    Set<string>
  >(new Set());

  // Add loading state for revert/apply operations
  const [isRevertingOrApplying, setIsRevertingOrApplying] = useState(false);

  // Enhanced status system
  const [operationStatus, setOperationStatus] = useState<
    import("@/types/chat").OperationStatus
  >(OperationStatusValues.IDLE);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const [operationProgress, setOperationProgress] = useState<{
    current: number;
    total: number;
    unit?: string;
  } | null>(null);

  // Token tracking
  const [totalTokens, setTotalTokens] = useState<number>(0);
  const [inputTokens, setInputTokens] = useState<number>(0);
  const [outputTokens, setOutputTokens] = useState<number>(0);

  /**
   * Resets the status indicator to idle.
   * Used for manual user-initiated resets (ESC key, New Chat, Cancel button).
   * Provides immediate UI feedback while the server processes the request.
   */
  const resetStatusIndicator = useCallback(() => {
    setOperationStatus(OperationStatusValues.IDLE);
    setStatusMessage(null);
    setOperationProgress(null);
  }, [setOperationStatus, setStatusMessage, setOperationProgress]);

  // Activity logs
  const [activityLogs, setActivityLogs] = useState<ActivityLog[]>([]);
  const [showLogs, setShowLogs] = useState(false);
  const [isGeneratingChangelog, setIsGeneratingChangelog] = useState(false);
  const [isLoadingLogs, setIsLoadingLogs] = useState(false);

  const loadingState: import("@/types/chat").OperationStatus = operationStatus;

  const setLoadingState = useCallback(
    (state: import("@/types/chat").OperationStatus) => {
      setOperationStatus(state);

      // When returning to idle, also clear any lingering message/progress.
      if (state === OperationStatusValues.IDLE) {
        setStatusMessage(null);
        setOperationProgress(null);
      }
    },
    [setOperationStatus, setStatusMessage, setOperationProgress],
  );

  const legacyLoadingState: import("@/types/chat").OperationStatus =
    operationStatus === OperationStatusValues.WAITING_APPROVAL
      ? OperationStatusValues.WAITING_APPROVAL
      : operationStatus === OperationStatusValues.IDLE
      ? OperationStatusValues.IDLE
      : OperationStatusValues.GENERATING_LLM;

  // prompt history navigation hook is initialized after conversation history

  const {
    approvedTools,
    setApprovedTools,
    rejectedTools,
    setRejectedTools,
    viewedTools,
    setViewedTools,
    erroredTools,
    setErroredTools,
    autoApproveEnabled,
    setAutoApproveEnabled,
    approveAllEnabled,
    setApproveAllEnabled,
    toolGroups,
    setToolGroups,
    toolGroupsRef,
    toolGroupDecisions,
    setToolGroupDecisions,
    toolDecisions,
    setToolDecisions,
    findToolGroup,
    isToolPending,
    getPendingToolCount,
    resetToolState,
  } = useToolState();
  const pendingToolCount = getPendingToolCount();
  const hasPendingUiRequest = useMemo(
    () =>
      chatHistory.some((message) =>
        (message.contentBlocks || []).some(
          (block) =>
            block.type === ContentBlockTypes.UI_REQUEST &&
            !block.uiRequestResponse,
        ),
      ),
    [chatHistory],
  );
  const shouldShowWaitingForUser =
    operationStatus !== OperationStatusValues.ERROR &&
    (pendingToolCount > 0 || hasPendingUiRequest);
  const statusIndicatorStatus = shouldShowWaitingForUser
    ? OperationStatusValues.WAITING_APPROVAL
    : operationStatus;
  const statusIndicatorMessage = shouldShowWaitingForUser
    ? "Waiting for user"
    : statusMessage;

  const { activeStickyMessageId, handleScroll, scrollToBottom, isNearBottom } =
    useChatViewport({
      chatHistory,
      isStreaming,
      getPendingToolCount,
      toolGroups,
      toolGroupDecisions,
      chatContainerRef,
      responseEndRef,
    });

  const {
    currentWorkbookKey,
    setCurrentWorkbookKey,
    showHistory,
    setShowHistory,
    activeConversationId,
    setIsViewingHistoryConversation,
    isViewingHistoryConversation,
    toggleHistoryPanel,
    handleHistoryConversationLoaded,
    handleHistoryConversationSelected,
    handleHistoryConversationLoadError,
    resetHistoryState,
  } = useConversationHistory({
    chatHistoryLength: chatHistory.length,
    setChatHistory,
    responseEndRef,
    resetToolState,
    setRejectedTools,
    setApprovedTools,
    setToolDecisions,
  });

  const {
    promptHistoryEntries,
    promptHistoryIndex,
    isPromptHistoryLoading,
    isNavigatingHistory,
    setIsNavigatingHistory,
    setPromptFromHistory,
    loadPromptHistoryEntries,
    setPromptHistoryIndex,
    handlePromptChange,
    resetNavigationAfterSend,
  } = usePromptHistoryNavigation({
    setPrompt,
    updateSlashCommandTrigger,
    updateMentionTrigger,
    currentWorkbookKey,
    setCurrentWorkbookKey,
  });

  const {
    showInlineComment,
    commentText,
    setCommentText,
    handleScoreMessage,
    handleCommentSubmit,
    handleCommentCancel,
  } = useFeedback({ chatHistory, setChatHistory, scrollToBottom });

  const {
    handleApproveTools,
    handleRejectTool,
    handleRejectSubmit,
    handleRejectCancel,
    handleApproveAll,
    handleViewTool,
    handleRevertFromTool,
    handleApplyFromTool,
    handleToolDecision,
    rejectingToolId,
    rejectMessage,
    setRejectMessage,
  } = useToolActions({
    chatHistory,
    setChatHistory,
    setRevertedConversations,
    setIsRevertingOrApplying,
    setApprovedTools,
    setRejectedTools,
    viewedTools,
    setViewedTools,
    setToolGroups,
    toolGroupsRef,
    toolGroupDecisions,
    setToolGroupDecisions,
    setApproveAllEnabled,
    findToolGroup,
    setToolDecisions,
  });

  // Handle token usage updates from WebSocket stream
  const handleUsageUpdate = (inputTokens: number, outputTokens: number) => {
    setTotalTokens((prev) => prev + inputTokens + outputTokens);
  };

  const handleUiRequestSubmit = useCallback(
    async (
      toolUseId: string,
      response: UiRequestResponse,
      summary?: string,
    ) => {
      setLoadingState(OperationStatusValues.GENERATING_LLM);
      setIsStreaming(true);
      try {
        await webSocketChatService.sendToolResult(toolUseId, response);
        setChatHistory((prev) =>
          prev.map((msg) => ({
            ...msg,
            contentBlocks: (msg.contentBlocks || []).map((block) =>
              block.toolUseId === toolUseId &&
              block.type === ContentBlockTypes.UI_REQUEST
                ? {
                    ...block,
                    uiRequestResponse: response,
                    uiRequestSummary: summary,
                  }
                : block,
            ),
          })),
        );
      } catch (error) {
        setIsStreaming(false);
        setLoadingState(OperationStatusValues.IDLE);
        throw error;
      }
    },
    [setChatHistory, setIsStreaming, setLoadingState],
  );

  const {
    getKeyInfo,
    updateCurrentAssistantMessage,
    currentAssistantMessageRef,
    blockIndexMap,
    keyInfoCacheRef,
    attachWebSocketHandlers,
  } = useWebSocketStream({
    toolCounter,
    setToolCounter,
    setChatHistory,
    setIsStreaming,
    setLoadingState, // Now uses the backward-compatible wrapper
    isCancelling,
    setIsCancelling,
    setHasShownConnectionError,
    setApprovedTools,
    setViewedTools,
    setErroredTools,
    setToolGroups,
    toolGroupsRef,
    setToolGroupDecisions,
    setApproveAllEnabled,
    findToolGroup,
    handleViewTool,
    handleToolDecision,
    createDenoServerErrorContent,
    getErrorDisplayMessage,
    approveAllEnabled,
    setShowTools,
    onUsageUpdate: handleUsageUpdate,
    // New status system handlers
    setOperationStatus,
    setStatusMessage,
    setOperationProgress,
    setInputTokens,
    setOutputTokens,
    setTotalTokens,
  });

  const {
    attachedFiles,
    setAttachedFiles,
    isConvertingFile,
    isDragOver,
    fileInputRef,
    handlePaste,
    handleDragOver,
    handleDragEnter,
    handleDragLeave,
    handleDrop,
    handleFileInputChange,
    removeAttachedFile,
    clearAllAttachedFiles,
    formatFileSize,
    openFileDialog,
  } = useAttachments();

  // WebSocket state
  const [wsConnected, setWsConnected] = useState(false);
  const chatHistoryLengthRef = useRef(0);
  const [wsUrl, setWsUrl] = useState("ws://localhost:8085"); // Default fallback

  // Fetch available tools from backend
  const { availableTools } = useAvailableTools();

  const { handleSend, handleCancel, handleRestartState } = useSendMessage({
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
  });

  // Function to update WebSocket URL for display
  const updateWebSocketUrl = useCallback(async () => {
    try {
      const url = await webViewBridge.getWebSocketUrl();
      setWsUrl(url);
    } catch (error) {
      console.warn("Failed to get WebSocket URL for display:", error);
    }
  }, []);

  // Set document title with branch name
  useEffect(() => {
    const branchName = import.meta.env.VITE_BRANCH_NAME;
    if (branchName && branchName !== "unknown") {
      document.title = `X21 Chat [${branchName}]`;
    } else {
      document.title = "X21 Chat";
    }
  }, []);

  // Initialize activeTools when availableTools is loaded
  useEffect(() => {
    if (availableTools.length === 0) return;

    // Keep only tools that the backend reports, and reinitialize if none remain
    setActiveTools((prev) => {
      const allowed = new Set<string>(
        availableTools
          .filter(
            (tool) =>
              tool.id !== ToolNames.VBA_CREATE &&
              tool.id !== ToolNames.VBA_UPDATE,
          )
          .map((tool) => tool.id),
      );

      if (prev.size === 0) {
        return new Set(allowed);
      }

      const filtered = new Set([...prev].filter((id) => allowed.has(id)));
      return filtered.size > 0 ? filtered : new Set(allowed);
    });
  }, [availableTools]);

  // Utility function to parse summary text into ActivityLog format
  const parseSummaryToActivityLog = useCallback(
    (
      summaryText: string,
      timestamp: number,
      logId: string,
      sheetsAffected?: number,
      comparisonType?: "self" | "external",
      comparisonFilePath?: string | null,
    ): ActivityLog | null => {
      if (!summaryText || !summaryText.trim()) {
        return null;
      }

      // Parse checkbox items with hierarchy support
      const lines = summaryText.split("\n");
      const checkboxItems: Array<{
        line: string;
        level: number;
        text: string;
      }> = [];

      for (const line of lines) {
        // Match checkbox pattern with optional leading spaces for indentation
        const match = line.match(/^(\s*)-\s*\[\s*[xX ]?\s*\]\s*(.+)$/);
        if (match) {
          const [, leadingSpaces, text] = match;
          // Calculate level: 0 spaces = level 0, 2 spaces = level 1, 4 spaces = level 2, etc.
          const level = Math.floor(leadingSpaces.length / 2);
          checkboxItems.push({ line, level, text: text.trim() });
        }
      }

      // Build hierarchical structure
      const buildHierarchy = (
        items: Array<{ line: string; level: number; text: string }>,
        startIndex: number,
        currentLevel: number,
      ): { summaries: ActivitySummary[]; nextIndex: number } => {
        const summaries: ActivitySummary[] = [];
        let i = startIndex;

        while (i < items.length) {
          const item = items[i];

          // If this item is at a deeper level, it belongs to the previous parent
          if (item.level > currentLevel) {
            break;
          }

          // If this item is at a shallower level, we've moved up the hierarchy
          if (item.level < currentLevel) {
            break;
          }

          // This item is at our current level
          const summary: ActivitySummary = {
            type: "workbook_action" as const,
            count: 1,
            description: item.text,
            details: [],
            firstOccurrence: new Date(timestamp).toISOString(),
            lastOccurrence: new Date(timestamp).toISOString(),
            level: item.level,
            children: [],
          };

          // Check if next items are children
          i++;
          if (i < items.length && items[i].level > currentLevel) {
            const childResult = buildHierarchy(items, i, items[i].level);
            summary.children = childResult.summaries;
            i = childResult.nextIndex;
          }

          summaries.push(summary);
        }

        return { summaries, nextIndex: i };
      };

      // If we found checkbox items, create hierarchical summaries
      const summaries =
        checkboxItems.length > 0
          ? buildHierarchy(checkboxItems, 0, 0).summaries
          : [
              {
                type: "workbook_action" as const,
                count: sheetsAffected || 1,
                description: summaryText,
                details: [],
                firstOccurrence: new Date(timestamp).toISOString(),
                lastOccurrence: new Date(timestamp).toISOString(),
                level: 0,
              },
            ];

      return {
        id: logId,
        timestamp,
        data: {
          timeRange: {
            start: new Date(timestamp).toISOString(),
            end: new Date(timestamp).toISOString(),
          },
          totalEvents: summaries.length,
          summaries,
        },
        rawSummaryText: summaryText,
        comparisonType,
        comparisonFilePath,
        comparisonFileName: comparisonFilePath
          ? comparisonFilePath.split(/[\\/]/).pop() ?? null
          : null,
      };
    },
    [],
  );

  // Load saved summaries from the database
  const loadSavedSummaries = useCallback(async () => {
    setIsLoadingLogs(true);
    try {
      const workbookName = await webViewBridge.getWorkbookName();
      if (!workbookName) {
        console.log("No workbook name available, skipping summary load");
        setIsLoadingLogs(false);
        return;
      }

      const summaries = await fetchChangelogSummaries(workbookName, 50);

      // Convert summaries to ActivityLog format
      const newLogs: ActivityLog[] = summaries
        .map((summary) => {
          return parseSummaryToActivityLog(
            summary.summaryText,
            summary.createdAt,
            summary.id,
            summary.sheetsAffected,
            summary.comparisonType,
            summary.comparisonFilePath,
          );
        })
        .filter((log: ActivityLog | null): log is ActivityLog => log !== null)
        .sort((a: ActivityLog, b: ActivityLog) => b.timestamp - a.timestamp); // Most recent first

      if (newLogs.length > 0) {
        // Merge with existing logs, avoiding duplicates by ID
        setActivityLogs((prev) => {
          const existingIds = new Set(prev.map((log) => log.id));
          const uniqueNewLogs = newLogs.filter(
            (log) => !existingIds.has(log.id),
          );
          // Combine and sort by timestamp (most recent first)
          const combined = [...prev, ...uniqueNewLogs].sort(
            (a, b) => b.timestamp - a.timestamp,
          );
          return combined;
        });
        console.log(`Loaded ${newLogs.length} saved summaries`);
      }
      setIsLoadingLogs(false);
    } catch (error) {
      console.error("Error loading saved summaries:", error);
      setIsLoadingLogs(false);
    }
  }, [parseSummaryToActivityLog]);

  // Handle workbook change summary notifications
  const handleChangeSummary = useCallback(
    (payload: ChangeSummaryPayload) => {
      console.log("Change summary received in App:", payload);

      // Always clear the generating state when summary message is received
      setIsGeneratingChangelog(false);

      // Store workbook change summary in logs instead of chat (only if there's content)
      if (payload?.summary && payload.summary.trim()) {
        const summaryId =
          typeof payload.id === "string" && payload.id.trim().length > 0
            ? payload.id
            : `change-summary-${Date.now()}-${Math.random()}`;
        const summaryTimestamp =
          typeof payload.timestamp === "number"
            ? payload.timestamp
            : Date.now();

        const newLog = parseSummaryToActivityLog(
          payload.summary,
          summaryTimestamp,
          summaryId,
          payload.sheetsAffected,
          payload.comparisonType,
          payload.comparisonFilePath,
        );

        if (newLog) {
          setActivityLogs((prev) => {
            // Avoid duplicates when the same summary is also loaded from /api/workbook-summaries
            if (prev.some((log) => log.id === summaryId)) {
              return prev;
            }
            return [newLog, ...prev];
          });
        }
      }
    },
    [parseSummaryToActivityLog],
  );

  useEffect(() => {
    chatHistoryLengthRef.current = chatHistory.length;
  }, [chatHistory.length]);

  const getChatHistoryLength = useCallback(
    () => chatHistoryLengthRef.current,
    [],
  );

  // Initialize WebSocket service
  useEffect(() => {
    // Get the current WebSocket URL for display
    updateWebSocketUrl();
    attachWebSocketHandlers(
      setWsConnected,
      updateWebSocketUrl,
      getChatHistoryLength,
      handleChangeSummary,
    );

    return () => {
      webSocketChatService.disconnect();
    };
  }, [
    attachWebSocketHandlers,
    getChatHistoryLength,
    handleChangeSummary,
    updateWebSocketUrl,
  ]); // Include handleChangeSummary in dependencies

  const toggleBlockExpansion = (blockId: string) => {
    setExpandedBlocks((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(blockId)) {
        newSet.delete(blockId);
      } else {
        newSet.add(blockId);
      }
      return newSet;
    });
  };

  const handleOpenSettings = () => {
    setShowConfigDialog(true);
  };

  const handleCloseSettings = () => {
    setShowConfigDialog(false);
  };

  const toggleStickyMessageExpansion = (messageId: string) => {
    setExpandedStickyMessages((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(messageId)) {
        newSet.delete(messageId);
      } else {
        newSet.add(messageId);
      }
      return newSet;
    });
  };

  const toggleUserMessageExpansion = (messageId: string) => {
    setExpandedUserMessages((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(messageId)) {
        newSet.delete(messageId);
      } else {
        newSet.add(messageId);
      }
      return newSet;
    });
  };

  // Handle tool auto-approval notification from WebSocket service

  // Handle tool error from WebSocket

  /**
   * Centralized function to create Deno server connection error content.
   * @returns Markdown string for connection failure.
   */
  function createDenoServerErrorContent(): string {
    return `🔌 **Cannot Connect to Deno Server**\n\n**Please restart Excel to restore the AI features.**`;
  }
  const { selectedRange, handleExcelRangeNavigate, handleExcelSheetNavigate } =
    useSelectedRange({
      isStreaming,
      textareaRef,
      showTools,
      showHistory,
      showCommands,
    });

  useEffect(() => {
    const handleUserIdReady = (data: { userId: string }) => {
      if (!posthog.__loaded) return;

      if (loading) {
        pendingPosthogUserIdRef.current = data.userId;
        return;
      }

      if (user?.email) {
        console.log("PostHog user already identified via email:", user.email);
        return;
      }

      console.log("PostHog identifying user:", data.userId);
      posthog.identify(data.userId);
    };

    webViewBridge.on("userIdReady", handleUserIdReady);

    return () => {
      webViewBridge.off("userIdReady");
    };
  }, [loading, user?.email]);

  useEffect(() => {
    if (!posthog.__loaded || loading || user?.email) return;

    const pendingUserId = pendingPosthogUserIdRef.current;
    if (!pendingUserId) return;

    console.log("PostHog identifying user:", pendingUserId);
    posthog.identify(pendingUserId);
    pendingPosthogUserIdRef.current = null;
  }, [loading, user?.email]);

  // Auto-focus text input when TaskPane first opens (component mounts)
  // Uses retry logic to handle async WebView2 rendering delays
  useFocusOnMount(textareaRef, { maxAttempts: 3, retryDelay: 50 });

  const toggleTool = (toolId: string) => {
    setActiveTools((prev) => {
      const newSet = new Set(prev);
      if (newSet.has(toolId)) {
        newSet.delete(toolId);
      } else {
        newSet.add(toolId);
      }
      return newSet;
    });
  };

  const handleToolSelect = (toolId: string) => {
    toggleTool(toolId);
  };

  const handleToggleTools = useCallback(() => {
    setShowTools((prev) => {
      if (prev) {
        // Reset highlight index when closing
        setToolHighlightIndex(0);
        // Return focus to textarea
        setTimeout(() => {
          textareaRef.current?.focus();
        }, 0);
      }
      return !prev;
    });
  }, [setToolHighlightIndex, setShowTools]);

  const handleCloseTools = () => {
    setShowTools(false);
    setToolHighlightIndex(0);
    // Return focus to textarea
    setTimeout(() => {
      textareaRef.current?.focus();
    }, 0);
  };

  // Global keyboard shortcuts for overlays
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      const isModifierPressed = event.ctrlKey || event.metaKey;
      const isShiftPressed = event.shiftKey;
      const key = event.key.toLowerCase();

      // ESC - Cancel streaming request if active
      if (key === "escape" && isStreaming) {
        event.preventDefault();
        resetStatusIndicator();
        void handleCancel();
        return;
      }

      // ESC - Close history or overlays when not streaming
      if (key === "escape") {
        if (showHistory) {
          event.preventDefault();
          setShowHistory(false);
          return;
        }
        if (isSlashCommandActive || isSheetMentionActive) {
          event.preventDefault();
          closeSlashCommandPalette();
          closeSheetMentionPalette();
          clearActiveSlashCommand();
          return;
        }
      }

      // Ctrl+Shift+T (or Cmd+Shift+T on Mac) - Toggle tools overlay
      if (isModifierPressed && isShiftPressed && key === "t") {
        event.preventDefault();
        handleToggleTools();
      }

      // Ctrl+H or Ctrl+Shift+H (or Cmd equivalents) - Toggle history panel
      if (isModifierPressed && key === "h") {
        event.preventDefault();
        void toggleHistoryPanel();
      }

      // Note: Commands palette is opened by typing "/" in the chat input
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => {
      window.removeEventListener("keydown", handleKeyDown);
    };
  }, [
    handleToggleTools,
    toggleHistoryPanel,
    isStreaming,
    handleCancel,
    resetStatusIndicator,
    showHistory,
    closeSlashCommandPalette,
    closeSheetMentionPalette,
    clearActiveSlashCommand,
    isSlashCommandActive,
    isSheetMentionActive,
    setShowHistory,
  ]);

  const handleSendWithHistoryReset = () => {
    resetNavigationAfterSend();
    void handleSend();
  };
  const handleInsertSelectedRange = useInsertSelectedRange({
    selectedRange,
    isStreaming,
    loadingState: legacyLoadingState,
    getPendingToolCount,
    setPrompt,
    updateSlashCommandTrigger,
    textareaRef,
  });

  const handleSlashCommandSelect = (command: SlashCommandDefinition) => {
    if (
      isStreaming ||
      legacyLoadingState !== OperationStatusValues.IDLE ||
      getPendingToolCount() > 0
    ) {
      return;
    }
    resetMentionState();
    console.info("[SlashCommand] selected", {
      id: command.id,
      title: command.title,
      requiresInput: command.requiresInput,
      prompt: command.prompt?.slice(0, 200),
      defaultInput: command.defaultInput,
    });

    const before = slashCommandState
      ? prompt.slice(0, slashCommandState.startIndex)
      : prompt;
    const after = slashCommandState
      ? prompt.slice(slashCommandState.endIndex)
      : "";
    const updatedPrompt = command.defaultInput
      ? `${before}${command.defaultInput}${after}`
      : `${before}${after}`;
    const commandRequiresInput = command.requiresInput !== false;

    resetSlashCommandState();
    setPrompt(updatedPrompt);

    if (!commandRequiresInput) {
      setPromptHistoryIndex(-1);
      void handleSend({
        overridePrompt: command.prompt,
        slashCommandId: command.id,
        nextPromptValue: updatedPrompt,
      });
      return;
    }

    setActiveSlashCommandId(command.id);

    requestAnimationFrame(() => {
      if (!textareaRef.current) return;
      const caretPosition = command.defaultInput
        ? before.length + command.defaultInput.length
        : before.length;
      textareaRef.current.selectionStart = caretPosition;
      textareaRef.current.selectionEnd = caretPosition;
      textareaRef.current.focus();
    });
  };

  const handleNewChatClick = () => {
    // Only show confirmation dialog if there's an ongoing request
    const hasOngoingRequest =
      isStreaming ||
      getPendingToolCount() > 0 ||
      loadingState !== OperationStatusValues.IDLE;

    if (hasOngoingRequest) {
      setShowNewChatDialog(true);
    } else {
      // No ongoing request, directly start new chat
      setShowNewChatDialog(false);
      resetStatusIndicator();
      setTotalTokens(0); // Reset token counter
      handleRestartState();
    }
  };

  const handleConfigSaved = () => {
    setShowConfigDialog(false);
    // Start a new chat after config change to avoid mixing different provider contexts
    handleNewChatClick();
  };

  const toggleLogsPanel = () => {
    setShowLogs((prev) => {
      const willOpen = !prev;
      // Load summaries from database when opening the panel
      if (willOpen) {
        loadSavedSummaries();
      }
      return willOpen;
    });
  };

  const handleDeleteLog = useCallback(async (logId: string) => {
    const ok = await deleteChangelogEntry(logId);
    if (!ok) {
      return;
    }
    setActivityLogs((prev) => prev.filter((log) => log.id !== logId));
  }, []);

  const handleEditLog = useCallback(
    async (logId: string, newRawSummaryText: string) => {
      const ok = await updateChangelogEntry(logId, newRawSummaryText);
      if (!ok) {
        return;
      }

      setActivityLogs((prev) => {
        const existing = prev.find((log) => log.id === logId);
        const timestamp = existing?.timestamp ?? Date.now();
        const updated = parseSummaryToActivityLog(
          newRawSummaryText,
          timestamp,
          logId,
        );

        if (!updated) {
          return prev;
        }

        return prev.map((log) => (log.id === logId ? updated : log));
      });
    },
    [parseSummaryToActivityLog],
  );

  const handleGenerateSummary = async (comparisonFilePath?: string) => {
    setIsGeneratingChangelog(true);

    try {
      console.log(
        "Generating changelog...",
        comparisonFilePath ? `with comparison file: ${comparisonFilePath}` : "",
      );
      const result = await webViewBridge.send<{
        success: boolean;
        message?: string;
        error?: string;
      }>(
        "generateChangelog",
        comparisonFilePath ? { comparisonFilePath } : {},
        true,
      );

      if (result?.success) {
        console.log("Changelog generation started:", result.message);
      } else {
        console.error("Failed to generate changelog:", result?.error);
        setIsGeneratingChangelog(false); // Clear generating state on error
      }
    } catch (error) {
      console.error("Error generating changelog:", error);
      setIsGeneratingChangelog(false); // Clear generating state on error
    }
  };

  const handlePromptKeyDown = createPromptKeyDownHandler({
    isSlashCommandActive,
    slashCommandCount,
    setSlashCommandHighlightIndex,
    highlightedSlashCommand,
    handleSlashCommandSelect,
    resetSlashCommandState: closeSlashCommandPalette,
    activeSlashCommandId,
    clearActiveSlashCommand,
    isStreaming,
    prompt,
    setPrompt,
    setPromptFromHistory,
    attachedFilesLength: attachedFiles.length,
    handleSend: handleSendWithHistoryReset,
    hasActiveSlashCommand: activeSlashCommand !== null,
    isSheetMentionActive,
    sheetMentionCount: mentionMatches.length,
    setSheetMentionHighlightIndex: setMentionHighlightIndex,
    highlightedSheetName,
    handleSheetMentionSelect,
    resetSheetMentionState: closeSheetMentionPalette,
    handleNewChat: handleNewChatClick,
    promptHistoryEntries,
    promptHistoryIndex,
    isPromptHistoryLoading,
    loadPromptHistory: loadPromptHistoryEntries,
    setPromptHistoryIndex,
    isNavigatingHistory,
    setIsNavigatingHistory,
  });

  const conversationHistoryProps = {
    open: showHistory,
    onClose: () => setShowHistory(false),
    activeConversationId,
    onEditPrompt: (promptText: string) => {
      setPrompt(promptText);
      setShowHistory(false);
      setTimeout(() => {
        if (textareaRef.current) {
          textareaRef.current.focus();
          // Set cursor to the end of the text
          const length = promptText.length;
          textareaRef.current.setSelectionRange(length, length);
          // Scroll to the bottom to make cursor visible
          textareaRef.current.scrollTop = textareaRef.current.scrollHeight;
        }
      }, 0);
    },
    onConversationSelected: (conversationId: string) =>
      handleHistoryConversationSelected(conversationId),
    onConversationLoaded: (
      messages: ChatMessage[],
      conversationId: string,
      toolDecisionsParam?: Map<string, ToolDecisionData>,
    ) =>
      handleHistoryConversationLoaded(
        messages as ChatMessage[],
        conversationId,
        toolDecisionsParam,
      ),
    onConversationLoadError: () => handleHistoryConversationLoadError(),
    textareaRef,
  };

  const chatThreadProps = {
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
    onEditPrompt: (text: string) => {
      setPrompt(text);
      requestAnimationFrame(() => {
        if (textareaRef.current) {
          textareaRef.current.focus();
          // Set cursor to the end of the text
          const length = text.length;
          textareaRef.current.setSelectionRange(length, length);
          // Scroll to the bottom to make cursor visible
          textareaRef.current.scrollTop = textareaRef.current.scrollHeight;
        }
      });
    },
    handleScoreMessage,
    handleCommentSubmit,
    handleCommentCancel,
    setCommentText,
    revertedConversations,
    onCommandSelect: handleSlashCommandSelect,
    onOpenCommands: openSlashCommandPaletteHandler,
    selectedRange,
    onOpenSettings: handleOpenSettings,
    onUiRequestSubmit: handleUiRequestSubmit,
  };

  const chatInputProps = {
    prompt,
    promptPlaceholder,
    activeSlashCommand,
    isSlashCommandActive,
    slashCommandMatches,
    slashCommandHighlightIndex,
    slashCommandQuery,
    setSlashCommandHighlightIndex,
    isSheetMentionActive,
    sheetMentionMatches: mentionMatches,
    sheetMentionHighlightIndex: mentionHighlightIndex,
    sheetMentionQuery: mentionQuery,
    setSheetMentionHighlightIndex: setMentionHighlightIndex,
    handleSheetMentionSelect,
    currentSheetNames: sheetNames,
    currentWorkbookName,
    openWorkbooks,
    resetSheetMentionState: closeSheetMentionPalette,
    openSheetMentionPaletteHandler,
    otherWorkbookSheets,
    isRequestingOtherWorkbooks,
    clearActiveSlashCommand,
    resetSlashCommandState: closeSlashCommandPalette,
    openSlashCommandPaletteHandler,
    handleSlashCommandSelect,
    handlePromptChange,
    handlePromptKeyDown,
    handleSend: handleSendWithHistoryReset,
    handleCancel: () => {
      resetStatusIndicator();
      void handleCancel();
    },
    textareaRef,
    isStreaming,
    loadingState,
    getPendingToolCount,
    attachedFiles,
    isDragOver,
    isConvertingFile,
    clearAllAttachedFiles,
    removeAttachedFile,
    formatFileSize,
    handleDragOver,
    handleDragEnter,
    handleDragLeave,
    handleDrop,
    handlePaste,
    fileInputRef,
    handleFileInputChange,
    openFileDialog,
    autoApproveEnabled,
    setAutoApproveEnabled,
    onToggleTools: handleToggleTools,
    activeToolsCount: activeTools.size,
    selectedRange,
    onInsertSelectedRange: handleInsertSelectedRange,
    handleRangeNavigate: handleExcelRangeNavigate,
    handleSheetNavigate: handleExcelSheetNavigate,
    totalTokens,
    inputTokens,
    outputTokens,
    wsConnected,
    wsUrl,
    hasMessages: chatHistory.length > 0,
  };

  return (
    <TooltipProvider>
      <ExcelContextProvider>
        <ProtectedRoute>
          <div className="relative h-screen flex flex-col px-1.5">
            <AppHeader
              onToggleHistory={() => void toggleHistoryPanel()}
              onToggleLogs={toggleLogsPanel}
              onNewChat={handleNewChatClick}
              onOpenSettings={handleOpenSettings}
            />

            <SettingsConfigDialog
              open={showConfigDialog}
              onCancel={handleCloseSettings}
              onSave={handleConfigSaved}
              allowOtherWorkbookReads={allowOtherWorkbookReads}
              onAllowOtherWorkbookReadsChange={setAllowOtherWorkbookReads}
            />

            {/* Conversation history panel */}
            <ConversationHistoryPanel {...conversationHistoryProps} />

            {/* Activity logs panel */}
            <ActivityPanel
              open={showLogs}
              onClose={() => setShowLogs(false)}
              logs={activityLogs}
              onRangeClick={handleExcelRangeNavigate}
              textareaRef={textareaRef}
              isGenerating={isGeneratingChangelog}
              isLoading={isLoadingLogs}
              onGenerateSummary={handleGenerateSummary}
              onDeleteLog={handleDeleteLog}
              onEditLog={handleEditLog}
            />
            {/* Chat messages area */}
            <DragDropLayer
              isDragOver={isDragOver}
              onDragOver={handleDragOver}
              onDragEnter={handleDragEnter}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className="relative flex-1 overflow-y-auto overflow-x-hidden my-2"
              ref={chatContainerRef}
              onScroll={handleScroll}
            >
              <div className="space-y-1.5 max-w-4xl mx-auto">
                <ChatThread {...chatThreadProps} />
                <div ref={responseEndRef} />
              </div>
            </DragDropLayer>

            <div className="relative">
              {!isNearBottom && chatHistory.length > 0 && (
                // Offset keeps the button just above the status panel.
                <div className="absolute -top-12 left-1/2 z-50 -translate-x-1/2">
                  <Button
                    type="button"
                    variant="secondary"
                    size="icon"
                    className="rounded-full shadow-md hover:bg-blue-50 dark:hover:bg-blue-950/30"
                    aria-label="Scroll to latest message"
                    onClick={() => scrollToBottom("smooth")}
                  >
                    <ChevronDown className="h-4 w-4" />
                  </Button>
                </div>
              )}
              <StatusIndicator
                status={statusIndicatorStatus}
                message={statusIndicatorMessage}
                progress={operationProgress}
              />
            </div>
            <ToolTogglePanel
              isOpen={showTools}
              activeTools={activeTools}
              highlightedIndex={toolHighlightIndex}
              onSelect={handleToolSelect}
              onHighlight={setToolHighlightIndex}
              onClose={handleCloseTools}
              onOpenSettings={handleOpenSettings}
              textareaRef={textareaRef}
              availableTools={availableTools}
            />
            <ChatInputArea {...chatInputProps} />
          </div>

          <NewChatDialog
            open={showNewChatDialog}
            isStreaming={isStreaming}
            onCancel={() => setShowNewChatDialog(false)}
            onConfirm={() => {
              setShowNewChatDialog(false);
              resetStatusIndicator();
              setTotalTokens(0); // Reset token counter
              void handleRestartState();
            }}
          />
        </ProtectedRoute>
      </ExcelContextProvider>
    </TooltipProvider>
  );
}

export default App;
