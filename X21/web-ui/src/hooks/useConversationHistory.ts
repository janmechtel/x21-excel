import { useState } from "react";

import type { ChatMessage, ToolDecisionData } from "@/types/chat";
import {
  listRecentConversationsForFile,
  loadConversationForFile,
} from "@/services/conversationHistoryService";
import type { RecentChatItem } from "@/services/historyService";
import { webViewBridge } from "@/services/webViewBridge";

interface Params {
  chatHistoryLength: number;
  setChatHistory: React.Dispatch<React.SetStateAction<ChatMessage[]>>;
  responseEndRef: React.RefObject<HTMLDivElement>;
  resetToolState: () => void;
  setRejectedTools: React.Dispatch<React.SetStateAction<Set<string>>>;
  setApprovedTools: React.Dispatch<React.SetStateAction<Set<string>>>;
  setToolDecisions: React.Dispatch<
    React.SetStateAction<Map<string, ToolDecisionData>>
  >;
}

export function useConversationHistory({
  setChatHistory,
  responseEndRef,
  resetToolState,
  setRejectedTools,
  setApprovedTools,
  setToolDecisions,
}: Params) {
  const [currentWorkbookKey, setCurrentWorkbookKey] = useState<string | null>(
    null,
  );
  const [showHistory, setShowHistory] = useState(false);
  const [recentChats, setRecentChats] = useState<RecentChatItem[]>([]);
  const [isHistoryLoading, setIsHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState<string | null>(null);
  const [activeConversationId, setActiveConversationId] = useState<
    string | null
  >(null);
  const [isViewingHistoryConversation, setIsViewingHistoryConversation] =
    useState(false);

  const applyHistoricalToolDecisions = (
    toolDecisions?: Map<string, ToolDecisionData>,
  ) => {
    resetToolState();

    const safeDecisions = toolDecisions ?? new Map();
    const rejectedIds = Array.from(safeDecisions.entries())
      .filter(([, data]) => data.decision === "rejected")
      .map(([toolId]) => toolId);

    setRejectedTools(new Set(rejectedIds));
    setApprovedTools(new Set());
    setToolDecisions(new Map(safeDecisions));
  };

  // Auto-restore removed - chat always starts fresh
  // Users can manually load conversations from the history panel

  const toggleHistoryPanel = async () => {
    setHistoryError(null);

    const willOpen = !showHistory;

    // CRITICAL: Open panel IMMEDIATELY (like tools panel) - don't wait for data
    setShowHistory(willOpen);

    // Load data in the background AFTER opening the panel
    if (willOpen) {
      // Use setTimeout to ensure panel renders first, then load data
      setTimeout(async () => {
        setIsHistoryLoading(true);

        try {
          const workbookKey =
            currentWorkbookKey ?? (await webViewBridge.getWorkbookPath());

          if (!workbookKey) {
            setHistoryError("No workbook path found for this Excel workbook.");
            setRecentChats([]);
            return;
          }

          setCurrentWorkbookKey(workbookKey);
          const items = await listRecentConversationsForFile(workbookKey, 20);
          setRecentChats(items);
        } catch (error) {
          console.error("Failed to load recent chats", error);
          setHistoryError("Failed to load history. Please try again.");
        } finally {
          setIsHistoryLoading(false);
        }
      }, 0);
    }
  };

  const handleSelectConversation = async (conversationId: string) => {
    try {
      setIsHistoryLoading(true);
      setHistoryError(null);

      const workbookKey =
        currentWorkbookKey ?? (await webViewBridge.getWorkbookPath());
      if (!workbookKey) {
        setHistoryError("No workbook path found for this Excel workbook.");
        return;
      }

      setCurrentWorkbookKey(workbookKey);
      const { messages, toolDecisions } = await loadConversationForFile(
        workbookKey,
        conversationId,
      );
      console.log("[History] Loaded conversation from history", {
        conversationId,
        workbookKey,
        messageCount: messages.length,
      });
      applyHistoricalToolDecisions(toolDecisions);
      setChatHistory(messages as ChatMessage[]);
      setActiveConversationId(conversationId);
      setIsViewingHistoryConversation(true);
      setShowHistory(false);

      setTimeout(() => {
        responseEndRef.current?.scrollIntoView({ behavior: "smooth" });
      }, 0);
    } catch (error) {
      console.error("Failed to load conversation from history", error);
      setHistoryError("Failed to load this conversation. Please try again.");
    } finally {
      setIsHistoryLoading(false);
    }
  };

  const handleHistoryConversationLoaded = (
    messages: ChatMessage[],
    conversationId: string,
    toolDecisions?: Map<string, ToolDecisionData>,
  ) => {
    setIsViewingHistoryConversation(true);
    applyHistoricalToolDecisions(toolDecisions);
    setChatHistory(messages as ChatMessage[]);
    setActiveConversationId(conversationId);
    setTimeout(() => {
      responseEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }, 0);
  };

  const handleHistoryConversationSelected = (conversationId: string) => {
    // Match tools panel implementation: simple synchronous state updates
    // React will batch these together efficiently, just like tools panel
    setShowHistory(false);
    setIsViewingHistoryConversation(true);
    setActiveConversationId(conversationId);
    // Clear chat history to show loading state
    setChatHistory([]);
  };

  const handleHistoryConversationLoadError = () => {
    // Reset state if loading fails
    setIsViewingHistoryConversation(false);
    setActiveConversationId(null);
  };

  const resetHistoryState = () => {
    setShowHistory(false);
    setActiveConversationId(null);
    setIsViewingHistoryConversation(false);
    setRecentChats([]);
  };

  return {
    currentWorkbookKey,
    setCurrentWorkbookKey,
    showHistory,
    setShowHistory,
    recentChats,
    isHistoryLoading,
    historyError,
    activeConversationId,
    setActiveConversationId,
    isViewingHistoryConversation,
    setIsViewingHistoryConversation,
    toggleHistoryPanel,
    handleSelectConversation,
    handleHistoryConversationLoaded,
    handleHistoryConversationSelected,
    handleHistoryConversationLoadError,
    resetHistoryState,
  };
}
