import { useCallback, useEffect, useRef } from "react";

import {
  DenoServerConnectionError,
  webSocketChatService,
} from "@/services/webSocketChatService";
import type { ChatMessage } from "@/types/chat";
import {
  OperationStatusValues,
  WebSocketMessageTypes,
  ClaudeContentTypes,
  ClaudeEventTypes,
  ContentBlockTypes,
} from "@/types/chat";
import { formatKeyInfo } from "@/utils/toolDisplay";

export interface UseWebSocketStreamParams {
  toolCounter: number;
  setToolCounter: (n: number) => void;
  setChatHistory: React.Dispatch<React.SetStateAction<ChatMessage[]>>;
  setIsStreaming: (v: boolean) => void;
  setLoadingState: (v: import("@/types/chat").OperationStatus) => void;
  isCancelling: boolean;
  setIsCancelling: (v: boolean) => void;
  setHasShownConnectionError: (v: boolean) => void;
  setApprovedTools: React.Dispatch<React.SetStateAction<Set<string>>>;
  setViewedTools: React.Dispatch<React.SetStateAction<Set<string>>>;
  setErroredTools: React.Dispatch<React.SetStateAction<Map<string, string>>>;
  setToolGroups: React.Dispatch<React.SetStateAction<Map<string, string[]>>>;
  toolGroupsRef: React.MutableRefObject<Map<string, string[]>>;
  setToolGroupDecisions: React.Dispatch<
    React.SetStateAction<
      Map<
        string,
        Map<string, { decision: "approved" | "rejected"; message?: string }>
      >
    >
  >;
  setApproveAllEnabled: (v: boolean) => void;
  findToolGroup: (toolId: string) => string | null;
  handleViewTool: (toolId: string, isAutoView?: boolean) => Promise<void>;
  handleToolDecision: (
    toolId: string,
    decision: "approved" | "rejected",
    message?: string,
  ) => Promise<void>;
  createDenoServerErrorContent: () => string;
  getErrorDisplayMessage: (errorType: string) => {
    title: string;
    description: string;
  };
  approveAllEnabled: boolean;
  setShowTools: (val: boolean) => void;
  onUsageUpdate?: (inputTokens: number, outputTokens: number) => void;
  // New enhanced status system
  setOperationStatus?: (status: import("@/types/chat").OperationStatus) => void;
  setStatusMessage?: (message: string | null) => void;
  setOperationProgress?: (
    progress: { current: number; total: number; unit?: string } | null,
  ) => void;
  setInputTokens?: (tokens: number) => void;
  setOutputTokens?: (tokens: number) => void;
  setTotalTokens?: (tokens: number) => void;
}

export function useWebSocketStream({
  toolCounter,
  setToolCounter,
  setChatHistory,
  setIsStreaming,
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
  onUsageUpdate,
  // New enhanced status system
  setOperationStatus,
  setStatusMessage,
  setOperationProgress,
  setInputTokens,
  setOutputTokens,
  setTotalTokens,
}: UseWebSocketStreamParams) {
  const currentAssistantMessageRef = useRef<string | null>(null);
  const blockIndexMap = useRef<Map<number, string>>(new Map());
  const keyInfoCacheRef = useRef<Map<string, string>>(new Map());
  const isCancellingRef = useRef(isCancelling);
  const handleStreamCompleteRef = useRef<(data: any) => void>(() => {});
  const handleStreamDeltaRef = useRef<(data: any) => void>(() => {});
  const handleStreamErrorMessageRef = useRef<(error: any) => void>(() => {});
  const handleConnectionLostRef = useRef<
    (hasShownConnectionError: boolean, chatHistoryLength: number) => void
  >(() => {});

  useEffect(() => {
    isCancellingRef.current = isCancelling;
  }, [isCancelling]);

  // Store the updater function in a ref so handlers can always access the latest version
  const updateCurrentAssistantMessageRef = useRef<
    (updater: (msg: ChatMessage) => ChatMessage) => void
  >(() => {});

  // Helper function to update the current assistant message
  const updateCurrentAssistantMessage = (
    updater: (msg: ChatMessage) => ChatMessage,
  ) => {
    if (!currentAssistantMessageRef.current) return;
    setChatHistory((prev) =>
      prev.map((msg) =>
        msg.id === currentAssistantMessageRef.current ? updater(msg) : msg,
      ),
    );
  };

  // Update the ref to point to the latest version of the function
  updateCurrentAssistantMessageRef.current = updateCurrentAssistantMessage;

  const getKeyInfo = (toolName: string, content: string, blockId: string) => {
    if (!content.trim()) {
      return keyInfoCacheRef.current.get(blockId) || "";
    }
    try {
      const params = JSON.parse(content);
      const keyInfo = formatKeyInfo(toolName, params);
      if (keyInfo) {
        keyInfoCacheRef.current.set(blockId, keyInfo);
      }
      return keyInfo;
    } catch {
      try {
        const partialJson = content.includes("{")
          ? `${content}}`
          : `{${content}}`;
        const params = JSON.parse(partialJson);
        const keyInfo = formatKeyInfo(toolName, params);
        if (keyInfo) {
          keyInfoCacheRef.current.set(blockId, keyInfo);
        }
        return keyInfo;
      } catch {
        return keyInfoCacheRef.current.get(blockId) || "";
      }
    }
  };

  const handleStreamComplete = (data: any) => {
    if (data && data.type) {
      switch (data.type) {
        case WebSocketMessageTypes.STREAM_CANCELLED:
          setIsStreaming(false);
          setIsCancelling(false);
          // Clear the ref so that idle status (sent after stream:cancelled) can be accepted
          currentAssistantMessageRef.current = null;
          return;
        case WebSocketMessageTypes.TOOL_PERMISSION:
          handleToolPermissionResponse(data);
          return;
        case WebSocketMessageTypes.TOOL_AUTO_APPROVED:
          handleToolAutoApproved(data);
          return;
        case WebSocketMessageTypes.TOOL_ERROR:
          handleToolErrorResponse(data);
          return;
        case WebSocketMessageTypes.UI_REQUEST: {
          const uiBlock = {
            id: `ui-request-block-${Date.now()}`,
            type: ContentBlockTypes.UI_REQUEST,
            content: data.payload?.description || "",
            toolUseId: data.toolUseId,
            isComplete: true,
            uiRequest: data.payload,
            uiRequestResponse: data.existingResponse || undefined,
            uiRequestSummary: data.summary || undefined,
          };

          if (currentAssistantMessageRef.current) {
            const messageId = currentAssistantMessageRef.current;
            setChatHistory((prev) =>
              prev.map((msg) =>
                msg.id === messageId
                  ? {
                      ...msg,
                      isStreaming: false,
                      contentBlocks: [...(msg.contentBlocks || []), uiBlock],
                    }
                  : msg,
              ),
            );
          } else {
            const uiRequestMessage: ChatMessage = {
              id: `assistant-ui-request-${Date.now()}-${Math.random()}`,
              role: "assistant",
              timestamp: Date.now(),
              isStreaming: false,
              contentBlocks: [uiBlock],
            };
            setChatHistory((prev) => [...prev, uiRequestMessage]);
          }

          setIsStreaming(false);
          return;
        }
      }
    }
    if (!currentAssistantMessageRef.current) {
      return;
    }
    setIsStreaming(false);
    const messageId = currentAssistantMessageRef.current;
    setChatHistory((prev) =>
      prev.map((msg) =>
        msg.id === messageId ? { ...msg, isStreaming: false } : msg,
      ),
    );
    if (approveAllEnabled) {
      setApproveAllEnabled(false);
      webSocketChatService.setApproveAll(false);
    }
    currentAssistantMessageRef.current = null;
  };

  const handleToolAutoApproved = (data: any) => {
    const toolIds = data.toolIds || [];
    setApprovedTools((prev) => {
      const newSet = new Set(prev);
      toolIds.forEach((toolId: string) => newSet.add(toolId));
      return newSet;
    });
  };

  const handleToolPermissionResponse = (data: any) => {
    const toolPermissions = data.toolPermissions || [data];
    if (!Array.isArray(toolPermissions) || toolPermissions.length === 0) {
      return;
    }
    if (data.message && setStatusMessage) {
      setStatusMessage(data.message);
    }
    const groupId = `group-${Date.now()}-${Math.random()}`;
    const toolsNeedingApproval = toolPermissions.map(
      (tool: any) => tool.toolId,
    );
    const newGroups = new Map(toolGroupsRef.current).set(
      groupId,
      toolsNeedingApproval,
    );
    toolGroupsRef.current = newGroups;
    setToolGroups(newGroups);
    setToolGroupDecisions((prev) => new Map(prev).set(groupId, new Map()));
    for (const toolId of toolsNeedingApproval) {
      void handleViewTool(toolId, true);
    }
  };

  const handleToolErrorResponse = (data: any) => {
    const toolId = data.toolId;
    const code = data.errorCode;
    const baseError =
      data.errorMessage || data.error || "Unknown error occurred";
    const errorMessage = code ? `${code}: ${baseError}` : baseError;
    // const errorMessage = data.errorMessage || data.error ||
    //   "Unknown error occurred";
    if (!toolId) return;
    setViewedTools((prev) => {
      const newSet = new Set(prev);
      newSet.delete(toolId);
      return newSet;
    });
    setErroredTools((prev) => new Map(prev).set(toolId, errorMessage));
    const groupId = findToolGroup(toolId);
    if (!groupId) {
      return;
    }
    void handleToolDecision(
      toolId,
      "rejected",
      `Tool execution failed: ${errorMessage}`,
    );
  };

  const handleStreamErrorMessage = (error: any) => {
    setIsStreaming(false);
    if (isCancelling) {
      return;
    }
    let errorContent = "";
    if (error instanceof DenoServerConnectionError) {
      errorContent = createDenoServerErrorContent();
    } else {
      let errorTitle = "Error";
      let errorDescription = "An error occurred while processing your request.";
      if (error && error.type) {
        const errorInfo = getErrorDisplayMessage(error.type);
        errorTitle = errorInfo.title;
        errorDescription = errorInfo.description;
      }
      errorContent = `❌ **${errorTitle}**\n\n${errorDescription}`;
    }
    const errorMessage: ChatMessage = {
      id: `stream-error-${Date.now()}-${Math.random()}`,
      role: "assistant",
      content: "",
      timestamp: Date.now(),
      contentBlocks: [
        {
          id: `stream-error-block-${Date.now()}`,
          type: "text" as const,
          content: errorContent,
          isComplete: true,
        },
      ],
      isStreaming: false,
    };
    setChatHistory((prev) => [...prev, errorMessage]);
    if (approveAllEnabled) {
      setApproveAllEnabled(false);
      webSocketChatService.setApproveAll(false);
    }
    currentAssistantMessageRef.current = null;
  };

  const handleConnectionLost = (
    hasShownConnectionError: boolean,
    chatHistoryLength: number,
  ) => {
    if (
      !hasShownConnectionError &&
      !currentAssistantMessageRef.current &&
      chatHistoryLength > 0
    ) {
      setHasShownConnectionError(true);
      const errorMessage: ChatMessage = {
        id: `connection-error-${Date.now()}-${Math.random()}`,
        role: "assistant",
        content: "",
        timestamp: Date.now(),
        contentBlocks: [
          {
            id: `connection-error-block-${Date.now()}`,
            type: "text" as const,
            content:
              "🔌 **Connection Lost to Deno Server**\n\n**Please restart Excel to restore the connection.**\n\nThe Deno server that powers the AI features has stopped responding. Restarting Excel will reinitialize the connection.\n\n_If the problem persists, ensure the Deno server is running properly._",
            isComplete: true,
          },
        ],
        isStreaming: false,
      };
      setChatHistory((prev) => [...prev, errorMessage]);
    }
  };

  const handleStreamDelta = (data: any) => {
    if (!currentAssistantMessageRef.current) {
      return;
    }
    try {
      const jsonData = JSON.parse(data);
      switch (jsonData.type) {
        case ClaudeEventTypes.CONTENT_BLOCK_START: {
          if (jsonData.content_block?.type === ContentBlockTypes.TEXT) {
            const blockId = `text-${Date.now()}-${Math.random()}`;
            const blockIndex = jsonData.index;
            blockIndexMap.current.set(blockIndex, blockId);
            updateCurrentAssistantMessageRef.current((msg) => ({
              ...msg,
              contentBlocks: [
                ...(msg.contentBlocks || []),
                {
                  id: blockId,
                  type: ContentBlockTypes.TEXT,
                  content: "",
                  isComplete: false,
                },
              ],
            }));
          } else if (
            jsonData.content_block?.type === ContentBlockTypes.THINKING
          ) {
            const blockId = `thinking-${Date.now()}-${Math.random()}`;
            const blockIndex = jsonData.index;
            const startTime = Date.now();
            blockIndexMap.current.set(blockIndex, blockId);
            updateCurrentAssistantMessageRef.current((msg) => ({
              ...msg,
              contentBlocks: [
                ...(msg.contentBlocks || []),
                {
                  id: blockId,
                  type: ContentBlockTypes.THINKING,
                  content: "",
                  isComplete: false,
                  startTime: startTime,
                },
              ],
            }));
          } else if (
            jsonData.content_block?.type === ClaudeContentTypes.TOOL_USE
          ) {
            const blockId = `tool-${Date.now()}-${Math.random()}`;
            const blockIndex = jsonData.index;
            const newToolNumber = toolCounter + 1;
            setToolCounter(newToolNumber);
            blockIndexMap.current.set(blockIndex, blockId);
            updateCurrentAssistantMessageRef.current((msg) => ({
              ...msg,
              contentBlocks: [
                ...(msg.contentBlocks || []),
                {
                  id: blockId,
                  type: ClaudeContentTypes.TOOL_USE,
                  content: "",
                  toolName: jsonData.content_block.name || "unknown",
                  toolUseId: jsonData.content_block.id,
                  toolNumber: newToolNumber,
                  isComplete: false,
                },
              ],
            }));
          }
          break;
        }
        case ClaudeEventTypes.CONTENT_BLOCK_DELTA: {
          if (jsonData.delta?.type === "text_delta") {
            const newText = jsonData.delta.text;
            const blockIndex = jsonData.index;
            let targetBlockId = blockIndexMap.current.get(blockIndex);

            // Fallback: if we never received a content_block_start, create one on the fly
            if (!targetBlockId) {
              targetBlockId = `text-${Date.now()}-${Math.random()}`;
              blockIndexMap.current.set(blockIndex, targetBlockId);
              const blockId = targetBlockId;
              const textToAdd = newText;
              updateCurrentAssistantMessageRef.current((msg) => ({
                ...msg,
                contentBlocks: [
                  ...(msg.contentBlocks || []),
                  {
                    id: blockId,
                    type: ContentBlockTypes.TEXT,
                    content: textToAdd,
                    isComplete: false,
                  },
                ],
              }));
            } else {
              updateCurrentAssistantMessageRef.current((msg) => ({
                ...msg,
                contentBlocks: (msg.contentBlocks || []).map((block) =>
                  block.id === targetBlockId
                    ? { ...block, content: block.content + newText }
                    : block,
                ),
              }));
            }
          } else if (jsonData.delta?.type === "thinking_delta") {
            const newThinking = jsonData.delta.thinking;
            const blockIndex = jsonData.index;
            const targetBlockId = blockIndexMap.current.get(blockIndex);
            if (targetBlockId) {
              updateCurrentAssistantMessageRef.current((msg) => ({
                ...msg,
                contentBlocks: (msg.contentBlocks || []).map((block) =>
                  block.id === targetBlockId
                    ? { ...block, content: block.content + newThinking }
                    : block,
                ),
              }));
            }
          } else if (jsonData.delta?.type === "input_json_delta") {
            const newJsonDelta = jsonData.delta.partial_json;
            const blockIndex = jsonData.index;
            const targetBlockId = blockIndexMap.current.get(blockIndex);
            if (targetBlockId) {
              updateCurrentAssistantMessageRef.current((msg) => ({
                ...msg,
                contentBlocks: (msg.contentBlocks || []).map((block) =>
                  block.id === targetBlockId &&
                  block.type === ClaudeContentTypes.TOOL_USE
                    ? { ...block, content: block.content + newJsonDelta }
                    : block,
                ),
              }));
            }
          }
          break;
        }
        case ClaudeEventTypes.CONTENT_BLOCK_STOP: {
          const stopBlockIndex = jsonData.index;
          const stopTargetBlockId = blockIndexMap.current.get(stopBlockIndex);
          if (stopTargetBlockId) {
            updateCurrentAssistantMessageRef.current((msg) => ({
              ...msg,
              contentBlocks: (msg.contentBlocks || []).map((block) => {
                if (block.id === stopTargetBlockId) {
                  const updatedBlock = {
                    ...block,
                    isComplete: true,
                    endTime:
                      block.type === ContentBlockTypes.THINKING
                        ? Date.now()
                        : block.endTime,
                  };
                  return updatedBlock;
                }
                return block;
              }),
            }));
            blockIndexMap.current.delete(stopBlockIndex);
          }
          break;
        }
        case ClaudeEventTypes.MESSAGE_START: {
          // Capture usage data from message_start
          if (jsonData.message?.usage && onUsageUpdate) {
            const { input_tokens = 0, output_tokens = 0 } =
              jsonData.message.usage;
            onUsageUpdate(input_tokens, output_tokens);
          }
          break;
        }
        case ClaudeEventTypes.MESSAGE_DELTA: {
          // Capture output token updates from message_delta
          if (jsonData.usage?.output_tokens && onUsageUpdate) {
            onUsageUpdate(0, jsonData.usage.output_tokens);
          }
          break;
        }
        case WebSocketMessageTypes.MESSAGE_STOP: {
          // Ignore message_stop; usage already captured
          break;
        }
        default: {
          const blockId = `info-${Date.now()}-${Math.random()}`;
          let formattedContent = "";
          if (jsonData.type && jsonData.type === OperationStatusValues.ERROR) {
            formattedContent = `❌ Error: ${
              jsonData.error || jsonData.message || "Unknown error occurred"
            }`;
            if (jsonData.error_type) {
              formattedContent += `\nType: ${jsonData.error_type}`;
            }
          } else {
            formattedContent = `ℹ️ ${jsonData.type || "Unknown message type"}`;
            if (jsonData.message) {
              formattedContent += `\n${jsonData.message}`;
            } else if (jsonData.content) {
              formattedContent += `\n${jsonData.content}`;
            } else {
              formattedContent += `\n${JSON.stringify(jsonData, null, 2)}`;
            }
          }
          updateCurrentAssistantMessageRef.current((msg) => ({
            ...msg,
            contentBlocks: [
              ...(msg.contentBlocks || []),
              {
                id: blockId,
                type: ContentBlockTypes.TEXT,
                content: formattedContent,
                isComplete: true,
              },
            ],
          }));
        }
      }
    } catch (error) {
      const blockId = `raw-${Date.now()}-${Math.random()}`;
      updateCurrentAssistantMessageRef.current((msg) => ({
        ...msg,
        contentBlocks: [
          ...(msg.contentBlocks || []),
          {
            id: blockId,
            type: "text",
            content: data,
            isComplete: true,
          },
        ],
      }));
    }
  };

  useEffect(() => {
    handleStreamCompleteRef.current = handleStreamComplete;
    handleStreamDeltaRef.current = handleStreamDelta;
    handleStreamErrorMessageRef.current = handleStreamErrorMessage;
    handleConnectionLostRef.current = handleConnectionLost;
  }, [
    handleStreamComplete,
    handleStreamDelta,
    handleStreamErrorMessage,
    handleConnectionLost,
  ]);

  const attachWebSocketHandlers = useCallback(
    (
      setWsConnected: (v: boolean) => void,
      updateWebSocketUrl: () => void,
      getChatHistoryLength: () => number,
      onChangeSummaryReceived?: (payload: any) => void,
    ) => {
      webSocketChatService.setEventHandlers({
        onStreamDelta: (data) => {
          if (typeof data === "string") {
            handleStreamDeltaRef.current(data);
          } else {
            handleStreamDeltaRef.current(JSON.stringify(data));
          }
        },
        onStreamComplete: (data) => {
          handleStreamCompleteRef.current(data);
        },
        onStreamError: (err) => {
          handleStreamErrorMessageRef.current(err);
        },
        onConnectionChange: (connected) => {
          setWsConnected(connected);
          if (connected) {
            updateWebSocketUrl();
            setHasShownConnectionError(false);
          } else {
            handleConnectionLostRef.current(false, getChatHistoryLength());
          }
        },
        // New status and token update handlers
        onStatusUpdate: (payload) => {
          if (isCancellingRef.current) {
            return;
          }

          // Race condition protection for "idle" status from server
          // Backend now sends stream:end BEFORE idle status, so currentAssistantMessageRef
          // is always cleared before idle arrives. If idle comes while ref is set,
          // it must be from a previous operation.
          if (
            payload.status === OperationStatusValues.IDLE &&
            currentAssistantMessageRef.current !== null
          ) {
            console.debug(
              `[onStatusUpdate] Ignoring stale "idle" from previous operation during active operation ${currentAssistantMessageRef.current}`,
            );
            return;
          }

          if (setOperationStatus) {
            setOperationStatus(payload.status);
          }
          if (setStatusMessage) {
            setStatusMessage(payload.message || null);
          }
          if (setOperationProgress) {
            setOperationProgress(payload.progress || null);
          }
        },
        onTokenUpdate: (payload) => {
          if (setInputTokens) {
            setInputTokens(payload.inputTokens);
          }
          if (setOutputTokens) {
            setOutputTokens(payload.outputTokens);
          }
          if (setTotalTokens) {
            setTotalTokens(payload.totalTokens);
          }
        },
        onChangeSummary: (payload) => {
          console.log("Change summary received in hook:", payload);
          onChangeSummaryReceived?.(payload);
        },
      });
    },
    [
      setHasShownConnectionError,
      setOperationStatus,
      setStatusMessage,
      setOperationProgress,
      setInputTokens,
      setOutputTokens,
      setTotalTokens,
    ],
  );

  const resetStreamRefs = () => {
    currentAssistantMessageRef.current = null;
    blockIndexMap.current.clear();
    keyInfoCacheRef.current.clear();
  };

  return {
    handleStreamDelta,
    handleStreamComplete,
    handleStreamErrorMessage,
    handleConnectionLost,
    getKeyInfo,
    updateCurrentAssistantMessage,
    currentAssistantMessageRef,
    blockIndexMap,
    keyInfoCacheRef,
    attachWebSocketHandlers,
    resetStreamRefs,
  };
}
