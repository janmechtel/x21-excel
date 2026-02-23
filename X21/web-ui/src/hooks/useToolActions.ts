import {
  type Dispatch,
  type SetStateAction,
  useCallback,
  useState,
} from "react";

import { webSocketChatService } from "@/services/webSocketChatService";
import type {
  ChatMessage,
  ToolDecision,
  ToolDecisionData,
  ToolDecisionList,
} from "@/types/chat";
import { ClaudeContentTypes } from "@/types/chat";
import { groupMessagesIntoConversations } from "@/utils/chat";

interface Params {
  chatHistory: ChatMessage[];
  setChatHistory: Dispatch<SetStateAction<ChatMessage[]>>;
  setRevertedConversations: Dispatch<SetStateAction<Set<string>>>;
  setIsRevertingOrApplying: (val: boolean) => void;
  setApprovedTools: Dispatch<SetStateAction<Set<string>>>;
  setRejectedTools: Dispatch<SetStateAction<Set<string>>>;
  viewedTools: Set<string>;
  setViewedTools: Dispatch<SetStateAction<Set<string>>>;
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
  setApproveAllEnabled: (val: boolean) => void;
  findToolGroup: (toolId: string) => string | null;
  setToolDecisions: Dispatch<SetStateAction<Map<string, ToolDecisionData>>>;
}

export function useToolActions({
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
}: Params) {
  const [rejectingToolId, setRejectingToolId] = useState<string | null>(null);
  const [rejectMessage, setRejectMessage] = useState("");

  const sendGroupResponse = useCallback(
    async (groupId: string, passedGroupDecisions?: ToolDecisionList) => {
      const groupTools = toolGroupsRef.current.get(groupId) || [];
      const groupDecisions: ToolDecisionList =
        passedGroupDecisions || toolGroupDecisions.get(groupId) || new Map();

      const toolResponses = groupTools.map((toolId) => {
        const decisionData = groupDecisions.get(toolId);
        const decision = decisionData?.decision || "approved";
        const userMessage = decisionData?.message;
        const response: any = { toolId, decision };
        if (decision === "rejected" && userMessage) {
          response.userMessage = userMessage;
        }
        return response;
      });

      try {
        const success = await webSocketChatService.sendToolPermissionResponse(
          groupId,
          toolResponses,
        );
        if (success) {
          setToolGroups((prev) => {
            const newMap = new Map(prev);
            newMap.delete(groupId);
            toolGroupsRef.current = newMap;
            return newMap;
          });
          setToolGroupDecisions((prev) => {
            const newMap = new Map(prev);
            newMap.delete(groupId);
            return newMap;
          });
        }
      } catch (error) {
        console.error(
          `❌ Failed to send group response for ${groupId}:`,
          error,
        );
      }
    },
    [toolGroupsRef, toolGroupDecisions, setToolGroups, setToolGroupDecisions],
  );

  const handleToolDecision = useCallback(
    async (toolId: string, decision: ToolDecision, message?: string) => {
      try {
        const groupId = findToolGroup(toolId);
        if (!groupId) {
          console.warn(
            `⚠️ Tool ${toolId} is not part of any group - cannot process decision`,
          );
          return;
        }

        const groupDecisions: ToolDecisionList = new Map(
          toolGroupDecisions.get(groupId) || [],
        );
        groupDecisions.set(toolId, { decision, message });
        const updatedGroupDecisions = groupDecisions;

        setToolGroupDecisions((prev) => {
          const newMap = new Map(prev);
          newMap.set(groupId, groupDecisions);
          return newMap;
        });

        setToolDecisions((prev) => {
          const next = new Map(prev);
          next.set(toolId, { decision, message });
          return next;
        });

        if (decision === "approved") {
          setApprovedTools((prev) => new Set([...prev, toolId]));
        } else {
          setRejectedTools((prev) => new Set([...prev, toolId]));
        }

        const groupTools = toolGroupsRef.current.get(groupId) || [];
        const isComplete = groupTools.every((tid) =>
          updatedGroupDecisions.has(tid),
        );

        if (isComplete) {
          await sendGroupResponse(groupId, updatedGroupDecisions);
        }
      } catch (error) {
        console.error(
          `❌ Failed to ${
            decision === "approved" ? "approve" : "reject"
          } tool ${toolId}:`,
          error,
        );
      }
    },
    [
      findToolGroup,
      toolGroupDecisions,
      setToolGroupDecisions,
      setApprovedTools,
      setRejectedTools,
      toolGroupsRef,
      sendGroupResponse,
      setToolDecisions,
    ],
  );

  const handleApproveTools = useCallback(
    async (toolId: string) => {
      const groupId = findToolGroup(toolId);
      if (!groupId) {
        console.warn(
          `⚠️ Tool ${toolId} is not part of any group - cannot approve`,
        );
        return;
      }
      await handleToolDecision(toolId, "approved");
    },
    [findToolGroup, handleToolDecision],
  );

  const executeRejectTool = useCallback(
    async (toolId: string, message?: string) => {
      const groupId = findToolGroup(toolId);
      if (!groupId) {
        console.warn(
          `⚠️ Tool ${toolId} is not part of any group - cannot reject`,
        );
        return;
      }
      await handleToolDecision(toolId, "rejected", message);
    },
    [findToolGroup, handleToolDecision],
  );

  const handleRejectSubmit = useCallback(() => {
    if (rejectingToolId) {
      void executeRejectTool(
        rejectingToolId,
        rejectMessage.trim() || undefined,
      );
      setRejectingToolId(null);
      setRejectMessage("");
    }
  }, [executeRejectTool, rejectMessage, rejectingToolId]);

  const handleRejectCancel = useCallback(() => {
    setRejectingToolId(null);
    setRejectMessage("");
  }, []);

  const handleRejectTool = useCallback(async (toolId: string) => {
    setRejectingToolId(toolId);
    setRejectMessage("");
  }, []);

  const handleApproveAll = useCallback(
    async (toolId: string) => {
      try {
        setApproveAllEnabled(true);
        webSocketChatService.setApproveAll(true);

        const groupId = findToolGroup(toolId);
        if (!groupId) {
          console.error(`❌ Tool ${toolId} is not part of any group`);
          return;
        }

        const groupTools = toolGroupsRef.current.get(groupId) || [];
        const groupDecisions: ToolDecisionList =
          toolGroupDecisions.get(groupId) || new Map();

        let updatedGroupDecisions: ToolDecisionList = new Map(groupDecisions);
        for (const tid of groupTools) {
          if (!updatedGroupDecisions.has(tid)) {
            updatedGroupDecisions.set(tid, { decision: "approved" });
            setApprovedTools((prev) => new Set([...prev, tid]));
          }
        }

        setToolDecisions((prev) => {
          const next = new Map(prev);
          for (const tid of groupTools) {
            const decisionData = updatedGroupDecisions.get(tid);
            next.set(tid, {
              decision: decisionData?.decision || "approved",
              message: decisionData?.message,
            });
          }
          return next;
        });

        setToolGroupDecisions((prev) => {
          const newMap = new Map(prev);
          newMap.set(groupId, updatedGroupDecisions);
          return newMap;
        });

        await sendGroupResponse(groupId, updatedGroupDecisions);
      } catch (error) {
        console.error(`❌ Failed to approve all tools:`, error);
      }
    },
    [
      findToolGroup,
      setApproveAllEnabled,
      toolGroupsRef,
      toolGroupDecisions,
      setApprovedTools,
      setToolGroupDecisions,
      sendGroupResponse,
      setToolDecisions,
    ],
  );

  const handleViewTool = useCallback(
    async (toolId: string, _isAutoView: boolean = false) => {
      try {
        const isCurrentlyViewed = viewedTools.has(toolId);
        if (isCurrentlyViewed) {
          const unviewSuccess = await webSocketChatService.unviewTool(toolId);
          if (!unviewSuccess) {
            throw new Error(
              "Failed to send unview request - WebSocket not connected",
            );
          }
          setViewedTools((prev) => {
            const newSet = new Set(prev);
            newSet.delete(toolId);
            return newSet;
          });
        } else {
          const viewSuccess = await webSocketChatService.viewTool(toolId);
          if (!viewSuccess) {
            throw new Error(
              "Failed to send view request - WebSocket not connected",
            );
          }
          setViewedTools((prev) => new Set([...prev, toolId]));
        }
      } catch (error) {
        console.error(
          `❌ Failed to ${
            viewedTools.has(toolId) ? "unview" : "view"
          } tool ${toolId}:`,
          error,
        );
      }
    },
    [setViewedTools, viewedTools],
  );

  const executeRevertOperation = useCallback(
    async (toolUseId: string, toolName: string) => {
      if (!toolUseId || !toolName) return;
      setIsRevertingOrApplying(true);

      try {
        let toolNumber = 0;
        for (const message of chatHistory) {
          if (message.contentBlocks) {
            for (const block of message.contentBlocks) {
              if (
                block.type === ClaudeContentTypes.TOOL_USE &&
                block.toolUseId === toolUseId
              ) {
                toolNumber = block.toolNumber || 0;
                break;
              }
            }
          }
          if (toolNumber > 0) break;
        }

        const revertSuccess = await webSocketChatService.revertTool(toolUseId);
        if (!revertSuccess) {
          throw new Error(
            "Failed to send revert request - WebSocket not connected",
          );
        }

        const conversations = groupMessagesIntoConversations(chatHistory);
        for (const conversation of conversations) {
          for (const assistantMessage of conversation.assistantMessages) {
            if (assistantMessage.contentBlocks) {
              for (const block of assistantMessage.contentBlocks) {
                if (
                  block.type === ClaudeContentTypes.TOOL_USE &&
                  block.toolUseId === toolUseId
                ) {
                  setRevertedConversations((prev) => {
                    const next = new Set(prev);
                    next.add(conversation.userMessage.id);
                    return next;
                  });
                  return;
                }
              }
            }
          }
        }
      } catch (error) {
        const errorMessage: ChatMessage = {
          id: `revert-error-${Date.now()}-${Math.random()}`,
          role: "assistant",
          timestamp: Date.now(),
          contentBlocks: [
            {
              id: `revert-error-block-${Date.now()}`,
              type: "text",
              content: `❌ Failed to revert changes: ${error}`,
              isComplete: true,
            },
          ],
        };
        setChatHistory((prev) => [...prev, errorMessage]);
      } finally {
        setIsRevertingOrApplying(false);
      }
    },
    [
      chatHistory,
      setChatHistory,
      setIsRevertingOrApplying,
      setRevertedConversations,
    ],
  );

  const executeApplyOperation = useCallback(
    async (toolUseId: string, toolName: string) => {
      if (!toolUseId || !toolName) return;
      setIsRevertingOrApplying(true);

      try {
        let toolNumber = 0;
        for (const message of chatHistory) {
          if (message.contentBlocks) {
            for (const block of message.contentBlocks) {
              if (
                block.type === ClaudeContentTypes.TOOL_USE &&
                block.toolUseId === toolUseId
              ) {
                toolNumber = block.toolNumber || 0;
                break;
              }
            }
          }
          if (toolNumber > 0) break;
        }

        const applySuccess = await webSocketChatService.applyTool(toolUseId);
        if (!applySuccess) {
          throw new Error(
            "Failed to send apply request - WebSocket not connected",
          );
        }

        const conversations = groupMessagesIntoConversations(chatHistory);
        for (const conversation of conversations) {
          for (const assistantMessage of conversation.assistantMessages) {
            if (assistantMessage.contentBlocks) {
              for (const block of assistantMessage.contentBlocks) {
                if (
                  block.type === ClaudeContentTypes.TOOL_USE &&
                  block.toolUseId === toolUseId
                ) {
                  setRevertedConversations((prev) => {
                    const newSet = new Set(prev);
                    newSet.delete(conversation.userMessage.id);
                    return newSet;
                  });
                  return;
                }
              }
            }
          }
        }
      } catch (error) {
        const errorMessage: ChatMessage = {
          id: `apply-error-${Date.now()}-${Math.random()}`,
          role: "assistant",
          timestamp: Date.now(),
          contentBlocks: [
            {
              id: `apply-error-block-${Date.now()}`,
              type: "text",
              content: `❌ Failed to apply changes: ${error}`,
              isComplete: true,
            },
          ],
        };
        setChatHistory((prev) => [...prev, errorMessage]);
      } finally {
        setIsRevertingOrApplying(false);
      }
    },
    [
      chatHistory,
      setChatHistory,
      setIsRevertingOrApplying,
      setRevertedConversations,
    ],
  );

  const handleRevertFromTool = useCallback(
    async (toolUseId: string, toolName: string) => {
      await executeRevertOperation(toolUseId, toolName);
    },
    [executeRevertOperation],
  );

  const handleApplyFromTool = useCallback(
    async (toolUseId: string, toolName: string) => {
      await executeApplyOperation(toolUseId, toolName);
    },
    [executeApplyOperation],
  );

  return {
    handleToolDecision,
    handleApproveTools,
    handleRejectTool,
    handleRejectSubmit,
    handleRejectCancel,
    handleApproveAll,
    handleViewTool,
    handleRevertFromTool,
    handleApplyFromTool,
    rejectingToolId,
    rejectMessage,
    setRejectMessage,
  };
}
