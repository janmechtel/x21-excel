import { useEffect, useRef, useState } from "react";

import type {
  ToolDecisionData,
  ToolGroupDecisions,
  ToolGroups,
} from "@/types/chat";
import { webSocketChatService } from "@/services/webSocketChatService";

export function useToolState() {
  const [approvedTools, setApprovedTools] = useState<Set<string>>(new Set());
  const [rejectedTools, setRejectedTools] = useState<Set<string>>(new Set());
  const [viewedTools, setViewedTools] = useState<Set<string>>(new Set());
  const [erroredTools, setErroredTools] = useState<Map<string, string>>(
    new Map(),
  );
  const [autoApproveEnabled, setAutoApproveEnabled] = useState<boolean>(false);
  const [approveAllEnabled, setApproveAllEnabled] = useState<boolean>(false);
  const [toolGroups, setToolGroups] = useState<ToolGroups>(new Map());
  const toolGroupsRef = useRef<ToolGroups>(new Map());
  const [toolGroupDecisions, setToolGroupDecisions] =
    useState<ToolGroupDecisions>(new Map());
  const [toolDecisions, setToolDecisions] = useState<
    Map<string, ToolDecisionData>
  >(new Map());

  const findToolGroup = (toolId: string): string | null => {
    for (const [groupId, toolIds] of toolGroupsRef.current.entries()) {
      if (toolIds.includes(toolId)) {
        return groupId;
      }
    }
    return null;
  };

  const isToolPending = (toolId: string): boolean => {
    const groupId = findToolGroup(toolId);
    if (!groupId) return false;
    const decisions = toolGroupDecisions.get(groupId);
    return !decisions?.has(toolId);
  };

  const getPendingToolCount = (): number => {
    let count = 0;
    for (const [groupId, toolIds] of toolGroups.entries()) {
      const decisions = toolGroupDecisions.get(groupId) || new Map();
      for (const toolId of toolIds) {
        if (!decisions.has(toolId)) {
          count++;
        }
      }
    }
    return count;
  };

  const resetToolState = () => {
    setApprovedTools(new Set());
    setRejectedTools(new Set());
    setViewedTools(new Set());
    setErroredTools(new Map());
    setToolGroups(new Map());
    toolGroupsRef.current = new Map();
    setToolGroupDecisions(new Map());
    setApproveAllEnabled(false);
    setToolDecisions(new Map());
  };

  useEffect(() => {
    webSocketChatService.setAutoApprove(autoApproveEnabled);
  }, [autoApproveEnabled]);

  useEffect(() => {
    webSocketChatService.setApproveAll(approveAllEnabled);
  }, [approveAllEnabled]);

  return {
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
  };
}
