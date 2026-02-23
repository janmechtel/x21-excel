import {
  fetchMessages,
  fetchRecentChats,
  fetchRecentUserMessages,
  getAttachmentsFromContent,
  getUserMessage,
  isToolResultMessage,
  type MessageRow,
  parseAssistantContentBlocks,
  type RecentChatItem,
  searchRecentChats,
} from "./historyService";
import type { AttachedFile, ToolDecisionData } from "@/types/chat";
import {
  ClaudeContentTypes,
  ContentBlockTypes,
  type ContentBlockType,
} from "@/types/chat";

export interface HistoryContentBlock {
  id: string;
  type: ContentBlockType;
  content: string;
  toolName?: string;
  toolUseId?: string;
  toolNumber?: number;
  isComplete: boolean;
  startTime?: number;
  endTime?: number;
  uiRequest?: any;
  uiRequestResponse?: any;
  uiRequestSummary?: string;
}

export interface HistoryChatMessage {
  id: string;
  role: "user" | "assistant";
  content?: string;
  timestamp: number;
  contentBlocks?: HistoryContentBlock[];
  isStreaming?: boolean;
  score?: "up" | "down" | null;
  slashCommandId?: string;
  attachedFiles?: AttachedFile[];
}

export async function listRecentConversationsForFile(
  workbookKey: string,
  limit: number = 20,
): Promise<RecentChatItem[]> {
  return fetchRecentChats(workbookKey, limit, "file");
}

export async function listRecentConversations(
  workbookKey: string | null,
  limit: number = 20,
  scope: "file" | "all" = "file",
): Promise<RecentChatItem[]> {
  return fetchRecentChats(workbookKey, limit, scope);
}

export async function listRecentUserMessagesForFile(
  workbookKey: string,
  limit: number = 50,
): Promise<MessageRow[]> {
  return fetchRecentUserMessages(workbookKey, limit);
}

export async function searchConversations(
  workbookKey: string | null,
  query: string,
  limit: number = 50,
  scope: "file" | "all" = "file",
): Promise<RecentChatItem[]> {
  return searchRecentChats(workbookKey, query, limit, scope);
}

export async function loadLatestConversationForFile(
  workbookKey: string,
): Promise<{
  conversationId: string | null;
  messages: HistoryChatMessage[];
  toolDecisions: Map<string, ToolDecisionData>;
}> {
  const recents = await fetchRecentChats(workbookKey, 1);
  if (!recents || recents.length === 0) {
    return { conversationId: null, messages: [], toolDecisions: new Map() };
  }

  const conversationId = recents[0].conversationId;
  const { messages, toolDecisions } = await loadConversationForFile(
    workbookKey,
    conversationId,
  );
  return { conversationId, messages, toolDecisions };
}

export async function loadConversationForFile(
  workbookKey: string,
  conversationId: string,
): Promise<{
  messages: HistoryChatMessage[];
  toolDecisions: Map<string, ToolDecisionData>;
}> {
  const rows: MessageRow[] = await fetchMessages(workbookKey, conversationId);
  const toolDecisions = extractToolDecisions(rows);
  const uiRequestResponses = extractUiRequestResponses(rows);
  const messages = mapRowsToChatMessages(rows, uiRequestResponses);
  return { messages, toolDecisions };
}

export function mapRowsToChatMessages(
  rows: MessageRow[],
  uiRequestResponses?: Map<string, { response: any; summary?: string }>,
): HistoryChatMessage[] {
  let toolCount = 0;

  return rows
    .filter((r) => {
      // Only include user and assistant messages, but skip tool result messages
      return (
        (r.role === "user" || r.role === "assistant") &&
        !(r.role === "user" && isToolResultMessage(r.content))
      );
    })
    .map<HistoryChatMessage>((r) => {
      if (r.role === "assistant") {
        const blocks = parseAssistantContentBlocks(r.content).map((b) => {
          if (b.type === ClaudeContentTypes.TOOL_USE) {
            toolCount += 1;
            return { ...b, toolNumber: toolCount };
          }
          if (
            b.type === ContentBlockTypes.UI_REQUEST &&
            b.toolUseId &&
            uiRequestResponses
          ) {
            const responseData = uiRequestResponses.get(b.toolUseId);
            if (responseData) {
              return {
                ...b,
                uiRequestResponse: responseData.response,
                uiRequestSummary: responseData.summary,
              };
            }
          }
          return b;
        });

        return {
          id: r.id,
          role: "assistant",
          timestamp: r.createdAt,
          contentBlocks: blocks,
          isStreaming: false,
        };
      } else {
        const attachments = getAttachmentsFromContent(r.content);
        return {
          id: r.id,
          role: "user",
          content: getUserMessage(r.content),
          timestamp: r.createdAt,
          attachedFiles: attachments,
        };
      }
    });
}

export function extractToolDecisions(
  rows: MessageRow[],
): Map<string, ToolDecisionData> {
  const decisions = new Map<string, ToolDecisionData>();
  const rejectionPrefix = "User Message on Tool Call Rejection:";

  for (const row of rows) {
    if (row.role !== "user" || !isToolResultMessage(row.content)) continue;

    try {
      const parsed = JSON.parse(row.content);
      if (!Array.isArray(parsed)) continue;

      for (const item of parsed) {
        if (!item || typeof item !== "object") continue;
        if ((item as any).type !== ClaudeContentTypes.TOOL_RESULT) continue;

        const toolUseId = (item as any).tool_use_id || (item as any).toolUseId;
        if (!toolUseId) continue;

        const normalized = normalizeToolResultContent((item as any).content);
        const lower = normalized.toLowerCase();

        if (normalized.startsWith(rejectionPrefix)) {
          const message = normalized.slice(rejectionPrefix.length).trim();
          decisions.set(toolUseId, { decision: "rejected", message });
        } else if (lower.includes("tool execution failed")) {
          decisions.set(toolUseId, {
            decision: "rejected",
            message: normalized,
          });
        } else if (
          lower.includes("cancelled by the user") ||
          lower.includes("cancelled by user")
        ) {
          decisions.set(toolUseId, {
            decision: "rejected",
            message: normalized,
          });
        }
      }
    } catch {
      // ignore malformed tool_result history rows
    }
  }

  return decisions;
}

export function extractUiRequestResponses(
  rows: MessageRow[],
): Map<string, { response: any; summary?: string }> {
  const responses = new Map<string, { response: any; summary?: string }>();

  for (const row of rows) {
    if (row.role !== "user" || !isToolResultMessage(row.content)) continue;

    try {
      const parsed = JSON.parse(row.content);
      if (!Array.isArray(parsed)) continue;

      for (const item of parsed) {
        if (!item || typeof item !== "object") continue;
        if ((item as any).type !== ClaudeContentTypes.TOOL_RESULT) continue;

        const toolUseId = (item as any).tool_use_id || (item as any).toolUseId;
        if (!toolUseId) continue;

        const content = (item as any).content;
        if (typeof content === "string") {
          try {
            // Try to parse as JSON (ui_request response)
            const response = JSON.parse(content);
            if (response && typeof response === "object") {
              responses.set(toolUseId, { response });
            }
          } catch {
            // Not JSON, skip
          }
        } else if (content && typeof content === "object") {
          responses.set(toolUseId, { response: content });
        }
      }
    } catch {
      // ignore malformed tool_result history rows
    }
  }

  return responses;
}

function normalizeToolResultContent(raw: unknown): string {
  if (typeof raw === "string") {
    const trimmed = raw.trim();
    try {
      const parsed = JSON.parse(trimmed);
      if (typeof parsed === "string") {
        return parsed;
      }
    } catch {
      // not JSON-encoded string
    }
    return trimmed;
  }

  if (raw === null || raw === undefined) {
    return "";
  }

  if (typeof raw === "object") {
    try {
      return JSON.stringify(raw);
    } catch {
      return String(raw);
    }
  }

  return String(raw);
}
