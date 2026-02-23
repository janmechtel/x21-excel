import { getApiBase } from "./apiBase";
import {
  ToolNames,
  ClaudeContentTypes,
  ContentBlockTypes,
  type ContentBlockType,
  type AttachedFile,
} from "@/types/chat";

export interface RecentChatItem {
  conversationId: string;
  workbookKey: string;
  firstMessage: string;
  firstMessagePreview: string;
  lastActivityAt: number;
  participants: string[];
}

export interface MessageRow {
  id: string;
  workbookKey: string;
  conversationId: string;
  role: "user" | "assistant" | "system";
  content: string;
  createdAt: number;
}

export async function fetchRecentChats(
  workbookKey: string | null,
  limit: number = 1,
  scope: "file" | "all" = "file",
): Promise<RecentChatItem[]> {
  const base = await getApiBase();
  let url: string;
  if (scope === "all") {
    url = `${base}/api/recent-chats?scope=all&limit=${limit}`;
  } else {
    if (!workbookKey) {
      throw new Error('workbookKey is required when scope is "file"');
    }
    url = `${base}/api/recent-chats?workbookKey=${encodeURIComponent(
      workbookKey,
    )}&limit=${limit}`;
  }
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch recent chats: ${res.status}`);
  }
  const data = await res.json();
  return data?.items ?? [];
}

export async function fetchRecentUserMessages(
  workbookKey: string,
  limit: number = 50,
): Promise<MessageRow[]> {
  const base = await getApiBase();
  if (!workbookKey) {
    throw new Error("workbookKey is required");
  }
  const url = `${base}/api/recent-user-messages?workbookKey=${encodeURIComponent(
    workbookKey,
  )}&limit=${limit}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch recent user messages: ${res.status}`);
  }
  const data = await res.json();
  return data?.items ?? [];
}

export async function searchRecentChats(
  workbookKey: string | null,
  query: string,
  limit: number = 20,
  scope: "file" | "all" = "file",
): Promise<RecentChatItem[]> {
  const base = await getApiBase();
  const encodedQuery = encodeURIComponent(query);
  let url: string;
  if (scope === "all") {
    url = `${base}/api/search-chats?scope=all&limit=${limit}&q=${encodedQuery}`;
  } else {
    if (!workbookKey) {
      throw new Error('workbookKey is required when scope is "file"');
    }
    url = `${base}/api/search-chats?workbookKey=${encodeURIComponent(
      workbookKey,
    )}&limit=${limit}&q=${encodedQuery}`;
  }
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to search chats: ${res.status}`);
  }
  const data = await res.json();
  return data?.items ?? [];
}

// Detect if a stored user message is actually a tool-result payload
export function isToolResultMessage(raw: string): boolean {
  try {
    const parsed = JSON.parse(raw);
    return (
      Array.isArray(parsed) &&
      parsed.length > 0 &&
      parsed.every((item: any) => item?.type === ClaudeContentTypes.TOOL_RESULT)
    );
  } catch {
    return false;
  }
}

export function getAttachmentsFromContent(
  raw: string,
): AttachedFile[] | undefined {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      const isDocumentArray =
        parsed.length > 0 &&
        parsed.every(
          (item: any) =>
            (item?.type === "document" || item?.type === "image") &&
            item?.source?.type === "base64",
        );
      if (!isDocumentArray) {
        return undefined;
      }
      return parsed.map((item: any, index: number) => ({
        name: String(item?.name || `attachment-${index + 1}`),
        type: String(item?.source?.media_type || item?.type || ""),
        size: 0,
        base64: "",
      }));
    }
    if (
      parsed &&
      typeof parsed === "object" &&
      Array.isArray((parsed as any).attachments)
    ) {
      const attachments = (parsed as any).attachments as Array<any>;
      if (attachments.length === 0) {
        return undefined;
      }
      return attachments.map((att: any, index: number) => ({
        name: String(att?.name || `attachment-${index + 1}`),
        type: String(att?.type || ""),
        size: Number(att?.size || 0),
        base64: "",
      }));
    }
    return undefined;
  } catch {
    return undefined;
  }
}

export function hasAttachmentsInContent(raw: string): boolean {
  return (getAttachmentsFromContent(raw)?.length ?? 0) > 0;
}

export async function fetchMessages(
  workbookKey: string,
  conversationId: string,
): Promise<MessageRow[]> {
  const base = await getApiBase();
  const url = `${base}/api/messages?workbookKey=${encodeURIComponent(
    workbookKey,
  )}&conversationId=${encodeURIComponent(conversationId)}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch messages: ${res.status}`);
  }
  const data = await res.json();
  return data?.items ?? [];
}

// Extract userMessage from stored user message content
// Note: tool result messages are filtered out by the caller
export function getUserMessage(raw: string): string {
  try {
    const parsed = JSON.parse(raw);
    // Handle { userMessage: "...", excelContext: {...} } format
    if (parsed?.userMessage && typeof parsed.userMessage === "string") {
      return parsed.userMessage;
    }
    // Fallback: if it's already a string, return it
    return typeof parsed === "string" ? parsed : raw;
  } catch {
    // Not JSON: treat as plain text
    return raw;
  }
}

// Parse raw content JSON into UI content blocks for assistant messages (history)
export function parseAssistantContentBlocks(raw: string): Array<{
  id: string;
  type: ContentBlockType;
  content: string;
  toolName?: string;
  toolUseId?: string;
  isComplete: boolean;
  uiRequest?: any;
  uiRequestResponse?: any;
  uiRequestSummary?: string;
}> {
  try {
    const parsed = JSON.parse(raw);
    const blocks: Array<{
      id: string;
      type: ContentBlockType;
      content: string;
      toolName?: string;
      toolUseId?: string;
      isComplete: boolean;
      uiRequest?: any;
      uiRequestResponse?: any;
      uiRequestSummary?: string;
    }> = [];
    if (Array.isArray(parsed)) {
      for (const item of parsed) {
        if (!item || typeof item !== "object") continue;
        // Assistant text
        if (
          (item as any).type === ContentBlockTypes.TEXT &&
          typeof (item as any).text === "string"
        ) {
          blocks.push({
            id: `hist-text-${Date.now()}-${Math.random()}`,
            type: ContentBlockTypes.TEXT,
            content: (item as any).text,
            isComplete: true,
          });
        } // Assistant tool use
        else if ((item as any).type === ClaudeContentTypes.TOOL_USE) {
          const toolName = (item as any).name || "unknown";
          const toolUseId = (item as any).id || "";
          const input = (item as any).input ?? {};
          let content = "";
          try {
            content = JSON.stringify(input);
          } catch {
            // Fallback to raw string if input isn't serializable
            content = String(input ?? "");
          }
          blocks.push({
            id: `hist-tool-${Date.now()}-${Math.random()}`,
            type: ClaudeContentTypes.TOOL_USE,
            content,
            toolName,
            toolUseId,
            isComplete: true,
          });

          // Check if this is a collect_input (ui_request) tool
          if (toolName === ToolNames.COLLECT_INPUT) {
            blocks.push({
              id: `hist-ui-request-${Date.now()}-${Math.random()}`,
              type: ContentBlockTypes.UI_REQUEST,
              content: input.description || "",
              toolUseId,
              isComplete: true,
              uiRequest: input,
            });
          } else {
            let content = "";
            try {
              content = JSON.stringify(input);
            } catch {
              // Fallback to raw string if input isn't serializable
              content = String(input ?? "");
            }
            blocks.push({
              id: `hist-tool-${Date.now()}-${Math.random()}`,
              type: ClaudeContentTypes.TOOL_USE,
              content,
              toolName,
              toolUseId,
              isComplete: true,
            });
          }
        } // Thinking (if present in stored history)
        else if (
          (item as any).type === ContentBlockTypes.THINKING &&
          typeof (item as any).thinking === "string"
        ) {
          blocks.push({
            id: `hist-thinking-${Date.now()}-${Math.random()}`,
            type: ContentBlockTypes.THINKING,
            content: (item as any).thinking,
            isComplete: true,
          });
        }
        // Tool result blocks are merged during live streaming; skip standalone tool_result here.
      }
      return blocks;
    }
    if (typeof parsed === "string") {
      return [
        {
          id: `hist-text-${Date.now()}-${Math.random()}`,
          type: "text",
          content: parsed,
          isComplete: true,
        },
      ];
    }
    // Unknown shape: fall back to raw string
    return [
      {
        id: `hist-text-${Date.now()}-${Math.random()}`,
        type: "text",
        content: raw,
        isComplete: true,
      },
    ];
  } catch {
    // Not JSON: treat as plain text
    return [
      {
        id: `hist-text-${Date.now()}-${Math.random()}`,
        type: "text",
        content: raw,
        isComplete: true,
      },
    ];
  }
}
