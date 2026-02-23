import { getDb, nowMs } from "./sqlite.ts";
import { ClaudeContentTypes } from "../types/index.ts";

function uuid(): string {
  return crypto.randomUUID();
}

export type MessageRow = {
  id: string;
  workbookKey: string;
  conversationId: string;
  role: string;
  content: string;
  createdAt: number;
};

// Detect if a stored user message is actually a tool-result payload
function isToolResultMessage(raw: string): boolean {
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) &&
      parsed.length > 0 &&
      parsed.every((item: any) =>
        item?.type === ClaudeContentTypes.TOOL_RESULT
      );
  } catch {
    return false;
  }
}

function isAttachmentOnlyMessage(raw: string): boolean {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.length > 0 &&
        parsed.every((item: any) =>
          (item?.type === "document" || item?.type === "image") &&
          item?.source?.type === "base64"
        );
    }
    if (parsed && typeof parsed === "object") {
      const attachments = Array.isArray((parsed as any).attachments)
        ? (parsed as any).attachments
        : [];
      const userMessage = typeof (parsed as any).userMessage === "string"
        ? (parsed as any).userMessage.trim()
        : "";
      return attachments.length > 0 && userMessage.length === 0;
    }
    return false;
  } catch {
    return false;
  }
}

function sanitizeAttachmentPreview(row: RecentChatItem): RecentChatItem {
  if (!isAttachmentOnlyMessage(row.firstMessage)) {
    return row;
  }

  return {
    ...row,
    firstMessage: "",
    firstMessagePreview: "(attachment)",
  };
}

function normalizeAttachmentContent(raw: string): string {
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      const attachments = parsed.map((item: any, index: number) => ({
        name: String(item?.name || `attachment-${index + 1}`),
        type: String(item?.source?.media_type || item?.type || ""),
        size: Number(item?.size || 0),
      }));
      return JSON.stringify({
        userMessage: "(attachment)",
        attachments,
      });
    }
    if (parsed && typeof parsed === "object") {
      const attachments = Array.isArray((parsed as any).attachments)
        ? (parsed as any).attachments
        : [];
      const userMessage = typeof (parsed as any).userMessage === "string"
        ? (parsed as any).userMessage.trim()
        : "";
      if (attachments.length > 0 && userMessage.length === 0) {
        return JSON.stringify({
          ...parsed,
          userMessage: "(attachment)",
          attachments,
        });
      }
    }
    return raw;
  } catch {
    return raw;
  }
}

function mapRecentChatRows(
  rows: Array<[string, string | null, string, number]>,
): RecentChatItem[] {
  return rows.map((r) => ({
    conversationId: r[0],
    workbookKey: r[1] ?? "",
    firstMessage: r[2]?.toString() ?? "",
    firstMessagePreview: r[2]?.toString().slice(0, 200) ?? "",
    lastActivityAt: Number(r[3]),
    participants: ["user", "assistant"],
  })).map(sanitizeAttachmentPreview);
}

export function insertMessage(
  row: Omit<MessageRow, "id" | "createdAt"> & { createdAt?: number },
): string {
  const db = getDb();
  const id = uuid();
  const createdAt = row.createdAt ?? nowMs();
  db.query(
    "INSERT INTO messages (id, workbook_key, conversation_id, role, content, created_at) VALUES (?, ?, ?, ?, ?, ?);",
    [
      id,
      row.workbookKey,
      row.conversationId,
      row.role,
      row.content,
      createdAt,
    ],
  );
  return id;
}

export function updateToolResultContent(
  workbookKey: string,
  conversationId: string,
  toolUseId: string,
  updatedContent: string,
): number {
  const db = getDb();
  const pattern = `%${toolUseId}%`;

  db.query(
    `UPDATE messages
     SET content = ?
     WHERE workbook_key = ? AND conversation_id = ? AND role = 'user' AND content LIKE ?;`,
    [updatedContent, workbookKey, conversationId, pattern],
  );

  return db.changes;
}

export type RecentChatItem = {
  conversationId: string;
  workbookKey: string;
  firstMessage: string;
  firstMessagePreview: string;
  lastActivityAt: number;
  participants: string[];
};

export function listRecentChatsAll(
  limit: number = 1,
): RecentChatItem[] {
  const db = getDb();

  const stmt = db.prepareQuery<[string, string | null, string, number]>(`
SELECT first_m.conversation_id, first_m.workbook_key, first_m.content, latest.last_created_at
FROM (
  SELECT conversation_id, MIN(created_at) AS first_created_at
  FROM messages
  GROUP BY conversation_id
) first
JOIN messages first_m
  ON first_m.conversation_id = first.conversation_id
 AND first_m.created_at = first.first_created_at
JOIN (
  SELECT conversation_id, MAX(created_at) AS last_created_at
  FROM messages
  GROUP BY conversation_id
) latest
  ON latest.conversation_id = first_m.conversation_id
ORDER BY latest.last_created_at DESC
LIMIT ?;`);

  const rows = stmt.all([limit]);
  stmt.finalize();

  return mapRecentChatRows(rows);
}

export function searchRecentChatsAll(
  query: string,
  limit: number = 20,
): RecentChatItem[] {
  const db = getDb();
  const pattern = `%${query}%`;

  const stmt = db.prepareQuery<[string, string | null, string, number]>(`
SELECT first_m.conversation_id, first_m.workbook_key, first_m.content, latest.last_created_at
FROM (
  SELECT conversation_id, MIN(created_at) AS first_created_at
  FROM messages
  GROUP BY conversation_id
) first
JOIN messages first_m
  ON first_m.conversation_id = first.conversation_id
 AND first_m.created_at = first.first_created_at
JOIN (
  SELECT conversation_id, MAX(created_at) AS last_created_at
  FROM messages
  GROUP BY conversation_id
) latest
  ON latest.conversation_id = first_m.conversation_id
WHERE first_m.content LIKE ?
ORDER BY latest.last_created_at DESC
LIMIT ?;`);

  const rows = stmt.all([pattern, limit]);
  stmt.finalize();

  return mapRecentChatRows(rows);
}

export function listMessagesByConversation(
  workbookKey: string,
  conversationId: string,
): MessageRow[] {
  const db = getDb();
  const stmt = db.prepareQuery<
    [string, string | null, string, string, string, number]
  >(`
SELECT id, workbook_key, conversation_id, role, content, created_at
FROM messages
WHERE workbook_key = ? AND conversation_id = ?
ORDER BY created_at ASC;`);

  const rows = stmt.all([workbookKey, conversationId]);
  stmt.finalize();

  const mapped = rows.map((r) => ({
    id: r[0],
    workbookKey: r[1] ?? "",
    conversationId: r[2],
    role: r[3],
    content: r[3] === "user" && isAttachmentOnlyMessage(r[4])
      ? normalizeAttachmentContent(r[4])
      : r[4],
    createdAt: Number(r[5]),
  }));

  return mapped;
}

export function listRecentChats(
  workbookKey: string,
  limit: number = 1,
): RecentChatItem[] {
  const db = getDb();

  // Get first message per conversation (as preview) for the given file,
  // ordered by latest activity (most recent message) first
  const stmt = db.prepareQuery<[string, string | null, string, number]>(`
SELECT first_m.conversation_id, first_m.workbook_key, first_m.content, latest.last_created_at
FROM (
  SELECT conversation_id, MIN(created_at) AS first_created_at
  FROM messages
  WHERE workbook_key = ?
  GROUP BY conversation_id
) first
JOIN messages first_m
  ON first_m.conversation_id = first.conversation_id
 AND first_m.created_at = first.first_created_at
JOIN (
  SELECT conversation_id, MAX(created_at) AS last_created_at
  FROM messages
  WHERE workbook_key = ?
  GROUP BY conversation_id
) latest
  ON latest.conversation_id = first_m.conversation_id
WHERE first_m.workbook_key = ?
ORDER BY latest.last_created_at DESC
LIMIT ?;`);

  const rows = stmt.all([workbookKey, workbookKey, workbookKey, limit]);
  stmt.finalize();

  // For participants, this minimal implementation assumes user and assistant
  // which matches our current system roles and UI expectations.
  return mapRecentChatRows(rows);
}

export function listRecentUserMessages(
  workbookKey: string,
  limit: number = 50,
): MessageRow[] {
  const db = getDb();
  const fetchLimit = Math.min(Math.max(limit * 3, limit), 500);
  const stmt = db.prepareQuery<
    [string, string | null, string, string, string, number]
  >(`
SELECT id, workbook_key, conversation_id, role, content, created_at
FROM messages
WHERE workbook_key = ?
  AND role = 'user'
ORDER BY created_at DESC
LIMIT ?;`);

  const rows = stmt.all([workbookKey, fetchLimit]);
  stmt.finalize();

  const mapped = rows.map((r) => ({
    id: r[0],
    workbookKey: r[1] ?? "",
    conversationId: r[2],
    role: r[3],
    content: r[4],
    createdAt: Number(r[5]),
  }));

  return mapped
    .filter((row) =>
      !isToolResultMessage(row.content) &&
      !isAttachmentOnlyMessage(row.content)
    )
    .slice(0, limit);
}

export function searchRecentChats(
  workbookKey: string,
  query: string,
  limit: number = 20,
): RecentChatItem[] {
  const db = getDb();
  const pattern = `%${query}%`;

  const stmt = db.prepareQuery<[string, string | null, string, number]>(`
SELECT first_m.conversation_id, first_m.workbook_key, first_m.content, latest.last_created_at
FROM (
  SELECT conversation_id, MIN(created_at) AS first_created_at
  FROM messages
  WHERE workbook_key = ?
  GROUP BY conversation_id
) first
JOIN messages first_m
  ON first_m.conversation_id = first.conversation_id
 AND first_m.created_at = first.first_created_at
JOIN (
  SELECT conversation_id, MAX(created_at) AS last_created_at
  FROM messages
  WHERE workbook_key = ?
  GROUP BY conversation_id
) latest
  ON latest.conversation_id = first_m.conversation_id
WHERE first_m.workbook_key = ?
  AND first_m.content LIKE ?
ORDER BY latest.last_created_at DESC
LIMIT ?;`);

  const rows = stmt.all([
    workbookKey,
    workbookKey,
    workbookKey,
    pattern,
    limit,
  ]);
  stmt.finalize();

  return mapRecentChatRows(rows);
}
