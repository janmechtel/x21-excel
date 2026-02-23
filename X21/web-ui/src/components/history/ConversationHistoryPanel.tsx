import { useEffect, useState } from "react";
import { ChevronDown, History, Loader2, Pencil } from "lucide-react";
import { webViewBridge } from "@/services/webViewBridge";
import {
  type HistoryChatMessage,
  listRecentConversations,
  loadConversationForFile,
  searchConversations,
} from "@/services/conversationHistoryService";
import type { RecentChatItem } from "@/services/historyService";
import { hasAttachmentsInContent } from "@/services/historyService";
import { SearchableOverlay } from "@/components/shared/SearchableOverlay";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import type { ToolDecisionData } from "@/types/chat";

function extractUserMessage(raw: string): string {
  // Try parse as JSON and extract userMessage if present
  try {
    const parsed = JSON.parse(raw);
    if (
      parsed &&
      typeof parsed === "object" &&
      typeof (parsed as any).userMessage === "string"
    ) {
      return (parsed as any).userMessage;
    }
  } catch {
    // Ignore parse errors; fall back to heuristic extraction
  }

  // Heuristic: handle truncated JSON that still contains "userMessage":"..."
  const match = raw.match(/"userMessage"\s*:\s*"([^"]*)/);
  if (match && match[1]) {
    return match[1];
  }

  // Fallback to raw string (trimmed)
  return raw.trim();
}

function formatPreview(raw: string): string {
  if (!raw) return "(empty prompt)";
  return extractUserMessage(raw) || "(empty prompt)";
}

function formatEditPrompt(raw: string): string {
  if (!raw) return "";
  return extractUserMessage(raw);
}

function formatTimeAgo(timestamp: number): string {
  const now = Date.now();
  const diffMs = now - timestamp;
  const diffSeconds = Math.max(0, Math.floor(diffMs / 1000));

  if (diffSeconds < 60) return "Just now";

  const diffMinutes = Math.floor(diffSeconds / 60);
  if (diffMinutes < 60) return `${diffMinutes}m`;

  const diffHours = Math.floor(diffMinutes / 60);
  if (diffHours < 24) return `${diffHours}h`;

  const diffDays = Math.floor(diffHours / 24);
  return `${diffDays}d`;
}

interface ConversationHistoryPanelProps {
  open: boolean;
  onClose: () => void;
  onConversationSelected: (conversationId: string) => void;
  onConversationLoaded: (
    messages: HistoryChatMessage[],
    conversationId: string,
    toolDecisions?: Map<string, ToolDecisionData>,
  ) => void;
  onConversationLoadError: () => void;
  activeConversationId: string | null;
  onEditPrompt: (prompt: string) => void;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
}

export function ConversationHistoryPanel({
  open,
  onClose,
  onConversationSelected,
  onConversationLoaded,
  onConversationLoadError,
  activeConversationId,
  onEditPrompt,
  textareaRef,
}: ConversationHistoryPanelProps) {
  const PAGE_SIZE = 10;
  const [recentChats, setRecentChats] = useState<RecentChatItem[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const [scope, setScope] = useState<"file" | "all">("file");
  const [isScopeMenuOpen, setIsScopeMenuOpen] = useState(false);

  useEffect(() => {
    if (!open) return;

    let cancelled = false;

    const load = async () => {
      setIsLoading(true);
      setError(null);

      try {
        const workbookKey =
          scope === "file" ? await webViewBridge.getWorkbookPath() : null;

        if (scope === "file" && !workbookKey) {
          if (!cancelled) {
            setError("No workbook path found for this Excel workbook.");
            setRecentChats([]);
          }
          return;
        }

        const normalizedQuery = searchQuery.trim();
        const items = normalizedQuery
          ? await searchConversations(workbookKey, normalizedQuery, 100, scope)
          : await listRecentConversations(workbookKey, 100, scope);

        if (!cancelled) {
          setRecentChats(items);
          setCurrentPage(1);
        }
      } catch (_e) {
        if (!cancelled) {
          setError("Failed to load history. Please try again.");
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    };

    void load();

    return () => {
      cancelled = true;
    };
  }, [open, scope, searchQuery]);

  useEffect(() => {
    if (open) {
      setHighlightedIndex(0);
    }
  }, [open, searchQuery, recentChats]);

  const handleSelectConversation = async (chat: RecentChatItem) => {
    // Immediately update UI state - simple state update like tools panel
    // This closes the panel instantly, matching tools panel behavior
    onConversationSelected(chat.conversationId);

    // Load messages in the background (don't block UI)
    try {
      let workbookKey: string | null;
      if (scope === "file") {
        workbookKey = await webViewBridge.getWorkbookPath();
      } else {
        workbookKey = chat.workbookKey;
      }

      if (!workbookKey) {
        // Reset the selection if we can't load
        onConversationLoadError();
        return;
      }

      // Load messages in the background
      const { messages, toolDecisions } = await loadConversationForFile(
        workbookKey,
        chat.conversationId,
      );
      // Update chat history with loaded messages
      onConversationLoaded(messages, chat.conversationId, toolDecisions);
    } catch (_e) {
      // Reset the selection on error
      onConversationLoadError();
    }
  };

  const normalizedQuery = searchQuery.trim();

  let totalPages: number;
  let safePage: number;
  let pageItems: RecentChatItem[];

  if (normalizedQuery) {
    // When searching (server-side), ignore pagination and show all results
    totalPages = 1;
    safePage = 1;
    pageItems = recentChats;
  } else {
    totalPages =
      recentChats.length === 0 ? 1 : Math.ceil(recentChats.length / PAGE_SIZE);
    safePage = Math.min(currentPage, totalPages);
    const startIndex = (safePage - 1) * PAGE_SIZE;
    pageItems = recentChats.slice(startIndex, startIndex + PAGE_SIZE);
  }

  const handleSelect = () => {
    if (pageItems[highlightedIndex]) {
      void handleSelectConversation(pageItems[highlightedIndex]);
    }
  };

  const searchBarActions = (
    <div className="relative inline-block text-left text-[11px]">
      <button
        type="button"
        onClick={() => setIsScopeMenuOpen((open) => !open)}
        className="inline-flex items-center gap-1 px-3 py-1 border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/60 text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
      >
        <span>{scope === "all" ? "All files" : "This file"}</span>
        <ChevronDown className="w-3 h-3" />
      </button>
      {isScopeMenuOpen && (
        <div className="absolute mt-1 w-36 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 shadow-lg z-[80]">
          <button
            type="button"
            onClick={() => {
              setScope("file");
              setIsScopeMenuOpen(false);
            }}
            className={`block w-full text-left px-3 py-1.5 text-[11px] ${
              scope === "file"
                ? "bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-50 font-medium"
                : "text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800"
            }`}
          >
            This file
          </button>
          <button
            type="button"
            onClick={() => {
              setScope("all");
              setIsScopeMenuOpen(false);
            }}
            className={`block w-full text-left px-3 py-1.5 text-[11px] ${
              scope === "all"
                ? "bg-slate-100 dark:bg-slate-800 text-slate-900 dark:text-slate-50 font-medium"
                : "text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800"
            }`}
          >
            All files
          </button>
        </div>
      )}
    </div>
  );

  const footer =
    !isLoading && !error && !normalizedQuery && recentChats.length > 0 ? (
      <div className="px-3 py-1.5 flex items-center justify-between text-[11px] text-slate-500 dark:text-slate-400">
        <span>
          Page {safePage} of {totalPages}
        </span>
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            disabled={safePage <= 1}
            className={`px-2 py-0.5 border text-[11px] ${
              safePage <= 1
                ? "border-slate-200 text-slate-300 dark:border-slate-700 dark:text-slate-600 cursor-not-allowed"
                : "border-slate-300 text-slate-600 dark:border-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800"
            }`}
          >
            Prev
          </button>
          <button
            type="button"
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            disabled={safePage >= totalPages}
            className={`px-2 py-0.5 border text-[11px] ${
              safePage >= totalPages
                ? "border-slate-200 text-slate-300 dark:border-slate-700 dark:text-slate-600 cursor-not-allowed"
                : "border-slate-300 text-slate-600 dark:border-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800"
            }`}
          >
            Next
          </button>
        </div>
      </div>
    ) : undefined;

  return (
    <SearchableOverlay
      isOpen={open}
      onClose={onClose}
      title="Conversation history"
      icon={History}
      searchValue={searchQuery}
      onSearchChange={setSearchQuery}
      searchPlaceholder="Search recent tasks"
      highlightedIndex={highlightedIndex}
      itemCount={pageItems.length}
      onHighlight={setHighlightedIndex}
      onSelect={handleSelect}
      searchBarActions={searchBarActions}
      footer={footer}
      textareaRef={textareaRef}
    >
      {isLoading && (
        <div className="flex items-center justify-center py-6 text-xs text-slate-500 dark:text-slate-400 gap-2">
          <Loader2 className="w-3 h-3 animate-spin" />
          <span>Loading conversations…</span>
        </div>
      )}
      {!isLoading && error && (
        <div className="px-3 py-3 text-xs text-red-600 dark:text-red-400">
          {error}
        </div>
      )}
      {!isLoading && !error && recentChats.length === 0 && !normalizedQuery && (
        <div className="px-3 py-3 text-xs text-slate-500 dark:text-slate-400">
          No previous conversations yet for this workbook.
        </div>
      )}
      {!isLoading && !error && normalizedQuery && recentChats.length === 0 && (
        <div className="px-3 py-3 text-xs text-slate-500 dark:text-slate-400">
          No conversations match your search.
        </div>
      )}
      {!isLoading && !error && pageItems.length > 0 && (
        <div className="py-1">
          {pageItems.map((chat, index) => {
            const isActive = index === highlightedIndex;
            const isAttachmentOnly =
              chat.firstMessagePreview === "(attachment)";
            const hasAttachmentMeta = hasAttachmentsInContent(
              chat.firstMessage,
            );
            const disableEdit = isAttachmentOnly || hasAttachmentMeta;
            return (
              <button
                key={chat.conversationId}
                type="button"
                onClick={(e) => {
                  e.preventDefault();
                  e.stopPropagation();
                  void handleSelectConversation(chat);
                }}
                onMouseEnter={() => setHighlightedIndex(index)}
                className={`w-full text-left px-3 py-2 text-xs border-b border-slate-100 dark:border-slate-800 transition-colors ${
                  isActive
                    ? "bg-blue-50 dark:bg-blue-900/30 text-blue-900 dark:text-blue-100"
                    : chat.conversationId === activeConversationId
                    ? "bg-slate-100 dark:bg-slate-800"
                    : "hover:bg-slate-50 dark:hover:bg-slate-800/60"
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <div className="font-medium text-slate-800 dark:text-slate-100 line-clamp-2">
                    {formatPreview(chat.firstMessagePreview)}
                  </div>
                  <div className="flex items-center gap-2 flex-shrink-0">
                    <span className="text-[10px] text-slate-500 dark:text-slate-400">
                      {formatTimeAgo(chat.lastActivityAt)}
                    </span>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <span>
                          <button
                            type="button"
                            disabled={disableEdit}
                            onClick={(e) => {
                              if (disableEdit) return;
                              e.stopPropagation();
                              onEditPrompt(formatEditPrompt(chat.firstMessage));
                              onClose();
                            }}
                            onKeyDown={(e) => {
                              if (disableEdit) return;
                              if (e.key === "Enter" || e.key === " ") {
                                e.preventDefault();
                                e.stopPropagation();
                                onEditPrompt(
                                  formatEditPrompt(chat.firstMessage),
                                );
                                onClose();
                              }
                            }}
                            className="p-1 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-500 dark:text-slate-400 cursor-pointer disabled:cursor-not-allowed"
                          >
                            <Pencil className="w-3 h-3" />
                            <span className="sr-only">Edit prompt</span>
                          </button>
                        </span>
                      </TooltipTrigger>
                      {disableEdit && (
                        <TooltipContent>
                          <p>
                            Edit is not available for prompts with attachments.
                          </p>
                        </TooltipContent>
                      )}
                    </Tooltip>
                  </div>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </SearchableOverlay>
  );
}
