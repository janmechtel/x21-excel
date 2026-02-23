import { useState } from "react";
import {
  ChevronRight,
  FileText,
  Loader2,
  Sparkles,
  Pencil,
  Trash2,
  File,
} from "lucide-react";
import { SearchableOverlay } from "@/components/shared/SearchableOverlay";
import { ExcelRangePill } from "@/components/excel/ExcelRangePill";
import type { ActivityLog, ActivitySummary } from "@/types/chat";
import { rangePattern } from "@/utils/toolDisplay";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { SaveCopiesCheckbox } from "@/components/settings/SaveCopiesCheckbox";
import { webViewBridge } from "@/services/webViewBridge";

interface ActivityPanelProps {
  open: boolean;
  onClose: () => void;
  logs: ActivityLog[];
  onRangeClick: (range: string) => void;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
  isGenerating?: boolean;
  isLoading?: boolean;
  onGenerateSummary?: (comparisonFilePath?: string) => void;
  onDeleteLog?: (logId: string) => void;
  onEditLog?: (logId: string, newRawSummaryText: string) => void;
}

const formatLogTimestamp = (timestamp: number): string => {
  const date = new Date(timestamp);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);
  const diffHours = Math.floor(diffMs / 3600000);
  const diffDays = Math.floor(diffMs / 86400000);

  // Recent timestamps (< 1 hour)
  if (diffMins < 1) return "just now";
  if (diffMins < 60) return `${diffMins}m ago`;

  // Today (< 24 hours)
  if (diffHours < 24 && date.getDate() === now.getDate()) {
    return date.toLocaleTimeString("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    });
  }

  // This week
  if (diffDays < 7) {
    return date.toLocaleDateString("en-US", {
      weekday: "short",
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
    });
  }

  // Older
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
};

// Strip AI reasoning and noise from text
const stripNoise = (text: string): string => {
  return (
    text
      // Remove AI reasoning prefixes
      .replace(
        /^(?:Based on|According to|From|Looking at|I can see that|The diff shows that|Previously)\s+.+?,\s*/i,
        "",
      )
      .replace(/^(?:This indicates|This suggests|This means)\s+that\s+/i, "")
      // Remove meta-commentary
      .replace(/\s+\((?:rows?\s+\d+[-–]\d+|spans?[^)]+)\)/g, "")
      // Clean up
      .trim()
  );
};

// Extract collapsed (single-line) vs expanded (detailed) text
interface ProcessedChange {
  collapsed: string;
  expanded: string;
  type: "add" | "edit" | "delete" | "rename" | "duplicate" | "fill" | "other";
  entity?: string; // The thing that changed (column name, etc)
  scope?: string; // Where it changed (table name, sheet, etc)
}

const processChangeText = (description: string): ProcessedChange => {
  // Clean markdown
  const clean = description
    .replace(/^-\s*\[\s*[xX]?\s*\]\s*/gm, "")
    .replace(/^[•\-]\s+/gm, "")
    .trim();

  // Check if LLM provided collapsed | expanded format
  const dualFormatMatch = clean.match(/^(.+?)\s*\|\s*(.+)$/);
  if (dualFormatMatch) {
    const [, collapsedText, expandedText] = dualFormatMatch;
    // Detect type from expanded text for icon
    const type = detectChangeType(expandedText);
    const scope = extractScope(expandedText);

    // Enforce max length for collapsed (safety measure) - matches prompt's 40 char limit
    let finalCollapsed = collapsedText.trim();
    if (finalCollapsed.length > 45) {
      // Truncate at word boundary (only if LLM ignores the 40-char instruction)
      const truncated = finalCollapsed.substring(0, 40);
      const lastSpace = truncated.lastIndexOf(" ");
      finalCollapsed =
        (lastSpace > 25 ? truncated.substring(0, lastSpace) : truncated) + "…";
    }

    return {
      collapsed: finalCollapsed,
      expanded: expandedText.trim(),
      type,
      scope,
    };
  }

  // Fallback: generate collapsed/expanded from single text
  const noNoise = stripNoise(clean);

  // Pattern matching for type detection and entity extraction
  const patterns = [
    {
      regex:
        /(?:added|created|introduced)\s+(?:a\s+)?["']([^"']+)["']\s+(?:column|attribute|metric|field)(?:\s+(?:to|in)\s+(?:the\s+)?(.+?))?(?:\s+to\s+(.+?))?$/i,
      type: "add" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: m[1],
        scope: m[2] || m[3],
        collapsed: `Added ${m[1]}`,
        expanded: noNoise,
      }),
    },
    {
      regex:
        /(?:renamed|changed)\s+(?:the\s+)?["']([^"']+)["']\s+(?:to|→)\s+["']([^"']+)["'](?:\s+in\s+(?:the\s+)?(.+?))?$/i,
      type: "rename" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: `${m[1]} → ${m[2]}`,
        scope: m[3],
        collapsed: `${m[1]} → ${m[2]}`,
        expanded: noNoise,
      }),
    },
    {
      regex:
        /(?:updated|modified|changed)\s+(?:the\s+)?["']([^"']+)["'](?:\s+in\s+(?:the\s+)?(.+?))?$/i,
      type: "edit" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: m[1],
        scope: m[2],
        collapsed: `Updated ${m[1]}`,
        expanded: noNoise,
      }),
    },
    {
      regex:
        /(?:deleted|removed)\s+(?:the\s+)?["']([^"']+)["'](?:\s+from\s+(?:the\s+)?(.+?))?$/i,
      type: "delete" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: m[1],
        scope: m[2],
        collapsed: `Deleted ${m[1]}`,
        expanded: noNoise,
      }),
    },
    {
      regex:
        /duplicated?\s+(?:the\s+)?["']([^"']+)["'](?:\s+in\s+(?:the\s+)?(.+?))?$/i,
      type: "duplicate" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: m[1],
        scope: m[2],
        collapsed: `Duplicated ${m[1]}`,
        expanded: noNoise,
      }),
    },
    {
      regex:
        /(?:filled|populated)\s+(?:missing\s+)?(?:values?\s+)?(?:in\s+)?(?:the\s+)?["']([^"']+)["'](?:\s+in\s+(?:the\s+)?(.+?))?$/i,
      type: "fill" as const,
      extract: (m: RegExpMatchArray) => ({
        entity: m[1],
        scope: m[2],
        collapsed: `Filled ${m[1]}`,
        expanded: noNoise,
      }),
    },
  ];

  for (const pattern of patterns) {
    const match = noNoise.match(pattern.regex);
    if (match) {
      const result = pattern.extract(match);
      return {
        type: pattern.type,
        entity: result.entity,
        scope: result.scope,
        collapsed: result.collapsed,
        expanded: result.expanded,
      };
    }
  }

  // Fallback: truncate intelligently
  const truncated = truncateIntelligently(noNoise, 70);
  return {
    type: "other",
    collapsed: truncated,
    expanded: noNoise,
  };
};

// Detect change type from text (for icon selection)
const detectChangeType = (text: string): ProcessedChange["type"] => {
  if (/\b(added|created|introduced)\b/i.test(text)) return "add";
  if (/\b(renamed|changed)\b/i.test(text)) return "rename";
  if (/\b(updated|modified)\b/i.test(text)) return "edit";
  if (/\b(deleted|removed)\b/i.test(text)) return "delete";
  if (/\b(duplicated|copied)\b/i.test(text)) return "duplicate";
  if (/\b(filled|populated)\b/i.test(text)) return "fill";
  return "other";
};

// Extract scope from text (table/dataset name)
const extractScope = (text: string): string | undefined => {
  // Look for "in/to the X table/dataset"
  const scopeMatch = text.match(
    /(?:in|to|from)\s+(?:the\s+)?([^.]+?)\s+(?:table|dataset|feature|category|sheet)/i,
  );
  if (scopeMatch) {
    return scopeMatch[1].trim();
  }
  return undefined;
};

// Truncate at natural boundaries
const truncateIntelligently = (text: string, maxLen: number): string => {
  if (text.length <= maxLen) return text;

  // Try to break at clause boundary
  const clauseMatch = text
    .substring(0, maxLen + 20)
    .match(/^(.{20,}?)(?:\s+(?:to|for|in order to|that|which)\s+)/);
  if (clauseMatch && clauseMatch[1].length <= maxLen) {
    return clauseMatch[1].trim();
  }

  // Break at word boundary
  const truncated = text.substring(0, maxLen);
  const lastSpace = truncated.lastIndexOf(" ");
  return (
    (lastSpace > maxLen * 0.7 ? truncated.substring(0, lastSpace) : truncated) +
    "…"
  );
};

const renderTextWithRangePills = (
  text: string,
  onRangeClick: (range: string) => void,
) => {
  const matches = text.match(rangePattern);
  if (!matches || matches.length === 0) {
    return text;
  }

  const parts: (string | JSX.Element)[] = [];
  let remaining = text;
  let key = 0;

  matches.forEach((match) => {
    const index = remaining.indexOf(match);
    if (index > 0) {
      parts.push(remaining.substring(0, index));
    }
    parts.push(
      <ExcelRangePill
        key={`range-${key++}`}
        range={match}
        onClick={onRangeClick}
        className="mx-0.5"
      />,
    );
    remaining = remaining.substring(index + match.length);
  });

  if (remaining.length > 0) {
    parts.push(remaining);
  }

  return <>{parts}</>;
};

// Hierarchical change row - supports nested children
function HierarchicalChangeRow({
  summary,
  isExpanded,
  onToggle,
  onRangeClick,
  path,
  expandedPaths,
  onTogglePath,
}: {
  summary: ActivitySummary;
  isExpanded: boolean;
  onToggle: () => void;
  onRangeClick: (range: string) => void;
  path: number[];
  expandedPaths: Set<number | string>;
  onTogglePath: (path: number[]) => void;
}) {
  const processed = processChangeText(summary.description);
  const level = summary.level || 0;
  const hasChildren = summary.children && summary.children.length > 0;
  const pathKey = path.join("-");

  return (
    <div className="border-b border-slate-200 dark:border-slate-700 last:border-0">
      {/* Collapsed: ONE LINE ONLY with indentation for hierarchy */}
      <button
        onClick={onToggle}
        className={`w-full text-left flex items-center gap-2.5 py-2.5 px-3 hover:bg-slate-50/70 dark:hover:bg-slate-800/50 transition-colors group/row ${
          level > 0 ? "bg-slate-50/30 dark:bg-slate-800/10" : ""
        }`}
        style={{ paddingLeft: `${12 + level * 10}px` }}
      >
        {/* Chevron - always show for all items (they can all be expanded) */}
        <ChevronRight
          className={`w-3.5 h-3.5 text-slate-400 dark:text-slate-500 transition-transform flex-shrink-0 ${
            isExpanded ? "rotate-90" : ""
          }`}
        />

        {/* Primary: <Verb> <object> - clean, no scope */}
        <span className="text-[13px] text-slate-700 dark:text-slate-200 truncate flex-1">
          {renderTextWithRangePills(processed.collapsed, onRangeClick)}
        </span>
      </button>

      {/* Expanded: inline details and children */}
      {isExpanded && (
        <div className="space-y-0">
          {/* Expanded description */}
          <div
            className="px-3 pb-2 pt-1 space-y-1 text-[12px] leading-relaxed text-slate-600 dark:text-slate-300 bg-slate-50/30 dark:bg-slate-800/20"
            style={{ paddingLeft: `${12 + level * 10 + 20}px` }}
          >
            {renderTextWithRangePills(processed.expanded, onRangeClick)}
          </div>

          {/* Render children if they exist */}
          {hasChildren && (
            <div className="border-l-2 border-slate-200 dark:border-slate-700 ml-1">
              {summary.children!.map((child, childIdx) => (
                <HierarchicalChangeRow
                  key={childIdx}
                  summary={child}
                  isExpanded={expandedPaths.has(`${pathKey}-${childIdx}`)}
                  onToggle={() => onTogglePath([...path, childIdx])}
                  onRangeClick={onRangeClick}
                  path={[...path, childIdx]}
                  expandedPaths={expandedPaths}
                  onTogglePath={onTogglePath}
                />
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function LogItem({
  log,
  onRangeClick,
  onDeleteLog,
  onEditLog,
}: {
  log: ActivityLog;
  onRangeClick: (range: string) => void;
  onDeleteLog?: (logId: string) => void;
  onEditLog?: (logId: string, newRawSummaryText: string) => void;
}) {
  const [expandedPaths, setExpandedPaths] = useState<Set<string | number>>(
    new Set(),
  );
  const [isEditing, setIsEditing] = useState(false);
  const [editText, setEditText] = useState(log.rawSummaryText || "");
  const [confirmDelete, setConfirmDelete] = useState(false);

  const toggleExpanded = (idx: number) => {
    const pathKey = idx.toString();
    setExpandedPaths((prev) => {
      const next = new Set(prev);
      if (next.has(pathKey)) {
        next.delete(pathKey);
      } else {
        next.add(pathKey);
      }
      return next;
    });
  };

  const togglePath = (path: number[]) => {
    const pathKey = path.join("-");
    setExpandedPaths((prev) => {
      const next = new Set(prev);
      if (next.has(pathKey)) {
        next.delete(pathKey);
      } else {
        next.add(pathKey);
      }
      return next;
    });
  };

  const handleDelete = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!onDeleteLog) return;
    setConfirmDelete(true);
  };

  const handleEditToggle = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!isEditing) {
      setEditText(log.rawSummaryText || "");
    }
    setIsEditing((prev) => !prev);
  };

  const handleEditSave = () => {
    if (!onEditLog) {
      setIsEditing(false);
      return;
    }
    const trimmed = editText.trim();
    if (!trimmed) {
      // Don't allow saving completely empty text
      return;
    }
    onEditLog(log.id, trimmed);
    setIsEditing(false);
  };

  const handleDeleteConfirm = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (!onDeleteLog) return;
    onDeleteLog(log.id);
    setConfirmDelete(false);
  };

  const handleDeleteCancel = (e: React.MouseEvent) => {
    e.stopPropagation();
    setConfirmDelete(false);
  };

  const handleEditCancel = () => {
    setEditText(log.rawSummaryText || "");
    setIsEditing(false);
  };

  // Handle no-change / empty case
  const isEmptyLog =
    log.data.summaries.length === 0 ||
    (log.data.summaries.length === 1 &&
      /no\s+(?:meaningful\s+)?changes?\s+(?:detected|found)/i.test(
        log.data.summaries[0].description,
      ));

  if (isEmptyLog) {
    return (
      <div className="group/log border-b border-slate-100 dark:border-slate-800">
        <div className="flex items-center justify-between px-3 py-1.5 bg-slate-50/30 dark:bg-slate-800/20">
          <div className="flex items-center gap-2 min-w-0">
            <time className="shrink-0 whitespace-nowrap text-[10px] font-medium text-slate-400 dark:text-slate-500 uppercase tracking-wide">
              {formatLogTimestamp(log.timestamp)}
            </time>
            {!!log.comparisonFileName && (
              <Tooltip>
                <TooltipTrigger asChild>
                  <div
                    className="flex items-center gap-1 text-[10px] text-slate-500 dark:text-slate-400 min-w-0 max-w-[160px] sm:max-w-[220px] whitespace-nowrap"
                    title={log.comparisonFileName}
                  >
                    <File className="w-3 h-3 shrink-0 opacity-70" />
                    <span className="truncate">{log.comparisonFileName}</span>
                  </div>
                </TooltipTrigger>
                <TooltipContent>{log.comparisonFileName}</TooltipContent>
              </Tooltip>
            )}
          </div>
          <div className="flex items-center gap-0.5 opacity-0 group-hover/log:opacity-100 transition-opacity">
            {confirmDelete && (
              <div className="flex items-center gap-1 mr-1 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2 py-1 text-[10px] text-slate-700 dark:text-slate-200">
                <span>Remove?</span>
                <button
                  onClick={handleDeleteConfirm}
                  className="text-red-600 hover:text-red-700 font-semibold"
                >
                  Yes
                </button>
                <button
                  onClick={handleDeleteCancel}
                  className="text-slate-500 hover:text-slate-700"
                >
                  No
                </button>
              </div>
            )}
            {onEditLog && (
              <button
                onClick={handleEditToggle}
                className="p-0.5 rounded text-slate-300 hover:text-slate-500 dark:text-slate-600 dark:hover:text-slate-400 transition-colors"
              >
                <Pencil className="w-2.5 h-2.5" />
              </button>
            )}
            {onDeleteLog && (
              <button
                onClick={handleDelete}
                className="p-0.5 rounded text-slate-300 hover:text-red-500 dark:text-slate-600 dark:hover:text-red-400 transition-colors"
              >
                <Trash2 className="w-2.5 h-2.5" />
              </button>
            )}
          </div>
        </div>

        {isEditing ? (
          <div className="px-3 py-2 bg-slate-50/60 dark:bg-slate-800/40 space-y-2">
            <p className="text-[11px] text-slate-500 dark:text-slate-400">
              Edit the raw changelog text. Changes will be saved for this entry.
            </p>
            <textarea
              className="w-full h-24 text-[12px] font-mono rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 px-2 py-1.5 resize-vertical"
              value={editText}
              onChange={(e) => setEditText(e.target.value)}
            />
            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={handleEditCancel}
                className="px-2.5 py-1 text-[11px] rounded border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleEditSave}
                className="px-2.5 py-1 text-[11px] rounded bg-blue-600 text-white hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
                disabled={!editText.trim()}
              >
                Save
              </button>
            </div>
          </div>
        ) : (
          <div className="px-3 py-2 text-[13px] text-slate-400 dark:text-slate-500">
            No meaningful changes detected
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="group/log">
      {/* Timestamp + feedback - compact header */}
      <div className="flex items-center justify-between px-3 py-1.5 bg-slate-50/30 dark:bg-slate-800/20 border-b border-slate-100 dark:border-slate-800">
        <div className="flex items-center gap-2 min-w-0">
          <time className="shrink-0 whitespace-nowrap text-[10px] font-medium text-slate-400 dark:text-slate-500 uppercase tracking-wide">
            {formatLogTimestamp(log.timestamp)}
          </time>
          {!!log.comparisonFileName && (
            <Tooltip>
              <TooltipTrigger asChild>
                <div
                  className="flex items-center gap-1 text-[10px] text-slate-500 dark:text-slate-400 min-w-0 max-w-[160px] sm:max-w-[220px] whitespace-nowrap"
                  title={log.comparisonFileName}
                >
                  <File className="w-3 h-3 shrink-0 opacity-70" />
                  <span className="truncate">{log.comparisonFileName}</span>
                </div>
              </TooltipTrigger>
              <TooltipContent>{log.comparisonFileName}</TooltipContent>
            </Tooltip>
          )}
        </div>
        <div className="flex items-center gap-0.5 opacity-0 group-hover/log:opacity-100 transition-opacity">
          {confirmDelete && (
            <div className="flex items-center gap-1 mr-1 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2 py-1 text-[10px] text-slate-700 dark:text-slate-200">
              <span>Remove?</span>
              <button
                onClick={handleDeleteConfirm}
                className="text-red-600 hover:text-red-700 font-semibold"
              >
                Yes
              </button>
              <button
                onClick={handleDeleteCancel}
                className="text-slate-500 hover:text-slate-700"
              >
                No
              </button>
            </div>
          )}
          {onEditLog && (
            <button
              onClick={handleEditToggle}
              className="p-0.5 rounded text-slate-300 hover:text-slate-500 dark:text-slate-600 dark:hover:text-slate-400 transition-colors"
            >
              <Pencil className="w-2.5 h-2.5" />
            </button>
          )}
          {onDeleteLog && (
            <button
              onClick={handleDelete}
              className="p-0.5 rounded text-slate-300 hover:text-red-500 dark:text-slate-600 dark:hover:text-red-400 transition-colors"
            >
              <Trash2 className="w-2.5 h-2.5" />
            </button>
          )}
        </div>
      </div>

      {/* Either show raw editable text or parsed hierarchical changes */}
      {isEditing ? (
        <div className="px-3 py-2 bg-slate-50/60 dark:bg-slate-800/40 border-t border-slate-100 dark:border-slate-800 space-y-2">
          <p className="text-[11px] text-slate-500 dark:text-slate-400">
            Edit the raw changelog text (checkbox list). Changes will be saved
            for this entry only.
          </p>
          <textarea
            className="w-full h-28 text-[12px] font-mono rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 px-2 py-1.5 resize-vertical"
            value={editText}
            onChange={(e) => setEditText(e.target.value)}
          />
          <div className="flex justify-end gap-2">
            <button
              type="button"
              onClick={handleEditCancel}
              className="px-2.5 py-1 text-[11px] rounded border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800"
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={handleEditSave}
              className="px-2.5 py-1 text-[11px] rounded bg-blue-600 text-white hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
              disabled={!editText.trim()}
            >
              Save
            </button>
          </div>
        </div>
      ) : (
        <div>
          {log.data.summaries.map((summary, idx) => (
            <HierarchicalChangeRow
              key={idx}
              summary={summary}
              isExpanded={expandedPaths.has(idx.toString())}
              onToggle={() => toggleExpanded(idx)}
              onRangeClick={onRangeClick}
              path={[idx]}
              expandedPaths={expandedPaths}
              onTogglePath={togglePath}
            />
          ))}
        </div>
      )}
    </div>
  );
}

export function ActivityPanel({
  open,
  onClose,
  logs,
  onRangeClick,
  textareaRef,
  isGenerating = false,
  isLoading = false,
  onGenerateSummary,
  onDeleteLog,
  onEditLog,
}: ActivityPanelProps) {
  const [searchValue, setSearchValue] = useState("");
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const [saveSnapshotsEnabled, setSaveSnapshotsEnabled] = useState(false);
  const [sourceMode, setSourceMode] = useState<"current" | "other">("current");
  const [selectedFile, setSelectedFile] = useState<{
    name: string;
    path: string;
  } | null>(null);

  // Recursively search in summary and its children
  const matchesSearch = (
    summary: ActivitySummary,
    searchLower: string,
  ): boolean => {
    // Check if this summary matches
    if (
      summary.description.toLowerCase().includes(searchLower) ||
      summary.details.some(
        (detail) =>
          detail.sheet?.toLowerCase().includes(searchLower) ||
          detail.cell?.toLowerCase().includes(searchLower) ||
          detail.range?.toLowerCase().includes(searchLower),
      )
    ) {
      return true;
    }

    // Recursively check children
    if (summary.children && summary.children.length > 0) {
      return summary.children.some((child) =>
        matchesSearch(child, searchLower),
      );
    }

    return false;
  };

  // Filter logs based on search
  const filteredLogs = logs.filter((log) => {
    if (!searchValue) return true;
    const searchLower = searchValue.toLowerCase();

    // Search in summaries (recursively including children)
    return log.data.summaries.some((summary) =>
      matchesSearch(summary, searchLower),
    );
  });

  const handleSelect = () => {
    // Nothing to select in logs panel, just close
    onClose();
  };

  const handleGenerateClick = () => {
    if (onGenerateSummary && !isGenerating) {
      // Pass filepath if in "other file" mode and a file is selected
      const filePath =
        sourceMode === "other" && selectedFile ? selectedFile.path : undefined;
      onGenerateSummary(filePath);
    }
  };

  const handleSourceModeChange = async (mode: "current" | "other") => {
    if (mode === "other") {
      // Switch to other file mode and open file picker
      setSourceMode("other");
      await handleFilePicker();
    } else {
      // Switch back to current file mode - forget external file
      setSourceMode("current");
      setSelectedFile(null);
    }
  };

  const handleFilePicker = async () => {
    try {
      const filePath = await webViewBridge.pickFile({
        extensions: [".xlsx", ".xlsm", ".xls"],
      });
      if (filePath) {
        const fileName = filePath.split(/[/\\]/).pop() || filePath;
        setSelectedFile({
          name: fileName,
          path: filePath,
        });
      }
    } catch (error) {
      console.warn("Failed to pick file:", error);
    }
  };

  // Determine if generate button should be enabled
  const isGenerateEnabled =
    sourceMode === "other" ? selectedFile !== null : saveSnapshotsEnabled;

  return (
    <SearchableOverlay
      isOpen={open}
      onClose={onClose}
      title="Activity"
      icon={FileText}
      searchValue={searchValue}
      onSearchChange={setSearchValue}
      searchPlaceholder="Search changes..."
      highlightedIndex={highlightedIndex}
      itemCount={filteredLogs.length}
      onHighlight={setHighlightedIndex}
      onSelect={handleSelect}
      textareaRef={textareaRef}
      maxWidth="max-w-2xl"
      maxHeight="max-h-[85vh]"
      emptyStateMessage={
        isGenerating || isLoading ? (
          ""
        ) : logs.length === 0 ? (
          <div className="px-4 py-5 text-xs text-slate-500 dark:text-slate-400">
            No activity yet
          </div>
        ) : (
          <div className="px-4 py-5 text-xs text-slate-500 dark:text-slate-400">
            No matching changes
          </div>
        )
      }
      searchBarActions={
        onGenerateSummary && (
          <div className="flex flex-col gap-1.5">
            {/* Source selection buttons */}
            <div className="px-4 py-2 border-b border-slate-200 dark:border-slate-700 bg-slate-50/30 dark:bg-slate-800/20">
              <div className="flex items-center gap-2 mb-2">
                <span className="text-[11px] font-medium text-slate-600 dark:text-slate-400 uppercase tracking-wide">
                  Source
                </span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleSourceModeChange("current")}
                  className={`px-2.5 py-1 rounded text-xs font-medium transition-colors ${
                    sourceMode === "current"
                      ? "bg-blue-500 text-white shadow-sm"
                      : "bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-300 dark:hover:bg-slate-600"
                  }`}
                >
                  Current file
                </button>
                <button
                  type="button"
                  onClick={() => handleSourceModeChange("other")}
                  className={`px-2.5 py-1 rounded text-xs font-medium transition-colors ${
                    sourceMode === "other"
                      ? "bg-blue-500 text-white shadow-sm"
                      : "bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-300 dark:hover:bg-slate-600"
                  }`}
                >
                  Other file
                </button>
              </div>

              {/* Show checkbox only in current file mode - keep mounted to avoid DB reload */}
              <div
                className={`mt-1.5 ${sourceMode === "other" ? "hidden" : ""}`}
              >
                <SaveCopiesCheckbox
                  onChange={setSaveSnapshotsEnabled}
                  size="small"
                />
              </div>

              {/* Show selected file in other file mode */}
              {sourceMode === "other" && selectedFile && (
                <div className="mt-2.5 flex items-center gap-2">
                  <File className="w-3.5 h-3.5 text-slate-500 dark:text-slate-400 flex-shrink-0" />
                  <span className="text-xs text-slate-700 dark:text-slate-300 flex-1 truncate">
                    {selectedFile.name}
                  </span>
                  <button
                    type="button"
                    onClick={handleFilePicker}
                    className="text-xs text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 font-medium flex-shrink-0"
                  >
                    Change
                  </button>
                </div>
              )}
            </div>

            {/* Generate button */}
            <div className="px-4 pb-2">
              <Tooltip>
                <TooltipTrigger asChild>
                  <button
                    type="button"
                    onClick={handleGenerateClick}
                    disabled={isGenerating || !isGenerateEnabled}
                    className="w-full border-2 border-blue-500 bg-blue-50 dark:bg-blue-950/30 text-blue-700 dark:text-blue-300 hover:bg-blue-100 dark:hover:bg-blue-950/50 transition-all disabled:opacity-50 disabled:cursor-not-allowed disabled:border-slate-300 dark:disabled:border-slate-600 disabled:bg-slate-50 dark:disabled:bg-slate-800/50 disabled:text-slate-400 dark:disabled:text-slate-500 flex items-center justify-center gap-1.5 px-2.5 py-1.5 rounded-md font-semibold text-xs shadow-sm hover:shadow-md"
                    aria-label="Generate summary"
                  >
                    {isGenerating ? (
                      <>
                        <Loader2 className="w-3.5 h-3.5 animate-spin" />
                        <span>Generating...</span>
                      </>
                    ) : (
                      <>
                        <Sparkles className="w-3.5 h-3.5" />
                        <span>Generate summary</span>
                      </>
                    )}
                  </button>
                </TooltipTrigger>
                <TooltipContent>
                  <p>
                    {isGenerating
                      ? "Generating summary..."
                      : sourceMode === "other"
                      ? "Generate summary comparing to selected file"
                      : "Summarize recent workbook changes"}
                  </p>
                </TooltipContent>
              </Tooltip>
            </div>

            {/* Helper text only in current file mode when checkbox is unchecked */}
            {sourceMode === "current" && !saveSnapshotsEnabled && (
              <p className="text-[10px] text-slate-500 dark:text-slate-400 px-4 pb-1.5">
                Requires milestone copies
              </p>
            )}
          </div>
        )
      }
    >
      {(isGenerating || isLoading) && logs.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-12 px-6">
          <Loader2 className="w-8 h-8 text-slate-400 dark:text-slate-500 animate-spin mb-4" />
          <p className="text-sm text-slate-600 dark:text-slate-400">
            {isLoading ? "Loading activity..." : "Analyzing changes..."}
          </p>
        </div>
      ) : filteredLogs.length > 0 ? (
        <div className="divide-y divide-slate-200 dark:divide-slate-700">
          {filteredLogs.map((log) => (
            <LogItem
              key={log.id}
              log={log}
              onRangeClick={onRangeClick}
              onDeleteLog={onDeleteLog}
              onEditLog={onEditLog}
            />
          ))}
        </div>
      ) : null}
    </SearchableOverlay>
  );
}
