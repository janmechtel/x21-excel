/* eslint-disable jsdoc/require-returns */
/* eslint-disable jsdoc/require-param */
/* eslint-disable jsdoc/require-jsdoc */
import {
  BookOpen,
  ChevronDown,
  ChevronRight,
  Columns,
  Folder,
  FileText,
  Palette,
  PenTool,
  Pipette,
  Rows,
  Trash2,
  Wrench,
  FileSpreadsheet,
} from "lucide-react";
import { useEffect, useState } from "react";
import type { Dispatch, ReactNode, SetStateAction } from "react";
import { ToolNames, ClaudeContentTypes } from "@/types/chat";

import { ActionButtons } from "@/components/chat/ActionButtons";
import { ExcelRangePill } from "@/components/excel/ExcelRangePill";
import { Textarea } from "@/components/ui/textarea";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@/components/ui/collapsible";
import { ChatMessage, ContentBlock, ToolDecisionData } from "@/types/chat";
import { rangePattern } from "@/utils/toolDisplay";
import { webViewBridge } from "@/services/webViewBridge";
import { formatWriteValuesAction } from "@/utils/toolStatus";
import {
  ColumnWidthModes,
  getColumnWidthMessage,
  type ColumnWidthMode,
} from "@/utils/columnWidth";
import { getApiBase } from "@/services/apiBase";

import { IconContainer, WaveAnimation } from "./Common";

interface ToolBlockProps {
  block: ContentBlock;
  message?: ChatMessage;
  isExpanded: boolean;
  onToggle: (id: string) => void;
  approvedTools: Set<string>;
  rejectedTools: Set<string>;
  viewedTools: Set<string>;
  erroredTools: Map<string, string>;
  autoApproveEnabled: boolean;
  toolDecisions: Map<string, ToolDecisionData>;
  isToolPending: (toolId: string) => boolean;
  onViewTool: (toolId: string, isAutoView?: boolean) => void;
  onApproveTool: (toolId: string) => void;
  onRejectTool: (toolId: string) => void;
  onApproveAll: (toolId: string) => void;
  rejectingToolId: string | null;
  rejectMessage: string;
  onRejectMessageChange: Dispatch<SetStateAction<string>>;
  onRejectSubmit: () => void;
  onRejectCancel: () => void;
  extractKeyInfo: (
    toolName: string,
    content: string,
    blockId: string,
  ) => string;
  onRangeClick: (range: string, workbookName?: string) => void;
  onOpenSettings: () => void;
}

const TOOL_ICON_MAP: Record<string, JSX.Element> = {
  read_values_batch: (
    <BookOpen className="w-3 h-3 text-blue-500 dark:text-blue-400" />
  ),
  read_format_batch: (
    <Pipette className="w-3 h-3 text-purple-500 dark:text-purple-400" />
  ),
  write_values_batch: (
    <PenTool className="w-3 h-3 text-green-500 dark:text-green-400" />
  ),
  drag_formula: (
    <PenTool className="w-3 h-3 text-green-500 dark:text-green-400" />
  ),
  write_format_batch: (
    <Palette className="w-3 h-3 text-pink-500 dark:text-pink-400" />
  ),
  add_columns: (
    <Columns className="w-3 h-3 text-emerald-500 dark:text-emerald-400" />
  ),
  remove_columns: <Trash2 className="w-3 h-3 text-red-500 dark:text-red-400" />,
  add_rows: <Rows className="w-3 h-3 text-emerald-500 dark:text-emerald-400" />,
  remove_rows: <Trash2 className="w-3 h-3 text-red-500 dark:text-red-400" />,
  add_sheets: <FileText className="w-3 h-3 text-cyan-500 dark:text-cyan-400" />,
  [ToolNames.MERGE_FILES]: (
    <Folder className="w-3 h-3 text-amber-600 dark:text-amber-400" />
  ),
};

const getToolIcon = (toolName?: string) => {
  if (!toolName) {
    return <Wrench className="w-3 h-3 text-slate-500 dark:text-slate-400" />;
  }
  return (
    TOOL_ICON_MAP[toolName] ?? (
      <Wrench className="w-3 h-3 text-slate-500 dark:text-slate-400" />
    )
  );
};

const COLUMN_WIDTH_MODE_KEY = "column_width_mode";

/** Render a single tool block with status and controls. */
export function ToolBlock({
  block,
  message,
  isExpanded,
  onToggle,
  approvedTools,
  rejectedTools,
  viewedTools,
  erroredTools,
  autoApproveEnabled,
  toolDecisions,
  isToolPending,
  onViewTool,
  onApproveTool,
  onRejectTool,
  onApproveAll,
  rejectingToolId,
  rejectMessage,
  onRejectMessageChange,
  onRejectSubmit,
  onRejectCancel,
  extractKeyInfo,
  onRangeClick,
  onOpenSettings,
}: ToolBlockProps) {
  let parameters: Record<string, unknown> = {};
  let hasValidParams = false;
  const toolDecision = block.toolUseId
    ? toolDecisions.get(block.toolUseId)
    : undefined;
  const rejectionReason =
    toolDecision?.decision === "rejected" ? toolDecision.message : undefined;

  const errorMessage = block.toolUseId
    ? erroredTools.get(block.toolUseId)
    : undefined;
  const isErrored = !!(block.toolUseId && erroredTools.has(block.toolUseId));
  const [columnWidthMode, setColumnWidthMode] = useState<ColumnWidthMode>(
    ColumnWidthModes.SMART,
  );

  if (block.content.trim()) {
    try {
      parameters = JSON.parse(block.content);
      hasValidParams = true;
    } catch {
      try {
        const partialJson = block.content.includes("{")
          ? `${block.content}}`
          : `{${block.content}}`;
        parameters = JSON.parse(partialJson);
        hasValidParams = true;
      } catch {
        hasValidParams = false;
      }
    }
  }

  const getToolDisplayInfo = () => {
    const toolName = block.toolName || "unknown";
    const keyInfo = extractKeyInfo(toolName, block.content, block.id);

    const toolId = block.toolUseId || "";
    const isApproved = approvedTools.has(toolId);
    const isRejected = rejectedTools.has(toolId);
    const isDecided = isApproved || isRejected;

    if (toolName === ToolNames.READ_VALUES_BATCH) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Read"
            : "Read (rejected)"
          : "Reading";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.READ_FORMAT_BATCH) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Read format"
            : "Read format (rejected)"
          : "Reading format";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.WRITE_VALUES_BATCH) {
      const action = formatWriteValuesAction({
        isComplete: block.isComplete,
        isApproved,
        isRejected,
        isErrored,
        errorMessage,
      });
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.WRITE_FORMAT_BATCH) {
      const wroteLabel = "Wrote format";
      const writingLabel = "Writing format";
      const rejectedLabel = "Write format (rejected)";
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? wroteLabel
            : rejectedLabel
          : writingLabel;
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.DRAG_FORMULA) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Dragged"
            : "Drag (rejected)"
          : "Dragging";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.ADD_COLUMNS) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Added columns"
            : "Add columns (rejected)"
          : "Adding columns";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.REMOVE_COLUMNS) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Removed columns"
            : "Remove columns (rejected)"
          : "Removing columns";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.ADD_ROWS) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Added rows"
            : "Add rows (rejected)"
          : "Adding rows";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.REMOVE_ROWS) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Removed rows"
            : "Remove rows (rejected)"
          : "Removing rows";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.ADD_SHEETS) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Added sheet"
            : "Add sheet (rejected)"
          : "Adding sheet";
      return { action, range: keyInfo };
    }
    if (toolName === ToolNames.MERGE_FILES) {
      const action =
        block.isComplete && isDecided
          ? isApproved
            ? "Merged files"
            : "Merge files (rejected)"
          : "Merging files";
      return { action, range: keyInfo };
    }
    const action =
      block.isComplete && isDecided
        ? isApproved
          ? toolName
          : `${toolName} (rejected)`
        : `${toolName}...`;
    return { action, range: "" };
  };

  const { action, range } = getToolDisplayInfo();

  const renderToolDisplay = () => {
    const renderReadBatchRanges = (): ReactNode | null => {
      const toolName = block.toolName;
      if (
        toolName !== ToolNames.READ_VALUES_BATCH &&
        toolName !== ToolNames.READ_FORMAT_BATCH
      ) {
        return null;
      }

      const operations = Array.isArray((parameters as any)?.operations)
        ? ((parameters as any).operations as Array<Record<string, unknown>>)
        : [];
      if (operations.length === 0) return null;

      const fallbackWorkbook =
        typeof (parameters as any)?.workbookName === "string"
          ? (parameters as any).workbookName
          : undefined;

      const pills = operations
        .map((op: Record<string, unknown>, index: number) => {
          const worksheet =
            typeof op?.worksheet === "string" ? op.worksheet : "";
          const rangeValue = typeof op?.range === "string" ? op.range : "";
          const label =
            worksheet && rangeValue
              ? `${worksheet}!${rangeValue}`
              : rangeValue || worksheet;
          if (!label) return null;
          const workbookName =
            typeof op?.workbookName === "string"
              ? op.workbookName
              : fallbackWorkbook;

          return (
            <span key={`op-${index}`} className="inline-flex items-center">
              {index > 0 ? (
                <span className="mx-1 text-slate-400">,</span>
              ) : null}
              <ExcelRangePill
                range={label}
                workbookName={workbookName}
                onClick={(clickedRange, wbName) =>
                  onRangeClick(clickedRange, wbName)
                }
                className="ml-0.5"
              />
            </span>
          );
        })
        .filter((pill: JSX.Element | null): pill is JSX.Element =>
          Boolean(pill),
        );

      if (pills.length === 0) return null;

      return (
        <>
          <span>{action} </span>
          {pills}
        </>
      );
    };

    const readBatchDisplay = renderReadBatchRanges();
    if (readBatchDisplay) {
      return readBatchDisplay;
    }

    if (!range) {
      return action;
    }

    const normalizedRange = range.trim();
    const getWorkbookNameForTool = () => {
      const workbookName = parameters?.workbookName;

      if (workbookName) return workbookName;

      const name = block.toolName;

      if (
        name === ToolNames.WRITE_FORMAT_BATCH &&
        Array.isArray(parameters?.operations)
      ) {
        const firstOp = parameters.operations[0];
        if (firstOp?.workbookName) return firstOp.workbookName;
      }

      return undefined;
    };

    const workbookName = getWorkbookNameForTool();
    const stripWorkbookPrefix = () => {
      if (!workbookName) return null;
      const lowerPrefix = workbookName.toLowerCase() + "!";
      const lowerRange = normalizedRange.toLowerCase();
      if (lowerRange.startsWith(lowerPrefix)) {
        return normalizedRange.slice(workbookName.length + 1);
      }
      return null;
    };

    const maybeRangeWithoutWorkbook = stripWorkbookPrefix();
    if (maybeRangeWithoutWorkbook) {
      return (
        <>
          <span>{action} </span>
          <ExcelRangePill
            range={maybeRangeWithoutWorkbook}
            workbookName={workbookName}
            onClick={(clickedRange, wbName) =>
              onRangeClick(clickedRange, wbName)
            }
            className="ml-0.5"
          />
        </>
      );
    }

    const matches = range.match(rangePattern);
    if (!matches || matches.length === 0) {
      return `${action}${range}`;
    }

    const parts: ReactNode[] = [];
    let processedText = range;
    let partKey = 0;

    parts.push(<span key="action">{action}</span>);

    matches.forEach((match) => {
      const rangeIndex = processedText.indexOf(match);

      if (rangeIndex > 0) {
        parts.push(
          <span key={`text-${partKey}`}>
            {processedText.substring(0, rangeIndex)}
          </span>,
        );
      }

      parts.push(
        <ExcelRangePill
          key={`range-${partKey++}`}
          range={match}
          workbookName={workbookName}
          onClick={(clickedRange, wbName) => onRangeClick(clickedRange, wbName)}
          className="ml-0.5"
        />,
      );

      processedText = processedText.substring(rangeIndex + match.length);
    });

    if (processedText.length > 0) {
      parts.push(<span key="text-end">{processedText}</span>);
    }

    return <>{parts}</>;
  };

  const needsApproval =
    block.isComplete &&
    block.toolUseId &&
    isToolPending(block.toolUseId) &&
    !autoApproveEnabled &&
    !erroredTools.has(block.toolUseId) &&
    !approvedTools.has(block.toolUseId) &&
    !rejectedTools.has(block.toolUseId);

  const pendingToolsInMessage =
    message?.contentBlocks?.filter(
      (b) =>
        b.type === ClaudeContentTypes.TOOL_USE &&
        b.toolUseId &&
        isToolPending(b.toolUseId),
    ).length || 0;

  useEffect(() => {
    if (block.toolName !== ToolNames.WRITE_VALUES_BATCH) return;
    let isActive = true;

    const fetchPreference = async () => {
      try {
        const base = await getApiBase();
        const response = await fetch(
          `${base}/api/user-preference?key=${COLUMN_WIDTH_MODE_KEY}&type=string&default=${ColumnWidthModes.SMART}`,
        );
        if (!response.ok) return;
        const data = await response.json();
        const raw = String(data.preferenceValue || "").toLowerCase();
        const next =
          raw === ColumnWidthModes.ALWAYS ||
          raw === ColumnWidthModes.SMART ||
          raw === ColumnWidthModes.NEVER
            ? raw
            : ColumnWidthModes.SMART;
        if (isActive) {
          setColumnWidthMode(next);
        }
      } catch (error) {
        console.warn("Failed to fetch column width preference:", error);
      }
    };

    fetchPreference();

    return () => {
      isActive = false;
    };
  }, [block.toolName]);

  const columnWidthNote =
    block.toolName === ToolNames.WRITE_VALUES_BATCH ? (
      <div className="text-[11px] text-slate-500 dark:text-slate-400">
        <span>{getColumnWidthMessage(columnWidthMode)} (</span>
        <button
          type="button"
          onClick={onOpenSettings}
          className="text-[11px] text-blue-600 dark:text-blue-400 hover:underline"
        >
          Open settings
        </button>
        <span>)</span>
      </div>
    ) : null;

  // Check if this is a completed merge_files tool with a result to display
  const isApprovedMergeFiles =
    block.toolName === ToolNames.MERGE_FILES &&
    block.isComplete &&
    block.toolUseId &&
    approvedTools.has(block.toolUseId);

  const parseMergeFilesResult = () => {
    console.log("[ToolBlock] parseMergeFilesResult called", {
      toolName: block.toolName,
      isComplete: block.isComplete,
      toolUseId: block.toolUseId,
      isApproved: block.toolUseId ? approvedTools.has(block.toolUseId) : false,
      isApprovedMergeFiles,
    });

    if (!isApprovedMergeFiles) {
      console.log("[ToolBlock] Not an approved merge_files tool, skipping");
      return null;
    }

    // Look for the next text block in the message that contains the tool result
    const blockIndex =
      message?.contentBlocks?.findIndex((b) => b.id === block.id) ?? -1;
    console.log(
      "[ToolBlock] Block index:",
      blockIndex,
      "Total blocks:",
      message?.contentBlocks?.length,
    );

    if (blockIndex === -1 || !message?.contentBlocks) {
      console.log(
        "[ToolBlock] Block not found in message or no content blocks",
      );
      return null;
    }

    // Check the next few blocks for a text block with JSON result
    for (
      let i = blockIndex + 1;
      i < Math.min(blockIndex + 5, message.contentBlocks.length);
      i++
    ) {
      const nextBlock = message.contentBlocks[i];
      console.log(`[ToolBlock] Checking block ${i}:`, {
        type: nextBlock.type,
        hasOutputPath: nextBlock.content.includes("outputPath"),
        contentPreview: nextBlock.content.substring(0, 100),
      });

      if (
        nextBlock.type === "text" &&
        nextBlock.content.includes("outputPath")
      ) {
        try {
          // Try to extract JSON from the content
          const jsonMatch = nextBlock.content.match(
            /\{[\s\S]*?"outputPath"[\s\S]*?\}/,
          );
          console.log(
            "[ToolBlock] JSON match:",
            jsonMatch ? "Found" : "Not found",
          );

          if (jsonMatch) {
            console.log("[ToolBlock] Matched JSON:", jsonMatch[0]);
            const result = JSON.parse(jsonMatch[0]);
            console.log("[ToolBlock] Parsed result:", result);

            if (result.outputPath) {
              console.log(
                "[ToolBlock] ✅ Found valid merge_files result:",
                result,
              );
              return result;
            }
          }
        } catch (error) {
          console.error("[ToolBlock] Error parsing JSON:", error);
        }
      }
    }

    console.log("[ToolBlock] No merge_files result found in subsequent blocks");
    return null;
  };

  const mergeFilesResult = parseMergeFilesResult();

  const handleOpenMergedFile = async (filePath: string) => {
    console.log("[ToolBlock] Opening merged file:", filePath);
    try {
      const success = await webViewBridge.openWorkbook(filePath);
      console.log("[ToolBlock] Open workbook result:", success);
      if (!success) {
        console.error("[ToolBlock] Failed to open merged workbook:", filePath);
      }
    } catch (error) {
      console.error("[ToolBlock] Error opening merged workbook:", error);
    }
  };

  return (
    <div className="group">
      <Collapsible
        open={isExpanded}
        onOpenChange={() => hasValidParams && onToggle(block.id)}
      >
        {needsApproval ? (
          <div className="px-1 py-2 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors relative">
            <div className="absolute inset-0 pending-approval-pulse" />

            <div className="relative z-10 space-y-3">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-400 flex-1">
                  <IconContainer>{getToolIcon(block.toolName)}</IconContainer>
                  <span className="font-medium">{renderToolDisplay()}</span>
                  {pendingToolsInMessage > 1 && (
                    <span className="text-xs text-orange-600 dark:text-orange-400 font-medium">
                      ({pendingToolsInMessage} tools)
                    </span>
                  )}
                </div>

                {hasValidParams && (
                  <CollapsibleTrigger asChild>
                    <button className="p-1 hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors">
                      {isExpanded ? (
                        <ChevronDown className="w-3 h-3 text-slate-400" />
                      ) : (
                        <ChevronRight className="w-3 h-3 text-slate-400" />
                      )}
                    </button>
                  </CollapsibleTrigger>
                )}
              </div>

              <ActionButtons
                isViewingTool={
                  !!block.toolUseId && viewedTools.has(block.toolUseId)
                }
                onViewTool={() =>
                  block.toolUseId && onViewTool(block.toolUseId)
                }
                onApproveTool={() =>
                  block.toolUseId && onApproveTool(block.toolUseId)
                }
                onRejectTool={() =>
                  block.toolUseId && onRejectTool(block.toolUseId)
                }
                onApproveAll={() =>
                  block.toolUseId && onApproveAll(block.toolUseId)
                }
              />
              {columnWidthNote}
            </div>
          </div>
        ) : (
          <>
            <div className="flex items-center justify-between px-1 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors relative py-0.5">
              <div className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-400 flex-1 relative z-10">
                <IconContainer>{getToolIcon(block.toolName)}</IconContainer>
                <span className="font-medium">{renderToolDisplay()}</span>
                {!block.isComplete && (
                  <WaveAnimation className="text-slate-400 dark:text-slate-500" />
                )}
              </div>

              <div className="flex items-center gap-1 relative z-10">
                {block.toolUseId &&
                  approvedTools.has(block.toolUseId) &&
                  !erroredTools.has(block.toolUseId) && (
                    <span className="text-xs text-green-600 dark:text-green-400">
                      ✓
                    </span>
                  )}
                {block.toolUseId && rejectedTools.has(block.toolUseId) && (
                  <span className="text-xs text-red-600 dark:text-red-400">
                    ✗
                  </span>
                )}
                {block.toolUseId && erroredTools.has(block.toolUseId) && (
                  <span
                    className="text-xs text-red-600 dark:text-red-400"
                    title={erroredTools.get(block.toolUseId)}
                  >
                    !
                  </span>
                )}

                {hasValidParams &&
                  block.toolUseId &&
                  (approvedTools.has(block.toolUseId) ||
                    rejectedTools.has(block.toolUseId)) && (
                    <CollapsibleTrigger asChild>
                      <button className="opacity-0 group-hover:opacity-100 transition-opacity p-1 hover:bg-slate-200 dark:hover:bg-slate-700 rounded">
                        {isExpanded ? (
                          <ChevronDown className="w-3 h-3 text-slate-400" />
                        ) : (
                          <ChevronRight className="w-3 h-3 text-slate-400" />
                        )}
                      </button>
                    </CollapsibleTrigger>
                  )}
              </div>
            </div>
            {columnWidthNote && (
              <div className="px-1 pb-1">{columnWidthNote}</div>
            )}
          </>
        )}

        <CollapsibleContent>
          {hasValidParams && (
            <div className="ml-4 mt-2 p-2 bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700 text-xs">
              <div className="font-medium text-slate-700 dark:text-slate-300 mb-1">
                Parameters:
              </div>
              <pre className="text-slate-600 dark:text-slate-400 whitespace-pre-wrap overflow-x-auto">
                {JSON.stringify(parameters, null, 2)}
              </pre>
            </div>
          )}
        </CollapsibleContent>

        {block.toolUseId && rejectionReason && (
          <div className="ml-4 mt-2 text-[11px] text-red-600 dark:text-red-400">
            Rejection reason: {rejectionReason}
          </div>
        )}

        {mergeFilesResult && (
          <div className="ml-4 mt-2 p-3 bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-800 rounded">
            <div className="flex items-center gap-2 mb-2">
              <FileSpreadsheet className="w-4 h-4 text-green-600 dark:text-green-400" />
              <span className="text-xs font-semibold text-green-800 dark:text-green-200">
                Merged Workbook Created
              </span>
            </div>
            <div className="space-y-2">
              <div className="text-[11px] text-slate-700 dark:text-slate-300">
                <span className="font-medium">Files merged:</span>{" "}
                {mergeFilesResult.fileCount}
                {" • "}
                <span className="font-medium">Total sheets:</span>{" "}
                {mergeFilesResult.sheetCount}
              </div>
              <button
                onClick={() =>
                  handleOpenMergedFile(mergeFilesResult.outputPath)
                }
                className="w-full flex items-center justify-center gap-2 px-3 py-2 bg-green-600 hover:bg-green-700 text-white text-xs font-medium rounded transition-colors"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                Open{" "}
                {mergeFilesResult.outputPath.split("\\").pop() ||
                  mergeFilesResult.outputPath.split("/").pop() ||
                  "Merged File"}
              </button>
              <div
                className="text-[10px] text-slate-500 dark:text-slate-400 truncate"
                title={mergeFilesResult.outputPath}
              >
                {mergeFilesResult.outputPath}
              </div>
            </div>
          </div>
        )}

        {block.toolUseId &&
          rejectingToolId === block.toolUseId &&
          !erroredTools.has(block.toolUseId) && (
            <div className="ml-4 mt-2 p-3 bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700">
              <div className="text-xs text-slate-600 dark:text-slate-400 font-medium mb-2">
                Why are you rejecting this tool? (Optional)
              </div>
              <div className="flex gap-2">
                <Textarea
                  value={rejectMessage}
                  onChange={(e) => onRejectMessageChange(e.target.value)}
                  placeholder="Type your reason for rejecting this tool..."
                  className="flex-1 resize-none border border-slate-300 dark:border-slate-600 focus:border-slate-500 focus:ring-2 focus:ring-slate-500/20 min-h-[60px] px-2 py-1.5 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 placeholder:text-slate-500 dark:placeholder:text-slate-400 text-xs"
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && !e.shiftKey) {
                      e.preventDefault();
                      onRejectSubmit();
                    }
                    if (e.key === "Escape") {
                      onRejectCancel();
                    }
                  }}
                  autoFocus
                />
                <div className="flex flex-col gap-1">
                  <button
                    onClick={onRejectSubmit}
                    className="px-3 py-1.5 text-xs bg-red-500 hover:bg-red-600 text-white transition-colors"
                  >
                    Reject
                  </button>
                  <button
                    onClick={onRejectCancel}
                    className="px-3 py-1.5 text-xs text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            </div>
          )}
      </Collapsible>
    </div>
  );
}
