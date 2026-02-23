import type { Dispatch, SetStateAction } from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { Loader2, Paperclip, X } from "lucide-react";

import { SelectedRangeBadge } from "@/components/chat/SelectedRangeBadge";
import { StatusMetaRow } from "@/components/chat/StatusMeta";
import {
  ChatInputForm,
  type PromptSegment,
} from "@/components/chat/ChatInputForm";
import { SlashCommandPalette } from "@/components/slash-commands/SlashCommandPalette";
import { Button } from "@/components/ui/button";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { SheetMentionPalette } from "@/components/excel/SheetMentionPalette";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import type { AttachedFile, OperationStatus } from "@/types/chat";
import { rangePattern } from "@/utils/toolDisplay";
import { OperationStatusValues } from "@/types/chat";
import {
  buildSheetLookup,
  buildWorkbookNameSet,
  isKnownSheetName,
  isKnownWorkbook,
  normalizeSheetName,
} from "@/utils/excelSheetValidation";

interface ChatInputAreaProps {
  prompt: string;
  promptPlaceholder: string;
  activeSlashCommand: SlashCommandDefinition | null;
  isSlashCommandActive: boolean;
  slashCommandMatches: SlashCommandDefinition[];
  slashCommandHighlightIndex: number;
  slashCommandQuery: string;
  setSlashCommandHighlightIndex: Dispatch<SetStateAction<number>>;
  isSheetMentionActive: boolean;
  sheetMentionMatches: string[];
  sheetMentionHighlightIndex: number;
  sheetMentionQuery: string;
  setSheetMentionHighlightIndex: Dispatch<SetStateAction<number>>;
  handleSheetMentionSelect: (sheetName: string) => void;
  currentSheetNames?: string[];
  currentWorkbookName?: string;
  openWorkbooks?: Array<{ workbookName: string; workbookFullName?: string }>;
  otherWorkbookSheets: Array<{
    workbookName: string;
    sheetName: string;
    workbookFullName?: string;
  }>;
  isRequestingOtherWorkbooks: boolean;
  resetSheetMentionState: () => void;
  openSheetMentionPaletteHandler: (cursorPosition?: number) => void;
  clearActiveSlashCommand: () => void;
  resetSlashCommandState: () => void;
  openSlashCommandPaletteHandler: () => void;
  handleSlashCommandSelect: (command: SlashCommandDefinition) => void;
  handlePromptChange: React.ChangeEventHandler<HTMLTextAreaElement>;
  handlePromptKeyDown: React.KeyboardEventHandler<HTMLTextAreaElement>;
  handleSend: () => void;
  handleCancel: () => void;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
  isStreaming: boolean;
  loadingState: OperationStatus;
  getPendingToolCount: () => number;
  attachedFiles: AttachedFile[];
  isDragOver: boolean;
  isConvertingFile: boolean;
  clearAllAttachedFiles: () => void;
  removeAttachedFile: (index: number) => void;
  formatFileSize: (size: number) => string;
  handleDragOver: React.DragEventHandler;
  handleDragEnter: React.DragEventHandler;
  handleDragLeave: React.DragEventHandler;
  handleDrop: React.DragEventHandler;
  handlePaste: React.ClipboardEventHandler<HTMLTextAreaElement>;
  fileInputRef: React.RefObject<HTMLInputElement>;
  handleFileInputChange: React.ChangeEventHandler<HTMLInputElement>;
  openFileDialog: () => void;
  autoApproveEnabled: boolean;
  setAutoApproveEnabled: (val: boolean) => void;
  onToggleTools: () => void;
  activeToolsCount: number;
  selectedRange: string;
  onInsertSelectedRange: () => void;
  handleRangeNavigate: (range: string, workbookName?: string) => void;
  handleSheetNavigate: (sheetName: string) => void;
  // Status meta
  totalTokens: number;
  inputTokens: number;
  outputTokens: number;
  wsConnected: boolean;
  wsUrl: string;
  hasMessages: boolean;
}

export function ChatInputArea({
  prompt,
  promptPlaceholder,
  activeSlashCommand,
  isSlashCommandActive,
  slashCommandMatches,
  slashCommandHighlightIndex,
  slashCommandQuery,
  setSlashCommandHighlightIndex,
  isSheetMentionActive,
  sheetMentionMatches,
  sheetMentionHighlightIndex,
  sheetMentionQuery,
  setSheetMentionHighlightIndex,
  handleSheetMentionSelect,
  currentSheetNames = [],
  currentWorkbookName,
  openWorkbooks = [],
  otherWorkbookSheets,
  isRequestingOtherWorkbooks,
  resetSheetMentionState,
  openSheetMentionPaletteHandler,
  clearActiveSlashCommand,
  resetSlashCommandState,
  openSlashCommandPaletteHandler,
  handleSlashCommandSelect,
  handlePromptChange,
  handlePromptKeyDown,
  handleSend,
  handleCancel,
  textareaRef,
  isStreaming,
  loadingState,
  getPendingToolCount,
  attachedFiles,
  isDragOver,
  isConvertingFile,
  removeAttachedFile,
  formatFileSize,
  handleDragOver,
  handleDragEnter,
  handleDragLeave,
  handleDrop,
  handlePaste,
  fileInputRef,
  handleFileInputChange,
  openFileDialog,
  autoApproveEnabled,
  setAutoApproveEnabled,
  onToggleTools,
  activeToolsCount,
  selectedRange,
  onInsertSelectedRange,
  handleRangeNavigate,
  handleSheetNavigate,
  totalTokens,
  inputTokens,
  outputTokens,
  wsConnected,
  wsUrl,
  hasMessages,
}: ChatInputAreaProps) {
  const pendingCount = getPendingToolCount();
  const disablePrompt =
    isStreaming ||
    pendingCount > 0 ||
    loadingState !== OperationStatusValues.IDLE;
  const disableSend =
    !prompt.trim() &&
    attachedFiles.length === 0 &&
    !activeSlashCommand &&
    !isStreaming;
  const formatRangeForDisplay = useCallback((raw: string) => {
    if (raw.includes("[") && raw.includes("]")) {
      return raw;
    }
    const parts = raw.split("!");
    if (parts.length >= 3) {
      return raw;
    }
    const [sheetPart, rangePart] = raw.split("!");
    if (!rangePart) return raw;

    const sheetName = normalizeSheetName(sheetPart ?? "");
    const needsQuoting = sheetName.length > 0 && /\s/.test(sheetName);
    const safeSheet = sheetName
      ? needsQuoting
        ? `'${sheetName.replace(/'/g, "''")}'`
        : sheetName
      : "";

    const cleanedRange = rangePart.trim();

    return safeSheet ? `${safeSheet}!${cleanedRange}` : cleanedRange;
  }, []);

  const sheetMentionPattern =
    /@?(?:\[[^\]]+\])?(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)!(?=$|\s|[.,;:)\]])/;
  const workbookRangePattern =
    /(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)!(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)!\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?/;
  const bracketWorkbookRangePattern =
    /\[[^\]]+\][^!]+!\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?/;

  const knownWorkbooks = useMemo(
    () =>
      buildWorkbookNameSet(
        openWorkbooks,
        currentWorkbookName,
        openWorkbooks.length === 0 ? otherWorkbookSheets : undefined,
      ),
    [openWorkbooks, currentWorkbookName, otherWorkbookSheets],
  );
  const filteredOtherWorkbookSheets = useMemo(() => {
    if (openWorkbooks.length === 0) return otherWorkbookSheets;
    return otherWorkbookSheets.filter((entry) => {
      const workbookToken = entry.workbookName || entry.workbookFullName || "";
      return isKnownWorkbook(workbookToken, knownWorkbooks);
    });
  }, [openWorkbooks.length, otherWorkbookSheets, knownWorkbooks]);
  const sheetLookup = useMemo(
    () =>
      buildSheetLookup(
        currentSheetNames,
        undefined,
        currentWorkbookName,
        filteredOtherWorkbookSheets,
      ),
    [currentSheetNames, currentWorkbookName, filteredOtherWorkbookSheets],
  );

  const isValidSheetToken = useCallback(
    (token: string) => {
      const bangIndex = token.indexOf("!");
      if (bangIndex < 0) return true;
      const parts = token.split("!");
      if (parts.length >= 3) {
        const workbookPart = parts[0];
        const sheetPart = parts[1];
        if (!isKnownWorkbook(workbookPart, knownWorkbooks)) return false;
        return isKnownSheetName(sheetPart, sheetLookup, workbookPart);
      }

      const sheetPart = parts[0];
      const bracketMatch = sheetPart.match(/^\[([^\]]+)\](.+)$/);
      if (bracketMatch) {
        const workbookPart = bracketMatch[1];
        const extractedSheet = bracketMatch[2];
        if (!isKnownWorkbook(workbookPart, knownWorkbooks)) return false;
        return isKnownSheetName(extractedSheet, sheetLookup, workbookPart);
      }
      return isKnownSheetName(sheetPart, sheetLookup);
    },
    [sheetLookup, knownWorkbooks],
  );

  const promptSegments = useMemo<PromptSegment[]>(() => {
    const segments: PromptSegment[] = [];
    const combinedRegexSource = `${workbookRangePattern.source}|${bracketWorkbookRangePattern.source}|${sheetMentionPattern.source}|${rangePattern.source}`;
    const combinedRegex = new RegExp(combinedRegexSource, "g");
    const strictSheetMentionRegex = new RegExp(
      `^${sheetMentionPattern.source}$`,
    );
    let lastIndex = 0;

    for (const match of prompt.matchAll(combinedRegex)) {
      const matchText = match[0];
      const start = match.index ?? 0;
      const end = start + matchText.length;

      if (!isValidSheetToken(matchText)) {
        continue;
      }

      // Skip matches that are part of a larger alphanumeric token (e.g., words like "X21" or "AI23")
      const prevChar = start > 0 ? prompt[start - 1] : "";
      const nextChar = end < prompt.length ? prompt[end] : "";
      const isAdjacentToWord =
        /[A-Za-z0-9_]/.test(prevChar) || /[A-Za-z0-9_]/.test(nextChar);
      if (isAdjacentToWord) {
        continue;
      }

      if (start > lastIndex) {
        segments.push({ type: "text", value: prompt.slice(lastIndex, start) });
      }

      if (strictSheetMentionRegex.test(matchText)) {
        segments.push({ type: "sheetMention", value: matchText });
      } else {
        segments.push({
          type: "range",
          value: formatRangeForDisplay(matchText),
        });
      }

      lastIndex = end;
    }

    if (lastIndex < prompt.length) {
      segments.push({ type: "text", value: prompt.slice(lastIndex) });
    }

    combinedRegex.lastIndex = 0;

    return segments.length ? segments : [{ type: "text", value: prompt }];
  }, [prompt, formatRangeForDisplay, isValidSheetToken]);

  const handleRangePillClick = (range: string) => {
    handleRangeNavigate(range);
  };

  const parseSheetMentionTarget = useCallback((raw: string) => {
    const cleaned = raw.replace(/^@/, "").replace(/!$/, "").trim();
    const bracketMatch = cleaned.match(/^\[([^\]]+)\](.+)$/);
    const workbookName = bracketMatch?.[1];
    const sheetToken = bracketMatch?.[2] ?? cleaned;
    const sheetName = sheetToken
      .replace(/^'+|'+$/g, "")
      .replace(/''/g, "'")
      .trim();
    return { sheetName, workbookName };
  }, []);

  const handleSheetMentionClick = (sheetName: string) => {
    const target = parseSheetMentionTarget(sheetName);
    if (!target.sheetName) return;
    if (target.workbookName) {
      const needsQuoting =
        /[\s,]/.test(target.sheetName) || target.sheetName.includes("'");
      const safeSheet = needsQuoting
        ? `'${target.sheetName.replace(/'/g, "''")}'`
        : target.sheetName;
      handleRangeNavigate(`${safeSheet}!A1`, target.workbookName);
      return;
    }
    handleSheetNavigate(target.sheetName);
  };

  // Resize state - default to minimum height
  // Minimum: selection row (~24px) + input box (40px) + commands/buttons (40px) + spacing/padding (~20px), doubled for more space
  // When slash command is active, add extra space for the command header (~35px)
  const BASE_MIN_HEIGHT = 160;
  const SLASH_COMMAND_HEADER_HEIGHT = 35;
  const MIN_HEIGHT = activeSlashCommand
    ? BASE_MIN_HEIGHT + SLASH_COMMAND_HEADER_HEIGHT
    : BASE_MIN_HEIGHT;
  const [inputAreaHeight, setInputAreaHeight] =
    useState<number>(BASE_MIN_HEIGHT);
  const [isResizing, setIsResizing] = useState(false);
  const resizeStartY = useRef<number>(0);
  const resizeStartHeight = useRef<number>(MIN_HEIGHT);
  const resizeHandleRef = useRef<HTMLDivElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  // Adjust input area height when slash command state changes
  useEffect(() => {
    // If current height is at or near the minimum, adjust it when slash command activates/deactivates
    const newMinHeight = activeSlashCommand
      ? BASE_MIN_HEIGHT + SLASH_COMMAND_HEADER_HEIGHT
      : BASE_MIN_HEIGHT;
    if (inputAreaHeight <= BASE_MIN_HEIGHT + SLASH_COMMAND_HEADER_HEIGHT) {
      setInputAreaHeight(newMinHeight);
    }
  }, [activeSlashCommand]);

  // Handle resize drag
  useEffect(() => {
    if (!isResizing) return;

    const handleMouseMove = (e: MouseEvent) => {
      // Calculate the delta (how much the mouse moved up - negative means up)
      const deltaY = resizeStartY.current - e.clientY;

      // New height = start height + delta (delta is positive when dragging up)
      const newHeight = resizeStartHeight.current + deltaY;

      // Use the current MIN_HEIGHT which adjusts based on slash command state
      const minHeight = MIN_HEIGHT;
      // Maximum height - allow up to 50% of viewport
      const maxHeight = window.innerHeight * 0.5;

      setInputAreaHeight(Math.max(minHeight, Math.min(maxHeight, newHeight)));
    };

    const handleMouseUp = () => {
      setIsResizing(false);
    };

    document.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseup", handleMouseUp);

    return () => {
      document.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
    };
  }, [isResizing]);

  const handleResizeStart = (e: React.MouseEvent) => {
    e.preventDefault();
    resizeStartY.current = e.clientY;
    resizeStartHeight.current = inputAreaHeight;
    setIsResizing(true);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
  };

  return (
    <div
      ref={containerRef}
      className="mb-1 relative flex flex-col"
      style={{ height: `${inputAreaHeight}px` }}
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* Resize handle */}
      <div
        ref={resizeHandleRef}
        onMouseDown={handleResizeStart}
        className={`h-1.5 cursor-ns-resize z-10 group transition-colors flex-shrink-0 relative ${
          isResizing
            ? "bg-blue-400 dark:bg-blue-500"
            : "bg-transparent hover:bg-slate-300 dark:hover:bg-slate-600"
        }`}
        style={{
          background: isResizing
            ? undefined
            : "linear-gradient(to bottom, transparent 0%, rgba(148, 163, 184, 0.3) 50%, transparent 100%)",
        }}
      >
        {/* Visual indicator dots */}
        <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 flex gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none">
          <div className="w-1 h-1 rounded-full bg-slate-400 dark:bg-slate-500"></div>
          <div className="w-1 h-1 rounded-full bg-slate-400 dark:bg-slate-500"></div>
          <div className="w-1 h-1 rounded-full bg-slate-400 dark:bg-slate-500"></div>
        </div>
      </div>

      <div className="flex-1 overflow-hidden flex flex-col min-h-0">
        {isConvertingFile && (
          <div className="mb-3 max-w-4xl mx-auto flex-shrink-0">
            <div className="flex items-center gap-2 p-2 bg-blue-50/50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-700">
              <Loader2 className="w-4 h-4 text-blue-500 animate-spin" />
              <span className="text-xs text-blue-600 dark:text-blue-400 font-medium">
                Converting PDF to base64...
              </span>
            </div>
          </div>
        )}

        {!isDragOver && attachedFiles.length > 0 && (
          <div
            data-testid="attachment-row"
            className="my-2 max-w-4xl flex-shrink-0 min-h-[24px] flex items-center justify-between gap-2"
          >
            <div
              data-testid="attached-files"
              className="flex items-center gap-0.5 flex-wrap justify-start"
            >
              {attachedFiles.map((file, index) => (
                <div
                  key={index}
                  className="flex items-center gap-0.5 bg-blue-100 dark:bg-blue-900/30 px-1.5 h-6 text-[10px] border border-blue-200 dark:border-blue-700"
                >
                  <Paperclip className="w-2 h-2 text-blue-600 dark:text-blue-400 flex-shrink-0" />
                  <span
                    className="text-blue-800 dark:text-blue-200 truncate max-w-[50px]"
                    title={`${file.name} (${formatFileSize(file.size)})`}
                  >
                    {file.name}
                  </span>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        onClick={() => removeAttachedFile(index)}
                        variant="ghost"
                        size="sm"
                        className="h-2.5 w-2.5 p-0 text-blue-600 dark:text-blue-400 hover:text-red-600 dark:hover:text-red-400"
                      >
                        <X className="w-2 h-2" />
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent>
                      <p>Remove file</p>
                    </TooltipContent>
                  </Tooltip>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="my-1 max-w-4xl flex-shrink-0 flex items-center justify-between gap-3">
          <StatusMetaRow
            hasMessages={hasMessages}
            totalTokens={totalTokens}
            inputTokens={inputTokens}
            outputTokens={outputTokens}
            wsConnected={wsConnected}
            wsUrl={wsUrl}
          />
          {selectedRange && (
            <SelectedRangeBadge
              selectedRange={selectedRange}
              onInsert={onInsertSelectedRange}
              disabled={disablePrompt}
            />
          )}
        </div>

        <div className="relative w-full flex-1 flex flex-col min-h-0">
          <SlashCommandPalette
            isOpen={isSlashCommandActive}
            commands={slashCommandMatches}
            highlightedIndex={slashCommandHighlightIndex}
            query={slashCommandQuery}
            onSelect={handleSlashCommandSelect}
            onHighlight={setSlashCommandHighlightIndex}
            onClose={resetSlashCommandState}
            textareaRef={textareaRef}
          />
          <SheetMentionPalette
            isOpen={isSheetMentionActive && !isSlashCommandActive}
            sheets={sheetMentionMatches}
            otherWorkbookSheets={otherWorkbookSheets}
            isRequestingOtherWorkbooks={isRequestingOtherWorkbooks}
            highlightedIndex={sheetMentionHighlightIndex}
            query={sheetMentionQuery}
            onSelect={handleSheetMentionSelect}
            onHighlight={setSheetMentionHighlightIndex}
            onClose={resetSheetMentionState}
            textareaRef={textareaRef}
          />

          <ChatInputForm
            prompt={prompt}
            promptPlaceholder={promptPlaceholder}
            promptSegments={promptSegments}
            activeSlashCommand={activeSlashCommand}
            disablePrompt={disablePrompt}
            disableSend={disableSend}
            isStreaming={isStreaming}
            isConvertingFile={isConvertingFile}
            isDragOver={isDragOver}
            autoApproveEnabled={autoApproveEnabled}
            activeToolsCount={activeToolsCount}
            textareaRef={textareaRef}
            fileInputRef={fileInputRef}
            openFileDialog={openFileDialog}
            openSheetMentionPaletteHandler={openSheetMentionPaletteHandler}
            openSlashCommandPaletteHandler={openSlashCommandPaletteHandler}
            clearActiveSlashCommand={clearActiveSlashCommand}
            setAutoApproveEnabled={setAutoApproveEnabled}
            onToggleTools={onToggleTools}
            handlePromptChange={handlePromptChange}
            handlePromptKeyDown={handlePromptKeyDown}
            handlePaste={handlePaste}
            handleDragOver={handleDragOver}
            handleDragEnter={handleDragEnter}
            handleDragLeave={handleDragLeave}
            handleDrop={handleDrop}
            handleSend={handleSend}
            handleCancel={handleCancel}
            handleFileInputChange={handleFileInputChange}
            handleSubmit={handleSubmit}
            onRangePillClick={handleRangePillClick}
            onSheetMentionClick={handleSheetMentionClick}
          />
        </div>
      </div>
    </div>
  );
}
