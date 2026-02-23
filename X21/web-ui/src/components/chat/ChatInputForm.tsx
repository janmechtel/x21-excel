import type {
  ChangeEventHandler,
  ClipboardEventHandler,
  DragEventHandler,
  FormEventHandler,
  KeyboardEventHandler,
  RefObject,
} from "react";

import {
  AlertCircle,
  AtSign,
  ChevronsRight,
  Command,
  Loader2,
  Plus,
  SlidersHorizontal,
  Square,
  X,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { RichChatEditor } from "@/components/chat/RichChatEditor";
import type { SlashCommandDefinition } from "@/types/slash-commands";

export type PromptSegment = {
  type: "text" | "range" | "sheetMention";
  value: string;
};

interface ChatInputFormProps {
  prompt: string;
  promptPlaceholder: string;
  promptSegments: PromptSegment[];
  activeSlashCommand: SlashCommandDefinition | null;
  disablePrompt: boolean;
  disableSend: boolean;
  isStreaming: boolean;
  isConvertingFile: boolean;
  isDragOver: boolean;
  autoApproveEnabled: boolean;
  activeToolsCount: number;
  textareaRef: RefObject<HTMLTextAreaElement>;
  fileInputRef: RefObject<HTMLInputElement>;
  openFileDialog: () => void;
  openSheetMentionPaletteHandler: (cursorPosition?: number) => void;
  openSlashCommandPaletteHandler: () => void;
  clearActiveSlashCommand: () => void;
  setAutoApproveEnabled: (val: boolean) => void;
  onToggleTools: () => void;
  handlePromptChange: ChangeEventHandler<HTMLTextAreaElement>;
  handlePromptKeyDown: KeyboardEventHandler<HTMLTextAreaElement>;
  handlePaste: ClipboardEventHandler<HTMLTextAreaElement>;
  handleDragOver: DragEventHandler;
  handleDragEnter: DragEventHandler;
  handleDragLeave: DragEventHandler;
  handleDrop: DragEventHandler;
  handleSend: () => void;
  handleCancel: () => void;
  handleFileInputChange: ChangeEventHandler<HTMLInputElement>;
  handleSubmit: FormEventHandler;
  onRangePillClick: (range: string) => void;
  onSheetMentionClick: (sheetName: string) => void;
}

export function ChatInputForm({
  prompt,
  promptPlaceholder,
  promptSegments,
  activeSlashCommand,
  disablePrompt,
  disableSend,
  isStreaming,
  isConvertingFile,
  isDragOver,
  autoApproveEnabled,
  activeToolsCount,
  textareaRef,
  fileInputRef,
  openFileDialog,
  openSheetMentionPaletteHandler,
  openSlashCommandPaletteHandler,
  clearActiveSlashCommand,
  setAutoApproveEnabled,
  onToggleTools,
  handlePromptChange,
  handlePromptKeyDown,
  handlePaste,
  handleDragOver,
  handleDragEnter,
  handleDragLeave,
  handleDrop,
  handleSend,
  handleCancel,
  handleFileInputChange,
  handleSubmit,
  onRangePillClick,
  onSheetMentionClick,
}: ChatInputFormProps) {
  return (
    <form
      onSubmit={handleSubmit}
      className="flex gap-4 items-end w-full h-full flex-1 min-h-0"
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      <div
        className={`relative flex-1 border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 shadow-sm transition-all duration-200 focus-within:border-blue-500 focus-within:ring-2 focus-within:ring-blue-500/20 flex flex-col h-full ${
          isDragOver
            ? "border-blue-400 dark:border-blue-500 bg-blue-50/50 dark:bg-blue-950/30 ring-2 ring-blue-500/20"
            : ""
        }`}
        title={
          disablePrompt
            ? "Please cancel or wait for the current request to finish"
            : undefined
        }
        onDragOver={handleDragOver}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {activeSlashCommand && (
          <div className="flex items-center justify-between px-4 pt-2 pb-1 text-[11px] text-blue-700 dark:text-blue-300 border-b border-blue-100/70 dark:border-blue-900/40">
            <div className="flex items-center gap-2 font-semibold uppercase tracking-wide">
              <Command className="w-3 h-3" />
              <span>/{activeSlashCommand.title}</span>
            </div>
            <button
              type="button"
              onClick={clearActiveSlashCommand}
              className="text-slate-400 hover:text-slate-600 dark:text-slate-500 dark:hover:text-slate-300 transition-colors"
            >
              <X className="w-3 h-3" />
              <span className="sr-only">Clear slash command</span>
            </button>
          </div>
        )}
        <div
          className="flex-1 overflow-hidden flex flex-col relative"
          data-has-text={prompt.length > 0 ? "true" : "false"}
        >
          <RichChatEditor
            promptSegments={promptSegments}
            placeholder={promptPlaceholder}
            disabled={disablePrompt}
            hasActiveSlashCommand={Boolean(activeSlashCommand)}
            textareaRef={textareaRef}
            onRangePillClick={onRangePillClick}
            onSheetMentionClick={onSheetMentionClick}
            onPromptChange={handlePromptChange}
            onPromptKeyDown={handlePromptKeyDown}
            onPaste={handlePaste}
            onDragOver={handleDragOver}
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          />
        </div>

        {/* Dividing line */}
        <div className="border-t border-slate-300 dark:border-slate-600" />

        {/* Bottom controls */}
        <div className="flex items-center justify-between px-2 py-2 min-h-[40px]">
          {/* Left side: Controls */}
          <div className="flex items-center gap-2">
            {/* Attach button */}
            <input
              ref={fileInputRef}
              type="file"
              accept=".pdf,.gif,.png,.jpg,.jpeg,.webm"
              multiple
              onChange={handleFileInputChange}
              className="hidden"
            />
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  type="button"
                  onClick={openFileDialog}
                  variant="ghost"
                  className="w-6 h-6 rounded-md p-0 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                  disabled={isStreaming || isConvertingFile}
                >
                  {isConvertingFile ? (
                    <Loader2 className="w-3 h-3 animate-spin text-slate-500" />
                  ) : (
                    <Plus className="w-3 h-3 text-slate-600 dark:text-slate-400" />
                  )}
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>Attach files</p>
              </TooltipContent>
            </Tooltip>

            {/* Sheet mention button */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  type="button"
                  onClick={() =>
                    openSheetMentionPaletteHandler(
                      textareaRef?.current?.selectionStart ?? undefined,
                    )
                  }
                  variant="ghost"
                  className="w-6 h-6 rounded-md p-0 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center justify-center"
                  disabled={disablePrompt}
                >
                  <AtSign className="w-3 h-3 text-slate-600 dark:text-slate-400" />
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>Sheets (@)</p>
              </TooltipContent>
            </Tooltip>

            {/* Commands button */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  type="button"
                  onClick={openSlashCommandPaletteHandler}
                  variant="ghost"
                  className="w-6 h-6 rounded-md p-0 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors flex items-center justify-center"
                >
                  <span className="text-sm font-semibold text-slate-600 dark:text-slate-400">
                    /
                  </span>
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>Commands (/)</p>
              </TooltipContent>
            </Tooltip>

            {/* Tools button */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  type="button"
                  onClick={onToggleTools}
                  variant="ghost"
                  className="w-6 h-6 rounded-md p-0 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors relative"
                >
                  <SlidersHorizontal className="w-3 h-3 text-slate-600 dark:text-slate-400" />
                  {activeToolsCount > 0 && (
                    <span className="absolute -top-0.5 -right-0.5 w-2.5 h-2.5 bg-slate-200 dark:bg-slate-600 text-slate-700 dark:text-slate-300 text-[8px] rounded-full flex items-center justify-center font-medium">
                      {activeToolsCount}
                    </span>
                  )}
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>{activeToolsCount} tools enabled (Ctrl+Shift+T)</p>
              </TooltipContent>
            </Tooltip>

            {/* Ask before edits toggle */}
            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  type="button"
                  onClick={() =>
                    !isStreaming && setAutoApproveEnabled(!autoApproveEnabled)
                  }
                  className={`flex items-center gap-1 h-6 px-2 rounded-md text-[10px] font-medium transition-colors text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 ${
                    isStreaming
                      ? "cursor-not-allowed opacity-50"
                      : "cursor-pointer"
                  }`}
                  disabled={isStreaming}
                >
                  {autoApproveEnabled ? (
                    <>
                      <ChevronsRight className="w-3 h-3" />
                      <span>Edit automatically</span>
                    </>
                  ) : (
                    <>
                      <AlertCircle className="w-3 h-3" />
                      <span>Ask before edits</span>
                    </>
                  )}
                </button>
              </TooltipTrigger>
              <TooltipContent>
                <p>
                  {isStreaming
                    ? "Cannot change mode while request is running"
                    : autoApproveEnabled
                    ? "Click to ask before edits"
                    : "Click to edit automatically"}
                </p>
              </TooltipContent>
            </Tooltip>
          </div>

          {/* Right side: Send button */}
          <div className="flex items-center gap-1.5">
            <Button
              type="button"
              onClick={() => {
                if (isStreaming) {
                  handleCancel();
                } else {
                  handleSend();
                }
              }}
              variant={isStreaming ? "destructive" : "default"}
              className={`w-6 h-6 font-medium transition-all duration-200 ${
                isStreaming
                  ? "bg-gradient-to-r from-red-500 to-red-600 hover:from-red-600 hover:to-red-700 shadow-lg hover:shadow-xl"
                  : "bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 shadow-lg hover:shadow-xl"
              } ${
                disableSend
                  ? "opacity-50 cursor-not-allowed"
                  : "hover:scale-[1.02] active:scale-[0.98]"
              }`}
              disabled={disableSend}
            >
              {isStreaming ? (
                <Square className="w-2 h-2" />
              ) : (
                <svg
                  className="w-2 h-2"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth={2}
                    d="M3 12h18m-9-9l9 9-9 9"
                  />
                </svg>
              )}
            </Button>
          </div>
        </div>
      </div>
    </form>
  );
}
