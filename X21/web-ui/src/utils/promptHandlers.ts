import type React from "react";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import { webViewBridge } from "@/services/webViewBridge";
import { getUserMessage } from "@/services/historyService";

interface PromptKeyDownConfig {
  isSlashCommandActive: boolean;
  slashCommandCount: number;
  setSlashCommandHighlightIndex: (updater: (prev: number) => number) => void;
  highlightedSlashCommand: SlashCommandDefinition | null;
  handleSlashCommandSelect: (command: SlashCommandDefinition) => void;
  resetSlashCommandState: () => void;
  activeSlashCommandId: string | null;
  clearActiveSlashCommand: () => void;
  isStreaming: boolean;
  prompt: string;
  setPrompt: (value: string) => void;
  setPromptFromHistory: (value: string) => void;
  attachedFilesLength: number;
  handleSend: () => void;
  hasActiveSlashCommand: boolean;
  isSheetMentionActive: boolean;
  sheetMentionCount: number;
  setSheetMentionHighlightIndex: (updater: (prev: number) => number) => void;
  highlightedSheetName: string | null;
  handleSheetMentionSelect: (sheetName: string) => void;
  resetSheetMentionState: () => void;
  handleNewChat: () => void;
  promptHistoryEntries: string[];
  promptHistoryIndex: number;
  isPromptHistoryLoading: boolean;
  loadPromptHistory: () => Promise<string[]>;
  setPromptHistoryIndex: React.Dispatch<React.SetStateAction<number>>;
  setIsNavigatingHistory: (value: boolean, source: string) => void;
  isNavigatingHistory: boolean;
}

type OverlayNavigationConfig<T> = {
  isActive: boolean;
  count: number;
  setHighlightIndex: (updater: (prev: number) => number) => void;
  highlightedItem: T | null;
  onSelect: (item: T) => void;
  onReset: () => void;
};

const handleOverlayNavigation = <T>(
  event: React.KeyboardEvent<HTMLTextAreaElement>,
  config: OverlayNavigationConfig<T>,
): boolean => {
  if (!config.isActive) return false;

  switch (event.key) {
    case "ArrowDown":
      event.preventDefault();
      if (config.count > 0) {
        config.setHighlightIndex((prev) =>
          Math.min(prev + 1, config.count - 1),
        );
      }
      return true;
    case "ArrowUp":
      event.preventDefault();
      if (config.count > 0) {
        config.setHighlightIndex((prev) => Math.max(prev - 1, 0));
      }
      return true;
    case "Enter":
      if (event.shiftKey) return false;
      event.preventDefault();
      if (config.highlightedItem) {
        config.onSelect(config.highlightedItem);
      }
      return true;
    case "Tab":
      event.preventDefault();
      if (config.highlightedItem) {
        config.onSelect(config.highlightedItem);
      }
      return true;
    case "Escape":
      event.preventDefault();
      config.onReset();
      return true;
    default:
      return false;
  }
};

const canNavigateHistory = (
  key: string,
  promptIsEmpty: boolean,
  isNavigatingHistory: boolean,
): boolean => {
  if (key === "ArrowUp") {
    return promptIsEmpty || isNavigatingHistory;
  }

  if (key === "ArrowDown") {
    return isNavigatingHistory;
  }

  return false;
};

export function createPromptKeyDownHandler({
  isSlashCommandActive,
  slashCommandCount,
  setSlashCommandHighlightIndex,
  highlightedSlashCommand,
  handleSlashCommandSelect,
  resetSlashCommandState,
  activeSlashCommandId,
  clearActiveSlashCommand,
  isStreaming,
  prompt,
  setPrompt,
  attachedFilesLength,
  handleSend,
  hasActiveSlashCommand,
  isSheetMentionActive,
  sheetMentionCount,
  setSheetMentionHighlightIndex,
  highlightedSheetName,
  handleSheetMentionSelect,
  resetSheetMentionState,
  handleNewChat,
  promptHistoryEntries,
  promptHistoryIndex,
  isPromptHistoryLoading,
  loadPromptHistory,
  setPromptHistoryIndex,
  setIsNavigatingHistory,
  isNavigatingHistory,
  setPromptFromHistory,
}: PromptKeyDownConfig): React.KeyboardEventHandler<HTMLTextAreaElement> {
  return (event) => {
    const mentionHandled = handleOverlayNavigation(event, {
      isActive: isSheetMentionActive,
      count: sheetMentionCount,
      setHighlightIndex: setSheetMentionHighlightIndex,
      highlightedItem: highlightedSheetName,
      onSelect: handleSheetMentionSelect,
      onReset: resetSheetMentionState,
    });
    if (mentionHandled) return;

    const slashHandled = handleOverlayNavigation(event, {
      isActive: isSlashCommandActive,
      count: slashCommandCount,
      setHighlightIndex: setSlashCommandHighlightIndex,
      highlightedItem: highlightedSlashCommand,
      onSelect: handleSlashCommandSelect,
      onReset: resetSlashCommandState,
    });
    if (slashHandled) return;

    const isNewChatShortcut =
      (event.ctrlKey || event.metaKey) &&
      !event.shiftKey &&
      event.key.toLowerCase() === "n";
    if (isNewChatShortcut) {
      event.preventDefault();
      handleNewChat();
      return;
    }

    // Ctrl+Shift+A (or Cmd+Shift+A on Mac) - Return focus to Excel
    const isReturnFocusShortcut =
      (event.ctrlKey || event.metaKey) &&
      event.shiftKey &&
      event.key.toLowerCase() === "a";
    if (isReturnFocusShortcut) {
      event.preventDefault();
      webViewBridge.send("focusHost", {});
      return;
    }

    // ESC key when slash command dropdown is NOT active
    if (event.key === "Escape") {
      event.preventDefault();
      // First priority: If there's text in the prompt, clear it
      if (prompt.trim()) {
        setPrompt("");
        setPromptHistoryIndex(-1);
        setIsNavigatingHistory(false, "escape-clear");
        return;
      }
      // Second priority: Clear active slash command panel if one is selected
      if (activeSlashCommandId) {
        clearActiveSlashCommand();
        return;
      }
      // Only return focus to Excel if no active command and prompt is empty
      webViewBridge.send("focusHost", {});
      return;
    }

    const promptIsEmpty = prompt.trim() === "";
    const applyHistoryEntry = (index: number, entries: string[]) => {
      resetSheetMentionState();
      const rawEntry = entries[index];
      const userMessage = getUserMessage(rawEntry).trim();
      console.info("[PromptHistory] applyHistoryEntry:before", {
        index,
        entry: rawEntry,
        userMessage,
      });
      setPromptHistoryIndex(index);
      setPromptFromHistory(userMessage);
      setIsNavigatingHistory(true, "history-entry");
      console.info("[PromptHistory] applyHistoryEntry:after", {
        index,
        entry: rawEntry,
        userMessage,
      });
    };

    if (event.key === "ArrowUp") {
      console.info("[PromptHistory] ArrowUp pressed", {
        promptIsEmpty,
        isNavigatingHistory,
        promptHistoryIndex,
      });
      if (!canNavigateHistory(event.key, promptIsEmpty, isNavigatingHistory)) {
        return;
      }

      event.preventDefault();
      if (isPromptHistoryLoading) return;

      void loadPromptHistory().then((entries) => {
        if (!entries.length) return;
        const nextIndex = isNavigatingHistory
          ? Math.min(promptHistoryIndex + 1, entries.length - 1)
          : 0;
        applyHistoryEntry(nextIndex, entries);
        setIsNavigatingHistory(true, "history-load");
      });
      return;
    }

    if (event.key === "ArrowDown") {
      console.info("[PromptHistory] ArrowDown pressed", {
        promptIsEmpty,
        isNavigatingHistory,
        promptHistoryIndex,
      });
      if (!canNavigateHistory(event.key, promptIsEmpty, isNavigatingHistory)) {
        return;
      }
      if (promptHistoryIndex <= -1) {
        return;
      }

      event.preventDefault();
      if (promptHistoryIndex === 0) {
        setPromptHistoryIndex(-1);
        setPrompt("");
        setIsNavigatingHistory(false, "arrow-down-reset");
        return;
      }

      applyHistoryEntry(promptHistoryIndex - 1, promptHistoryEntries);
      return;
    }

    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      if (
        !isStreaming &&
        (prompt.trim() || attachedFilesLength > 0 || hasActiveSlashCommand)
      ) {
        void handleSend();
      }
    }
  };
}
