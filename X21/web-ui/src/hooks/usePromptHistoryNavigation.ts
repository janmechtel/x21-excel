import { useCallback, useRef, useState } from "react";

import { listRecentUserMessagesForFile } from "@/services/conversationHistoryService";
import { getUserMessage, isToolResultMessage } from "@/services/historyService";
import { webViewBridge } from "@/services/webViewBridge";

interface Params {
  setPrompt: (value: string) => void;
  updateSlashCommandTrigger: (value: string, caret: number) => void;
  updateMentionTrigger: (value: string, caret: number) => void;
  currentWorkbookKey: string | null;
  setCurrentWorkbookKey: (key: string) => void;
}

export function usePromptHistoryNavigation({
  setPrompt,
  updateSlashCommandTrigger,
  updateMentionTrigger,
  currentWorkbookKey,
  setCurrentWorkbookKey,
}: Params) {
  const [promptHistoryEntries, setPromptHistoryEntries] = useState<string[]>(
    [],
  );
  const [promptHistoryIndex, setPromptHistoryIndex] = useState(-1);
  const [isPromptHistoryLoading, setIsPromptHistoryLoading] = useState(false);
  const [isNavigatingHistory, setIsNavigatingHistoryState] = useState(false);
  const promptUpdateSourceRef = useRef<"history" | null>(null);

  const setIsNavigatingHistory = useCallback(
    (value: boolean, source: string) => {
      setIsNavigatingHistoryState((prev) => {
        console.info("[PromptHistory] isNavigatingHistory change", {
          from: prev,
          to: value,
          source,
        });
        return value;
      });
    },
    [],
  );

  const setPromptFromHistory = useCallback(
    (value: string) => {
      promptUpdateSourceRef.current = "history";
      console.info("[PromptHistory] setPromptFromHistory ->", value);
      setPrompt(value);
    },
    [setPrompt],
  );

  const loadPromptHistoryEntries = useCallback(
    async (limit = 50): Promise<string[]> => {
      if (isPromptHistoryLoading) {
        return [];
      }

      setIsPromptHistoryLoading(true);
      try {
        const workbookKey =
          currentWorkbookKey ?? (await webViewBridge.getWorkbookPath());
        if (!workbookKey) {
          console.warn(
            "[Prompt] Unable to fetch chat history: workbook key is unavailable.",
          );
          return [];
        }

        setCurrentWorkbookKey(workbookKey);
        const recentMessages = await listRecentUserMessagesForFile(
          workbookKey,
          limit,
        );
        const entries = recentMessages
          .filter(
            (item) =>
              item.role === "user" && !isToolResultMessage(item.content),
          )
          .map((item) => getUserMessage(item.content).trim())
          .filter((value): value is string => Boolean(value));

        console.log(
          `[Prompt] Latest ${entries.length} user prompts from arrow navigation:`,
          recentMessages,
        );

        setPromptHistoryEntries(entries);
        return entries;
      } catch (error) {
        console.error("[Prompt] Failed to fetch latest chat histories", error);
        return [];
      } finally {
        setIsPromptHistoryLoading(false);
      }
    },
    [currentWorkbookKey, isPromptHistoryLoading, setCurrentWorkbookKey],
  );

  const handlePromptChange = useCallback(
    (event: React.ChangeEvent<HTMLTextAreaElement>) => {
      const { value, selectionStart } = event.target;
      const updateSource = promptUpdateSourceRef.current;
      console.info("[PromptHistory] handlePromptChange:before", {
        updateSource,
        value,
        promptHistoryIndex,
        isNavigatingHistory,
      });
      setPrompt(value);
      if (updateSource === "history") {
        promptUpdateSourceRef.current = null;
      } else {
        if (value.trim()) {
          setIsNavigatingHistory(false, "user-typing");
        }
        setPromptHistoryIndex(-1);
      }
      updateSlashCommandTrigger(value, selectionStart ?? value.length);
      updateMentionTrigger(value, selectionStart ?? value.length);
      console.info("[PromptHistory] handlePromptChange:after", {
        updateSourceCleared: promptUpdateSourceRef.current,
        value,
        promptHistoryIndex,
        isNavigatingHistory,
      });
    },
    [
      isNavigatingHistory,
      promptHistoryIndex,
      setIsNavigatingHistory,
      setPrompt,
      updateMentionTrigger,
      updateSlashCommandTrigger,
    ],
  );

  const resetNavigationAfterSend = useCallback(() => {
    setPromptHistoryIndex(-1);
    setIsNavigatingHistory(false, "send");
  }, [setIsNavigatingHistory]);

  return {
    promptHistoryEntries,
    promptHistoryIndex,
    isPromptHistoryLoading,
    isNavigatingHistory,
    setIsNavigatingHistory,
    setPromptFromHistory,
    loadPromptHistoryEntries,
    setPromptHistoryIndex,
    handlePromptChange,
    resetNavigationAfterSend,
  };
}
