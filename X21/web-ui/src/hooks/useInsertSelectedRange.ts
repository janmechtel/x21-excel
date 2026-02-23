import { useCallback } from "react";
import type { OperationStatus } from "@/types/chat";
import { OperationStatusValues } from "@/types/chat";

interface UseInsertSelectedRangeParams {
  selectedRange: string;
  isStreaming: boolean;
  loadingState: OperationStatus;
  getPendingToolCount: () => number;
  setPrompt: (value: string | ((prev: string) => string)) => void;
  updateSlashCommandTrigger: (value: string, cursor: number) => void;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
}

export function useInsertSelectedRange({
  selectedRange,
  isStreaming,
  loadingState,
  getPendingToolCount,
  setPrompt,
  updateSlashCommandTrigger,
  textareaRef,
}: UseInsertSelectedRangeParams) {
  const formatRangeWithSheet = useCallback((raw: string) => {
    if (!raw) return "";

    const [sheetPart, rangePart] = raw.split("!");
    if (!rangePart) {
      return raw;
    }

    const normalizedSheet = sheetPart?.replace(/^'+|'+$/g, "").trim() ?? "";
    const safeSheet = normalizedSheet
      ? `'${normalizedSheet.replace(/'/g, "''")}'`
      : "";

    const cleanedRange = rangePart.trim();

    return safeSheet ? `${safeSheet}!${cleanedRange}` : cleanedRange;
  }, []);

  return useCallback(() => {
    if (
      !selectedRange ||
      isStreaming ||
      loadingState !== OperationStatusValues.IDLE ||
      getPendingToolCount() > 0
    ) {
      return;
    }

    const formattedRange = formatRangeWithSheet(selectedRange);

    setPrompt((prev) => {
      const nextValue = `${prev}${formattedRange}\n`;
      console.log(
        "[useInsertSelectedRange] Inserting selected range into prompt:",
        formattedRange,
      );

      requestAnimationFrame(() => {
        updateSlashCommandTrigger(nextValue, nextValue.length);
        if (textareaRef.current) {
          textareaRef.current.focus();
          textareaRef.current.setSelectionRange(
            nextValue.length,
            nextValue.length,
          );
        }
      });

      return nextValue;
    });
  }, [
    selectedRange,
    isStreaming,
    loadingState,
    getPendingToolCount,
    formatRangeWithSheet,
    setPrompt,
    updateSlashCommandTrigger,
    textareaRef,
  ]);
}
