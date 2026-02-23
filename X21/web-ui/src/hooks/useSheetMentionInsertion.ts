import { useCallback } from "react";

import type { RefObject } from "react";
import type { SheetMentionState } from "./useSheetMentionState";

interface UseSheetMentionInsertionParams {
  prompt: string;
  setPrompt: (value: string) => void;
  sheetMentionState: SheetMentionState | null;
  resetMentionState: () => void;
  textareaRef: RefObject<HTMLTextAreaElement>;
}

export function useSheetMentionInsertion({
  prompt,
  setPrompt,
  sheetMentionState,
  resetMentionState,
  textareaRef,
}: UseSheetMentionInsertionParams) {
  const formatSheetNameForInsertion = useCallback((sheetName: string) => {
    const trimmed = sheetName.trim();
    const escaped = trimmed.replace(/'/g, "''");

    // Include trailing "!" so the mention is parsed and rendered as a pill, plus a newline to move the caret after.
    return `'${escaped}'!\n`;
  }, []);

  const formatWorkbookSheetForInsertion = useCallback(
    (workbookName: string, sheetName: string) => {
      const quoteIfNeeded = (value: string) => {
        const needsQuotes = /[\s,]/.test(value) || value.includes("'");
        return needsQuotes ? `'${value.replace(/'/g, "''")}'` : value;
      };

      const trimmedWorkbook = workbookName.trim().replace(/^\[|\]$/g, "");
      const safeWorkbook = trimmedWorkbook;
      const safeSheet = quoteIfNeeded(sheetName.trim());

      // Use [Workbook]Sheet! to insert a sheet reference (without a cell range).
      return `[${safeWorkbook}]${safeSheet}!\n`;
    },
    [],
  );

  const handleSheetMentionSelect = useCallback(
    (sheetName: string) => {
      if (!sheetMentionState) return;

      const workbookSplit = sheetName.split("/", 2);
      const looksLikeWorkbook =
        workbookSplit.length === 2 &&
        /\.(xlsx|xlsm|xls|xlsb)$/i.test(workbookSplit[0]);
      const insertion = looksLikeWorkbook
        ? formatWorkbookSheetForInsertion(workbookSplit[0], workbookSplit[1])
        : formatSheetNameForInsertion(sheetName);
      const before = prompt.slice(0, sheetMentionState.startIndex);
      const after = prompt.slice(sheetMentionState.endIndex);
      // Replace the trigger with the mention text (no leading "@") to avoid re-triggering the palette.
      const updatedPrompt = `${before}${insertion}${after}`;
      console.log("Inserting sheet mention:", {
        sheetName,
        insertion,
        updatedPrompt,
      });
      resetMentionState();
      setPrompt(updatedPrompt);

      requestAnimationFrame(() => {
        if (!textareaRef.current) return;
        const caretPosition = before.length + insertion.length;
        textareaRef.current.setSelectionRange(caretPosition, caretPosition);
        textareaRef.current.focus();
      });
    },
    [
      formatWorkbookSheetForInsertion,
      formatSheetNameForInsertion,
      prompt,
      resetMentionState,
      setPrompt,
      sheetMentionState,
      textareaRef,
    ],
  );

  return { formatSheetNameForInsertion, handleSheetMentionSelect };
}
