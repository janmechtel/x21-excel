import { useCallback } from "react";
import type React from "react";

import type { SheetMentionState } from "./useSheetMentionState";

interface UseSheetMentionCloseOptions {
  prompt: string;
  setPrompt: (value: string) => void;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
  resetMentionState: () => void;
  sheetMentionState: SheetMentionState | null;
}

export function useSheetMentionClose({
  prompt,
  setPrompt,
  textareaRef,
  resetMentionState,
  sheetMentionState,
}: UseSheetMentionCloseOptions) {
  return useCallback(() => {
    if (sheetMentionState) {
      const { startIndex, endIndex } = sheetMentionState;
      const before = prompt.slice(0, startIndex);
      const after = prompt.slice(endIndex);
      const nextPrompt = `${before}${after}`;

      if (nextPrompt !== prompt) {
        setPrompt(nextPrompt);
        requestAnimationFrame(() => {
          if (!textareaRef.current) return;
          const caret = before.length;
          textareaRef.current.selectionStart = caret;
          textareaRef.current.selectionEnd = caret;
          textareaRef.current.focus();
        });
      }
    }
    resetMentionState();
  }, [sheetMentionState, prompt, resetMentionState, setPrompt, textareaRef]);
}
