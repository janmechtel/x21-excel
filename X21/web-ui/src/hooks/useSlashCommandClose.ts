import { useCallback } from "react";
import type React from "react";

import type { SlashCommandState } from "./useSlashCommandState";

interface UseSlashCommandCloseOptions {
  prompt: string;
  setPrompt: (value: string) => void;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
  resetSlashCommandState: () => void;
  slashCommandState: SlashCommandState | null;
}

export function useSlashCommandClose({
  prompt,
  setPrompt,
  textareaRef,
  resetSlashCommandState,
  slashCommandState,
}: UseSlashCommandCloseOptions) {
  return useCallback(() => {
    if (slashCommandState) {
      const { startIndex, endIndex } = slashCommandState;
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
    resetSlashCommandState();
  }, [
    slashCommandState,
    prompt,
    resetSlashCommandState,
    setPrompt,
    textareaRef,
  ]);
}
