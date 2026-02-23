import { useEffect, useMemo, useState } from "react";
import type { Dispatch, SetStateAction } from "react";

import {
  findSlashCommandById,
  refreshExcelSlashCommands,
  searchSlashCommands,
  subscribeToSlashCommands,
} from "@/lib/slashCommands";
import type { SlashCommandDefinition } from "@/types/slash-commands";

export interface SlashCommandState {
  query: string;
  startIndex: number;
  endIndex: number;
}

interface UseSlashCommandStateResult {
  slashCommandState: SlashCommandState | null;
  slashCommandMatches: SlashCommandDefinition[];
  slashCommandHighlightIndex: number;
  setSlashCommandHighlightIndex: Dispatch<SetStateAction<number>>;
  activeSlashCommandId: string | null;
  setActiveSlashCommandId: (id: string | null) => void;
  activeSlashCommand: SlashCommandDefinition | null;
  highlightedSlashCommand: SlashCommandDefinition | null;
  isSlashCommandActive: boolean;
  promptPlaceholder: string;
  updateSlashCommandTrigger: (
    value: string,
    cursorPosition: number | null,
  ) => void;
  resetSlashCommandState: () => void;
  clearActiveSlashCommand: () => void;
  openSlashCommandPalette: (prompt: string) => void;
}

export function useSlashCommandState(): UseSlashCommandStateResult {
  const [slashCommandState, setSlashCommandState] =
    useState<SlashCommandState | null>(null);
  const [slashCommandHighlightIndex, setSlashCommandHighlightIndex] =
    useState(0);
  const [activeSlashCommandId, setActiveSlashCommandId] = useState<
    string | null
  >(null);
  const [slashCommandRevision, setSlashCommandRevision] = useState(0);

  const slashCommandMatches = useMemo(
    () =>
      slashCommandState ? searchSlashCommands(slashCommandState.query) : [],
    [slashCommandState, slashCommandRevision],
  );

  const slashCommandQuery = slashCommandState?.query ?? "";
  const slashCommandCount = slashCommandMatches.length;

  const activeSlashCommand = useMemo(
    () => findSlashCommandById(activeSlashCommandId),
    [activeSlashCommandId, slashCommandRevision],
  );

  const isSlashCommandActive = Boolean(slashCommandState);

  const highlightedSlashCommand =
    isSlashCommandActive && slashCommandMatches.length > 0
      ? slashCommandMatches[
          Math.min(slashCommandHighlightIndex, slashCommandMatches.length - 1)
        ]
      : null;

  const promptPlaceholder = activeSlashCommand
    ? activeSlashCommand.inputPlaceholder ||
      `/${activeSlashCommand.id} selected — describe the parameters or context to apply`
    : "Plan, build, fix anything";

  useEffect(() => {
    if (!slashCommandState) {
      setSlashCommandHighlightIndex(0);
      return;
    }

    if (slashCommandCount === 0) {
      setSlashCommandHighlightIndex(0);
      return;
    }

    setSlashCommandHighlightIndex((prev) =>
      Math.min(prev, slashCommandCount - 1),
    );
  }, [slashCommandState, slashCommandCount]);

  useEffect(() => {
    if (slashCommandState) {
      setSlashCommandHighlightIndex(0);
    }
  }, [slashCommandQuery, slashCommandState]);

  useEffect(() => {
    const unsubscribe = subscribeToSlashCommands(() =>
      setSlashCommandRevision((rev) => rev + 1),
    );
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    refreshExcelSlashCommands();
  }, []);

  useEffect(() => {
    if (slashCommandState) {
      refreshExcelSlashCommands();
    }
  }, [slashCommandState]);

  const resetSlashCommandState = () => {
    setSlashCommandState(null);
    setSlashCommandHighlightIndex(0);
  };

  const updateSlashCommandTrigger = (
    value: string,
    cursorPosition: number | null,
  ) => {
    if (cursorPosition === null || cursorPosition === undefined) {
      resetSlashCommandState();
      return;
    }

    const uptoCursor = value.slice(0, cursorPosition);
    const slashMatch = uptoCursor.match(/(^|\s)\/([^\s]*)$/);

    if (!slashMatch) {
      setSlashCommandState(null);
      return;
    }

    const query = slashMatch[2] ?? "";
    const startIndex = cursorPosition - query.length - 1;
    const endIndex = cursorPosition;

    setSlashCommandState({ query, startIndex, endIndex });
  };

  const clearActiveSlashCommand = () => {
    setActiveSlashCommandId(null);
  };

  const openSlashCommandPalette = (currentPrompt: string) => {
    const cursorPos = currentPrompt.length;
    updateSlashCommandTrigger(`${currentPrompt}/`, cursorPos + 1);
    refreshExcelSlashCommands();
  };

  return {
    slashCommandState,
    slashCommandMatches,
    slashCommandHighlightIndex,
    setSlashCommandHighlightIndex,
    activeSlashCommandId,
    setActiveSlashCommandId,
    activeSlashCommand: activeSlashCommand || null,
    highlightedSlashCommand,
    isSlashCommandActive,
    promptPlaceholder,
    updateSlashCommandTrigger,
    resetSlashCommandState,
    clearActiveSlashCommand,
    openSlashCommandPalette,
  };
}
