import { type RefObject, useState } from "react";

import { useExcelIntegration } from "@/hooks/useExcelIntegration";

interface UseSelectedRangeOptions {
  isStreaming: boolean;
  textareaRef: RefObject<HTMLTextAreaElement>;
  showTools: boolean;
  showHistory: boolean;
  showCommands: boolean;
}

export function useSelectedRange({
  isStreaming,
  textareaRef,
  showTools,
  showHistory,
  showCommands,
}: UseSelectedRangeOptions) {
  const [selectedRange, setSelectedRange] = useState("");

  const { handleExcelRangeNavigate, handleExcelSheetNavigate } =
    useExcelIntegration({
      isStreaming,
      setSelectedRange,
      textareaRef,
      showTools,
      showHistory,
      showCommands,
    });

  const clearSelectedRange = () => setSelectedRange("");

  return {
    selectedRange,
    setSelectedRange,
    clearSelectedRange,
    handleExcelRangeNavigate,
    handleExcelSheetNavigate,
  };
}
