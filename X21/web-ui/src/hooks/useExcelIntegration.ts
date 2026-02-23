import { useCallback, useEffect, useRef } from "react";

import { useWebViewEvents } from "@/hooks/useWebViewEvents";
import { webViewBridge } from "@/services/webViewBridge";
import { webSocketChatService } from "@/services/webSocketChatService";

interface UseExcelIntegrationParams {
  isStreaming: boolean;
  setSelectedRange: (range: string) => void;
  textareaRef: React.RefObject<HTMLTextAreaElement>;
  showTools: boolean;
  showHistory: boolean;
  showCommands: boolean;
}

export function useExcelIntegration({
  isStreaming,
  setSelectedRange,
  textareaRef,
  showTools,
  showHistory,
  showCommands,
}: UseExcelIntegrationParams) {
  const lastSheetRef = useRef<string>("");

  const handleSelectionChanged = useCallback(
    async (data: any) => {
      if (isStreaming) return;

      const rawRange =
        typeof data?.selectedRange === "string" ? data.selectedRange : "";
      const activeSheetFromEvent =
        typeof data?.activeSheet === "string" ? data.activeSheet : "";
      if (!rawRange) {
        console.log(
          "[ExcelIntegration] Selection changed with no range, clearing.",
        );
        setSelectedRange("");
        return;
      }

      try {
        const workbookName = await webViewBridge.getWorkbookName();
        const activeSheet = activeSheetFromEvent || "";

        console.log("[ExcelIntegration] Excel context retrieved:", {
          workbookName,
          activeSheet,
        });

        console.log("[ExcelIntegration] Selection changed:", {
          rawRange,
          activeSheet,
        });

        const sheetName = activeSheet
          ? activeSheet.replace(/!+$/, "")
          : lastSheetRef.current;
        if (sheetName && sheetName !== lastSheetRef.current) {
          console.log("[ExcelIntegration] Sheet changed:", {
            previousSheet: lastSheetRef.current,
            currentSheet: sheetName,
            rawRange,
          });
        }
        const hasSheetInRange = rawRange.includes("!");

        const rangeWithSheet =
          !hasSheetInRange && sheetName ? `${sheetName}!${rawRange}` : rawRange;
        if (sheetName) {
          lastSheetRef.current = sheetName;
        }

        setSelectedRange(rangeWithSheet);
      } catch (error) {
        console.warn(
          "[ExcelIntegration] Failed to resolve sheet name from metadata",
          error,
        );
        // Fallback to raw range
        setSelectedRange(rawRange);
      }
    },
    [isStreaming, setSelectedRange],
  );

  const handleWorkbookContext = useCallback((data: any) => {
    const workbookPath =
      typeof data?.workbookPath === "string" ? data.workbookPath : "";
    const workbookName =
      typeof data?.workbookName === "string" ? data.workbookName : "";
    const workbookId = workbookPath || workbookName;
    if (!workbookId) return;

    webSocketChatService.setPinnedWorkbookIdentifier(workbookId);
  }, []);

  useWebViewEvents({
    onSelectionChanged: handleSelectionChanged,
    onWorkbookContext: handleWorkbookContext,
  });

  useEffect(() => {
    webViewBridge.send("FrontEndIsReady", {});
  }, []);

  useEffect(() => {
    const handleWindowFocus = () => {
      webViewBridge.send("webViewFocusChanged", { hasFocus: true });
    };

    const handleWindowBlur = () => {
      webViewBridge.send("webViewFocusChanged", { hasFocus: false });
    };

    window.addEventListener("focus", handleWindowFocus);
    window.addEventListener("blur", handleWindowBlur);

    return () => {
      window.removeEventListener("focus", handleWindowFocus);
      window.removeEventListener("blur", handleWindowBlur);
    };
  }, []);

  useEffect(() => {
    const handleFocusTextInput = () => {
      // Check if any overlay is open - if so, let the overlay handle focus
      const isAnyOverlayOpen = showTools || showHistory || showCommands;

      if (isAnyOverlayOpen) {
        console.log("[Focus] Overlay is open, skipping textarea focus");
        // The SearchableOverlay component will handle focus automatically
        return;
      }

      // No overlay open, focus the textarea as normal
      if (textareaRef.current) {
        console.log("[Focus] Focusing textarea");
        textareaRef.current.focus();
      }
    };

    webViewBridge.on("focusTextInput", handleFocusTextInput);
    return () => {
      webViewBridge.off("focusTextInput");
    };
  }, [textareaRef, showTools, showHistory, showCommands]);

  const handleExcelRangeNavigate = (range: string, workbookName?: string) => {
    try {
      webViewBridge.send("navigateToRange", { range, workbookName }, false);
    } catch (error) {
      console.error("Error navigating to range:", error);
    }
  };

  const handleExcelSheetNavigate = (sheetName: string) => {
    try {
      const cleanedSheet = sheetName.replace(/^@/, "").replace(/!$/, "").trim();
      const needsQuoting =
        /\s/.test(cleanedSheet) || cleanedSheet.includes("'");
      const safeSheet = needsQuoting
        ? `'${cleanedSheet.replace(/'/g, "''")}'`
        : cleanedSheet;
      const target = `${safeSheet}!A1`;
      webViewBridge.send("navigateToRange", { range: target }, false);
    } catch (error) {
      console.error("Error navigating to sheet:", error);
    }
  };

  return { handleExcelRangeNavigate, handleExcelSheetNavigate };
}
