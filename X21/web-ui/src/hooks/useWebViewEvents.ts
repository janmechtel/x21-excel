import { useEffect } from "react";
import { webViewBridge } from "../services/webViewBridge";

interface UseWebViewEventsProps {
  onStreamDelta?: (data: any) => void;
  onStreamComplete?: (data: any) => void;
  onStreamError?: (error: any) => void;
  onSelectionChanged?: (data: any) => void;
  onWorkbookContext?: (data: any) => void;
}

export const useWebViewEvents = ({
  onStreamDelta,
  onStreamComplete,
  onStreamError,
  onSelectionChanged,
  onWorkbookContext,
}: UseWebViewEventsProps) => {
  useEffect(() => {
    const handleStreamDelta = (data: any) => {
      onStreamDelta?.(data);
    };

    const handleStreamComplete = (data: any) => {
      onStreamComplete?.(data);
    };

    const handleStreamError = (data: any) => {
      onStreamError?.(data);
    };

    const handleSelectionChanged = (data: any) => {
      onSelectionChanged?.(data);
    };

    const handleWorkbookContext = (data: any) => {
      onWorkbookContext?.(data);
    };

    if (onStreamDelta) webViewBridge.on("streamDelta", handleStreamDelta);
    if (onStreamComplete) {
      webViewBridge.on("streamComplete", handleStreamComplete);
    }
    if (onStreamError) webViewBridge.on("streamError", handleStreamError);
    if (onSelectionChanged) {
      webViewBridge.on("selectionChanged", handleSelectionChanged);
    }
    if (onWorkbookContext) {
      webViewBridge.on("workbookContext", handleWorkbookContext);
    }

    return () => {
      webViewBridge.off("streamDelta");
      webViewBridge.off("streamComplete");
      webViewBridge.off("streamError");
      webViewBridge.off("selectionChanged");
      webViewBridge.off("workbookContext");
    };
  }, [
    onStreamDelta,
    onStreamComplete,
    onStreamError,
    onSelectionChanged,
    onWorkbookContext,
  ]);
};
