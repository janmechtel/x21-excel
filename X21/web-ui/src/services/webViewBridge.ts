import { WebViewBridgeMessage } from "../types";

class WebViewBridge {
  private messageHandlers: Map<string, (data: any) => void> = new Map();
  private requestId = 0;
  private pendingRequests: Map<number, (response: any) => void> = new Map();

  constructor() {
    if (typeof window !== "undefined" && (window as any).chrome?.webview) {
      (window as any).chrome.webview.addEventListener(
        "message",
        this.handleMessage.bind(this),
      );
    }
  }

  private handleMessage(event: MessageEvent) {
    try {
      console.log("[DEBUG] handleMessage - Raw message:", event, event.data);

      // Some browser/devtools messages arrive as plain objects that aren't part of the bridge contract.
      if (
        typeof event.data !== "string" &&
        (typeof event.data !== "object" || event.data === null)
      ) {
        console.warn(
          "[DEBUG] Ignoring non-string, non-object WebView message:",
          event.data,
        );
        return;
      }

      let message: WebViewBridgeMessage;

      if (typeof event.data === "string") {
        message = JSON.parse(event.data);
      } else if ("type" in event.data) {
        message = event.data as WebViewBridgeMessage;
      } else {
        console.warn(
          "[DEBUG] Ignoring message without a type (likely devtools noise):",
          event.data,
        );
        return;
      }

      console.log("[DEBUG] handleMessage - Parsed message:", message);

      if (message.type === "response" && message.payload.requestId) {
        const requestId = message.payload.requestId;
        const requestIdNum =
          typeof requestId === "number" ? requestId : parseInt(requestId, 10);
        console.log(
          "[DEBUG] handleMessage - Looking for resolver, requestId:",
          requestId,
          "type:",
          typeof requestId,
        );
        console.log(
          "[DEBUG] handleMessage - pendingRequests has:",
          Array.from(this.pendingRequests.keys()),
        );
        if (Number.isNaN(requestIdNum)) {
          console.error(
            "[DEBUG] handleMessage - Invalid requestId, cannot resolve:",
            requestId,
          );
          return;
        }
        const resolver = this.pendingRequests.get(requestIdNum);
        console.log("[DEBUG] handleMessage - Found resolver:", !!resolver);
        if (resolver) {
          console.log(
            "[DEBUG] handleMessage - Calling resolver with data:",
            message.payload.data,
          );
          resolver(message.payload.data);
          this.pendingRequests.delete(requestIdNum);
          console.log("[DEBUG] handleMessage - Resolver called successfully");
        } else {
          console.error(
            "[DEBUG] handleMessage - NO RESOLVER FOUND for requestId:",
            requestId,
          );
        }
        return;
      }
      console.log(
        "[DEBUG] handleMessage - Not a response with requestId, continuing...",
      );

      if (message.type === "event") {
        const eventType = (message as any).eventType;
        const eventData = (message as any).data;
        console.log("Event received:", eventType, eventData);

        const handler = this.messageHandlers.get(eventType);
        if (handler) {
          handler(eventData);
        } else {
          console.warn("No handler for event type:", eventType);
        }
        return;
      }

      const handler = this.messageHandlers.get(message.type);
      if (handler) {
        handler(message.payload);
      } else {
        console.warn("No handler for message type:", message.type);
      }
    } catch (error) {
      console.error(
        "Error parsing WebView message:",
        error,
        "Raw data:",
        event.data,
      );
    }
  }

  on(messageType: string, handler: (data: any) => void) {
    this.messageHandlers.set(messageType, handler);
  }

  off(messageType: string) {
    this.messageHandlers.delete(messageType);
  }

  async send<T = any>(
    type: string,
    payload: any = {},
    expectResponse: boolean = false,
  ): Promise<T> {
    if (typeof window === "undefined" || !(window as any).chrome?.webview) {
      console.warn("WebView2 not available");
      return Promise.resolve(undefined as T);
    }

    const message: WebViewBridgeMessage = { type, payload };

    if (expectResponse) {
      const requestId = ++this.requestId;
      message.payload = { ...payload, requestId };

      console.log("Sending message with response expected:", message);

      return new Promise<T>((resolve, reject) => {
        const timeoutId = setTimeout(() => {
          this.pendingRequests.delete(requestId);
          reject(new Error(`Request timeout for ${type}`));
        }, 60000 * 30); // 30 second timeout

        this.pendingRequests.set(requestId, (response) => {
          clearTimeout(timeoutId);
          resolve(response);
        });

        (window as any).chrome.webview.postMessage(JSON.stringify(message));
      });
    } else {
      console.log("Sending message without response:", message);
      (window as any).chrome.webview.postMessage(JSON.stringify(message));
      return Promise.resolve(undefined as T);
    }
  }
}

const bridgeInstance = new WebViewBridge();

// Convenience methods for common operations
export const webViewBridge = Object.assign(bridgeInstance, {
  async getWorkbookName(): Promise<string> {
    try {
      const result = await bridgeInstance.send<string>(
        "getWorkbookName",
        {},
        true,
      );
      // Empty string indicates no active workbook; do not fabricate a placeholder
      return result || "";
    } catch (error) {
      console.warn("Failed to get workbook name from bridge:", error);
      return "";
    }
  },

  async getWebSocketUrl(): Promise<string> {
    console.log("[DEBUG] getWebSocketUrl() called - about to send request");
    try {
      console.log("[DEBUG] Calling bridgeInstance.send()...");
      const result = await bridgeInstance.send<string>(
        "getWebSocketUrl",
        {},
        true,
      );
      console.log("[DEBUG] bridgeInstance.send() returned:", result);
      return result || "ws://localhost:8000";
    } catch (error) {
      console.warn("[DEBUG] Failed to get WebSocket URL from bridge:", error);
      return "ws://localhost:8000";
    }
  },

  async getWorksheetNames(): Promise<string[]> {
    try {
      const result = await bridgeInstance.send<string[] | string>(
        "getWorksheetNames",
        {},
        true,
      );
      if (Array.isArray(result)) {
        return result;
      }
      if (typeof result === "string") {
        return result
          .split(",")
          .map((name) => name.trim())
          .filter(Boolean);
      }
      return [];
    } catch (error) {
      console.warn("Failed to get worksheet names from bridge:", error);
      return [];
    }
  },

  async getWorkbookPath(): Promise<string> {
    try {
      const result = await bridgeInstance.send<string>(
        "getWorkbookPath",
        {},
        true,
      );
      return result || "";
    } catch (error) {
      console.warn("Failed to get workbook path from bridge:", error);
      return "";
    }
  },

  async pickFolder(options?: {
    allowFileListing?: boolean;
    extensions?: string[];
  }): Promise<{ path: string; files?: string[] } | null> {
    try {
      const payload = {
        allowFileListing: options?.allowFileListing ?? true,
        extensions: options?.extensions,
      };
      const result = await bridgeInstance.send<
        { path?: string; files?: string[] } | string
      >("pickFolder", payload, true);
      if (typeof result === "string") {
        return result ? { path: result, files: undefined } : null;
      }
      if (result && typeof result === "object" && result.path) {
        return {
          path: result.path,
          files: Array.isArray(result.files) ? result.files : undefined,
        };
      }
    } catch (error) {
      console.warn("Failed to pick folder from bridge:", error);
    }
    return null;
  },

  async pickFile(options?: {
    extensions?: string[];
    title?: string;
    filterLabel?: string;
  }): Promise<string | null> {
    try {
      const payload = {
        extensions: options?.extensions || [".xlsx", ".xlsm", ".xls"],
        title: options?.title,
        filterLabel: options?.filterLabel,
      };
      const result = await bridgeInstance.send<string>(
        "pickFile",
        payload,
        true,
      );
      return result || null;
    } catch (error) {
      console.warn("Failed to pick file from bridge:", error);
    }
    return null;
  },

  async requestRangeSelection(): Promise<string | null> {
    try {
      const result = await bridgeInstance.send<string>(
        "requestRangeSelection",
        {},
        true,
      );
      if (typeof result === "string" && result.trim().length > 0) {
        return result.trim();
      }
    } catch (error) {
      console.warn("Failed to request range selection from bridge:", error);
    }
    return null;
  },

  async openWorkbook(filePath: string): Promise<boolean> {
    console.log("[webViewBridge] openWorkbook called with filePath:", filePath);
    try {
      console.log("[webViewBridge] Sending openWorkbook message to C#...");
      const result = await bridgeInstance.send<{ success: boolean }>(
        "openWorkbook",
        { filePath },
        true,
      );
      console.log("[webViewBridge] Received response from C#:", result);
      const success = result?.success ?? false;
      console.log("[webViewBridge] Final success value:", success);
      return success;
    } catch (error) {
      console.error(
        "[webViewBridge] Failed to open workbook from bridge:",
        error,
      );
      return false;
    }
  },
});
