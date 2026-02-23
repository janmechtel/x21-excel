import { createWsClient, WSHandlers } from "./wsClient";
import { webViewBridge } from "./webViewBridge";
import { toolsNotRequiringApproval } from "@/utils/tools";
import { WebSocketMessageTypes, ToolNames } from "@/types/chat";

// Dedicated error class for Deno server connection failures
export class DenoServerConnectionError extends Error {
  constructor(
    message: string = "Failed to connect to Deno server - please restart Excel",
  ) {
    super(message);
    this.name = "DenoServerConnectionError";
  }
}

interface StreamWorkflowPayload {
  prompt: string;
  activeTools: string[];
  autoApproveEnabled: boolean;
  workbookName?: string;
  allowOtherWorkbookReads?: boolean;
  documentsBase64?: Array<{
    name: string;
    type: string;
    size: number;
    base64: string;
  }>;
}

interface WSMessage {
  type: string;
  [key: string]: any;
}

const WORKBOOK_CACHE_MS = 250;
let cachedWorkbookId = "";
let cachedWorkbookAt = 0;
let pinnedWorkbookId = "";

export class WebSocketChatService {
  private wsClient: any = null;
  private isConnected = false;
  private isConnecting = false;
  private reconnectTimeout: number | null = null;
  private autoApproveEnabled: boolean = false;
  private approveAllEnabled: boolean = false;
  private currentUserEmail: string | null = null;

  // Event handlers
  private onStreamDelta: ((data: string) => void) | null = null;
  private onStreamComplete: ((data: any) => void) | null = null;
  private onStreamError: ((error: any) => void) | null = null;
  private onConnectionChange: ((connected: boolean) => void) | null = null;
  private onStatusUpdate: ((data: any) => void) | null = null;
  private onTokenUpdate: ((data: any) => void) | null = null;
  private onChangeSummary: ((data: any) => void) | null = null;

  constructor() {
    this.connect();

    if (typeof window !== "undefined" && import.meta.env.DEV) {
      (window as any).__x21DebugSocketMessage = (message: WSMessage) => {
        this.handleWebSocketMessage(message);
      };
      (window as any).__x21WebSocketService = this;
    }
  }

  setPinnedWorkbookIdentifier(workbookId: string): void {
    const nextId = workbookId?.trim() ?? "";
    if (!nextId) {
      return;
    }
    if (nextId === pinnedWorkbookId) {
      return;
    }

    pinnedWorkbookId = nextId;
    cachedWorkbookId = nextId;
    cachedWorkbookAt = Date.now();

    if (this.isConnected && this.wsClient) {
      const message = {
        type: "workbook:register",
        workbookName: nextId,
      };
      this.wsClient.send(message);
      console.log("Registered workbook with WebSocket:", nextId);
    }
  }

  // Set event handlers
  setEventHandlers(handlers: {
    onStreamDelta?: (data: string) => void;
    onStreamComplete?: (data: any) => void;
    onStreamError?: (error: any) => void;
    onConnectionChange?: (connected: boolean) => void;
    onStatusUpdate?: (data: any) => void;
    onTokenUpdate?: (data: any) => void;
    onChangeSummary?: (data: any) => void;
  }) {
    this.onStreamDelta = handlers.onStreamDelta || null;
    this.onStreamComplete = handlers.onStreamComplete || null;
    this.onStreamError = handlers.onStreamError || null;
    this.onConnectionChange = handlers.onConnectionChange || null;
    this.onStatusUpdate = handlers.onStatusUpdate || null;
    this.onTokenUpdate = handlers.onTokenUpdate || null;
    this.onChangeSummary = handlers.onChangeSummary || null;
  }

  private async connect() {
    console.log("[DEBUG] connect() called");
    if (this.isConnected || this.wsClient || this.isConnecting) {
      console.log("Already connected, skipping...");
      return;
    }

    this.isConnecting = true;
    try {
      console.log("[DEBUG] About to call webViewBridge.getWebSocketUrl()");
      // Get dynamic WebSocket URL from backend
      const wsUrl = await webViewBridge.getWebSocketUrl();
      console.log(`[DEBUG] Got WebSocket URL: ${wsUrl}`);

      this.wsClient = createWsClient(wsUrl);
    } catch (error) {
      this.isConnecting = false;
      console.error("Failed to initialize WebSocket client:", error);
      throw error;
    }

    const handlers: WSHandlers = {
      onOpen: () => {
        console.log("WebSocket connected");
        this.isConnected = true;
        this.isConnecting = false;
        this.onConnectionChange?.(true);

        // Automatically send user email when connection opens
        this.sendUserEmailOnConnect();

        // Register workbook with WebSocket immediately
        this.registerWorkbookOnConnect();
      },

      onMessage: (data: WSMessage) => {
        console.log("WebSocket message received:", data);
        this.handleWebSocketMessage(data);
      },

      onClose: (event) => {
        console.log("WebSocket disconnected:", event);
        this.isConnected = false;
        this.isConnecting = false;
        this.onConnectionChange?.(false);
        this.scheduleReconnect();
      },

      onError: (event) => {
        console.error("WebSocket error:", event);
        this.isConnecting = false;
        this.scheduleReconnect();
      },
    };

    this.wsClient.connect(handlers);
  }

  // Add method to update the setting
  setAutoApprove(enabled: boolean) {
    this.autoApproveEnabled = enabled;
  }

  // Add method to update the approve all setting
  setApproveAll(enabled: boolean) {
    this.approveAllEnabled = enabled;
  }

  private async handleWebSocketMessage(data: WSMessage) {
    const type = data.type;
    const payload = data?.payload;

    console.log("WebSocket message received:", type, payload);

    switch (type) {
      case "stream:delta":
        if (!payload) {
          console.error("WebSocket message payload is missing");
          return;
        }

        this.onStreamDelta?.(JSON.stringify(payload));
        break;

      case "stream:end":
        this.onStreamComplete?.(payload);
        break;

      case "stream:error":
        this.onStreamError?.(payload);
        break;

      case WebSocketMessageTypes.STREAM_CANCELLED:
        console.log("Stream was cancelled by user");
        // Notify UI that stream was cancelled
        this.onStreamComplete?.({
          type: WebSocketMessageTypes.STREAM_CANCELLED,
          message: payload?.message || "Request was cancelled",
          requestId: payload?.requestId,
        });
        break;

      case "status:update":
        this.onStatusUpdate?.(payload);
        break;

      case "stream:token_update":
        this.onTokenUpdate?.(payload);
        break;

      case "welcome":
        // Surface welcome/info events to the UI so they can be rendered
        this.onStreamComplete?.({
          type: "system:welcome",
          message: payload?.message || data.message || "Connected to server",
        });
        break;

      case WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY:
        // Handle workbook change summary notifications
        console.log("Workbook change summary received:", payload);
        this.onChangeSummary?.(payload);
        break;

      case WebSocketMessageTypes.TOOL_PERMISSION:
        const toolPermissions = data.payload?.toolPermissions || [data.payload];

        if (!Array.isArray(toolPermissions) || toolPermissions.length === 0) {
          console.error("No tool permissions provided");
          break;
        }

        // Check if any tools need manual approval
        const toolsNeedingApproval = toolPermissions.filter((tool: any) => {
          const toolName = tool.toolName;
          const toolDoesNotRequireApproval =
            toolsNotRequiringApproval.includes(toolName);
          const requiresManualApproval =
            toolName === ToolNames.LIST_OPEN_WORKBOOKS;
          const shouldAutoApprove =
            !requiresManualApproval &&
            (this.autoApproveEnabled ||
              this.approveAllEnabled ||
              toolDoesNotRequireApproval);

          console.log(
            `Tool ${toolName} does not require approval: ${toolDoesNotRequireApproval}`,
          );
          console.log(
            `Tool ${toolName} should auto-approve: ${shouldAutoApprove}`,
          );
          return !shouldAutoApprove;
        });

        if (toolsNeedingApproval.length === 0) {
          // Auto-approve all tools by sending a single tool:permission:response
          console.log(`Auto-approving ${toolPermissions.length} tools`);

          const workbookName: string = await getWorkbookIdentifier();

          // Create permission response with all tools approved
          const toolResponses: ToolResponse[] = toolPermissions.map(
            (tool: any) => ({
              toolId: tool.toolId,
              decision: "approved" as const,
            }),
          );

          const permissionResponse = {
            type: "tool:permission:response",
            workbookName: workbookName,
            groupId: `group_${Date.now()}`,
            toolResponses: toolResponses,
          };

          // Send permission response to backend
          this.ensureConnectedAndSend(permissionResponse)
            .then(() => {
              console.log(
                `✅ Sent tool:permission:response with ${toolResponses.length} auto-approved tools`,
              );
            })
            .catch((error) => {
              console.error("Failed to send auto-approval response:", error);
            });

          // Notify UI to update state (mark tools as approved)
          this.onStreamComplete?.({
            type: WebSocketMessageTypes.TOOL_AUTO_APPROVED,
            toolIds: toolPermissions.map((t: any) => t.toolId),
            message: `${toolPermissions.length} tools auto-approved`,
          });
        } else {
          // Send to UI for manual approval
          this.onStreamComplete?.({
            type: WebSocketMessageTypes.TOOL_PERMISSION,
            toolPermissions: toolPermissions,
            message: `${toolPermissions.length} tools need approval`,
          });
        }
        break;

      case WebSocketMessageTypes.TOOL_ERROR:
        const errorToolId = data.payload?.toolId;
        const errorMessage = data.payload?.errorMessage;
        const error = data.payload?.error;

        console.log(
          `Tool error received for ${errorToolId}:`,
          errorMessage,
          error,
        );

        // Send to UI for error state handling
        this.onStreamComplete?.({
          type: WebSocketMessageTypes.TOOL_ERROR,
          toolId: errorToolId,
          errorMessage: errorMessage,
          error: error,
          message: "Tool execution failed",
        });
        break;

      case "email:request":
        console.log("Server requesting user email");
        this.handleEmailRequest();
        break;

      case WebSocketMessageTypes.UI_REQUEST: {
        const toolUseId = data.toolUseId || payload?.toolUseId;
        const requestPayload = payload?.payload || payload;
        if (!toolUseId || !requestPayload) {
          console.error("Invalid ui:request payload received");
          break;
        }
        this.onStreamComplete?.({
          type: WebSocketMessageTypes.UI_REQUEST,
          toolUseId,
          payload: requestPayload,
        });
        break;
      }

      default:
        console.log("Unknown WebSocket message type:", data.type, data);
        break;
    }
  }

  /**
   * Internal method to ensure connection and send message
   * This centralizes the connection checking and reconnection logic
   */
  private async ensureConnectedAndSend(message: any): Promise<void> {
    // Check if we're connected, if not try to reconnect
    if (!this.isConnected || !this.wsClient) {
      console.log("WebSocket not connected, attempting to reconnect...");
      const reconnected = await this.tryReconnect();

      if (!reconnected) {
        console.error(
          "Failed to establish WebSocket connection to Deno server",
        );
        throw new DenoServerConnectionError(
          "Failed to establish WebSocket connection to Deno server",
        );
      }
    }

    console.log("Sending WebSocket message:", message);
    const sent = this.wsClient.send(message);
    if (!sent) {
      throw new DenoServerConnectionError(
        "Failed to send WebSocket message - connection lost",
      );
    }
  }

  async sendMessage(payload: StreamWorkflowPayload): Promise<boolean> {
    // Get workbook name if not provided
    const workbookName = await getWorkbookIdentifier();
    const workbookPath = await webViewBridge.getWorkbookPath();

    const message = {
      type: "stream:start",
      workbookName: workbookName,
      workbookPath: workbookPath || undefined,
      payload: {
        prompt: payload.prompt,
        activeTools: payload.activeTools,
        allowOtherWorkbookReads: payload.allowOtherWorkbookReads ?? false,
        documentsBase64: payload.documentsBase64,
      },
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async cancelCurrentRequest(): Promise<boolean> {
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "stream:cancel",
      workbookName: workbookName,
      payload: {},
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async restartState(): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "chat:restart",
      workbookName: workbookName,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async revertTool(toolUseId: string, userId?: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:revert",
      workbookName: workbookName,
      toolUseId: toolUseId,
      userId: userId,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async applyTool(toolUseId: string, userId?: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:apply",
      workbookName: workbookName,
      toolUseId: toolUseId,
      userId: userId,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async approveTool(toolId: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:approve",
      workbookName: workbookName,
      toolId: toolId,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async rejectTool(toolId: string, userMessage?: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:reject",
      workbookName: workbookName,
      toolId: toolId,
      userMessage: userMessage || undefined,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async viewTool(toolId: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:view",
      workbookName: workbookName,
      toolId: toolId,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async unviewTool(toolId: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "tool:unview",
      workbookName: workbookName,
      toolId: toolId,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async sendScoreOnly(score: number): Promise<boolean> {
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "score:score",
      workbookName: workbookName,
      score: score,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  async sendFeedback(comment: string): Promise<boolean> {
    // Get workbook name
    const workbookName = await getWorkbookIdentifier();

    const message = {
      type: "score:feedback",
      workbookName: workbookName,
      comment: comment,
    };

    await this.ensureConnectedAndSend(message);
    return true;
  }

  disconnect() {
    if (this.reconnectTimeout) {
      clearTimeout(this.reconnectTimeout);
      this.reconnectTimeout = null;
    }

    if (this.wsClient) {
      this.wsClient.close(1000, "Client disconnect"); // Proper close code
      this.wsClient = null;
    }

    this.isConnected = false;
    this.onConnectionChange?.(false);
  }

  getConnectionStatus(): boolean {
    return this.isConnected;
  }

  private scheduleReconnect(): void {
    if (this.reconnectTimeout || this.isConnecting) return;
    this.reconnectTimeout = window.setTimeout(async () => {
      this.reconnectTimeout = null;
      if (!this.isConnected) {
        try {
          await this.tryReconnect();
        } catch (error) {
          console.error("Reconnection attempt failed:", error);
        }
      }
    }, 1500);
  }

  // Public method to manually trigger a reconnection attempt
  async tryReconnect(): Promise<boolean> {
    if (this.isConnected) {
      console.log("Already connected, no need to reconnect");
      return true;
    }

    console.log("Manual reconnection attempt triggered...");

    // Clear any existing reconnection timeout and reset state
    if (this.reconnectTimeout) {
      clearTimeout(this.reconnectTimeout);
      this.reconnectTimeout = null;
    }

    // Close existing connection if any
    if (this.wsClient) {
      this.wsClient.close(1000, "Manual reconnection");
      this.wsClient = null;
    }
    this.isConnected = false;

    try {
      await this.connect();
      // Give a small delay to allow connection to establish
      await new Promise((resolve) => setTimeout(resolve, 500));
      return this.isConnected;
    } catch (error) {
      console.error("Manual reconnection failed:", error);
      return false;
    }
  }

  /**
   * Automatically send user email when WebSocket connection opens
   */
  private sendUserEmailOnConnect(): void {
    const userEmail = this.getCurrentUserEmail();
    this.sendUserEmail(
      userEmail,
      "Automatically sent user email on connection",
    );
  }

  /**
   * Handle server request for user email
   */
  private handleEmailRequest(): void {
    const userEmail = this.getCurrentUserEmail();
    this.sendUserEmail(userEmail, "Sent user email to server");
  }

  /**
   * Get current user email from various sources
   */
  private getCurrentUserEmail(): string | null {
    if (this.currentUserEmail) {
      return this.currentUserEmail;
    }

    // Try to get email from sessionStorage/localStorage first
    try {
      const storedAuth = localStorage.getItem(
        "sb-qvycnlwxhhmuobjzzoos-auth-token",
      );
      if (storedAuth) {
        const authData = JSON.parse(storedAuth);
        if (authData?.user?.email) {
          console.log("Found user email in localStorage:", authData.user.email);
          return authData.user.email;
        }
      }
    } catch (error) {
      console.warn("Could not get email from localStorage:", error);
    }

    // Try other storage keys
    try {
      const keys = Object.keys(localStorage);
      for (const key of keys) {
        if (key.includes("supabase") || key.includes("auth")) {
          const value = localStorage.getItem(key);
          if (value) {
            const data = JSON.parse(value);
            if (data?.user?.email) {
              console.log(`Found user email in ${key}:`, data.user.email);
              return data.user.email;
            }
          }
        }
      }
    } catch (error) {
      console.warn("Could not scan localStorage for email:", error);
    }

    console.log("No user email found in storage");
    return null;
  }

  private sendUserEmail(userEmail: string | null, logPrefix: string): void {
    const response = {
      type: "user:email_response",
      email: userEmail,
    };

    if (this.wsClient) {
      this.wsClient.send(response);
      console.log(`${logPrefix}:`, userEmail || "null");
    }
  }

  updateUserEmail(email: string | null): void {
    this.currentUserEmail = email;

    if (this.isConnected && this.wsClient) {
      this.sendUserEmail(email, "Updated user email sent to server");
    }
  }

  async sendToolPermissionResponse(
    groupId: string,
    toolResponses: Array<{ toolId: string; decision: "approved" | "rejected" }>,
  ): Promise<boolean> {
    try {
      const workbookName: string = await getWorkbookIdentifier();
      const message = {
        type: "tool:permission:response",
        workbookName: workbookName,
        groupId: groupId,
        toolResponses: toolResponses,
      };

      console.log("Sending tool permission response:", message);
      await this.ensureConnectedAndSend(message);
      return true;
    } catch (error) {
      console.error("Error sending tool permission response:", error);
      return false;
    }
  }

  /**
   * Register workbook with WebSocket as soon as connection opens
   */
  private registerWorkbookOnConnect(): void {
    // Get workbook name asynchronously and send registration
    getWorkbookIdentifier()
      .then((workbookName) => {
        if (!workbookName) {
          console.warn("No workbook name available for WebSocket registration");
          return;
        }

        const message = {
          type: "workbook:register",
          workbookName: workbookName,
        };

        if (this.wsClient) {
          this.wsClient.send(message);
          console.log("Registered workbook with WebSocket:", workbookName);
        }
      })
      .catch((error) => {
        console.error("Failed to register workbook on connect:", error);
      });
  }

  async sendToolResult(toolUseId: string, output: any): Promise<boolean> {
    const workbook = await getWorkbookIdentifier();
    const message = {
      type: "tool:result",
      workbookName: workbook,
      toolUseId,
      output,
    };
    await this.ensureConnectedAndSend(message);
    return true;
  }
}

async function getWorkbookIdentifier(): Promise<string> {
  if (pinnedWorkbookId) {
    return pinnedWorkbookId;
  }

  const now = Date.now();
  if (cachedWorkbookId && now - cachedWorkbookAt < WORKBOOK_CACHE_MS) {
    return cachedWorkbookId;
  }

  const path = await webViewBridge.getWorkbookPath();
  if (path) {
    cachedWorkbookId = path;
    cachedWorkbookAt = now;
    return path;
  }

  const name = await webViewBridge.getWorkbookName();
  cachedWorkbookId = name || "";
  cachedWorkbookAt = now;
  return cachedWorkbookId;
}

interface ToolResponse {
  toolId: string;
  decision: "approved" | "rejected";
  userMessage?: string;
}

// Singleton instance
export const webSocketChatService = new WebSocketChatService();
