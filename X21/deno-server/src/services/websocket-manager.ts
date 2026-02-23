import { createLogger } from "../utils/logger.ts";
import type { OperationStatus } from "../types/index.ts";
import {
  OperationStatusValues,
  WebSocketMessageTypes,
} from "../types/index.ts";
import process from "node:process";

const logger = createLogger("WebSocketManager");

export class WebSocketManager {
  private static instance: WebSocketManager;
  // private socket: WebSocket | null = null;
  private sockets: Map<string, WebSocket | null> = new Map();
  private deltaState: Map<string, { index?: number; deltaType?: string }> =
    new Map();

  static getInstance(): WebSocketManager {
    if (!this.instance) this.instance = new WebSocketManager();
    return this.instance;
  }

  getSocket(workbookName: string): WebSocket {
    const socket = this.sockets.get(workbookName);
    if (socket) {
      return socket as WebSocket;
    } else {
      throw new Error("No socket found");
    }
  }

  setSocket(workbookName: string, socket: WebSocket) {
    const existingSocket = this.sockets.get(workbookName);

    if (existingSocket && existingSocket !== socket) {
      logger.info(`Replacing old socket for workbook: ${workbookName}`);
      try {
        existingSocket.close(1000, "Replaced by new connection");
      } catch (error) {
        logger.warn(`Failed to close old socket for ${workbookName}:`, error);
      }
    }

    // Set the new socket
    this.sockets.set(workbookName, socket);

    // Set up cleanup when this socket closes
    const cleanup = () => {
      // Only remove if this is still the current socket for this workbook
      if (this.sockets.get(workbookName) === socket) {
        this.sockets.delete(workbookName);
        // Clean up delta state
        if (this.deltaState.has(workbookName)) {
          process.stdout.write("\n");
          this.deltaState.delete(workbookName);
        }
        logger.info(`Cleaned up socket for workbook: ${workbookName}`);
      }
    };

    socket.addEventListener("close", cleanup);
    socket.addEventListener("error", cleanup);
  }

  private writeDeltaContent(workbookName: string, payload: any): void {
    if (!payload) return;

    const state = this.deltaState.get(workbookName) || {};
    const currentIndex = payload.index;
    const currentDeltaType = payload.delta?.type;

    // Check if we need to start a new line (context changed)
    const contextChanged = state.index !== currentIndex ||
      state.deltaType !== currentDeltaType ||
      payload.type !== "content_block_delta";

    // Handle content_block_delta with various delta types
    if (payload.type === "content_block_delta" && payload.delta) {
      const delta = payload.delta;

      // Start new line if context changed
      if (contextChanged) {
        if (state.index !== undefined || state.deltaType !== undefined) {
          process.stdout.write("\n");
        }
        const index = currentIndex !== undefined ? `[${currentIndex}]` : "";
        process.stdout.write(`[${workbookName}] delta${index} ${delta.type}: `);
        this.deltaState.set(workbookName, {
          index: currentIndex,
          deltaType: currentDeltaType,
        });
      }

      // Write just the content
      if (delta.type === "input_json_delta" && delta.partial_json) {
        process.stdout.write(delta.partial_json);
      } else if (delta.type === "text_delta" && delta.text) {
        process.stdout.write(delta.text);
      } else if (delta.type === "thinking_delta" && delta.thinking) {
        process.stdout.write(delta.thinking);
      } else if (delta.type === "tool_use_delta") {
        if (delta.name) process.stdout.write(delta.name);
        if (delta.input) process.stdout.write(delta.input);
      }
      return;
    }

    // For other event types, start a new line
    if (state.index !== undefined || state.deltaType !== undefined) {
      process.stdout.write("\n");
      this.deltaState.delete(workbookName);
    }

    // Handle message_delta
    if (payload.type === "message_delta") {
      const parts: string[] = [`[${workbookName}] message_delta`];
      if (payload.delta?.stop_reason) {
        parts.push(`stop: ${payload.delta.stop_reason}`);
      }
      if (payload.usage) {
        parts.push(`tokens: ${payload.usage.output_tokens || 0}`);
      }
      process.stdout.write(parts.join(" ") + "\n");
      return;
    }

    // Handle content_block_start
    if (payload.type === "content_block_start") {
      const blockType = payload.content_block?.type || "unknown";
      const index = payload.index !== undefined ? ` [${payload.index}]` : "";
      process.stdout.write(`[${workbookName}] start ${blockType}${index}\n`);
      return;
    }

    // Handle content_block_stop
    if (payload.type === "content_block_stop") {
      const index = payload.index !== undefined ? ` [${payload.index}]` : "";
      process.stdout.write(`[${workbookName}] stop${index}\n`);
      return;
    }
  }

  send(workbookName: string, type: string, payload: any): boolean {
    const socket = this.sockets.get(workbookName);
    if (!socket || socket.readyState !== socket.OPEN) {
      logger.warn(
        `Socket not found or not open for workbook: ${workbookName}`,
        {
          available: Array.from(this.sockets.keys()),
        },
      );
      return false;
    }
    try {
      const data: any = { type: type, payload: payload };
      socket.send(JSON.stringify(data));

      // Enhanced logging for status:update messages
      if (type === "status:update") {
        logger.info(`Message sent to ${workbookName}`, {
          type,
          status: payload?.status,
          message: payload?.message,
          progress: payload?.progress,
          metadata: payload?.metadata,
        });
      } else if (type === "stream:delta") {
        // Write deltas continuously on the same line
        this.writeDeltaContent(workbookName, payload);
      } else {
        logger.info(`Message sent to ${workbookName}`, { type, payload });
      }
      return true;
    } catch (error) {
      logger.error(`Failed to send message to ${workbookName}`, error);
      return false;
    }
  }

  isConnected(workbookName: string): boolean {
    const socket = this.sockets.get(workbookName);

    return !!socket && socket.readyState === socket.OPEN;
  }

  close(workbookName: string, code?: number, reason?: string) {
    const socket = this.sockets.get(workbookName);
    try {
      socket?.close(code, reason);
    } finally {
      this.sockets.delete(workbookName);
    }
  }

  // Helper method to send status updates
  sendStatus(
    workbookName: string,
    status: OperationStatus,
    message?: string,
    progress?: {
      current: number;
      total: number;
      unit?: string;
    },
    metadata?: {
      operation?: string;
      range?: string;
      toolName?: string;
      estimatedMs?: number;
    },
  ): boolean {
    const payload = {
      status,
      message,
      progress,
      metadata,
    };

    const sent = this.send(workbookName, "status:update", payload);

    if (!sent) {
      logger.warn(
        `Failed to send status:update (socket missing or not open) for workbook=${workbookName}`,
        { status, message, hasProgress: !!progress, metadata },
      );
    }
    // Note: Success logging is handled by send() method for status:update messages

    return sent;
  }

  // Helper method to send token updates
  sendTokenUpdate(
    workbookName: string,
    inputTokens: number,
    outputTokens: number,
    totalTokens: number,
  ): boolean {
    return this.send(workbookName, "stream:token_update", {
      inputTokens,
      outputTokens,
      totalTokens,
    });
  }

  /**
   * Safely end a stream by sending stream:end followed by idle status.
   * This ensures correct message ordering to avoid race conditions on the frontend.
   *
   * Message order matters:
   * 1. stream:end → clears currentAssistantMessageRef on frontend
   * 2. idle status → can be safely accepted (ref is null)
   *
   * @param workbookName - The workbook identifier
   * @param usage - LLM usage information (tokens, can be any object)
   * @param model - Model identifier (e.g., "claude-3-5-sonnet-20241022")
   */
  endStream(
    workbookName: string,
    usage?: Record<string, any>,
    model?: string | null,
  ): void {
    // Send stream:end first
    this.send(workbookName, "stream:end", {
      usage: usage || {},
      model: model || null,
    });

    // Clean up delta state and add newline if there was an active delta
    if (this.deltaState.has(workbookName)) {
      process.stdout.write("\n");
      this.deltaState.delete(workbookName);
    }

    // Then send idle status
    this.sendStatus(workbookName, OperationStatusValues.IDLE);

    logger.info(`Stream ended for workbook: ${workbookName}`);
  }

  /**
   * Safely cancel a stream by sending stream:cancelled followed by idle status.
   * This ensures correct message ordering to avoid race conditions on the frontend.
   *
   * Message order matters:
   * 1. stream:cancelled → clears currentAssistantMessageRef on frontend
   * 2. idle status → can be safely accepted (ref is null)
   *
   * @param workbookName - The workbook identifier
   * @param message - Cancellation message
   * @param requestId - Request identifier for tracing
   */
  cancelStream(
    workbookName: string,
    message?: string,
    requestId?: string,
  ): void {
    // Send stream:cancelled first
    this.send(workbookName, WebSocketMessageTypes.STREAM_CANCELLED, {
      message: message || "Request was cancelled by user",
      requestId,
    });

    // Clean up delta state and add newline if there was an active delta
    if (this.deltaState.has(workbookName)) {
      process.stdout.write("\n");
      this.deltaState.delete(workbookName);
    }

    // Then send idle status
    this.sendStatus(workbookName, OperationStatusValues.IDLE);

    logger.info(`Stream cancelled for workbook: ${workbookName}`);
  }

  /**
   * Safely end a stream with an error by sending stream:error followed by idle status.
   * This ensures correct message ordering to avoid race conditions on the frontend.
   *
   * Message order matters:
   * 1. stream:error → clears currentAssistantMessageRef on frontend
   * 2. idle status → can be safely accepted (ref is null)
   *
   * @param workbookName - The workbook identifier
   * @param payload - Error payload to send
   */
  errorStream(
    workbookName: string,
    payload: any,
  ): void {
    // Send stream:error first
    this.send(workbookName, "stream:error", payload);

    // Clean up delta state and add newline if there was an active delta
    if (this.deltaState.has(workbookName)) {
      process.stdout.write("\n");
      this.deltaState.delete(workbookName);
    }

    // Then send idle status
    this.sendStatus(workbookName, OperationStatusValues.IDLE);

    logger.info(`Stream error sent for workbook: ${workbookName}`);
  }
}
