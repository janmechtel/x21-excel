import Anthropic from "@anthropic-ai/sdk";
import { createLogger } from "../utils/logger.ts";
import { insertMessage, updateToolResultContent } from "../db/dal.ts";
import { ClaudeContentTypes } from "../types/index.ts";

const logger = createLogger("StateManager");

class StateManager {
  private workbookStates = new Map<string, State>();

  constructor() {
    logger.info("StateManager initialized");
  }

  deleteWorkbookState(workbookName: string): void {
    if (!this.workbookStates.has(workbookName)) {
      logger.info("StateManager: Workbook state not found for deletion", {
        workbookName,
      });
      return;
    }
    logger.info("StateManager: DELETING workbook state (including snapshot!)", {
      workbookName,
    });
    this.workbookStates.delete(workbookName);
  }

  addMessage(
    workbookName: string,
    message: Anthropic.MessageParam,
    options?: { persist?: boolean },
  ): Anthropic.MessageParam[] {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.conversationHistory.push(message);

    if (options?.persist !== false) {
      // Persist to SQLite (best-effort)
      try {
        const conversationId = workbookState.sessionId;
        const workbookKey = workbookState.workbookKey ?? workbookName;
        const content = typeof message.content === "string"
          ? message.content
          : JSON.stringify(message.content);

        insertMessage({
          workbookKey,
          conversationId,
          role: message.role as string,
          content,
        });
      } catch (e) {
        logger.warn("Failed to persist message to SQLite", {
          workbookName,
          error: (e as Error)?.message,
        });
      }
    }

    return workbookState.conversationHistory;
  }

  setConversationHistory(
    workbookName: string,
    conversationHistory: Anthropic.MessageParam[],
  ): void {
    const workbookState = this.workbookStates.get(workbookName);

    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    workbookState.conversationHistory = conversationHistory;
  }

  clearConversationHistory(workbookName: string): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.conversationHistory = [];
  }

  clearToolChangeHistory(workbookName: string): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.toolChangeHistory = new Map();
  }

  getConversationHistory(workbookName: string): any[] {
    const workbookState = this.getState(workbookName);
    return workbookState.conversationHistory;
  }

  isToolChangeApproved(workbookName: string, toolId: string): boolean {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    return workbookState.toolChangeHistory.get(toolId)?.approved ?? false;
  }

  isToolChangeApplied(workbookName: string, toolId: string): boolean {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    return workbookState.toolChangeHistory.get(toolId)?.applied ?? false;
  }

  getToolChange(workbookName: string, toolId: string): ToolChangeInterface {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      logger.error("Workbook state not found for ", { workbookName });
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    const toolChange = workbookState.toolChangeHistory.get(toolId);

    if (!toolChange) {
      logger.error("Tool change not found for ", { workbookName, toolId });
      throw new Error(`Tool change ${toolId} not found`);
    }

    return toolChange;
  }

  updateToolChange(
    workbookName: string,
    toolId: string,
    change: Partial<ToolChangeInterface>,
  ): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    const existingToolChange = workbookState.toolChangeHistory.get(toolId);
    if (!existingToolChange) {
      throw new Error(`Tool change ${toolId} not found`);
    }

    const newToolChange = {
      ...existingToolChange,
      ...change,
    };
    workbookState.toolChangeHistory.set(toolId, newToolChange);
  }

  addToolChange(
    workbookName: string,
    toolId: string,
    change: ToolChangeInterface,
  ): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.toolChangeHistory.set(toolId, change);
  }

  addInitialToolChange(
    workbookName: string,
    tool: Anthropic.ToolUseBlock,
    requestId: string,
  ): void {
    const toolId = tool.id;

    const change: ToolChangeInterface = {
      requestId: requestId,
      toolId: toolId,
      toolName: tool.name,
      workbookName: workbookName,
      timestamp: new Date(),
      inputData: tool.input,
      applied: false,
      approved: false,
      pending: true,
    };
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.toolChangeHistory.set(toolId, change);
  }

  anyPendingToolChanges(workbookName: string): boolean {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    for (const change of workbookState.toolChangeHistory.values()) {
      if (change.pending) return true;
    }

    return false;
  }

  getSessionId(workbookName: string): string {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    return workbookState.sessionId ?? "RND";
  }

  setLatestRequestId(workbookName: string): string {
    const requestUUID = globalThis.crypto.randomUUID();
    const workbookState = this.getState(workbookName);
    workbookState.latestRequestId = requestUUID;
    return requestUUID;
  }

  getWorkbookName(requestId: string): string {
    logger.info("All states", { workbookStates: this.workbookStates.size });
    logger.info("Request ID", { requestId });
    for (const [workbookName, workbookState] of this.workbookStates.entries()) {
      if (workbookState.latestRequestId === requestId) {
        return workbookName;
      }
    }
    throw new Error(`Workbook state not found for ${requestId}`);
  }

  getLatestRequestId(workbookName: string): string {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    return workbookState.latestRequestId;
  }

  getToolChanges(workbookName: string): ToolChangeInterface[] {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    return Array.from(workbookState.toolChangeHistory.values());
  }

  startState(workbookName: string): void {
    this.workbookStates.set(workbookName, {
      sessionId: globalThis.crypto.randomUUID(),
      latestRequestId: globalThis.crypto.randomUUID(),
      conversationHistory: [],
      toolChangeHistory: new Map(),
      abortController: new Map(),
      requestMetada: new Map(),
      workbookKey: workbookName,
      lastSnapshot: undefined,
    });
  }

  setWorkbookKey(workbookName: string, workbookKey: string): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.workbookKey = workbookKey;
  }

  getState(workbookName: string): State {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    return workbookState;
  }

  getOrCreateState(workbookName: string): State {
    try {
      logger.info("Getting state", { workbookName });
      return this.getState(workbookName);
    } catch (_) {
      logger.info("Starting state", { workbookName });
      this.startState(workbookName);
      return this.getState(workbookName);
    }
  }

  saveRequestMetadata(
    workbookName: string,
    requestId: string,
    metadata: RequestMetadata,
  ): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.requestMetada.set(requestId, metadata);
  }

  getRequestMetadata(workbookName: string, requestId: string): RequestMetadata {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    const requestMetadata = workbookState.requestMetada.get(requestId);
    if (!requestMetadata) {
      throw new Error(`Request metadata not found for ${requestId}`);
    }
    return requestMetadata;
  }

  deleteRequestMetadata(workbookName: string, requestId: string): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.requestMetada.delete(requestId);
  }

  creatingAbortController(workbookName: string, requestId: string): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }
    workbookState.abortController.set(requestId, new AbortController());
  }

  getAbortController(
    workbookName: string,
    requestId: string,
  ): AbortController | null {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      return null;
    }
    const abortController = workbookState.abortController.get(requestId);
    return abortController || null;
  }

  getAbortControllerSignal(
    workbookName: string,
    requestId: string,
  ): AbortSignal | undefined {
    const abortController = this.getAbortController(workbookName, requestId);
    return abortController?.signal;
  }

  abortAbortController(workbookName: string, requestId: string): void {
    const abortController = this.getAbortController(workbookName, requestId);
    if (abortController) {
      abortController.abort();
      this.workbookStates.get(workbookName)?.abortController.delete(requestId);
    }
  }

  updateToolResponseInConversationHistory(
    workbookName: string,
    result: any,
    toolId: string,
  ) {
    const workbookState = this.getState(workbookName);
    const conversationId = workbookState.sessionId;
    const workbookKey = workbookState.workbookKey ?? workbookName;
    let updatedContentForDb: any = null;

    const conversationHistory = this.getConversationHistory(workbookName);

    // Find and update the user message that matches the toolId
    const updatedHistory = conversationHistory.map((message: any) => {
      if (message.role === "user" && Array.isArray(message.content)) {
        const hasMatchingToolResult = message.content.some((content: any) =>
          content.type === ClaudeContentTypes.TOOL_RESULT &&
          content.tool_use_id === toolId
        );

        if (hasMatchingToolResult) {
          const updatedContent = message.content.map((content: any) => {
            if (
              content.type === ClaudeContentTypes.TOOL_RESULT &&
              content.tool_use_id === toolId
            ) {
              return {
                ...content,
                content: JSON.stringify(result),
              };
            }
            return content;
          });

          updatedContentForDb = updatedContent;

          // Update the content with the new result
          return {
            ...message,
            content: updatedContent,
          };
        }
      }
      return message;
    });

    // Update the conversation history
    this.setConversationHistory(workbookName, updatedHistory);

    // Persist the updated tool result to SQLite (best-effort)
    if (updatedContentForDb) {
      try {
        const serializedContent = typeof updatedContentForDb === "string"
          ? updatedContentForDb
          : JSON.stringify(updatedContentForDb);
        const changes = updateToolResultContent(
          workbookKey,
          conversationId,
          toolId,
          serializedContent,
        );
        if (changes === 0) {
          logger.warn("No SQLite rows updated for tool result", {
            workbookName,
            toolId,
          });
        }
      } catch (error) {
        logger.warn("Failed to persist tool result to SQLite", {
          workbookName,
          toolId,
          error: (error as Error)?.message,
        });
      }
    }
  }

  setLastSnapshot(workbookName: string, snapshot: any): void {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    const sheetCount = snapshot?.sheetXmls
      ? Object.keys(snapshot.sheetXmls).length
      : 0;
    logger.info("StateManager: Setting last snapshot", {
      workbookName,
      sheetCount,
      hasWorkbookXml: !!snapshot?.workbookXml,
      hasSharedStrings: !!snapshot?.sharedStringsXml,
    });

    workbookState.lastSnapshot = snapshot;
  }

  getLastSnapshot(workbookName: string): any | undefined {
    const workbookState = this.workbookStates.get(workbookName);
    if (!workbookState) {
      throw new Error(`Workbook state not found for ${workbookName}`);
    }

    const snapshot = workbookState.lastSnapshot;
    const sheetCount = snapshot?.sheetXmls
      ? Object.keys(snapshot.sheetXmls).length
      : 0;

    logger.info("StateManager: Getting last snapshot", {
      workbookName,
      hasSnapshot: !!snapshot,
      sheetCount,
      hasWorkbookXml: !!snapshot?.workbookXml,
      hasSharedStrings: !!snapshot?.sharedStringsXml,
    });

    return snapshot;
  }
}

export interface State {
  sessionId: string;
  latestRequestId: string; // traceId
  conversationHistory: Anthropic.MessageParam[];
  toolChangeHistory: Map<string, ToolChangeInterface>; //Map<toolId, ToolChangeInterface>
  abortController: Map<string, AbortController>; //Map<requestId, AbortController>
  requestMetada: Map<string, RequestMetadata>; //Map<requestId, RequestMetadata>
  workbookKey?: string;
  lastSnapshot?: any; // Store last workbook snapshot for diff calculation
}

export interface RequestMetadata {
  activeTools: string[];
  workbookName: string;
  worksheets: string[];
  activeWorksheet: string;
}

// Interface for storing tool change data
export interface ToolChangeInterface {
  timestamp: Date;
  workbookName: string;
  worksheet?: string;
  requestId: string;
  toolId: string;
  toolName: string;
  applied: boolean;
  approved: boolean;
  pending: boolean;
  inputData?: any;
  inputDataRevert?: any;
  result?: any;
}

export const stateManager = new StateManager();

// Setter to store a stable file identifier (from client) on the workbook state
