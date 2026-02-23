import Anthropic from "@anthropic-ai/sdk";
import { streamClaudeResponseToWebSocket } from "./llm.ts";
import { streamNativeOpenAIResponseToWebSocket } from "./openai-llm.ts";
import { stateManager, ToolChangeInterface } from "../state/state-manager.ts";
import { WebSocketManager } from "../services/websocket-manager.ts";
import { tracing } from "../tracing/tracing.ts";
import { getToolByName } from "../tools/index.ts";
import { createLogger } from "../utils/logger.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";
import { resolveWorkbookForWrite } from "../utils/workbook-resolver.ts";
import { sanitizeWorkbookName } from "../utils/workbook-name.ts";
import {
  AddColumnsRequest,
  AddRowsRequest,
  CellValue,
  ClaudeContentTypes,
  ClaudeStopReasons,
  DeleteCellsRequest,
  OperationStatus,
  OperationStatusValues,
  RangeFormatPair,
  ReadFormatFinalResponseList,
  RemoveColumnsRequest,
  RemoveSheetsRequest,
  RevertOperationKeys,
  ToolNames,
  WebSocketMessageTypes,
  WorkbookResolutionPaths,
  WriteValuesRequest,
} from "../types/index.ts";
import { UiRequestPayload } from "../types/ui-request.ts";
import { ToolExecutionError } from "../errors/tool-execution-error.ts";
import { UserCancellationError } from "../errors/user-cancellation-error.ts";
import { getLLMProvider } from "../llm-client/provider.ts";

const socket = WebSocketManager.getInstance();
const logger = createLogger("ToolLogic");
const excelReadyChecks = new Map<string, Promise<void>>();
const WRITE_TOOLS = new Set<string>([
  ToolNames.COPY_PASTE,
  ToolNames.WRITE_VALUES_BATCH,
  ToolNames.WRITE_FORMAT_BATCH,
  ToolNames.DRAG_FORMULA,
  ToolNames.ADD_COLUMNS,
  ToolNames.REMOVE_COLUMNS,
  ToolNames.ADD_ROWS,
  ToolNames.REMOVE_ROWS,
  ToolNames.ADD_SHEETS,
]);
const EXCEL_READY_SKIP_TOOLS = new Set<string>([
  ToolNames.WORKBOOK_CHANGELOG,
]);

async function waitForExcelReadyBeforeApproval(
  workbookName: string,
  toolName: string,
): Promise<void> {
  if (EXCEL_READY_SKIP_TOOLS.has(toolName)) {
    return;
  }
  if (!workbookName || workbookName.trim() === "") {
    return;
  }

  const existing = excelReadyChecks.get(workbookName);
  if (existing) {
    await existing;
    return;
  }

  const client = ExcelApiClient.getInstance();
  const waitPromise = (async () => {
    try {
      await client.getMetadata({ workbookName });
    } catch (error) {
      logger.warn("Excel readiness check failed; continuing to approval", {
        workbookName,
        toolName,
        error: error instanceof Error ? error.message : String(error),
      });
    }
  })();

  excelReadyChecks.set(workbookName, waitPromise);
  try {
    await waitPromise;
  } finally {
    excelReadyChecks.delete(workbookName);
  }
}

export async function streamClaudeResponseAndHandleToolUsage(
  requestId: string,
  client: Anthropic,
  params: Anthropic.MessageCreateParamsNonStreaming,
): Promise<Anthropic.Message> {
  const provider = getLLMProvider();
  logger.info("Getting abort controller");
  const workbookName = stateManager.getWorkbookName(requestId);

  const abortController = stateManager.getAbortController(
    workbookName,
    requestId,
  );

  if (!abortController) {
    logger.info("Abort controller not found, request was likely cancelled", {
      workbookName,
      requestId,
    });
    throw new UserCancellationError("Request was cancelled", requestId);
  }

  try {
    const toolIdsHandledDuringStream = new Set<string>();

    const handleToolUseBlock = (tool: Anthropic.ToolUseBlock) => {
      if (!tool?.id || !tool?.name) return;
      if (toolIdsHandledDuringStream.has(tool.id)) return;
      toolIdsHandledDuringStream.add(tool.id);

      // Add initial tool change (idempotent-ish: skip if already exists)
      try {
        stateManager.getToolChange(workbookName, tool.id);
      } catch {
        stateManager.addInitialToolChange(workbookName, tool, requestId);
      }

      // UI-request tools are handled immediately (form rendering)
      if (tool.name === ToolNames.COLLECT_INPUT) {
        socket.send(workbookName, WebSocketMessageTypes.UI_REQUEST, {
          toolUseId: tool.id,
          payload: tool.input as UiRequestPayload,
        });
        socket.sendStatus(
          workbookName,
          OperationStatusValues.WAITING_APPROVAL,
          "Waiting for your answer...",
        );
        return;
      }

      // Send permission request for executable tools once Excel is ready.
      void (async () => {
        await waitForExcelReadyBeforeApproval(workbookName, tool.name);
        socket.send(workbookName, WebSocketMessageTypes.TOOL_PERMISSION, {
          toolPermissions: [{ toolId: tool.id, toolName: tool.name }],
        });
        socket.sendStatus(
          workbookName,
          OperationStatusValues.WAITING_APPROVAL,
          "Waiting for tool approval...",
        );
      })().catch((error) => {
        logger.warn("Failed to send tool approval request", {
          workbookName,
          toolId: tool.id,
          toolName: tool.name,
          error: error instanceof Error ? error.message : String(error),
        });
      });
    };

    const finalMessage = provider === "anthropic"
      ? await streamClaudeResponseToWebSocket(
        requestId,
        client,
        params,
        abortController,
      )
      : await streamNativeOpenAIResponseToWebSocket(
        requestId,
        params,
        abortController,
      );

    stateManager.addMessage(workbookName, {
      role: "assistant",
      content: finalMessage.content,
    });

    if (finalMessage.stop_reason === ClaudeStopReasons.TOOL_USE) {
      // Find ALL tools in the content, not just the first one
      const tools = finalMessage.content.filter(
        (content: any): content is Anthropic.ToolUseBlock =>
          content.type === ClaudeContentTypes.TOOL_USE,
      );

      if (tools.length === 0) {
        throw new Error("No tools found in message");
      }

      logger.info("Detected tool_use blocks", {
        count: tools.length,
        toolNames: tools.map((tool) => tool.name),
      });
      logger.info("tools", { tools });

      const uiRequestTools = tools.filter((tool) =>
        tool.name === ToolNames.COLLECT_INPUT
      );
      const executableTools = tools.filter((tool) =>
        tool.name !== ToolNames.COLLECT_INPUT
      );

      // Create tool permission payloads for executable tools
      const toolPermissionPayloads: ToolPermissionPayload = {
        toolPermissions: executableTools.map((tool) => ({
          toolId: tool.id,
          toolName: tool.name,
        })),
      };

      // Add initial tool changes for all tools
      tools.forEach((tool) => {
        stateManager.addInitialToolChange(workbookName, tool, requestId);
      });

      // Create placeholder responses for all tools
      const placeholderResponse: Anthropic.MessageParam = {
        role: "user",
        content: tools.map((tool) => ({
          type: ClaudeContentTypes.TOOL_RESULT,
          tool_use_id: tool.id,
          content: "Placeholder - Response Not Received",
        })),
      };

      stateManager.addMessage(workbookName, placeholderResponse);

      // Send UI request payloads to the client for form rendering
      if (uiRequestTools.length > 0) {
        uiRequestTools.forEach((tool) => {
          const payload = tool.input as UiRequestPayload;

          socket.send(workbookName, WebSocketMessageTypes.UI_REQUEST, {
            toolUseId: tool.id,
            payload,
          });
        });

        socket.sendStatus(
          workbookName,
          OperationStatusValues.WAITING_APPROVAL,
          "Waiting for your answer...",
        );
      }

      // Send executable tool permissions via socket
      if (toolPermissionPayloads.toolPermissions.length > 0) {
        socket.send(
          workbookName,
          WebSocketMessageTypes.TOOL_PERMISSION,
          toolPermissionPayloads,
        );

        // Send waiting_approval status for executable tools (centralized status management)
        if (uiRequestTools.length === 0) {
          socket.sendStatus(
            workbookName,
            OperationStatusValues.WAITING_APPROVAL,
            "Waiting for tool approval...",
          );
        }
      }
    } else {
      const metadata = {
        endTime: new Date().toISOString(),
      };
      tracing.endTrace(requestId, metadata);

      socket.endStream(
        workbookName,
        finalMessage.usage || {},
        finalMessage.model || null,
      );

      stateManager.deleteRequestMetadata(workbookName, requestId);
    }

    return finalMessage;
  } catch (error: any) {
    const metadata = {
      endTime: new Date().toISOString(),
      output: {
        success: false,
        error: error.message,
        stack: error.stack,
      },
    };
    tracing.endTrace(requestId, metadata);
    throw error;
  }
}

export async function executeTool(
  toolName: string,
  inputData: any,
): Promise<string> {
  logger.info("getting tool: ", { toolName });
  const tool = getToolByName(toolName);

  if (!tool) {
    throw new Error(`Tool ${toolName} not found`);
  }

  logger.info("executing tool: ", { toolName });
  const result = await tool.execute(inputData);
  return result;
}

export async function toolExecutionFlow(
  requestId: string,
  toolName: string,
  inputData: any,
  toolChange: ToolChangeInterface,
  _workbookName?: string,
): Promise<any> {
  logger.info("starting tool span");
  const spanId = tracing.logToolCallStart(requestId, toolName, inputData);
  const sessionWorkbookNameRaw = _workbookName ||
    stateManager.getWorkbookName(requestId);
  const sessionWorkbookName = sanitizeWorkbookName(sessionWorkbookNameRaw) ||
    sessionWorkbookNameRaw;
  const providedWorkbookRaw = inputData?.workbookName ||
    inputData?.workbook ||
    inputData?.targetWorkbook;
  const providedWorkbook = sanitizeWorkbookName(providedWorkbookRaw) ||
    providedWorkbookRaw;

  let resolvedWorkbookName = sessionWorkbookName;
  let resolutionPath = WorkbookResolutionPaths.SESSION;
  let openWorkbooks: string[] | undefined;

  try {
    if (isWriteTool(toolName)) {
      const resolution = await resolveWorkbookForWrite({
        sessionWorkbookName,
        providedWorkbookName: providedWorkbook,
      });
      resolvedWorkbookName = resolution.workbookName;
      resolutionPath = resolution.resolutionPath;
      openWorkbooks = resolution.openWorkbooks;
    }

    // Send status update before tool execution
    const preStatus = buildToolStatus(toolName, inputData);
    socket.sendStatus(
      sessionWorkbookName,
      preStatus.status,
      preStatus.message,
      preStatus.progress,
      preStatus.metadata,
    );

    logger.info("executing tool", {
      toolName,
      sessionWorkbookName,
      resolvedWorkbookName,
      resolutionPath,
      providedWorkbook,
    });
    const normalizedInput = normalizeToolInput(toolName, inputData);
    const preparedInput = applyWorkbookGuards(
      toolName,
      normalizedInput,
      resolvedWorkbookName,
      providedWorkbook,
    );
    if (toolChange) {
      toolChange.inputData = preparedInput;
    }

    const result = await executeTool(toolName, preparedInput);

    // Send status update after tool execution - back to generating
    socket.sendStatus(
      sessionWorkbookName,
      OperationStatusValues.GENERATING_LLM,
      "Analyzing results...",
    );

    logger.info("ending tool span");
    tracing.logToolCallEnd(spanId, result);
    return result;
  } catch (error: any) {
    logger.error("Error executing tool", {
      error,
      toolName,
      sessionWorkbookName,
      resolvedWorkbookName,
      resolutionPath,
      providedWorkbook,
      openWorkbooks,
    });
    tracing.logToolCallEnd(spanId, error);

    socket.sendStatus(
      sessionWorkbookName,
      OperationStatusValues.ERROR,
      `Tool execution failed: ${error?.message || "Unknown error"}`,
    );

    throw new ToolExecutionError(
      `Tool execution failed: ${error?.message || "Unknown error"}`,
      toolChange.toolId,
      toolName,
      error,
    );
  }
}

// Helper function to get user-friendly tool descriptions
function getToolDescription(toolName: string): string {
  const descriptions: Record<string, string> = {
    [ToolNames.COPY_PASTE]: "Copying and pasting cells",
    [ToolNames.READ_VALUES_BATCH]: "Reading cell values from Excel (batch)",
    [ToolNames.WRITE_VALUES_BATCH]: "Writing values to Excel (batch)",
    [ToolNames.READ_FORMAT_BATCH]: "Reading cell formatting (batch)",
    [ToolNames.WRITE_FORMAT_BATCH]: "Applying cell formatting",
    [ToolNames.DRAG_FORMULA]: "Dragging formula across cells",
    [ToolNames.ADD_SHEETS]: "Adding new sheets",
    [ToolNames.REMOVE_SHEETS]: "Removing sheets",
    [ToolNames.ADD_COLUMNS]: "Adding columns",
    [ToolNames.REMOVE_COLUMNS]: "Removing columns",
    [ToolNames.ADD_ROWS]: "Adding rows",
    [ToolNames.REMOVE_ROWS]: "Removing rows",
    [ToolNames.VBA_CREATE]: "Creating VBA macro",
    [ToolNames.VBA_READ]: "Reading VBA code",
    [ToolNames.VBA_UPDATE]: "Updating VBA macro",
    [ToolNames.COLLECT_INPUT]: "Waiting for user input",
  };
  return descriptions[toolName] || `Executing ${toolName}`;
}

type ToolStatus = {
  status: OperationStatus;
  message: string;
  progress?: {
    current: number;
    total: number;
    unit?: string;
  };
  metadata?: {
    operation?: string;
    range?: string;
    toolName?: string;
    estimatedMs?: number;
  };
};

function buildToolStatus(toolName: string, inputData: any): ToolStatus {
  const friendlyName = toolName.replace(/_/g, " ");
  const baseMetadata = {
    toolName,
    operation: getToolDescription(toolName),
    range: inputData?.range,
  };

  if (toolName === ToolNames.WRITE_FORMAT_BATCH) {
    const operations = Array.isArray(inputData?.operations)
      ? inputData.operations
      : [];
    const opCount = operations.length;
    const firstOp = operations[0];
    const firstLocation = firstOp?.worksheet && firstOp?.range
      ? `${firstOp.worksheet}!${firstOp.range}`
      : firstOp?.range || firstOp?.worksheet;

    return {
      status: OperationStatusValues.WRITING_EXCEL_FORMAT,
      message: opCount > 0
        ? `Applying batch formatting (${opCount} ops)${
          firstLocation ? ` starting at ${firstLocation}` : ""
        }`
        : "Applying batch formatting...",
      progress: opCount > 0
        ? { current: 0, total: opCount, unit: "ops" }
        : undefined,
      metadata: {
        ...baseMetadata,
        range: firstOp?.range || baseMetadata.range,
      },
    };
  }

  if (toolName === ToolNames.WRITE_VALUES_BATCH) {
    const operations = Array.isArray(inputData?.operations)
      ? inputData.operations
      : [];
    const opCount = operations.length;
    const firstOp = operations[0];
    const firstLocation = firstOp?.worksheet && firstOp?.range
      ? `${firstOp.worksheet}!${firstOp.range}`
      : firstOp?.range || firstOp?.worksheet;

    return {
      status: OperationStatusValues.WRITING_EXCEL,
      message: opCount > 0
        ? `Writing values (${opCount} ops)${
          firstLocation ? ` starting at ${firstLocation}` : ""
        }`
        : "Writing values (batch)...",
      progress: opCount > 0
        ? { current: 0, total: opCount, unit: "ops" }
        : undefined,
      metadata: {
        ...baseMetadata,
        range: firstOp?.range || baseMetadata.range,
      },
    };
  }

  if (toolName === ToolNames.COPY_PASTE) {
    const source = inputData?.sourceRange;
    const destinationWorksheet = inputData?.destinationWorksheet ||
      inputData?.worksheet ||
      inputData?.sourceWorksheet;
    const destinationRange = inputData?.destinationRange;
    const destinationLocation = destinationWorksheet && destinationRange
      ? `${destinationWorksheet}!${destinationRange}`
      : destinationRange || destinationWorksheet;

    return {
      status: "writing_excel",
      message: destinationLocation
        ? `Copying ${source || "range"} to ${destinationLocation}`
        : "Copy/pasting in Excel...",
      metadata: {
        ...baseMetadata,
        range: destinationRange || baseMetadata.range,
      },
    };
  }

  if (toolName === ToolNames.READ_VALUES_BATCH) {
    const operations = Array.isArray(inputData?.operations)
      ? inputData.operations
      : [];
    const opCount = operations.length;
    const firstOp = operations[0];
    const firstLocation = firstOp?.worksheet && firstOp?.range
      ? `${firstOp.worksheet}!${firstOp.range}`
      : firstOp?.range || firstOp?.worksheet;

    return {
      status: OperationStatusValues.READING_EXCEL,
      message: opCount > 0
        ? `Reading values (${opCount} ops)${
          firstLocation ? ` starting at ${firstLocation}` : ""
        }`
        : "Reading values (batch)...",
      progress: opCount > 0
        ? { current: 0, total: opCount, unit: "ops" }
        : undefined,
      metadata: {
        ...baseMetadata,
        range: firstOp?.range || baseMetadata.range,
      },
    };
  }

  if (toolName === ToolNames.READ_FORMAT_BATCH) {
    const operations = Array.isArray(inputData?.operations)
      ? inputData.operations
      : [];
    const opCount = operations.length;
    const firstOp = operations[0];
    const firstLocation = firstOp?.worksheet && firstOp?.range
      ? `${firstOp.worksheet}!${firstOp.range}`
      : firstOp?.range || firstOp?.worksheet;

    return {
      status: OperationStatusValues.READING_EXCEL_FORMAT,
      message: opCount > 0
        ? `Reading formatting (${opCount} ops)${
          firstLocation ? ` starting at ${firstLocation}` : ""
        }`
        : "Reading cell formatting (batch)...",
      progress: opCount > 0
        ? { current: 0, total: opCount, unit: "ops" }
        : undefined,
      metadata: {
        ...baseMetadata,
        range: firstOp?.range || baseMetadata.range,
      },
    };
  }
  return {
    status: OperationStatusValues.EXECUTING_TOOL,
    message: `Executing: ${friendlyName}`,
    metadata: baseMetadata,
  };
}

/**
 * Heuristic for detecting a Claude/Anthropic rare malformed tool payload where
 * the `operations` field is emitted as a JSON string containing the array.
 */
function looksLikeSplitOperations(value: string): boolean {
  return (
    typeof value === "string" &&
    value.trim().startsWith("[") &&
    value.includes('"formats":')
  );
}

/**
 * Reconstructs tool input when Claude/Anthropic rarely splits the `operations`
 * array into a string field, producing malformed JSON for our tools.
 */
function reconstructSplitToolInput(data: any): any {
  if (!data || typeof data !== "object") return data;

  if (
    typeof data.operations === "string" &&
    looksLikeSplitOperations(data.operations)
  ) {
    const reconstructedJson = `{
      "operations": ${data.operations}
    }`;

    logger.info("reconstructedJson: ", reconstructedJson);

    try {
      return JSON.parse(reconstructedJson);
    } catch {
      // If this fails, fall back to original data
      return data;
    }
  }

  return data;
}

function normalizeToolInput(toolName: string, inputData: any): any {
  const sanitizeName = (value: any) =>
    sanitizeWorkbookName(value) ||
    (typeof value === "string" ? value.trim() : undefined);

  let data = typeof inputData === "string"
    ? JSON.parse(inputData)
    : { ...(inputData || {}) };

  data = reconstructSplitToolInput(data);

  if (
    toolName === ToolNames.WRITE_VALUES_BATCH &&
    typeof data.operations === "string"
  ) {
    try {
      const parsed = JSON.parse(data.operations);
      if (Array.isArray(parsed)) {
        data.operations = parsed;
      }
    } catch (error) {
      logger.warn("Failed to parse write_values_batch operations string", {
        error,
      });
    }
  }

  // Ensure arrays
  if (!Array.isArray(data.operations)) data.operations = [];
  if (data.formats && !Array.isArray(data.formats)) data.formats = [];

  const fallbackWorkbook = sanitizeName(
    data.workbookName || data.currentWorkbookName ||
      data.selectedWorkbookName,
  );

  const fallbackWorksheet = data.worksheet || data.selectedWorksheet ||
    data.activeSheet;

  if (
    toolName === ToolNames.WRITE_VALUES_BATCH ||
    toolName === ToolNames.WRITE_FORMAT_BATCH ||
    toolName === ToolNames.READ_VALUES_BATCH ||
    toolName === ToolNames.READ_FORMAT_BATCH
  ) {
    data.workbookName = fallbackWorkbook;
    data.operations = data.operations.map((op: any) => {
      const opCopy = { ...(op || {}) };
      const opWorkbook = sanitizeName(opCopy.workbookName);
      opCopy.workbookName = opWorkbook ?? fallbackWorkbook;
      if (!opCopy.worksheet) opCopy.worksheet = fallbackWorksheet;
      if (toolName === ToolNames.WRITE_VALUES_BATCH) {
        const opRequested = sanitizeName(opCopy.requestedWorkbook);
        opCopy.requestedWorkbook = opRequested ?? opCopy.workbookName;
      }
      return opCopy;
    });
  }

  return data;
}

export function applyWorkbookGuards(
  toolName: string,
  inputData: any,
  activeWorkbookName: string,
  providedWorkbookName?: string,
): any {
  const activeWorkbook = sanitizeWorkbookName(activeWorkbookName) ??
    activeWorkbookName;
  const _providedWorkbook = sanitizeWorkbookName(providedWorkbookName) ??
    providedWorkbookName;
  const data = inputData ? { ...inputData } : {};
  const readToolsWithWorkbook = new Set<string>([
    ToolNames.READ_VALUES_BATCH,
    ToolNames.READ_FORMAT_BATCH,
    ToolNames.GET_METADATA,
    ToolNames.LIST_SHEETS,
    ToolNames.WORKBOOK_CHANGELOG,
  ]);

  if (WRITE_TOOLS.has(toolName)) {
    const requestedWorkbook = sanitizeWorkbookName(
      data.workbookName ||
        data.workbook ||
        data.targetWorkbook ||
        providedWorkbookName,
    ) ?? (data.workbookName || data.workbook || data.targetWorkbook ||
      providedWorkbookName);
    if (
      requestedWorkbook &&
      !equalsIgnoreCase(requestedWorkbook, activeWorkbook)
    ) {
      logger.warn("Provided workbookName differs from active; ignoring", {
        toolName,
        requestedWorkbook,
        activeWorkbook,
      });
    }

    data.workbookName = activeWorkbook;
    data.activeWorkbookName = activeWorkbook;
    if (Array.isArray(data.operations)) {
      data.operations = data.operations.map((op: any) => ({
        ...(op || {}),
        requestedWorkbook: sanitizeWorkbookName(op?.workbookName) ||
          requestedWorkbook,
        workbookName: activeWorkbook,
      }));
    }
    return data;
  }

  if (readToolsWithWorkbook.has(toolName)) {
    if (!data.workbookName && activeWorkbookName) {
      data.workbookName = activeWorkbookName;
    }

    if (Array.isArray(data.operations)) {
      data.operations = data.operations.map((op: any) => ({
        ...(op || {}),
        workbookName: op?.workbookName || data.workbookName,
      }));
    }
  }

  return data;
}

function isWriteTool(toolName: string): boolean {
  return WRITE_TOOLS.has(toolName);
}

function equalsIgnoreCase(a?: string, b?: string): boolean {
  return (a ?? "").toLowerCase() === (b ?? "").toLowerCase();
}

export function getRange(toolChange: ToolChangeInterface, result: any): string {
  switch (toolChange.toolName) {
    case ToolNames.DRAG_FORMULA:
      return result.range;
    case ToolNames.COPY_PASTE:
      return result?.destinationRange ||
        toolChange.inputData?.destinationRange;
    default:
      return "";
  }
}

export function getInputDataRevert(
  result: any,
  workbookName: string,
  toolChange: ToolChangeInterface,
): Record<string, any> {
  const tool = getToolByName(toolChange.toolName);
  if (!tool) {
    throw new Error(`Tool ${toolChange.toolName} not found`);
  }
  switch (toolChange.toolName) {
    case ToolNames.DRAG_FORMULA: {
      const oldValues = result.oldValues;
      const range = getRange(toolChange, result);
      return getRevertValuesForWriteValues(
        oldValues,
        range,
        workbookName,
        toolChange,
      );
    }
    case ToolNames.WRITE_VALUES_BATCH: {
      return getRevertValuesForWriteValuesBatch(
        result,
        workbookName,
        toolChange,
      );
    }
    case ToolNames.COPY_PASTE: {
      const copyResult = result?.copyPasteResult || result;
      const rawDestinationRange = copyResult?.destinationRange ||
        toolChange.inputData?.destinationRange;
      const destinationWorksheet = copyResult?.destinationWorksheet ||
        toolChange.inputData?.destinationWorksheet ||
        toolChange.inputData?.worksheet;
      const oldFormats = copyResult?.oldFormats || result?.oldFormats;
      const pasteType = copyResult?.pasteType ||
        toolChange.inputData?.pasteType;
      const destinationRange = resolveCopyPasteDestinationRange(
        rawDestinationRange,
        toolChange.inputData?.sourceRange,
      );
      const insertMode = (copyResult?.insertMode ||
        toolChange.inputData?.insertMode ||
        "none").toLowerCase();

      if (
        insertMode && insertMode !== "none" && destinationWorksheet &&
        destinationRange
      ) {
        const deleteCellsRequest: DeleteCellsRequest = {
          workbookName,
          worksheet: destinationWorksheet,
          range: destinationRange,
          shiftDirection: insertMode.includes("down") ? "up" : "left",
        };
        return { "delete-cells": deleteCellsRequest };
      }

      const oldValues = copyResult?.oldValues;
      if (!oldValues || !destinationRange || !destinationWorksheet) {
        logger.warn(
          "Missing data to build revert payload for copy_paste",
          {
            hasOldValues: !!oldValues,
            destinationRange,
            destinationWorksheet,
          },
        );
        return {} as Record<string, any>;
      }

      const patchedToolChange: ToolChangeInterface = {
        ...toolChange,
        inputData: {
          ...(toolChange.inputData || {}),
          worksheet: destinationWorksheet,
        },
      };

      const inputDataRevert = getRevertValuesForWriteValues(
        oldValues,
        destinationRange,
        workbookName,
        patchedToolChange,
      );
      const formatRevertEntries = getRevertValuesForCopyPasteFormat(
        oldFormats,
        workbookName,
        destinationWorksheet,
        destinationRange,
        pasteType,
      );
      if (Object.keys(formatRevertEntries).length > 0) {
        Object.assign(inputDataRevert, formatRevertEntries);
      }
      return inputDataRevert;
    }
    case ToolNames.WRITE_FORMAT_BATCH:
      return getRevertValuesForWriteFormatBatch(
        toolChange,
        result,
        workbookName,
      );
    case ToolNames.ADD_SHEETS:
      return getRevertValuesForAddSheets(toolChange.inputData);
    case ToolNames.REMOVE_COLUMNS:
      return getRevertValuesForRemoveColumns(toolChange, result);
    case ToolNames.REMOVE_ROWS:
      return getRevertValuesForRemoveRows(toolChange, result);
    case ToolNames.ADD_COLUMNS:
      return getRevertValuesForAddColumns(toolChange.inputData);
    case ToolNames.ADD_ROWS:
      return getRevertValuesForAddRows(toolChange.inputData);
    default:
      return {} as Record<string, any>;
  }
}

function getRevertValuesForWriteValues(
  oldValues: any,
  range: string,
  workbookName: string,
  toolChange: ToolChangeInterface,
): any {
  if (!range) {
    logger.warn("write_values revert skipped: missing range");
    return {} as Record<string, any>;
  }
  if (!oldValues || !oldValues.cellValues) {
    logger.warn("write_values revert skipped: missing oldValues.cellValues");
    return {} as Record<string, any>;
  }
  const cleanedRange = stripSheetPrefix(range);
  logger.info("getRevertValuesForWriteValues start", {
    oldValuesKeys: Object.keys(oldValues || {}),
    range: cleanedRange,
  });

  const oldValuesClean: Record<string, string> = {};
  for (const [address, value] of Object.entries(oldValues.cellValues)) {
    const valueObj = value as CellValue;
    oldValuesClean[address] = valueObj.formula || valueObj.value || "";
  }
  logger.info("oldValuesClean created", {
    cleanKeys: Object.keys(oldValuesClean),
  });

  const revertValues: string[][] = convertMapToValuesForRevert(
    oldValuesClean,
    cleanedRange,
    true,
  );
  logger.info("revert values", {
    revertValues: `${revertValues.slice(0, 5)}...`,
  });

  const writeValuesRevertParams: WriteValuesRequest = {
    workbookName: workbookName,
    worksheet: toolChange.inputData?.worksheet,
    range: cleanedRange,
    values: revertValues,
  };
  logger.info("writeValuesRevertParams created");

  const inputDataRevert: Record<string, any> = {
    [RevertOperationKeys.WRITE_VALUES_BATCH]: writeValuesRevertParams,
  };
  logger.info("inputDataRevert created, returning");

  return inputDataRevert;
}

function getRevertValuesForWriteValuesBatch(
  result: any,
  workbookName: string,
  toolChange: ToolChangeInterface,
): Record<string, any> {
  const ops = Array.isArray(toolChange.inputData?.operations)
    ? toolChange.inputData.operations
    : [];
  const oldValuesList = Array.isArray(result?.oldValues)
    ? result.oldValues
    : [];

  if (ops.length === 0) {
    logger.warn("write_values_batch revert skipped: no operations");
    return {} as Record<string, any>;
  }

  const itemCount = Math.min(ops.length, oldValuesList.length);
  if (oldValuesList.length !== ops.length) {
    logger.warn(
      "write_values_batch revert: oldValues length mismatch; using overlap only",
      {
        oldValuesLength: oldValuesList.length,
        operationsLength: ops.length,
      },
    );
    if (itemCount === 0) {
      logger.warn(
        "write_values_batch revert skipped: no overlapping operations/oldValues",
      );
      return {} as Record<string, any>;
    }
  }

  const revertRequests: WriteValuesRequest[] = [];

  for (let idx = 0; idx < itemCount; idx++) {
    const op = ops[idx];
    const oldValues = oldValuesList[idx];
    if (!oldValues) {
      logger.warn(
        "write_values_batch revert: missing oldValues; skipping operation",
        { index: idx },
      );
      continue;
    }
    if (!op?.range) {
      logger.warn(
        "write_values_batch revert: missing range; skipping operation",
        { index: idx },
      );
      continue;
    }

    const patchedToolChange: ToolChangeInterface = {
      ...toolChange,
      inputData: {
        ...(toolChange.inputData || {}),
        worksheet: op.worksheet,
      },
    };

    const revertEntry = getRevertValuesForWriteValues(
      oldValues,
      op.range,
      workbookName,
      patchedToolChange,
    );
    const writeParams = revertEntry[RevertOperationKeys.WRITE_VALUES_BATCH];
    if (writeParams) {
      revertRequests.push(writeParams);
    }
  }

  if (revertRequests.length === 0) {
    logger.warn("write_values_batch revert produced no revert requests");
    return {} as Record<string, any>;
  }

  if (revertRequests.length !== ops.length) {
    logger.warn("write_values_batch revert is partial", {
      requested: ops.length,
      built: revertRequests.length,
    });
  }

  return {
    [RevertOperationKeys.WRITE_VALUES_BATCH]: revertRequests,
  };
}

function getRevertValuesForAddColumns(inputData: any): Record<string, any> {
  const removeColumnsInputData: RemoveColumnsRequest = inputData;
  const inputDataRevert: Record<string, any> = {
    [RevertOperationKeys.REMOVE_COLUMNS]: removeColumnsInputData,
  };
  return inputDataRevert;
}

function getRevertValuesForAddRows(inputData: any): Record<string, any> {
  const addRowsInputData: AddRowsRequest = inputData;
  const inputDataRevert: Record<string, any> = {
    [RevertOperationKeys.ADD_ROWS]: addRowsInputData,
  };
  return inputDataRevert;
}

function getRevertValuesForCopyPasteFormat(
  oldFormats: ReadFormatFinalResponseList | undefined,
  workbookName: string,
  worksheet: string | undefined,
  fallbackRange: string | undefined,
  pasteType: string | undefined,
): Record<string, any> {
  const inputDataRevert: Record<string, any> = {};
  if (!worksheet) {
    return inputDataRevert;
  }

  const shouldClearBorders = shouldClearBordersForPasteType(pasteType);

  if (oldFormats && oldFormats.length > 0) {
    const flattenedFormats: RangeFormatPair[] = flattenToRangeFormatPairs(
      oldFormats,
    );
    for (const rangeFormatList of flattenedFormats) {
      const formatPayload = shouldClearBorders
        ? { ...rangeFormatList.format, clearBorders: true }
        : rangeFormatList.format;
      const entryKey = `write-format-${rangeFormatList.range}`;
      inputDataRevert[entryKey] = {
        workbookName,
        worksheet,
        range: rangeFormatList.range,
        format: formatPayload,
      };
    }
    return inputDataRevert;
  }

  if (fallbackRange) {
    const entryKey = `write-format-${fallbackRange}`;
    inputDataRevert[entryKey] = {
      workbookName,
      worksheet,
      range: fallbackRange,
      format: shouldClearBorders ? { clearBorders: true } : {},
    };
  }

  return inputDataRevert;
}

function shouldClearBordersForPasteType(pasteType?: string): boolean {
  const normalized = (pasteType || "all").toLowerCase();
  return normalized === "all" || normalized === "formats";
}

function resolveCopyPasteDestinationRange(
  destinationRange: string | undefined,
  sourceRange: string | undefined,
): string | undefined {
  if (!destinationRange) return destinationRange;
  if (!sourceRange) return destinationRange;

  const hasColumnLetters = /[A-Za-z]/.test(destinationRange);
  if (hasColumnLetters) return destinationRange;

  const sourceBounds = parseRangeBounds(sourceRange);
  if (!sourceBounds) return destinationRange;

  const destinationBounds = parseRangeBounds(destinationRange);
  if (!destinationBounds) return destinationRange;

  const startColName = numberToColumn(sourceBounds.startCol);
  const endColName = numberToColumn(
    sourceBounds.startCol + (sourceBounds.endCol - sourceBounds.startCol),
  );
  const startRow = destinationBounds.startRow;
  const endRow = destinationBounds.endRow;

  const startCell = `${startColName}${startRow}`;
  const endCell = `${endColName}${endRow}`;
  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function getRevertValuesForRemoveRows(
  toolChange: ToolChangeInterface,
  result: any,
): Record<string, any> {
  logger.info("getting revert values for remove rows");
  const oldValuesMap = result.oldValues;
  const inputData = toolChange.inputData;

  const oldValuesClean: Record<string, string> = {};

  for (const [address, value] of Object.entries(oldValuesMap.cellValues)) {
    logger.info("address", { address });
    const valueObj = value as CellValue;
    oldValuesClean[address] = valueObj.formula || valueObj.value || "";
  }

  logger.info("old values clean", { oldValuesClean });
  const range = getExcelRangeFromCells(oldValuesClean);

  logger.info("range", { range });
  const revertValues = convertMapToValuesForRevert(oldValuesClean, range, true);

  logger.info("revert values", {
    revertValues: `${revertValues.slice(0, 5)}...`,
  });
  const addRowsInputData: AddRowsRequest = inputData;
  const writeValuesInputData: WriteValuesRequest = {
    workbookName: inputData.workbookName,
    worksheet: inputData.worksheet,
    range: range,
    values: revertValues,
  };

  const inputDataRevert: Record<string, any> = {
    [RevertOperationKeys.ADD_ROWS]: addRowsInputData,
    [RevertOperationKeys.WRITE_VALUES_BATCH]: writeValuesInputData,
  };
  return inputDataRevert;
}

function convertMapToValuesForRevert(
  oldValues: Record<string, string>,
  range: string,
  includeEmpty: boolean = false,
): string[][] {
  logger.info("converting map to values for revert", {
    range,
    includeEmpty,
  });

  const bounds = parseRangeBounds(range);
  const normalizedOldValues: Record<string, string> = {};
  for (const [address, value] of Object.entries(oldValues)) {
    const normalizedAddress = normalizeCellAddress(address);
    if (normalizedAddress) {
      normalizedOldValues[normalizedAddress] = value;
    }
  }

  if (!bounds) {
    const normalizedSingle = normalizeCellAddress(range);
    const fallbackValue = normalizedSingle
      ? normalizedOldValues[normalizedSingle]
      : undefined;
    return [[
      includeEmpty || fallbackValue !== undefined ? fallbackValue || "" : "",
    ]];
  }

  const rowCount = bounds.endRow - bounds.startRow + 1;
  const colCount = bounds.endCol - bounds.startCol + 1;
  const values: string[][] = Array.from(
    { length: rowCount },
    () => Array.from({ length: colCount }, () => ""),
  );

  for (let row = bounds.startRow; row <= bounds.endRow; row++) {
    for (let col = bounds.startCol; col <= bounds.endCol; col++) {
      const address = `${numberToColumn(col)}${row}`;
      const value = normalizedOldValues[address];
      if (includeEmpty || value !== undefined) {
        values[row - bounds.startRow][col - bounds.startCol] = value || "";
      }
    }
  }

  logger.info("converted values for revert", {
    valuesLength: values.length,
    values: values.length > 0 ? values : "empty",
  });

  return values;
}

function parseRangeBounds(range: string): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} | null {
  const parts = range.includes(":") ? range.split(":") : [range, range];
  if (parts.length !== 2) return null;

  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  if (!start || !end) return null;

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

function parseCellAddress(
  address: string,
): { row: number; col: number } | null {
  const normalized = normalizeCellAddress(address);
  if (!normalized) return null;
  const match = normalized.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const col = columnToNumber(match[1]);
  const row = parseInt(match[2], 10);
  if (!row || !col) return null;
  return { row, col };
}

function normalizeCellAddress(address: string): string | null {
  if (!address) return null;
  const noSheet = stripSheetPrefix(address);
  const trimmed = noSheet.replace(/\$/g, "").trim();
  const match = trimmed.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;
  return `${match[1].toUpperCase()}${match[2]}`;
}

function stripSheetPrefix(value: string): string {
  const bangIndex = value.lastIndexOf("!");
  if (bangIndex === -1) return value;
  return value.slice(bangIndex + 1);
}
function columnToNumber(col: string): number {
  return col.split("").reduce(
    (acc, char) => acc * 26 + (char.charCodeAt(0) - 64),
    0,
  );
}

function numberToColumn(num: number): string {
  let result = "";
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

interface ToolPermissionPayload {
  toolPermissions: ToolPermission[];
}

interface ToolPermission {
  toolId: string;
  toolName: string;
}

function getRevertValuesForAddSheets(inputData: any): any {
  const removeSheetsInputData: RemoveSheetsRequest = inputData;

  const addSheetsRevertParams: Record<string, any> = {
    [RevertOperationKeys.REMOVE_SHEETS]: removeSheetsInputData,
  };
  return addSheetsRevertParams;
}

function getRevertValuesForRemoveColumns(
  toolChange: ToolChangeInterface,
  result: any,
): Record<string, any> {
  logger.info("getting revert values for remove columns");
  const oldValuesMap = result.oldValues;
  const inputData = toolChange.inputData;

  const oldValuesClean: Record<string, string> = {};
  for (const [address, value] of Object.entries(oldValuesMap.cellValues)) {
    const valueObj = value as CellValue;
    oldValuesClean[address] = valueObj.formula || valueObj.value || "";
  }

  const range = getExcelRangeFromCells(oldValuesClean);

  const revertValues = convertMapToValuesForRevert(oldValuesClean, range, true);

  const writeValuesRevertParams: WriteValuesRequest = {
    workbookName: inputData.workbookName,
    worksheet: inputData.worksheet,
    range: range,
    values: revertValues,
  };

  const addColumnsRevertParams: AddColumnsRequest = {
    ...inputData,
  };

  const inputDataRevert: Record<string, any> = {
    [RevertOperationKeys.WRITE_VALUES_BATCH]: writeValuesRevertParams,
    [RevertOperationKeys.ADD_COLUMNS]: addColumnsRevertParams,
  };

  logger.info("inputDataRevert created for remove_columns", {
    range,
    hasValues: revertValues.length > 0,
  });

  return inputDataRevert;
}

function getExcelRangeFromCells(cellData: Record<string, string>): string {
  const addresses = Object.keys(cellData);

  if (addresses.length === 0) {
    return "";
  }

  if (addresses.length === 1) {
    return addresses[0]; // Single cell like "H1"
  }

  // Parse all cell addresses to find min/max bounds
  const parsedCells = addresses.map((addr) => {
    const colMatch = addr.match(/[A-Z]+/)?.[0] || "A";
    const rowMatch = parseInt(addr.match(/\d+/)?.[0] || "1");
    return {
      col: colMatch,
      row: rowMatch,
      colNum: columnToNumber(colMatch),
    };
  });

  // Find bounds
  const minRow = Math.min(...parsedCells.map((c) => c.row));
  const maxRow = Math.max(...parsedCells.map((c) => c.row));
  const minColNum = Math.min(...parsedCells.map((c) => c.colNum));
  const maxColNum = Math.max(...parsedCells.map((c) => c.colNum));

  const startCol = numberToColumn(minColNum);
  const endCol = numberToColumn(maxColNum);

  return `${startCol}${minRow}:${endCol}${maxRow}`;
}

function getRevertValuesForWriteFormatBatch(
  toolChange: ToolChangeInterface,
  result: any,
  workbookName: string,
): Record<string, any> {
  const oldFormats:
    | ReadFormatFinalResponseList
    | ReadFormatFinalResponseList[] = result.oldFormats;
  const inputData = toolChange.inputData;

  logger.info("Revert(build) write_format_batch", {
    toolName: toolChange.toolName,
    hasOldFormats: !!oldFormats,
    oldFormatsCount: Array.isArray(oldFormats) ? oldFormats.length : undefined,
    operationsCount: Array.isArray(inputData?.operations)
      ? inputData.operations.length
      : undefined,
  });

  const inputDataRevert: Record<string, any> = {};
  const revertEntryKeys: string[] = [];

  const ops = Array.isArray(inputData?.operations) ? inputData.operations : [];
  const oldFormatLists: ReadFormatFinalResponseList[] = Array.isArray(
      oldFormats,
    )
    ? oldFormats as ReadFormatFinalResponseList[]
    : oldFormats
    ? [oldFormats]
    : [];

  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    const formatsForOp: ReadFormatFinalResponseList | undefined =
      oldFormatLists.length === ops.length
        ? oldFormatLists[i]
        : oldFormatLists.length === 1
        ? oldFormatLists[0]
        : undefined;

    const flattened = formatsForOp
      ? flattenToRangeFormatPairs(formatsForOp)
      : [];

    if (flattened.length > 0) {
      for (const pair of flattened) {
        const range = pair.range;
        const entryKey = `${RevertOperationKeys.WRITE_FORMAT_PREFIX}${range}`;
        inputDataRevert[entryKey] = {
          workbookName: op.workbookName || workbookName,
          worksheet: op.worksheet,
          range,
          format: pair.format,
        };
        revertEntryKeys.push(entryKey);
        logger.info("Batch write_format revert entry added", {
          entryKey,
          workbookName: op.workbookName || workbookName,
          worksheet: op.worksheet,
          range,
          formatKeys: pair.format ? Object.keys(pair.format) : [],
        });
      }
    } else if (op?.range) {
      // No old formats captured; fallback to clearing the batch-applied format.
      const entryKey = `${RevertOperationKeys.WRITE_FORMAT_PREFIX}${op.range}`;
      inputDataRevert[entryKey] = {
        workbookName: op.workbookName || workbookName,
        worksheet: op.worksheet,
        range: op.range,
        format: {},
      };
      revertEntryKeys.push(entryKey);
      logger.info(
        "Batch write_format revert entry added with empty format fallback",
        {
          entryKey,
          workbookName: op.workbookName || workbookName,
          worksheet: op.worksheet,
          range: op.range,
        },
      );
    }
  }

  logger.info("Completed building write_format_batch revert payload", {
    entryCount: revertEntryKeys.length,
    entryKeys: revertEntryKeys,
  });

  return inputDataRevert;
}

function flattenToRangeFormatPairs(
  list: ReadFormatFinalResponseList,
): RangeFormatPair[] {
  return list.flatMap((item) =>
    item.ranges.map((range) => ({
      range,
      format: item.format,
    }))
  );
}
