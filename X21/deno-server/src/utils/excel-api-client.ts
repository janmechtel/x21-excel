import { ExcelApiConfigService } from "./excel-api-config.ts";
import { createLogger } from "./logger.ts";
import {
  OpenWorkbook,
  OperationStatusValues,
  WorkbookMetadata,
} from "../types/index.ts";
import { WebSocketManager } from "../services/websocket-manager.ts";

const logger = createLogger("ExcelApiClient");

/**
 * Generic Excel API client for making HTTP requests to Excel actions
 */
export class ExcelApiError extends Error {
  constructor(
    message: string,
    public readonly status: number,
    public readonly code?: string,
    public readonly details?: unknown,
  ) {
    super(message);
    this.name = "ExcelApiError";
  }
}

export class ExcelApiClient {
  private static instance: ExcelApiClient;
  private excelApi: ExcelApiConfigService;
  private static readonly ExcelNotReadyCode = "EXCEL_NOT_READY";
  private static readonly WaitBaseDelayMs = 250;
  private static readonly WaitMaxDelayMs = 2000;
  private static readonly WaitMaxMs = 10 * 60 * 1000;
  private static readonly WaitMaxAttempts = 1200;

  private constructor() {
    this.excelApi = ExcelApiConfigService.getInstance();
  }

  public static getInstance(): ExcelApiClient {
    if (!ExcelApiClient.instance) {
      ExcelApiClient.instance = new ExcelApiClient();
    }
    return ExcelApiClient.instance;
  }

  /**
   * Execute a generic Excel action
   * @param action - The action name (e.g., ToolNames.ADD_COLUMNS, ToolNames.ADD_ROWS, etc.)
   * @param params - The parameters for the action
   * @param validationFn - Optional validation function for the parameters
   * @returns Promise with the API response
   */
  public async executeAction<TRequest, TResponse>(
    action: string,
    params: TRequest,
  ): Promise<TResponse> {
    logger.info(`📡 Making HTTP call to Excel API for ${action}...`);

    const body = {
      action,
      ...params,
    };

    return await this.executeWithExcelReadyRetry(
      action,
      params,
      async () => {
        const response = await fetch(this.excelApi.getApiUrlActionExecution(), {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(body),
        });

        if (!response.ok) {
          throw await this.buildError(response);
        }

        const result = await response.json();
        logger.info(`✅ ${action} completed successfully via Excel API`);

        return result;
      },
    );
  }

  /**
   * Convenience method for Excel actions - no built-in validation
   */
  public async executeExcelAction<TRequest, TResponse>(
    action: string,
    params: TRequest,
  ): Promise<TResponse> {
    if (!params) {
      throw new Error("params is required");
    }

    ExcelApiClient.validateAllPropertiesNotNull(params);
    logger.info("Excel action payload summary", {
      action,
      workbookName: (params as any)?.workbookName,
      worksheet: (params as any)?.worksheet,
      range: (params as any)?.range,
    });
    return await this.executeAction<TRequest, TResponse>(action, params);
  }

  /**
   * Validates that all properties in the request object are not null or undefined
   * Note: Only validates properties that are actually present in the object.
   * Optional properties (undefined) are allowed and will be skipped.
   */
  public static validateAllPropertiesNotNull<T extends Record<string, any>>(
    params: T,
  ): void {
    for (const [key, value] of Object.entries(params)) {
      // Only validate if the property is explicitly set (not undefined)
      // This allows optional properties to be omitted
      if (value === null) {
        throw new Error(`${key} is required and cannot be null`);
      }
      // undefined is allowed - it means the optional property was not provided
    }
  }

  private async buildError(response: Response): Promise<ExcelApiError> {
    let message = `HTTP error! status: ${response.status}`;
    let code: string | undefined;
    let details: unknown;

    try {
      const body = await response.json();
      message = body?.error || body?.message || message;
      code = body?.errorCode || body?.code;
      details = body?.candidates || body;
    } catch {
      // ignore JSON parse errors and fall back to default message
    }

    return new ExcelApiError(message, response.status, code, details);
  }

  public async getMetadata(
    params?: { workbookName?: string },
  ): Promise<WorkbookMetadata> {
    const url = this.excelApi.getApiUrlMetadata(params?.workbookName);
    return await this.executeWithExcelReadyRetry(
      "get_metadata",
      params,
      async () => {
        const response = await fetch(url, {
          method: "GET",
          headers: { "Content-Type": "application/json" },
        });

        if (!response.ok) {
          throw await this.buildError(response);
        }

        return await response.json();
      },
    );
  }

  public async listOpenWorkbooks(): Promise<OpenWorkbook[]> {
    const urlAll = this.excelApi.getApiUrlOpenWorkbooksAll();
    try {
      const response = await this.executeWithExcelReadyRetry(
        "list_open_workbooks",
        undefined,
        async () => {
          const res = await fetch(urlAll, { method: "GET" });
          if (!res.ok) {
            throw await this.buildError(res);
          }
          return res;
        },
      );

      const payload = await response.json();
      const workbooks = payload?.workbooks ?? [];
      logger.info("[OpenWorkbooksAll] Response", {
        count: workbooks.length,
        withSheets: workbooks.filter((wb: OpenWorkbook) =>
          Array.isArray(wb?.sheets) && wb.sheets.length > 0
        ).length,
      });
      return workbooks;
    } catch (error) {
      if (ExcelApiClient.isExcelNotReadyError(error)) {
        throw error;
      }

      const primaryError = error as ExcelApiError;
      logger.warn(
        "[OpenWorkbooksAll] Request failed, falling back to /api/openWorkbooks",
        {
          status: primaryError.status,
          message: primaryError.message,
          code: primaryError.code,
        },
      );

      const fallbackUrl = this.excelApi.getApiUrlOpenWorkbooks();
      const fallbackResponse = await this.executeWithExcelReadyRetry(
        "list_open_workbooks_fallback",
        undefined,
        async () => {
          const res = await fetch(fallbackUrl, { method: "GET" });
          if (!res.ok) {
            throw await this.buildError(res);
          }
          return res;
        },
      );

      const fallbackPayload = await fallbackResponse.json();
      const fallbackWorkbooks = fallbackPayload?.workbooks ?? [];
      logger.info("[OpenWorkbooks] Fallback response", {
        count: fallbackWorkbooks.length,
        withSheets: fallbackWorkbooks.filter((wb: OpenWorkbook) =>
          Array.isArray(wb?.sheets) && wb.sheets.length > 0
        ).length,
      });
      return fallbackWorkbooks;
    }
  }

  private async executeWithExcelReadyRetry<T>(
    action: string,
    params: unknown,
    request: () => Promise<T>,
  ): Promise<T> {
    const start = Date.now();
    let attempt = 0;
    let delayMs = ExcelApiClient.WaitBaseDelayMs;
    let lastStatusAt = 0;
    const workbookName = ExcelApiClient.extractWorkbookName(params);

    while (attempt < ExcelApiClient.WaitMaxAttempts) {
      try {
        return await request();
      } catch (error) {
        if (!ExcelApiClient.isExcelNotReadyError(error)) {
          logger.error(`❌ Error calling Excel API for ${action}:`, {
            message: (error as Error)?.message,
            name: (error as Error)?.name,
          });
          throw error;
        }

        attempt += 1;
        logger.info("Excel not ready, will retry", {
          action,
          attempt,
          delayMs,
          workbookName,
        });
        if (attempt >= ExcelApiClient.WaitMaxAttempts) {
          logger.error("Excel stayed busy too long, aborting retry", {
            action,
            attempts: attempt,
            elapsedMs: Date.now() - start,
            reason: "max_attempts",
          });
          throw error;
        }
        const now = Date.now();
        if (now - start > ExcelApiClient.WaitMaxMs) {
          logger.error("Excel stayed busy too long, aborting retry", {
            action,
            attempts: attempt,
            elapsedMs: now - start,
            reason: "max_elapsed",
          });
          throw error;
        }

        if (workbookName) {
          if (attempt === 1 || now - lastStatusAt > 5000) {
            WebSocketManager.getInstance().sendStatus(
              workbookName,
              OperationStatusValues.EXECUTING_TOOL,
              ExcelApiClient.getExcelNotReadyStatusMessage(error),
            );
            lastStatusAt = now;
          }
        }

        await ExcelApiClient.sleep(delayMs);
        delayMs = Math.min(
          ExcelApiClient.WaitMaxDelayMs,
          Math.round(delayMs * 1.5),
        );
      }
    }

    throw new Error(
      `Excel stayed busy too long, aborting retry (max attempts ${ExcelApiClient.WaitMaxAttempts})`,
    );
  }

  private static isExcelNotReadyError(error: unknown): boolean {
    if (!error || typeof error !== "object") return false;
    const excelError = error as ExcelApiError;
    if (excelError.code === ExcelApiClient.ExcelNotReadyCode) return true;
    const message = (excelError.message || "").toLowerCase();
    return message.includes("excel is busy") ||
      message.includes("calculation in progress") ||
      message.includes("edit mode") ||
      message.includes("not ready");
  }

  private static getExcelNotReadyStatusMessage(error: unknown): string {
    const message = (error as ExcelApiError)?.message ?? "";
    const lower = message.toLowerCase();
    if (lower.includes("calculation in progress")) {
      return "Waiting for Excel (calculation in progress)...";
    }
    if (lower.includes("edit mode")) {
      return "Waiting for Excel (exit edit mode)...";
    }
    if (lower.includes("modal dialog")) {
      return "Waiting for Excel (close modal dialog)...";
    }
    return "Waiting for Excel...";
  }

  private static extractWorkbookName(params: unknown): string | undefined {
    if (!params || typeof params !== "object") return undefined;
    const anyParams = params as Record<string, any>;
    const direct = anyParams.workbookName || anyParams.workbook ||
      anyParams.activeWorkbookName;
    if (typeof direct === "string" && direct.trim()) return direct;

    const operations = Array.isArray(anyParams.operations)
      ? anyParams.operations
      : [];
    for (const op of operations) {
      const name = op?.workbookName || op?.requestedWorkbook;
      if (typeof name === "string" && name.trim()) return name;
    }

    return undefined;
  }

  private static sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}
