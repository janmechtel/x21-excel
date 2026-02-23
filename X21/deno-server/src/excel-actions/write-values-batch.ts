import {
  ToolNames,
  WriteValuesBatchRequest,
  WriteValuesBatchResponse,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";
import { createLogger } from "../utils/logger.ts";
import { chunkArray } from "../utils/batching.ts";
import { sanitizeWorkbookName } from "../utils/workbook-name.ts";

const logger = createLogger("WriteValuesBatchAction");

export async function writeValuesBatch(
  params: WriteValuesBatchRequest,
): Promise<WriteValuesBatchResponse> {
  const columnWidthMode = params.columnWidthMode;
  const operations = (params.operations ?? []).map((op) => {
    const rawWorkbook = op.workbookName;
    const rawRequested = op.requestedWorkbook;
    const workbookName = sanitizeWorkbookName(rawWorkbook) || rawWorkbook;
    const requestedWorkbook = sanitizeWorkbookName(rawRequested) ||
      sanitizeWorkbookName(rawWorkbook) ||
      rawRequested ||
      rawWorkbook;
    return {
      ...op,
      workbookName,
      requestedWorkbook,
    };
  });
  if (operations.length === 0) {
    throw new Error("write_values_batch requires at least one operation.");
  }

  logger.info(
    `write_values_batch request: ops=${operations.length}`,
    operations.map((op) => `${op.workbookName}/${op.worksheet}/${op.range}`),
  );

  const client = ExcelApiClient.getInstance();
  const batches = chunkArray(operations);
  const results: WriteValuesBatchResponse["results"] = [];
  const operationsResults: WriteValuesBatchResponse["operations"] = [];
  let overallSuccess = true;
  let applied = 0;

  for (const [index, batch] of batches.entries()) {
    const start = Date.now();
    const batchResult = await client.executeExcelAction<
      WriteValuesBatchRequest,
      WriteValuesBatchResponse
    >(
      ToolNames.WRITE_VALUES_BATCH,
      { operations: batch, columnWidthMode },
    );
    const elapsedMs = Date.now() - start;

    const batchResults = Array.isArray(batchResult?.results)
      ? batchResult.results
      : [];
    results.push(...batchResults);
    overallSuccess = overallSuccess && !!batchResult?.success;
    applied += batchResults.filter((res) => res?.success).length;

    // Map results to operation stubs for compatibility
    batch.forEach((op, idx) => {
      const res = batchResults[idx];
      operationsResults.push({
        worksheet: op.worksheet,
        range: op.range,
        requestedWorkbook: sanitizeWorkbookName(op.requestedWorkbook) ||
          sanitizeWorkbookName(op.workbookName) ||
          op.requestedWorkbook || op.workbookName,
        resolvedWorkbook: sanitizeWorkbookName(op.workbookName) ||
          op.workbookName,
        policyAction: "wrote",
        status: res?.success ? "success" : "error",
        oldValues: null,
        newValues: op.values as any[][],
        errorCode: res?.success ? undefined : "WRITE_FAILED",
        errorMessage: res?.message,
      });
    });

    logger.info("write_values_batch batch complete", {
      operationType: ToolNames.WRITE_VALUES_BATCH,
      totalOps: operations.length,
      batches: batches.length,
      batchIndex: index + 1,
      batchSize: batch.length,
      elapsedMs,
    });
  }

  const message =
    `Batch write completed (${applied}/${operations.length} successful)`;

  logger.info(
    `write_values_batch result: success=${overallSuccess}, message="${message}", results=${results.length}`,
  );

  const activeWorkbook = sanitizeWorkbookName(operations[0]?.workbookName) ||
    operations[0]?.workbookName;
  const requestedWorkbook = sanitizeWorkbookName(
    operations[0]?.requestedWorkbook,
  ) || operations[0]?.requestedWorkbook;

  return {
    success: overallSuccess,
    message,
    results,
    tool: ToolNames.WRITE_VALUES_BATCH,
    activeWorkbook,
    requestedWorkbook,
    resolvedWorkbook: activeWorkbook,
    policyAction: overallSuccess ? "wrote" : "partial",
    operations: operationsResults,
    applied,
    batches: batches.length,
  };
}
