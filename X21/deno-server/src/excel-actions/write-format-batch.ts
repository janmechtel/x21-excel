import { ExcelApiClient } from "../utils/excel-api-client.ts";
import {
  ToolNames,
  WriteFormatBatchRequest,
  WriteFormatBatchResponse,
} from "../types/index.ts";
import { createLogger } from "../utils/logger.ts";
import { chunkArray } from "../utils/batching.ts";

const logger = createLogger("WriteFormatBatchAction");

export async function writeFormatBatch(
  params: WriteFormatBatchRequest,
): Promise<WriteFormatBatchResponse> {
  const operations = params.operations ?? [];
  if (operations.length === 0) {
    throw new Error("write_format_batch requires at least one operation.");
  }

  logger.info(
    `write_format_batch request: ops=${operations.length}`,
    operations.map((op) => `${op.workbookName}/${op.worksheet}/${op.range}`),
  );

  const client = ExcelApiClient.getInstance();
  const batches = chunkArray(operations);
  const results: WriteFormatBatchResponse["results"] = [];
  let overallSuccess = true;
  let applied = 0;

  for (const [index, batch] of batches.entries()) {
    const start = Date.now();
    const batchResult = await client.executeExcelAction<
      WriteFormatBatchRequest,
      WriteFormatBatchResponse
    >(ToolNames.WRITE_FORMAT_BATCH, { operations: batch });
    const elapsedMs = Date.now() - start;

    const batchResults = Array.isArray(batchResult?.results)
      ? batchResult.results
      : [];
    results.push(...batchResults);
    overallSuccess = overallSuccess && !!batchResult?.success;
    applied += batchResults.filter((res) => res?.success).length;

    logger.info("write_format_batch batch complete", {
      operationType: ToolNames.WRITE_FORMAT_BATCH,
      totalOps: operations.length,
      batches: batches.length,
      batchIndex: index + 1,
      batchSize: batch.length,
      elapsedMs,
    });
  }

  const message =
    `Successfully formatted ${applied}/${operations.length} range(s)`;

  logger.info(
    `write_format_batch result: success=${overallSuccess}, message="${message}", results=${results.length}`,
  );

  return {
    success: overallSuccess,
    message,
    results,
    applied,
    batches: batches.length,
  };
}
