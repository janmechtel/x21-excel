import {
  ReadValuesBatchRequest,
  ReadValuesBatchResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";
import { createLogger } from "../utils/logger.ts";
import { chunkArray } from "../utils/batching.ts";

const logger = createLogger("ReadValuesBatchAction");

export async function readValuesBatch(
  params: ReadValuesBatchRequest,
): Promise<ReadValuesBatchResponse> {
  const operations = params.operations ?? [];
  if (operations.length === 0) {
    throw new Error("read_values_batch requires at least one operation.");
  }

  logger.info(
    `read_values_batch request: ops=${operations.length}`,
    operations.map((op) => `${op.workbookName}/${op.worksheet}/${op.range}`),
  );

  const client = ExcelApiClient.getInstance();
  const batches = chunkArray(operations);
  const results: ReadValuesBatchResponse["results"] = [];
  let overallSuccess = true;
  let readCount = 0;

  for (const [index, batch] of batches.entries()) {
    const start = Date.now();
    const batchResult = await client.executeExcelAction<
      ReadValuesBatchRequest,
      ReadValuesBatchResponse
    >(ToolNames.READ_VALUES_BATCH, { operations: batch });
    const elapsedMs = Date.now() - start;

    const batchResults = Array.isArray(batchResult?.results)
      ? batchResult.results
      : [];
    results.push(...batchResults);
    overallSuccess = overallSuccess && !!batchResult?.success;
    readCount += batchResults.filter((res) => res?.success).length;

    logger.info("read_values_batch batch complete", {
      operationType: ToolNames.READ_VALUES_BATCH,
      totalOps: operations.length,
      batches: batches.length,
      batchIndex: index + 1,
      batchSize: batch.length,
      elapsedMs,
    });
  }

  const message = `Read values for ${readCount}/${operations.length} range(s)`;

  logger.info(
    `read_values_batch result: success=${overallSuccess}, message="${message}", results=${results.length}`,
  );

  return {
    success: overallSuccess,
    message,
    results,
  };
}
