import { groupCellFormatsByRanges } from "../utils/format-grouping.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";
import {
  ReadFormatBatchRequest,
  ReadFormatBatchResponse,
  ReadFormatFinalResponseList,
  ToolNames,
} from "../types/index.ts";
import { createLogger } from "../utils/logger.ts";
import { chunkArray } from "../utils/batching.ts";

const logger = createLogger("ReadFormatBatch");

/**
 * Reads formats for multiple ranges in a single round-trip to the Excel API.
 * Returns a list of grouped format responses in the same order as the request operations.
 */
export async function readFormatBatch(
  params: ReadFormatBatchRequest,
): Promise<ReadFormatFinalResponseList[]> {
  const operations = params.operations ?? [];
  if (operations.length === 0) {
    throw new Error("read_format_batch requires at least one operation.");
  }

  logger.info(
    `readFormatBatch request: ops=${operations.length}`,
    operations.map((op) => `${op.workbookName}/${op.worksheet}/${op.range}`),
  );
  const client = ExcelApiClient.getInstance();
  const batches = chunkArray(operations);
  const groupedResults: ReadFormatFinalResponseList[] = [];
  let overallSuccess = true;

  for (const [index, batch] of batches.entries()) {
    const start = Date.now();
    const batchResult = await client.executeExcelAction<
      ReadFormatBatchRequest,
      ReadFormatBatchResponse
    >(ToolNames.READ_FORMAT_BATCH, { operations: batch });
    const elapsedMs = Date.now() - start;

    const batchResults = Array.isArray(batchResult?.results)
      ? batchResult.results
      : [];
    groupedResults.push(
      ...batchResults.map((res) => groupCellFormatsByRanges(res.cellFormats)),
    );
    overallSuccess = overallSuccess && !!batchResult?.success;

    logger.info("read_format_batch batch complete", {
      operationType: ToolNames.READ_FORMAT_BATCH,
      totalOps: operations.length,
      batches: batches.length,
      batchIndex: index + 1,
      batchSize: batch.length,
      elapsedMs,
    });
  }

  logger.info(
    `readFormatBatch result: success=${overallSuccess}, results=${groupedResults.length}`,
  );

  return groupedResults;
}
