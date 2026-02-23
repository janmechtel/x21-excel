import {
  ReadFormatBatchRequest,
  ReadFormatFinalResponseList,
  ReadFormatRequest,
} from "../types/index.ts";
import { createLogger } from "../utils/logger.ts";
import { readFormatBatch } from "./read-format-batch.ts";

const logger = createLogger("ReadFormat");

export async function readFormat(
  params: ReadFormatRequest,
): Promise<ReadFormatFinalResponseList> {
  const batchRequest: ReadFormatBatchRequest = { operations: [params] };
  const batchResult = await readFormatBatch(batchRequest);
  const first = batchResult[0];
  if (!first) {
    throw new Error("read_format_batch returned no results");
  }
  logger.info("Result from readFormat (batch)", {
    ranges: first.map((res) => res.ranges).flat(),
  });
  return first;
}
