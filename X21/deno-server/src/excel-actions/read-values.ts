import {
  ReadValuesBatchRequest,
  ReadValuesRequest,
  ReadValuesResponse,
} from "../types/index.ts";
import { readValuesBatch } from "./read-values-batch.ts";

export async function readValues(
  params: ReadValuesRequest,
): Promise<ReadValuesResponse> {
  const batchRequest: ReadValuesBatchRequest = { operations: [params] };
  const batchResult = await readValuesBatch(batchRequest);
  const first = batchResult.results?.[0];
  if (!first) {
    throw new Error("read_values_batch returned no results");
  }
  if (first.success === false) {
    throw new Error(first.message || "read_values_batch failed");
  }
  return { cellValues: first.cellValues };
}
