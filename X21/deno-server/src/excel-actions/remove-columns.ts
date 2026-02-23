import {
  RemoveColumnsRequest,
  RemoveColumnsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function removeColumns(
  params: RemoveColumnsRequest,
): Promise<RemoveColumnsResponse> {
  // console.log('params', params);
  // const modifiedParams = {
  //     ...params,
  //     columnRange: ensureColumnRangeWithValidation(params.columnRange)
  // }

  // console.log('modifiedParams', modifiedParams);

  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<
    RemoveColumnsRequest,
    RemoveColumnsResponse
  >(
    ToolNames.REMOVE_COLUMNS,
    params,
  );
}

/**
 * Converts a column letter specification to a proper Excel column range
 * @param columnSpec Column specification like "A", "AB", "A:A", "B:Z", etc.
 * @returns Proper Excel column range like "A:A", "AB:AB", "A:A", "B:Z"
 */
export function ensureColumnRange(columnSpec: string): string {
  // Trim whitespace
  const trimmed = columnSpec.trim().toUpperCase();

  // If it already contains a colon, it's already a range - return as is
  if (trimmed.includes(":")) {
    return trimmed;
  }

  // If it's a single column letter/letters, convert to range format
  // Check if it's a valid column letter pattern (only letters A-Z)
  if (/^[A-Z]+$/.test(trimmed)) {
    return `${trimmed}:${trimmed}`;
  }

  // If it doesn't match expected patterns, return as is (let Excel handle the error)
  return trimmed;
}

/**
 * Converts a column letter specification to a proper Excel column range with validation
 * @param columnSpec Column specification like "A", "AB", "A:A", "B:Z", etc.
 * @returns Proper Excel column range like "A:A", "AB:AB", "A:A", "B:Z"
 * @throws Error if the column specification is invalid
 */
export function ensureColumnRangeWithValidation(columnSpec: string): string {
  // Trim whitespace and convert to uppercase
  const trimmed = columnSpec.trim().toUpperCase();

  if (!trimmed) {
    throw new Error("Column specification cannot be empty");
  }

  // If it already contains a colon, validate it's a proper range
  if (trimmed.includes(":")) {
    const parts = trimmed.split(":");
    if (parts.length !== 2) {
      throw new Error(
        `Invalid range format: ${columnSpec}. Expected format like 'A:Z'`,
      );
    }

    const [start, end] = parts;
    if (!/^[A-Z]+$/.test(start) || !/^[A-Z]+$/.test(end)) {
      throw new Error(`Invalid column letters in range: ${columnSpec}`);
    }

    return trimmed;
  }

  // If it's a single column letter/letters, convert to range format
  if (/^[A-Z]+$/.test(trimmed)) {
    // Additional validation: check if it's a valid Excel column (max is XFD)
    if (trimmed.length > 3) {
      throw new Error(
        `Invalid column: ${columnSpec}. Excel columns go up to XFD`,
      );
    }

    return `${trimmed}:${trimmed}`;
  }

  throw new Error(
    `Invalid column specification: ${columnSpec}. Expected format like 'A', 'AB', 'A:Z'`,
  );
}
