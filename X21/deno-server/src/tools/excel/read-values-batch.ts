import {
  ReadValuesBatchRequest,
  ReadValuesBatchResult,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { readValuesBatch } from "../../excel-actions/read-values-batch.ts";
import { columnToNumber, parseRange } from "../../utils/excel-range.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("ReadValuesBatchTool");
const MAX_CELLS_PER_CALL = 1000;
const WARN_CELLS_PER_CALL = Math.round(MAX_CELLS_PER_CALL * 1.5);

function estimateCellCount(a1Range: string): number {
  const parsed = parseRange(a1Range);
  const startCol = columnToNumber(parsed.startColumn);
  const endCol = columnToNumber(parsed.endColumn);
  const rowCount = Math.abs(parsed.endRow - parsed.startRow) + 1;
  const colCount = Math.abs(endCol - startCol) + 1;
  return rowCount * colCount;
}

export class ReadValuesBatchTool implements Tool {
  name = ToolNames.READ_VALUES_BATCH;
  description = [
    "Read values and formulas from multiple ranges in a single call.",
    "IMPORTANT: keep total cells read per call (sum of all operations, rows×cols)",
    `to ~${MAX_CELLS_PER_CALL} to avoid token limits; for larger reads, split into`,
    "multiple calls.",
  ].join(" ");
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["operations"],
    properties: {
      operations: {
        type: "array",
        description: [
          "List of read operations to perform sequentially.",
          "IMPORTANT: keep total cells read per call (sum of all operations, rows×cols)",
          `to ~${MAX_CELLS_PER_CALL}; for larger reads, split into multiple calls.`,
        ].join(" "),
        minItems: 1,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["worksheet", "range"],
          properties: {
            workbookName: {
              type: "string",
              description:
                "Optional workbook name. Defaults to the active workbook when omitted.",
            },
            worksheet: {
              type: "string",
              description: "Worksheet name (e.g., Sheet1)",
            },
            range: {
              type: "string",
              description: "Range to read (e.g., A1:C10)",
            },
          },
        },
      },
    },
  };

  async execute(params: ReadValuesBatchRequest): Promise<{
    results: ReadValuesBatchResult[];
  }> {
    if (!params.operations || params.operations.length === 0) {
      throw new Error("read_values_batch requires at least one operation.");
    }

    const totalCells = params.operations.reduce((sum, op) => {
      try {
        return sum + estimateCellCount(op.range);
      } catch {
        // If parsing fails, don't block execution here; Excel API will error with details.
        return sum;
      }
    }, 0);

    if (totalCells > WARN_CELLS_PER_CALL) {
      logger.warn("read_values_batch request exceeds recommended cell limit", {
        totalCells,
        recommendedLimit: WARN_CELLS_PER_CALL,
        operations: params.operations.length,
      });
    }

    logger.info(
      `read_values_batch start: operations=${params.operations.length}`,
    );
    const result = await readValuesBatch(params);
    logger.info(
      `read_values_batch complete: results=${result.results.length}, success=${result.success}`,
    );

    return { results: result.results };
  }
}
