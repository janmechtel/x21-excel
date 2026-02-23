import { readFormatBatch } from "../../excel-actions/read-format-batch.ts";
import {
  ReadFormatBatchRequest,
  ReadFormatFinalResponseList,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("ReadFormatBatchTool");

export class ReadFormatBatchTool implements Tool {
  name = ToolNames.READ_FORMAT_BATCH;
  description =
    "Read formatting for multiple ranges in a single call. For a single range, pass one operation.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["operations"],
    properties: {
      operations: {
        type: "array",
        description: "List of read format operations to perform sequentially.",
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
            propertiesToRead: {
              type: "array",
              description:
                "Optional list of format properties to read (e.g., bold, fontColor).",
              items: { type: "string" },
            },
          },
        },
      },
    },
  };

  async execute(params: ReadFormatBatchRequest): Promise<{
    results: ReadFormatFinalResponseList[];
  }> {
    if (!params.operations || params.operations.length === 0) {
      throw new Error("read_format_batch requires at least one operation.");
    }

    logger.info(
      `read_format_batch start: operations=${params.operations.length}`,
    );
    const results = await readFormatBatch(params);
    logger.info(`read_format_batch complete: results=${results.length}`);

    return { results };
  }
}
