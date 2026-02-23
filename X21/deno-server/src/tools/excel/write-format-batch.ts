import { readFormatBatch } from "../../excel-actions/read-format-batch.ts";
import { writeFormatBatch } from "../../excel-actions/write-format-batch.ts";
import {
  OperationStatus,
  ReadFormatBatchOperation,
  ReadFormatBatchRequest,
  ReadFormatFinalResponseList,
  Tool,
  ToolNames,
  WriteFormatBatchRequest,
  WriteFormatResponse,
} from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";
import { WebSocketManager } from "../../services/websocket-manager.ts";
import { normalizeCurrencyNumberFormat } from "../../utils/number-format.ts";

const logger = createLogger("WriteFormatBatchTool");

export class WriteFormatBatchTool implements Tool {
  name = ToolNames.WRITE_FORMAT_BATCH;
  description =
    "⚠️ CRITICAL: This tool REQUIRES explicit user consent via COLLECT_INPUT before use. NEVER call this tool directly without first asking the user for permission using COLLECT_INPUT. Write formatting (colors, fonts, alignment, etc.) to one or more ranges in a single call. For a single range, pass one operation. Note: Number formatting should be applied via WRITE_VALUES_BATCH tool's formats array, not this tool.";

  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["operations"],
    properties: {
      operations: {
        type: "array",
        description: "List of write format operations to perform sequentially.",
        minItems: 1,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["worksheet", "range", "format"],
          properties: {
            worksheet: {
              type: "string",
              description: "Worksheet name (e.g., Sheet1)",
            },
            range: {
              type: "string",
              description: "Range to format (e.g., A1:C10)",
            },
            format: {
              type: "object",
              description: "Formatting options to apply to this range.",
              additionalProperties: false,
              properties: {
                bold: { type: "boolean", description: "Make text bold" },
                italic: { type: "boolean", description: "Make text italic" },
                underline: {
                  type: "boolean",
                  description: "Make text underlined",
                },
                fontColor: {
                  type: "string",
                  description: "Font color hex (e.g., #FF0000)",
                },
                backgroundColor: {
                  type: "string",
                  description:
                    "Background color hex (e.g., #FFFF00). Use 'none' to clear.",
                },
                alignment: {
                  type: "string",
                  description: "Text alignment: left, center, right",
                },
                numberFormat: {
                  type: "string",
                  description:
                    "Excel number format. Examples: '#,##0', '0.0%', '$#,##0;($#,##0);-', '#,##0\"€\";(#,##0\"€\")', '0.0x'.",
                },
                fontSize: {
                  type: "integer",
                  description: "Font size (e.g., 12, 14, 16)",
                },
                fontName: {
                  type: "string",
                  description: "Font family (e.g., Arial)",
                },
              },
            },
          },
        },
      },
      readOldFormats: {
        type: "boolean",
        description:
          "If true, read existing formats before writing (sequential COM-safe). Default false. May be overridden by DISABLE_FORMAT_REVERT env flag.",
      },
      collapseReadRanges: {
        type: "boolean",
        description:
          "If true, deduplicate pre-read calls by workbook/worksheet/range. Default true.",
      },
    },
  };

  async execute(params: WriteFormatBatchRequest): Promise<{
    writeResults: WriteFormatResponse[];
    oldFormats?: ReadFormatFinalResponseList[];
  }> {
    if (!params.operations || params.operations.length === 0) {
      throw new Error("write_format_batch requires at least one operation.");
    }

    const revertDisabled =
      (Deno.env.get("DISABLE_FORMAT_REVERT") ?? "true").toLowerCase() ===
        "true";
    const readOldFormats = !revertDisabled && params.readOldFormats !== false;

    const collapseReadRanges = params.collapseReadRanges !== false;
    const normalizedOperations = params.operations.map((op) => {
      const normalizedFormat = op.format
        ? {
          ...op.format,
          numberFormat: normalizeCurrencyNumberFormat(op.format.numberFormat),
        }
        : op.format;
      return {
        ...op,
        format: normalizedFormat,
      };
    });

    const totalOps = normalizedOperations.length;
    const statusWorkbook = normalizedOperations.find((op) => !!op.workbookName)
      ?.workbookName;

    const sendProgress = (
      status: OperationStatus,
      current: number,
      message?: string,
    ) => {
      if (!statusWorkbook) return;
      WebSocketManager.getInstance().sendStatus(
        statusWorkbook,
        status,
        message,
        { current, total: totalOps, unit: "ops" },
      );
    };

    logger.info(
      `write_format_batch start: operations=${params.operations.length}, readOldFormats=${readOldFormats} (revertDisabled=${revertDisabled}), collapseReadRanges=${collapseReadRanges}`,
    );

    const readCache = new Map<string, ReadFormatFinalResponseList>();
    const writeResults: WriteFormatResponse[] = [];
    const oldFormats: ReadFormatFinalResponseList[] = [];

    if (readOldFormats) {
      const operationsToRead:
        (ReadFormatBatchOperation & { cacheKey: string })[] = [];
      for (const op of normalizedOperations) {
        const cacheKey = `${op.workbookName}::${op.worksheet}::${op.range}`;
        if (collapseReadRanges && readCache.has(cacheKey)) continue;

        operationsToRead.push({
          cacheKey,
          workbookName: op.workbookName,
          worksheet: op.worksheet,
          range: op.range,
          propertiesToRead: undefined,
        });
      }

      if (operationsToRead.length > 0) {
        const batchRequest: ReadFormatBatchRequest = {
          operations: operationsToRead.map(({ cacheKey: _cacheKey, ...rest }) =>
            rest
          ),
        };

        logger.info(
          `write_format_batch pre-read start: uniqueReads=${operationsToRead.length}`,
          operationsToRead.map((o) =>
            `${o.workbookName}/${o.worksheet}/${o.range}`
          ),
        );

        const batchResults = await readFormatBatch(batchRequest);

        batchResults.forEach((result, idx) => {
          const op = operationsToRead[idx];
          if (!op) return;
          readCache.set(op.cacheKey, result);
        });

        logger.info(
          `write_format_batch pre-read done: cached=${readCache.size}`,
        );
      }
    }

    const batchResult = await writeFormatBatch({
      operations: normalizedOperations,
    });
    if (Array.isArray(batchResult.results)) {
      writeResults.push(...batchResult.results);
    }

    for (let idx = 0; idx < normalizedOperations.length; idx++) {
      const op = normalizedOperations[idx];
      const result = writeResults[idx];
      logger.info(
        `write_format_batch wrote -> workbook=${
          result?.workbookName ?? op.workbookName
        }, sheet=${result?.worksheet ?? op.worksheet}, success=${
          result?.success ?? false
        }, message="${result?.message ?? "No response"}"`,
      );
      sendProgress(
        "writing_excel_format",
        idx + 1,
        `Applying batch formatting (${totalOps} ops)`,
      );

      if (readOldFormats) {
        const cacheKey = `${op.workbookName}::${op.worksheet}::${op.range}`;
        const cached = readCache.get(cacheKey);
        if (cached) {
          oldFormats.push(cached);
        }
      }
    }

    logger.info(
      `write_format_batch complete: writes=${writeResults.length}, reads=${
        readOldFormats ? oldFormats.length : 0
      }`,
    );

    return {
      writeResults,
      oldFormats: readOldFormats ? oldFormats : undefined,
    };
  }
}
