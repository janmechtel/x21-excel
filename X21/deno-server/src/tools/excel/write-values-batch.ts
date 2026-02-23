import { readValuesBatch } from "../../excel-actions/read-values-batch.ts";
import { writeValuesBatch } from "../../excel-actions/write-values-batch.ts";
import { readFormatBatch } from "../../excel-actions/read-format-batch.ts";
import { writeFormatBatch } from "../../excel-actions/write-format-batch.ts";
import { WebSocketManager } from "../../services/websocket-manager.ts";
import { createLogger } from "../../utils/logger.ts";
import {
  columnToNumber,
  numberToColumn,
  parseRange,
  rangeToParsedString,
} from "../../utils/excel-range.ts";
import {
  ColumnWidthMode,
  OperationStatus,
  OperationStatusValues,
  PolicyAction,
  ReadFormatBatchRequest,
  ReadFormatFinalResponseList,
  ReadValuesBatchRequest,
  ReadValuesResponse,
  Tool,
  ToolNames,
  ToolParametersSchema,
  WriteFormatBatchRequest,
  WriteNumberFormatOperation,
  WriteValuesBatchRequest,
  WriteValuesBatchResponse,
  WriteValuesResponse,
} from "../../types/index.ts";
import { chunkArray, EXCEL_BATCH_MAX_OPS } from "../../utils/batching.ts";
import { normalizeCurrencyNumberFormat } from "../../utils/number-format.ts";
import { sanitizeWorkbookName } from "../../utils/workbook-name.ts";
import { UserPreferencesService } from "../../services/user-preferences.ts";

const logger = createLogger("WriteValuesBatchTool");
const COLUMN_WIDTH_MODE_KEY = "column_width_mode";
const DEFAULT_COLUMN_WIDTH_MODE: ColumnWidthMode = "smart";
const COLUMN_WIDTH_MODE_VALUES = new Set<ColumnWidthMode>([
  "always",
  "smart",
  "never",
]);
const NUMBER_FORMAT_DESCRIPTION = [
  "Excel number format applied to the entire range.",
  "CRITICAL for attachments (PDFs/images): preserve the EXACT format from the source.",
  "Observe negative style (parentheses vs minus), currency symbol, placement/spacing, decimal places, and thousands separators.",
  'Prefer raw numbers with numberFormat over embedding currency/unit text in values (avoid strings like "100 RUB").',
  "Combined currency + negative examples: Source $(123) → '$#,##0;($#,##0)'; Source -€1,234 → '#,##0\"€\";-#,##0\"€\"'.",
  "Validation: after writing, the displayed format must match the source exactly.",
  'Default format rules: Years as text ("2024"); Currency: $#,##0 with units in headers; Zeros as "-": $#,##0;($#,##0);-; Percentages: 0.0%; Multiples: 0.0x; Negative numbers: parentheses (123).',
  "Examples: '#,##0', '0.0%', '$#,##0;($#,##0);-', '#,##0\"€\";(#,##0\"€\")', '0.0x'.",
].join(" ");

export class WriteValuesBatchTool implements Tool {
  name = ToolNames.WRITE_VALUES_BATCH;

  description = `Write values to Excel ranges in one call.

Usage:
- Use this for ALL writes. For a single range, include one operation.
- For multiple disjoint ranges, include multiple operations (not separate calls).
- Values MUST be a rectangular 2D array (every row the same length).
- Use null or "" for blanks.
- Number formatting (formats[numberFormat]) is applied with values and does NOT require
  separate user consent.
- Old values are captured for revert and returned (memory grows with cell count).
- Operations are automatically chunked into batches of up to ${EXCEL_BATCH_MAX_OPS}.

Formula rules (ALWAYS):
- Formulas are written via Excel COM Range.Formula (NOT FormulaLocal).
- Use English/invariant formulas (e.g., =IF(...), =SUM(...)).
- Use "." as decimal separator and "," as argument separator.
- Never use thousands separators inside formulas (write 1000 not 1,000).
- Avoid TEXT() unless the output must be a string. Prefer numeric/date values.
- TEXT() format_text is locale-aware. Use excelContext locale fields when needed:
  - excelContext.listSeparator
  - excelContext.decimalSeparator
  - excelContext.thousandsSeparator
  - excelContext.dateLanguage:
    - German: "TT.MM.JJJJ"
    - English: "mm/dd/yyyy"`;

  strict = true;

  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["operations"],
    properties: {
      operations: {
        type: "array",
        description: "List of write operations to perform sequentially.",
        minItems: 1,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["worksheet", "range", "values"],
          properties: {
            workbookName: {
              type: "string",
              description:
                "Workbook name (defaults to the active workbook). Cross-workbook writes are blocked.",
            },
            worksheet: {
              type: "string",
              description: "Worksheet name (e.g., Sheet1).",
            },
            range: {
              type: "string",
              description: "Range to write (e.g., A1:C10).",
            },
            values: {
              type: "array",
              description:
                'Rectangular 2D array where each sub-array is a row and every row has the same length. Use null or "" for blanks. Range must match array dimensions.',
              items: {
                type: "array",
                items: {
                  type: ["string", "number", "boolean", "null"],
                },
              },
            },
          },
          examples: [
            // 1) Simple header + one data row
            {
              worksheet: "Sheet1",
              range: "A1:B2",
              values: [
                ["Name", "Age"],
                ["John", 25],
              ],
            },
            // 2) Larger rectangular block with mixed types
            {
              worksheet: "Sales",
              range: "A1:D3",
              values: [
                ["Product", "Price", "Stock", "Category"],
                ["Widget A", 19.99, 150, "Electronics"],
                ["Widget B", 29.99, 75, "Electronics"],
              ],
            },
            // 3) Using null for blanks inside the rectangle
            {
              worksheet: "Data",
              range: "A1:C4",
              values: [
                ["ID", "Name", "Email"],
                [1, "Alice", "alice@example.com"],
                [2, "Bob", "bob@example.com"],
                [3, "Charlie", null],
              ],
            },
            // 4) Explicit workbookName example
            {
              workbookName: "Budget2025.xlsx",
              worksheet: "Plan",
              range: "A1:C2",
              values: [
                ["Month", "Revenue", "Expenses"],
                ["Jan", 120000, 80000],
              ],
            },
          ],
        },
      },

      formats: {
        type: "array",
        description:
          "List of formatting overlays applied after all values are written.",
        minItems: 0,
        items: {
          type: "object",
          additionalProperties: false,
          required: ["worksheet", "range"],
          properties: {
            workbookName: {
              type: "string",
              description:
                "Workbook name (defaults to the active workbook). Must match value writes if provided.",
            },
            worksheet: {
              type: "string",
              description: "Worksheet name (e.g., Sheet1).",
            },
            range: {
              type: "string",
              description: "Target range for formatting (e.g., B2:B10).",
            },
            numberFormat: {
              type: "string",
              description: NUMBER_FORMAT_DESCRIPTION,
            },
          },
        },
        examples: [
          {
            worksheet: "Sales",
            range: "B2:B10",
            numberFormat: "\\$#,##0.00",
          },
          {
            worksheet: "Sales",
            range: "C2:C10",
            numberFormat: '#,##0"€"',
          },
          {
            worksheet: "Sales",
            range: "A2:A10",
            numberFormat: "mm/dd/yyyy",
          },
          {
            worksheet: "Financial",
            range: "E2:E5",
            numberFormat: "0.0%",
          },
        ],
      },
    },
    examples: [
      // 1) Simple write, no formatting
      {
        operations: [
          {
            worksheet: "Sheet1",
            range: "A1:B2",
            values: [
              ["Name", "Age"],
              ["John", 25],
            ],
          },
        ],
      },

      // 2) Write + price formatting on a subset
      {
        operations: [
          {
            worksheet: "Sales",
            range: "A1:D3",
            values: [
              ["Product", "Price", "Stock", "Category"],
              ["Widget A", 19.99, 150, "Electronics"],
              ["Widget B", 29.99, 75, "Electronics"],
            ],
          },
        ],
        formats: [
          {
            worksheet: "Sales",
            range: "B2:B3",
            numberFormat: "\\$#,##0.00",
          },
        ],
      },

      // 3) Multiple sheets + numeric/date formats
      {
        operations: [
          {
            worksheet: "Data",
            range: "A1:C4",
            values: [
              ["ID", "Name", "Email"],
              [1, "Alice", "alice@example.com"],
              [2, "Bob", "bob@example.com"],
              [3, "Charlie", null],
            ],
          },
          {
            worksheet: "Summary",
            range: "A1:B2",
            values: [
              ["Total Sales", 50000],
              ["Total Units", 1000],
            ],
          },
        ],
        formats: [
          {
            worksheet: "Data",
            range: "A2:A4",
            numberFormat: "#,##0",
          },
          {
            worksheet: "Summary",
            range: "B1:B1",
            numberFormat: "\\$#,##0",
          },
          {
            worksheet: "Summary",
            range: "B2:B2",
            numberFormat: "#,##0",
          },
        ],
      },

      // 4) Explicit workbookName example
      {
        operations: [
          {
            workbookName: "Budget2025.xlsx",
            worksheet: "Plan",
            range: "A1:C2",
            values: [
              ["Month", "Revenue", "Expenses"],
              ["Jan", 120000, 80000],
            ],
          },
        ],
        formats: [
          {
            workbookName: "Budget2025.xlsx",
            worksheet: "Plan",
            range: "B2:C2",
            numberFormat: "#,##0",
          },
        ],
      },
    ],
  } as ToolParametersSchema;

  async execute(
    params: WriteValuesBatchRequest,
  ): Promise<WriteValuesBatchResponse> {
    if (!params.operations || params.operations.length === 0) {
      throw new Error("write_values_batch requires at least one operation.");
    }

    const totalOps = params.operations.length;
    const rawActiveWorkbook = params.activeWorkbookName ||
      params.operations[0]?.workbookName;
    const activeWorkbook = sanitizeWorkbookName(rawActiveWorkbook) ||
      rawActiveWorkbook;
    if (!activeWorkbook) {
      throw new Error("No active workbook provided for write_values_batch.");
    }

    const sanitizedOperations = params.operations.map((op) => {
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

    const statusWorkbook = sanitizeWorkbookName(
      sanitizedOperations.find((op) => !!op.workbookName)?.workbookName,
    ) ||
      sanitizedOperations.find((op) => !!op.workbookName)?.workbookName ||
      activeWorkbook;

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

    const revertDisabled =
      (Deno.env.get("DISABLE_FORMAT_REVERT") ?? "true").toLowerCase() ===
        "true";

    const formats = Array.isArray(params.formats) ? params.formats : [];
    const hasFormats = formats.length > 0;
    const columnWidthMode = resolveColumnWidthModePreference();

    logger.info(
      `write_values_batch start: operations=${totalOps}, formats=${formats.length}, revertDisabled=${revertDisabled}`,
    );

    const sanitizedFormats: WriteNumberFormatOperation[] = formats.map(
      (fmt) => {
        const rawWorkbook = fmt.workbookName;
        const workbookName = sanitizeWorkbookName(rawWorkbook) ||
          rawWorkbook ||
          activeWorkbook;
        return {
          ...fmt,
          workbookName,
        };
      },
    );

    const mismatchedFormat = sanitizedFormats.find((fmt) =>
      fmt.workbookName && !isSameWorkbook(fmt.workbookName, activeWorkbook)
    );
    if (mismatchedFormat) {
      const requestedWorkbook = mismatchedFormat.workbookName;
      const blockedOperations = sanitizedOperations.map((op) => ({
        worksheet: op.worksheet,
        range: op.range,
        requestedWorkbook: op.requestedWorkbook ||
          sanitizeWorkbookName(op.workbookName) ||
          op.workbookName ||
          requestedWorkbook,
        resolvedWorkbook: null,
        policyAction: "blocked" as PolicyAction,
        status: "error" as const,
        oldValues: null,
        newValues: null,
        errorCode: "CROSS_WORKBOOK_WRITE_BLOCKED",
        errorMessage:
          `Writes are restricted to the active workbook '${activeWorkbook}'. Requested '${requestedWorkbook}'.`,
      }));
      return {
        tool: ToolNames.WRITE_VALUES_BATCH,
        success: false,
        message: "Blocked cross-workbook format write request.",
        activeWorkbook,
        requestedWorkbook,
        resolvedWorkbook: null,
        policyAction: "blocked",
        operations: blockedOperations,
        results: blockedOperations.map((op) => ({
          success: false,
          message: op.errorMessage || "blocked",
        })),
      };
    }

    const opResults: WriteValuesBatchResponse["operations"] = [];
    const normalizedOps: WriteValuesBatchRequest["operations"] = [];
    for (const op of sanitizedOperations) {
      const requestedWorkbook = op.requestedWorkbook ||
        sanitizeWorkbookName(op.workbookName) ||
        op.workbookName;
      const resolvedWorkbook = sanitizeWorkbookName(op.workbookName) ||
        activeWorkbook;
      if (
        requestedWorkbook &&
        !isSameWorkbook(requestedWorkbook, activeWorkbook)
      ) {
        opResults.push({
          worksheet: op.worksheet,
          range: op.range,
          requestedWorkbook,
          resolvedWorkbook: null,
          policyAction: "blocked",
          status: "error",
          oldValues: null,
          newValues: null,
          errorCode: "CROSS_WORKBOOK_WRITE_BLOCKED",
          errorMessage:
            `Writes are restricted to the active workbook '${activeWorkbook}'. Requested '${requestedWorkbook}'.`,
        });
        continue;
      }
      const { normalizedRange, normalizedValues } =
        normalizeRangeAndValuesForOperation(op.range, op.values);
      normalizedOps.push({
        ...op,
        range: normalizedRange,
        values: normalizedValues as any,
        requestedWorkbook,
        workbookName: resolvedWorkbook,
      });
      logger.info(
        `write_values_batch apply -> workbook=${op.workbookName}, sheet=${op.worksheet}, range=${normalizedRange}`,
      );
    }

    if (opResults.some((r) => r.policyAction === "blocked")) {
      return {
        tool: ToolNames.WRITE_VALUES_BATCH,
        success: false,
        message: "Blocked cross-workbook write request.",
        activeWorkbook,
        requestedWorkbook: opResults[0]?.requestedWorkbook,
        resolvedWorkbook: null,
        policyAction: "blocked",
        operations: opResults,
        results: opResults.map((r) => ({
          success: false,
          message: r.errorMessage || "blocked",
        })),
      };
    }

    // Read old formats before writing if formats will be applied and reverting is enabled
    const oldFormats: ReadFormatFinalResponseList[] = [];
    if (hasFormats && !revertDisabled) {
      const defaultWorkbookName = activeWorkbook;
      const formatReadRequest: ReadFormatBatchRequest = {
        operations: sanitizedFormats.map((fmt) => ({
          workbookName: fmt.workbookName || defaultWorkbookName,
          worksheet: fmt.worksheet,
          range: fmt.range,
          propertiesToRead: undefined,
        })),
      };
      const formatReadResults = await readFormatBatch(formatReadRequest);
      oldFormats.push(...formatReadResults);
      logger.info(
        `write_values_batch pre-read formats: ${formatReadResults.length} ranges`,
      );
      formatReadResults.forEach((formatList, idx) => {
        logger.info(
          `write_values_batch pre-read format[${idx}]: ${formatList.length} format groups`,
          formatList.map((fg) => ({
            ranges: fg.ranges,
            formatKeys: Object.keys(fg.format || {}),
            numberFormat: fg.format?.numberFormat,
          })),
        );
      });
    }

    // Chunk and process operations
    const batches = chunkArray(normalizedOps);
    const oldValues: ReadValuesResponse[] = [];
    const writeResults: WriteValuesResponse[] = [];
    let overallSuccess = true;

    for (const [index, batchOps] of batches.entries()) {
      logger.info(
        `write_values_batch pre-read start: batch=${
          index + 1
        }/${batches.length}, ops=${batchOps.length}`,
      );
      const readRequest: ReadValuesBatchRequest = {
        operations: batchOps.map((op) => ({
          workbookName: op.workbookName,
          worksheet: op.worksheet,
          range: op.range,
        })),
      };
      const readResult = await readValuesBatch(readRequest);
      if (readResult.results.length !== batchOps.length) {
        throw new Error(
          `read_values_batch returned ${readResult.results.length} results for ${batchOps.length} operations (batch ${
            index + 1
          })`,
        );
      }
      oldValues.push(
        ...readResult.results.map((res) => ({ cellValues: res.cellValues })),
      );
      logger.info(
        `write_values_batch pre-read done: batch=${
          index + 1
        }, reads=${readResult.results.length}`,
      );

      const batchResult: WriteValuesBatchResponse = await writeValuesBatch({
        operations: batchOps,
        columnWidthMode,
      });
      const batchWriteResults = Array.isArray(batchResult?.results)
        ? batchResult.results
        : [];
      writeResults.push(...batchWriteResults);
      overallSuccess = overallSuccess && !!batchResult?.success;

      const processed = writeResults.length;
      sendProgress(
        OperationStatusValues.WRITING_EXCEL,
        processed,
        `Writing values batch ${
          index + 1
        }/${batches.length} (${processed}/${totalOps})`,
      );

      logger.info(
        `write_values_batch batch complete: batch=${
          index + 1
        }, size=${batchOps.length}, processed=${processed}/${totalOps}, success=${!!batchResult
          ?.success}`,
      );
    }

    // Apply formats after all values are written
    if (hasFormats) {
      logger.info(
        `write_values_batch applying formats: ${sanitizedFormats.length} operations`,
      );
      const formatWriteRequest: WriteFormatBatchRequest = {
        operations: sanitizedFormats.map((fmt) => ({
          workbookName: fmt.workbookName,
          worksheet: fmt.worksheet,
          range: fmt.range,
          format: {
            numberFormat: normalizeCurrencyNumberFormat(fmt.numberFormat),
          },
        })),
        readOldFormats: false, // Already read above
        collapseReadRanges: true,
      };

      const formatResult = await writeFormatBatch(formatWriteRequest);
      if (!formatResult.success) {
        logger.warn(
          `write_values_batch format application had errors: ${formatResult.message}`,
        );
        overallSuccess = false;
      } else {
        logger.info(
          `write_values_batch formats applied: ${
            formatResult.results?.length ?? 0
          } operations`,
        );
      }
    }

    const writesSuccess = writeResults.every((r) => r.success);
    const success = overallSuccess && writesSuccess;

    const operationsResults: WriteValuesBatchResponse["operations"] =
      normalizedOps
        .map((op, idx) => ({
          worksheet: op.worksheet,
          range: op.range,
          requestedWorkbook: op.requestedWorkbook,
          resolvedWorkbook: op.workbookName,
          policyAction: "wrote" as PolicyAction,
          status: writeResults[idx]?.success ? "success" : "error",
          oldValues: (oldValues[idx]?.cellValues as any) ?? null,
          newValues: op.values as any[][],
          errorCode: writeResults[idx]?.success ? undefined : "WRITE_FAILED",
          errorMessage: writeResults[idx]?.message,
        }));

    const message = success
      ? `Completed ${writeResults.length} write operations`
      : `Completed ${writeResults.length} write operations with errors`;

    logger.info(
      `write_values_batch complete: writes=${writeResults.length}, reads=${oldValues.length}, formats=${
        !revertDisabled ? oldFormats.length : 0
      }, batches=${batches.length}, success=${success}`,
    );

    return {
      tool: ToolNames.WRITE_VALUES_BATCH,
      success,
      message,
      activeWorkbook,
      requestedWorkbook: operationsResults[0]?.requestedWorkbook,
      resolvedWorkbook: activeWorkbook,
      policyAction: success ? "wrote" : "partial",
      operations: operationsResults,
      results: writeResults,
      applied: operationsResults.filter((op) => op.status === "success").length,
      batches: batches.length,
      oldValues,
      oldFormats: hasFormats && !revertDisabled ? oldFormats : undefined,
    };
  }
}

function resolveColumnWidthModePreference(): ColumnWidthMode {
  const prefs = UserPreferencesService.getInstance();
  const raw = prefs.getPreference(COLUMN_WIDTH_MODE_KEY);
  return normalizeColumnWidthMode(raw);
}

function normalizeColumnWidthMode(raw: string | null): ColumnWidthMode {
  if (!raw) {
    return DEFAULT_COLUMN_WIDTH_MODE;
  }
  const normalized = raw.toLowerCase();
  if (COLUMN_WIDTH_MODE_VALUES.has(normalized as ColumnWidthMode)) {
    return normalized as ColumnWidthMode;
  }
  return DEFAULT_COLUMN_WIDTH_MODE;
}

function normalizeRangeAndValuesForOperation(
  range: string,
  values: unknown,
): { normalizedRange: string; normalizedValues: unknown } {
  const safeRange = (range || "").trim().toUpperCase();
  if (!safeRange) {
    throw new Error("write_values_batch requires a range for each operation.");
  }

  if (!Array.isArray(values) || values.length === 0) {
    throw new Error(
      "write_values_batch requires a non-empty 2D values array for each operation.",
    );
  }

  const rowCount = values.length;
  const colCount = Math.max(
    ...values.map((row) => Array.isArray(row) ? row.length : 0),
  );

  if (colCount === 0) {
    throw new Error(
      "write_values_batch requires each row to have at least one value.",
    );
  }

  const normalizedValues: unknown[] = [];
  for (const [index, row] of values.entries()) {
    if (!Array.isArray(row)) {
      throw new Error(
        `write_values_batch row ${index + 1} is not an array.`,
      );
    }
    if (row.length > colCount) {
      throw new Error(
        `write_values_batch row ${index + 1} length (${
          Array.isArray(row) ? row.length : 0
        }) does not match expected column count (${colCount}).`,
      );
    }

    if (row.length !== colCount) {
      logger.warn("Padding write_values_batch row with blanks", {
        row: index + 1,
        currentLength: row.length,
        expectedLength: colCount,
      });
    }

    const padded = row.concat(Array(colCount - row.length).fill(""));
    normalizedValues.push(padded);
  }

  const parsed = parseRange(safeRange);
  const startColNum = columnToNumber(parsed.startColumn);
  const endColNum = startColNum + colCount - 1;
  const endRow = parsed.startRow + rowCount - 1;
  const normalizedParsed = {
    startColumn: parsed.startColumn,
    startRow: parsed.startRow,
    endColumn: numberToColumn(endColNum),
    endRow,
  };

  const normalizedRange = rangeToParsedString(normalizedParsed);
  if (normalizedRange !== safeRange) {
    logger.warn("Adjusted write_values_batch range to match values", {
      providedRange: range,
      normalizedRange,
      rowCount,
      colCount,
    });
  }

  return { normalizedRange, normalizedValues };
}

function equalsIgnoreCase(a?: string, b?: string): boolean {
  return (a ?? "").toLowerCase() === (b ?? "").toLowerCase();
}

function isSameWorkbook(a?: string, b?: string): boolean {
  if (!a || !b) return false;
  if (equalsIgnoreCase(a, b)) return true;

  const baseA = a.toLowerCase().split(/[/\\]/).pop() || "";
  const baseB = b.toLowerCase().split(/[/\\]/).pop() || "";
  return baseA.length > 0 && baseA === baseB;
}
