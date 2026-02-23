import { copyPaste } from "../../excel-actions/copy-paste.ts";
import { readFormat } from "../../excel-actions/read-format.ts";
import { readValues } from "../../excel-actions/read-values.ts";
import {
  CopyPasteRequest,
  ReadFormatFinalResponseList,
  ReadFormatRequest,
  ReadValuesRequest,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("CopyPasteTool");

export class CopyPasteTool implements Tool {
  name = ToolNames.COPY_PASTE;
  description =
    "Copy a range and paste it elsewhere (like Ctrl+C/Ctrl+V). Supports paste options (all/values/formulas/formats) and optional insert mode (shift_right or shift_down) to insert cells instead of overwriting.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: [
      "sourceWorksheet",
      "sourceRange",
      "destinationWorksheet",
      "destinationRange",
    ],
    properties: {
      workbookName: {
        type: "string",
        description:
          "Target workbook (defaults to the active workbook; writes are restricted to the active workbook).",
      },
      sourceWorksheet: {
        type: "string",
        description: "Worksheet to copy from.",
      },
      sourceRange: {
        type: "string",
        description: "Range to copy (e.g., A1:C10).",
      },
      destinationWorksheet: {
        type: "string",
        description: "Worksheet to paste into.",
      },
      destinationRange: {
        type: "string",
        description:
          "Top-left cell or target range for the paste (will resize to the source size).",
      },
      pasteType: {
        type: "string",
        enum: [
          "all",
          "values",
          "formats",
          "formulas",
          "formulas_and_number_formats",
          "values_and_number_formats",
          "column_widths",
        ],
        description:
          "What to paste. Defaults to all. Use values/formulas/formats for paste-special behavior.",
      },
      insertMode: {
        type: "string",
        enum: ["none", "shift_right", "shift_down"],
        description:
          "Use shift_right or shift_down to insert cells (like Insert Copied Cells) instead of overwriting.",
      },
      skipBlanks: {
        type: "boolean",
        description: "Skip blank cells when pasting (default: false).",
      },
      transpose: {
        type: "boolean",
        description: "Transpose rows/columns while pasting (default: false).",
      },
      includeColumnWidths: {
        type: "boolean",
        description: "Also copy column widths from the source range.",
      },
    },
  };

  async execute(params: CopyPasteRequest): Promise<any> {
    logger.info("Executing copy_paste", {
      workbookName: params.workbookName,
      sourceWorksheet: params.sourceWorksheet,
      destinationWorksheet: params.destinationWorksheet,
      sourceRange: params.sourceRange,
      destinationRange: params.destinationRange,
      pasteType: params.pasteType,
      insertMode: params.insertMode,
    });

    const destinationWorksheet = params.destinationWorksheet ||
      params.sourceWorksheet;
    const formatReadRange = getTargetRange(
      params.sourceRange,
      params.destinationRange,
    );
    let oldFormats: ReadFormatFinalResponseList | undefined;

    if (formatReadRange && params.workbookName && destinationWorksheet) {
      const formatParams: ReadFormatRequest = {
        workbookName: params.workbookName,
        worksheet: destinationWorksheet,
        range: formatReadRange,
      };
      try {
        oldFormats = await readFormat(formatParams);
      } catch (error) {
        logger.warn("Failed to read old formats before copy_paste", {
          error: (error as Error)?.message,
          worksheet: formatParams.worksheet,
          range: formatParams.range,
        });
      }
    }

    const copyPasteResult = await copyPaste(params);

    let newValues: any = undefined;
    const readParams: ReadValuesRequest | null =
      copyPasteResult.success && (copyPasteResult.destinationRange ||
          params.destinationRange)
        ? {
          workbookName: params.workbookName,
          worksheet: params.destinationWorksheet || params.sourceWorksheet,
          range: copyPasteResult.destinationRange || params.destinationRange,
        }
        : null;

    if (readParams) {
      try {
        newValues = await readValues(readParams);
      } catch (error) {
        logger.warn("Failed to read new values after copy_paste", {
          error: (error as Error)?.message,
          worksheet: readParams.worksheet,
          range: readParams.range,
        });
      }
    }

    return {
      copyPasteResult,
      oldValues: copyPasteResult.oldValues,
      oldFormats,
      newValues,
    };
  }
}

function getTargetRange(
  sourceRange: string,
  destinationRange: string,
): string | null {
  const sourceBounds = parseRangeBounds(sourceRange);
  if (!sourceBounds) return null;
  const destinationTopLeft = parseCellAddress(
    destinationRange.split(":")[0],
  );
  if (!destinationTopLeft) return null;

  const rowCount = sourceBounds.endRow - sourceBounds.startRow + 1;
  const colCount = sourceBounds.endCol - sourceBounds.startCol + 1;

  const endRow = destinationTopLeft.row + rowCount - 1;
  const endCol = destinationTopLeft.col + colCount - 1;
  const endColName = numberToColumn(endCol);

  const startColName = numberToColumn(destinationTopLeft.col);
  const startCell = `${startColName}${destinationTopLeft.row}`;
  const endCell = `${endColName}${endRow}`;

  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function parseRangeBounds(range: string): {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} | null {
  const parts = range.includes(":") ? range.split(":") : [range, range];
  if (parts.length !== 2) return null;

  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  if (!start || !end) return null;

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

function parseCellAddress(
  address: string,
): { row: number; col: number } | null {
  const normalized = normalizeCellAddress(address);
  if (!normalized) return null;
  const match = normalized.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const col = columnToNumber(match[1]);
  const row = parseInt(match[2], 10);
  if (!row || !col) return null;
  return { row, col };
}

function normalizeCellAddress(address: string): string | null {
  if (!address) return null;
  const trimmed = address.replace(/\$/g, "").trim();
  const match = trimmed.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;
  return `${match[1].toUpperCase()}${match[2]}`;
}

function columnToNumber(col: string): number {
  return col.split("").reduce(
    (acc, char) => acc * 26 + (char.charCodeAt(0) - 64),
    0,
  );
}

function numberToColumn(num: number): string {
  let result = "";
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}
