/// <reference lib="deno.ns" />
import { stateManager, ToolChangeInterface } from "../state/state-manager.ts";
import { createLogger } from "../utils/logger.ts";
import { writeValuesBatch } from "../excel-actions/write-values-batch.ts";
import { removeSheets } from "../excel-actions/remove-sheets.ts";
import { removeRows } from "../excel-actions/remove-rows.ts";
import { removeColumns } from "../excel-actions/remove-columns.ts";
import { addRows } from "../excel-actions/add-rows.ts";
import { addColumns } from "../excel-actions/add-columns.ts";
import { writeFormatBatch } from "../excel-actions/write-format-batch.ts";
import { deleteCells } from "../excel-actions/delete-cells.ts";
import {
  RevertOperationKeys,
  ToolNames,
  WriteFormatRequest,
  WriteValuesRequest,
} from "../types/index.ts";

const logger = createLogger("RevertService");
const TOOLS_WITHOUT_REVERT: string[] = [
  ToolNames.READ_VALUES_BATCH,
  ToolNames.READ_FORMAT_BATCH,
  ToolNames.VBA_CREATE,
  ToolNames.VBA_READ,
  ToolNames.VBA_UPDATE,
  ToolNames.MERGE_FILES,
];

export async function revertSingleToolChange(
  change: ToolChangeInterface,
  isToolPending: boolean,
): Promise<void> {
  logger.info("Starting revert for tool change", {
    toolId: change.toolId,
    toolName: change.toolName,
    workbookName: change.workbookName,
    isToolPending,
    hasInputDataRevert: !!change.inputDataRevert,
    revertKeys: change.inputDataRevert
      ? Object.keys(change.inputDataRevert)
      : [],
  });

  switch (change.toolName) {
    case ToolNames.WRITE_VALUES_BATCH:
      await revertWriteValues(change);
      break;
    case ToolNames.DRAG_FORMULA:
      await revertDragFormula(change);
      break;
    case ToolNames.ADD_SHEETS:
      await revertAddSheets(change);
      break;
    case ToolNames.WRITE_FORMAT_BATCH:
      await revertWriteFormatBatch(change);
      break;
    case ToolNames.ADD_ROWS:
      await revertAddRows(change);
      break;
    case ToolNames.REMOVE_ROWS:
      await revertRemoveRows(change);
      break;
    case ToolNames.ADD_COLUMNS:
      await revertAddColumns(change);
      break;
    case ToolNames.REMOVE_COLUMNS:
      await revertRemoveColumns(change);
      break;
    case ToolNames.COPY_PASTE:
      await revertCopyPaste(change);
      break;
    default:
      if (TOOLS_WITHOUT_REVERT.includes(change.toolName)) {
        logger.info(
          `Skipping revert for ${change.toolName} - no changes to revert.`,
        );
      } else {
        logger.warn(`Unknown tool type for revert: ${change.toolName}`);
      }
      break;
  }

  const updatedChange = {
    ...change,
    applied: false,
    pending: isToolPending,
  };
  stateManager.updateToolChange(
    change.workbookName,
    change.toolId,
    updatedChange,
  );
  return;
}

function isNotFoundError(message?: string): boolean {
  return !!message && message.toLowerCase().includes("not found");
}

async function executeWriteValuesBatch(
  requests: WriteValuesRequest[],
  failurePrefix: string,
): Promise<void> {
  if (requests.length === 0) {
    throw new Error("Missing required data for write_values revert");
  }

  const batchResult = await writeValuesBatch({
    operations: requests.map((request) => ({ ...request })),
  });

  const results = Array.isArray(batchResult?.results)
    ? batchResult.results
    : [];

  if (results.length === 0) {
    throw new Error(`${failurePrefix}: empty batch result`);
  }

  const failures: string[] = [];
  let notFoundCount = 0;

  for (let i = 0; i < requests.length; i++) {
    const result = results[i];
    const request = requests[i];
    const location = request?.worksheet && request?.range
      ? `${request.worksheet}!${request.range}`
      : request?.range || request?.worksheet || `op ${i + 1}`;

    if (!result) {
      failures.push(`missing result for ${location}`);
      continue;
    }

    if (!result.success) {
      const message = result.message || "Unknown error";
      // Check if the error is due to worksheet not found - this is not critical for revert
      if (isNotFoundError(message)) {
        notFoundCount++;
        logger.warn(
          `Worksheet not found during revert - skipping: ${message}`,
          { location, index: i },
        );
        continue;
      }
      failures.push(`${location}: ${message}`);
    }
  }

  if (notFoundCount > 0) {
    logger.warn("Skipped revert entries for missing worksheets", {
      notFoundCount,
      totalOperations: requests.length,
    });
  }

  if (failures.length > 0) {
    throw new Error(`${failurePrefix}: ${failures[0]}`);
  }
}

function normalizeWriteValuesRequests(
  writeValuesParams:
    | WriteValuesRequest
    | WriteValuesRequest[]
    | WriteValuesRequest[][]
    | undefined,
): WriteValuesRequest[] {
  if (!writeValuesParams) {
    return [];
  }

  if (Array.isArray(writeValuesParams)) {
    return (writeValuesParams as Array<
      WriteValuesRequest | WriteValuesRequest[]
    >)
      .flat(2)
      .filter(Boolean) as WriteValuesRequest[];
  }

  return [writeValuesParams];
}

async function revertWriteValues(change: ToolChangeInterface): Promise<void> {
  logger.info("change inputDataRevert", change.inputDataRevert);

  const inputDataRevert = change.inputDataRevert;
  if (
    !inputDataRevert || !inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH]
  ) {
    throw new Error("Missing required data for write_values revert");
  }

  const writeValuesParams =
    inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH];
  const requests = normalizeWriteValuesRequests(writeValuesParams);
  await executeWriteValuesBatch(requests, "Failed to revert write_values");
}

async function revertWriteFormatBatch(
  change: ToolChangeInterface,
): Promise<void> {
  logger.info("change inputDataRevert", change.inputDataRevert);

  const inputDataRevert: Record<string, WriteFormatRequest> =
    change.inputDataRevert;
  if (!inputDataRevert) {
    throw new Error("Missing required data for write_format_batch revert");
  }

  const revertEntries = Object.entries(inputDataRevert);
  if (revertEntries.length === 0) {
    throw new Error("Missing required data for write_format_batch revert");
  }

  logger.info("Prepared write_format_batch revert entries", {
    entryCount: revertEntries.length,
    entryKeys: revertEntries.map(([key]) => key),
  });

  const operations = revertEntries.map(([, inputParams]) => ({
    workbookName: inputParams?.workbookName,
    worksheet: inputParams?.worksheet,
    range: inputParams?.range,
    format: inputParams?.format,
  }));

  const revertResult = await writeFormatBatch({ operations });
  const results = Array.isArray(revertResult.results)
    ? revertResult.results
    : [];

  if (results.length === 0 && !revertResult.success) {
    throw new Error(
      `Failed to revert write_format_batch: ${revertResult.message}`,
    );
  }

  results.forEach((result, idx) => {
    const entryKey = revertEntries[idx]?.[0];
    if (result?.success) {
      logger.info("write_format_batch revert completed", {
        entryKey,
        success: result.success,
        message: result.message,
      });
      return;
    }

    if (isNotFoundError(result?.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${result.message}`,
      );
      return;
    }

    throw new Error(
      `Failed to revert write_format_batch: ${result?.message}`,
    );
  });
}

async function revertDragFormula(change: ToolChangeInterface) {
  logger.info("Reverting drag_formula");

  const inputDataRevert = change.inputDataRevert;

  if (
    !inputDataRevert || !inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH]
  ) {
    throw new Error("Missing required data for drag_formula revert");
  }

  const writeValuesParams =
    inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH];
  await executeWriteValuesBatch(
    [writeValuesParams],
    "Failed to revert drag_formula",
  );
}

async function revertAddSheets(change: ToolChangeInterface) {
  logger.info("Reverting add_sheets");

  const inputDataRevert = change.inputDataRevert;

  if (!inputDataRevert || !inputDataRevert[RevertOperationKeys.REMOVE_SHEETS]) {
    throw new Error("Missing required data for add_sheets revert");
  }

  const removeSheetsParams = inputDataRevert[RevertOperationKeys.REMOVE_SHEETS];
  const revertResult = await removeSheets(removeSheetsParams);

  if (!revertResult.success) {
    if (isNotFoundError(revertResult.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${revertResult.message}`,
      );
      return;
    }
    throw new Error(`Failed to revert add_sheets: ${revertResult.message}`);
  }
}

async function revertAddRows(change: ToolChangeInterface) {
  logger.info("Reverting add_rows");

  const inputDataRevert = change.inputDataRevert;

  if (!inputDataRevert || !inputDataRevert[RevertOperationKeys.REMOVE_ROWS]) {
    throw new Error("Missing required data for add_rows revert");
  }

  const removeRowsParams = inputDataRevert[RevertOperationKeys.REMOVE_ROWS];
  const revertResult = await removeRows(removeRowsParams);

  if (!revertResult.success) {
    if (isNotFoundError(revertResult.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${revertResult.message}`,
      );
      return;
    }
    throw new Error(`Failed to revert add_rows: ${revertResult.message}`);
  }
}

async function revertAddColumns(change: ToolChangeInterface) {
  logger.info("Reverting add_columns");

  const inputDataRevert = change.inputDataRevert;

  if (
    !inputDataRevert || !inputDataRevert[RevertOperationKeys.REMOVE_COLUMNS]
  ) {
    throw new Error("Missing required data for add_columns revert");
  }

  const removeColumnsParams =
    inputDataRevert[RevertOperationKeys.REMOVE_COLUMNS];

  logger.info(`Reverting add_columns: ${removeColumnsParams}`);

  const revertResult = await removeColumns(removeColumnsParams);

  if (!revertResult.success) {
    if (isNotFoundError(revertResult.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${revertResult.message}`,
      );
      return;
    }
    throw new Error(`Failed to revert add_columns: ${revertResult.message}`);
  }
}

async function revertRemoveRows(change: ToolChangeInterface) {
  logger.info("Reverting remove_rows");

  const inputDataRevert = change.inputDataRevert;
  if (
    !inputDataRevert || !inputDataRevert[RevertOperationKeys.ADD_ROWS] ||
    !inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH]
  ) {
    throw new Error("Missing required data for remove_rows revert");
  }

  const addRowsParams = inputDataRevert[RevertOperationKeys.ADD_ROWS];
  const revertResultAddRows = await addRows(addRowsParams);
  if (!revertResultAddRows.success) {
    if (isNotFoundError(revertResultAddRows.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${revertResultAddRows.message}`,
      );
      return;
    }
    throw new Error(
      `Failed to revert remove_rows: ${revertResultAddRows.message}`,
    );
  }

  const writeValuesParams =
    inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH];
  await executeWriteValuesBatch(
    [writeValuesParams],
    "Failed to revert remove_rows",
  );
}

async function revertRemoveColumns(change: ToolChangeInterface) {
  logger.info("Reverting remove_columns");

  const inputDataRevert = change.inputDataRevert;
  if (
    !inputDataRevert ||
    !inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH] ||
    !inputDataRevert[RevertOperationKeys.ADD_COLUMNS]
  ) {
    throw new Error("Missing required data for remove_columns revert");
  }

  const addColumnsParams = inputDataRevert[RevertOperationKeys.ADD_COLUMNS];
  const revertResultAddColumns = await addColumns(addColumnsParams);
  if (!revertResultAddColumns.success) {
    if (isNotFoundError(revertResultAddColumns.message)) {
      logger.warn(
        `Worksheet not found during revert - skipping: ${revertResultAddColumns.message}`,
      );
      return;
    }
    throw new Error(
      `Failed to revert remove_columns: ${revertResultAddColumns.message}`,
    );
  }

  const writeValuesParams =
    inputDataRevert[RevertOperationKeys.WRITE_VALUES_BATCH];
  await executeWriteValuesBatch(
    [writeValuesParams],
    "Failed to revert remove_columns",
  );
}

async function revertCopyPaste(change: ToolChangeInterface): Promise<void> {
  logger.info("Reverting copy_paste", {
    hasRevertData: !!change.inputDataRevert,
    revertKeys: change.inputDataRevert
      ? Object.keys(change.inputDataRevert)
      : [],
  });

  const inputDataRevert = change.inputDataRevert;
  if (!inputDataRevert) {
    throw new Error("Missing required data for copy_paste revert");
  }

  if (inputDataRevert["delete-cells"]) {
    const deleteParams = inputDataRevert["delete-cells"];
    const deleteResult = await deleteCells(deleteParams);

    if (!deleteResult.success) {
      if (isNotFoundError(deleteResult.message)) {
        logger.warn(
          `Worksheet not found during revert - skipping: ${deleteResult.message}`,
        );
        return;
      }
      throw new Error(
        `Failed to revert copy_paste (delete cells): ${deleteResult.message}`,
      );
    }
    return;
  }

  const writeValuesKey = RevertOperationKeys.WRITE_VALUES_BATCH;
  const writeValuesParams = inputDataRevert[writeValuesKey] ||
    inputDataRevert["write-values"];
  if (writeValuesParams) {
    let valueError: Error | null = null;
    try {
      await executeWriteValuesBatch(
        [writeValuesParams],
        "Failed to revert copy_paste (write values)",
      );
    } catch (error) {
      valueError = error as Error;
    }

    try {
      await revertCopyPasteFormats(inputDataRevert);
    } catch (formatError) {
      logger.error("Failed to revert copy_paste formats", formatError);
      if (!valueError) {
        throw formatError;
      }
    }

    if (valueError) {
      throw valueError;
    }
    return;
  }

  throw new Error("Missing required data for copy_paste revert");
}

async function revertCopyPasteFormats(
  inputDataRevert: Record<string, any>,
): Promise<void> {
  const formatEntries = Object.entries(inputDataRevert).filter(([key]) =>
    key.startsWith("write-format-")
  );
  if (formatEntries.length === 0) {
    return;
  }

  logger.info("Reverting copy_paste formats", {
    entryCount: formatEntries.length,
    entryKeys: formatEntries.map(([key]) => key),
  });

  const operations = formatEntries.map(([, inputParams]) => ({
    workbookName: inputParams?.workbookName,
    worksheet: inputParams?.worksheet,
    range: inputParams?.range,
    format: inputParams?.format,
  }));

  const revertResult = await writeFormatBatch({ operations });
  const results = Array.isArray(revertResult.results)
    ? revertResult.results
    : [];

  if (results.length === 0 && !revertResult.success) {
    throw new Error(
      `Failed to revert copy_paste (write format): ${revertResult.message}`,
    );
  }

  results.forEach((result, idx) => {
    const entryKey = formatEntries[idx]?.[0];
    if (result?.success) {
      logger.info("copy_paste format revert completed", {
        entryKey,
        success: result.success,
        message: result.message,
      });
      return;
    }

    if (isNotFoundError(result?.message)) {
      logger.warn(
        `Worksheet not found during format revert - skipping: ${result.message}`,
      );
      return;
    }

    throw new Error(
      `Failed to revert copy_paste (write format): ${result?.message}`,
    );
  });
}
