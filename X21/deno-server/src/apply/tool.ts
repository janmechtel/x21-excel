import { addColumns } from "../excel-actions/add-columns.ts";
import { addRows } from "../excel-actions/add-rows.ts";
import { addSheets } from "../excel-actions/add-sheets.ts";
import { dragFormula } from "../excel-actions/drag-formula.ts";
import { copyPaste } from "../excel-actions/copy-paste.ts";
import { removeColumns } from "../excel-actions/remove-columns.ts";
import { removeRows } from "../excel-actions/remove-rows.ts";
import { writeFormatBatch } from "../excel-actions/write-format-batch.ts";
import { writeValuesBatch } from "../excel-actions/write-values-batch.ts";
import { ToolChangeInterface } from "../state/state-manager.ts";
import { createLogger } from "../utils/logger.ts";
import { ToolNames } from "../types/index.ts";

const logger = createLogger("ApplyTool");

export async function applySingleToolChange(
  change: ToolChangeInterface,
): Promise<void> {
  const toolName = change.toolName;

  // Skip read operations as they don't make changes
  if (
    toolName === ToolNames.READ_VALUES_BATCH ||
    toolName === ToolNames.READ_FORMAT_BATCH
  ) {
    logger.debug(`Skipping apply for ${toolName} - no changes to apply`);
    return;
  }

  switch (toolName) {
    case ToolNames.ADD_SHEETS:
      await addSheets(change.inputData);
      break;
    case ToolNames.WRITE_FORMAT_BATCH:
      await writeFormatBatch(change.inputData);
      break;
    case ToolNames.WRITE_VALUES_BATCH: {
      const operations = Array.isArray(change.inputData?.operations)
        ? change.inputData.operations
        : [];
      const fallbackWorkbook = change.inputData?.workbookName;
      await writeValuesBatch({
        operations: operations.map((op: any) => ({
          ...op,
          workbookName: op.workbookName ?? fallbackWorkbook,
        })),
      });
      break;
    }
    case ToolNames.DRAG_FORMULA:
      await dragFormula(change.inputData);
      break;
    case ToolNames.COPY_PASTE:
      await copyPaste(change.inputData);
      break;
    case ToolNames.ADD_ROWS:
      await addRows(change.inputData);
      break;
    case ToolNames.REMOVE_ROWS:
      await removeRows(change.inputData);
      break;
    case ToolNames.ADD_COLUMNS:
      await addColumns(change.inputData);
      break;
    case ToolNames.REMOVE_COLUMNS:
      await removeColumns(change.inputData);
      break;
    default:
      logger.warn(`Tool ${toolName} not found`);
      break;
  }
}
