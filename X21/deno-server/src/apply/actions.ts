import { stateManager, ToolChangeInterface } from "../state/state-manager.ts";
import { createLogger } from "../utils/logger.ts";
import { applySingleToolChange } from "./tool.ts";

const logger = createLogger("ApplyActions");

export async function applyFromToolIdOnwards(
  workbookName: string,
  toolId: string,
): Promise<ToolChangeInterface[]> {
  // Get all tool changes to find the target
  const allToolChanges = stateManager.getToolChanges(workbookName);
  console.log(`All tool changes: ${JSON.stringify(allToolChanges)}`);

  const targetToolChange = allToolChanges.find((change) =>
    change.toolId === toolId
  );
  if (!targetToolChange) {
    throw new Error(
      `Target tool change not found for workbook ${workbookName} and tool ID ${toolId}`,
    );
  }

  // Only apply changes that are approved
  const approvedToolChanges = allToolChanges.filter((change) =>
    change.approved
  );
  const changesToApply = approvedToolChanges.filter((change) =>
    !change.applied && change.timestamp >= targetToolChange.timestamp
  );

  if (changesToApply.length === 0) {
    logger.info(
      `No approved tool changes to apply from tool ${toolId} onwards`,
    );
    return [];
  }

  await applyToolChanges(changesToApply);

  for (const change of changesToApply) {
    stateManager.updateToolChange(workbookName, change.toolId, {
      applied: true,
    });
  }

  return changesToApply;
}

async function applyToolChanges(
  toolChanges: ToolChangeInterface[],
): Promise<void> {
  const workbookName = toolChanges[0].workbookName ?? "unknown";
  // Sort by timestamp in ascending order (oldest first) to apply in chronological order
  const sortedChanges = toolChanges.sort((a, b) =>
    a.timestamp.getTime() - b.timestamp.getTime()
  );

  logger.info(
    `Starting apply of ${sortedChanges.length} tool changes for workbook ${workbookName}`,
  );

  for (const change of sortedChanges) {
    try {
      await applySingleToolChange(change);
      logger.info(
        `Successfully applied tool ${change.toolName} with timestamp ${change.timestamp}`,
      );
    } catch (error) {
      logger.error(`Failed to apply tool ${change.toolName}: ${error}`);
      throw error; // Stop applying if one fails
    }
  }

  logger.info(
    `Completed apply of all tool changes for workbook ${workbookName}`,
  );
}
