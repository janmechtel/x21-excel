import { stateManager, ToolChangeInterface } from "../state/state-manager.ts";
import { createLogger } from "../utils/logger.ts";
import { revertSingleToolChange } from "./tool.ts";

const logger = createLogger("RevertActions");

async function revertToolChanges(
  toolChanges: ToolChangeInterface[],
): Promise<void> {
  // Sort by timestamp in descending order (newest first) to revert in reverse order
  const sortedChanges = toolChanges.sort((a, b) =>
    b.timestamp.getTime() - a.timestamp.getTime()
  );

  const workbookName = toolChanges[0].workbookName ?? "unknown";

  logger.info(
    `Starting revert of ${sortedChanges.length} tool changes for workbook ${workbookName}`,
  );

  for (const change of sortedChanges) {
    try {
      await revertSingleToolChange(change, false);
      logger.info(
        `Successfully reverted tool ${change.toolName} with timestamp ${change.timestamp}`,
      );
    } catch (error) {
      logger.error(`Failed to revert tool ${change.toolName}: ${error}`);
      throw error; // Stop reverting if one fails
    }
  }

  logger.info(
    `Completed revert of all tool changes for workbook ${workbookName}`,
  );
}

export async function revertToolChangesForWorkbookFromToolIdOnwards(
  workbookName: string,
  toolId: string,
): Promise<ToolChangeInterface[]> {
  // Get all tool changes to find the target
  const allToolChanges = stateManager.getToolChanges(workbookName);
  logger.info("Loaded tool changes for revert", {
    count: allToolChanges.length,
    workbookName,
  });

  const targetToolChange = allToolChanges.find((change) =>
    change.toolId === toolId
  );
  if (!targetToolChange) {
    throw new Error(
      `Target tool change not found for workbook ${workbookName} and tool ID ${toolId}`,
    );
  }

  // Only revert changes that are applied
  const appliedToolChanges = allToolChanges.filter((change) => change.applied);
  const changesToRevert = appliedToolChanges.filter((change) =>
    change.timestamp >= targetToolChange.timestamp
  );

  if (changesToRevert.length === 0) {
    logger.info(
      `No applied tool changes to revert from tool ${toolId} onwards`,
    );
    return [];
  }

  await revertToolChanges(changesToRevert);

  for (const change of changesToRevert) {
    stateManager.updateToolChange(workbookName, change.toolId, {
      applied: false,
    });
  }

  return changesToRevert;
}
