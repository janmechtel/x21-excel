import { createLogger } from "../utils/logger.ts";
import { stateManager } from "../state/state-manager.ts";
import { WebSocketManager } from "./websocket-manager.ts";
import { generateWorkbookDiff } from "./workbook-diff.ts";
import { runAutonomousAnalysis } from "./autonomous-diff-agent.ts";
import {
  getWorkbookSnapshot,
  saveWorkbookDiffs,
  saveWorkbookSnapshot,
  saveWorkbookSummary,
} from "../db/workbook-diff-storage.ts";
import { WebSocketMessageTypes } from "../types/index.ts";

const logger = createLogger("ChangelogService");

export interface ChangelogResult {
  success: boolean;
  message: string;
  isInitialSnapshot: boolean;
  diffResults?: any[];
  diffId?: string | null;
  comparisonMetadata?: {
    comparisonType: "self" | "external";
    comparisonFilePath: string | null;
    comparisonFileModifiedAt: number | null;
  };
}

/**
 * Service responsible for handling workbook changelog generation
 * including snapshot management, diff generation, and autonomous analysis.
 *
 * Snapshot storage:
 * - Snapshots are stored only in the database (no in-memory storage)
 * - Initial snapshots are saved when a workbook is first opened
 * - Consecutive snapshots are only saved to DB when there are actual changes
 *   (preserves baseline snapshot when file is reopened without changes)
 */
export class ChangelogService {
  private static instance: ChangelogService;

  static getInstance(): ChangelogService {
    if (!this.instance) {
      this.instance = new ChangelogService();
    }
    return this.instance;
  }

  /**
   * Process consecutive snapshot - generates diffs and saves them, but does NOT generate summary.
   * Requires a previous snapshot to exist in the database.
   *
   * Note: Initial snapshots should be handled via handleInitialSnapshot().
   * This method is used by generateChangelog().
   *
   * @returns Diff results and ID when there's a previous snapshot
   * @throws Error if no previous snapshot exists (should use handleInitialSnapshot instead)
   */
  processSnapshot(
    workbookName: string,
    currentSnapshot: any,
    comparisonSnapshot?: any,
  ): ChangelogResult {
    try {
      logger.info("Processing snapshot", {
        workbookName,
        hasComparisonSnapshot: !!comparisonSnapshot,
      });

      // Ensure workbook state exists
      stateManager.getOrCreateState(workbookName);

      // Determine comparison metadata
      let comparisonType: "self" | "external" = "self";
      let comparisonFilePath: string | null = null;
      let comparisonFileModifiedAt: number | null = null;

      // Get previous snapshot - use comparisonSnapshot if provided, otherwise from database
      let previousSnapshot: any | null = null;
      if (comparisonSnapshot) {
        previousSnapshot = comparisonSnapshot;
        comparisonType = "external";
        comparisonFilePath = comparisonSnapshot.filePath || null;
        comparisonFileModifiedAt = comparisonSnapshot.lastModified || null;
        const comparisonFileName = comparisonFilePath?.split(/[\\/]/).pop() ??
          null;
        logger.info("Using comparison snapshot from file", {
          workbookName,
          comparisonFileName,
          comparisonFileModifiedAt,
        });
      } else {
        try {
          previousSnapshot = getWorkbookSnapshot(workbookName);
          if (previousSnapshot) {
            logger.info("Loaded previous snapshot from database", {
              workbookName,
            });
          }
        } catch (error) {
          logger.error("Failed to load previous snapshot from database", {
            workbookName,
            error,
          });
        }
      }

      this.logSnapshotInfo(workbookName, previousSnapshot, currentSnapshot);

      // If no previous snapshot, automatically treat as initial snapshot
      // This handles the case where user disabled save copies, opened workbook, then enabled it
      if (!previousSnapshot) {
        logger.info(
          "No previous snapshot found - treating current as initial snapshot (user may have enabled save copies mid-session)",
          { workbookName },
        );

        // Automatically handle as initial snapshot to establish baseline
        const _result = this.handleInitialSnapshot(
          workbookName,
          currentSnapshot,
        );

        return {
          success: true,
          message:
            "Recording started. This snapshot will serve as the baseline for future changes.",
          isInitialSnapshot: true,
        };
      }

      // Generate diff and save
      const { diffResults, diffId, comparisonMetadata } = this
        .handleSnapshotDiff(
          workbookName,
          previousSnapshot,
          currentSnapshot,
          comparisonType,
          comparisonFilePath,
          comparisonFileModifiedAt,
        );

      return {
        success: true,
        message: "Snapshot processed, diffs saved",
        isInitialSnapshot: false,
        diffResults,
        diffId,
        comparisonMetadata,
      };
    } catch (error) {
      logger.error("Error processing snapshot", error);
      throw error;
    }
  }

  /**
   * Generate changelog with summary - processes consecutive snapshot, saves diffs, AND generates summary.
   *
   * Flow:
   * 1. Processes snapshot and generates diffs (via processSnapshot)
   * 2. Saves diffs to database (only if there are actual changes)
   * 3. Updates DB snapshot (only if there are actual changes, preserves baseline)
   * 4. Runs autonomous analysis in background to generate summary
   */
  generateChangelog(
    workbookName: string,
    currentSnapshot: any,
    comparisonSnapshot?: any,
  ): ChangelogResult {
    try {
      const comparisonFileName = comparisonSnapshot?.filePath?.split(/[\\/]/)
        .pop() ?? comparisonSnapshot?.fileName ?? null;
      logger.info("Generating changelog with summary", {
        workbookName,
        hasComparisonSnapshot: !!comparisonSnapshot,
        comparisonFileName,
      });

      // Use processSnapshot to handle common snapshot processing logic
      const result = this.processSnapshot(
        workbookName,
        currentSnapshot,
        comparisonSnapshot,
      );

      // If it was an initial snapshot, return early (no analysis needed)
      if (result.isInitialSnapshot) {
        return result;
      }

      // Extract diff results, diff ID, and comparison metadata for analysis
      const diffResults = result.diffResults || [];
      const diffId = result.diffId ?? null;
      const comparisonMetadata = result.comparisonMetadata || {
        comparisonType: "self" as const,
        comparisonFilePath: null,
        comparisonFileModifiedAt: null,
      };

      // If there are no actual changes, we should still notify the UI so it can
      // clear its "Generating..." state. (The UI clears generating state on any
      // `workbook:change_summary` message.)
      const hasActualChanges = diffResults.some((d) => d?.hasChanges === true);
      if (!hasActualChanges) {
        const wsManager = WebSocketManager.getInstance();
        wsManager.send(
          workbookName,
          WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY,
          {
            summary: "No changes detected.",
            sheetsAffected: 0,
            timestamp: Date.now(),
            comparisonType: comparisonMetadata.comparisonType,
            comparisonFilePath: comparisonMetadata.comparisonFilePath,
          },
        );

        return {
          success: true,
          message: "No changes detected",
          isInitialSnapshot: false,
        };
      }

      // Run autonomous analysis in background (only when there are actual changes)
      if (diffResults.length > 0 && diffId !== null) {
        logger.info("Starting autonomous analysis for changelog generation");
        this.runAnalysisInBackground(
          workbookName,
          diffResults,
          diffId,
          comparisonMetadata,
        );
      }

      return {
        success: true,
        message: "Analysis started",
        isInitialSnapshot: false,
      };
    } catch (error) {
      logger.error("Error generating changelog", error);
      throw error;
    }
  }

  /**
   * Handle initial snapshot storage - called when a workbook is first opened.
   * Only saves if no snapshot exists in DB to avoid overwriting existing snapshots.
   *
   * Called directly from router when C# sends isInitialSnapshot: true.
   * This establishes the baseline snapshot for future diff comparisons.
   *
   * @param workbookName - Name of the workbook
   * @param snapshot - The initial snapshot to store
   * @returns ChangelogResult indicating success/failure
   */
  handleInitialSnapshot(
    workbookName: string,
    snapshot: any,
  ): ChangelogResult {
    // Check if snapshot already exists in DB before saving
    // This prevents overwriting existing snapshots if getWorkbookSnapshot failed earlier
    try {
      const existing = getWorkbookSnapshot(workbookName);
      if (existing) {
        logger.info(
          "Snapshot already exists in DB, not overwriting.",
          { workbookName },
        );
        // Return success but don't save - preserve existing snapshot
        this.sendInitialSnapshotNotification(workbookName);
        return {
          success: true,
          message: "Snapshot already exists in database",
          isInitialSnapshot: true,
        };
      }
    } catch (error) {
      // If check fails, proceed with save (better to have a snapshot than not)
      logger.warn(
        "Failed to check for existing snapshot, proceeding with save",
        {
          workbookName,
          error,
        },
      );
    }

    logger.info("First snapshot for workbook - storing initial state", {
      workbookName,
    });

    // Persist the initial snapshot so we can resume from it in future sessions
    try {
      saveWorkbookSnapshot(workbookName, snapshot);
      logger.info("Stored initial snapshot", {
        workbookName,
        sheetCount: snapshot?.sheetXmls
          ? Object.keys(snapshot.sheetXmls).length
          : 0,
      });
    } catch (error) {
      logger.error("Failed to persist initial snapshot to database", {
        workbookName,
        error,
      });
      // Continue even if snapshot persistence fails
    }

    // Send WebSocket notification if connected
    this.sendInitialSnapshotNotification(workbookName);

    return {
      success: true,
      message: "Initial snapshot stored",
      isInitialSnapshot: true,
    };
  }

  /**
   * Handle diff generation and saving for consecutive snapshots.
   *
   * Behavior:
   * - Generates diffs between previous and current snapshot
   * - Saves diffs to database only if there are actual changes
   * - Updates DB snapshot only if there are actual changes (preserves baseline)
   * - Always reads previous snapshot from database (no in-memory storage)
   *
   * @returns Diff results and ID (null if no changes detected)
   */
  private handleSnapshotDiff(
    workbookName: string,
    previousSnapshot: any,
    currentSnapshot: any,
    comparisonType: "self" | "external" = "self",
    comparisonFilePath: string | null = null,
    comparisonFileModifiedAt: number | null = null,
  ): {
    diffResults: any[];
    diffId: string | null;
    comparisonMetadata: {
      comparisonType: "self" | "external";
      comparisonFilePath: string | null;
      comparisonFileModifiedAt: number | null;
    };
  } {
    logger.info("About to generate diff", {
      workbookName,
      hasPrevious: !!previousSnapshot,
      hasCurrent: !!currentSnapshot,
      comparisonType,
    });

    // Generate diffs synchronously
    const diffResults = generateWorkbookDiff(
      previousSnapshot,
      currentSnapshot,
    );

    logger.info(`Generated ${diffResults.length} diff results`, {
      workbookName,
      diffCount: diffResults.length,
      diffTypes: diffResults.map((d) => ({
        sheet: d.sheetName,
        hasChanges: d.hasChanges,
        isActualSheet: d.isActualSheet,
      })),
    });

    // Save diffs to database and get diff_id for linking
    // Only save if there are actual changes
    let diffId: string | null = null;
    const hasActualChanges = diffResults.some((d) => d.hasChanges);
    if (hasActualChanges) {
      try {
        diffId = (saveWorkbookDiffs as any)(
          workbookName,
          diffResults,
          comparisonType,
        );
      } catch (error) {
        logger.error("Failed to save diffs to database", error);
        // Continue even if save fails
      }
    } else {
      logger.info("No actual changes detected, skipping diff save", {
        workbookName,
      });
    }

    // Only update DB snapshot if there are actual changes
    // This preserves the baseline when file is reopened without changes
    if (hasActualChanges) {
      try {
        saveWorkbookSnapshot(workbookName, currentSnapshot);
        logger.info("Stored updated snapshot for next comparison", {
          workbookName,
        });
      } catch (error) {
        logger.error("Failed to persist updated snapshot to database", {
          workbookName,
          error,
        });
        // Continue even if snapshot persistence fails
      }
    } else {
      logger.info("No changes detected, preserving baseline snapshot in DB", {
        workbookName,
      });
    }

    return {
      diffResults,
      diffId,
      comparisonMetadata: {
        comparisonType,
        comparisonFilePath,
        comparisonFileModifiedAt,
      },
    };
  }

  /**
   * Log snapshot information for debugging
   */
  private logSnapshotInfo(
    workbookName: string,
    previousSnapshot: any,
    currentSnapshot: any,
  ): void {
    logger.info("Snapshot check", {
      workbookName,
      hasPreviousSnapshot: !!previousSnapshot,
      previousSheetCount: previousSnapshot?.sheetXmls
        ? Object.keys(previousSnapshot.sheetXmls).length
        : 0,
      currentSheetCount: currentSnapshot?.sheetXmls
        ? Object.keys(currentSnapshot.sheetXmls).length
        : 0,
      previousHasWorkbook: !!previousSnapshot?.workbookXml,
      currentHasWorkbook: !!currentSnapshot?.workbookXml,
    });
  }

  /**
   * Send WebSocket notification for initial snapshot
   */
  private sendInitialSnapshotNotification(workbookName: string): void {
    const wsManager = WebSocketManager.getInstance();
    if (wsManager.isConnected(workbookName)) {
      wsManager.send(
        workbookName,
        WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY,
        {
          summary:
            "Activity analyser initialized | Activity analysis is now active. Future changes will be summarized.",
          sheetsAffected: 0,
          timestamp: Date.now(),
        },
      );
    } else {
      logger.info("WebSocket not connected yet - skipping initial message", {
        workbookName,
      });
    }
  }

  /**
   * Run autonomous analysis in background without blocking
   */
  private runAnalysisInBackground(
    workbookName: string,
    diffResults: any[],
    diffId: string | null,
    comparisonMetadata: {
      comparisonType: "self" | "external";
      comparisonFilePath: string | null;
      comparisonFileModifiedAt: number | null;
    },
  ): void {
    const wsManager = WebSocketManager.getInstance();

    // Start autonomous analysis in background
    (async () => {
      try {
        logger.info("Starting autonomous diff analysis in background");
        const summary = await runAutonomousAnalysis(
          workbookName,
          diffResults,
          {
            diffId,
            comparisonType: comparisonMetadata.comparisonType,
            comparisonFilePath: comparisonMetadata.comparisonFilePath,
            comparisonFileModifiedAt:
              comparisonMetadata.comparisonFileModifiedAt,
          },
        );

        logger.info("Autonomous analysis complete", {
          summary,
          summaryLength: summary?.length || 0,
          isEmpty: !summary || !summary.trim(),
        });

        // Validate summary before processing
        if (!summary || !summary.trim()) {
          logger.warn(
            "Received empty summary from autonomous analysis - sending fallback notification",
          );
          wsManager.send(
            workbookName,
            WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY,
            {
              summary:
                "Changes detected in workbook, but no summary could be generated.",
              sheetsAffected: diffResults.filter((r) => r.isActualSheet).length,
              timestamp: Date.now(),
              comparisonType: comparisonMetadata.comparisonType,
              comparisonFilePath: comparisonMetadata.comparisonFilePath,
            },
          );
          return;
        }

        // Save summary to database (linked to diff via diff_id if available)
        // Use comparison metadata passed from the diff generation
        let summaryId: string | null = null;
        let summaryTimestamp: number = Date.now();
        try {
          const saved = (saveWorkbookSummary as any)(
            workbookName,
            summary,
            diffId,
            comparisonMetadata.comparisonType,
            comparisonMetadata.comparisonFilePath,
            comparisonMetadata.comparisonFileModifiedAt,
          );
          summaryId = saved.id;
          summaryTimestamp = saved.timestamp;
          if (!diffId) {
            logger.warn(
              "Saved summary without diff_id. Diffs may not have been saved.",
            );
          }
        } catch (error) {
          logger.error("Failed to save summary to database", error);
          // Continue even if save fails
        }

        // Broadcast result via WebSocket
        const payload = {
          // Include DB id and timestamp so UI can de-duplicate entries when also loading from API
          id: summaryId,
          summary,
          sheetsAffected: diffResults.filter((r) => r.isActualSheet).length,
          timestamp: summaryTimestamp,
          comparisonType: comparisonMetadata.comparisonType,
          comparisonFilePath: comparisonMetadata.comparisonFilePath,
        };

        wsManager.send(
          workbookName,
          WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY,
          payload,
        );
      } catch (error) {
        logger.error("Autonomous analysis failed", error);

        // Send error notification
        wsManager.send(
          workbookName,
          WebSocketMessageTypes.WORKBOOK_CHANGE_SUMMARY,
          {
            summary: `Analysis encountered an error | ${
              error instanceof Error ? error.message : String(error)
            }`,
            sheetsAffected: 0,
            timestamp: Date.now(),
          },
        );
      }
    })();
  }
}

export const changelogService = ChangelogService.getInstance();
