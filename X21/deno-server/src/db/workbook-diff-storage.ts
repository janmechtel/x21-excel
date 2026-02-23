import { getDatabasePath, getDb } from "./sqlite.ts";
import { createLogger } from "../utils/logger.ts";
import type { DiffResult } from "../services/workbook-diff.ts";
import { formatDiffResults } from "../services/workbook-diff.ts";

const logger = createLogger("WorkbookDiffStorage");

export interface WorkbookSummary {
  id: string;
  workbookName: string;
  summaryText: string;
  createdAt: number;
  diffId: string | null; // optional link to diff
  comparisonType: "self" | "external";
  comparisonFilePath: string | null;
  comparisonFileModifiedAt: number | null;
}

export interface WorkbookDiffRecord {
  id: string;
  workbookName: string;
  unifiedDiff: string;
  createdAt: number;
  comparisonType: "self" | "external";
}

export interface WorkbookSnapshotRecord {
  workbookName: string;
  snapshotJson: string;
  updatedAt: number;
}

/**
 * Save diffs for a workbook change event
 * Combines all diffs into a single unified_diff per workbook
 * Returns the diff id for linking to summary
 */
export function saveWorkbookDiffs(
  workbookName: string,
  diffs: DiffResult[],
  comparisonType: "self" | "external" = "self",
): string {
  const db = getDb();
  const timestamp = Date.now();

  try {
    // Combine all diffs into a single unified diff
    const unifiedDiff = formatDiffResults(diffs);
    const id = `${workbookName}-${timestamp}-${Math.random()}`;

    db.query(
      `INSERT INTO workbook_diffs
       (id, workbook_name, unified_diff, created_at, comparison_type)
       VALUES (?, ?, ?, ?, ?)`,
      [id, workbookName, unifiedDiff, timestamp, comparisonType],
    );

    logger.info(
      `Saved unified diff for workbook: ${workbookName} with id: ${id}, comparison_type: ${comparisonType}`,
    );
    return id;
  } catch (error) {
    logger.error("Error saving workbook diffs", error);
    throw error;
  }
}

/**
 * Save a summary for a workbook change event
 * Links to the diff via diff_id (if provided)
 */
export function saveWorkbookSummary(
  workbookName: string,
  summaryText: string,
  diffId: string | null,
  comparisonType: "self" | "external" = "self",
  comparisonFilePath: string | null = null,
  comparisonFileModifiedAt: number | null = null,
): { id: string; timestamp: number } {
  const db = getDb();
  const timestamp = Date.now();
  const id = `${workbookName}-${timestamp}-${Math.random()}`;

  try {
    db.query(
      `INSERT INTO workbook_summaries
       (id, workbook_name, summary_text, created_at, diff_id, comparison_type, comparison_file_path, comparison_file_modified_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        id,
        workbookName,
        summaryText,
        timestamp,
        diffId,
        comparisonType,
        comparisonFilePath,
        comparisonFileModifiedAt,
      ],
    );

    logger.info(`Saved summary for workbook: ${workbookName}`);
    return { id, timestamp };
  } catch (error) {
    logger.error("Error saving workbook summary", error);
    throw error;
  }
}

/**
 * Get all summaries for a specific workbook, ordered by most recent first
 */
export function getWorkbookSummaries(
  workbookName: string,
  limit?: number,
): WorkbookSummary[] {
  const db = getDb();

  try {
    const query = limit
      ? `SELECT id, workbook_name, summary_text, created_at, diff_id, comparison_type, comparison_file_path, comparison_file_modified_at
         FROM workbook_summaries
         WHERE workbook_name = ?
         ORDER BY created_at DESC
         LIMIT ?`
      : `SELECT id, workbook_name, summary_text, created_at, diff_id, comparison_type, comparison_file_path, comparison_file_modified_at
         FROM workbook_summaries
         WHERE workbook_name = ?
         ORDER BY created_at DESC`;

    const params = limit ? [workbookName, limit] : [workbookName];
    const rows = db.query<
      [
        string,
        string,
        string,
        number,
        string | null,
        string,
        string | null,
        number | null,
      ]
    >(query, params);

    return rows.map((row) => ({
      id: row[0],
      workbookName: row[1],
      summaryText: row[2],
      createdAt: row[3],
      diffId: row[4],
      comparisonType: row[5] as "self" | "external",
      comparisonFilePath: row[6],
      comparisonFileModifiedAt: row[7],
    }));
  } catch (error) {
    logger.error("Error getting workbook summaries", error);
    throw error;
  }
}

/**
 * Get summaries for a workbook within a date range
 */
export function getWorkbookSummariesByDateRange(
  workbookName: string,
  startTimestamp: number | null,
  endTimestamp: number | null,
  limit?: number,
): WorkbookSummary[] {
  const db = getDb();

  try {
    let query =
      `SELECT id, workbook_name, summary_text, created_at, diff_id, comparison_type, comparison_file_path, comparison_file_modified_at
                 FROM workbook_summaries
                 WHERE workbook_name = ?`;
    const params: any[] = [workbookName];

    if (startTimestamp !== null) {
      query += ` AND created_at >= ?`;
      params.push(startTimestamp);
    }

    if (endTimestamp !== null) {
      query += ` AND created_at <= ?`;
      params.push(endTimestamp);
    }

    query += ` ORDER BY created_at DESC`;

    if (limit) {
      query += ` LIMIT ?`;
      params.push(limit);
    }

    const rows = db.query<
      [
        string,
        string,
        string,
        number,
        string | null,
        string,
        string | null,
        number | null,
      ]
    >(query, params);

    return rows.map((row) => ({
      id: row[0],
      workbookName: row[1],
      summaryText: row[2],
      createdAt: row[3],
      diffId: row[4],
      comparisonType: row[5] as "self" | "external",
      comparisonFilePath: row[6],
      comparisonFileModifiedAt: row[7],
    }));
  } catch (error) {
    logger.error("Error getting workbook summaries by date range", error);
    throw error;
  }
}

/**
 * Get summaries with their diffs by date range
 */
export function getSummariesWithDiffsByDateRange(
  workbookName: string,
  startTimestamp: number | null,
  endTimestamp: number | null,
  limit?: number,
): Array<{ summary: WorkbookSummary; unifiedDiff: string | null }> {
  const summaries = getWorkbookSummariesByDateRange(
    workbookName,
    startTimestamp,
    endTimestamp,
    limit,
  );

  return summaries.map((summary) => {
    const unifiedDiff = summary.diffId ? getDiffById(summary.diffId) : null;
    return { summary, unifiedDiff };
  });
}

/**
 * Get a diff by its ID
 */
function getDiffById(diffId: string): string | null {
  const db = getDb();

  try {
    const rows = db.query<[string]>(
      `SELECT unified_diff
       FROM workbook_diffs
       WHERE id = ?
       LIMIT 1`,
      [diffId],
    );

    return rows.length > 0 ? rows[0][0] : null;
  } catch (error) {
    logger.error("Error getting diff by ID", error);
    return null;
  }
}

/**
 * Get the most recent summary for a workbook
 */
export function getLatestWorkbookSummary(
  workbookName: string,
): WorkbookSummary | null {
  const summaries = getWorkbookSummaries(workbookName, 1);
  return summaries.length > 0 ? summaries[0] : null;
}

/**
 * Get the most recent unified diff for a workbook
 */
export function getLatestWorkbookDiff(
  workbookName: string,
): string | null {
  const db = getDb();

  try {
    const rows = db.query<[string]>(
      `SELECT unified_diff
       FROM workbook_diffs
       WHERE workbook_name = ?
       ORDER BY created_at DESC
       LIMIT 1`,
      [workbookName],
    );

    return rows.length > 0 ? rows[0][0] : null;
  } catch (error) {
    logger.error("Error getting latest workbook diff", error);
    throw error;
  }
}

/**
 * Get summary and its linked diff by diff_id
 */
export function getSummaryWithDiffByDiffId(
  diffId: string,
): { summary: WorkbookSummary | null; unifiedDiff: string | null } {
  const db = getDb();

  try {
    // Get summary
    const summaryRows = db.query<
      [
        string,
        string,
        string,
        number,
        string | null,
        string,
        string | null,
        number | null,
      ]
    >(
      `SELECT id, workbook_name, summary_text, created_at, diff_id, comparison_type, comparison_file_path, comparison_file_modified_at
       FROM workbook_summaries
       WHERE diff_id = ?
       LIMIT 1`,
      [diffId],
    );

    const summary = summaryRows.length > 0
      ? {
        id: summaryRows[0][0],
        workbookName: summaryRows[0][1],
        summaryText: summaryRows[0][2],
        createdAt: summaryRows[0][3],
        diffId: summaryRows[0][4],
        comparisonType: summaryRows[0][5] as "self" | "external",
        comparisonFilePath: summaryRows[0][6],
        comparisonFileModifiedAt: summaryRows[0][7],
      }
      : null;

    // Get diff
    const diffRows = db.query<[string]>(
      `SELECT unified_diff
       FROM workbook_diffs
       WHERE id = ?
       LIMIT 1`,
      [diffId],
    );

    const unifiedDiff = diffRows.length > 0 ? diffRows[0][0] : null;

    return { summary, unifiedDiff };
  } catch (error) {
    logger.error("Error getting summary with diff by diff_id", error);
    throw error;
  }
}

/**
 * Get the most recent summary and diff for a workbook
 */
export function getLatestSummaryWithDiff(
  workbookName: string,
): { summary: WorkbookSummary | null; unifiedDiff: string | null } {
  const summary = getLatestWorkbookSummary(workbookName);
  if (!summary || !summary.diffId) {
    return { summary, unifiedDiff: null };
  }
  return getSummaryWithDiffByDiffId(summary.diffId);
}

/**
 * Get all unique workbook keys that have summaries
 */
export function getWorkbooksWithSummaries(): string[] {
  const db = getDb();

  try {
    const rows = db.query<[string]>(
      `SELECT DISTINCT workbook_name
       FROM workbook_summaries
       ORDER BY workbook_name`,
    );

    return rows.map((row) => row[0]);
  } catch (error) {
    const dbPath = getDatabasePath();
    logger.error("Error getting workbooks with summaries", {
      error,
      databasePath: dbPath,
      message: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }
}

/**
 * Update the summary text for a specific workbook summary entry
 */
export function updateWorkbookSummaryText(
  id: string,
  summaryText: string,
): void {
  const db = getDb();

  try {
    db.query(
      `UPDATE workbook_summaries
       SET summary_text = ?
       WHERE id = ?`,
      [summaryText, id],
    );

    logger.info(`Updated workbook summary with id: ${id}`);
  } catch (error) {
    logger.error("Error updating workbook summary", { id, error });
    throw error;
  }
}

/**
 * Delete a specific workbook summary entry
 */
export function deleteWorkbookSummary(id: string): void {
  const db = getDb();

  try {
    db.query(
      `DELETE FROM workbook_summaries
       WHERE id = ?`,
      [id],
    );

    logger.info(`Deleted workbook summary with id: ${id}`);
  } catch (error) {
    logger.error("Error deleting workbook summary", { id, error });
    throw error;
  }
}

/**
 * Load all file-specific summaries at startup
 * Returns a map of workbook key to latest summary
 */
export function loadAllFileSummaries(): Map<string, WorkbookSummary> {
  const workbooks = getWorkbooksWithSummaries();
  const summaries = new Map<string, WorkbookSummary>();

  for (const workbookName of workbooks) {
    const summary = getLatestWorkbookSummary(workbookName);
    if (summary) {
      summaries.set(workbookName, summary);
    }
  }

  logger.info(`Loaded ${summaries.size} file-specific summaries at startup`);
  return summaries;
}

/**
 * Save or update the latest snapshot for a workbook
 * Uses INSERT OR REPLACE to ensure only one snapshot per workbook
 */
export function saveWorkbookSnapshot(
  workbookName: string,
  snapshot: any,
): void {
  const db = getDb();
  const timestamp = Date.now();

  try {
    const snapshotJson = JSON.stringify(snapshot);

    db.query(
      `INSERT INTO workbook_snapshots (workbook_name, snapshot_json, updated_at)
       VALUES (?, ?, ?)
       ON CONFLICT(workbook_name) DO UPDATE SET
         snapshot_json = excluded.snapshot_json,
         updated_at = excluded.updated_at`,
      [workbookName, snapshotJson, timestamp],
    );

    logger.info(`Saved snapshot for workbook: ${workbookName}`);
  } catch (error) {
    logger.error("Error saving workbook snapshot", error);
    throw error;
  }
}

/**
 * Get the latest snapshot for a workbook
 * Returns null if no snapshot exists or if parsing fails
 */
export function getWorkbookSnapshot(workbookName: string): any | null {
  const db = getDb();

  try {
    const rows = db.query<[string]>(
      `SELECT snapshot_json
       FROM workbook_snapshots
       WHERE workbook_name = ?
       LIMIT 1`,
      [workbookName],
    );

    if (rows.length === 0) {
      return null;
    }

    const snapshotJson = rows[0][0];
    return JSON.parse(snapshotJson);
  } catch (error) {
    logger.error("Error getting workbook snapshot", {
      workbookName,
      error,
    });
    return null;
  }
}

/**
 * Copy snapshot from source workbook to target workbook (e.g., on Save As)
 * Only copies if source snapshot exists and target snapshot doesn't exist
 * Returns true if copy was successful, false otherwise
 */
export function copyWorkbookSnapshot(
  sourceWorkbookName: string,
  targetWorkbookName: string,
): boolean {
  const db = getDb();
  const timestamp = Date.now();

  try {
    // Get source snapshot
    const sourceSnapshot = getWorkbookSnapshot(sourceWorkbookName);
    if (!sourceSnapshot) {
      logger.warn(
        `Source snapshot not found for copy: ${sourceWorkbookName}`,
      );
      return false;
    }

    // Check if target already has a snapshot
    const targetSnapshot = getWorkbookSnapshot(targetWorkbookName);
    if (targetSnapshot) {
      logger.info(
        `Target snapshot already exists, skipping copy: ${targetWorkbookName}`,
      );
      return false;
    }

    // Copy the snapshot
    const snapshotJson = JSON.stringify(sourceSnapshot);
    db.query(
      `INSERT INTO workbook_snapshots (workbook_name, snapshot_json, updated_at)
       VALUES (?, ?, ?)`,
      [targetWorkbookName, snapshotJson, timestamp],
    );

    logger.info(
      `Copied snapshot from ${sourceWorkbookName} to ${targetWorkbookName}`,
    );
    return true;
  } catch (error) {
    logger.error("Error copying workbook snapshot", {
      sourceWorkbookName,
      targetWorkbookName,
      error,
    });
    return false;
  }
}
