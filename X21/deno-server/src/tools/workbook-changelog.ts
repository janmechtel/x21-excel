import {
  type Tool,
  ToolNames,
  type WorkbookChangelogRequest,
  type WorkbookChangelogResponse,
} from "../types/index.ts";
import {
  getLatestSummaryWithDiff,
  getSummariesWithDiffsByDateRange,
} from "../db/workbook-diff-storage.ts";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("WorkbookChangelogTool");

export class WorkbookChangelogTool implements Tool<WorkbookChangelogRequest> {
  name = ToolNames.WORKBOOK_CHANGELOG;
  description =
    "Retrieve changelog summaries for the current Excel workbook. " +
    "Use this whenever the user asks for a changelog, summary of recent Excel changes, what changed recently, " +
    "or asks about changes since a specific date (e.g., 'what did I change since last month/week/yesterday'). " +
    "Date filtering expects timestamps (ms since epoch). The LLM should convert natural language dates into " +
    "startTimestampMs/endTimestampMs using nowMs as the reference for relative phrases. " +
    "This tool does NOT trigger new analysis – it surfaces stored changelog summaries from the database.";

  input_schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      workbookName: {
        type: "string",
        description:
          "Optional workbook name. If omitted, the active workbook is used.",
      },
      includeDiff: {
        type: "boolean",
        description:
          "Whether to include the underlying unified diff text along with the summary. Defaults to false.",
      },
      startTimestampMs: {
        type: "number",
        description:
          "Optional start timestamp in milliseconds since epoch. Provide this to filter changelogs from this time onward.",
      },
      endTimestampMs: {
        type: "number",
        description:
          "Optional end timestamp in milliseconds since epoch. Defaults to current time when omitted.",
      },
    },
  };

  execute(
    params: WorkbookChangelogRequest,
  ): Promise<WorkbookChangelogResponse> {
    const workbookName = params.workbookName;

    if (!workbookName || workbookName.trim().length === 0) {
      // This should normally be filled in by applyWorkbookGuards,
      // but we guard here as well to give a clear error to the model.
      logger.warn(
        "WorkbookChangelogTool called without workbookName after guards",
      );
      return Promise.resolve({
        success: false,
        workbookName: "",
        message:
          "No workbook name was provided and no active workbook was detected. Please retry with a workbookName.",
      });
    }

    try {
      if (typeof params.startTimestampMs === "number") {
        const startTimestamp = params.startTimestampMs;
        const nowMs = Date.now();
        const endTimestamp = typeof params.endTimestampMs === "number"
          ? params.endTimestampMs
          : nowMs;

        if (Number.isNaN(startTimestamp)) {
          return Promise.resolve({
            success: false,
            workbookName,
            message:
              "Invalid startTimestampMs: must be a number (ms since epoch).",
          });
        }

        if (Number.isNaN(endTimestamp)) {
          return Promise.resolve({
            success: false,
            workbookName,
            message:
              "Invalid endTimestampMs: must be a number (ms since epoch).",
          });
        }

        logger.info("Fetching changelogs by date range", {
          workbookName,
          startTimestamp,
          endTimestamp,
        });

        const results = getSummariesWithDiffsByDateRange(
          workbookName,
          startTimestamp,
          endTimestamp,
        );

        if (results.length === 0) {
          return Promise.resolve({
            success: false,
            workbookName,
            message:
              `No activity information is available for this workbook between ${
                new Date(startTimestamp).toISOString()
              } and ${new Date(endTimestamp).toISOString()}. ` +
              `Activity is recorded only when you manually click the "Generate summary" button in the Activity panel (📄 symbol). ` +
              `Please generate a summary first, then try again.`,
            dateRange: {
              start: startTimestamp,
              end: endTimestamp,
            },
          });
        }

        const summaries = results.map(({ summary, unifiedDiff }) => ({
          id: summary.id,
          createdAt: summary.createdAt,
          text: summary.summaryText,
          diffId: summary.diffId,
          ...(params.includeDiff && unifiedDiff ? { unifiedDiff } : {}),
        }));

        return Promise.resolve({
          success: true,
          workbookName,
          message:
            `Found ${summaries.length} activity information(s) in the specified date range.`,
          summaries,
          dateRange: {
            start: startTimestamp,
            end: endTimestamp,
          },
        });
      }

      logger.info("Fetching latest changelog for workbook", { workbookName });

      const { summary, unifiedDiff } = getLatestSummaryWithDiff(workbookName);

      if (!summary) {
        logger.info("No changelog summaries found for workbook", {
          workbookName,
        });
        return Promise.resolve({
          success: false,
          workbookName,
          message: `No activity information is available for this workbook. ` +
            `Activity is recorded only when you manually click the "Generate summary" button in the Activity panel (📄 symbol). ` +
            `Please generate a summary first, then try again.`,
        });
      }

      const response: WorkbookChangelogResponse = {
        success: true,
        workbookName,
        message: "Latest workbook changelog retrieved successfully.",
        summary: {
          id: summary.id,
          createdAt: summary.createdAt,
          text: summary.summaryText,
          diffId: summary.diffId,
        },
      };

      if (params.includeDiff && unifiedDiff && response.summary) {
        response.summary.unifiedDiff = unifiedDiff;
      }

      return Promise.resolve(response);
    } catch (error) {
      logger.error("Failed to fetch workbook changelog", {
        workbookName,
        error,
      });
      return Promise.resolve({
        success: false,
        workbookName,
        message:
          "Failed to retrieve changelog for this workbook. The database may be unavailable or corrupted.",
      });
    }
  }
}
