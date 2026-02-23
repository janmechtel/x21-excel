import { ErrorHandler } from "../middleware/error-handler.ts";
import { createLogger } from "../utils/logger.ts";
import { getEnvironment, getLogsEnabled } from "../utils/environment.ts";
import { WebSocketManager } from "../services/websocket-manager.ts";
import { stateManager, ToolChangeInterface } from "../state/state-manager.ts";
import { tracing } from "../tracing/tracing.ts";
import { createPrompt } from "../stream/prompt-builder.ts";
import {
  buildAnthropicParamsForConversatoin as buildAnthropicParamsForConversation,
  buildAnthropicParamsForTokenCount,
  createAnthropicClient,
  tokenLimit,
  withAnthropicBetas,
} from "../llm-client/anthropic.ts";
import {
  getAnthropicConfig,
  getAzureOpenAIConfig,
  getLLMProvider,
  reloadLLMConfig,
} from "../llm-client/provider.ts";
import {
  getInputDataRevert,
  streamClaudeResponseAndHandleToolUsage,
  toolExecutionFlow,
} from "../stream/tool-logic.ts";
import { ToolExecutionError } from "../errors/tool-execution-error.ts";
import { Anthropic } from "@anthropic-ai/sdk";
import { revertToolChangesForWorkbookFromToolIdOnwards } from "../revert/actions.ts";
import { applyFromToolIdOnwards } from "../apply/actions.ts";
import { UserService } from "../services/user.ts";
import {
  listMessagesByConversation,
  listRecentChats,
  listRecentChatsAll,
  listRecentUserMessages,
  searchRecentChats,
  searchRecentChatsAll,
} from "../db/dal.ts";
import {
  getLlmKeysConfigByProvider,
  upsertLlmKeysConfigByProvider,
} from "../db/llm-keys-dal.ts";
import {
  compactChatHistory,
  isHistoryExceedingTokenLimit,
} from "../compacting/compacting.ts";
import { revertSingleToolChange } from "../revert/tool.ts";
import {
  extractClaudeErrorType,
  mapClaudeErrorToPayload,
} from "../stream/llm.ts";
import { UserCancellationError } from "../errors/user-cancellation-error.ts";
import { slashCommandService } from "../services/slash-commands.ts";
import { validateUiRequestResult } from "../utils/ui-request.ts";
import { listOpenWorkbooks } from "../excel-actions/list-open-workbooks.ts";
import { getMetadata } from "../excel-actions/get-metadata.ts";
import { changelogService } from "../services/changelog-service.ts";
import {
  copyWorkbookSnapshot,
  deleteWorkbookSummary,
  getWorkbookSnapshot,
  getWorkbookSummaries,
  updateWorkbookSummaryText,
} from "../db/workbook-diff-storage.ts";
import {
  OperationStatusValues,
  ToolNames,
  WebSocketMessageTypes,
} from "../types/index.ts";

const logger = createLogger("Router");
const MAX_LOG_MESSAGE_LENGTH = 2000;
const MAX_LOG_LEVEL_LENGTH = 16;

export class Router {
  constructor() {
  }

  async handleRequest(req: Request): Promise<Response> {
    const url = new URL(req.url);

    // WebSocket endpoint (also exposed via dedicated WS server)
    if (url.pathname === "/ws" && req.headers.get("upgrade") === "websocket") {
      return await this.handleWebSocketRequest(req);
    }

    if (url.pathname === "/api/messages") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const workbookKey = url.searchParams.get("workbookKey") || "";
        const conversationId = url.searchParams.get("conversationId") || "";
        if (!workbookKey || !conversationId) {
          return ErrorHandler.createResponse(
            new Error(
              "Missing 'workbookKey' or 'conversationId' query parameter",
            ),
            400,
          );
        }
        const items = listMessagesByConversation(workbookKey, conversationId);
        return new Response(JSON.stringify({ items }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to fetch messages", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (
      url.pathname === "/api/excel/open-workbooks" ||
      url.pathname === "/api/openWorkbooks"
    ) {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const workbooks = await listOpenWorkbooks();
        logger.info("[OpenWorkbooks] Response", {
          count: workbooks.length,
          workbooks,
        });
        return new Response(JSON.stringify({ workbooks }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to list open workbooks", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (
      url.pathname === "/api/excel/workbook-metadata" ||
      url.pathname === "/api/getMetadata"
    ) {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const workbookName = url.searchParams.get("workbookName") ||
          url.searchParams.get("workbook_name") || undefined;
        const metadata = await getMetadata({ workbookName });
        return new Response(JSON.stringify(metadata), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to fetch workbook metadata", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/progress") {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type, X-Workbook-Name",
          },
        });
      }

      if (req.method !== "POST") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const body = await req.json();
        const workbookName = req.headers.get("x-workbook-name") ||
          body?.workbookName || "";
        const status = body?.status || OperationStatusValues.PROCESSING;
        const message = body?.message || null;
        const progress = body?.progress;
        const metadata = body?.metadata;

        if (workbookName) {
          WebSocketManager.getInstance().send(workbookName, "status:update", {
            status,
            message,
            progress,
            metadata,
          });
        } else {
          logger.warn("Progress update missing workbook name header/body");
        }

        return new Response(JSON.stringify({ success: true }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to handle progress update", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/recent-user-messages") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const workbookKey = url.searchParams.get("workbookKey") || "";
        const limitParam = url.searchParams.get("limit") || "50";
        const limit = Math.max(1, Math.min(200, Number(limitParam) || 50));

        if (!workbookKey) {
          return ErrorHandler.createResponse(
            new Error("Missing 'workbookKey' query parameter"),
            400,
          );
        }

        const items = listRecentUserMessages(workbookKey, limit);
        return new Response(JSON.stringify({ items }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to fetch recent user messages", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/environment") {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      return new Response(
        JSON.stringify({
          environment: getEnvironment(),
          logsEnabled: getLogsEnabled(),
        }),
        {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        },
      );
    }

    if (url.pathname === "/api/logs") {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "POST") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const body = await req.json().catch(() => ({}));
        const levelRaw = typeof body?.level === "string" ? body.level : "";
        const level = levelRaw.toLowerCase();
        const message = typeof body?.message === "string" ? body.message : "";
        const context = body?.context ?? {};

        if (!message) {
          return ErrorHandler.createResponse(
            new Error("Missing 'message' in request body"),
            400,
          );
        }

        if (message.length > MAX_LOG_MESSAGE_LENGTH) {
          return ErrorHandler.createResponse(
            new Error("Message too long"),
            400,
          );
        }

        if (levelRaw.length > MAX_LOG_LEVEL_LENGTH) {
          return ErrorHandler.createResponse(
            new Error("Invalid 'level' (too long)"),
            400,
          );
        }

        const payload = {
          source: "web-ui",
          level,
          message,
          context,
        };

        if (level === "warn" || level === "warning") {
          logger.warn("Frontend warning", payload);
        } else if (level === "error") {
          logger.error("Frontend error", payload);
        } else {
          return ErrorHandler.createResponse(
            new Error("Invalid 'level' (use 'warning' or 'error')"),
            400,
          );
        }

        return new Response(JSON.stringify({ success: true }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to ingest frontend log", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/recent-chats") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const scope = url.searchParams.get("scope") || "file";
        const workbookKey = url.searchParams.get("workbookKey") || "";
        const limitParam = url.searchParams.get("limit") || "1";
        const limit = Math.max(1, Math.min(50, Number(limitParam) || 1));

        let items;
        if (scope === "all") {
          items = listRecentChatsAll(limit);
        } else {
          if (!workbookKey) {
            return ErrorHandler.createResponse(
              new Error("Missing 'workbookKey' query parameter"),
              400,
            );
          }
          items = listRecentChats(workbookKey, limit);
        }
        return new Response(JSON.stringify({ items }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to fetch recent chats", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/search-chats") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const scope = url.searchParams.get("scope") || "file";
        const workbookKey = url.searchParams.get("workbookKey") || "";
        const query = url.searchParams.get("q") || "";
        const limitParam = url.searchParams.get("limit") || "20";
        const limit = Math.max(1, Math.min(100, Number(limitParam) || 20));

        if (!query) {
          return ErrorHandler.createResponse(
            new Error("Missing 'q' query parameter"),
            400,
          );
        }

        let items;
        if (scope === "all") {
          items = searchRecentChatsAll(query, limit);
        } else {
          if (!workbookKey) {
            return ErrorHandler.createResponse(
              new Error("Missing 'workbookKey' query parameter"),
              400,
            );
          }
          items = searchRecentChats(workbookKey, query, limit);
        }

        return new Response(JSON.stringify({ items }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to search chats", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/slash-commands") {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const workbookName = url.searchParams.get("workbookName");
        const separated = url.searchParams.get("separated") === "true";

        if (separated) {
          const { baseCommands, excelCommands } = await slashCommandService
            .getCommandsSeparated(workbookName);
          return new Response(JSON.stringify({ baseCommands, excelCommands }), {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          });
        }

        const commands = await slashCommandService.getCommands(workbookName);
        return new Response(JSON.stringify({ commands }), {
          status: 200,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        logger.error("Failed to fetch slash commands", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/workbook-summaries") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, PATCH, DELETE, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method === "GET") {
        try {
          const workbookKey = url.searchParams.get("workbookKey") || "";
          const limitParam = url.searchParams.get("limit") || "50";
          const limit = Math.max(1, Math.min(200, Number(limitParam) || 50));

          if (!workbookKey) {
            return ErrorHandler.createResponse(
              new Error("Missing 'workbookKey' query parameter"),
              400,
            );
          }

          const summaries = getWorkbookSummaries(workbookKey, limit);
          return new Response(JSON.stringify({ summaries }), {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          });
        } catch (error) {
          logger.error("Failed to fetch workbook summaries", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      if (req.method === "PATCH") {
        try {
          const body = await req.json();
          const id = body?.id as string | undefined;
          const summaryText = body?.summaryText as string | undefined;

          if (!id || typeof id !== "string") {
            return ErrorHandler.createResponse(
              new Error("Missing or invalid 'id' in request body"),
              400,
            );
          }

          if (!summaryText || typeof summaryText !== "string") {
            return ErrorHandler.createResponse(
              new Error("Missing or invalid 'summaryText' in request body"),
              400,
            );
          }

          updateWorkbookSummaryText(id, summaryText);

          return new Response(
            JSON.stringify({ success: true }),
            {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            },
          );
        } catch (error) {
          logger.error("Failed to update workbook summary", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      if (req.method === "DELETE") {
        try {
          const body = await req.json().catch(() => ({}));
          const id = (body?.id as string | undefined) ??
            (url.searchParams.get("id") || undefined);

          if (!id || typeof id !== "string") {
            return ErrorHandler.createResponse(
              new Error("Missing or invalid 'id' to delete"),
              400,
            );
          }

          deleteWorkbookSummary(id);

          return new Response(
            JSON.stringify({ success: true }),
            {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            },
          );
        } catch (error) {
          logger.error("Failed to delete workbook summary", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    if (url.pathname === "/api/user-preference") {
      // Handle user preference requests
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method === "GET") {
        try {
          const preferenceKey = url.searchParams.get("key");
          if (!preferenceKey) {
            return ErrorHandler.createResponse(
              new Error("Missing 'key' query parameter"),
              400,
            );
          }
          const preferenceType = (url.searchParams.get("type") || "boolean")
            .toLowerCase();
          const defaultValue = url.searchParams.get("default");
          const { defaultBool, fallbackValue } = parseDefaultValue(
            preferenceType,
            defaultValue,
          );
          const fallbackString = typeof fallbackValue === "string"
            ? fallbackValue
            : "";

          const userService = UserService.getInstance();
          const userEmail = userService.getUserEmail();

          if (!userEmail || userEmail === "Email not set") {
            // Return default value if user not logged in
            return new Response(
              JSON.stringify({
                preferenceKey,
                preferenceValue: preferenceType === "string"
                  ? fallbackString
                  : defaultBool,
              }),
              {
                status: 200,
                headers: {
                  "Access-Control-Allow-Origin": "*",
                  "Content-Type": "application/json",
                },
              },
            );
          }

          const { UserPreferencesService } = await import(
            "../services/user-preferences.ts"
          );
          const prefsService = UserPreferencesService.getInstance();
          if (preferenceType === "string") {
            const value = prefsService.getPreference(preferenceKey);
            return new Response(
              JSON.stringify({
                preferenceKey,
                preferenceValue: value ?? fallbackString,
              }),
              {
                status: 200,
                headers: {
                  "Access-Control-Allow-Origin": "*",
                  "Content-Type": "application/json",
                },
              },
            );
          }
          const value = prefsService.getPreferenceBool(
            preferenceKey,
            defaultBool,
          );

          return new Response(
            JSON.stringify({ preferenceKey, preferenceValue: value }),
            {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            },
          );
        } catch (error) {
          logger.error("Failed to get user preference", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      if (req.method === "POST") {
        try {
          const body = await req.json();
          const { preferenceKey, preferenceValue } = body;

          if (!preferenceKey) {
            return ErrorHandler.createResponse(
              new Error("Missing 'preferenceKey' in request body"),
              400,
            );
          }

          const userService = UserService.getInstance();
          const userEmail = userService.getUserEmail();

          if (!userEmail || userEmail === "Email not set") {
            return ErrorHandler.createResponse(
              new Error("User not logged in"),
              401,
            );
          }

          const { UserPreferencesService } = await import(
            "../services/user-preferences.ts"
          );
          const prefsService = UserPreferencesService.getInstance();

          if (typeof preferenceValue === "boolean") {
            prefsService.setPreferenceBool(preferenceKey, preferenceValue);
          } else {
            prefsService.setPreference(preferenceKey, String(preferenceValue));
          }

          return new Response(
            JSON.stringify({ success: true, preferenceKey, preferenceValue }),
            {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            },
          );
        } catch (error) {
          logger.error("Failed to set user preference", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    if (url.pathname === "/api/workbook-snapshot") {
      // Handle workbook snapshot processing
      // API contract:
      // - isInitialSnapshot: true  -> Initial snapshot (on workbook open), saves baseline
      // - isInitialSnapshot: false -> Consecutive snapshot, generates changelog with summary
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "POST") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const body = await req.json();
        const workbookName = body?.workbookName || "Unknown";
        const currentSnapshot = body?.currentSnapshot;
        const isInitialSnapshot = body?.isInitialSnapshot === true;
        const comparisonSnapshot = body?.comparisonSnapshot;

        if (!currentSnapshot) {
          return ErrorHandler.createResponse(
            new Error("Missing currentSnapshot"),
            400,
          );
        }
        const comparisonFileName = comparisonSnapshot?.filePath?.split(/[\\/]/)
          .pop() ?? comparisonSnapshot?.fileName ?? null;
        const comparisonFileModifiedAt = comparisonSnapshot?.lastModified ??
          null;
        logger.info("Workbook snapshot request received", {
          workbookName,
          isInitialSnapshot,
          hasComparisonSnapshot: !!comparisonSnapshot,
          comparisonFileName,
          comparisonFileModifiedAt,
        });

        // Handle initial snapshot from C# (on workbook open)
        if (isInitialSnapshot) {
          const result = changelogService.handleInitialSnapshot(
            workbookName,
            currentSnapshot,
          );
          return new Response(
            JSON.stringify(result),
            {
              status: 202, // Accepted
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            },
          );
        }

        // For consecutive snapshots: always generate changelog with summary
        // If comparisonSnapshot is provided, it will be used instead of DB snapshot
        const result = (changelogService.generateChangelog as any)(
          workbookName,
          currentSnapshot,
          comparisonSnapshot,
        );

        return new Response(
          JSON.stringify(result),
          {
            status: 202, // Accepted
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          },
        );
      } catch (error) {
        logger.error("Failed to process workbook snapshot", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/workbook-snapshot/exists") {
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "GET") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      const workbookName = url.searchParams.get("workbookName");
      if (!workbookName) {
        return ErrorHandler.createResponse(
          new Error("Missing workbookName"),
          400,
        );
      }

      try {
        const snapshot = getWorkbookSnapshot(workbookName);
        const exists = !!snapshot;
        return new Response(
          JSON.stringify({ exists }),
          {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          },
        );
      } catch (error) {
        logger.error("Failed to check workbook snapshot existence", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/workbook-snapshot/copy") {
      // Handle snapshot copying (e.g., on Save As)
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method !== "POST") {
        return ErrorHandler.createResponse(
          ErrorHandler.createNotFoundError("Method not allowed"),
          405,
        );
      }

      try {
        const body = await req.json();
        const sourceWorkbookName = body?.sourceWorkbookName;
        const targetWorkbookName = body?.targetWorkbookName;

        if (!sourceWorkbookName || !targetWorkbookName) {
          return ErrorHandler.createResponse(
            new Error("Missing sourceWorkbookName or targetWorkbookName"),
            400,
          );
        }

        if (sourceWorkbookName === targetWorkbookName) {
          return ErrorHandler.createResponse(
            new Error("Source and target workbook names must be different"),
            400,
          );
        }

        const success = copyWorkbookSnapshot(
          sourceWorkbookName,
          targetWorkbookName,
        );

        return new Response(
          JSON.stringify({ success }),
          {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          },
        );
      } catch (error) {
        logger.error("Failed to copy workbook snapshot", error);
        return ErrorHandler.createResponse(error, 500);
      }
    }

    if (url.pathname === "/api/llm-config") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      // GET: Retrieve LLM config
      if (req.method === "GET") {
        try {
          const provider = url.searchParams.get("provider") || "azure_openai";
          logger.info("📖 API: Fetching LLM configuration", { provider });

          const config = getLlmKeysConfigByProvider(provider);

          if (!config) {
            logger.info("ℹ️ API: No configuration found", { provider });
            return new Response(JSON.stringify({ config: null }), {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            });
          }

          logger.info("✓ API: Configuration retrieved", {
            provider,
            configId: config.id,
            hasEndpoint: !!config.azureOpenaiEndpoint,
            hasDeployment: !!config.azureOpenaiDeploymentName,
            hasModel: !!config.azureOpenaiModel,
            hasAnthropicModel: !!config.anthropicModel,
            hasAnthropicBaseUrl: !!config.anthropicBaseUrl,
            hasAnthropicCaBundlePath: !!config.anthropicCaBundlePath,
            isActive: config.isActive,
          });

          return new Response(JSON.stringify({ config }), {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          });
        } catch (error) {
          logger.error("❌ API: Failed to fetch LLM config", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      // POST: Save LLM config
      if (req.method === "POST") {
        try {
          const body = await req.json();
          const {
            provider,
            azureOpenaiEndpoint,
            azureOpenaiKey,
            azureOpenaiDeploymentName,
            azureOpenaiModel,
            openaiReasoningEffort,
            anthropicApiKey,
            anthropicModel,
            anthropicBaseUrl,
            anthropicCaBundlePath,
            isActive,
          } = body;

          if (!provider) {
            return ErrorHandler.createResponse(
              new Error("Missing 'provider' field in request body"),
              400,
            );
          }

          if (provider === "azure_openai") {
            // Validate Azure OpenAI endpoint URL format
            if (azureOpenaiEndpoint) {
              try {
                const url = new URL(azureOpenaiEndpoint);
                if (!url.protocol.startsWith("http")) {
                  return ErrorHandler.createResponse(
                    new Error(
                      "Azure OpenAI endpoint must be a valid HTTP(S) URL",
                    ),
                    400,
                  );
                }
              } catch {
                return ErrorHandler.createResponse(
                  new Error("Invalid Azure OpenAI endpoint URL format"),
                  400,
                );
              }
            }

            // Validate deployment name format
            let trimmedDeploymentName = azureOpenaiDeploymentName;
            if (azureOpenaiDeploymentName) {
              trimmedDeploymentName = azureOpenaiDeploymentName.trim();
              if (!trimmedDeploymentName) {
                return ErrorHandler.createResponse(
                  new Error("Deployment name cannot be empty"),
                  400,
                );
              }

              // Allow alphanumeric, hyphen, underscore, and dot; no leading/trailing separators.
              const deploymentNameRegex =
                /^[a-zA-Z0-9](?:[a-zA-Z0-9._-]*[a-zA-Z0-9])?$/;
              if (!deploymentNameRegex.test(trimmedDeploymentName)) {
                return ErrorHandler.createResponse(
                  new Error(
                    "Invalid deployment name. Use letters, numbers, hyphens, underscores, or dots, and avoid leading/trailing separators.",
                  ),
                  400,
                );
              }

              if (trimmedDeploymentName.length > 64) {
                return ErrorHandler.createResponse(
                  new Error("Deployment name must be 64 characters or less"),
                  400,
                );
              }
            }

            // Validate API key is not empty if provided
            if (
              azureOpenaiKey !== undefined && azureOpenaiKey !== null &&
              azureOpenaiKey.trim() === ""
            ) {
              return ErrorHandler.createResponse(
                new Error("Azure OpenAI API key cannot be empty"),
                400,
              );
            }

            // Validate model name is not empty if provided
            if (
              azureOpenaiModel !== undefined && azureOpenaiModel !== null &&
              azureOpenaiModel.trim() === ""
            ) {
              return ErrorHandler.createResponse(
                new Error("Azure OpenAI model cannot be empty"),
                400,
              );
            }

            logger.info("💾 Saving Azure OpenAI configuration:", {
              provider,
              endpoint: azureOpenaiEndpoint || "[not set]",
              deploymentName: azureOpenaiDeploymentName || "[not set]",
              model: azureOpenaiModel || "[not set]",
              modelMatchesDeployment:
                azureOpenaiModel === azureOpenaiDeploymentName,
              hasApiKey: !!azureOpenaiKey,
              apiKeyLength: azureOpenaiKey?.length || 0,
              reasoningEffort: openaiReasoningEffort || "[not set]",
              isActive: isActive !== false,
            });

            const id = upsertLlmKeysConfigByProvider({
              provider,
              azureOpenaiEndpoint: azureOpenaiEndpoint || null,
              azureOpenaiKey: azureOpenaiKey || null,
              azureOpenaiDeploymentName: trimmedDeploymentName || null,
              azureOpenaiModel: azureOpenaiModel || null,
              openaiReasoningEffort: openaiReasoningEffort || null,
              anthropicApiKey: null,
              anthropicModel: null,
              anthropicBaseUrl: null,
              anthropicCaBundlePath: null,
              isActive: isActive !== false,
            });

            logger.info("✓ Azure OpenAI configuration saved successfully", {
              configId: id,
              provider,
            });

            return new Response(JSON.stringify({ success: true, id }), {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            });
          }

          if (provider === "anthropic") {
            const trimmedBaseUrl = typeof anthropicBaseUrl === "string"
              ? anthropicBaseUrl.trim()
              : "";
            const trimmedCaBundlePath =
              typeof anthropicCaBundlePath === "string"
                ? anthropicCaBundlePath.trim()
                : "";

            if (trimmedBaseUrl) {
              const hasScheme = /^https?:\/\//i.test(trimmedBaseUrl);
              if (hasScheme) {
                try {
                  const url = new URL(trimmedBaseUrl);
                  if (!url.protocol.startsWith("http")) {
                    return ErrorHandler.createResponse(
                      new Error(
                        "Anthropic base URL must be a valid HTTP(S) URL",
                      ),
                      400,
                    );
                  }
                } catch {
                  return ErrorHandler.createResponse(
                    new Error("Invalid Anthropic base URL format"),
                    400,
                  );
                }
              }
            }

            if (trimmedCaBundlePath) {
              if (!/\.(pem|crt|cer)$/i.test(trimmedCaBundlePath)) {
                return ErrorHandler.createResponse(
                  new Error(
                    "Anthropic CA bundle must be a .pem, .crt, or .cer file",
                  ),
                  400,
                );
              }

              try {
                const info = await Deno.stat(trimmedCaBundlePath);
                if (!info.isFile) {
                  return ErrorHandler.createResponse(
                    new Error(
                      "Anthropic CA bundle path must point to a file",
                    ),
                    400,
                  );
                }
              } catch (_error) {
                return ErrorHandler.createResponse(
                  new Error(
                    "Anthropic CA bundle file was not found or is not readable",
                  ),
                  400,
                );
              }
            }

            if (
              anthropicApiKey !== undefined && anthropicApiKey !== null &&
              anthropicApiKey.trim() === ""
            ) {
              return ErrorHandler.createResponse(
                new Error("Anthropic API key cannot be empty"),
                400,
              );
            }

            if (
              anthropicModel !== undefined && anthropicModel !== null &&
              anthropicModel.trim() === ""
            ) {
              return ErrorHandler.createResponse(
                new Error("Anthropic model cannot be empty"),
                400,
              );
            }

            if (isActive !== false && !anthropicApiKey) {
              return ErrorHandler.createResponse(
                new Error(
                  "Anthropic API key is required when enabling provider",
                ),
                400,
              );
            }

            logger.info("💾 Saving Anthropic configuration:", {
              provider,
              model: anthropicModel || "[not set]",
              baseUrl: trimmedBaseUrl || "[default]",
              hasCaBundlePath: !!trimmedCaBundlePath,
              hasApiKey: !!anthropicApiKey,
              apiKeyLength: anthropicApiKey?.length || 0,
              isActive: isActive !== false,
            });

            const id = upsertLlmKeysConfigByProvider({
              provider,
              azureOpenaiEndpoint: null,
              azureOpenaiKey: null,
              azureOpenaiDeploymentName: null,
              azureOpenaiModel: null,
              openaiReasoningEffort: null,
              anthropicApiKey: anthropicApiKey || null,
              anthropicModel: anthropicModel || null,
              anthropicBaseUrl: trimmedBaseUrl || null,
              anthropicCaBundlePath: trimmedCaBundlePath || null,
              isActive: isActive !== false,
            });

            logger.info("✓ Anthropic configuration saved successfully", {
              configId: id,
              provider,
            });

            return new Response(JSON.stringify({ success: true, id }), {
              status: 200,
              headers: {
                "Access-Control-Allow-Origin": "*",
                "Content-Type": "application/json",
              },
            });
          }

          return ErrorHandler.createResponse(
            new Error(`Unsupported provider: ${provider}`),
            400,
          );
        } catch (error) {
          logger.error("Failed to save LLM config", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    // POST: Reload LLM configuration
    if (url.pathname === "/api/llm-config/reload") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method === "POST") {
        try {
          logger.info("🔄 API: Reloading LLM configuration requested");
          reloadLLMConfig();
          logger.info("✓ API: LLM configuration reload complete");

          return new Response(JSON.stringify({ success: true }), {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          });
        } catch (error) {
          logger.error("✗ API: Failed to reload LLM config", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    // POST: Test LLM connection
    if (url.pathname === "/api/llm-config/test") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method === "POST") {
        try {
          const body = await req.json().catch(() => ({}));
          const provider = body?.provider || "azure_openai";

          logger.info("🔍 API: Testing LLM connection", { provider });

          if (provider === "azure_openai") {
            const config = getAzureOpenAIConfig();
            if (!config) {
              return new Response(
                JSON.stringify({
                  success: false,
                  error: "No Azure OpenAI configuration found",
                  hint: "Please configure Azure OpenAI settings first",
                }),
                {
                  status: 400,
                  headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json",
                  },
                },
              );
            }

            const baseURL = `${config.endpoint}/openai/v1/`;

            logger.info("🔍 Connection test details:", {
              endpoint: config.endpoint,
              baseURL: baseURL,
              fullURL: `${baseURL}responses`,
              deploymentName: config.deploymentName,
              model: config.model,
              hasApiKey: !!config.apiKey,
              apiKeyLength: config.apiKey.length,
            });

            return new Response(
              JSON.stringify({
                success: true,
                message:
                  "Configuration loaded. Check server logs for connection details.",
                config: {
                  endpoint: config.endpoint,
                  baseURL: baseURL,
                  fullURL: `${baseURL}responses`,
                  deploymentName: config.deploymentName,
                  model: config.model,
                  hasApiKey: !!config.apiKey,
                  apiKeyLength: config.apiKey.length,
                  apiKeyPrefix: config.apiKey.substring(0, 4) + "...",
                },
              }),
              {
                status: 200,
                headers: {
                  "Access-Control-Allow-Origin": "*",
                  "Content-Type": "application/json",
                },
              },
            );
          }

          if (provider === "anthropic") {
            const config = getAnthropicConfig();
            if (!config) {
              return new Response(
                JSON.stringify({
                  success: false,
                  error: "No Anthropic configuration found",
                  hint: "Please configure Anthropic settings first",
                }),
                {
                  status: 400,
                  headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json",
                  },
                },
              );
            }

            try {
              const claudeClient = createAnthropicClient();
              const testResponse = await claudeClient.messages.create({
                model: config.model,
                max_tokens: 1,
                messages: [{
                  role: "user",
                  content: "ping",
                }],
              });

              return new Response(
                JSON.stringify({
                  success: true,
                  message: "Configuration loaded. Message request succeeded.",
                  config: {
                    model: config.model,
                    baseUrl: config.baseUrl || null,
                    caBundlePath: config.caBundlePath || null,
                    hasApiKey: !!config.apiKey,
                    apiKeyLength: config.apiKey.length,
                    apiKeyPrefix: config.apiKey.substring(0, 4) + "...",
                    responseId: testResponse.id,
                  },
                }),
                {
                  status: 200,
                  headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json",
                  },
                },
              );
            } catch (error: any) {
              const status = typeof error?.status === "number"
                ? error.status
                : 502;
              const details = error?.error && typeof error.error === "object"
                ? error.error
                : undefined;

              logger.error("❌ API: Anthropic connection test failed", {
                error: error?.message,
                status,
                requestId: error?.requestID,
                details,
              });

              return new Response(
                JSON.stringify({
                  success: false,
                  error: error?.message || "Anthropic connection test failed",
                  status,
                  requestId: error?.requestID,
                  details,
                }),
                {
                  status,
                  headers: {
                    "Access-Control-Allow-Origin": "*",
                    "Content-Type": "application/json",
                  },
                },
              );
            }
          }

          return ErrorHandler.createResponse(
            new Error(`Unsupported provider: ${provider}`),
            400,
          );
        } catch (error: any) {
          logger.error("❌ API: Connection test failed", {
            error: error.message,
            stack: error.stack,
          });
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    // GET: Available tools
    if (url.pathname === "/api/tools") {
      // Simple CORS support
      if (req.method === "OPTIONS") {
        return new Response(null, {
          status: 204,
          headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type",
          },
        });
      }

      if (req.method === "GET") {
        try {
          const { tools } = await import("../tools/index.ts");
          const toolsList = tools.map((tool) => ({
            id: tool.name,
            name: tool.name
              .split("_")
              .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
              .join(" "),
            description: tool.description || "",
          }));

          return new Response(JSON.stringify({ tools: toolsList }), {
            status: 200,
            headers: {
              "Access-Control-Allow-Origin": "*",
              "Content-Type": "application/json",
            },
          });
        } catch (error) {
          logger.error("Failed to get tools", error);
          return ErrorHandler.createResponse(error, 500);
        }
      }

      return ErrorHandler.createResponse(
        ErrorHandler.createNotFoundError("Method not allowed"),
        405,
      );
    }

    return ErrorHandler.createResponse(
      ErrorHandler.createNotFoundError("Endpoint not found"),
      404,
    );
  }

  handleWebSocketRequest(req: Request): Response {
    try {
      const { socket, response } = (globalThis as any).Deno.upgradeWebSocket(
        req,
      );

      let currentWorkbookName: string | null = null;

      socket.onopen = () => {
        logger.info("WebSocket connection opened");
        socket.send(JSON.stringify({ type: "welcome", message: "Connected" }));
      };

      socket.onmessage = async (event: MessageEvent) => {
        try {
          const dataRaw = typeof event.data === "string" ? event.data : "";
          const data = dataRaw ? JSON.parse(dataRaw) : {};
          logger.info("WebSocket message received", { data });

          const type = data?.type;
          const workbookName = data?.workbookName;

          // Register socket for workbook when we first see the workbook name
          if (workbookName && !currentWorkbookName) {
            currentWorkbookName = workbookName;
            WebSocketManager.getInstance().setSocket(workbookName, socket);
            logger.info(`Registered socket for workbook: ${workbookName}`);
          }

          switch (type) {
            case "workbook:register":
              logger.info("Workbook registration confirmed", { workbookName });
              return;

            case "stream:start":
              return await this.handleStartStream(data);

            case "stream:cancel":
              return await this.handleCancelStream(data);

            case "chat:restart":
              return await this.handleRestartState(data);

            case "tool:revert":
              return await this.handleRevertTool(data);

            case "tool:apply":
              return await this.handleApplyTool(data);

            case "tool:permission:response":
              return await this.handleToolPermissionResponse(data);

            case "tool:result":
              return await this.handleToolResult(data);

            case "tool:view":
              return await this.handleToolView(data);

            case "tool:unview":
              return await this.handleToolUnview(data);

            case "score:score":
              return await this.handleScore(data);

            case "score:feedback":
              return await this.handleFeedback(data);

            case "user:email_response":
              return await this.handleEmailResponse(data);

            default:
              return;
          }
        } catch (error: any) {
          if (error instanceof UserCancellationError) {
            const requestId = error.requestId;
            logger.info("User cancelled the request", {
              currentWorkbookName,
              requestId,
            });

            // Safely log cancellation span - don't crash if trace is missing
            try {
              tracing.logUserCancellationSpan(requestId);
              tracing.sendScore(requestId, -0.5);
              logger.info("Sent cancellation score to Langfuse", {
                requestId,
                score: -0.5,
              });
            } catch (tracingError) {
              logger.warn("Failed to log cancellation to tracing service", {
                requestId,
                error: tracingError,
              });
            }
            return;
          }

          if (error instanceof ToolExecutionError) {
            logger.warn(
              `Tool ${error.toolId} errored during execution, sending error to client`,
            );
            logger.error("Tool execution error details", {
              toolId: error.toolId,
              toolName: error.toolName,
              errorMessage: error.message,
              errorCode: (error.originalError as any)?.code,
              originalError: error.originalError,
            });

            const toolErrorPayload = {
              error: JSON.stringify(error.originalError),
              errorMessage: error.message,
              toolId: error.toolId,
              errorCode: (error.originalError as any)?.code,
            };

            if (currentWorkbookName) {
              WebSocketManager.getInstance().send(
                currentWorkbookName,
                WebSocketMessageTypes.TOOL_ERROR,
                toolErrorPayload,
              );
            }

            return;
          }

          logger.info("errorClass: ", error.constructor.name);

          // Handle other errors (stream errors, etc.)
          const errorMessage = error?.message;
          const errorType = extractClaudeErrorType(errorMessage);

          const payload = mapClaudeErrorToPayload(errorType);

          logger.error("Error handling WebSocket message", {
            currentWorkbookName,
            errorMessage,
            errorType,
            payload,
          });

          // Use errorStream() to ensure correct message ordering:
          // stream:error BEFORE idle (matches endStream() and cancelStream() pattern)
          // This prevents race conditions where idle arrives before error
          if (currentWorkbookName) {
            WebSocketManager.getInstance().errorStream(
              currentWorkbookName,
              payload,
            );
          } else {
            // Fallback if no workbook name (shouldn't happen, but be safe)
            socket.send(
              JSON.stringify({ type: "stream:error", payload: payload }),
            );
          }
        }
      };

      socket.onclose = (event: CloseEvent) => {
        logger.info("WebSocket connection closed", {
          code: event.code,
          reason: event.reason,
        });
        if (currentWorkbookName) {
          logger.info(
            `Cleaning up connection for workbook: ${currentWorkbookName}`,
          );
          // WebSocketManager will auto-cleanup via event listeners set in setSocket()
        }
      };

      socket.onerror = (event: Event | ErrorEvent) => {
        logger.error("WebSocket error", event);
      };

      return response;
    } catch (error) {
      logger.error("WebSocket upgrade failed", error);
      return ErrorHandler.createResponse(error);
    }
  }

  async handleToolUnview(data: any) {
    const workbookName = data?.workbookName as string;
    const toolId = data?.toolId as string;

    logger.info("tool id", toolId);
    logger.info("getting tool change", { workbookName, toolId });
    const toolChange: ToolChangeInterface = stateManager.getToolChange(
      workbookName,
      toolId,
    );

    try {
      await revertSingleToolChange(toolChange, true);
      logger.info(
        `Successfully reverted tool ${toolChange.toolName} during unview`,
      );

      // Send idle status after unview revert completes
      // (VSTO progress updates may have been sent during the operation)
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
    } catch (error) {
      logger.error(
        `Failed to revert tool ${toolChange.toolName} during unview: ${error}`,
      );
      // Don't re-throw - log and continue gracefully
      // Still send idle even on error to clear any progress status
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
    }
  }

  private async handleToolView(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const toolId = data?.toolId as string;

    logger.info("WebSocket tool view message received", {
      workbookName,
      toolId,
    });
    const toolChange: ToolChangeInterface = stateManager.getToolChange(
      workbookName,
      toolId,
    );

    logger.info("checking if tool is already approved");
    if (toolChange.approved) {
      throw new Error("Tool is already approved");
    }

    logger.info("checking tool is pending");
    if (!toolChange.pending) {
      throw new Error("Tool is not pending");
    }

    logger.info("getting request id");
    const requestId = stateManager.getLatestRequestId(workbookName);

    logger.info("tool execute flow");
    const result = await toolExecutionFlow(
      requestId,
      toolChange.toolName,
      toolChange.inputData,
      toolChange,
      workbookName,
    );

    logger.info("getting input data revert");
    const inputDataRevert = getInputDataRevert(
      result,
      workbookName,
      toolChange,
    );
    logger.info("input data revert obtained", {
      inputDataRevertKeys: Object.keys(inputDataRevert || {}),
    });

    logger.info("updating tool change");
    try {
      const partialToolChange = {
        ...toolChange,
        inputDataRevert: inputDataRevert,
        approved: false,
        applied: true,
        pending: true,
        result: result,
      };
      logger.info("partial tool change created", {
        toolId,
        hasInputDataRevert: !!partialToolChange.inputDataRevert,
      });

      stateManager.updateToolChange(workbookName, toolId, partialToolChange);
      logger.info("tool change updated successfully", { toolId });

      // Send success message back to client
      WebSocketManager.getInstance().send(workbookName, "tool:view:success", {
        toolId,
        message: "Tool preview applied successfully",
      });
      logger.info("sent tool:view:success message to client", { toolId });

      // Send idle status after tool view completes
      // (toolExecutionFlow sends "generating_llm" but we're just previewing, not continuing)
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
    } catch (error) {
      logger.error("Error updating tool change", { error, toolId });

      // Send error message back to client
      const errorMessage = error instanceof Error
        ? error.message
        : "Unknown error during tool view";
      WebSocketManager.getInstance().send(workbookName, "tool:view:error", {
        toolId,
        error: errorMessage,
      });

      // Send idle status even on error to clear any progress status
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );

      throw error;
    }
  }

  private handleEmailResponse(data: any) {
    const userService = UserService.getInstance();
    userService.setUserEmail(data.email);
  }

  private async handleStartStream(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const workbookPath = data?.workbookPath as string | undefined;
    const payload = data?.payload as any;

    logger.info("Getting or creating state", { workbookName });
    stateManager.getOrCreateState(workbookName);

    if (workbookPath && workbookPath.length > 0) {
      try {
        stateManager.setWorkbookKey(workbookName, workbookPath);
        logger.info("Stored workbookKey for workbook", {
          workbookName,
          workbookPath,
        });
      } catch (e) {
        logger.warn("Failed to set workbookKey", {
          workbookName,
          workbookPath,
          error: (e as Error)?.message,
        });
      }
    }

    logger.info("Setting latest request id", { workbookName });
    const requestId = stateManager.setLatestRequestId(workbookName);

    logger.info("Creating abort controller");
    stateManager.creatingAbortController(workbookName, requestId);

    logger.info("Getting session id", { workbookName });
    const sessionId = stateManager.getSessionId(workbookName);

    logger.info("Saving request metadata");
    const requestMetadata = {
      activeTools: payload.activeTools,
      workbookName: workbookName,
      worksheets: payload.worksheets,
      activeWorksheet: payload.activeWorksheet,
    };
    stateManager.saveRequestMetadata(workbookName, requestId, requestMetadata);

    logger.info("Getting user email");
    const userEmail = UserService.getInstance().getUserEmail();

    // Extract attachment information for Langfuse metadata
    const attachmentInfo = extractAttachmentMetadata(payload);

    const metadata = {
      name: "Excel AI Workflow",
      userEmail: userEmail,
      input: payload.prompt,
      sessionId: sessionId,
      ...attachmentInfo,
    };

    logger.info("Starting trace", { requestId, metadata });
    tracing.startTrace(requestId, metadata, workbookName);

    logger.info("Add Documents if attached");
    const documentEntries = parseDocumentFromRequest(payload);
    if (documentEntries && documentEntries?.length > 0) {
      for (const entry of documentEntries) {
        stateManager.addMessage(workbookName, entry, { persist: false });
      }
    }

    const attachmentSummaries = Array.isArray(payload.documentsBase64)
      ? payload.documentsBase64.map((doc: any) => ({
        name: String(doc?.name ?? ""),
        type: String(doc?.type ?? ""),
        size: Number(doc?.size ?? 0),
      })).filter((doc: { name: string; type: string; size: number }) =>
        Boolean(doc.name || doc.type || doc.size)
      )
      : [];

    const prompt = await createPrompt(payload.prompt, {
      allowOtherWorkbooks: payload.allowOtherWorkbookReads === true,
      activeWorkbookName: workbookName,
      attachments: attachmentSummaries.length > 0
        ? attachmentSummaries
        : undefined,
    });

    logger.info("Adding message to conversation history");
    stateManager.addMessage(workbookName, prompt);

    logger.info("Creating active tools");
    const activeTools = payload.activeTools ||
      [
        ToolNames.READ_VALUES_BATCH,
        ToolNames.READ_FORMAT_BATCH,
        ToolNames.GET_METADATA,
        ToolNames.LIST_SHEETS,
        ToolNames.LIST_OPEN_WORKBOOKS,
      ];

    const llmProvider = getLLMProvider();
    logger.info("Creating Claude client / provider", { llmProvider });
    const client = createAnthropicClient();

    logger.info("Get Conversation History");
    let conversationHistory = stateManager.getConversationHistory(workbookName);

    let exceedsTokenLimit = false;
    if (llmProvider === "anthropic") {
      logger.info("Checking if conversation history exceeds token limit");
      exceedsTokenLimit = await isHistoryExceedingTokenLimit(
        conversationHistory,
        tokenLimit,
        activeTools,
      );
    }
    if (exceedsTokenLimit) {
      conversationHistory = await this.handleConversationCompacting(
        workbookName,
        requestId,
        conversationHistory,
        activeTools,
      );
    }

    logger.info("Building Claude params");
    const llmParams = buildAnthropicParamsForConversation(
      conversationHistory,
      activeTools,
    );

    logger.info("Streaming Claude response and handling tool usage");
    await streamClaudeResponseAndHandleToolUsage(requestId, client, llmParams);
  }

  /**
   * Handle conversation compacting with tracing
   */
  private async handleConversationCompacting(
    workbookName: string,
    requestId: string,
    conversationHistory: Anthropic.MessageParam[],
    activeTools: string[],
  ): Promise<Anthropic.MessageParam[]> {
    logger.info("Conversation history exceeds token limit");

    // Get initial token count
    const claudeClient = createAnthropicClient();
    const initialTokenParams = buildAnthropicParamsForTokenCount(
      conversationHistory,
      activeTools,
    );
    const initialTokenResponse = await claudeClient.beta.messages.countTokens(
      withAnthropicBetas(initialTokenParams),
    ) as Anthropic.MessageTokensCount;
    const initialTokenCount = initialTokenResponse.input_tokens;

    // Log compacting process to Langfuse
    const compactingSpanId = tracing.logCompactingStart(requestId, {
      originalMessageCount: conversationHistory.length,
      initialTokenCount: initialTokenCount,
      tokenLimit: tokenLimit,
    });

    // Get current request metadata for compacting
    const currentRequestMetadata = stateManager.getRequestMetadata(
      workbookName,
      requestId,
    );
    const compactedHistory = await compactChatHistory(
      currentRequestMetadata,
      conversationHistory,
    );

    // Get final token count after compacting
    const finalTokenParams = buildAnthropicParamsForTokenCount(
      compactedHistory,
      activeTools,
    );
    const finalTokenResponse = await claudeClient.beta.messages.countTokens(
      withAnthropicBetas(finalTokenParams),
    ) as Anthropic.MessageTokensCount;
    const finalTokenCount = finalTokenResponse.input_tokens;

    tracing.logCompactingEnd(compactingSpanId, {
      compactedMessageCount: compactedHistory.length,
      wasCompacted: true,
      finalTokenCount: finalTokenCount,
      tokenReduction: initialTokenCount - finalTokenCount,
      compressionRatio: (finalTokenCount / initialTokenCount * 100).toFixed(2) +
        "%",
    });

    return compactedHistory;
  }

  handleCancelStream(data: any): void {
    const workbookName = data?.workbookName as string;
    logger.info("WebSocket cancel stream message received", { workbookName });

    try {
      const requestId = stateManager.getLatestRequestId(workbookName);
      stateManager.abortAbortController(workbookName, requestId);

      // Send cancellation messages immediately with correct ordering
      // The abort will also trigger UserCancellationError in the stream,
      // but we send the cancellation message here to ensure immediate feedback
      WebSocketManager.getInstance().cancelStream(
        workbookName,
        "Request was cancelled by user",
        requestId,
      );
    } catch (error) {
      logger.info("No workbook state found to cancel", { workbookName, error });
      // If there's no active request, we can safely send idle directly
      // since there's no stream to cancel
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
      return;
    }
  }

  private handleRestartState(data: any): void {
    const workbookName = data?.workbookName as string;
    logger.info("WebSocket restart state message received", { workbookName });

    if (workbookName) {
      this.handleCancelStream(data);
      stateManager.deleteWorkbookState(workbookName);
    }
  }

  private async handleApplyTool(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const toolId = data?.toolUseId as string;

    logger.info("WebSocket apply tool message received", workbookName);
    logger.info("tool id", toolId);

    if (!workbookName || !toolId) {
      throw new Error("Missing workbookName or toolId field");
    }

    logger.info("Getting user email");
    const userEmail = UserService.getInstance().getUserEmail();

    logger.info("Getting session id");
    const sessionId = stateManager.getSessionId(workbookName);

    const traceMetadata = {
      name: "Excel Re-Apply Operation",
      userEmail: userEmail,
      toolUseId: toolId,
      workbookName: workbookName,
      sessionId: sessionId,
      input: {
        workbookName: workbookName,
        toolId: toolId,
      },
    };

    logger.info("Starting trace", { traceMetadata });
    const traceId = crypto.randomUUID();
    tracing.startTrace(traceId, traceMetadata, workbookName);

    try {
      logger.info("Applying tool from tool id", { toolId, workbookName });
      const result = await applyFromToolIdOnwards(workbookName, toolId);

      logger.info("Updating History");
      const applyMessage: Anthropic.MessageParam = {
        role: "user",
        content:
          `All tool calls until ${toolId} included, have been applied by the user`,
      };
      stateManager.addMessage(workbookName, applyMessage);

      logger.info("Apply tool result", { result });
      tracing.endTrace(traceId, {
        output: {
          success: true,
          appliedToolChanges: result.map((change) => ({
            toolId: change.toolId,
            toolName: change.toolName,
          })),
        },
      });

      // Send idle status after apply operation completes
      // (VSTO progress updates may have been sent during the operation)
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
    } catch (error: any) {
      tracing.endTrace(traceId, {
        endTime: new Date().toISOString(),
        output: {
          success: false,
          error: error.message,
          stack: error.stack,
        },
      });
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        "error",
        error?.message || "Failed to apply tool changes",
      );
      throw error;
    }
  }

  private async handleRevertTool(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const toolId = data?.toolUseId as string;

    logger.info("WebSocket revert tool message received", workbookName);
    logger.info("tool id", toolId);

    if (!workbookName || !toolId) {
      throw new Error("Missing workbookName or toolId field");
    }

    logger.info("Getting user email");
    const userEmail = UserService.getInstance().getUserEmail();

    logger.info("Getting session id");
    const sessionId = stateManager.getSessionId(workbookName);

    const traceMetadata = {
      name: "Excel Revert Operation",
      userEmail: userEmail,
      toolUseId: toolId,
      workbookName: workbookName,
      sessionId: sessionId,
      input: {
        workbookName: workbookName,
        toolId: toolId,
      },
    };

    logger.info("Starting trace", { traceMetadata });
    const traceId = crypto.randomUUID();
    tracing.startTrace(traceId, traceMetadata, workbookName);

    try {
      logger.info("Reverting tool from tool id", { toolId, workbookName });
      const result = await revertToolChangesForWorkbookFromToolIdOnwards(
        workbookName,
        toolId,
      );

      logger.info("Updating History");
      const revertMessage: Anthropic.MessageParam = {
        role: "user",
        content:
          `All tool calls until ${toolId} included, have been reverted by the user`,
      };

      stateManager.addMessage(workbookName, revertMessage);

      logger.info("Revert tool result", { result });
      tracing.endTrace(traceId, {
        output: {
          success: true,
          revertedToolChanges: result.map((change) => ({
            toolId: change.toolId,
            toolName: change.toolName,
          })),
        },
      });

      // Send idle status after revert operation completes
      // (VSTO progress updates may have been sent during the operation)
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        OperationStatusValues.IDLE,
      );
    } catch (error: any) {
      tracing.endTrace(traceId, {
        endTime: new Date().toISOString(),
        output: {
          success: false,
          error: error.message,
          stack: error.stack,
        },
      });
      WebSocketManager.getInstance().sendStatus(
        workbookName,
        "error",
        error?.message || "Failed to revert tool changes",
      );
      throw error;
    }
  }

  private handleFeedback(data: any): void {
    const workbookName = data?.workbookName as string;
    const requestId = stateManager.getLatestRequestId(workbookName);
    const comment = data?.comment as string;

    logger.info("Sending feedback", { requestId, comment });
    tracing.sendFeedback(requestId, comment);
  }

  private handleScore(data: any): void {
    const workbookName = data?.workbookName as string;
    const requestId = stateManager.getLatestRequestId(workbookName);
    const score = data?.score as number;

    logger.info("Sending score", { requestId, score });
    tracing.sendScore(requestId, score);
  }

  private async handleToolResult(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const toolUseId = data?.toolUseId as string;
    const output = data?.output;

    logger.info("Handling tool:result message", {
      workbookName,
      toolUseId,
      output,
    });

    if (!workbookName || !toolUseId) {
      throw new Error("Missing workbookName or toolUseId for tool:result");
    }

    const toolChange = stateManager.getToolChange(workbookName, toolUseId);
    if (toolChange.toolName !== ToolNames.COLLECT_INPUT) {
      throw new Error(
        `tool:result is only supported for collect_input, received ${toolChange.toolName}`,
      );
    }

    const controls = toolChange.inputData?.controls;
    const validated = validateUiRequestResult(controls, output);
    logger.info("tool:result validated payload", {
      workbookName,
      toolUseId,
      validated,
    });

    stateManager.updateToolResponseInConversationHistory(
      workbookName,
      validated,
      toolChange.toolId,
    );

    const updatedToolChange = {
      ...toolChange,
      approved: true,
      applied: true,
      pending: false,
      result: validated,
    };
    stateManager.updateToolChange(workbookName, toolUseId, updatedToolChange);

    WebSocketManager.getInstance().sendStatus(
      workbookName,
      OperationStatusValues.GENERATING_LLM,
      "Processing your input...",
    );

    await this.continueStreamingAfterToolProcessing(
      workbookName,
      toolChange.requestId,
    );
  }

  /**
   * Common method to continue streaming after tool processing
   * This handles the LLM continuation logic that's shared between
   * tool approval, rejection, and permission response handling
   */
  private async continueStreamingAfterToolProcessing(
    workbookName: string,
    requestId: string,
  ): Promise<void> {
    logger.info("Continuing streaming after tool processing", {
      workbookName,
      requestId,
    });

    const abortController = stateManager.getAbortController(
      workbookName,
      requestId,
    );
    if (!abortController || abortController.signal.aborted) {
      logger.info("Request was cancelled, not continuing stream", {
        workbookName,
        requestId,
      });
      throw new UserCancellationError("User cancelled the request", requestId);
    }

    const llmProvider = getLLMProvider();
    logger.info("Creating Claude client");
    const client = createAnthropicClient();

    logger.info("Getting active tools");
    const activeTools =
      stateManager.getRequestMetadata(workbookName, requestId).activeTools ||
      [ToolNames.READ_VALUES_BATCH];

    logger.info("Getting conversation history");
    let conversationHistory = stateManager.getConversationHistory(workbookName);

    let exceedsTokenLimit = false;
    if (llmProvider === "anthropic") {
      logger.info("Checking if conversation history exceeds token limit");
      exceedsTokenLimit = await isHistoryExceedingTokenLimit(
        conversationHistory,
        tokenLimit,
        activeTools,
      );
    }
    if (exceedsTokenLimit) {
      conversationHistory = await this.handleConversationCompacting(
        workbookName,
        requestId,
        conversationHistory,
        activeTools,
      );
    }

    logger.info("Building Claude params");
    const llmParams = buildAnthropicParamsForConversation(
      conversationHistory,
      activeTools,
    );

    logger.info("Streaming Claude response and handling tool usage");
    await streamClaudeResponseAndHandleToolUsage(requestId, client, llmParams);
  }

  private async handleToolPermissionResponse(data: any): Promise<void> {
    const workbookName = data?.workbookName as string;
    const groupId = data?.groupId as string;
    const toolResponses = data?.toolResponses as Array<
      {
        toolId: string;
        decision: "approved" | "rejected";
        userMessage?: string;
      }
    >;

    logger.info("Handling tool permission response", {
      workbookName,
      groupId,
      toolResponses,
    });

    if (!toolResponses || !Array.isArray(toolResponses)) {
      throw new Error("Invalid tool responses provided");
    }

    const requestId = stateManager.getLatestRequestId(workbookName);

    const abortController = stateManager.getAbortController(
      workbookName,
      requestId,
    );
    if (!abortController || abortController.signal.aborted) {
      logger.info("Request was cancelled, ignoring tool permission response", {
        workbookName,
        requestId,
      });
      throw new UserCancellationError("User cancelled the request", requestId);
    }

    const approvedTools: string[] = [];
    const rejectedTools: string[] = [];
    const failedTools: Array<{ toolId: string; error: ToolExecutionError }> =
      [];

    // Process each tool response by leveraging existing logic
    for (const response of toolResponses) {
      const toolId = response.toolId;
      const decision = response.decision;

      logger.info("Processing tool response", { toolId, decision });

      const toolChange: ToolChangeInterface = stateManager.getToolChange(
        workbookName,
        toolId,
      );

      if (decision === "approved") {
        try {
          // Use the same logic as handleToolApproval but without streaming continuation
          await this.processToolApproval(
            workbookName,
            toolId,
            requestId,
            toolChange,
          );

          // Only add to approved tools if execution succeeded
          approvedTools.push(toolId);
        } catch (error: any) {
          if (error instanceof ToolExecutionError) {
            logger.warn(
              `Tool ${toolId} failed during execution, will process after batch`,
            );
            logger.error("Tool execution error details", {
              toolId: error.toolId,
              toolName: error.toolName,
              errorMessage: error.message,
              errorCode: (error.originalError as any)?.code,
              originalError: error.originalError,
            });
            failedTools.push({ toolId, error });

            // Send error to client
            const toolErrorPayload = {
              error: JSON.stringify(error.originalError),
              errorMessage: error.message,
              toolId: error.toolId,
              errorCode: (error.originalError as any)?.code,
            };
            WebSocketManager.getInstance().send(
              workbookName,
              WebSocketMessageTypes.TOOL_ERROR,
              toolErrorPayload,
            );
          } else {
            // Re-throw non-tool errors
            throw error;
          }
        }
      } else {
        rejectedTools.push(toolId);
        const userMessage = response.userMessage || "Rejected by the user";
        // Use the same logic as handleToolReject but without streaming continuation
        await this.processToolRejection(
          workbookName,
          toolId,
          requestId,
          toolChange,
          userMessage,
        );
      }
    }

    // Handle failed tools after processing all tools in batch
    for (const { toolId, error } of failedTools) {
      logger.info(`Handling failed tool ${toolId} after batch processing`);

      // Update conversation history with the error so LLM can see it and retry
      const errorMessage =
        `Tool execution failed: ${error.message}\nOriginal error: ${
          error.originalError?.message || "Unknown error"
        }`;
      stateManager.updateToolResponseInConversationHistory(
        workbookName,
        errorMessage,
        error.toolId,
      );

      // Mark tool as failed in state
      const toolChange = stateManager.getToolChange(workbookName, error.toolId);
      const updatedToolChange = {
        ...toolChange,
        approved: false,
        applied: false,
        pending: false,
        result: errorMessage,
      };
      stateManager.updateToolChange(
        workbookName,
        error.toolId,
        updatedToolChange,
      );
    }

    logger.info("Tool permission response processed", {
      groupId,
      approvedCount: approvedTools.length,
      rejectedCount: rejectedTools.length,
      failedCount: failedTools.length,
    });

    try {
      // Only continue streaming if there are approved tools or failed tools
      // If ALL tools were rejected, we should stop to avoid infinite loops
      const shouldContinue = approvedTools.length > 0 || failedTools.length > 0;

      if (shouldContinue) {
        logger.info("Continuing with LLM processing after tool responses", {
          approvedCount: approvedTools.length,
          rejectedCount: rejectedTools.length,
          failedCount: failedTools.length,
        });

        // Continue streaming using the common method
        await this.continueStreamingAfterToolProcessing(
          workbookName,
          requestId,
        );
      } else {
        // All tools were rejected or no tools to process - end the stream
        logger.info("All tools rejected or no tools processed, ending stream", {
          approvedCount: approvedTools.length,
          rejectedCount: rejectedTools.length,
          failedCount: failedTools.length,
        });

        const metadata = {
          endTime: new Date().toISOString(),
        };
        tracing.endTrace(requestId, metadata);

        const socket = WebSocketManager.getInstance();
        // Use endStream() to safely send stream:end followed by idle status
        // This ensures correct message ordering to avoid race conditions
        socket.endStream(workbookName, {}, null);

        stateManager.deleteRequestMetadata(workbookName, requestId);
      }
    } catch (error: any) {
      tracing.endTrace(requestId, {
        endTime: new Date().toISOString(),
        output: {
          success: false,
          error: error.message,
          stack: error.stack,
        },
      });
      throw error;
    }
  }

  /**
   * Process tool approval without streaming continuation
   * Extracted from handleToolApproval for reuse
   */
  private async processToolApproval(
    workbookName: string,
    toolId: string,
    requestId: string,
    toolChange: ToolChangeInterface,
  ): Promise<void> {
    logger.info("Starting tool approval span");
    const spanId = tracing.logToolApprovalSpan(
      requestId,
      toolId,
      toolChange.toolName,
    );

    try {
      logger.info("checking if tool is already approved");
      if (toolChange.approved) {
        throw new Error("Tool is already approved");
      }

      logger.info("checking tool is pending");
      if (!toolChange.pending) {
        throw new Error("Tool is not pending");
      }

      let partialToolChange: any;
      let result: any;

      if (!toolChange.applied) {
        logger.info("tool execute flow");
        result = await toolExecutionFlow(
          requestId,
          toolChange.toolName,
          toolChange.inputData,
          toolChange,
          workbookName,
        );

        logger.info("getting input data revert");

        const inputDataRevert = getInputDataRevert(
          result,
          workbookName,
          toolChange,
        );

        partialToolChange = {
          ...toolChange,
          inputDataRevert: inputDataRevert,
          approved: true,
          applied: true,
          pending: false,
          result: result,
        };
      } else {
        result = toolChange.result;
        partialToolChange = {
          ...toolChange,
          approved: true,
          applied: true,
          pending: false,
        };
      }

      const operations = Array.isArray(result?.operations)
        ? result.operations
        : [];
      const hasOperationError = operations.some((op: any) =>
        op?.status === "error" || op?.policyAction === "blocked"
      );
      const isBlocked = result?.policyAction === "blocked";
      const isFailure = result?.success === false || hasOperationError ||
        isBlocked;
      if (isFailure) {
        const firstErrorCode = result?.errorCode ||
          operations.find((op: any) => op?.errorCode)?.errorCode;
        const errorMessage = result?.message ||
          (isBlocked
            ? "Blocked cross-workbook write"
            : "Tool reported failure");
        const failedOperations = operations.filter((op: any) =>
          op?.status === "error" || op?.policyAction === "blocked"
        ).map((op: any) => ({
          worksheet: op?.worksheet,
          range: op?.range,
          errorCode: op?.errorCode,
          errorMessage: op?.errorMessage,
          policyAction: op?.policyAction,
          status: op?.status,
          requestedWorkbook: op?.requestedWorkbook,
          resolvedWorkbook: op?.resolvedWorkbook,
        }));
        logger.error("Tool reported failure details", {
          toolId,
          toolName: toolChange.toolName,
          errorMessage,
          errorCode: firstErrorCode,
          success: result?.success,
          policyAction: result?.policyAction,
          failedOperationCount: failedOperations.length,
          failedOperations: failedOperations.slice(0, 10),
        });
        if (workbookName) {
          WebSocketManager.getInstance().send(
            workbookName,
            WebSocketMessageTypes.TOOL_ERROR,
            {
              toolId,
              errorMessage,
              errorCode: firstErrorCode,
            },
          );
        }
      }

      logger.info("updating tool change");
      stateManager.updateToolChange(workbookName, toolId, partialToolChange);

      logger.info("Updating Tool Result Placeholder");
      stateManager.updateToolResponseInConversationHistory(
        workbookName,
        result,
        toolChange.toolId,
      );

      if (spanId) {
        logger.info("Ending tool approval span");
        tracing.logToolApprovalSpanEnd(spanId, {
          success: true,
          toolId: toolId,
          toolName: toolChange.toolName,
          approved: true,
          applied: partialToolChange.applied,
        });
      }
    } catch (error: any) {
      logger.error("Error during tool approval", error);
      if (spanId) {
        tracing.logToolApprovalSpanEnd(spanId, {
          success: false,
          error: error.message,
          stack: error.stack,
          toolId: toolId,
          toolName: toolChange.toolName,
        });
      }
      throw error;
    }
  }

  /**
   * Process tool rejection without streaming continuation
   * Extracted from handleToolReject for reuse
   */
  private async processToolRejection(
    workbookName: string,
    toolId: string,
    requestId: string,
    toolChange: ToolChangeInterface,
    userMessage: string,
  ): Promise<void> {
    logger.info("Starting tool rejection span");
    const spanId = tracing.logToolRejectionSpan(
      requestId,
      toolId,
      toolChange.toolName,
      userMessage,
    );

    try {
      if (toolChange.applied) {
        try {
          await revertSingleToolChange(toolChange, false);
          logger.info(
            `Successfully reverted tool ${toolChange.toolName} during rejection`,
          );
          // After revert completes, send idle to ensure status is correct
          // (revert operations send progress updates that are now properly awaited)
          WebSocketManager.getInstance().sendStatus(
            workbookName,
            OperationStatusValues.IDLE,
          );
        } catch (error) {
          logger.error(
            `Failed to revert tool ${toolChange.toolName} during rejection: ${error}`,
          );
          // Don't re-throw - log and continue gracefully
        }
      }

      logger.info("Updating Tool Result Placeholder");
      let response: any =
        `User Message on Tool Call Rejection: ${userMessage}` ||
        "Tool call rejected by the user";
      if (toolChange.toolName === ToolNames.LIST_OPEN_WORKBOOKS) {
        response = {
          errorCode: "PERMISSION_DENIED",
          message: userMessage || "Permission denied by user",
        };
      }
      stateManager.updateToolResponseInConversationHistory(
        workbookName,
        response,
        toolChange.toolId,
      );

      logger.info("updating tool change");
      const partialToolChange = {
        ...toolChange,
        approved: false,
        applied: false,
        pending: false,
      };
      stateManager.updateToolChange(workbookName, toolId, partialToolChange);

      if (spanId) {
        logger.info("Ending tool rejection span");
        tracing.logToolRejectionSpanEnd(spanId, {
          success: true,
          toolId: toolId,
          toolName: toolChange.toolName,
          approved: toolChange.approved,
          applied: toolChange.applied,
        });
      }
    } catch (error: any) {
      logger.error("Error during tool rejection", error);
      if (spanId) {
        tracing.logToolRejectionSpanEnd(spanId, {
          success: false,
          error: error.message,
          stack: error.stack,
          toolId: toolId,
          toolName: toolChange.toolName,
        });
      }
      throw error;
    }
  }
}

function parseDefaultValue(
  preferenceType: string,
  rawValue: string | null,
): { defaultBool: boolean; fallbackValue: string | boolean } {
  if (preferenceType === "string") {
    return { defaultBool: false, fallbackValue: rawValue ?? "" };
  }

  const defaultBool = rawValue
    ? rawValue.toLowerCase() === "true" || rawValue === "1"
    : false;

  return { defaultBool, fallbackValue: defaultBool };
}

function extractAttachmentMetadata(payload: any): Record<string, any> {
  const attachmentMetadata: Record<string, any> = {};

  if (payload.documentsBase64 && payload.documentsBase64.length > 0) {
    logger.info(
      `Extracting metadata for ${payload.documentsBase64.length} attachment(s)`,
    );

    const attachments = payload.documentsBase64.map((doc: any) => ({
      name: doc.name,
      type: doc.type,
      size: doc.size,
    }));

    // Count different file types
    const pdfFiles = attachments.filter((att: any) =>
      att.type === "application/pdf"
    );
    const imageFiles = attachments.filter((att: any) =>
      att.type.startsWith("image/")
    );

    const pdfCount = pdfFiles.length;
    const imageCount = imageFiles.length;

    // Calculate total size
    const totalSize = attachments.reduce(
      (sum: number, att: any) => sum + (att.size || 0),
      0,
    );

    // Calculate PDF-specific metrics
    const totalPdfSize = pdfFiles.reduce(
      (sum: number, att: any) => sum + (att.size || 0),
      0,
    );
    // Note: PDF page count would need to be extracted from the actual PDF content
    // For now, we'll estimate based on size (rough estimate: ~100KB per page)
    const estimatedTotalPages = pdfFiles.reduce((sum: number, att: any) => {
      const estimatedPages = Math.max(1, Math.round((att.size || 0) / 100000)); // 100KB per page estimate
      return sum + estimatedPages;
    }, 0);

    // Calculate image-specific metrics
    const totalImageSize = imageFiles.reduce(
      (sum: number, att: any) => sum + (att.size || 0),
      0,
    );

    attachmentMetadata.hasAttachments = true;
    attachmentMetadata.attachmentCount = attachments.length;
    attachmentMetadata.attachments = attachments;
    attachmentMetadata.pdfCount = pdfCount;
    attachmentMetadata.imageCount = imageCount;
    attachmentMetadata.totalAttachmentSize = totalSize;
    attachmentMetadata.totalPdfSize = totalPdfSize;
    attachmentMetadata.totalImageSize = totalImageSize;
    attachmentMetadata.estimatedTotalPdfPages = estimatedTotalPages;

    // Get file type breakdown
    const fileTypes = [...new Set(attachments.map((att: any) => att.type))];
    attachmentMetadata.fileTypes = fileTypes;

    logger.info("Attachment metadata extracted", attachmentMetadata);
  } else {
    attachmentMetadata.hasAttachments = false;
    attachmentMetadata.attachmentCount = 0;
  }

  return attachmentMetadata;
}

function parseDocumentFromRequest(requestBody: any) {
  const documentEntries: Anthropic.MessageParam[] = [];
  if (requestBody.documentsBase64 && requestBody.documentsBase64.length > 0) {
    logger.info(
      `Processing ${requestBody.documentsBase64.length} document(s) for traceId: ${requestBody.traceId}`,
    );

    for (const [_, document] of requestBody.documentsBase64.entries()) {
      if (document.base64 && document.base64.trim()) {
        const contentType = document.type == "application/pdf"
          ? "document"
          : "image";
        const mediaType = document.type;

        const documentEntry: Anthropic.MessageParam = {
          role: "user",
          content: [{
            type: contentType,
            source: {
              type: "base64",
              media_type: mediaType,
              data: document.base64,
            },
          } as Anthropic.DocumentBlockParam],
        };

        documentEntries.push(documentEntry);
      } else {
        logger.warn(
          `Skipping document ${
            document.name || "unnamed"
          } - missing or empty base64 data`,
        );
      }
    }

    return documentEntries;
  }
}
