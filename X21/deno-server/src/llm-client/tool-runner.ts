import Anthropic from "@anthropic-ai/sdk";
import { ClaudeContentTypes, ExcelInteropToolNames } from "../types/index.ts";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("ToolRunner");

const DEFAULT_PARALLEL_LIMIT = 5;

const excelToolNames = new Set<string>(ExcelInteropToolNames);

type ExecutionMode = "excel-serial" | "parallel";

type ToolUseEntry = {
  index: number;
  block: Anthropic.ToolUseBlock;
};

export interface ExecuteToolUsesOptions {
  executor: (tool: Anthropic.ToolUseBlock) => Promise<any>;
  parallelLimit?: number;
  formatResultContent?: (result: any) => any;
  loggerContext?: Record<string, unknown>;
}

const defaultFormatter = (result: any) => {
  if (typeof result === "string") return result;
  if (result === undefined) return "";
  return JSON.stringify(result);
};

const isExcelTool = (name: string): boolean => excelToolNames.has(name);

const createLimiter = (limit: number) => {
  let active = 0;
  const queue: Array<() => void> = [];

  const next = () => {
    if (queue.length === 0 || active >= limit) return;
    const fn = queue.shift();
    if (fn) fn();
  };

  return async <T>(fn: () => Promise<T>): Promise<T> =>
    await new Promise<T>((resolve, reject) => {
      const run = () => {
        active++;
        fn()
          .then(resolve)
          .catch(reject)
          .finally(() => {
            active--;
            next();
          });
      };

      if (active < limit) {
        run();
      } else {
        queue.push(run);
      }
    });
};

async function runTool(
  entry: ToolUseEntry,
  executor: (tool: Anthropic.ToolUseBlock) => Promise<any>,
  formatResultContent: (result: any) => any,
  mode: ExecutionMode,
  loggerContext?: Record<string, unknown>,
): Promise<Anthropic.ToolResultBlockParam> {
  const startedAt = Date.now();
  try {
    const output = await executor(entry.block);
    return {
      type: ClaudeContentTypes.TOOL_RESULT,
      tool_use_id: entry.block.id,
      content: formatResultContent(output),
    };
  } catch (error: any) {
    const errorPayload: Record<string, unknown> = {
      message: error?.message || "Tool execution failed",
    };
    if (error?.stack) errorPayload.stack = error.stack;
    if (error?.code) errorPayload.code = error.code;
    if (error?.note) errorPayload.note = error.note;

    return {
      type: ClaudeContentTypes.TOOL_RESULT,
      tool_use_id: entry.block.id,
      is_error: true,
      content: formatResultContent(errorPayload),
    };
  } finally {
    const durationMs = Date.now() - startedAt;
    logger.info("Tool execution finished", {
      mode,
      toolName: entry.block.name,
      toolUseId: entry.block.id,
      durationMs,
      ...(loggerContext || {}),
    });
  }
}

export async function executeToolUsesWithConcurrency(
  toolUses: Anthropic.ToolUseBlock[],
  options: ExecuteToolUsesOptions,
): Promise<Anthropic.ToolResultBlockParam[]> {
  const {
    executor,
    parallelLimit = DEFAULT_PARALLEL_LIMIT,
    formatResultContent = defaultFormatter,
    loggerContext,
  } = options;

  logger.info("Detected tool_use blocks", {
    count: toolUses.length,
    toolNames: toolUses.map((t) => t.name),
    ...(loggerContext || {}),
  });

  const results: Anthropic.ToolResultBlockParam[] = new Array(toolUses.length);
  const limiter = createLimiter(Math.max(1, parallelLimit));
  let inFlight: Promise<void>[] = [];

  for (const [index, block] of toolUses.entries()) {
    const entry: ToolUseEntry = { index, block };

    if (!isExcelTool(block.name)) {
      const p = limiter(async () => {
        results[entry.index] = await runTool(
          entry,
          executor,
          formatResultContent,
          "parallel",
          loggerContext,
        );
      });
      inFlight.push(p);
      continue;
    }

    // For Excel tools, wait for all previously launched non-Excel tools to finish to preserve ordering.
    if (inFlight.length > 0) {
      await Promise.all(inFlight);
      inFlight = [];
    }

    results[entry.index] = await runTool(
      entry,
      executor,
      formatResultContent,
      "excel-serial",
      loggerContext,
    );
  }

  // Flush any remaining non-Excel tasks
  if (inFlight.length > 0) {
    await Promise.all(inFlight);
  }

  logger.info("tool_result bundle", {
    tool_result_count: results.length,
    has_errors: results.some((r) => r?.is_error),
    ...(loggerContext || {}),
  });

  return results;
}
