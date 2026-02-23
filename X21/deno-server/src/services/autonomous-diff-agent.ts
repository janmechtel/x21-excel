import Anthropic from "@anthropic-ai/sdk";
import {
  createAnthropicClient,
  withAnthropicBetas,
} from "../llm-client/anthropic.ts";
import {
  getAnthropicConfig,
  getLLMProvider,
  type LLMProvider,
} from "../llm-client/provider.ts";
import { createLogger } from "../utils/logger.ts";
import { createDiffSummaryForLLM, DiffResult } from "./workbook-diff.ts";
import {
  getAutonomousDiffAnalysisPrompt,
  getWorkbookDiffUserMessage,
} from "../prompts/workbook-diff.ts";
import { executeTool } from "../stream/tool-logic.ts";
import { executeToolUsesWithConcurrency } from "../llm-client/tool-runner.ts";
import { tools } from "../tools/index.ts";
import { tracing } from "../tracing/tracing.ts";
import { stateManager } from "../state/state-manager.ts";
import { UserService } from "./user.ts";
import {
  ClaudeContentTypes,
  ClaudeStopReasons,
  ContentBlockTypes,
  ToolNames,
} from "../types/index.ts";

const logger = createLogger("AutonomousDiffAgent");

const READ_ONLY_TOOLS = [
  ToolNames.READ_VALUES_BATCH,
  ToolNames.READ_FORMAT_BATCH,
] as const;

/**
 * Get read-only tools for autonomous agent
 */
function getReadOnlyTools() {
  return tools
    .filter((tool) =>
      READ_ONLY_TOOLS.includes(tool.name as typeof READ_ONLY_TOOLS[number])
    )
    .map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: tool.input_schema,
      strict: tool.strict,
    }));
}

/**
 * Create a non-streaming message using the appropriate provider
 */
async function createNonStreamingMessage(
  provider: LLMProvider,
  params: Anthropic.MessageCreateParamsNonStreaming,
): Promise<Anthropic.Message> {
  if (provider === "anthropic") {
    const client = createAnthropicClient();
    return await client.beta.messages.create(
      withAnthropicBetas(params),
    ) as Anthropic.Message;
  } else {
    // Azure OpenAI
    const {
      createNativeOpenAIClient,
      convertAnthropicToResponsesInput,
      convertAnthropicToolsToResponsesFormat,
    } = await import("../llm-client/openai-native.ts");

    const { client, model: azureModel } = createNativeOpenAIClient();

    // Convert messages and tools to Responses API format
    const responsesInput = convertAnthropicToResponsesInput(
      params.messages,
      params.system as string | undefined,
    );
    const responsesTools = convertAnthropicToolsToResponsesFormat(
      params.tools as any,
    );

    // Create non-streaming request
    const createParams = {
      model: azureModel,
      input: responsesInput,
      tools: responsesTools.length > 0 ? responsesTools : undefined,
      tool_choice: responsesTools.length > 0 ? ("auto" as const) : undefined,
      stream: false as const,
    };

    const response = await client.responses.create(createParams) as any;

    logger.info("Azure response received (autonomous agent):", {
      hasOutput: !!response.output,
      outputLength: response.output?.length || 0,
      hasOutputText: !!response.output_text,
      outputText: response.output_text,
    });

    // Convert Azure response to Anthropic format
    const content: Anthropic.ContentBlock[] = [];
    let textContent = "";
    const toolUseBlocks: Anthropic.ToolUseBlock[] = [];

    // First check for output_text at the response level
    if (response.output_text && response.output_text.trim()) {
      textContent = response.output_text;
      logger.info("Found text in output_text:", {
        textLength: textContent.length,
      });
    }

    // Then check output array for function calls, messages, and reasoning
    if (response.output) {
      for (const outputItem of response.output) {
        logger.info("Processing output item:", {
          type: outputItem.type,
          hasName: !!outputItem.name,
          hasArguments: !!outputItem.arguments,
          keys: Object.keys(outputItem),
        });

        if (outputItem.type === "function_call") {
          // Azure format: name and arguments are directly on outputItem
          let parsedInput: any = {};
          try {
            parsedInput = JSON.parse(outputItem.arguments || "{}");
          } catch (e) {
            logger.error("Failed to parse function arguments:", e);
            parsedInput = {};
          }

          toolUseBlocks.push({
            type: ClaudeContentTypes.TOOL_USE,
            id: outputItem.id,
            name: outputItem.name,
            input: parsedInput,
          });

          logger.info("Added tool use block:", {
            id: outputItem.id,
            name: outputItem.name,
            inputKeys: Object.keys(parsedInput),
          });
        } else if (outputItem.type === "message" && outputItem.content) {
          // Extract text from message type output items
          const messageContent = Array.isArray(outputItem.content)
            ? outputItem.content
            : [outputItem.content];

          for (const contentItem of messageContent) {
            if (typeof contentItem === "string") {
              textContent += contentItem;
            } else if (
              contentItem && typeof contentItem === "object" && contentItem.text
            ) {
              // Handle content objects with text property
              textContent += contentItem.text;
            }
          }

          if (textContent) {
            logger.info("Found text in message output item:", {
              textLength: textContent.length,
            });
          }
        }
        // Note: We ignore "reasoning" type items for now as they're internal
      }
    }

    logger.info("Parsed Azure response:", {
      textContentLength: textContent.length,
      textContent: textContent.substring(0, 200),
      toolUseBlocksCount: toolUseBlocks.length,
      hasOutputText: !!response.output_text,
      outputArrayLength: response.output?.length || 0,
    });

    if (textContent) {
      content.push({
        type: ContentBlockTypes.TEXT,
        text: textContent,
      } as Anthropic.TextBlock);
    } else if (toolUseBlocks.length === 0) {
      // No text and no tools - this is unusual, log it
      logger.warn("Azure response has no text content and no tool calls", {
        hasOutputText: !!response.output_text,
        outputArrayLength: response.output?.length || 0,
        outputTypes: response.output?.map((o: any) => o.type) || [],
      });
    }

    content.push(...toolUseBlocks);

    // If we have both text and tools, we should still finish if the model indicates end_turn
    // But Azure might not provide that signal, so we check: if we have text and no tools, it's end_turn
    // If we have tools, it's tool_use (even if we also have text, the tools need to be executed)
    const stopReason = toolUseBlocks.length > 0
      ? ClaudeStopReasons.TOOL_USE
      : "end_turn";

    return {
      id: `azure_${Date.now()}`,
      type: "message",
      role: "assistant",
      model: azureModel,
      content,
      stop_reason: stopReason,
      stop_sequence: null,
      usage: {
        input_tokens: response.usage?.input_tokens || 0,
        output_tokens: response.usage?.output_tokens || 0,
      } as Anthropic.Usage,
    };
  }
}

/**
 * Extract text content from Claude response
 */
function extractTextContent(response: Anthropic.Message): string {
  const textBlocks = response.content.filter(
    (block): block is Anthropic.TextBlock =>
      block.type === ContentBlockTypes.TEXT,
  );

  logger.info("Extracting text content", {
    totalBlocks: response.content.length,
    textBlocksCount: textBlocks.length,
    hasToolUse: response.content.some((b) =>
      b.type === ClaudeContentTypes.TOOL_USE
    ),
  });

  if (textBlocks.length === 0) {
    logger.warn("No text blocks found in response - returning empty summary");
    return "";
  }

  const text = textBlocks.map((block) => block.text).join("\n").trim();

  if (!text) {
    logger.warn("Text blocks exist but are empty - returning empty summary");
    return "";
  }

  // Clean up: extract only checkbox lines if text contains non-checkbox content
  const lines = text.split("\n");
  const checkboxLines = lines.filter((line) => {
    const trimmed = line.trim();
    // Match checkbox pattern and ensure there's actual content after the checkbox
    const match = trimmed.match(/^-\s*\[\s*[xX ]?\s*\]\s*(.+)$/);
    return match && match[1].trim().length > 0;
  });

  // If we found checkboxes with content, return only those (agent included extra text)
  if (checkboxLines.length > 0) {
    const result = checkboxLines.join("\n");
    logger.info("Extracted checkbox lines", {
      count: checkboxLines.length,
      preview: result.substring(0, 100),
    });
    return result;
  }

  // Otherwise return the full text (might be an error message or edge case)
  logger.info("Returning full text (no checkboxes found)", {
    length: text.length,
    preview: text.substring(0, 100),
  });
  return text;
}

/**
 * Run autonomous analysis with tool-enhanced exploration
 */
export async function runAutonomousAnalysis(
  workbookName: string,
  diffs: DiffResult[],
  traceContext: {
    diffId?: string | null;
    comparisonType?: "self" | "external";
    comparisonFilePath?: string | null;
    comparisonFileModifiedAt?: number | null;
  } = {},
): Promise<string> {
  logger.info("Starting autonomous analysis", { workbookName });

  const provider = getLLMProvider();
  logger.info("Using LLM provider", { provider });

  const diffSummary = createDiffSummaryForLLM(diffs);
  const userMessage = getWorkbookDiffUserMessage(workbookName, diffSummary);

  // Pull best-available metadata from state/user services for tracing
  const userEmail = UserService.getInstance().getUserEmail();
  const state = stateManager.getOrCreateState(workbookName);
  const parentRequestId = state.latestRequestId ?? null;

  const sheetsWithChanges = diffs
    .filter((d) => d?.hasChanges === true && d?.isActualSheet === true)
    .map((d) => d.sheetName);
  const internalFilesChanged = diffs
    .filter((d) => d?.hasChanges === true && d?.isActualSheet === false)
    .map((d) => d.sheetPath);

  // Create a unique trace ID for this autonomous analysis
  const safeWorkbookName = workbookName.replace(/[^\w.-]+/g, "_").slice(0, 80);
  const traceId = `autonomous-${safeWorkbookName}-${Date.now()}`;

  // Get session ID from state manager if workbook exists
  let sessionId: string | undefined;
  try {
    sessionId = stateManager.getSessionId(workbookName);
  } catch {
    // Workbook may not have a session yet, that's ok
    sessionId = undefined;
  }

  const traceMetadata = {
    name: "Autonomous Diff Analysis",
    workbookName,
    sessionId,
    parentRequestId,
    userEmail,
    diffsCount: diffs.length,
    comparisonType: traceContext.comparisonType ?? null,
    comparisonFileName: traceContext.comparisonFilePath?.split(/[\\/]/).pop() ??
      null,
    comparisonFilePath: traceContext.comparisonFilePath ?? null,
    comparisonFileModifiedAt: traceContext.comparisonFileModifiedAt ?? null,
    sheetsWithChangesCount: sheetsWithChanges.length,
    internalFilesChangedCount: internalFilesChanged.length,
    provider,
    startTimestamp: new Date().toISOString(),
  };

  // Start Langfuse trace for autonomous agent
  tracing.startTrace(
    traceId,
    traceMetadata,
    workbookName,
  );
  // Keep the full prompt visible as the trace "input" without storing it in metadata.
  try {
    tracing.getTrace(traceId)?.update({ input: userMessage } as any);
  } catch {
    // Best-effort only
  }

  const messages: Anthropic.MessageParam[] = [
    {
      role: "user",
      // Use content blocks so centralized tracing (content[0]) doesn't log only the first character
      content: [{ type: ContentBlockTypes.TEXT, text: userMessage }] as any,
    },
  ];

  let iterations = 0;
  const MAX_ITERATIONS = 15;
  const startTime = Date.now();
  const TIMEOUT = 60000; // 60 seconds

  while (iterations < MAX_ITERATIONS) {
    // Check timeout
    if (Date.now() - startTime > TIMEOUT) {
      logger.warn("Autonomous agent timeout - returning partial results");

      // End trace with timeout
      tracing.endTrace(traceId, {
        ...traceMetadata,
        success: false,
        timeout: true,
        iterations: iterations + 1,
        endTimestamp: new Date().toISOString(),
      });

      return "Analysis timed out.";
    }

    logger.info(`Autonomous agent iteration ${iterations + 1}`);

    try {
      // Get model based on provider
      let model: string;
      if (provider === "anthropic") {
        const config = getAnthropicConfig();
        if (!config) {
          throw new Error("Anthropic configuration not found");
        }
        model = config.model;
      } else {
        // Azure OpenAI
        const { getAzureOpenAIConfig } = await import(
          "../llm-client/provider.ts"
        );
        const config = getAzureOpenAIConfig();
        if (!config) {
          throw new Error("Azure OpenAI configuration not found");
        }
        model = config.model;
      }

      // Log generation start
      const generationParams: Anthropic.MessageCreateParamsNonStreaming = {
        model,
        max_tokens: 4000,
        system: getAutonomousDiffAnalysisPrompt(),
        messages,
        tools: getReadOnlyTools() as Anthropic.MessageCreateParamsNonStreaming[
          "tools"
        ],
      };

      const generationId = tracing.logGenerationStart(
        traceId,
        generationParams,
        {
          iteration: iterations + 1,
          provider,
          autonomousAgent: true,
          workbookName,
          userEmail,
        },
      );

      // Create non-streaming message using the appropriate provider
      const response = await createNonStreamingMessage(
        provider,
        generationParams,
      );

      // Log generation end with response
      tracing.logGenerationEnd(generationId, {
        output: response.content,
        metadata: {
          inputTokens: response.usage?.input_tokens || 0,
          outputTokens: response.usage?.output_tokens || 0,
          stop_reason: response.stop_reason,
        },
      });

      // Add assistant response to history
      messages.push({
        role: "assistant",
        content: response.content,
      });

      // If agent finished, extract and return summary
      if (response.stop_reason === "end_turn") {
        const summary = extractTextContent(response);

        if (!summary || !summary.trim()) {
          logger.warn("Autonomous agent completed but summary is empty", {
            contentBlocks: response.content.length,
            contentTypes: response.content.map((b) => b.type),
            stopReason: response.stop_reason,
          });

          // End trace with warning
          tracing.endTrace(traceId, {
            ...traceMetadata,
            success: false,
            iterations: iterations + 1,
            emptySummary: true,
            endTimestamp: new Date().toISOString(),
          });

          return "Changes detected in workbook. Summary generation completed but no content was produced.";
        }

        logger.info("Autonomous agent completed successfully", {
          summaryLength: summary.length,
          summaryPreview: summary.substring(0, 150),
        });

        // End trace successfully
        tracing.endTrace(traceId, {
          ...traceMetadata,
          success: true,
          iterations: iterations + 1,
          summaryLength: summary.length,
          endTimestamp: new Date().toISOString(),
        });

        return summary;
      }

      // If agent wants to use tools, execute them
      if (response.stop_reason === ClaudeStopReasons.TOOL_USE) {
        const toolResults = await executeToolsAutonomously(
          traceId,
          workbookName,
          response.content,
        );

        messages.push({
          role: "user",
          content: toolResults,
        });
      }

      iterations++;
    } catch (error) {
      logger.error("Error in autonomous agent loop", error);

      // End trace with error
      tracing.endTrace(traceId, {
        ...traceMetadata,
        success: false,
        error: error instanceof Error ? error.message : "Unknown error",
        iterations: iterations + 1,
        endTimestamp: new Date().toISOString(),
      });

      return "Analysis encountered an error.";
    }
  }

  // Max iterations reached
  logger.warn("Max iterations reached without completion");

  // End trace with max iterations
  tracing.endTrace(traceId, {
    ...traceMetadata,
    success: false,
    maxIterations: true,
    iterations,
    endTimestamp: new Date().toISOString(),
  });

  return "Analysis incomplete.";
}

/**
 * Execute tools autonomously without user approval
 */
async function executeToolsAutonomously(
  traceId: string,
  _workbookName: string,
  content: Anthropic.ContentBlock[],
): Promise<Anthropic.ToolResultBlockParam[]> {
  const toolBlocks = content.filter(
    (block): block is Anthropic.ToolUseBlock =>
      block.type === ClaudeContentTypes.TOOL_USE,
  );

  const executor = async (tool: Anthropic.ToolUseBlock) => {
    const spanId = tracing.startSpan(traceId, {
      name: `Tool: ${tool.name}`,
      input: tool.input,
      metadata: {
        autonomousTool: true,
        toolId: tool.id,
      },
    });

    try {
      logger.info(`Autonomous agent calling tool: ${tool.name}`, tool.input);
      const result = await executeTool(tool.name, tool.input);
      tracing.endSpan(spanId, { output: result, success: true });
      return result;
    } catch (error: any) {
      logger.warn(`Tool ${tool.name} failed in autonomous mode`, error);
      const wrappedError = error instanceof Error
        ? error
        : new Error(error?.message || "Tool execution failed");
      (wrappedError as any).note = (wrappedError as any)?.note ||
        "Workbook may be closed or unavailable";

      tracing.endSpan(spanId, {
        output: {
          error: wrappedError.message || "Tool execution failed",
          note: (wrappedError as any)?.note,
        },
        success: false,
        error: wrappedError.message,
      });

      throw wrappedError;
    }
  };

  return await executeToolUsesWithConcurrency(toolBlocks, {
    executor,
    formatResultContent: (result) =>
      typeof result === "string" ? result : JSON.stringify(result, null, 2),
    loggerContext: { traceId, workbookName: _workbookName },
  });
}
