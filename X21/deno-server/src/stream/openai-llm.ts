import Anthropic from "@anthropic-ai/sdk";
import { createLogger } from "../utils/logger.ts";
import { tracing } from "../tracing/tracing.ts";
import { WebSocketManager } from "../services/websocket-manager.ts";
import { stateManager } from "../state/state-manager.ts";
import { UserCancellationError } from "../errors/user-cancellation-error.ts";
import { getAzureOpenAIConfig } from "../llm-client/provider.ts";
import {
  ClaudeContentTypes,
  ClaudeEventTypes,
  ClaudeStopReasons,
  ContentBlockTypes,
  OperationStatusValues,
} from "../types/index.ts";

const logger = createLogger("streamNativeOpenAIResponseToWebSocket");
const socket = WebSocketManager.getInstance();

type ReasoningEffort = "low" | "medium" | "high";
type ReasoningSummary = "auto" | "none";

function readReasoningEffort(): ReasoningEffort {
  const config = getAzureOpenAIConfig();
  const raw = (config?.reasoningEffort || "medium").toLowerCase();
  if (raw === "low" || raw === "medium" || raw === "high") {
    return raw;
  }
  return "medium";
}

function readReasoningSummary(): ReasoningSummary {
  // Fixed value; no longer configurable via environment variables.
  return "auto";
}

function readMaxOutputTokens(): number | undefined {
  // Fixed value; no longer configurable via environment variables.
  return 25000;
}

function shouldLogVerbose(): boolean {
  return (Deno.env.get("OPENAI_VERBOSE_LOGS") || "").toLowerCase() === "true";
}

/**
 * Stream OpenAI Responses API to WebSocket in Anthropic-compatible format
 */
export async function streamNativeOpenAIResponseToWebSocket(
  requestId: string,
  params: Anthropic.MessageCreateParamsNonStreaming,
  abortController: AbortController,
): Promise<Anthropic.Message> {
  const {
    createNativeOpenAIClient,
    convertAnthropicToResponsesInput,
    convertAnthropicToolsToResponsesFormat,
  } = await import("../llm-client/openai-native.ts");

  const workbookName: string = stateManager.getWorkbookName(requestId);

  logger.info("🚀 Starting Azure OpenAI streaming request", {
    requestId,
    workbookName,
    messageCount: params.messages.length,
    hasSystemPrompt: !!params.system,
    toolCount: params.tools?.length || 0,
    maxTokens: params.max_tokens,
  });

  // Create Azure OpenAI Responses API client - returns both client and the actual model to use
  const { client, model: azureModel } = createNativeOpenAIClient();

  logger.info("✓ Azure OpenAI client ready for request", {
    requestId,
    modelToUse: azureModel,
  });

  // Convert messages and tools directly to Responses API format
  const responsesInput = convertAnthropicToResponsesInput(
    params.messages,
    params.system as string | undefined,
  );
  const responsesTools = convertAnthropicToolsToResponsesFormat(
    params.tools as any,
  );
  const reasoningEffort = readReasoningEffort();
  const reasoningSummary = readReasoningSummary();
  const maxOutputTokens = readMaxOutputTokens();
  const verboseLogs = shouldLogVerbose();

  // Extract Excel context from the last user message for Langfuse logging
  let excelContext: any = {};
  try {
    const lastUserMessage = params.messages[params.messages.length - 1];
    if (lastUserMessage && typeof lastUserMessage.content === "string") {
      const parsed = JSON.parse(lastUserMessage.content);
      excelContext = parsed.excelContext || {};
    }
  } catch {
    // If parsing fails, context will be empty
  }

  const generationId = tracing.logGenerationStart(requestId, params, {
    provider: "azure-openai",
    model: azureModel,
    systemMessage: (params.system as string) || "",
    toolCount: responsesTools.length,
    toolNames: responsesTools.map((t: any) => t.name).join(", "),
    toolNamesList: responsesTools.map((t: any) => t.name),
    // Excel context details
    workbookName: excelContext.workbookName || workbookName,
    activeSheet: excelContext.activeSheet,
    selectedRange: excelContext.selectedRange,
    usedRange: excelContext.usedRange,
  });

  try {
    logger.info("NATIVE: Responses API input being sent to OpenAI:", {
      inputCount: responsesInput.length,
      input: responsesInput.map((msg, idx) => ({
        index: idx,
        role: msg.role,
        contentLength: msg.content?.[0]?.text?.length || 0,
        textPreview: msg.content?.[0]?.text?.substring(0, 200) || "",
      })),
    });

    logger.info("NATIVE: Responses API tools being sent to OpenAI:", {
      toolCount: responsesTools.length,
      tools: responsesTools.map((tool) => ({
        type: tool.type,
        name: tool.name,
        hasDescription: !!tool.description,
        descriptionPreview: tool.description?.substring(0, 100),
        parameterCount: tool.input_schema?.properties
          ? Object.keys(tool.input_schema.properties).length
          : 0,
        requiredParams: tool.input_schema?.required || [],
      })),
    });

    // Detect if this is a reasoning model (o1, o3, gpt-5) that supports reasoning_effort
    // Use the actual Azure model name, not the params.model which might be a Claude model
    const modelStr = azureModel.toLowerCase();
    const isReasoningModel = modelStr.includes("o1") ||
      modelStr.includes("o3") ||
      modelStr.includes("gpt-5") ||
      modelStr.includes("gpt5");

    if (isReasoningModel) {
      logger.info("NATIVE: Using reasoning model:", {
        model: azureModel,
        reasoning_effort: reasoningEffort,
        reasoning_summary: reasoningSummary,
        max_output_tokens: maxOutputTokens,
      });
    }

    // Build base parameters for Responses API
    const createParams = {
      model: azureModel,
      input: responsesInput,
      tools: responsesTools.length > 0 ? responsesTools : undefined,
      tool_choice: responsesTools.length > 0 ? ("auto" as const) : undefined,
      stream: true as const,
      reasoning: isReasoningModel
        ? {
          effort: reasoningEffort,
          summary: reasoningSummary,
        }
        : undefined,
      max_output_tokens: maxOutputTokens,
    };

    logger.info("📡 Calling Azure OpenAI Responses API with params:", {
      requestId,
      model: azureModel,
      inputCount: responsesInput.length,
      toolCount: responsesTools.length,
      toolChoice: createParams.tool_choice,
      isReasoningModel,
      reasoningEffort: isReasoningModel ? reasoningEffort : "N/A",
      reasoningSummary: isReasoningModel ? reasoningSummary : "N/A",
      maxOutputTokens,
      hasTools: !!createParams.tools,
      hasAbortSignal: !!abortController?.signal,
    });

    if (verboseLogs) {
      logger.info("NATIVE: Full API call params:", {
        fullParams: JSON.stringify(
          {
            ...createParams,
            input: createParams.input.map((msg: any, idx: number) => ({
              index: idx,
              role: msg.role,
              textPreview: msg.content?.[0]?.text?.substring(0, 150),
            })),
          },
          null,
          2,
        ),
      });
    }

    // Send connecting status
    socket.sendStatus(
      workbookName,
      OperationStatusValues.CONNECTING,
      "Connecting to Azure OpenAI...",
    );

    // Log the actual API call details
    logger.info("🔌 Attempting connection to Azure OpenAI:", {
      requestId,
      endpoint: "[will use client baseURL]",
      model: azureModel,
      streamEnabled: true,
    });

    let stream;
    try {
      logger.info("🔧 Debug: About to call client.responses.stream", {
        requestId,
        paramsKeys: Object.keys(createParams),
        clientType: client.constructor?.name,
      });
      stream = client.responses.stream(createParams);
    } catch (error: any) {
      logger.error("❌ Failed to initiate stream", {
        requestId,
        errorMessage: error.message,
        errorType: error.constructor?.name,
        errorCode: error.code,
        statusCode: error.status,
        errorDetails: error.response?.data || error.response ||
          "No response data",
        possibleCauses: [
          "Endpoint doesn't support Responses API (/openai/v1/responses)",
          "Invalid endpoint URL format",
          "Network connectivity issue to internal endpoint",
          "Invalid API key or authentication method",
          "Firewall blocking internal network access",
          "API version mismatch",
        ],
      });
      throw error;
    }

    logger.info(
      "NATIVE: OpenAI API call successful, starting to process stream",
    );

    // Send generating status when streaming starts
    socket.sendStatus(
      workbookName,
      OperationStatusValues.GENERATING_LLM,
      "Generating response...",
    );

    // Send message_start event according to Anthropic spec
    socket.send(workbookName, "stream:delta", {
      type: ClaudeEventTypes.MESSAGE_START,
      message: {
        id: `native_${Date.now()}`,
        type: "message",
        role: "assistant",
        content: [],
        model: azureModel,
        stop_reason: null,
        stop_sequence: null,
        usage: {
          input_tokens: 0,
          output_tokens: 0,
        },
      },
    });

    let textContent = "";
    let reasoningContent = "";
    let inputTokens = 0;
    let outputTokens = 0;

    const toolCallBuffer = new Map<
      string,
      { name: string; arguments: string }
    >();
    const toolIndexMap = new Map<string, number>();

    let nextIndex = 0;
    let textBlockStarted = false;
    let textBlockIndex = -1;
    let reasoningBlockStarted = false;
    let reasoningBlockIndex = -1;

    logger.info("NATIVE: Initialized toolCallBuffer for streaming:", {
      requestId,
      workbookName,
    });

    for await (const chunk of stream) {
      if (abortController.signal.aborted) {
        throw new UserCancellationError("Stream aborted by user", requestId);
      }

      const event = chunk as any;

      // Handle reasoning output item added - this is our trigger to show the thinking UI
      if (
        event.type === "response.output_item.added" &&
        event.item?.type === "reasoning"
      ) {
        if (!reasoningBlockStarted) {
          reasoningBlockStarted = true;
          reasoningBlockIndex = nextIndex++;
          logger.info("NATIVE: Starting reasoning block:", {
            index: reasoningBlockIndex,
          });
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_START,
            index: reasoningBlockIndex,
            content_block: { type: ContentBlockTypes.THINKING, thinking: "" },
          });
        }
        continue;
      }

      // Also handle possible reasoning deltas (if/when OpenAI exposes them)
      if (
        (event.type === "response.reasoning_summary_text.delta" ||
          event.type === "response.reasoning.delta") &&
        event.delta
      ) {
        const reasoningDelta = event.delta;
        logger.info("NATIVE: Reasoning summary delta received:", {
          reasoningLength: reasoningDelta.length,
          reasoningBlockStarted,
        });

        if (!reasoningBlockStarted) {
          reasoningBlockStarted = true;
          reasoningBlockIndex = nextIndex++;
          logger.info("NATIVE: Starting reasoning block (from delta):", {
            index: reasoningBlockIndex,
          });
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_START,
            index: reasoningBlockIndex,
            content_block: { type: ContentBlockTypes.THINKING, thinking: "" },
          });
        }

        reasoningContent += reasoningDelta;
        socket.send(workbookName, "stream:delta", {
          type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
          index: reasoningBlockIndex,
          delta: { type: "thinking_delta", thinking: reasoningDelta },
        });
        continue;
      }

      if (event.type === "response.reasoning_text.delta" && event.delta) {
        const reasoningDelta = event.delta;
        logger.info("NATIVE: Reasoning text delta received:", {
          reasoningLength: reasoningDelta.length,
          reasoningBlockStarted,
        });

        if (!reasoningBlockStarted) {
          reasoningBlockStarted = true;
          reasoningBlockIndex = nextIndex++;
          logger.info("NATIVE: Starting reasoning block (from text delta):", {
            index: reasoningBlockIndex,
          });
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_START,
            index: reasoningBlockIndex,
            content_block: { type: ContentBlockTypes.THINKING, thinking: "" },
          });
        }

        reasoningContent += reasoningDelta;
        socket.send(workbookName, "stream:delta", {
          type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
          index: reasoningBlockIndex,
          delta: { type: "thinking_delta", thinking: reasoningDelta },
        });
        continue;
      }

      // Handle function_call output item added (capture tool name)
      if (
        event.type === "response.output_item.added" &&
        event.item?.type === "function_call"
      ) {
        const toolId = event.item.id || event.item_id;
        const toolName = event.item.name ||
          event.item.function_call?.name ||
          event.item.function?.name ||
          event.name;

        logger.info("NATIVE: Function call output item added - FULL EVENT:", {
          eventType: event.type,
          eventKeys: Object.keys(event),
          itemKeys: event.item ? Object.keys(event.item) : [],
          toolId,
          toolName,
          eventName: event.name,
          itemName: event.item?.name,
          functionCallName: event.item?.function_call?.name,
          functionName: event.item?.function?.name,
          fullEvent: JSON.stringify(event),
        });

        if (toolId && toolName) {
          logger.info("NATIVE: Capturing tool name in buffer:", {
            toolId,
            toolName,
          });
          const existing = toolCallBuffer.get(toolId) || {
            name: toolName,
            arguments: "",
          };
          existing.name = toolName;
          toolCallBuffer.set(toolId, existing);

          logger.info("NATIVE: Buffer state after capturing tool name:", {
            toolId,
            bufferSize: toolCallBuffer.size,
            hasEntry: toolCallBuffer.has(toolId),
            entryName: toolCallBuffer.get(toolId)?.name,
          });
        } else {
          logger.warn(
            "NATIVE: Could not extract tool name from function_call event!",
            {
              toolId,
              toolName,
              hasToolId: !!toolId,
              hasToolName: !!toolName,
            },
          );
        }
        continue;
      }

      // Handle text output delta from Responses API
      if (event.type === "response.output_text.delta" && event.delta) {
        const textDelta = event.delta;

        logger.info("NATIVE: Text delta received:", {
          content: textDelta,
          contentLength: textDelta.length,
          textBlockStarted,
        });

        if (!textBlockStarted) {
          textBlockStarted = true;
          textBlockIndex = nextIndex++;
          logger.info("NATIVE: Starting text block:", {
            index: textBlockIndex,
          });
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_START,
            index: textBlockIndex,
            content_block: { type: ContentBlockTypes.TEXT, text: "" },
          });
        }

        textContent += textDelta;
        socket.send(workbookName, "stream:delta", {
          type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
          index: textBlockIndex,
          delta: { type: "text_delta", text: textDelta },
        });
        continue;
      }

      // Handle function call arguments delta (for tool calls)
      if (
        event.type === "response.function_call_arguments.delta" &&
        event.delta
      ) {
        const toolId = event.item_id;
        const argsFragment = event.delta;

        let bufferEntry = toolCallBuffer.get(toolId);
        const hadExisting = !!bufferEntry;
        if (bufferEntry) {
          bufferEntry.arguments += argsFragment;
        } else {
          bufferEntry = {
            name: "",
            arguments: argsFragment,
          };
        }
        toolCallBuffer.set(toolId, bufferEntry);

        logger.info("NATIVE: Function call arguments delta received:", {
          toolId,
          argsFragmentLength: argsFragment.length,
          hadExistingEntry: hadExisting,
          bufferEntryName: bufferEntry.name,
          bufferEntryArgsLength: bufferEntry.arguments.length,
          bufferSize: toolCallBuffer.size,
        });

        if (!toolIndexMap.has(toolId)) {
          const idx = nextIndex++;
          toolIndexMap.set(toolId, idx);
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_START,
            index: idx,
            content_block: {
              type: ClaudeContentTypes.TOOL_USE,
              id: toolId,
              name: bufferEntry.name || "function",
              input: "",
            },
          });
        }

        const idx = toolIndexMap.get(toolId) as number;
        socket.send(workbookName, "stream:delta", {
          type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
          index: idx,
          delta: {
            type: "input_json_delta",
            partial_json: argsFragment,
          },
        });
        continue;
      }

      // Handle function call arguments done (get function name)
      if (event.type === "response.function_call_arguments.done") {
        const toolId = event.item_id;
        const toolName = event.name;
        const fullArgs = event.arguments;

        logger.info("NATIVE: Function call arguments done:", {
          toolId,
          toolName,
          argsLength: fullArgs?.length,
        });

        const existing = toolCallBuffer.get(toolId) || {
          name: toolName || "",
          arguments: fullArgs || "",
        };

        if (toolName) existing.name = toolName;
        if (fullArgs && !existing.arguments) {
          existing.arguments = fullArgs;
        }
        toolCallBuffer.set(toolId, existing);

        logger.info("NATIVE: Buffer state after arguments done:", {
          toolId,
          bufferSize: toolCallBuffer.size,
          entryName: toolCallBuffer.get(toolId)?.name,
          entryArgsLength: toolCallBuffer.get(toolId)?.arguments.length,
        });
        continue;
      }

      // Handle response.completed event (extract tool names and token usage)
      if (event.type === "response.completed" && event.response) {
        let completedOutputText = "";

        // Capture token usage if available
        if (event.response.usage) {
          inputTokens = event.response.usage.input_tokens || 0;
          outputTokens = event.response.usage.output_tokens || 0;
          logger.info("NATIVE: Captured token usage from response:", {
            inputTokens,
            outputTokens,
            totalTokens: inputTokens + outputTokens,
          });
        }

        // Extract tool names and any completed output text
        if (event.response.output) {
          logger.info(
            "NATIVE: Response completed, extracting tool names from output:",
            {
              outputItems: event.response.output.length,
            },
          );

          for (const outputItem of event.response.output) {
            if (
              outputItem.type === "message" && Array.isArray(outputItem.content)
            ) {
              for (const contentItem of outputItem.content) {
                if (contentItem?.type === "output_text" && contentItem.text) {
                  completedOutputText += contentItem.text;
                }
              }
            } else if (outputItem.type === "output_text" && outputItem.text) {
              completedOutputText += outputItem.text;
            }

            if (
              outputItem.type === "function_call" && outputItem.function_call
            ) {
              const toolId = outputItem.id;
              const toolName = outputItem.function_call.name;

              logger.info(
                "NATIVE: Extracting tool name from completed response:",
                {
                  toolId,
                  toolName,
                },
              );

              if (toolId && toolName) {
                const existing = toolCallBuffer.get(toolId);
                if (existing) {
                  existing.name = toolName;
                  toolCallBuffer.set(toolId, existing);
                  logger.info("NATIVE: Updated tool name in buffer:", {
                    toolId,
                    toolName,
                  });
                }
              }
            }
          }
        }

        if (!textContent && completedOutputText) {
          if (!textBlockStarted) {
            textBlockStarted = true;
            textBlockIndex = nextIndex++;
            logger.info(
              "NATIVE: Starting text block (from completed output):",
              {
                index: textBlockIndex,
              },
            );
            socket.send(workbookName, "stream:delta", {
              type: ClaudeEventTypes.CONTENT_BLOCK_START,
              index: textBlockIndex,
              content_block: { type: ContentBlockTypes.TEXT, text: "" },
            });
          }

          textContent += completedOutputText;
          socket.send(workbookName, "stream:delta", {
            type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
            index: textBlockIndex,
            delta: { type: "text_delta", text: completedOutputText },
          });
        }
        continue;
      }

      // Other event types are ignored but logged above
    }

    // Close reasoning block if started (even if no reasoningContent)
    if (reasoningBlockStarted && reasoningBlockIndex >= 0) {
      socket.send(workbookName, "stream:delta", {
        type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
        index: reasoningBlockIndex,
      });
    }

    // Close text block if started
    if (textBlockStarted && textBlockIndex >= 0) {
      socket.send(workbookName, "stream:delta", {
        type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
        index: textBlockIndex,
      });
    }

    // Close each tool block
    for (const [_toolId, toolIndex] of toolIndexMap.entries()) {
      socket.send(workbookName, "stream:delta", {
        type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
        index: toolIndex,
      });
    }

    logger.info("NATIVE: Stream processing complete:", {
      totalReasoningLength: reasoningContent.length,
      totalTextLength: textContent.length,
      toolCallCount: toolCallBuffer.size,
      reasoningBlockStarted,
      textBlockStarted,
    });

    logger.info("NATIVE: Tool call buffer before creating tool use blocks:", {
      bufferSize: toolCallBuffer.size,
      bufferEntries: Array.from(toolCallBuffer.entries()).map(
        ([id, call]) => ({
          id,
          name: call.name,
          argumentsLength: call.arguments.length,
          argumentsPreview: call.arguments.substring(0, 100),
        }),
      ),
    });

    const toolUseBlocks = Array.from(toolCallBuffer.entries()).map(
      ([id, call]) => {
        let parsedInput: any = {};
        if (typeof call.arguments === "string" && call.arguments.length > 0) {
          try {
            parsedInput = JSON.parse(call.arguments);
          } catch (parseError) {
            logger.warn("NATIVE: Failed to parse tool arguments as JSON:", {
              id,
              arguments: call.arguments,
              error: parseError,
            });
            parsedInput = call.arguments;
          }
        } else {
          logger.warn("NATIVE: Tool arguments are empty:", {
            id,
            argumentsType: typeof call.arguments,
            arguments: call.arguments,
          });
        }

        logger.info("NATIVE: Created tool use block:", {
          id,
          name: call.name,
          inputEmpty: Object.keys(parsedInput).length === 0,
          input: parsedInput,
        });

        return {
          type: ClaudeContentTypes.TOOL_USE,
          id,
          name: call.name,
          input: parsedInput,
        } as Anthropic.ToolUseBlock;
      },
    );

    const stopReason = toolUseBlocks.length > 0
      ? ClaudeStopReasons.TOOL_USE
      : "end_turn";

    socket.send(workbookName, "stream:delta", {
      type: ClaudeEventTypes.MESSAGE_DELTA,
      delta: {
        stop_reason: stopReason,
        stop_sequence: null,
      },
      usage: {
        output_tokens: outputTokens,
      },
    });

    socket.send(workbookName, "stream:delta", {
      type: ClaudeEventTypes.MESSAGE_STOP,
    });

    const contentBlocks: Anthropic.ContentBlock[] = [];

    // ✅ Always add thinking block if reasoning UI was started,
    // even if reasoningContent is an empty string.
    if (reasoningBlockStarted) {
      contentBlocks.push({
        type: ContentBlockTypes.THINKING,
        thinking: reasoningContent || "",
      } as any);
    }

    if (textContent) {
      contentBlocks.push({
        type: ContentBlockTypes.TEXT,
        text: textContent,
        citations: [],
      } as Anthropic.TextBlock);
    }

    contentBlocks.push(...toolUseBlocks);

    const responseMessage: Anthropic.Message = {
      id: `native_${Date.now()}`,
      type: "message",
      role: "assistant",
      model: azureModel,
      content: contentBlocks,
      stop_reason: stopReason,
      stop_sequence: null,
      usage: {
        input_tokens: inputTokens,
        output_tokens: outputTokens,
        cache_creation_input_tokens: 0,
        cache_read_input_tokens: 0,
      } as Anthropic.Usage,
    };

    logger.info("NATIVE: OpenAI response completed successfully:", {
      requestId,
      workbookName,
      inputTokens,
      outputTokens,
      totalTokens: inputTokens + outputTokens,
      contentBlockCount: contentBlocks.length,
      hasReasoning: reasoningBlockStarted,
      reasoningLength: reasoningContent.length,
      textLength: textContent.length,
      toolUseBlockCount: toolUseBlocks.length,
      stopReason: responseMessage.stop_reason,
      model: azureModel,
    });

    logger.info("NATIVE: Complete OpenAI response message:", {
      responseMessage: JSON.stringify(responseMessage, null, 2),
      contentBlockCount: contentBlocks.length,
      hasReasoning: reasoningBlockStarted,
      reasoningLength: reasoningContent.length,
      textContent: textContent.substring(0, 200),
      toolUseBlockCount: toolUseBlocks.length,
      stopReason: responseMessage.stop_reason,
    });

    tracing.logGenerationEnd(generationId, {
      output: responseMessage,
      metadata: {
        provider: "azure-openai",
        model: azureModel,
        inputTokens: inputTokens,
        outputTokens: outputTokens,
        totalTokens: inputTokens + outputTokens,
        stopReason: responseMessage.stop_reason,
        contentBlockCount: contentBlocks.length,
        hasReasoning: reasoningBlockStarted,
        reasoningLength: reasoningContent.length,
        textLength: textContent.length,
        toolUseCount: toolUseBlocks.length,
        toolNames: toolUseBlocks.map((t) => t.name).join(", "),
        systemMessage: (params.system as string)?.substring(0, 500) || "",
        messageCount: params.messages.length,
        toolCount: responsesTools.length,
        responseText: textContent.substring(0, 500),
        toolCalls: toolUseBlocks.map((t) => ({
          name: t.name,
          inputPreview: JSON.stringify(t.input).substring(0, 200),
        })),
        // Excel context details
        workbookName: excelContext.workbookName || workbookName,
        activeSheet: excelContext.activeSheet,
        selectedRange: excelContext.selectedRange,
        usedRange: excelContext.usedRange,
      },
    });

    return responseMessage;
  } catch (error: any) {
    if (error instanceof UserCancellationError) {
      logger.info("⏹️ Azure OpenAI request cancelled by user", {
        requestId,
        workbookName,
      });
      throw error;
    }

    logger.error("❌ Error during Azure OpenAI streaming:", {
      requestId,
      errorMessage: error.message,
      errorType: error.constructor?.name,
      errorName: error.name,
      errorCode: error.code,
      statusCode: error.status,
      model: azureModel,
      workbookName,
      // Try to extract more details from the error
      errorCause: error.cause?.message || error.cause,
      errorResponse: error.response,
      errorHeaders: error.headers,
      stack: error.stack?.substring(0, 500),
      troubleshooting: {
        likelyIssue:
          "The endpoint might not support the Responses API path (/openai/v1/responses)",
        checkEndpoint:
          "Verify this is the correct endpoint URL and that it's accessible from your network",
        checkPath:
          "This endpoint might use a different path like /openai/deployments/{deployment-name}/chat/completions",
        checkApiKey:
          "Verify API key format and that it's valid for this endpoint",
        checkNetwork:
          "Check if VPN or network access is required for internal endpoints (.internal)",
        checkApiVersion:
          "The API might require a specific api-version query parameter",
        alternativeApproach:
          "Consider using standard Azure OpenAI Chat Completions API instead of Responses API",
      },
    });
    throw error;
  }
}
