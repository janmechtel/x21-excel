import Anthropic, { APIUserAbortError } from "@anthropic-ai/sdk";
import { createLogger } from "../utils/logger.ts";
import { tracing } from "../tracing/tracing.ts";
import { WebSocketManager } from "../services/websocket-manager.ts";
import { ErrorCode } from "../errors/error-codes.ts";
import { stateManager } from "../state/state-manager.ts";
import { UserCancellationError } from "../errors/user-cancellation-error.ts";
import {
  ClaudeEventTypes,
  ContentBlockTypes,
  OperationStatusValues,
} from "../types/index.ts";
import { withAnthropicBetas } from "../llm-client/anthropic.ts";

const logger = createLogger("streamClaudeResponseToWebSocket");
const socket = WebSocketManager.getInstance();

async function streamClaudeResponseToWebSocket(
  requestId: string,
  client: Anthropic,
  params: Anthropic.MessageCreateParamsNonStreaming,
  abortController: AbortController,
): Promise<Anthropic.Message> {
  const generationId = tracing.logGenerationStart(requestId, params, {
    provider: "anthropic",
  });
  const workbookName: string = stateManager.getWorkbookName(requestId);

  try {
    // Send connecting status
    socket.sendStatus(
      workbookName,
      OperationStatusValues.CONNECTING,
      "Connecting to Claude...",
    );

    const stream = client.beta.messages.stream(
      withAnthropicBetas(params),
      {
        signal: abortController?.signal,
      },
    );

    let customThinkingText = "";
    let customThinkingSignature = "";

    logger.info("Start Streaming");

    // Send generating status when streaming starts
    socket.sendStatus(
      workbookName,
      OperationStatusValues.GENERATING_LLM,
      "Generating response...",
    );

    for await (const event of stream) {
      if (abortController?.signal.aborted) {
        stream.abort();
        throw new UserCancellationError(
          "User cancelled the request",
          requestId,
        );
      }

      socket.send(workbookName, "stream:delta", event);

      // Stream token updates in real-time
      if (event.type === ClaudeEventTypes.MESSAGE_DELTA && event.usage) {
        const inputTokens = (event as any).usage?.input_tokens || 0;
        const outputTokens = event.usage.output_tokens || 0;
        socket.sendTokenUpdate(
          workbookName,
          inputTokens,
          outputTokens,
          inputTokens + outputTokens,
        );
      }

      if (
        event.type === ClaudeEventTypes.CONTENT_BLOCK_DELTA &&
        event?.delta.type === "thinking_delta"
      ) {
        customThinkingText += event.delta.thinking;
      }
      if (
        event.type === ClaudeEventTypes.CONTENT_BLOCK_START &&
        event?.content_block.type === ContentBlockTypes.THINKING
      ) {
        customThinkingSignature = event.content_block.signature || "";
      }
    }

    const finalMessage = await stream.finalMessage();
    const enhancedFinalMessage = enhanceMessageWithThinking2(
      finalMessage as any,
      customThinkingText,
      customThinkingSignature,
    );

    const output_tokens = enhancedFinalMessage.usage?.output_tokens || 0;
    const input_tokens = enhancedFinalMessage.usage?.input_tokens || 0;

    tracing.logGenerationEnd(generationId, {
      output: { ...enhancedFinalMessage },
      metadata: {
        inputTokens: input_tokens,
        outputTokens: output_tokens,
        totalTokens: input_tokens + output_tokens,
      },
    });

    return enhancedFinalMessage;
  } catch (error: any) {
    tracing.logGenerationEnd(generationId, {
      output: {
        success: false,
        error: `Stream Error: ${error.message}`,
        streamingFailed: true,
      },
      metadata: { inputTokens: 0, outputTokens: 0, totalTokens: 0 },
    });

    if (error instanceof APIUserAbortError) {
      throw new UserCancellationError("User cancelled the request", requestId);
    }

    // Handle malformed JSON errors from Anthropic SDK when parsing tool parameters
    if (
      error?.message?.includes("Unable to parse tool parameter JSON") ||
      error?.message?.includes("Bad control character in string literal")
    ) {
      logger.error(
        "Anthropic SDK failed to parse tool parameters due to malformed JSON:",
        {
          errorMessage: error.message,
          errorName: error.name,
          requestId,
          workbookName,
        },
      );
      // Re-throw with a more user-friendly message
      const friendlyError = new Error(
        "The model generated invalid tool parameters. Please try again or rephrase your request.",
      );
      (friendlyError as any).originalError = error;
      throw friendlyError;
    }

    throw error;
  }
}

/**
 * Maps Claude API error types to internal error code payloads
 * @param errorType - The Claude error type extracted from the error message
 * @returns Payload object with appropriate error code
 */
function mapClaudeErrorToPayload(
  errorType: string | null,
): { type: ErrorCode } {
  switch (errorType) {
    case "overloaded_error":
      return {
        type: ErrorCode.OVERLOADED,
      };
    case "request_too_large":
      return {
        type: ErrorCode.REQUEST_TOO_LARGE,
      };
    case "rate_limit_error":
      return {
        type: ErrorCode.RATE_LIMIT_ERROR,
      };
    case "invalid_request_error":
      return {
        type: ErrorCode.INVALID_REQUEST_ERROR,
      };
    default:
      return {
        type: ErrorCode.OTHER,
      };
  }
}

function enhanceMessageWithThinking2(
  finalMessage: Anthropic.Message,
  customThinkingText: string,
  customThinkingSignature: string,
): Anthropic.Message {
  const thinkingContent = finalMessage.content?.find((content) =>
    content.type === ContentBlockTypes.THINKING
  );
  const finalMessageThinking = thinkingContent?.thinking || "";
  const finalMessageSignature = thinkingContent?.signature || "";

  const thinkingText = finalMessageThinking || customThinkingText;
  const thinkingSignature = finalMessageSignature || customThinkingSignature;

  const enhancedFinalMessage = {
    ...finalMessage,
    content: finalMessage.content?.map((contentItem) => {
      if (contentItem.type === ContentBlockTypes.THINKING) {
        return {
          ...contentItem,
          thinking: thinkingText,
          signature: thinkingSignature,
        };
      }
      return contentItem;
    }) || [],
  };

  return enhancedFinalMessage;
}

/**
 * Extracts the Claude API error type from error message string
 * @param errorMessage - The error message string from Claude API
 * @returns The error type as a string, or null if not found
 */
function extractClaudeErrorType(errorMessage: string): string | null {
  try {
    // Remove HTTP status code prefix if present (e.g., "400 ")
    const jsonStart = errorMessage.indexOf("{");
    const jsonPart = jsonStart !== -1
      ? errorMessage.substring(jsonStart)
      : errorMessage;

    // Parse the JSON error message
    const errorData = JSON.parse(jsonPart);
    return errorData?.error?.type || null;
  } catch (_parseError) {
    // Fallback: try to extract using regex
    const typeMatch = errorMessage.match(/"error":\s*{\s*"type":\s*"([^"]+)"/);
    return typeMatch?.[1] || null;
  }
}

export {
  extractClaudeErrorType,
  mapClaudeErrorToPayload,
  streamClaudeResponseToWebSocket,
};
