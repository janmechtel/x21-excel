import OpenAI from "openai";
import type { AzureOpenAI } from "openai";
import Anthropic from "@anthropic-ai/sdk";
import { getChatConversationSystemMessage } from "../prompts/chat.ts";
import { createLogger } from "../utils/logger.ts";
import { getAzureOpenAIConfig } from "./provider.ts";
import { ClaudeContentTypes, ContentBlockTypes } from "../types/index.ts";

const logger = createLogger("OpenAINative");

export type LLMProvider = "anthropic" | "azure";

export interface OpenAIProviderOptions {
  model?: string;
}

function normalizeToolSchema(schema: unknown): Record<string, any> {
  if (!schema || typeof schema !== "object" || Array.isArray(schema)) {
    return { type: "object", properties: {}, additionalProperties: false };
  }

  const schemaObj = schema as Record<string, any>;
  const hasType = typeof schemaObj.type === "string";
  const hasProperties = schemaObj.properties &&
    typeof schemaObj.properties === "object" &&
    !Array.isArray(schemaObj.properties);

  if (!hasType && !hasProperties) {
    return {
      type: "object",
      properties: schemaObj,
      additionalProperties: false,
    };
  }

  const normalized = { ...schemaObj };
  if (!normalized.type) normalized.type = "object";
  if (!hasProperties) normalized.properties = {};

  if (
    normalized.additionalProperties === undefined ||
    (typeof normalized.additionalProperties !== "boolean" &&
      typeof normalized.additionalProperties !== "object") ||
    Array.isArray(normalized.additionalProperties)
  ) {
    normalized.additionalProperties = false;
  }

  normalized.properties = normalizeSchemaProperties(normalized.properties);
  const propertyKeys = Object.keys(normalized.properties || {});
  if (propertyKeys.length > 0) {
    if (!Array.isArray(normalized.required)) {
      normalized.required = propertyKeys;
    } else {
      const requiredSet = new Set(normalized.required);
      for (const key of propertyKeys) {
        requiredSet.add(key);
      }
      normalized.required = Array.from(requiredSet);
    }
  }
  return normalized;
}

function normalizeSchemaProperties(
  properties: Record<string, any>,
): Record<string, any> {
  const normalizedProps: Record<string, any> = {};

  for (const [key, value] of Object.entries(properties || {})) {
    normalizedProps[key] = normalizeSchemaValue(value);
  }

  return normalizedProps;
}

function normalizeSchemaValue(value: any): any {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const normalized = { ...value };
  const hasProperties = normalized.properties &&
    typeof normalized.properties === "object" &&
    !Array.isArray(normalized.properties);
  const isObjectSchema = normalized.type === "object" || hasProperties;

  if (isObjectSchema) {
    if (!normalized.type) normalized.type = "object";
    if (
      normalized.additionalProperties === undefined ||
      (typeof normalized.additionalProperties !== "boolean" &&
        typeof normalized.additionalProperties !== "object") ||
      Array.isArray(normalized.additionalProperties)
    ) {
      normalized.additionalProperties = false;
    }

    if (hasProperties) {
      normalized.properties = normalizeSchemaProperties(normalized.properties);
      const propertyKeys = Object.keys(normalized.properties || {});
      if (propertyKeys.length > 0) {
        if (!Array.isArray(normalized.required)) {
          normalized.required = propertyKeys;
        } else {
          const requiredSet = new Set(normalized.required);
          for (const key of propertyKeys) {
            requiredSet.add(key);
          }
          normalized.required = Array.from(requiredSet);
        }
      }
    }
  }

  if (normalized.type === "array" && normalized.items) {
    normalized.items = normalizeSchemaValue(normalized.items);
  }

  return normalized;
}

/**
 * Create Azure OpenAI client for Responses API
 * The Responses API uses v1 endpoint structure (no deployment in path)
 */
export function createNativeOpenAIClient(): {
  client: OpenAI | AzureOpenAI;
  model: string;
} {
  const config = getAzureOpenAIConfig();

  if (!config) {
    logger.error("✗ No Azure OpenAI configuration found");
    throw new Error(
      "Azure OpenAI not configured. Please configure via the app settings UI.",
    );
  }

  // Try Responses API path first (v1 endpoint)
  const apiVersion = config.apiVersion || "2024-08-01-preview";
  const baseURL = `${config.endpoint}/openai/v1/`;

  const clientConfig = {
    apiKey: config.apiKey,
    baseURL: baseURL,
    defaultHeaders: {
      "api-key": config.apiKey,
      "api-version": apiVersion,
    },
  };

  logger.info(
    "✓ Creating Azure OpenAI Responses API client:",
    {
      endpoint: config.endpoint,
      baseURL: baseURL,
      apiVersion: apiVersion,
      fullResponsesPath: `${baseURL}responses`,
      deploymentName: config.deploymentName,
      model: config.model,
      modelMatchesDeployment: config.model === config.deploymentName,
      reasoningEffort: config.reasoningEffort,
      hasApiKey: !!config.apiKey,
      apiKeyLength: config.apiKey.length,
      apiKeyPrefix: config.apiKey.substring(0, 4) + "...",
      headers: {
        "api-key": "[REDACTED]",
        "api-version": apiVersion,
      },
    },
  );

  try {
    const client = new OpenAI(clientConfig) as AzureOpenAI;

    logger.info("✓ Azure OpenAI client created successfully", {
      modelToUse: config.model,
      clientConfigured: true,
      baseURL: baseURL,
    });

    return { client, model: config.model };
  } catch (error: any) {
    logger.error("❌ Failed to create Azure OpenAI client", {
      error: error.message,
      endpoint: config.endpoint,
      baseURL: baseURL,
    });
    throw error;
  }
}

/**
 * Convert Anthropic message format to Responses API input format
 */
export function convertAnthropicToResponsesInput(
  messages: Anthropic.MessageParam[],
  system?: string,
): any[] {
  const converted: any[] = [];
  let messageIdCounter = 0;

  const systemMessage = system || getChatConversationSystemMessage();

  // Add system message as first user message
  if (systemMessage) {
    converted.push({
      id: `msg_${messageIdCounter++}_${Date.now()}`,
      status: "completed" as const,
      type: "message" as const,
      role: "user" as const,
      content: [{
        type: "input_text" as const,
        text: `[System]: ${systemMessage}`,
      }],
    });
  }

  logger.info("Converting Anthropic messages to Responses API format:", {
    messageCount: messages.length,
  });

  for (const message of messages) {
    const baseMessage = {
      id: `msg_${messageIdCounter++}_${Date.now()}`,
      status: "completed" as const,
      type: "message" as const,
    };

    // Handle simple string content
    if (typeof message.content === "string") {
      const isAssistant = message.role === "assistant";
      converted.push({
        ...baseMessage,
        role: isAssistant ? "assistant" as const : "user" as const,
        content: [{
          type: isAssistant ? "output_text" as const : "input_text" as const,
          text: message.content,
        }],
      });
      continue;
    }

    // Handle complex content blocks
    const contentBlocks = message.content;
    const textBlocks = contentBlocks.filter((c: any) =>
      c.type === ContentBlockTypes.TEXT
    );
    const toolUseBlocks = contentBlocks.filter((c: any) =>
      c.type === ClaudeContentTypes.TOOL_USE
    );
    const toolResultBlocks = contentBlocks.filter((c: any) =>
      c.type === ClaudeContentTypes.TOOL_RESULT
    );

    // Handle tool results (user messages with tool outputs)
    if (toolResultBlocks.length > 0) {
      for (const block of toolResultBlocks) {
        const toolResult = block as any;
        let contentStr = "";

        if (typeof toolResult.content === "string") {
          contentStr = toolResult.content;
        } else if (Array.isArray(toolResult.content)) {
          contentStr = toolResult.content
            .map((
              c: any,
            ) => (typeof c === "string" ? c : c.text || JSON.stringify(c)))
            .join("\n");
        } else if (toolResult.content) {
          contentStr = JSON.stringify(toolResult.content);
        }

        converted.push({
          ...baseMessage,
          id: `msg_${messageIdCounter++}_${Date.now()}`,
          role: "user" as const,
          content: [{
            type: "input_text" as const,
            text: `[Tool result for ${toolResult.tool_use_id}]: ${contentStr}`,
          }],
        });
      }
      continue;
    }

    // Handle assistant messages with tool calls
    if (toolUseBlocks.length > 0) {
      const text = textBlocks.map((block: any) => block.text || "").join("");
      const toolCallsText = toolUseBlocks
        .map((block: any) =>
          `${block.name}(${JSON.stringify(block.input || {})})`
        )
        .join(", ");

      logger.info("Creating assistant message with tool calls:", {
        toolCallCount: toolUseBlocks.length,
        toolCalls: toolUseBlocks.map((block: any) => ({
          id: block.id,
          name: block.name,
        })),
      });

      converted.push({
        ...baseMessage,
        role: "assistant" as const,
        content: [{
          type: "output_text" as const,
          text: text || `Using tools: ${toolCallsText}`,
        }],
      });
      continue;
    }

    // Regular text message
    const text = textBlocks.map((block: any) => block.text || "").join("");
    const isAssistant = message.role === "assistant";
    converted.push({
      ...baseMessage,
      role: isAssistant ? "assistant" as const : "user" as const,
      content: [{
        type: isAssistant ? "output_text" as const : "input_text" as const,
        text,
      }],
    });
  }

  return converted;
}

/**
 * Convert Anthropic tools to Responses API function tools format
 */
export function convertAnthropicToolsToResponsesFormat(
  tools: any[] | undefined,
): any[] {
  if (!tools || !Array.isArray(tools)) {
    logger.info("No tools to convert");
    return [];
  }

  logger.info("Converting Anthropic tools to Responses API format:", {
    toolCount: tools.length,
  });

  return tools.map((tool) => ({
    type: "function" as const,
    name: tool.name,
    description: tool.description,
    parameters: normalizeToolSchema(
      tool.input_schema || {
        type: "object",
        properties: {},
      },
    ),
    strict: true,
  }));
}
