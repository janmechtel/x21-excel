import Anthropic from "@anthropic-ai/sdk";
import { tools } from "../tools/index.ts";
import { createLogger } from "../utils/logger.ts";
import { getChatConversationSystemMessage } from "../prompts/chat.ts";
import { getCompactConversationSystemMessage } from "../prompts/compact.ts";
import { getAnthropicConfig } from "./provider.ts";

const logger = createLogger("Anthropic-LLM-Client");

export const maxTokens = 32000;
export const tokenLimit = (200000 - maxTokens) * 0.6;
export const anthropicBetas = [
  "interleaved-thinking-2025-05-14",
  "structured-outputs-2025-11-13",
] as const;
const thinkingBudgetTokens = 1600;

const caBundleFetchCache = new Map<string, typeof fetch>();

function getFetchWithCaBundle(caBundlePath?: string): typeof fetch | undefined {
  if (!caBundlePath) {
    logger.info("Anthropic CA bundle not configured; using default TLS trust");
    return undefined;
  }

  const cached = caBundleFetchCache.get(caBundlePath);
  if (cached) {
    logger.info("Using cached Anthropic CA bundle HTTP client", {
      caBundlePath,
    });
    return cached;
  }

  try {
    logger.info("Loading Anthropic CA bundle from disk", { caBundlePath });
    const pemContents = Deno.readTextFileSync(caBundlePath);
    if (!pemContents.trim()) {
      logger.error("Anthropic CA bundle file is empty", { caBundlePath });
      throw new Error("CA bundle file is empty");
    }
    logger.info("Anthropic CA bundle loaded", {
      caBundlePath,
      pemLength: pemContents.length,
    });

    const httpClient = Deno.createHttpClient({ caCerts: [pemContents] });
    const fetchWithClient: typeof fetch = (input, init) => {
      const initWithClient = {
        ...(init ?? {}),
      } as RequestInit & { client: Deno.HttpClient };
      initWithClient.client = httpClient;
      return fetch(input, initWithClient);
    };

    caBundleFetchCache.set(caBundlePath, fetchWithClient);
    logger.info("Anthropic CA bundle HTTP client ready", {
      caBundlePath,
    });
    return fetchWithClient;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error("Failed to load Anthropic CA bundle", {
      caBundlePath,
      error: message,
    });
    throw error;
  }
}

export function createAnthropicClient(): Anthropic {
  const config = getAnthropicConfig();

  if (!config) {
    logger.error("No Anthropic configuration found");
    throw new Error(
      "Anthropic API key not configured. Please set ANTHROPIC_API_KEY environment variable.",
    );
  }

  logger.info("Creating Anthropic client with centralized config", {
    baseUrl: config.baseUrl || "[default]",
    hasCaBundlePath: !!config.caBundlePath,
  });
  const fetchOverride = getFetchWithCaBundle(config.caBundlePath);
  return new Anthropic({
    apiKey: config.apiKey,
    baseURL: config.baseUrl,
    ...(fetchOverride ? { fetch: fetchOverride } : {}),
  });
}

function getModel(): string {
  const config = getAnthropicConfig();
  return config?.model!;
}

type AnthropicParams =
  | Anthropic.MessageCreateParamsNonStreaming
  | Anthropic.MessageCountTokensParams;

export function withAnthropicBetas<T extends AnthropicParams>(
  params: T,
): T & { betas: string[] } {
  return {
    ...params,
    betas: [...anthropicBetas],
  };
}

export function buildAnthropicParamsForConversatoin(
  messageHistory: Anthropic.MessageParam[],
  active_tools: string[],
): Anthropic.MessageCreateParamsNonStreaming {
  logger.info("Building Claude params with message history");
  logger.info("Active tools: ", active_tools);

  const filteredTools = tools.filter((tool) =>
    active_tools.includes(tool.name)
  );
  logger.info(`Filtered tools: ${filteredTools.map((t) => t.name).join(", ")}`);

  return {
    model: getModel(),
    max_tokens: maxTokens,
    system: getChatConversationSystemMessage(),
    messages: messageHistory,
    thinking: { type: "enabled", budget_tokens: thinkingBudgetTokens },
    tool_choice: {
      type: "auto",
      // Cast to allow disable_parallel_tool_use until SDK types catch up
      disable_parallel_tool_use: false,
    } as Anthropic.MessageCreateParamsNonStreaming["tool_choice"] & {
      disable_parallel_tool_use?: boolean;
    },
    tools: filteredTools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: tool.input_schema,
    })) as Anthropic.MessageCreateParamsNonStreaming["tools"],
  };
}

export function buildAnthropicParamsForTokenCount(
  messageHistory: Anthropic.MessageParam[],
  active_tools: string[],
): Anthropic.MessageCountTokensParams {
  logger.debug("Building Claude params with message history");
  logger.debug("Active tools: ", active_tools);

  const filteredTools = tools.filter((tool) =>
    active_tools.includes(tool.name)
  );
  logger.debug(
    `Filtered tools: ${filteredTools.map((t) => t.name).join(", ")}`,
  );

  return {
    model: getModel(),
    system: getChatConversationSystemMessage(),
    messages: messageHistory,
    thinking: { type: "enabled", budget_tokens: thinkingBudgetTokens },
    tool_choice: {
      type: "auto",
      // Cast to allow disable_parallel_tool_use until SDK types catch up
      disable_parallel_tool_use: false,
    } as unknown as { type: "auto"; disable_parallel_tool_use?: boolean },
    tools: filteredTools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: tool.input_schema,
    })) as Anthropic.MessageCountTokensParams["tools"],
  };
}

export function buildAnthropicParamsForCompact(
  messageHistory: Anthropic.MessageParam[],
): Anthropic.MessageCreateParamsNonStreaming {
  logger.debug("Building Claude params for compact");
  logger.debug("Message history: ", messageHistory);

  return {
    model: getModel(),
    max_tokens: 8000,
    system: getCompactConversationSystemMessage(),
    messages: messageHistory,
  };
}
