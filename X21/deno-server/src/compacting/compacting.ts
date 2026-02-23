import Anthropic from "@anthropic-ai/sdk";
import { RequestMetadata, stateManager } from "../state/state-manager.ts";
import {
  buildAnthropicParamsForCompact,
  buildAnthropicParamsForTokenCount,
  createAnthropicClient,
  withAnthropicBetas,
} from "../llm-client/anthropic.ts";
import { createLogger } from "../utils/logger.ts";
import { ContentBlockTypes } from "../types/index.ts";

const logger = createLogger("Compacting");

/**
 * Build a compacted prompt with conversation summary and current context
 */
async function buildCompactedPrompt(
  requestMetadata: RequestMetadata,
  lastEntry: Anthropic.MessageParam,
  allButLast: Anthropic.MessageParam[],
): Promise<string> {
  logger.info("Creating summary from history");
  const summary: string = await createSummaryFromHistory(
    allButLast,
  );

  const userPrompt: string = typeof lastEntry.content === "string"
    ? lastEntry.content
    : JSON.stringify(lastEntry.content);

  const newPrompt: string = `
    Summarized Context of previous conversation:

    ${summary}

    Current User Prompt:

    ${userPrompt}

    Current workbook name: ${requestMetadata.workbookName}

    Current worksheets: ${requestMetadata.worksheets}

    Active worksheet: ${requestMetadata.activeWorksheet}

    Current active tools: ${requestMetadata.activeTools}
    `;

  return newPrompt;
}

export async function compactChatHistory(
  requestMetadata: RequestMetadata,
  messages: Anthropic.MessageParam[],
): Promise<Anthropic.MessageParam[]> {
  logger.info("Compacting chat history");
  logger.info("Messages length: ", messages.length);

  logger.info("Slicing messages");
  const lastEntry = messages.slice(-1)[0];
  const allButLast = messages.slice(0, -1);

  logger.info("Building compacted prompt");
  const newPrompt = await buildCompactedPrompt(
    requestMetadata,
    lastEntry,
    allButLast,
  );

  logger.info("Creating new compacted conversation entry");
  const newCompactedConversationEntry: Anthropic.MessageParam = {
    role: "user",
    content: newPrompt,
  };

  const newCompactedConversationHistory: Anthropic.MessageParam[] = [
    newCompactedConversationEntry,
  ];

  logger.info("Setting new compacted conversation history");
  stateManager.setConversationHistory(
    requestMetadata.workbookName,
    newCompactedConversationHistory,
  );
  return newCompactedConversationHistory;
}

async function createSummaryFromHistory(
  allButLast: Anthropic.Messages.MessageParam[],
): Promise<string> {
  const conversationHistory: string = JSON.stringify(allButLast);

  logger.info("Creating prompt");
  const prompt = `
    Summarize the following conversation history?
    ${conversationHistory}
    `;

  logger.info("Creating prompt params");
  const promptParams: Anthropic.MessageParam[] = [
    {
      role: "user",
      content: prompt,
    },
  ];

  logger.info("Building Claude params");
  const claudeParams: Anthropic.MessageCreateParamsNonStreaming =
    buildAnthropicParamsForCompact(promptParams);

  logger.info("Creating Claude client");
  const claudeClient: Anthropic = createAnthropicClient();

  logger.info("Creating summary from history");

  try {
    logger.info("Calling Claude client");
    const response: Anthropic.Message = await claudeClient.beta.messages.create(
      withAnthropicBetas(claudeParams),
    ) as Anthropic.Message;
    logger.info("Claude client called");

    const firstBlock = response.content[0];
    if (firstBlock.type === ContentBlockTypes.TEXT) {
      logger.info("Returning summary");
      logger.info("Summary", { summary: firstBlock.text });
      return firstBlock.text;
    }
    throw new Error("Expected text response from Claude");
  } catch (error: any) {
    throw error;
  }
}

export async function isHistoryExceedingTokenLimit(
  conversationHistory: Anthropic.MessageParam[],
  tokenLimit: number,
  activeTools: string[],
): Promise<boolean> {
  const claudeClient: Anthropic = createAnthropicClient();
  const params = buildAnthropicParamsForTokenCount(
    conversationHistory,
    activeTools,
  );

  const response: Anthropic.MessageTokensCount = await claudeClient.beta
    .messages
    .countTokens(
      withAnthropicBetas(params),
    ) as Anthropic.MessageTokensCount;

  const totalTokens = response.input_tokens;

  logger.info("Input tokens: ", response.input_tokens);
  logger.info("Token limit: ", tokenLimit);

  return totalTokens > tokenLimit;
}
