import { assert, assertEquals } from "std/assert";
import { stateManager } from "../src/state/state-manager.ts";
import {
  buildAnthropicParamsForConversatoin,
  buildAnthropicParamsForTokenCount,
  maxTokens,
  tokenLimit,
} from "../src/llm-client/anthropic.ts";
import { createAnthropicClient } from "../src/llm-client/anthropic.ts";
import { streamClaudeResponseAndHandleToolUsage } from "../src/stream/tool-logic.ts";
import { tracing } from "../src/tracing/tracing.ts";
import { UserService } from "../src/services/user.ts";
import {
  compactChatHistory,
  isHistoryExceedingTokenLimit,
} from "../src/compacting/compacting.ts";
import { ClaudeStopReasons, ToolNames } from "../src/types/index.ts";

Deno.test({
  name: "test generation with 171k token history and max_tokens=32000",
  ignore: !Deno.env.get("ANTHROPIC_API_KEY"),
  sanitizeResources: false,
  sanitizeOps: false,
  async fn() {
    const workbookName = "test-worksheet.xlsx";
    const requestId = "req-test-171k";

    // Initialize workbook state
    stateManager.startState(workbookName);
    const workbookState = stateManager.getState(workbookName);
    workbookState.latestRequestId = requestId;
    stateManager.creatingAbortController(workbookName, requestId);

    // Populate history with approximately 175,000 tokens
    // Using the same approximation as token-limit-revert.test.ts:
    // ~2.5 characters per token, so 175k tokens ≈ 437,500 characters
    const longHistoryMessage = "A very long message".repeat(43000); // ~171k tokens
    stateManager.setConversationHistory(workbookName, [
      { role: "user", content: longHistoryMessage },
    ]);

    // Add a new user message for the generation request
    const newPrompt = "Please help me with this task.";
    stateManager.addMessage(workbookName, {
      role: "user",
      content: newPrompt,
    });

    // Initialize tracing
    const sessionId = stateManager.getSessionId(workbookName);
    tracing.startTrace(requestId, {
      name: "Excel AI Workflow",
      userEmail: "test@example.com",
      input: newPrompt,
      sessionId: sessionId,
    }, workbookName);

    // Set up request metadata
    const activeTools = [
      ToolNames.READ_VALUES_BATCH,
      ToolNames.READ_FORMAT_BATCH,
    ];
    const requestMetadata = {
      activeTools: activeTools,
      workbookName: workbookName,
      worksheets: [],
      activeWorksheet: "",
    };
    stateManager.saveRequestMetadata(workbookName, requestId, requestMetadata);

    // Get the conversation history
    let history = stateManager.getConversationHistory(workbookName);

    // Check if conversation history exceeds token limit
    const claudeClient = createAnthropicClient();
    const tokenCount = await claudeClient.messages.countTokens(
      buildAnthropicParamsForTokenCount(history, activeTools),
    );

    assert(
      tokenCount.input_tokens > tokenLimit &&
        tokenCount.input_tokens < 200000,
      `Token count (${tokenCount.input_tokens}) should be above token limit (${tokenLimit}) and below 200000`,
    );

    if (
      await isHistoryExceedingTokenLimit(history, tokenLimit, activeTools) ===
        true
    ) {
      console.log(
        `\nHistory exceeds token limit (${tokenLimit}), compacting...`,
      );
      history = await compactChatHistory(requestMetadata, history);
      console.log(`   - Compacted history messages: ${history.length}`);
    } else {
      console.log(
        `\nHistory does not exceed token limit (${tokenLimit}), not compacting...`,
      );
    }

    const params = buildAnthropicParamsForConversatoin(history, activeTools);

    // Verify max_tokens is set correctly
    assertEquals(
      params.max_tokens,
      maxTokens,
      `max_tokens should be ${maxTokens}`,
    );

    // Log token count information
    console.log(`\nTest setup:`);
    console.log(`   - History messages: ${history.length}`);
    console.log(`   - Max tokens: ${params.max_tokens}`);
    console.log(`   - Model: ${params.model}`);

    // Make an actual generation request
    console.log(`\nMaking generation request with max_tokens=${maxTokens}...`);
    UserService.getInstance().setUserEmail("test@example.com");

    // This will make an actual API call - comment out if you don't want to make real calls
    const response = await streamClaudeResponseAndHandleToolUsage(
      requestId,
      createAnthropicClient(),
      params,
    );

    // Assert that completion was successful
    assert(response !== null, "Response should not be null");
    assert(response !== undefined, "Response should not be undefined");
    assert(
      response.stop_reason === "end_turn" ||
        response.stop_reason === ClaudeStopReasons.TOOL_USE,
      `Completion should be successful. Expected stop_reason to be 'end_turn' or '${ClaudeStopReasons.TOOL_USE}', but got '${response.stop_reason}'`,
    );
    assert(
      Array.isArray(response.content) && response.content.length > 0,
      "Response should have non-empty content",
    );

    console.log(`\nGeneration completed successfully`);
    console.log(`   - Stop reason: ${response.stop_reason}`);
    console.log(
      `   - Content length: ${JSON.stringify(response.content).length}`,
    );
    console.log(
      `\nResponse content: ${JSON.stringify(response.content, null, 2)}`,
    );

    // Cleanup
    stateManager.deleteWorkbookState(workbookName);
    console.log(`\nCleanup completed`);
  },
});
