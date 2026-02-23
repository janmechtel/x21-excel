import Anthropic from "@anthropic-ai/sdk";
import { assert, assertEquals, assertStringIncludes } from "std/assert";
import { executeToolUsesWithConcurrency } from "../src/llm-client/tool-runner.ts";
import {
  ClaudeContentTypes,
  ClaudeStopReasons,
  ContentBlockTypes,
} from "../src/types/index.ts";

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

Deno.test("Multiple tool_use blocks are executed and returned in one user message", async () => {
  const toolUses: Anthropic.ToolUseBlock[] = [
    {
      type: ClaudeContentTypes.TOOL_USE,
      id: "tool-1",
      name: "read_values_batch",
      input: { a: 1 },
    },
    {
      type: ClaudeContentTypes.TOOL_USE,
      id: "tool-2",
      name: "get_metadata",
      input: { b: 2 },
    },
    {
      type: ClaudeContentTypes.TOOL_USE,
      id: "tool-3",
      name: "write_values_batch",
      input: { c: 3 },
    },
  ];

  const results = await executeToolUsesWithConcurrency(toolUses, {
    executor: async (tool) => {
      await delay(10);
      return { handled: tool.id };
    },
  });

  assertEquals(results.length, 3);
  assertEquals(
    results.map((r) => r.tool_use_id),
    ["tool-1", "tool-2", "tool-3"],
  );
  results.forEach((result) => {
    assertEquals(result.type, ClaudeContentTypes.TOOL_RESULT);
    assert(!result.is_error);
    const content = typeof result.content === "string"
      ? JSON.parse(result.content)
      : result.content;
    assertEquals(content.handled, result.tool_use_id);
  });
});

Deno.test("Assistant content retains thinking block and multiple tool_results are grouped", async () => {
  const messages: Anthropic.MessageParam[] = [];
  const response: Anthropic.Message = {
    id: "msg-1",
    role: "assistant",
    model: "claude-sonnet-4-20250514",
    content: [
      {
        type: ContentBlockTypes.THINKING,
        thinking: "step plan",
        signature: "sig",
      } as Anthropic.ThinkingBlock,
      {
        type: ClaudeContentTypes.TOOL_USE,
        id: "tool-1",
        name: "get_metadata",
        input: {},
      } as Anthropic.ToolUseBlock,
      {
        type: ClaudeContentTypes.TOOL_USE,
        id: "tool-2",
        name: "read_values_batch",
        input: {},
      } as Anthropic.ToolUseBlock,
    ],
    stop_reason: ClaudeStopReasons.TOOL_USE,
    stop_sequence: null,
    usage: { input_tokens: 1, output_tokens: 1 } as Anthropic.Usage,
    type: "message",
  };

  messages.push({ role: "assistant", content: response.content });
  const toolBlocks = response.content.filter(
    (block): block is Anthropic.ToolUseBlock =>
      block.type === ClaudeContentTypes.TOOL_USE,
  );
  const toolResults = await executeToolUsesWithConcurrency(toolBlocks, {
    executor: (tool) => Promise.resolve(tool.id),
  });
  messages.push({ role: "user", content: toolResults });

  const assistantContent = messages[0].content as Anthropic.ContentBlock[];
  assertEquals(assistantContent[0]?.type, ContentBlockTypes.THINKING);
  assertEquals((assistantContent[0] as any).thinking, "step plan");

  const resultContent = messages[1].content as Anthropic.ToolResultBlockParam[];
  assertEquals(resultContent.length, 2);
  resultContent.forEach((c) =>
    assertEquals(c.type, ClaudeContentTypes.TOOL_RESULT)
  );
});

Deno.test("Tool error handling surfaces is_error without crashing the loop", async () => {
  const toolUses: Anthropic.ToolUseBlock[] = [
    {
      type: ClaudeContentTypes.TOOL_USE,
      id: "ok-tool",
      name: "get_metadata",
      input: {},
    },
    {
      type: ClaudeContentTypes.TOOL_USE,
      id: "fail-tool",
      name: "read_values_batch",
      input: {},
    },
  ];

  const results = await executeToolUsesWithConcurrency(toolUses, {
    executor: (tool) => {
      if (tool.id === "fail-tool") {
        return Promise.reject(new Error("boom"));
      }
      return Promise.resolve("fine");
    },
  });

  const ok = results.find((r) => r.tool_use_id === "ok-tool")!;
  const err = results.find((r) => r.tool_use_id === "fail-tool")!;

  assertEquals(ok.is_error, undefined);
  assertEquals(err.is_error, true);
  const errContent = typeof err.content === "string"
    ? JSON.parse(err.content)
    : err.content;
  assertStringIncludes(errContent.message, "boom");
});
