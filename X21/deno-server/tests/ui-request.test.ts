import { assertEquals, assertThrows } from "std/assert";
import { Router } from "../src/router/index.ts";
import { stateManager } from "../src/state/state-manager.ts";
import { validateUiRequestResult } from "../src/utils/ui-request.ts";
import {
  ClaudeContentTypes,
  ToolNames,
  UiRequestControl,
} from "../src/types/index.ts";

Deno.test("validateUiRequestResult sanitizes and enforces schema", () => {
  const controls: UiRequestControl[] = [
    { id: "confirm", kind: "boolean", label: "Proceed?", required: true },
    {
      id: "choice",
      kind: "segmented",
      label: "Cadence",
      options: [
        { id: "monthly", label: "Monthly" },
        { id: "other", label: "Other", allowFreeText: true },
      ],
    },
  ];

  const output = {
    confirm: { value: true },
    choice: { choiceId: "other", freeText: "Quarterly" },
  };

  const result = validateUiRequestResult(controls, output);
  assertEquals(result, {
    confirm: { value: true },
    choice: { choiceId: "other", freeText: "Quarterly" },
  });

  assertThrows(
    () => validateUiRequestResult(controls, {}),
    Error,
    "Missing required response",
  );
});

Deno.test("handleToolResult updates conversation history for collect_input", async () => {
  const workbookName = "ui-request-test.xlsx";
  stateManager.startState(workbookName);
  const requestId = stateManager.getLatestRequestId(workbookName);
  const toolUseId = "tool-123";

  const payload = {
    title: "Need confirmation",
    mode: "blocking",
    controls: [{
      id: "confirm",
      kind: "boolean",
      label: "Proceed?",
      required: true,
    }],
  };

  stateManager.addInitialToolChange(
    workbookName,
    {
      id: toolUseId,
      name: ToolNames.COLLECT_INPUT,
      input: payload,
      type: ClaudeContentTypes.TOOL_USE,
    } as any,
    requestId,
  );

  const placeholder = {
    role: "user",
    content: [{
      type: ClaudeContentTypes.TOOL_RESULT,
      tool_use_id: toolUseId,
      content: "Placeholder - Response Not Received",
    }],
  };
  stateManager.addMessage(workbookName, placeholder as any);

  const router = new Router() as any;
  let continued = false;
  router.continueStreamingAfterToolProcessing = () => {
    continued = true;
    return Promise.resolve();
  };

  await router.handleToolResult({
    workbookName,
    toolUseId,
    output: { confirm: { value: true } },
  });

  const history = stateManager.getConversationHistory(workbookName);
  const updatedToolResult =
    (history.find((msg: any) => msg.role === "user")?.content || []).find(
      (c: any) => c.tool_use_id === toolUseId,
    );

  assertEquals(
    updatedToolResult?.content,
    JSON.stringify({ confirm: { value: true } }),
  );

  const change = stateManager.getToolChange(workbookName, toolUseId);
  assertEquals(change.pending, false);
  assertEquals(change.approved, true);
  assertEquals(continued, true);

  stateManager.deleteWorkbookState(workbookName);
});
