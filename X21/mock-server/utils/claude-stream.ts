/**
 * Claude API Streaming Format Simulator
 *
 * Simulates the exact streaming format that Claude API sends
 * Matches the event structure expected by useWebSocketStream hook
 */

import { ContentBlockTypes, ClaudeContentTypes, ClaudeEventTypes, WebSocketMessageTypes, type ContentBlockType } from "../../shared/types/index.ts";

interface ContentBlock {
  type: ContentBlockType;
  content: string;
  toolName?: string;
  toolInput?: Record<string, unknown>;
}

interface StreamScenario {
  blocks: ContentBlock[];
  delayMs?: number; // Delay between chunks
  requestToolApproval?: boolean;
  toolPermissions?: Array<{
    toolId: string;
    toolName: string;
    toolInput: Record<string, unknown>;
  }>;
  uiRequest?: {
    toolUseId: string;
    payload: Record<string, unknown>;
  };
}

/**
 * Simulates Claude API streaming events
 */
export async function simulateClaudeStream(
  socket: WebSocket,
  scenario: StreamScenario,
) {
  const delayMs = scenario.delayMs || 50;
  let blockIndex = 0;

  // Send message_start event
  sendEvent(socket, {
    type: ClaudeEventTypes.MESSAGE_START,
    message: {
      id: `msg_${Date.now()}`,
      type: "message",
      role: "assistant",
      content: [],
      model: "claude-sonnet-4",
      stop_reason: null,
      usage: { input_tokens: 100, output_tokens: 0 },
    },
  });

  await delay(delayMs);

  // Stream each content block
  for (const block of scenario.blocks) {
    await streamContentBlock(socket, block, blockIndex++, delayMs);
  }

  // Send message_delta and message_stop
  sendEvent(socket, {
    type: ClaudeEventTypes.MESSAGE_DELTA,
    delta: { stop_reason: "end_turn", stop_sequence: null },
    usage: { output_tokens: 150 },
  });

  await delay(delayMs);

  sendEvent(socket, {
    type: ClaudeEventTypes.MESSAGE_STOP,
  });

  // If scenario requires tool approval, send tool:permission after stream ends
  if (scenario.requestToolApproval && scenario.toolPermissions) {
    await delay(100);
    socket.send(JSON.stringify({
      type: WebSocketMessageTypes.TOOL_PERMISSION,
      toolPermissions: scenario.toolPermissions,
    }));
    console.log("🔧 Sent tool approval request");
  }

  // If scenario includes a UI request, send it after the stream
  if (scenario.uiRequest) {
    await delay(50);
    socket.send(JSON.stringify({
      type: WebSocketMessageTypes.UI_REQUEST,
      payload: {
        toolUseId: scenario.uiRequest.toolUseId,
        payload: scenario.uiRequest.payload,
      },
    }));
    console.log("📝 Sent ui:request form to client");
  }
}

async function streamContentBlock(
  socket: WebSocket,
  block: ContentBlock,
  index: number,
  delayMs: number,
) {
  const blockId = `block_${index}_${Date.now()}`;

  if (block.type === ContentBlockTypes.TEXT) {
    await streamTextBlock(socket, block.content, blockId, index, delayMs);
  } else if (block.type === ContentBlockTypes.THINKING) {
    await streamThinkingBlock(socket, block.content, blockId, index, delayMs);
  } else if (block.type === ClaudeContentTypes.TOOL_USE) {
    await streamToolUseBlock(
      socket,
      block.toolName!,
      block.toolInput!,
      blockId,
      index,
      delayMs,
    );
  }
}

async function streamTextBlock(
  socket: WebSocket,
  text: string,
  blockId: string,
  index: number,
  delayMs: number,
) {
  // content_block_start
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_START,
    index,
    content_block: {
      type: ContentBlockTypes.TEXT,
      text: "",
    },
  });

  await delay(delayMs);

  // Stream text in chunks (split by words for realistic streaming)
  const words = text.split(" ");
  let accumulated = "";

  for (let i = 0; i < words.length; i++) {
    const word = words[i];
    accumulated += (i > 0 ? " " : "") + word;

    sendEvent(socket, {
      type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
      index,
      delta: {
        type: "text_delta",
        text: (i > 0 ? " " : "") + word,
      },
    });

    await delay(delayMs);
  }

  // content_block_stop
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
    index,
  });

  await delay(delayMs);
}

async function streamThinkingBlock(
  socket: WebSocket,
  thinking: string,
  blockId: string,
  index: number,
  delayMs: number,
) {
  // content_block_start
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_START,
    index,
    content_block: {
      type: ContentBlockTypes.THINKING,
      thinking: "",
    },
  });

  await delay(delayMs);

  // Stream thinking in chunks
  const words = thinking.split(" ");

  for (let i = 0; i < words.length; i++) {
    const word = words[i];

    sendEvent(socket, {
      type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
      index,
      delta: {
        type: "thinking_delta",
        thinking: (i > 0 ? " " : "") + word,
      },
    });

    await delay(delayMs);
  }

  // content_block_stop
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
    index,
  });

  await delay(delayMs);
}

async function streamToolUseBlock(
  socket: WebSocket,
  toolName: string,
  toolInput: Record<string, unknown>,
  blockId: string,
  index: number,
  delayMs: number,
) {
  const toolUseId = `toolu_${Date.now()}_${index}`;

  // content_block_start
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_START,
    index,
    content_block: {
      type: ClaudeContentTypes.TOOL_USE,
      id: toolUseId,
      name: toolName,
      input: {},
    },
  });

  await delay(delayMs);

  // Stream tool input JSON in chunks
  const inputJson = JSON.stringify(toolInput, null, 2);
  const chunks = inputJson.match(/.{1,20}/g) || [inputJson]; // Split into ~20 char chunks

  for (const chunk of chunks) {
    sendEvent(socket, {
      type: ClaudeEventTypes.CONTENT_BLOCK_DELTA,
      index,
      delta: {
        type: "input_json_delta",
        partial_json: chunk,
      },
    });

    await delay(delayMs / 2); // Faster for JSON streaming
  }

  // content_block_stop
  sendEvent(socket, {
    type: ClaudeEventTypes.CONTENT_BLOCK_STOP,
    index,
  });

  await delay(delayMs);
}

function sendEvent(socket: WebSocket, event: Record<string, unknown>) {
  const message = {
    type: "stream:delta",
    payload: event,
  };
  socket.send(JSON.stringify(message));
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
