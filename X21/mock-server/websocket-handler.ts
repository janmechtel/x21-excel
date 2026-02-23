/**
 * WebSocket Message Handler
 *
 * Handles all message types from web-ui and simulates realistic backend responses
 */

import { simulateClaudeStream } from "./utils/claude-stream.ts";
import {
  getSimpleResponse,
  getToolApprovalScenario,
  getUiRequestScenario,
  getUiControlsShowcaseScenario,
} from "./scenarios/scenarios.ts";
import { WebSocketMessageTypes } from "../shared/types/index.ts";

interface WebSocketMessage {
  type: string;
  workbookName?: string;
  payload?: unknown;
  groupId?: string;
  toolId?: string;
  toolResponses?: Array<{ toolId: string; decision: string }>;
  userMessage?: string;
}

export function handleWebSocket(socket: WebSocket) {
  console.log("📱 WebSocket client connected");

  socket.onopen = () => {
    console.log("✅ WebSocket connection established");

    // Send email request (like real backend does)
    setTimeout(() => {
      socket.send(JSON.stringify({ type: "email:request" }));
    }, 500);
  };

  socket.onmessage = async (event) => {
    try {
      const message: WebSocketMessage = JSON.parse(event.data);
      console.log(`📨 Received: ${message.type}`, message);

      await handleMessage(socket, message);
    } catch (error) {
      console.error("❌ Error handling message:", error);
      socket.send(JSON.stringify({
        type: "stream:error",
        error: "PARSE_ERROR",
        message: "Failed to parse message",
      }));
    }
  };

  socket.onclose = () => {
    console.log("🔌 WebSocket client disconnected");
  };

  socket.onerror = (error) => {
    console.error("❌ WebSocket error:", error);
  };
}

async function handleMessage(socket: WebSocket, message: WebSocketMessage) {
  switch (message.type) {
    case "stream:start":
      await handleStreamStart(socket, message);
      break;

    case "stream:cancel":
      handleStreamCancel(socket, message);
      break;

    case "chat:restart":
      handleChatRestart(socket, message);
      break;

    case "tool:permission:response":
      await handleToolPermissionResponse(socket, message);
      break;

    case "tool:approve":
      handleToolApprove(socket, message);
      break;

    case "tool:reject":
      handleToolReject(socket, message);
      break;

    case "tool:view":
      handleToolView(socket, message);
      break;

    case "tool:unview":
      handleToolUnview(socket, message);
      break;

    case "tool:revert":
      handleToolRevert(socket, message);
      break;

    case "tool:apply":
      handleToolApply(socket, message);
      break;

    case "score:score":
      handleScore(socket, message);
      break;

    case "score:feedback":
      handleFeedback(socket, message);
      break;

    case "user:email_response":
      handleEmailResponse(socket, message);
      break;

    default:
      console.warn(`⚠️  Unknown message type: ${message.type}`);
  }
}

// ============================================================================
// Message Handlers
// ============================================================================

async function handleStreamStart(socket: WebSocket, message: WebSocketMessage) {
  const payload = message.payload as { prompt: string; activeTools?: string[] };
  const prompt = payload?.prompt || "";

  console.log(`🤖 Starting stream for prompt: "${prompt.substring(0, 50)}..."`);

  // Determine scenario based on prompt content
  const shouldIncludeTools = prompt.toLowerCase().includes("tool") ||
    prompt.toLowerCase().includes("format") ||
    prompt.toLowerCase().includes("write") ||
    prompt.toLowerCase().includes("read");
  const shouldRequestUiForm = prompt.toLowerCase().includes("form") ||
    prompt.toLowerCase().includes("ui request") ||
    prompt.toLowerCase().includes("amortization");
  const shouldShowcaseAllControls = prompt.toLowerCase().includes("ui demo") ||
    prompt.toLowerCase().includes("all controls") ||
    prompt.toLowerCase().includes("showcase");

  // Simulate streaming response
  try {
    if (shouldShowcaseAllControls) {
      await simulateClaudeStream(socket, getUiControlsShowcaseScenario());
    } else if (shouldRequestUiForm) {
      await simulateClaudeStream(socket, getUiRequestScenario());
    } else if (shouldIncludeTools) {
      await simulateClaudeStream(socket, getToolApprovalScenario(prompt));
    } else {
      await simulateClaudeStream(socket, getSimpleResponse(prompt));
    }

    socket.send(JSON.stringify({ type: "stream:end" }));
    console.log("✅ Stream completed");
  } catch (error) {
    console.error("❌ Stream error:", error);
    socket.send(JSON.stringify({
      type: "stream:error",
      error: "STREAM_ERROR",
      message: String(error),
    }));
  }
}

function handleStreamCancel(socket: WebSocket, message: WebSocketMessage) {
  console.log("🛑 Stream cancelled");
  socket.send(JSON.stringify({
    type: WebSocketMessageTypes.STREAM_CANCELLED,
    message: "Stream cancelled by user",
    requestId: Date.now().toString(),
  }));
}

function handleChatRestart(socket: WebSocket, message: WebSocketMessage) {
  console.log("🔄 Chat restarted");
  // No response needed - frontend just clears state
}

async function handleToolPermissionResponse(
  socket: WebSocket,
  message: WebSocketMessage,
) {
  const toolResponses = message.toolResponses || [];
  console.log(`✅ Tool permission response:`, toolResponses);

  // Simulate tools executing
  for (const response of toolResponses) {
    if (response.decision === "approved") {
      console.log(`⚙️  Executing tool: ${response.toolId}`);

      // Simulate brief execution delay
      await new Promise((resolve) => setTimeout(resolve, 200));

      // 90% success rate
      if (Math.random() > 0.1) {
        console.log(`✅ Tool ${response.toolId} executed successfully`);
      } else {
        // Simulate tool error
        socket.send(JSON.stringify({
          type: WebSocketMessageTypes.TOOL_ERROR,
          toolId: response.toolId,
          errorMessage: "Mock error: Operation failed",
          error: "EXECUTION_ERROR",
        }));
      }
    }
  }

  // Continue with follow-up response after tools execute
  await new Promise((resolve) => setTimeout(resolve, 300));
  await simulateClaudeStream(socket, {
    blocks: [
      {
        type: "text",
        content: "I've completed the requested operations. Is there anything else you'd like me to help with?",
      },
    ],
  });
  socket.send(JSON.stringify({ type: "stream:end" }));
}

function handleToolApprove(socket: WebSocket, message: WebSocketMessage) {
  console.log(`✅ Tool approved: ${message.toolId}`);
  // Individual approval handled same as batch
}

function handleToolReject(socket: WebSocket, message: WebSocketMessage) {
  console.log(`❌ Tool rejected: ${message.toolId}`, message.userMessage);
  // Tool rejection acknowledged
}

function handleToolView(socket: WebSocket, message: WebSocketMessage) {
  console.log(`👁️  Tool preview requested: ${message.toolId}`);
  // In real system, this triggers Excel preview
}

function handleToolUnview(socket: WebSocket, message: WebSocketMessage) {
  console.log(`👁️  Tool preview removed: ${message.toolId}`);
}

function handleToolRevert(socket: WebSocket, message: WebSocketMessage) {
  console.log(`↩️  Tool reverted: ${message.toolId}`);
}

function handleToolApply(socket: WebSocket, message: WebSocketMessage) {
  console.log(`✅ Tool applied: ${message.toolId}`);
}

function handleScore(socket: WebSocket, message: WebSocketMessage) {
  console.log(`📊 Feedback score received:`, message);
}

function handleFeedback(socket: WebSocket, message: WebSocketMessage) {
  console.log(`💬 Feedback comment received:`, message);
}

function handleEmailResponse(socket: WebSocket, message: WebSocketMessage) {
  console.log(`📧 User email received:`, message);
}
