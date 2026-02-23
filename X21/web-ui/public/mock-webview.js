/**
 * Mock WebView2 Bridge for Browser Testing
 *
 * Simulates the window.chrome.webview API that's normally provided by WebView2
 * This allows the web-ui to run in a regular browser for testing
 */

(function () {
  console.log("🧪 Loading Mock WebView2 Bridge...");

  // Check if already in WebView2 environment
  if (window.chrome?.webview) {
    console.log("✅ Real WebView2 detected, skipping mock");
    return;
  }

  // Store message handlers to send events
  let messageHandlers = [];
  let currentSelectedRange = "A1"; // Default mock selection

  // Function to send selectionChanged event
  function sendSelectionChangedEvent(range) {
    const event = {
      data: JSON.stringify({
        type: "event",
        eventType: "selectionChanged",
        data: {
          selectedRange: range,
        },
      }),
    };

    // Send to all registered message handlers
    messageHandlers.forEach((handler) => {
      try {
        handler(event);
      } catch (e) {
        console.error(
          "[Mock WebView2] Error sending selectionChanged event:",
          e,
        );
      }
    });

    console.log(`[Mock WebView2] Sent selectionChanged event: ${range}`);
  }

  // Expose function to change selection (for testing/debugging)
  window.__mockSetSelection = function (range) {
    currentSelectedRange = range;
    sendSelectionChangedEvent(range);
  };

  // Mock WebView2 API
  window.chrome = window.chrome || {};
  window.chrome.webview = {
    postMessage: function (message) {
      console.log("[Mock WebView2] postMessage:", message);

      // Parse the message
      let parsed;
      try {
        parsed = typeof message === "string" ? JSON.parse(message) : message;
      } catch (e) {
        console.error("[Mock WebView2] Failed to parse message:", e);
        return;
      }

      // Handle FrontEndIsReady - send current selection (like Excel does)
      if (parsed.type === "FrontEndIsReady") {
        setTimeout(() => {
          sendSelectionChangedEvent(currentSelectedRange);
        }, 200);
      }

      // Simulate responses for different request types
      setTimeout(() => {
        const response = getMockResponse(parsed);
        if (response) {
          const event = new MessageEvent("message", {
            data: JSON.stringify(response),
          });
          window.dispatchEvent(event);
        }
      }, 100); // Simulate small network delay
    },

    addEventListener: function (type, handler) {
      console.log("[Mock WebView2] addEventListener:", type);

      // Listen for messages from C# (simulated)
      if (type === "message") {
        window.addEventListener("message", handler);
        // Store handler for sending events
        if (!messageHandlers.includes(handler)) {
          messageHandlers.push(handler);
        }
      }

      // Simulate userIdReady event after a short delay
      setTimeout(() => {
        const userIdEvent = {
          data: JSON.stringify({
            type: "event",
            eventType: "userIdReady",
            data: {
              userId: "mock-user-" + Date.now(),
              timestamp: Date.now(),
            },
          }),
        };
        handler(userIdEvent);
        console.log("[Mock WebView2] Sent userIdReady event");
      }, 500);
    },

    removeEventListener: function (type, handler) {
      console.log("[Mock WebView2] removeEventListener:", type);
      if (type === "message") {
        window.removeEventListener("message", handler);
        // Remove handler from stored list
        messageHandlers = messageHandlers.filter((h) => h !== handler);
      }
    },
  };

  console.log("✅ Mock WebView2 Bridge loaded successfully");
})();

/**
 * Generate mock responses based on request type
 */
function getMockResponse(request) {
  const { type, payload } = request;
  const requestId = payload?.requestId;

  console.log(`[Mock WebView2] Handling request: ${type}`);

  const baseResponse = {
    type: "response",
    payload: {
      requestId,
    },
  };

  switch (type) {
    case "getWorkbookName":
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: "MockWorkbook.xlsx",
        },
      };

    case "getWebSocketUrl":
      // Point to our mock server (with /ws path)
      // Read port from window global (injected by Vite plugin) or fallback to default port 8085
      const mockServerPort = window.__MOCK_SERVER_PORT__ || 8085;
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: `ws://localhost:${mockServerPort}/ws`,
        },
      };

    case "getWorkbookPath":
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: "C:\\Mock\\Path\\MockWorkbook.xlsx",
        },
      };

    case "pickFolder":
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: {
            path: "C:\\Mock\\MergeFolder",
            files: ["January.xlsx", "February.xlsx", "March.xlsx"],
          },
        },
      };

    case "getSlashCommandsFromSheet":
      // Return mock slash commands
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: {
            success: true,
            commands: [
              {
                id: "mock-cmd-1",
                trigger: "/test",
                prompt: "This is a mock slash command for testing",
                category: "Testing",
              },
              {
                id: "mock-cmd-2",
                trigger: "/demo",
                prompt: "Demo command to showcase functionality",
                category: "Demo",
              },
            ],
          },
        },
      };

    case "getUserInfo":
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: {
            userId: "mock-user-" + Date.now(),
            email: "test@example.com",
            displayName: "Mock User",
          },
        },
      };

    default:
      console.warn(`[Mock WebView2] Unknown request type: ${type}`);
      return {
        ...baseResponse,
        payload: {
          ...baseResponse.payload,
          data: null,
          error: "Unknown request type",
        },
      };
  }
}
