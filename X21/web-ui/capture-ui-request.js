/**
 * Paste this in the browser console to capture UI requests
 *
 * Usage:
 * 1. Open DevTools (F12)
 * 2. Paste this code in the Console tab
 * 3. Trigger your prompt that creates a UI request
 * 4. The payload will be logged to the console
 * 5. Copy the payload and use it to create a new scenario
 */

(function () {
  console.log("🎣 UI Request Capture Tool Installed");
  console.log("   Waiting for UI requests...");

  // Store the original send method
  const svc = window.__x21WebSocketService;
  if (!svc) {
    console.error(
      "❌ WebSocket service not found. Make sure the app is loaded.",
    );
    return;
  }

  // Hook into the WebSocket message handler
  const originalOnMessage = svc.ws?.onmessage;
  if (svc.ws && originalOnMessage) {
    svc.ws.onmessage = function (event) {
      try {
        const data = JSON.parse(event.data);

        // Check if this is a UI request
        if (data.type === "ui:request") {
          console.log("📋 UI Request Captured!");
          console.log("=".repeat(80));
          console.log("Tool Use ID:", data.toolUseId);
          console.log("\nPayload:", JSON.stringify(data.payload, null, 2));
          console.log("=".repeat(80));
          console.log("\n💾 Copy the payload above to create a new scenario\n");

          // Store for easy access
          window.__lastUiRequest = data;
          console.log("💡 Tip: Access with window.__lastUiRequest");
        }
      } catch (e) {
        // Not JSON or parsing failed, ignore
      }

      // Call the original handler
      return originalOnMessage.call(this, event);
    };

    console.log("✅ Capture tool ready!");
  } else {
    console.error(
      "❌ Could not hook into WebSocket. Connection might not be established yet.",
    );
  }
})();
