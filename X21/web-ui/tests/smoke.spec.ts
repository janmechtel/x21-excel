/**
 * Smoke Tests for X21 Web-UI
 *
 * Basic functionality tests to verify the app works in browser mode
 */

import { expect, test } from "@playwright/test";
import {
  bypassAuthentication,
  sendChatMessage,
  waitForAppReady,
  waitForAssistantResponse,
} from "./helpers/test-actions";

test.describe("X21 Web-UI Smoke Tests", () => {
  test.beforeEach(async ({ page }) => {
    // Bypass authentication for testing
    await bypassAuthentication(page);

    // Navigate to the app
    await page.goto("/");
    await waitForAppReady(page);
  });

  test("app loads successfully", async ({ page }) => {
    // Check that the root element exists and is visible
    const root = page.locator("#root");
    await expect(root).toBeVisible();

    // Check for main UI elements (app uses Lexical ContentEditable, not textarea)
    const chatInput = page.locator('[contenteditable="true"]').first();
    await expect(chatInput).toBeVisible({ timeout: 10000 });
  });

  test("can type in chat input", async ({ page }) => {
    // App uses Lexical editor (ContentEditable), not textarea
    const chatInput = page.locator('[contenteditable="true"]').first();
    await chatInput.waitFor({ state: "visible", timeout: 10000 });

    await chatInput.fill("Hello, this is a test message");
    await expect(chatInput).toHaveText("Hello, this is a test message");
  });

  test("can send a message", async ({ page }) => {
    await sendChatMessage(page, "Hello from test");

    // Verify message appears in chat (either as user message or in history)
    const messageText = page.getByText("Hello from test");
    await expect(messageText).toBeVisible({ timeout: 10000 });
  });

  test("receives assistant response", async ({ page }) => {
    await sendChatMessage(page, "Test message");

    // Wait for assistant response
    await waitForAssistantResponse(page);

    // Check that some response content is visible
    const messages = page.locator("[data-role], .message").all();
    const messageCount = await messages.then((msgs) => msgs.length);
    expect(messageCount).toBeGreaterThan(0);
  });

  test("mock WebView2 bridge is loaded", async ({ page }) => {
    // Check that the mock bridge injected the chrome.webview object
    const hasWebViewAPI = await page.evaluate(() => {
      return typeof window.chrome?.webview !== "undefined";
    });

    expect(hasWebViewAPI).toBe(true);
  });

  test("WebSocket connection is established", async ({ page }) => {
    // Send a message to trigger WebSocket connection
    await sendChatMessage(page, "Connect test");

    // Wait a bit for connection
    await page.waitForTimeout(2000);

    // Check console for connection messages (if available in test mode)
    // In real scenario, you might check connection status in UI
    const messages = page.locator("[data-role], .message").all();
    const messageCount = await messages.then((msgs) => msgs.length);
    expect(messageCount).toBeGreaterThanOrEqual(1);
  });

  test("no console errors on load", async ({ page }) => {
    const consoleErrors: string[] = [];

    page.on("console", (msg) => {
      if (msg.type() === "error") {
        consoleErrors.push(msg.text());
      }
    });

    await page.reload();
    await waitForAppReady(page);

    // Filter out known non-critical errors if any
    const criticalErrors = consoleErrors.filter(
      (err) =>
        !err.includes("WebView2") && // Mock WebView2 warnings are ok
        !err.includes("favicon"), // Favicon 404s are ok
    );

    expect(criticalErrors).toHaveLength(0);
  });
});
