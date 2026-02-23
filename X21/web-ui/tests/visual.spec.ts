/**
 * Visual/Screenshot Tests for X21 Web-UI
 *
 * Captures screenshots of different UI states for visual regression testing
 */

import { expect, test } from "@playwright/test";
import {
  bypassAuthentication,
  isToolApprovalVisible,
  sendChatMessage,
  waitForAppReady,
  waitForAssistantResponse,
  waitForStreamingComplete,
} from "./helpers/test-actions";

test.describe("X21 Web-UI Visual Tests", () => {
  test.beforeEach(async ({ page }) => {
    // Fail test on console errors
    page.on("console", (msg) => {
      if (msg.type() === "error") {
        throw new Error(`Console error: ${msg.text()}`);
      }
    });

    // Fail test on page errors
    page.on("pageerror", (error) => {
      throw new Error(`Page error: ${error.message}`);
    });

    // Bypass authentication for testing
    await bypassAuthentication(page);

    await page.goto("/");
    await waitForAppReady(page);
  });

  test("empty state - initial load", async ({ page }) => {
    // Capture the initial empty state
    await page.screenshot({
      path: "tests/screenshots/01-empty-state.png",
      fullPage: true,
    });

    // Verify chat input is visible
    const chatInput = page.locator('textarea, input[type="text"]').first();
    await expect(chatInput).toBeVisible();
  });

  test("user message sent", async ({ page }) => {
    await sendChatMessage(page, "Hello, can you help me with Excel?");

    // Wait a moment for message to appear
    await page.waitForTimeout(500);

    await page.screenshot({
      path: "tests/screenshots/02-user-message.png",
      fullPage: true,
    });
  });

  test("assistant streaming response", async ({ page }) => {
    await sendChatMessage(page, "Tell me about Excel functions");

    // Wait for response to start
    await page.waitForTimeout(1000);

    // Capture during streaming (if still streaming)
    await page.screenshot({
      path: "tests/screenshots/03-streaming-response.png",
      fullPage: true,
    });
  });

  test("complete conversation", async ({ page }) => {
    // Send message and wait for complete response
    await sendChatMessage(page, "Simple question");
    await waitForAssistantResponse(page);
    await waitForStreamingComplete(page);

    await page.screenshot({
      path: "tests/screenshots/04-complete-conversation.png",
      fullPage: true,
    });
  });

  test("tool approval UI", async ({ page }) => {
    // Send message that triggers tool use
    await sendChatMessage(page, "Please format the cells in range A1:B5");

    // Wait for tool approval UI
    await page.waitForTimeout(2000);

    // Check if tool approval is visible
    const hasToolApproval = await isToolApprovalVisible(page);

    if (hasToolApproval) {
      await page.screenshot({
        path: "tests/screenshots/05-tool-approval.png",
        fullPage: true,
      });
    } else {
      console.log("Tool approval UI not visible (may have auto-approved)");
    }
  });

  test("multiple messages in conversation", async ({ page }) => {
    // Send multiple messages to create a longer conversation
    await sendChatMessage(page, "First message");
    await waitForStreamingComplete(page);
    await page.waitForTimeout(500);

    await sendChatMessage(page, "Second message");
    await waitForStreamingComplete(page);
    await page.waitForTimeout(500);

    await sendChatMessage(page, "Third message");
    await waitForStreamingComplete(page);

    await page.screenshot({
      path: "tests/screenshots/06-multiple-messages.png",
      fullPage: true,
    });
  });

  test("scrolled conversation", async ({ page }) => {
    // Create a long conversation to test scrolling
    for (let i = 1; i <= 5; i++) {
      await sendChatMessage(page, `Message number ${i}`);
      await page.waitForTimeout(1500);
    }

    await page.screenshot({
      path: "tests/screenshots/07-scrolled-conversation.png",
      fullPage: true,
    });
  });

  test("responsive layout - narrow viewport", async ({ page }) => {
    // Test mobile/narrow viewport
    await page.setViewportSize({ width: 375, height: 667 });
    await page.reload();
    await waitForAppReady(page);

    await sendChatMessage(page, "Testing mobile view");
    await page.waitForTimeout(1000);

    await page.screenshot({
      path: "tests/screenshots/08-mobile-viewport.png",
      fullPage: true,
    });
  });

  test("dark mode (if available)", async ({ page }) => {
    // Attempt to find and toggle dark mode
    const darkModeToggle = page
      .locator('button[aria-label*="dark"], button[aria-label*="theme"]')
      .first();

    if (await darkModeToggle.isVisible({ timeout: 1000 }).catch(() => false)) {
      await darkModeToggle.click();
      await page.waitForTimeout(500);

      await sendChatMessage(page, "Testing dark mode");
      await page.waitForTimeout(1000);

      await page.screenshot({
        path: "tests/screenshots/09-dark-mode.png",
        fullPage: true,
      });
    } else {
      console.log("Dark mode toggle not found");
    }
  });

  test("settings panel (if available)", async ({ page }) => {
    // Try to open settings
    const settingsButton = page
      .locator('button[aria-label*="Settings"], button:has-text("Settings")')
      .first();

    if (await settingsButton.isVisible({ timeout: 1000 }).catch(() => false)) {
      await settingsButton.click();
      await page.waitForTimeout(500);

      await page.screenshot({
        path: "tests/screenshots/10-settings-panel.png",
        fullPage: true,
      });
    } else {
      console.log("Settings button not found");
    }
  });
});

test.describe("Visual Regression - Edge Cases", () => {
  test.beforeEach(async ({ page }) => {
    // Fail test on console errors
    page.on("console", (msg) => {
      if (msg.type() === "error") {
        throw new Error(`Console error: ${msg.text()}`);
      }
    });

    // Fail test on page errors
    page.on("pageerror", (error) => {
      throw new Error(`Page error: ${error.message}`);
    });

    // Bypass authentication for testing
    await bypassAuthentication(page);

    await page.goto("/");
    await waitForAppReady(page);
  });

  test("long message text", async ({ page }) => {
    const longMessage =
      "This is a very long message that should test how the UI handles extensive text input. ".repeat(
        10,
      );

    await sendChatMessage(page, longMessage);
    await page.waitForTimeout(1000);

    await page.screenshot({
      path: "tests/screenshots/edge-01-long-message.png",
      fullPage: true,
    });
  });

  test("special characters in message", async ({ page }) => {
    await sendChatMessage(
      page,
      'Testing special chars: @#$%^&*() <script>alert("test")</script> 你好',
    );
    await page.waitForTimeout(1000);

    await page.screenshot({
      path: "tests/screenshots/edge-02-special-chars.png",
      fullPage: true,
    });
  });

  test("rapid message sending", async ({ page }) => {
    // Test UI under rapid interaction
    await sendChatMessage(page, "First");
    await page.waitForTimeout(100);
    await sendChatMessage(page, "Second");
    await page.waitForTimeout(100);
    await sendChatMessage(page, "Third");

    await page.waitForTimeout(1000);

    await page.screenshot({
      path: "tests/screenshots/edge-03-rapid-messages.png",
      fullPage: true,
    });
  });
});
