/**
 * Showcase Test - UI Request Questions Form
 *
 * This test triggers the showcase command and captures a screenshot
 * of the questions form for visual testing and documentation.
 */

import { expect, test } from "@playwright/test";
import {
  bypassAuthentication,
  sendChatMessage,
  waitForAppReady,
} from "./helpers/test-actions";

test.describe("UI Request Questions Showcase", () => {
  test.beforeEach(async ({ page }) => {
    // Bypass authentication for testing
    await bypassAuthentication(page);

    // Navigate to the app
    await page.goto("/");
    await waitForAppReady(page);
  });

  test("displays questions form showcase", async ({ page }) => {
    // Send the showcase command (sendChatMessage has internal waits)
    await page.waitForTimeout(2000); // Give app extra time to initialize
    await sendChatMessage(page, "/ui_questions_showcase");

    // Wait for the UI request form to appear
    const uiRequestCard = page.locator('[data-testid="ui-request-card"]');
    await expect(uiRequestCard).toBeVisible({ timeout: 15000 });

    // Find the chat container and scroll it to the top
    const chatContainer = page.locator(".overflow-y-auto").first();
    await chatContainer.evaluate((el) => (el.scrollTop = 0));
    await page.waitForTimeout(500);

    // Capture initial form (fullPage to see everything)
    await page.screenshot({
      path: "tests/screenshots/showcase-form-step1.png",
      fullPage: true,
    });

    // Also capture just the form card to see details better
    await uiRequestCard.screenshot({
      path: "tests/screenshots/showcase-form-card.png",
    });

    // Click Continue to advance through the showcase
    const continueButton = page.getByRole("button", { name: /continue/i });
    if (await continueButton.isVisible({ timeout: 2000 }).catch(() => false)) {
      await continueButton.click();
      await page.waitForTimeout(2000);

      // Capture the next form
      await page.screenshot({
        path: "tests/screenshots/showcase-form-step2.png",
        fullPage: true,
      });
    }

    // Final screenshot
    await page.screenshot({
      path: "tests/screenshots/showcase-questions-form.png",
      fullPage: true,
    });
  });
});
