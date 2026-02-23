import { expect, test } from "@playwright/test";
import {
  bypassAuthentication,
  sendChatMessage,
  waitForAppReady,
} from "../helpers/test-actions";

/**
 * TDD Loop Test - Auto-generated
 *
 * This test is automatically generated and modified by the /tdd-loop command.
 * Goals:
 * 1. Fix "No, pause" button label overflow
 * 2. Remove outlines from the form
 * 3. Scroll the first question into view when form appears
 *
 * DO NOT manually edit this file - it will be overwritten by the TDD loop.
 */

test("TDD iteration", async ({ page }) => {
  // Bypass authentication
  await bypassAuthentication(page);

  // Navigate to the application
  await page.goto("/");

  // Wait for app to be ready
  await waitForAppReady(page);

  // Send the showcase command
  await page.waitForTimeout(2000);
  await sendChatMessage(page, "/ui_questions_showcase");

  // Wait for the UI request form to appear
  const uiRequestCard = page.locator('[data-testid="ui-request-card"]');
  await expect(uiRequestCard).toBeVisible({ timeout: 15000 });

  // Wait for scroll animation and other animations to complete
  await page.waitForTimeout(2000);

  const iterationNumber = process.env.ITERATION || "1";

  // Click on the first answer to test checkmark on expanded question
  const yesButton = page.getByRole("button", { name: /yes, continue/i });
  if (await yesButton.isVisible({ timeout: 2000 }).catch(() => false)) {
    await yesButton.click();

    // Wait for state updates
    await page.waitForTimeout(500);
  }

  // Capture screenshot to verify checkmark shows on expanded answered question
  await page.screenshot({
    path: `tests/tdd/screenshots/iteration-${iterationNumber}.png`,
    fullPage: true,
  });
});
