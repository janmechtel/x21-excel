/**
 * Common Test Actions
 *
 * Reusable functions for interacting with the X21 UI during tests
 */

import { Page } from "@playwright/test";

/**
 * Send a chat message
 */
export async function sendChatMessage(page: Page, message: string) {
  // The app uses Lexical editor (ContentEditable div), not a textarea
  const chatInput = page.locator('[contenteditable="true"]').first();
  await chatInput.waitFor({ state: "visible", timeout: 10000 });

  // Click to focus the editor
  await chatInput.click();

  // Type the message
  await chatInput.fill(message);

  // Press Enter to send (or click send button if visible)
  const sendButton = page
    .locator('button[type="submit"], button:has-text("Send")')
    .first();
  if (await sendButton.isVisible({ timeout: 1000 }).catch(() => false)) {
    await sendButton.click();
  } else {
    await chatInput.press("Enter");
  }
}

/**
 * Wait for assistant response to appear
 */
export async function waitForAssistantResponse(page: Page, timeout = 10000) {
  await page.waitForSelector(
    '[data-role="assistant"], .message-assistant, [role="article"]',
    {
      timeout,
    },
  );
}

/**
 * Wait for streaming to complete
 */
export async function waitForStreamingComplete(page: Page, timeout = 10000) {
  // Wait for streaming indicator to disappear
  await page
    .waitForSelector('[data-streaming="true"], .streaming-indicator', {
      state: "hidden",
      timeout,
    })
    .catch(() => {
      // Ignore if selector doesn't exist
    });

  // Additional wait for any animations to complete
  await page.waitForTimeout(500);
}

/**
 * Get all chat messages
 */
export async function getChatMessages(page: Page) {
  const messages = await page.locator("[data-role], .message").all();
  return Promise.all(
    messages.map(async (msg) => ({
      role: await msg.getAttribute("data-role"),
      text: await msg.textContent(),
    })),
  );
}

/**
 * Check if tool approval UI is visible
 */
export async function isToolApprovalVisible(page: Page): Promise<boolean> {
  const approvalUI = page
    .locator(
      '[data-testid="tool-approval"], .tool-approval, button:has-text("Approve")',
    )
    .first();
  try {
    await approvalUI.waitForSelector({ timeout: 2000 });
    return approvalUI.isVisible();
  } catch {
    return false;
  }
}

/**
 * Approve all pending tools
 */
export async function approveAllTools(page: Page) {
  const approveAllButton = page
    .locator('button:has-text("Approve All"), button:has-text("Approve")')
    .first();
  await approveAllButton.click();
}

/**
 * Reject all pending tools
 */
export async function rejectAllTools(page: Page) {
  const rejectButton = page
    .locator('button:has-text("Reject"), button:has-text("Cancel")')
    .first();
  await rejectButton.click();
}

/**
 * Open settings panel
 */
export async function openSettings(page: Page) {
  const settingsButton = page
    .locator('button[aria-label*="Settings"], button:has-text("Settings")')
    .first();
  await settingsButton.click();
}

/**
 * Close settings panel
 */
export async function closeSettings(page: Page) {
  const closeButton = page
    .locator('button[aria-label*="Close"], button:has-text("Close")')
    .first();
  await closeButton.click();
}

/**
 * Enable/disable auto-approve
 */
export async function toggleAutoApprove(page: Page, enable: boolean) {
  await openSettings(page);

  const autoApproveCheckbox = page
    .locator('input[type="checkbox"][name*="auto"], label:has-text("Auto")')
    .first();
  const isChecked = await autoApproveCheckbox.isChecked();

  if (isChecked !== enable) {
    await autoApproveCheckbox.click();
  }

  await closeSettings(page);
}

/**
 * Clear chat history
 */
export async function clearChat(page: Page) {
  const clearButton = page
    .locator('button:has-text("Clear"), button:has-text("New")')
    .first();
  await clearButton.click();

  // Confirm if there's a confirmation dialog
  const confirmButton = page.locator(
    'button:has-text("Confirm"), button:has-text("Yes")',
  );
  if (await confirmButton.isVisible({ timeout: 1000 }).catch(() => false)) {
    await confirmButton.click();
  }
}

/**
 * Take a screenshot with consistent naming
 */
export async function takeScreenshot(page: Page, name: string) {
  const timestamp = Date.now();
  await page.screenshot({
    path: `tests/screenshots/${name}-${timestamp}.png`,
    fullPage: true,
  });
}

/**
 * Wait for the app to be fully loaded
 */
export async function waitForAppReady(page: Page) {
  // Wait for React root to be rendered
  await page.waitForSelector("#root > *", { timeout: 10000 });

  // Wait for any initial loading states to complete
  await page.waitForLoadState("networkidle");

  // Additional wait for animations
  await page.waitForTimeout(500);
}

/**
 * Mock authentication (bypass login)
 * Sets a valid Supabase session in localStorage so the app thinks user is logged in
 * NOTE: Must be called BEFORE page.goto() as it uses addInitScript
 */
export async function bypassAuthentication(
  page: Page,
  email = "test@kontext21.com",
) {
  // Use addInitScript to inject auth before page loads
  await page.addInitScript((userEmail) => {
    // Set proper Supabase auth token in localStorage
    const mockSession = {
      access_token: "mock-access-token-" + Date.now(),
      refresh_token: "mock-refresh-token-" + Date.now(),
      expires_in: 3600,
      expires_at: Math.floor(Date.now() / 1000) + 3600,
      token_type: "bearer",
      user: {
        id: "mock-user-id-" + Date.now(),
        email: userEmail,
        user_metadata: {
          first_name: "Test",
          last_name: "User",
        },
        app_metadata: {},
        aud: "authenticated",
        created_at: new Date().toISOString(),
      },
    };

    // Supabase storage key format: sb-{project-ref}-auth-token
    localStorage.setItem(
      "sb-qvycnlwxhhmuobjzzoos-auth-token",
      JSON.stringify(mockSession),
    );
  }, email);
}

/**
 * Mock Supabase API responses for testing login flow
 * Intercepts Supabase auth endpoints and returns successful mock responses
 */
export async function loginWithMockSupabase(
  page: Page,
  email = "test@kontext21.com",
  otp = "123456",
) {
  // Mock the OTP send request
  await page.route("**/auth/v1/otp", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({}),
    });
  });

  // Mock the OTP verify request
  await page.route("**/auth/v1/verify", async (route) => {
    const mockSession = {
      access_token: "mock-access-token-" + Date.now(),
      refresh_token: "mock-refresh-token-" + Date.now(),
      expires_in: 3600,
      expires_at: Math.floor(Date.now() / 1000) + 3600,
      token_type: "bearer",
      user: {
        id: "mock-user-id-" + Date.now(),
        email: email,
        user_metadata: {
          first_name: "Test",
          last_name: "User",
        },
        app_metadata: {},
        aud: "authenticated",
        created_at: new Date().toISOString(),
        confirmed_at: new Date().toISOString(),
      },
    };

    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify(mockSession),
    });
  });

  // Mock the session endpoint (for checking existing sessions)
  await page.route("**/auth/v1/token**", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({
        access_token: null,
        user: null,
      }),
    });
  });

  return { email, otp };
}
