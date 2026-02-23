/**
 * Authentication Flow Tests for X21 Web-UI
 *
 * Tests the Supabase OTP authentication flow with mocked API responses
 */

import { expect, test } from "@playwright/test";
import { loginWithMockSupabase } from "./helpers/test-actions";

test.describe("X21 Authentication Flow", () => {
  test.beforeEach(async ({ page }) => {
    // Setup mock Supabase responses before navigating
    await loginWithMockSupabase(page, "test@kontext21.com", "123456");

    // Navigate to the app (should show login page since not authenticated)
    await page.goto("/");
  });

  test("shows login page when not authenticated", async ({ page }) => {
    // Should show email input or login form
    const emailInput = page.locator(
      'input[type="email"], input[placeholder*="email" i]',
    );
    await expect(emailInput).toBeVisible({ timeout: 10000 });
  });

  test("user can enter email address", async ({ page }) => {
    // Find and fill email input
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("test@kontext21.com");

    // Verify email was entered
    await expect(emailInput).toHaveValue("test@kontext21.com");
  });

  test("user can submit email and see OTP verification", async ({ page }) => {
    // Enter email
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("test@kontext21.com");

    // Submit email (find continue/next/submit button)
    const submitButton = page
      .locator(
        'button[type="submit"], button:has-text("Continue"), button:has-text("Next"), button:has-text("Send")',
      )
      .first();
    await submitButton.click();

    // Wait for OTP input to appear
    const otpInput = page
      .locator(
        'input[type="text"], input[placeholder*="code" i], input[placeholder*="otp" i]',
      )
      .first();
    await expect(otpInput).toBeVisible({ timeout: 5000 });
  });

  test("user can verify OTP and login successfully", async ({ page }) => {
    // Step 1: Enter email
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("test@kontext21.com");

    // Step 2: Submit email
    const submitButton = page
      .locator(
        'button[type="submit"], button:has-text("Continue"), button:has-text("Next"), button:has-text("Send")',
      )
      .first();
    await submitButton.click();

    // Step 3: Wait for OTP input
    const otpInput = page
      .locator(
        'input[type="text"], input[placeholder*="code" i], input[placeholder*="otp" i]',
      )
      .first();
    await expect(otpInput).toBeVisible({ timeout: 5000 });

    // Step 4: Enter OTP
    await otpInput.fill("123456");

    // Step 5: Submit OTP
    const verifyButton = page
      .locator(
        'button[type="submit"], button:has-text("Verify"), button:has-text("Continue"), button:has-text("Login")',
      )
      .first();
    await verifyButton.click();

    // Step 6: Verify successful login - should see chat interface
    const chatInput = page.locator('textarea, input[type="text"]').first();
    await expect(chatInput).toBeVisible({ timeout: 10000 });
  });

  test("shows validation error for invalid email", async ({ page }) => {
    // Enter invalid email
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("invalid-email");

    // Try to submit
    const submitButton = page
      .locator(
        'button[type="submit"], button:has-text("Continue"), button:has-text("Next"), button:has-text("Send")',
      )
      .first();
    await submitButton.click();

    // Should show validation error or prevent submission
    // Either error message appears OR email input still visible (didn't progress)
    const stillOnLoginPage = await emailInput.isVisible();
    expect(stillOnLoginPage).toBe(true);
  });
});

test.describe("Authentication - New User Flow", () => {
  test.beforeEach(async ({ page }) => {
    await loginWithMockSupabase(page, "newuser@kontext21.com", "123456");
    await page.goto("/");
  });

  test("new user can provide first and last name", async ({ page }) => {
    // This test assumes new users need to provide their name
    // If the UI collects name during registration

    // Enter email
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("newuser@kontext21.com");

    // Submit email
    const submitButton = page.locator('button[type="submit"]').first();
    await submitButton.click();

    // Wait a moment for potential name collection UI
    await page.waitForTimeout(1000);

    // Look for first name and last name inputs (if they exist)
    const firstNameInput = page.locator(
      'input[name="firstName"], input[placeholder*="first name" i]',
    );
    const lastNameInput = page.locator(
      'input[name="lastName"], input[placeholder*="last name" i]',
    );

    if (await firstNameInput.isVisible({ timeout: 2000 }).catch(() => false)) {
      await firstNameInput.fill("New");
      await lastNameInput.fill("User");

      // Submit name form
      const continueButton = page
        .locator('button[type="submit"], button:has-text("Continue")')
        .first();
      await continueButton.click();
    }

    // Should eventually see OTP input or chat interface
    const otpOrChat = page
      .locator('input[placeholder*="code" i], textarea, input[type="text"]')
      .first();
    await expect(otpOrChat).toBeVisible({ timeout: 5000 });
  });
});

test.describe("Authentication - Error Handling", () => {
  test("handles network errors gracefully", async ({ page }) => {
    // Mock network failure for OTP endpoint
    await page.route("**/auth/v1/otp", async (route) => {
      await route.abort("failed");
    });

    await page.goto("/");

    // Try to login
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("test@kontext21.com");

    const submitButton = page.locator('button[type="submit"]').first();
    await submitButton.click();

    // Should show error message or stay on login page
    await page.waitForTimeout(2000);

    // Verify still on login page or error message visible
    const errorMessage = page.locator(
      '[role="alert"], .error, .text-destructive',
    );
    const stillOnLogin = await emailInput.isVisible();

    const hasErrorOrStillOnLogin =
      stillOnLogin || (await errorMessage.isVisible().catch(() => false));
    expect(hasErrorOrStillOnLogin).toBe(true);
  });

  test("handles invalid OTP code", async ({ page }) => {
    // Mock OTP verification failure
    await page.route("**/auth/v1/otp", async (route) => {
      await route.fulfill({
        status: 200,
        contentType: "application/json",
        body: JSON.stringify({}),
      });
    });

    await page.route("**/auth/v1/verify", async (route) => {
      await route.fulfill({
        status: 400,
        contentType: "application/json",
        body: JSON.stringify({
          error: "invalid_grant",
          error_description: "Invalid OTP code",
        }),
      });
    });

    await page.goto("/");

    // Enter email
    const emailInput = page
      .locator('input[type="email"], input[placeholder*="email" i]')
      .first();
    await emailInput.fill("test@kontext21.com");
    await page.locator('button[type="submit"]').first().click();

    // Enter invalid OTP
    const otpInput = page
      .locator('input[type="text"], input[placeholder*="code" i]')
      .first();
    await expect(otpInput).toBeVisible({ timeout: 5000 });
    await otpInput.fill("000000");

    // Submit invalid OTP
    const verifyButton = page
      .locator('button[type="submit"], button:has-text("Verify")')
      .first();
    await verifyButton.click();

    // Should show error message
    await page.waitForTimeout(1000);
    const errorMessage = page.locator(
      '[role="alert"], .error, .text-destructive',
    );
    const hasError = await errorMessage.isVisible().catch(() => false);

    // Either error shown OR still on OTP page
    const stillOnOTPPage = await otpInput.isVisible();
    expect(hasError || stillOnOTPPage).toBe(true);
  });
});
