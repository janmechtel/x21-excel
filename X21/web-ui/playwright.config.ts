import { defineConfig, devices } from "@playwright/test";
import dotenv from "dotenv";

// Load environment variables from .env.local (worktree-specific)
dotenv.config({ path: ".env.local" });

/**
 * Playwright configuration for X21 Web-UI testing
 * See https://playwright.dev/docs/test-configuration
 *
 * Environment variables:
 * - BASE_URL: Override the base URL (default: http://localhost:5174)
 *   Set automatically by dev-mock.js in .env.local
 */
export default defineConfig({
  testDir: "./tests",

  // Maximum time one test can run for
  timeout: 30 * 1000,

  // Test execution settings
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,

  // Reporter to use
  reporter: [["html", { outputFolder: "playwright-report" }], ["list"]],

  use: {
    // Base URL for the tests - can be overridden via BASE_URL env var
    baseURL: process.env.BASE_URL || "http://localhost:5174",

    // Viewport size (simulating Excel task pane)
    viewport: { width: 300, height: 700 },

    // Collect trace when retrying the failed test
    trace: "on-first-retry",

    // Screenshot on failure
    screenshot: "on",

    // Video on failure
    video: "retain-on-failure",
  },

  // Configure projects for major browsers
  projects: [
    {
      name: "chromium",
      use: {
        ...devices["Desktop Chrome"],
        viewport: { width: 300, height: 700 },
      },
    },
    // Uncomment to test on more browsers
    // {
    //   name: 'firefox',
    //   use: { ...devices['Desktop Firefox'] },
    // },
    // {
    //   name: 'webkit',
    //   use: { ...devices['Desktop Safari'] },
    // },
  ],
  // Run web server before tests (optional)
  // Uncomment if you want Playwright to start the dev server automatically
  // webServer: {
  //   command: 'npm run dev',
  //   url: 'http://localhost:5173',
  //   reuseExistingServer: !process.env.CI,
  //   timeout: 120 * 1000,
  // },
});
