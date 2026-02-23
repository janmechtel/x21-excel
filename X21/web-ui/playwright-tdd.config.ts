import { defineConfig, devices } from "@playwright/test";
import dotenv from "dotenv";

// Load environment variables from .env.local (worktree-specific)
dotenv.config({ path: ".env.local" });

/**
 * Playwright configuration for TDD Loop workflow
 *
 * This config is optimized for rapid iteration with AI evaluation:
 * - Single test execution (no parallelization)
 * - Always captures screenshots
 * - Supports multiple browser modes (headless, headed, attach)
 *
 * Environment variables:
 * - BASE_URL: Override the base URL (default: http://localhost:5174)
 *   Set automatically by dev-mock.js in .env.local
 * - BROWSER_MODE: 'headless' | 'headed' | 'attach' (default: 'headless')
 * - CDP_ENDPOINT: Chrome DevTools Protocol endpoint for 'attach' mode
 *   (default: 'http://localhost:9222')
 */

const browserMode = process.env.BROWSER_MODE || "headless";
const cdpEndpoint = process.env.CDP_ENDPOINT || "http://localhost:9222";

export default defineConfig({
  testDir: "./tests/tdd",

  // Shorter timeout for faster iteration
  timeout: 15 * 1000,

  // Single test execution (no parallelization for TDD)
  fullyParallel: false,
  workers: 1,
  retries: 0, // No retries in TDD mode

  // Minimal reporter for clean output
  reporter: [["list"]],

  use: {
    // Base URL for the tests - can be overridden via BASE_URL env var
    baseURL: process.env.BASE_URL || "http://localhost:5174",

    // Viewport size (simulating Excel task pane)
    viewport: { width: 300, height: 700 },

    // Always capture screenshots for AI evaluation
    screenshot: "on",

    // No video/trace in TDD mode (faster)
    video: "off",
    trace: "off",

    // Headed mode configuration
    headless: browserMode === "headless",

    // Slower actions for better visibility in headed mode
    ...(browserMode === "headed" && {
      launchOptions: {
        slowMo: 100, // Slow down actions by 100ms for visibility
      },
    }),
  },

  // Single browser project
  projects: [
    {
      name: "tdd-chromium",
      use: {
        ...devices["Desktop Chrome"],
        viewport: { width: 300, height: 700 },

        // Connect to existing Chrome if in 'attach' mode
        ...(browserMode === "attach" && {
          connectOptions: {
            wsEndpoint: `${cdpEndpoint}/devtools/browser`,
          },
        }),
      },
    },
  ],
  // Assume dev server is already running
  // (User should have it running before starting TDD loop)
});
