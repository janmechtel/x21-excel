import { ApiServe } from "./src/api/server.ts";
import { createLogger, setupLogger } from "./src/utils/logger.ts";
import { getLLMConfig } from "./src/llm-client/provider.ts";
import { initDatabase } from "./src/db/sqlite.ts";

// Global error handlers to prevent crashes
let logger: ReturnType<typeof createLogger>;

// Handle uncaught errors
globalThis.addEventListener("error", (event) => {
  if (logger) {
    logger.error("Uncaught error", {
      message: event.error?.message,
      stack: event.error?.stack,
      error: event.error,
    });
  } else {
    console.error("Uncaught error (logger not initialized):", event.error);
  }
  // Prevent the error from crashing the process
  event.preventDefault();
});

// Handle unhandled promise rejections
globalThis.addEventListener("unhandledrejection", (event) => {
  if (logger) {
    logger.error("Unhandled promise rejection", {
      reason: event.reason,
      promise: event.promise,
    });
  } else {
    console.error(
      "Unhandled promise rejection (logger not initialized):",
      event.reason,
    );
  }
  // Prevent the rejection from crashing the process
  event.preventDefault();
});

async function runAPI() {
  // Setup logger first before creating any loggers
  await setupLogger();

  logger = createLogger("Main");
  logger.info(`LANGFUSE_ENABLED: ${Deno.env.get("LANGFUSE_ENABLED")}`);

  // Initialize SQLite database before checking LLM config
  try {
    initDatabase();
    logger.info("SQLite database initialized");
  } catch (e) {
    logger.error("Failed to initialize SQLite database", e);
    throw e;
  }

  // Log LLM configuration
  const llmConfig = getLLMConfig();
  logger.info("LLM Configuration:", {
    provider: llmConfig.provider,
    model: llmConfig.model,
    apiKey: llmConfig.apiKey,
    endpoint: llmConfig.endpoint,
  });

  const server = new ApiServe();
  const portString = Deno.env.get("SERVER_PORT");
  const port = portString ? Number(portString) : 8000; // Default to 8000 if not set

  logger.info(`Using port: ${port}`);

  try {
    await server.start(port);
  } catch (error) {
    logger.error("Fatal error starting server", error);
    // Attempt to restart after a delay
    logger.info("Attempting to restart server in 5 seconds...");
    await new Promise((resolve) => setTimeout(resolve, 5000));
    return runAPI(); // Recursive restart
  }
}

async function main() {
  await runAPI();
}

if ((import.meta as any).main) {
  main();
}
