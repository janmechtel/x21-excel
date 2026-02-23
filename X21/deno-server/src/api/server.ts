import { Router } from "../router/index.ts";
import { createLogger } from "../utils/logger.ts";
import { findAvailablePort, writePortToFile } from "../utils/port.ts";
import { initDatabase } from "../db/sqlite.ts";

export class ApiServe {
  private router: Router;
  private logger = createLogger("ApiServe");

  constructor() {
    this.router = new Router();
  }

  async handleRequest(req: Request): Promise<Response> {
    return await this.router.handleRequest(req);
  }

  async start(preferredPort: number): Promise<void> {
    // Initialize SQLite database (idempotent)
    try {
      initDatabase();
      this.logger.info("SQLite database initialized");
    } catch (e) {
      this.logger.error("Failed to initialize SQLite database", e);
      throw e;
    }

    const actualPort = await findAvailablePort(preferredPort);

    await writePortToFile(actualPort);
    // Back-compat: write the same port for "websocket" so clients reading that file work
    await writePortToFile(actualPort, "websocket");

    (globalThis as unknown as { Deno: { serve: typeof Deno.serve } }).Deno
      .serve({
        port: actualPort,
        hostname: "127.0.0.1",
      }, (req: Request) => this.handleRequest(req));

    this.logger.info(`HTTP server running on http://localhost:${actualPort}`);
    this.logger.info(
      `WebSocket endpoint available at ws://localhost:${actualPort}/ws`,
    );
    if (actualPort !== preferredPort) {
      this.logger.info(
        `Note: Preferred port ${preferredPort} was not available, using port ${actualPort}`,
      );
    }
  }
}
