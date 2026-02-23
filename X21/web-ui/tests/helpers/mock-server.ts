/**
 * Mock Server Management for Playwright Tests
 *
 * Utilities to start/stop the mock WebSocket server during tests
 */

interface MockServerProcess {
  pid?: number;
  port: number;
}

let mockServer: MockServerProcess | null = null;

/**
 * Start the mock server for testing
 */
export async function startMockServer(): Promise<MockServerProcess> {
  if (mockServer) {
    console.log("Mock server already running");
    return mockServer;
  }

  console.log("Starting mock server...");

  // Note: In a real implementation, you would start the Deno server as a background process
  // For now, this assumes the mock server is already running or will be started manually
  // You can enhance this to actually spawn the process:
  //
  // const serverProcess = spawn('deno', ['run', '--allow-net', '../mock-server/mod.ts'], {
  //   detached: true,
  //   stdio: 'ignore'
  // });
  //
  // serverProcess.unref();

  mockServer = {
    port: 8085,
  };

  // Wait for server to be ready
  await waitForServer("http://localhost:8085/health", 10000);

  console.log("Mock server ready");
  return mockServer;
}

/**
 * Stop the mock server
 */
export async function stopMockServer(): Promise<void> {
  if (!mockServer) {
    return;
  }

  console.log("Stopping mock server...");

  // If you spawned the process, kill it here
  // if (mockServer.pid) {
  //   process.kill(mockServer.pid);
  // }

  mockServer = null;
  console.log("Mock server stopped");
}

/**
 * Wait for server to be ready by polling the health endpoint
 */
async function waitForServer(url: string, timeoutMs: number): Promise<void> {
  const startTime = Date.now();

  while (Date.now() - startTime < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) {
        return;
      }
    } catch (error) {
      // Server not ready yet, continue polling
    }

    await new Promise((resolve) => setTimeout(resolve, 500));
  }

  throw new Error(
    `Server at ${url} did not become ready within ${timeoutMs}ms`,
  );
}

/**
 * Check if mock server is running
 */
export async function isMockServerRunning(): Promise<boolean> {
  try {
    const response = await fetch("http://localhost:8085/health", {
      signal: AbortSignal.timeout(2000),
    });
    return response.ok;
  } catch {
    return false;
  }
}
