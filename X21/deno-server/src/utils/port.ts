/// <reference lib="deno.ns" />

import { createLogger } from "./logger.ts";

const logger = createLogger("port");

/**
 * Check if a port is available by attempting to listen on it
 */
async function isPortAvailable(port: number): Promise<boolean> {
  try {
    const listener = Deno.listen({ port, hostname: "127.0.0.1" });
    await listener.close();
    return true;
  } catch {
    return false;
  }
}

/**
 * Find an available port starting from the given port
 */
export async function findAvailablePort(
  startPort: number = 8013,
): Promise<number> {
  let port = startPort;
  const maxAttempts = 100; // Prevent infinite loops

  for (let i = 0; i < maxAttempts; i++) {
    if (await isPortAvailable(port)) {
      return port;
    }
    port++;
  }

  throw new Error(`No available port found starting from ${startPort}`);
}

/**
 * Get the X21 app data directory path (same parent as logs)
 */
function getX21AppDataDirectory(): string {
  // Get the LocalApplicationData directory
  const localAppData = Deno.env.get("LOCALAPPDATA");
  if (!localAppData) {
    throw new Error("LOCALAPPDATA environment variable not found");
  }

  // Use the same path structure as the X21 app (2 levels up from logs)
  return `${localAppData}\\X21`;
}

/**
 * Get the branch name for the filename
 */
function getBranchName(): string {
  // Try to get branch from environment variable (set during build)
  const envBranch = Deno.env.get("X21_ENVIRONMENT");
  if (envBranch) {
    return envBranch;
  }
  return "dev";
}

function getBranchNameVariants(branchName: string): string[] {
  const lower = branchName.toLowerCase();
  const upper = branchName.toUpperCase();
  const capitalized = branchName.length > 0
    ? branchName[0].toUpperCase() + branchName.slice(1)
    : branchName;
  // Deduplicate while preserving order
  const variants = [branchName, capitalized, lower, upper];
  return variants.filter((value, index) => variants.indexOf(value) === index);
}

/**
 * Write the port number to a file in the X21 app data directory for other processes to read
 */
export async function writePortToFile(
  port: number,
  serverType: string = "server",
): Promise<void> {
  try {
    const appDataDir = getX21AppDataDirectory();

    // Ensure the app data directory exists
    try {
      await Deno.stat(appDataDir);
    } catch {
      await Deno.mkdir(appDataDir, { recursive: true });
    }

    const branchName = getBranchName();
    const variants = getBranchNameVariants(branchName);

    for (const variant of variants) {
      const fileName = `deno-${serverType}-port-${variant}`;
      const filePath = `${appDataDir}\\${fileName}`;
      await Deno.writeTextFile(filePath, port.toString());
      logger.info(`Port ${port} written to: ${filePath}`);
    }
  } catch (error) {
    logger.error(`Failed to write port to app data directory:`, error);

    // Fallback to current directory
    const branchName = getBranchName();
    const fallbackPath = `./deno-${serverType}-${branchName}`;
    await Deno.writeTextFile(fallbackPath, port.toString());
    logger.info(`Port ${port} written to fallback location: ${fallbackPath}`);
  }
}
