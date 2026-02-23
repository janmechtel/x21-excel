import { join } from "@std/path";

/**
 * Gets the current environment name from X21_ENVIRONMENT variable
 * Defaults to "Dev" if not set
 */
export function getEnvironment(): string {
  const env = Deno.env.get("X21_ENVIRONMENT");
  console.log(
    `DEBUG: X21_ENVIRONMENT raw value: "${env}" (length: ${env?.length ?? 0})`,
  );
  return env ? env.trim() : "Debug";
}

/**
 * Gets whether logs should be enabled for the frontend.
 * Defaults to true only when POSTHOG_API_KEY is configured.
 * Can be overridden via X21_LOGS_ENABLED or POSTHOG_LOGS_ENABLED.
 */
export function getLogsEnabled(): boolean {
  const explicit = Deno.env.get("X21_LOGS_ENABLED") ??
    Deno.env.get("POSTHOG_LOGS_ENABLED");
  if (explicit) {
    return !["0", "false", "no"].includes(explicit.trim().toLowerCase());
  }
  const posthogKey = Deno.env.get("POSTHOG_API_KEY") ||
    Deno.env.get("POSTHOG_PROJECT_API_KEY");
  return !!posthogKey;
}

/**
 * Gets the base data directory for the current OS
 */
function getBaseDataDirectory(): string {
  const os = Deno.build.os;
  const home = Deno.env.get("HOME") ||
    Deno.env.get("USERPROFILE") ||
    Deno.env.get("HOMEPATH") ||
    ".";

  if (os === "windows") {
    return (
      Deno.env.get("LOCALAPPDATA") ||
      Deno.env.get("APPDATA") ||
      join(home, "AppData", "Local")
    );
  }

  if (os === "darwin") {
    return join(home, "Library", "Application Support");
  }

  // Linux and others
  return Deno.env.get("XDG_DATA_HOME") || join(home, ".local", "share");
}

/**
 * Gets the environment-specific log directory path for Deno backend
 * Returns: %LOCALAPPDATA%\X21\X21-deno-{Environment}\Logs (Windows)
 *          ~/Library/Application Support/X21/X21-deno-{Environment}/Logs (macOS)
 *          ~/.local/share/X21/X21-deno-{Environment}/Logs (Linux)
 */
export function getLogDirectory(): string {
  const environment = getEnvironment();
  const baseDir = getBaseDataDirectory();
  const os = Deno.build.os;

  if (os === "windows") {
    return `${baseDir}\\X21\\X21-deno-${environment}\\Logs`;
  }

  return join(baseDir, "X21", `X21-deno-${environment}`, "Logs");
}

/**
 * Gets the environment-specific database file path
 * Returns: %LOCALAPPDATA%\X21\x21-{Environment}.sqlite3 (Windows)
 *          ~/Library/Application Support/X21/x21-{Environment}.sqlite3 (macOS)
 *          ~/.local/share/X21/x21-{Environment}.sqlite3 (Linux)
 */
export function getDatabasePath(): string {
  // Allow override via env variable
  const fromEnv = Deno.env.get("X21_DB_PATH");
  if (fromEnv && fromEnv.trim().length > 0) {
    return fromEnv;
  }

  const environment = getEnvironment();
  const baseDir = getBaseDataDirectory();
  const dbFileName = `x21-${environment}.sqlite3`;

  return join(baseDir, "X21", dbFileName);
}

/**
 * Gets the X21 data directory
 * Returns: %LOCALAPPDATA%\X21 (Windows)
 *          ~/Library/Application Support/X21 (macOS)
 *          ~/.local/share/X21 (Linux)
 */
export function getDataDirectory(): string {
  const baseDir = getBaseDataDirectory();
  return join(baseDir, "X21");
}
