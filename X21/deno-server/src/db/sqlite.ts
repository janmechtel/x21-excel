import { createLogger } from "../utils/logger.ts";
import {
  getDatabasePath as getEnvDatabasePath,
  getEnvironment,
} from "../utils/environment.ts";
import { runMigrations } from "./migrations.ts";

// Use Deno SQLite. Pin a version for stability.
import { DB } from "sqlite";
import { dirname } from "@std/path";

const logger = createLogger("SQLite");

let dbInstance: DB | null = null;

function ensureDirExists(path: string): void {
  try {
    const dir = dirname(path);
    if (dir) {
      Deno.mkdirSync(dir, { recursive: true });
    }
  } catch {
    // ignore if exists
  }
}

export function getDatabasePath(): string {
  // Use centralized environment utility
  return getEnvDatabasePath();
}

export function getDb(): DB {
  if (!dbInstance) {
    throw new Error("Database not initialized. Call initDatabase() first.");
  }
  return dbInstance;
}

/**
 * Check if the database has been initialized
 */
export function isDbInitialized(): boolean {
  return dbInstance !== null;
}

export function initDatabase(): void {
  if (dbInstance) return;

  const dbPath = getDatabasePath();
  ensureDirExists(dbPath);

  // Log environment and database path for debugging
  const environment = getEnvironment();
  const envVar = Deno.env.get("X21_ENVIRONMENT");
  logger.info(`X21_ENVIRONMENT variable: ${envVar || "not set"}`);
  logger.info(`Resolved environment: ${environment}`);
  logger.info(`Opening SQLite database at ${dbPath}`);
  const db = new DB(dbPath, { memory: false });

  // Enable WAL for concurrency and performance
  try {
    db.execute("PRAGMA journal_mode = WAL;");
    db.execute("PRAGMA busy_timeout = 2000;");
  } catch {
    // Best effort
  }

  // Minimal schema supporting workbook-keyed messages
  db.execute(`
CREATE TABLE IF NOT EXISTS messages (
  id TEXT PRIMARY KEY,
  workbook_key TEXT NOT NULL,
  conversation_id TEXT NOT NULL,
  role TEXT NOT NULL,              -- 'user' | 'assistant' | 'system'
  content TEXT NOT NULL,
  created_at INTEGER NOT NULL
);
`);

  // Indexes
  db.execute(
    "CREATE INDEX IF NOT EXISTS idx_messages_conv_created ON messages(conversation_id, created_at DESC);",
  );
  db.execute(
    "CREATE INDEX IF NOT EXISTS idx_messages_workbook_created ON messages(workbook_key, created_at DESC);",
  );

  // Run database migrations
  runMigrations(db);

  dbInstance = db;
}

export function closeDatabase(): void {
  try {
    dbInstance?.close();
  } finally {
    dbInstance = null;
  }
}

export function nowMs(): number {
  return Date.now();
}
