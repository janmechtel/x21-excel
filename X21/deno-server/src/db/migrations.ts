import { DB } from "sqlite";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("Migrations");

export interface Migration {
  version: number;
  name: string;
  up: (db: DB) => void;
}

// Create schema_version table to track migrations
function initMigrationTracking(db: DB): void {
  db.execute(`
    CREATE TABLE IF NOT EXISTS schema_version (
      version INTEGER PRIMARY KEY,
      name TEXT NOT NULL,
      applied_at INTEGER NOT NULL
    );
  `);
}

// Get current schema version
function getCurrentVersion(db: DB): number {
  initMigrationTracking(db);

  const result = db.query<[number]>(
    "SELECT COALESCE(MAX(version), 0) as max_version FROM schema_version",
  );

  if (result.length > 0) {
    return result[0][0];
  }

  return 0;
}

// Record migration
function recordMigration(db: DB, migration: Migration): void {
  db.query(
    "INSERT INTO schema_version (version, name, applied_at) VALUES (?, ?, ?)",
    [migration.version, migration.name, Date.now()],
  );
}

// All migrations in order
const migrations: Migration[] = [
  {
    version: 1,
    name: "add_llm_keys_config_table",
    up: (db: DB) => {
      logger.info("Running migration 1: add_llm_keys_config_table");

      db.execute(`
        CREATE TABLE IF NOT EXISTS llm_keys_config (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          provider TEXT NOT NULL DEFAULT 'azure',
          azure_openai_endpoint TEXT,
          azure_openai_key TEXT,
          added_date INTEGER NOT NULL,
          modified_date INTEGER
        );
      `);

      logger.info("Created llm_keys_config table");
    },
  },
  {
    version: 2,
    name: "add_is_active_to_llm_keys_config",
    up: (db: DB) => {
      logger.info("Running migration 2: add_is_active_to_llm_keys_config");

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1;
      `);

      logger.info("Added is_active column to llm_keys_config table");
    },
  },
  {
    version: 3,
    name: "add_workbook_diffs_and_summaries_tables",
    up: (db: DB) => {
      logger.info(
        "Running migration 3: add_workbook_diffs_and_summaries_tables",
      );

      // Create workbook_diffs table (one diff per workbook change event)
      db.execute(`
        CREATE TABLE IF NOT EXISTS workbook_diffs (
          id TEXT PRIMARY KEY,
          workbook_name TEXT NOT NULL,
          unified_diff TEXT NOT NULL,
          created_at INTEGER NOT NULL
        );
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_diffs_workbook_created
        ON workbook_diffs(workbook_name, created_at DESC);
      `);

      // Create workbook_summaries table - simple structure
      // Each summary can optionally link to a diff
      db.execute(`
        CREATE TABLE IF NOT EXISTS workbook_summaries (
          id TEXT PRIMARY KEY,
          workbook_name TEXT NOT NULL,
          summary_text TEXT NOT NULL,
          created_at INTEGER NOT NULL,
          diff_id TEXT,
          FOREIGN KEY (diff_id) REFERENCES workbook_diffs(id)
        );
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_summaries_diff_id
        ON workbook_summaries(diff_id);
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_summaries_workbook_created
        ON workbook_summaries(workbook_name, created_at DESC);
      `);

      logger.info(
        "Created workbook_diffs and workbook_summaries tables with diff_id linking",
      );
    },
  },
  {
    version: 4,
    name: "add_workbook_snapshots_table",
    up: (db: DB) => {
      logger.info("Running migration 4: add_workbook_snapshots_table");

      // Table to persist the latest workbook snapshot per workbook
      db.execute(`
        CREATE TABLE IF NOT EXISTS workbook_snapshots (
          workbook_name TEXT PRIMARY KEY,
          snapshot_json TEXT NOT NULL,
          updated_at INTEGER NOT NULL
        );
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_snapshots_updated_at
        ON workbook_snapshots(updated_at DESC);
      `);

      logger.info("Created workbook_snapshots table");
    },
  },
  {
    version: 5,
    name: "add_user_preferences_table",
    up: (db: DB) => {
      logger.info("Running migration 5: add_user_preferences_table");

      // Table to store user preferences (e.g., consent for services)
      // Uses composite primary key to allow multiple preferences per user
      db.execute(`
        CREATE TABLE IF NOT EXISTS user_preferences (
          user_email TEXT NOT NULL,
          preference_key TEXT NOT NULL,
          preference_value TEXT NOT NULL,
          updated_at INTEGER NOT NULL,
          PRIMARY KEY (user_email, preference_key)
        );
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_user_preferences_email
        ON user_preferences(user_email);
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_user_preferences_key
        ON user_preferences(preference_key);
      `);

      logger.info("Created user_preferences table");
    },
  },
  {
    version: 6,
    name: "add_comparison_metadata_to_workbook_diffs_and_summaries",
    up: (db: DB) => {
      logger.info(
        "Running migration 6: add_comparison_metadata_to_workbook_diffs_and_summaries",
      );

      // Add comparison type to workbook_diffs (only type, not full metadata)
      db.execute(`
        ALTER TABLE workbook_diffs
        ADD COLUMN comparison_type TEXT NOT NULL DEFAULT 'self';
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_diffs_comparison_type
        ON workbook_diffs(comparison_type);
      `);

      // Add full comparison metadata to workbook_summaries
      db.execute(`
        ALTER TABLE workbook_summaries
        ADD COLUMN comparison_type TEXT NOT NULL DEFAULT 'self';
      `);

      db.execute(`
        ALTER TABLE workbook_summaries
        ADD COLUMN comparison_file_path TEXT;
      `);

      db.execute(`
        ALTER TABLE workbook_summaries
        ADD COLUMN comparison_file_modified_at INTEGER;
      `);

      // Add indexes for summaries
      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_summaries_comparison_type
        ON workbook_summaries(comparison_type);
      `);

      db.execute(`
        CREATE INDEX IF NOT EXISTS idx_workbook_summaries_workbook_comparison
        ON workbook_summaries(workbook_name, comparison_type);
      `);

      // Populate comparison_type in summaries from linked diffs for existing records
      db.execute(`
        UPDATE workbook_summaries
        SET comparison_type = COALESCE(
          (SELECT comparison_type FROM workbook_diffs WHERE id = workbook_summaries.diff_id),
          'self'
        )
        WHERE diff_id IS NOT NULL;
      `);

      logger.info(
        "Added comparison metadata to workbook_diffs and workbook_summaries tables",
      );
    },
  },
  {
    version: 7,
    name: "add_deployment_and_reasoning_to_llm_keys_config",
    up: (db: DB) => {
      logger.info(
        "Running migration 7: add_deployment_and_reasoning_to_llm_keys_config",
      );

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN azure_openai_deployment_name TEXT;
      `);

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN openai_reasoning_effort TEXT;
      `);

      logger.info(
        "Added azure_openai_deployment_name and openai_reasoning_effort columns to llm_keys_config table",
      );
    },
  },
  {
    version: 8,
    name: "add_model_to_llm_keys_config",
    up: (db: DB) => {
      logger.info(
        "Running migration 8: add_model_to_llm_keys_config",
      );

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN azure_openai_model TEXT;
      `);

      logger.info(
        "Added azure_openai_model column to llm_keys_config table",
      );
    },
  },
  {
    version: 9,
    name: "add_anthropic_fields_to_llm_keys_config",
    up: (db: DB) => {
      logger.info(
        "Running migration 9: add_anthropic_fields_to_llm_keys_config",
      );

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN anthropic_api_key TEXT;
      `);

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN anthropic_model TEXT;
      `);

      logger.info(
        "Added anthropic_api_key and anthropic_model columns to llm_keys_config table",
      );
    },
  },
  {
    version: 10,
    name: "add_anthropic_base_url_to_llm_keys_config",
    up: (db: DB) => {
      logger.info(
        "Running migration 10: add_anthropic_base_url_to_llm_keys_config",
      );

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN anthropic_base_url TEXT;
      `);

      logger.info(
        "Added anthropic_base_url column to llm_keys_config table",
      );
    },
  },
  {
    version: 11,
    name: "add_anthropic_ca_bundle_path_to_llm_keys_config",
    up: (db: DB) => {
      logger.info(
        "Running migration 11: add_anthropic_ca_bundle_path_to_llm_keys_config",
      );

      db.execute(`
        ALTER TABLE llm_keys_config
        ADD COLUMN anthropic_ca_bundle_path TEXT;
      `);

      logger.info(
        "Added anthropic_ca_bundle_path column to llm_keys_config table",
      );
    },
  },
];

// Run all pending migrations
export function runMigrations(db: DB): void {
  const currentVersion = getCurrentVersion(db);
  logger.info(`Current schema version: ${currentVersion}`);

  const pendingMigrations = migrations.filter((m) =>
    m.version > currentVersion
  );

  if (pendingMigrations.length === 0) {
    logger.info("No pending migrations");
    return;
  }

  logger.info(`Found ${pendingMigrations.length} pending migration(s)`);

  // Run migrations in a transaction
  db.execute("BEGIN TRANSACTION;");

  try {
    for (const migration of pendingMigrations) {
      logger.info(`Applying migration ${migration.version}: ${migration.name}`);
      migration.up(db);
      recordMigration(db, migration);
      logger.info(`Migration ${migration.version} applied successfully`);
    }

    db.execute("COMMIT;");
    logger.info("All migrations completed successfully");
  } catch (error) {
    db.execute("ROLLBACK;");
    logger.error("Migration failed, rolled back", error);
    throw error;
  }
}
