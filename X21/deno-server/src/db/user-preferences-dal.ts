import { getDb, isDbInitialized, nowMs } from "./sqlite.ts";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("UserPreferencesDAL");

export interface UserPreference {
  userEmail: string;
  preferenceKey: string;
  preferenceValue: string;
  updatedAt: number;
}

/**
 * Get a user preference value by email and key
 * Returns null if not found
 */
export function getUserPreference(
  userEmail: string,
  preferenceKey: string,
): string | null {
  if (!isDbInitialized()) {
    logger.debug("Database not initialized yet, skipping preference lookup");
    return null;
  }

  const db = getDb();

  const rows = db.query<[string, number]>(
    `SELECT preference_value, updated_at
     FROM user_preferences
     WHERE user_email = ? AND preference_key = ?`,
    [userEmail, preferenceKey],
  );

  if (rows.length === 0) {
    return null;
  }

  const [value] = rows[0];
  logger.debug(`Retrieved preference ${preferenceKey} for user ${userEmail}`);
  return value;
}

/**
 * Get a boolean user preference (converts "true"/"false" strings)
 * Returns defaultValue if not found
 */
export function getUserPreferenceBool(
  userEmail: string,
  preferenceKey: string,
  defaultValue: boolean = false,
): boolean {
  const value = getUserPreference(userEmail, preferenceKey);
  if (value === null) {
    return defaultValue;
  }
  return value.toLowerCase() === "true" || value === "1";
}

/**
 * Set a user preference value
 */
export function setUserPreference(
  userEmail: string,
  preferenceKey: string,
  preferenceValue: string,
): void {
  if (!isDbInitialized()) {
    logger.warn("Database not initialized, cannot save preference");
    return;
  }

  const db = getDb();
  const now = nowMs();

  db.query(
    `INSERT INTO user_preferences (user_email, preference_key, preference_value, updated_at)
     VALUES (?, ?, ?, ?)
     ON CONFLICT(user_email, preference_key)
     DO UPDATE SET
       preference_value = excluded.preference_value,
       updated_at = excluded.updated_at`,
    [userEmail, preferenceKey, preferenceValue, now],
  );

  logger.info(
    `Set preference ${preferenceKey} = ${preferenceValue} for user ${userEmail}`,
  );
}

/**
 * Set a boolean user preference (converts to "true"/"false" string)
 */
export function setUserPreferenceBool(
  userEmail: string,
  preferenceKey: string,
  value: boolean,
): void {
  setUserPreference(userEmail, preferenceKey, value ? "true" : "false");
}

/**
 * Get all preferences for a user
 */
export function getAllUserPreferences(
  userEmail: string,
): UserPreference[] {
  if (!isDbInitialized()) {
    logger.debug("Database not initialized yet, skipping preference lookup");
    return [];
  }

  const db = getDb();

  const rows = db.query<[string, string, string, number]>(
    `SELECT user_email, preference_key, preference_value, updated_at
     FROM user_preferences
     WHERE user_email = ?
     ORDER BY preference_key`,
    [userEmail],
  );

  return rows.map(([email, key, value, updatedAt]) => ({
    userEmail: email,
    preferenceKey: key,
    preferenceValue: value,
    updatedAt,
  }));
}

/**
 * Delete a user preference
 */
export function deleteUserPreference(
  userEmail: string,
  preferenceKey: string,
): void {
  if (!isDbInitialized()) {
    logger.warn("Database not initialized, cannot delete preference");
    return;
  }

  const db = getDb();

  db.query(
    `DELETE FROM user_preferences
     WHERE user_email = ? AND preference_key = ?`,
    [userEmail, preferenceKey],
  );

  logger.info(`Deleted preference ${preferenceKey} for user ${userEmail}`);
}

/**
 * Delete all preferences for a user
 */
export function deleteAllUserPreferences(userEmail: string): void {
  if (!isDbInitialized()) {
    logger.warn("Database not initialized, cannot delete preferences");
    return;
  }

  const db = getDb();

  db.query(
    `DELETE FROM user_preferences WHERE user_email = ?`,
    [userEmail],
  );

  logger.info(`Deleted all preferences for user ${userEmail}`);
}
