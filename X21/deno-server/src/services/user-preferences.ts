import { createLogger } from "../utils/logger.ts";
import {
  getUserPreference,
  getUserPreferenceBool,
  setUserPreference,
  setUserPreferenceBool,
} from "../db/user-preferences-dal.ts";
import { UserService } from "./user.ts";

const logger = createLogger("UserPreferencesService");

/**
 * Service for managing user preferences
 * Provides convenient methods for common preference operations
 */
export class UserPreferencesService {
  private static instance: UserPreferencesService | null = null;

  private constructor() {}

  public static getInstance(): UserPreferencesService {
    if (!UserPreferencesService.instance) {
      UserPreferencesService.instance = new UserPreferencesService();
    }
    return UserPreferencesService.instance;
  }

  /**
   * Get the current user's email from UserService
   */
  private getCurrentUserEmail(): string | null {
    const email = UserService.getInstance().getUserEmail();
    // UserService returns "Email not set" as a string when email is null
    if (email === null || email === "Email not set") {
      return null;
    }
    return email;
  }

  /**
   * Get a preference for the current user
   * Returns null if user is not logged in or preference doesn't exist
   */
  getPreference(preferenceKey: string): string | null {
    const email = this.getCurrentUserEmail();
    if (!email) {
      logger.debug("No user email available, cannot get preference");
      return null;
    }
    return getUserPreference(email, preferenceKey);
  }

  /**
   * Get a boolean preference for the current user
   * Returns defaultValue if user is not logged in or preference doesn't exist
   */
  getPreferenceBool(
    preferenceKey: string,
    defaultValue: boolean = false,
  ): boolean {
    const email = this.getCurrentUserEmail();
    if (!email) {
      logger.debug("No user email available, returning default value");
      return defaultValue;
    }
    return getUserPreferenceBool(email, preferenceKey, defaultValue);
  }

  /**
   * Set a preference for the current user
   * Returns false if user is not logged in
   */
  setPreference(preferenceKey: string, value: string): boolean {
    const email = this.getCurrentUserEmail();
    if (!email) {
      logger.warn("No user email available, cannot set preference");
      return false;
    }
    setUserPreference(email, preferenceKey, value);
    return true;
  }

  /**
   * Set a boolean preference for the current user
   * Returns false if user is not logged in
   */
  setPreferenceBool(preferenceKey: string, value: boolean): boolean {
    const email = this.getCurrentUserEmail();
    if (!email) {
      logger.warn("No user email available, cannot set preference");
      return false;
    }
    setUserPreferenceBool(email, preferenceKey, value);
    return true;
  }

  /**
   * Check if user has consented to Save Copies
   */
  hasSaveSnapshotsConsent(): boolean {
    return this.getPreferenceBool("save_snapshots", false);
  }

  /**
   * Set user's consent for Save Copies
   */
  setSaveSnapshotsConsent(consented: boolean): boolean {
    return this.setPreferenceBool("save_snapshots", consented);
  }
}
