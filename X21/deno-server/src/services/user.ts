import { createLogger } from "../utils/logger.ts";
import { setLogUserEmail } from "../utils/log-context.ts";

const logger = createLogger("UserService");

export class UserService {
  private static instance: UserService | null = null;
  private userEmail: string | null = null;

  private constructor() {}

  public static getInstance(): UserService {
    if (!UserService.instance) {
      UserService.instance = new UserService();
    }
    return UserService.instance;
  }

  setUserEmail(email: string | null): void {
    logger.info(`User email set to: ${email}`);
    this.userEmail = email;
    const normalizedEmail = !email || email === "Email not set" ? null : email;
    setLogUserEmail(normalizedEmail);
  }

  getUserEmail(): string | null {
    return this.userEmail || "Email not set";
  }
}
