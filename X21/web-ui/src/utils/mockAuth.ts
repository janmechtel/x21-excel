/**
 * Mock Authentication Utility
 *
 * Injects a mock Supabase session when VITE_SKIP_AUTH=true
 * This allows testing without requiring real email authentication
 */

const DEFAULT_SUPABASE_STORAGE_KEY = "sb-qvycnlwxhhmuobjzzoos-auth-token";
const SUPABASE_STORAGE_KEY =
  import.meta.env.VITE_SUPABASE_STORAGE_KEY || DEFAULT_SUPABASE_STORAGE_KEY;

export interface MockSession {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  expires_at: number;
  token_type: string;
  user: {
    id: string;
    email: string;
    user_metadata: {
      first_name: string;
      last_name: string;
    };
    app_metadata: Record<string, unknown>;
    aud: string;
    created_at: string;
  };
}

/**
 * Check if auth should be skipped (from env var or window global)
 */
export function shouldSkipAuth(): boolean {
  // Check window global (injected by Vite plugin)
  if (typeof window !== "undefined" && (window as any).__SKIP_AUTH__) {
    return (window as any).__SKIP_AUTH__ === true;
  }

  // Check import.meta.env (for module code)
  if (typeof import.meta !== "undefined" && import.meta.env?.VITE_SKIP_AUTH) {
    return import.meta.env.VITE_SKIP_AUTH === "true";
  }

  return false;
}

/**
 * Create a mock Supabase session
 */
export function createMockSession(
  email: string = "test@kontext21.com",
): MockSession {
  const now = Math.floor(Date.now() / 1000);

  return {
    access_token: `mock-access-token-${Date.now()}`,
    refresh_token: `mock-refresh-token-${Date.now()}`,
    expires_in: 3600,
    expires_at: now + 3600,
    token_type: "bearer",
    user: {
      id: `mock-user-id-${Date.now()}`,
      email: email,
      user_metadata: {
        first_name: "Test",
        last_name: "User",
      },
      app_metadata: {},
      aud: "authenticated",
      created_at: new Date().toISOString(),
    },
  };
}

/**
 * Inject mock authentication session into localStorage
 */
export function injectMockAuth(email: string = "test@kontext21.com"): void {
  if (typeof window === "undefined") {
    return; // Not in browser environment
  }

  const mockSession = createMockSession(email);

  try {
    localStorage.setItem(SUPABASE_STORAGE_KEY, JSON.stringify(mockSession));
    console.log("✅ Mock authentication injected:", email);
  } catch (error) {
    console.error("❌ Failed to inject mock authentication:", error);
  }
}

/**
 * Remove mock authentication from localStorage
 */
export function removeMockAuth(): void {
  if (typeof window === "undefined") {
    return;
  }

  try {
    localStorage.removeItem(SUPABASE_STORAGE_KEY);
    console.log("✅ Mock authentication removed");
  } catch (error) {
    console.error("❌ Failed to remove mock authentication:", error);
  }
}

/**
 * Initialize mock auth if VITE_SKIP_AUTH is enabled
 * Call this early in the app lifecycle (e.g., in main.tsx)
 */
export function initializeMockAuth(): void {
  if (shouldSkipAuth()) {
    console.log("🧪 Skipping authentication (VITE_SKIP_AUTH=true)");
    injectMockAuth();
  }
}
