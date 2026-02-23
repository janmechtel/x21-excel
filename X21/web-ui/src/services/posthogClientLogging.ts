import { getApiBase } from "./apiBase";

type LogContext = Record<string, unknown>;

const LOG_WINDOW_MS = 5 * 60 * 1000;
const MAX_LOGS_PER_WINDOW = 200;
const logTimestamps: number[] = [];
const DEFAULT_SUPABASE_STORAGE_KEY = "sb-qvycnlwxhhmuobjzzoos-auth-token";
const SUPABASE_STORAGE_KEY =
  import.meta.env.VITE_SUPABASE_STORAGE_KEY || DEFAULT_SUPABASE_STORAGE_KEY;
let initialized = false;
type EnvironmentInfo = {
  environment: string;
  logsEnabled: boolean;
};
let envPromise: Promise<EnvironmentInfo> | null = null;
let logsEnabled: boolean | null = null;

function shouldEnable(): boolean {
  return logsEnabled !== false;
}

function serializeArg(arg: unknown): string {
  if (arg instanceof Error) {
    return `${arg.name}: ${arg.message}`;
  }
  if (typeof arg === "string") return arg;
  try {
    return JSON.stringify(arg);
  } catch {
    return String(arg);
  }
}

function allowLog(): boolean {
  const now = Date.now();
  while (logTimestamps.length > 0 && now - logTimestamps[0] > LOG_WINDOW_MS) {
    logTimestamps.shift();
  }
  if (logTimestamps.length >= MAX_LOGS_PER_WINDOW) return false;
  logTimestamps.push(now);
  return true;
}

async function getEnvironmentInfo(): Promise<EnvironmentInfo> {
  if (!envPromise) {
    envPromise = getApiBase()
      .then((baseUrl) =>
        fetch(`${baseUrl}/api/environment`)
          .then((res) => (res.ok ? res.json() : null))
          .then((payload) => ({
            environment: payload?.environment || "Debug",
            logsEnabled: payload?.logsEnabled !== false,
          })),
      )
      .catch(() => ({ environment: "Debug", logsEnabled: true }));
  }
  return envPromise.then((info) => {
    logsEnabled = info.logsEnabled;
    return info;
  });
}

function getUserEmail(): string | null {
  try {
    const storedAuth = localStorage.getItem(SUPABASE_STORAGE_KEY);
    if (storedAuth) {
      const authData = JSON.parse(storedAuth);
      if (authData?.user?.email) {
        return authData.user.email;
      }
    }
  } catch {
    // ignore
  }

  try {
    const keys = Object.keys(localStorage);
    for (const key of keys) {
      if (key.includes("supabase") || key.includes("auth")) {
        const value = localStorage.getItem(key);
        if (!value) continue;
        const data = JSON.parse(value);
        if (data?.user?.email) {
          return data.user.email;
        }
      }
    }
  } catch {
    // ignore
  }

  return null;
}

function buildContext(args: unknown[], extra?: LogContext): LogContext {
  const error = args.find((arg) => arg instanceof Error) as Error | undefined;
  return {
    source: "web-ui",
    userEmail: getUserEmail(),
    args: args.map((arg) => {
      if (arg instanceof Error) {
        return {
          name: arg.name,
          message: arg.message,
          stack: arg.stack,
        };
      }
      return arg;
    }),
    error: error
      ? { name: error.name, message: error.message, stack: error.stack }
      : undefined,
    url: globalThis.location?.href,
    userAgent: globalThis.navigator?.userAgent,
    ...extra,
  };
}

function capture(
  eventName: string,
  message: string,
  context: LogContext,
): void {
  if (!shouldEnable()) return;
  if (!allowLog()) return;
  void getEnvironmentInfo().then(async (info) => {
    try {
      if (!info.logsEnabled) return;
      const baseUrl = await getApiBase();
      await fetch(`${baseUrl}/api/logs`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          level: eventName === "client_warning" ? "warning" : "error",
          message,
          context: {
            environment: info.environment,
            ...context,
          },
        }),
      });
    } catch {
      // Avoid recursive logging on failure.
    }
  });
}

function captureFromArgs(eventName: string, args: unknown[]): void {
  const message = args.map(serializeArg).join(" ");
  capture(eventName, message, buildContext(args));
}

export function initPosthogClientLogging(): void {
  if (initialized || !shouldEnable()) return;
  initialized = true;

  const originalWarn = console.warn.bind(console);
  const originalError = console.error.bind(console);

  console.warn = (...args: unknown[]) => {
    originalWarn(...args);
    captureFromArgs("client_warning", args);
  };

  console.error = (...args: unknown[]) => {
    originalError(...args);
    captureFromArgs("client_error", args);
  };

  globalThis.addEventListener("error", (event) => {
    const error = (event as ErrorEvent).error as Error | undefined;
    const message = event.message || "Uncaught error";
    capture(
      "client_error",
      message,
      buildContext([error ?? message], {
        filename: (event as ErrorEvent).filename,
        lineno: (event as ErrorEvent).lineno,
        colno: (event as ErrorEvent).colno,
      }),
    );
  });

  globalThis.addEventListener("unhandledrejection", (event) => {
    const reason = (event as PromiseRejectionEvent).reason;
    const message =
      reason instanceof Error ? reason.message : "Unhandled promise rejection";
    capture(
      "client_error",
      message,
      buildContext([reason], {
        rejection: true,
      }),
    );
  });
}
