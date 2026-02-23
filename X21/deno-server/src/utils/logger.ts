import * as log from "@std/log";
import { logs, SeverityNumber } from "@opentelemetry/api-logs";
import { OTLPLogExporter } from "@opentelemetry/exporter-logs-otlp-http";
import { Resource } from "@opentelemetry/resources";
import {
  BatchLogRecordProcessor,
  LoggerProvider,
} from "@opentelemetry/sdk-logs";
import { getEnvironment, getLogDirectory } from "./environment.ts";
import { getLogUserEmail } from "./log-context.ts";

// Custom file handler with daily rotation
class RotatingFileHandler extends log.BaseHandler {
  private baseFilename: string;
  private logDir: string;
  private currentLogFile: string;
  private currentDate: string;
  private isInitialized: boolean = false;

  constructor(
    levelName: log.LevelName,
    options: {
      baseFilename: string;
      logDir: string;
      formatter?: log.FormatterFunction;
    },
  ) {
    super(levelName, options);
    this.baseFilename = options.baseFilename;
    this.logDir = options.logDir;
    this.currentDate = this.getCurrentDateString();
    this.currentLogFile = `${this.logDir}/${this.baseFilename}.log`; // Always write to current file
  }

  private getCurrentDateString(): string {
    const now = new Date();
    return now.toISOString().split("T")[0]; // YYYY-MM-DD format
  }

  async log(msg: string): Promise<void> {
    try {
      // Check if we need to rotate to a new day (but only after first initialization)
      const today = this.getCurrentDateString();
      if (this.isInitialized && today !== this.currentDate) {
        await this.rotateToNewDay(today);
      }

      // Mark as initialized after first log
      if (!this.isInitialized) {
        this.isInitialized = true;
      }

      // Always write to the current log file (without date)
      await Deno.writeTextFile(this.currentLogFile, msg + "\n", {
        append: true,
      });
    } catch (error) {
      console.error("Failed to write to log file:", error);
    }
  }

  private async rotateToNewDay(newDate: string) {
    try {
      // Archive yesterday's log with the previous date
      const previousDateFile =
        `${this.logDir}/${this.baseFilename}_${this.currentDate}.log`;

      // Copy current log content to the dated archive file
      try {
        const currentContent = await Deno.readTextFile(this.currentLogFile);
        await Deno.writeTextFile(previousDateFile, currentContent);
        console.log(`📅 Archived previous day's logs to: ${previousDateFile}`);
      } catch (error) {
        // If current log doesn't exist yet, that's fine
        if (!(error instanceof Deno.errors.NotFound)) {
          console.error("Failed to archive previous day logs:", error);
        }
      }

      // Clear the current log file for the new day
      await Deno.writeTextFile(
        this.currentLogFile,
        `--- New day started: ${newDate} ---\n`,
      );

      // Update current date
      this.currentDate = newDate;

      console.log(`📅 Rotated to new day: ${newDate}`);
    } catch (error) {
      console.error("Failed to rotate to new day:", error);
    }
  }
}

class PosthogOtelHandler extends log.BaseHandler {
  private otelLogger: ReturnType<typeof logs.getLogger>;

  constructor(
    levelName: log.LevelName,
    otelLogger: ReturnType<typeof logs.getLogger>,
  ) {
    super(levelName, {});
    this.otelLogger = otelLogger;
  }

  log(msg: string, logRecord?: log.LogRecord): void {
    const levelName = logRecord?.levelName ?? "WARN";
    const severityNumber = mapSeverity(levelName);
    if (!severityNumber) return;
    const userEmail = getLogUserEmail();
    const attributes: Record<string, unknown> = {
      logger: logRecord?.loggerName,
      level: levelName,
    };
    if (userEmail) {
      attributes.userEmail = userEmail;
    }
    this.otelLogger.emit({
      severityText: levelName,
      severityNumber,
      body: msg,
      attributes,
      timestamp: logRecord?.datetime,
    });
  }
}

function mapSeverity(level: log.LevelName): SeverityNumber | undefined {
  switch (level) {
    case "WARN":
      return SeverityNumber.WARN;
    case "ERROR":
      return SeverityNumber.ERROR;
    case "CRITICAL":
      return SeverityNumber.FATAL;
    default:
      return undefined;
  }
}

// Logger wrapper that automatically JSON.stringify objects
class LoggerWrapper {
  constructor() {
    // No longer store a logger instance - always get it dynamically
  }

  private formatArgs(...args: unknown[]): string {
    const userEmail = getLogUserEmail();
    return args.map((arg) => {
      if (arg instanceof Error) {
        // JSON.stringify(new Error("x")) => "{}" because Error fields aren't enumerable.
        // Make errors readable in logs.
        const err = arg as any;
        const payload: Record<string, unknown> = {
          name: arg.name,
          message: arg.message,
          stack: arg.stack,
        };
        if (typeof err.cause !== "undefined") payload.cause = err.cause;
        if (arg instanceof AggregateError) {
          payload.errors = (arg as AggregateError).errors;
        }
        if (userEmail && !("userEmail" in payload)) {
          payload.userEmail = userEmail;
        }
        // Include any custom attached fields (best-effort)
        for (const key of Object.getOwnPropertyNames(arg)) {
          if (!(key in payload)) {
            try {
              payload[key] = (arg as any)[key];
            } catch {
              // ignore
            }
          }
        }
        try {
          return JSON.stringify(payload, null, 2);
        } catch {
          return `${arg.name}: ${arg.message}`;
        }
      }
      if (typeof arg === "object" && arg !== null) {
        if (Array.isArray(arg)) {
          try {
            return JSON.stringify(arg, null, 2);
          } catch (error: unknown) {
            return `[Array: ${
              error instanceof Error ? error.message : "Serialization failed"
            }]`;
          }
        }
        const payload = { ...(arg as Record<string, unknown>) };
        if (userEmail && !("userEmail" in payload)) {
          payload.userEmail = userEmail;
        }
        try {
          return JSON.stringify(payload, null, 2);
        } catch (error: unknown) {
          return `[Object: ${
            error instanceof Error ? error.message : "Serialization failed"
          }]`;
        }
      }
      return String(arg);
    }).join(" ");
  }

  private addUserEmailPrefix(message: string): string {
    const userEmail = getLogUserEmail();
    return userEmail ? `[userEmail=${userEmail}] ${message}` : message;
  }

  debug(message: string, ...args: unknown[]): void {
    const formattedMessage = args.length > 0
      ? `${message} ${this.formatArgs(...args)}`
      : message;
    log.getLogger().debug(this.addUserEmailPrefix(formattedMessage));
  }

  info(message: string, ...args: unknown[]): void {
    const formattedMessage = args.length > 0
      ? `${message} ${this.formatArgs(...args)}`
      : message;
    log.getLogger().info(this.addUserEmailPrefix(formattedMessage));
  }

  warn(message: string, ...args: unknown[]): void {
    const formattedMessage = args.length > 0
      ? `${message} ${this.formatArgs(...args)}`
      : message;
    log.getLogger().warn(this.addUserEmailPrefix(formattedMessage));
  }

  error(message: string, ...args: unknown[]): void {
    const formattedMessage = args.length > 0
      ? `${message} ${this.formatArgs(...args)}`
      : message;
    log.getLogger().error(this.addUserEmailPrefix(formattedMessage));
  }

  critical(message: string, ...args: unknown[]): void {
    const formattedMessage = args.length > 0
      ? `${message} ${this.formatArgs(...args)}`
      : message;
    log.getLogger().critical(this.addUserEmailPrefix(formattedMessage));
  }
}

let isConfigured = false;

export async function setupLogger() {
  if (isConfigured) return;

  const environment = getEnvironment();
  const logDir = getLogDirectory();
  const machineName = Deno.hostname();
  const baseLogFile = `deno-${machineName}`;
  const currentLogFile = `${logDir}/${baseLogFile}.log`; // Current log file (no date)

  try {
    await Deno.mkdir(logDir, { recursive: true });
  } catch (error) {
    if (!(error instanceof Deno.errors.AlreadyExists)) {
      console.error("Failed to create log directory:", error);
    }
  }

  console.log(`🔧 Environment: ${environment}`);
  console.log(`📁 Log directory: ${logDir}`);
  console.log(`📄 Current log file: ${currentLogFile}`);
  console.log(`📅 Daily archives: ${baseLogFile}_YYYY-MM-DD.log`);

  // Test if we can write to the log file location (with append: true)
  try {
    await Deno.writeTextFile(
      currentLogFile,
      `Logger initialized at ${new Date().toISOString()}\n`,
      { append: true },
    );
    console.log("✅ File write test successful");
  } catch (error) {
    console.error("❌ File write test failed:", error);
  }

  const logLevel = Deno.env.get("DENO_ENV") === "development"
    ? "DEBUG"
    : "INFO";
  const useColors = Deno.env.get("DENO_LOG_COLORS") !== "false"; // Default to true, set DENO_LOG_COLORS=false to disable
  const enableConsole = Deno.env.get("DENO_CONSOLE_LOGS") !== "false"; // Set to false to disable console output
  console.log(`🔧 Log level: ${logLevel}`);
  console.log(`🎨 Console colors: ${useColors ? "enabled" : "disabled"}`);
  console.log(`📺 Console logging: ${enableConsole ? "enabled" : "disabled"}`);

  try {
    const handlers: any = {
      file: new RotatingFileHandler(logLevel as log.LevelName, {
        baseFilename: baseLogFile,
        logDir: logDir,
        formatter: (logRecord) =>
          `${logRecord.datetime.toISOString()} [${logRecord.levelName}] ${logRecord.msg}`,
      }),
    };

    const posthogApiKey = Deno.env.get("POSTHOG_API_KEY");
    const legacyPosthogApiKey = Deno.env.get("POSTHOG_PROJECT_API_KEY");
    if (!posthogApiKey && legacyPosthogApiKey) {
      console.warn(
        "POSTHOG_PROJECT_API_KEY is deprecated; use POSTHOG_API_KEY instead.",
      );
    }
    const resolvedPosthogApiKey = posthogApiKey || legacyPosthogApiKey;
    const posthogLogsEndpoint = Deno.env.get("POSTHOG_LOGS_ENDPOINT") ||
      "https://us.i.posthog.com/i/v1/logs";
    const posthogLogsEnabled = resolvedPosthogApiKey &&
      Deno.env.get("POSTHOG_LOGS_ENABLED") !== "false";

    let includePosthogHandler = false;
    if (posthogLogsEnabled) {
      const resource = new Resource({
        "service.name": "x21-deno-backend",
        "deployment.environment": environment,
        "host.name": machineName,
      });
      const provider = new LoggerProvider({ resource });
      provider.addLogRecordProcessor(
        new BatchLogRecordProcessor(
          new OTLPLogExporter({
            url: posthogLogsEndpoint,
            headers: {
              Authorization: `Bearer ${resolvedPosthogApiKey}`,
            },
          }),
        ),
      );
      logs.setGlobalLoggerProvider(provider);
      const otelLogger = logs.getLogger("x21-deno-backend");
      handlers.posthog = new PosthogOtelHandler(
        "WARN",
        otelLogger,
      );
      includePosthogHandler = true;
      console.log(
        `📤 PostHog logs enabled (${posthogLogsEndpoint})`,
      );
    } else {
      console.log("📤 PostHog logs disabled");
    }

    // Only add console handler if enabled
    if (enableConsole) {
      handlers.console = new log.ConsoleHandler(logLevel as log.LevelName, {
        formatter: (logRecord) => `${logRecord.levelName} ${logRecord.msg}`,
        useColors: useColors,
      });
    }

    const handlerNames = [
      ...(enableConsole ? ["console"] : []),
      "file",
      ...(includePosthogHandler ? ["posthog"] : []),
    ];

    await log.setup({
      handlers: handlers,
      loggers: {
        default: {
          level: logLevel as log.LevelName,
          handlers: handlerNames,
        },
      },
    });
    console.log("✅ Logger setup completed successfully");
  } catch (error) {
    console.error("❌ Failed to setup logger:", error);
  }

  isConfigured = true;
}

export function createLogger(_name: string): LoggerWrapper {
  // Always get the current logger instance dynamically
  // This ensures loggers created before setupLogger() still work correctly
  return new LoggerWrapper();
}

export type { LogLevel } from "@std/log";
