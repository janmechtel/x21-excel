import { getMetadata } from "../excel-actions/get-metadata.ts";
import { listOpenWorkbooks } from "../excel-actions/list-open-workbooks.ts";
import { createLogger } from "./logger.ts";
import { sanitizeWorkbookName } from "./workbook-name.ts";
import {
  WorkbookResolutionPath,
  WorkbookResolutionPaths,
} from "../types/index.ts";

const logger = createLogger("WorkbookResolver");

export interface WorkbookResolutionResult {
  workbookName: string;
  resolutionPath: WorkbookResolutionPath;
  openWorkbooks?: string[];
  providedWorkbookName?: string;
}

export interface WorkbookResolverDeps {
  getHostMetadata: typeof getMetadata;
  listOpenWorkbooks: typeof listOpenWorkbooks;
}

const defaultDeps: WorkbookResolverDeps = {
  getHostMetadata: getMetadata,
  listOpenWorkbooks: listOpenWorkbooks,
};

function normalizeWorkbookId(value?: string | null): string {
  if (!value) return "";
  return value.trim().toLowerCase();
}

function basenameFromPath(value: string): string {
  const parts = value.split(/[/\\]/);
  return parts[parts.length - 1] || value;
}

function matchesWorkbookId(candidate: string, openItem: string): boolean {
  const candidateNorm = normalizeWorkbookId(candidate);
  const openNorm = normalizeWorkbookId(openItem);
  if (!candidateNorm || !openNorm) return false;
  if (candidateNorm === openNorm) return true;
  return basenameFromPath(candidateNorm) === basenameFromPath(openNorm);
}

export class WorkbookResolutionError extends Error {
  code: string;
  openWorkbooks?: string[];
  constructor(message: string, code: string, openWorkbooks?: string[]) {
    super(message);
    this.code = code;
    this.openWorkbooks = openWorkbooks;
  }
}

/**
 * Resolve the workbook used for write operations, preferring the host's active workbook.
 * Never invents a placeholder (e.g., Book1) unless the host reports it as active.
 */
export async function resolveWorkbookForWrite(
  options: {
    sessionWorkbookName?: string;
    providedWorkbookName?: string;
  },
  deps: WorkbookResolverDeps = defaultDeps,
): Promise<WorkbookResolutionResult> {
  const rawSession = options.sessionWorkbookName;
  const rawProvided = options.providedWorkbookName;
  const sessionWorkbook = sanitizeWorkbookName(rawSession) || rawSession;
  const providedWorkbook = sanitizeWorkbookName(rawProvided) || rawProvided;

  // 1) Look at open workbooks so we can validate session/provided names
  let openWorkbooks: string[] = [];
  try {
    const open = await deps.listOpenWorkbooks();
    openWorkbooks = open
      .flatMap((w) => [w.workbookName, w.workbookFullName])
      .filter(Boolean) as string[];
  } catch (error) {
    logger.warn("Failed to list open workbooks", { error });
  }

  // 2) Use existing session workbook (including placeholders like "Book1")
  if (sessionWorkbook) {
    const sessionIsOpen = openWorkbooks.length === 0 ||
      openWorkbooks.some((name) => matchesWorkbookId(sessionWorkbook, name));
    if (sessionIsOpen) {
      logger.info("Resolved workbook from session", {
        resolutionPath: WorkbookResolutionPaths.SESSION,
        sessionWorkbook,
        providedWorkbook,
      });
      return {
        workbookName: sessionWorkbook,
        resolutionPath: WorkbookResolutionPaths.SESSION,
        providedWorkbookName: providedWorkbook,
      };
    }
    logger.warn("Session workbook not found among open workbooks", {
      sessionWorkbook,
      openWorkbooks,
      providedWorkbook,
    });
  }

  // 3) Ask host for active workbook when session is unavailable
  try {
    const metadata = await deps.getHostMetadata({});
    const hostWorkbook = metadata?.workbookName;
    if (hostWorkbook && hostWorkbook.trim().length > 0) {
      logger.info("Resolved workbook from host metadata", {
        resolutionPath: WorkbookResolutionPaths.HOST_ACTIVE,
        hostWorkbook,
        providedWorkbook,
        sessionWorkbook,
      });
      return {
        workbookName: hostWorkbook,
        resolutionPath: WorkbookResolutionPaths.HOST_ACTIVE,
        providedWorkbookName: providedWorkbook,
      };
    }
  } catch (error) {
    logger.warn("Failed to resolve workbook via host metadata", {
      error,
      sessionWorkbook,
      providedWorkbook,
    });
  }

  // 4) Look at open workbooks to auto-correct
  if (openWorkbooks.length > 0) {
    if (
      sessionWorkbook &&
      openWorkbooks.some((name) => matchesWorkbookId(sessionWorkbook, name))
    ) {
      return {
        workbookName: sessionWorkbook,
        resolutionPath: WorkbookResolutionPaths.SESSION_OPEN,
        openWorkbooks,
        providedWorkbookName: providedWorkbook,
      };
    }

    if (
      providedWorkbook &&
      openWorkbooks.some((name) => matchesWorkbookId(providedWorkbook, name))
    ) {
      return {
        workbookName: providedWorkbook,
        resolutionPath: WorkbookResolutionPaths.PROVIDED_OPEN,
        openWorkbooks,
        providedWorkbookName: providedWorkbook,
      };
    }

    if (openWorkbooks.length === 1) {
      return {
        workbookName: openWorkbooks[0],
        resolutionPath: WorkbookResolutionPaths.SINGLE_OPEN,
        openWorkbooks,
        providedWorkbookName: providedWorkbook,
      };
    }
  }

  logger.error("No active workbook could be resolved", {
    sessionWorkbook,
    providedWorkbook,
    openWorkbooks,
  });

  throw new WorkbookResolutionError(
    "No active workbook is set for this chat. Open or activate a workbook in Excel and retry.",
    "NO_ACTIVE_WORKBOOK_SET",
    openWorkbooks,
  );
}
