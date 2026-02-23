import { createTwoFilesPatch } from "diff";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("WorkbookDiff");

export interface WorkbookSnapshot {
  fileName?: string;
  workbookXml: string | null;
  workbookRelsXml: string | null;
  sharedStringsXml: string | null;
  sheetXmls: Record<string, string>;
}

export interface DiffResult {
  sheetName: string;
  sheetPath: string;
  unifiedDiff: string;
  hasChanges: boolean;
  isActualSheet: boolean; // true for worksheet sheets, false for internal Excel files like sharedStrings
}

/**
 * Pretty-print XML for better diff readability
 */
function prettyPrintXml(xmlString: string): string {
  try {
    const lines = xmlString
      .replace(/>\s*</g, ">\n<")
      .split("\n");

    let indent = 0;
    const INDENT = "  ";

    return lines
      .map((line) => {
        const trimmed = line.trim();
        if (!trimmed) return "";

        // closing tag
        if (/^<\/.+>/.test(trimmed)) {
          indent = Math.max(indent - 1, 0);
        }

        const result = INDENT.repeat(indent) + trimmed;

        // opening tag that is NOT:
        // - closing
        // - self-closing
        // - inline (<a></a>)
        // - declaration / comment
        if (
          /^<[^!?/][^>]*>$/.test(trimmed) &&
          !trimmed.endsWith("/>") &&
          !trimmed.includes("</")
        ) {
          indent++;
        }

        return result;
      })
      .filter(Boolean)
      .join("\n");
  } catch {
    return xmlString;
  }
}

/**
 * Compare two workbook snapshots and generate unified diffs
 */
export function generateWorkbookDiff(
  previous: WorkbookSnapshot,
  current: WorkbookSnapshot,
): DiffResult[] {
  const results: DiffResult[] = [];

  logger.info("Starting workbook diff generation");

  const prevFileName = previous.fileName || "previous.xlsx";
  const currFileName = current.fileName || "current.xlsx";

  // Compare sharedStrings.xml (contains text values)
  if (previous.sharedStringsXml || current.sharedStringsXml) {
    const prevSharedStrings = previous.sharedStringsXml || "";
    const currSharedStrings = current.sharedStringsXml || "";

    if (prevSharedStrings !== currSharedStrings) {
      const unifiedDiff = generateUnifiedDiff(
        prevSharedStrings,
        currSharedStrings,
        prevFileName,
        currFileName,
        "xl/sharedStrings.xml",
        "Shared Strings (Text Values)",
      );

      if (unifiedDiff) {
        results.push({
          sheetName: "Shared Strings",
          sheetPath: "xl/sharedStrings.xml",
          unifiedDiff,
          hasChanges: true,
          isActualSheet: false, // This is an internal Excel file, not a worksheet
        });
      }
    }
  }

  // Parse sheet name mappings from workbook.xml for both snapshots
  const prevNameToPath = parseSheetNameMappings(
    previous.workbookXml,
    previous.workbookRelsXml,
    previous.sheetXmls,
  );
  const currNameToPath = parseSheetNameMappings(
    current.workbookXml,
    current.workbookRelsXml,
    current.sheetXmls,
  );

  // Get all unique sheet names from both workbooks
  const allSheetNames = new Set([
    ...prevNameToPath.keys(),
    ...currNameToPath.keys(),
  ]);

  // Compare sheets by name (not by position)
  for (const sheetName of allSheetNames) {
    const prevPath = prevNameToPath.get(sheetName);
    const currPath = currNameToPath.get(sheetName);
    const prevXml = prevPath ? previous.sheetXmls[prevPath] : undefined;
    const currXml = currPath ? current.sheetXmls[currPath] : undefined;

    if (!prevPath && currPath && currXml) {
      // Sheet was added (exists in current but not in previous)
      results.push({
        sheetName,
        sheetPath: currPath,
        unifiedDiff: `Sheet "${sheetName}" was added`,
        hasChanges: true,
        isActualSheet: true,
      });
    } else if (prevPath && !currPath && prevXml) {
      // Sheet was deleted (exists in previous but not in current)
      results.push({
        sheetName,
        sheetPath: prevPath,
        unifiedDiff: `Sheet "${sheetName}" was deleted`,
        hasChanges: true,
        isActualSheet: true,
      });
    } else if (prevPath && currPath && prevXml && currXml) {
      // Sheet exists in both - generate unified diff
      // Note: paths might differ if sheets were reordered, but names match
      const unifiedDiff = generateUnifiedDiff(
        prevXml,
        currXml,
        prevFileName,
        currFileName,
        currPath, // Use current path for display
        sheetName,
      );

      if (unifiedDiff) {
        results.push({
          sheetName,
          sheetPath: currPath,
          unifiedDiff,
          hasChanges: true,
          isActualSheet: true,
        });
      }
    }
    // Note: If both paths exist but XML is missing, we skip it (shouldn't happen in normal cases)
  }

  logger.info(`Generated ${results.length} diff results`);
  return results;
}

/**
 * Parse workbook.xml.rels to extract rId to path mappings
 * Structure: <Relationship Id="rId1" Type="..." Target="worksheets/sheet1.xml"/>
 */
function parseWorkbookRels(
  workbookRelsXml: string | null,
): Map<string, string> {
  const rIdToPath = new Map<string, string>();

  if (!workbookRelsXml) {
    return rIdToPath;
  }

  try {
    // Pattern: <Relationship Id="rIdX" ... Target="worksheets/sheetN.xml"/>
    const relRegex =
      /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*>/g;
    let match;
    while ((match = relRegex.exec(workbookRelsXml)) !== null) {
      const rId = match[1];
      let target = match[2];

      // Normalize path: "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
      if (!target.startsWith("xl/")) {
        target = `xl/${target}`;
      }

      rIdToPath.set(rId, target);
    }
  } catch (error) {
    logger.warn("Failed to parse workbook.xml.rels", {
      error: error instanceof Error ? error.message : String(error),
    });
  }

  return rIdToPath;
}

/**
 * Parse workbook.xml to extract sheet name to path mappings
 * workbook.xml structure: <sheet name="SheetName" sheetId="1" r:id="rId1"/>
 * Resolve rId to actual path via workbook.xml.rels
 */
function parseSheetNameMappings(
  workbookXml: string | null,
  workbookRelsXml: string | null,
  sheetXmls: Record<string, string>,
): Map<string, string> {
  const nameToPath = new Map<string, string>();
  const availablePaths = Object.keys(sheetXmls);

  if (!workbookXml) {
    logger.warn("workbook.xml not available, using path-based fallback", {
      availablePaths,
    });
    // Fallback: use path-based names if workbook.xml is not available
    for (const path of availablePaths) {
      const match = path.match(/sheet(\d+)\.xml/);
      const name = match ? `Sheet${match[1]}` : path;
      nameToPath.set(name, path);
    }
    return nameToPath;
  }

  try {
    // First, parse workbook.xml.rels to get rId -> path mappings
    const rIdToPath = parseWorkbookRels(workbookRelsXml);

    // Parse workbook.xml to extract sheet names and rIds
    // Pattern: <sheet name="..." sheetId="N" r:id="rIdX"/>
    // Note: r:id may appear before or after other attributes
    const sheetRegex =
      /<sheet\s+[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"[^>]*>|<sheet\s+[^>]*r:id="([^"]+)"[^>]*name="([^"]+)"[^>]*>/g;
    const sheetMatches: Array<{ name: string; rId: string }> = [];

    let match;
    while ((match = sheetRegex.exec(workbookXml)) !== null) {
      // Match can have name in position 1 or 4, rId in position 2 or 3
      const name = match[1] || match[4];
      const rId = match[2] || match[3];
      if (name && rId) {
        sheetMatches.push({ name, rId });
      }
    }

    // Map sheet names to paths using rId
    for (const { name, rId } of sheetMatches) {
      const path = rIdToPath.get(rId);
      if (path && sheetXmls[path]) {
        // Exact match via rId
        nameToPath.set(name, path);
      } else if (path) {
        // rId found but path doesn't exist in sheetXmls
        nameToPath.set(name, path);
        logger.warn(
          "Sheet name mapped via rId but path not found in sheetXmls",
          {
            name,
            rId,
            path,
            availablePaths,
          },
        );
      } else {
        // rId not found in workbook.xml.rels - fallback to sheetId-based matching
        // Extract sheetId as fallback
        const sheetIdMatch = workbookXml.match(
          new RegExp(
            `<sheet\\s+[^>]*name="${
              name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
            }"[^>]*sheetId="(\\d+)"`,
            "i",
          ),
        );
        if (sheetIdMatch) {
          const sheetId = parseInt(sheetIdMatch[1], 10);
          const expectedPath = `xl/worksheets/sheet${sheetId}.xml`;
          if (sheetXmls[expectedPath]) {
            nameToPath.set(name, expectedPath);
          } else {
            // Try fuzzy match
            const matchingPath = availablePaths.find((p) =>
              p.includes(`sheet${sheetId}`)
            );
            if (matchingPath) {
              nameToPath.set(name, matchingPath);
            } else {
              nameToPath.set(name, expectedPath);
              logger.warn(
                "Sheet name mapped via sheetId fallback (rId not found)",
                {
                  name,
                  rId,
                  sheetId,
                  expectedPath,
                  availablePaths,
                },
              );
            }
          }
        } else {
          logger.warn(
            "Could not map sheet name (rId and sheetId both failed)",
            {
              name,
              rId,
              availablePaths,
            },
          );
        }
      }
    }
  } catch (error) {
    logger.warn(
      "Failed to parse workbook.xml, falling back to path-based names",
      {
        error,
      },
    );
    // Fallback to path-based names
    for (const path of availablePaths) {
      const name = extractSheetNameFromPath(path);
      nameToPath.set(name, path);
    }
  }

  return nameToPath;
}

/**
 * Extract sheet name from path like "xl/worksheets/sheet1.xml"
 * This is a fallback when workbook.xml is not available
 */
function extractSheetNameFromPath(path: string): string {
  const match = path.match(/sheet(\d+)\.xml/);
  return match ? `Sheet${match[1]}` : path;
}

/**
 * Generate a unified diff (git-style) for two sheet XMLs
 */
function generateUnifiedDiff(
  prevXml: string,
  currXml: string,
  prevFileName: string,
  currFileName: string,
  sheetPath: string,
  sheetName: string,
): string | null {
  try {
    // Quick check if XMLs are identical
    if (prevXml === currXml) {
      return null;
    }

    // Pretty-print both XMLs for better diff readability
    const prettyPrev = prettyPrintXml(prevXml);
    const prettyCurr = prettyPrintXml(currXml);

    // Generate unified diff
    const patch = createTwoFilesPatch(
      `a/${prevFileName}/${sheetPath}`,
      `b/${currFileName}/${sheetPath}`,
      prettyPrev,
      prettyCurr,
      `Sheet: ${sheetName}`,
      `Sheet: ${sheetName}`,
      { context: 3 }, // Number of context lines
    );

    // Remove the file header lines if no actual changes
    if (!patch.includes("@@")) {
      return null;
    }

    // Add a header section
    let output = "\n" + "=".repeat(80) + "\n";
    output += `Comparing Sheet: '${sheetName}'\n`;
    output += `  File 1: ${prevFileName}\n`;
    output += `  File 2: ${currFileName}\n`;
    output += "=".repeat(80) + "\n\n";
    output += "📄 Found sheet XML files:\n";
    output += `  File 1: ${sheetPath}\n`;
    output += `  File 2: ${sheetPath}\n\n`;
    output += "─".repeat(80) + "\n";
    output += "📝 Differences found (unified diff):\n";
    output += "─".repeat(80) + "\n\n";
    output += patch;
    output += "\n" + "=".repeat(80) + "\n";

    return output;
  } catch (error) {
    logger.error(`Error generating unified diff for ${sheetName}:`, error);
    return null;
  }
}

/**
 * Count changes in a diff result
 */
export function countChanges(diffResult: DiffResult): {
  added: number;
  deleted: number;
  total: number;
} {
  const lines = diffResult.unifiedDiff.split("\n");
  const added =
    lines.filter((l) => l.startsWith("+") && !l.startsWith("+++")).length;
  const deleted =
    lines.filter((l) => l.startsWith("-") && !l.startsWith("---")).length;
  return { added, deleted, total: added + deleted };
}

/**
 * Count total changes across all diff results
 */
export function countTotalChanges(results: DiffResult[]): number {
  return results.reduce((acc, r) => acc + countChanges(r).total, 0);
}

/**
 * Format diff results as a human-readable string (with unified diffs)
 */
export function formatDiffResults(results: DiffResult[]): string {
  if (results.length === 0) {
    return "No changes detected in the workbook.";
  }

  let output = "";

  for (const result of results) {
    output += result.unifiedDiff;
    output += "\n";
  }

  return output;
}

/**
 * Create a concise summary for LLM consumption
 */
export function createDiffSummaryForLLM(results: DiffResult[]): string {
  if (results.length === 0) {
    return "No changes detected.";
  }

  let summary = `Excel File Changes Summary\n`;
  summary += `${"=".repeat(80)}\n\n`;
  summary += `Changes detected in ${results.length} sheet(s):\n\n`;

  for (const result of results) {
    summary += `Sheet: ${result.sheetName}\n`;

    // Count the types of changes in the unified diff
    const lines = result.unifiedDiff.split("\n");
    const addedLines = lines.filter((l) =>
      l.startsWith("+") && !l.startsWith("+++")
    ).length;
    const deletedLines =
      lines.filter((l) => l.startsWith("-") && !l.startsWith("---")).length;

    summary += `  - Added lines: ${addedLines}\n`;
    summary += `  - Deleted lines: ${deletedLines}\n`;
    summary += `  - Modified regions: ${
      lines.filter((l) => l.startsWith("@@")).length
    }\n`;
    summary += "\n";
  }

  // Include the full diff for the LLM to analyze
  summary += "\nFull Unified Diff:\n";
  summary += "─".repeat(80) + "\n\n";
  summary += formatDiffResults(results);

  return summary;
}
