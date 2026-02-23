import { ToolNames } from "@/types/chat";

const TOOL_NAME_WITH_RANGE = new Set<string>([
  ToolNames.WRITE_VALUES_BATCH,
  ToolNames.READ_VALUES_BATCH,
  ToolNames.WRITE_FORMAT_BATCH,
  ToolNames.READ_FORMAT_BATCH,
  ToolNames.DRAG_FORMULA,
  ToolNames.ADD_COLUMNS,
  ToolNames.REMOVE_COLUMNS,
  ToolNames.ADD_ROWS,
  ToolNames.REMOVE_ROWS,
  ToolNames.ADD_SHEETS,
]);

// Match Excel ranges, optionally prefixed with a sheet name.
// Supports sheet!A1, workbook!sheet!A1, and [workbook]sheet!A1 forms.
export const rangePattern =
  /(?:(?:\[[^\]]+\][^!]+|(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)!(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*)|(?:'[^']+'|[A-Za-z0-9_.-]+(?:\s+[A-Za-z0-9_.-]+)*))!)?[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?/g;

const quoteNameIfNeeded = (name?: string) => {
  if (!name) return "";
  const needsQuotes =
    /[\s,]/.test(name) || name.includes("'") || name.includes("!");
  if (!needsQuotes) return name;
  return `'${name.replace(/'/g, "''")}'`;
};

const asString = (value: unknown) =>
  typeof value === "string" ? value : undefined;

const formatAddress = (
  workbookName?: string,
  worksheet?: string,
  range?: string,
  activeWorkbook?: string,
) => {
  const safeWorkbook =
    workbookName && workbookName !== activeWorkbook
      ? quoteNameIfNeeded(workbookName)
      : "";
  const safeSheet = quoteNameIfNeeded(worksheet);

  if (safeWorkbook && safeSheet && range)
    return `${safeWorkbook}!${safeSheet}!${range}`;
  if (safeWorkbook && range) return `${safeWorkbook}!${range}`;
  if (safeSheet && range) return `${safeSheet}!${range}`;
  if (range) return range;
  if (safeSheet) return safeSheet;
  if (safeWorkbook) return safeWorkbook;
  return "";
};

export const formatKeyInfo = (toolName: string, params: any): string => {
  if (!params || typeof params !== "object") return "";

  if (toolName === ToolNames.WRITE_VALUES_BATCH) {
    const ops = Array.isArray(params.operations)
      ? (params.operations as Array<Record<string, unknown>>)
      : [];
    if (ops.length > 0) {
      const first = ops[0] ?? {};
      const labelWorksheet = asString(first.worksheet);
      const labelRange = asString(first.range);
      const labelWorkbook = asString(first.workbookName);
      const address = formatAddress(
        labelWorkbook,
        labelWorksheet,
        labelRange,
        asString(first.activeWorkbookName),
      );
      if (ops.length > 1 && address) return ` ${address} (+${ops.length - 1})`;
      if (ops.length > 1) return ` ${ops.length} op(s)`;
      if (address) return ` ${address}`;
      return ops.length > 0 ? ` ${ops.length} op(s)` : "";
    }

    const workbookName = params.workbookName;
    const worksheet = params.worksheet;
    const range = params.range;
    const address = formatAddress(
      workbookName,
      worksheet,
      range,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (
    toolName === ToolNames.READ_VALUES_BATCH ||
    toolName === ToolNames.READ_FORMAT_BATCH
  ) {
    const ops = Array.isArray(params.operations) ? params.operations : [];
    if (ops.length > 0) {
      const addresses = ops
        .map((op: Record<string, unknown>) =>
          formatAddress(
            typeof op?.workbookName === "string"
              ? op.workbookName
              : params?.workbookName,
            typeof op?.worksheet === "string" ? op.worksheet : undefined,
            typeof op?.range === "string" ? op.range : undefined,
            params?.activeWorkbookName,
          ),
        )
        .filter(Boolean);
      if (addresses.length > 0) return ` ${addresses.join(", ")}`;
      return ` ${ops.length} op(s)`;
    }
  }

  if (
    toolName === ToolNames.WRITE_FORMAT_BATCH ||
    toolName === ToolNames.READ_VALUES_BATCH ||
    toolName === ToolNames.READ_FORMAT_BATCH
  ) {
    const workbookName = params.workbookName;
    const worksheet = params.worksheet;
    const range = params.range;

    const address = formatAddress(
      workbookName,
      worksheet,
      range,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (toolName === ToolNames.WRITE_FORMAT_BATCH) {
    const ops = Array.isArray(params.operations)
      ? (params.operations as Array<Record<string, unknown>>)
      : [];
    const first = ops[0] ?? {};
    const labelWorksheet = asString(first.worksheet);
    const labelRange = asString(first.range);
    const labelWorkbook = asString(first.workbookName);
    const address = formatAddress(
      labelWorkbook,
      labelWorksheet,
      labelRange,
      asString(first.activeWorkbookName),
    );
    if (ops.length > 1 && address) return ` ${address} (+${ops.length - 1})`;
    if (ops.length > 1) return ` ${ops.length} op(s)`;
    if (address) return ` ${address}`;
    return ops.length > 0 ? ` ${ops.length} op(s)` : "";
  }

  if (
    toolName === ToolNames.ADD_COLUMNS ||
    toolName === ToolNames.REMOVE_COLUMNS
  ) {
    const workbookName = params.workbookName;
    const worksheet = params.worksheet;
    const columnRange = params.columnRange;

    const address = formatAddress(
      workbookName,
      worksheet,
      columnRange,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (toolName === ToolNames.ADD_ROWS || toolName === ToolNames.REMOVE_ROWS) {
    const workbookName = params.workbookName;
    const worksheet = params.worksheet;
    const rowRange = params.rowRange;

    const address = formatAddress(
      workbookName,
      worksheet,
      rowRange,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (toolName === ToolNames.ADD_SHEETS) {
    const workbookName = params.workbookName;
    const sheetName = params.sheetName;
    const address = formatAddress(
      workbookName,
      sheetName,
      undefined,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (toolName === ToolNames.DRAG_FORMULA) {
    const workbookName = params.workbookName;
    const worksheet = params.worksheet;
    const sourceRange = params.sourceRange;

    const address = formatAddress(
      workbookName,
      worksheet,
      sourceRange,
      params.activeWorkbookName,
    );
    return address ? ` ${address}` : "";
  }

  if (toolName === ToolNames.MERGE_FILES) {
    const sourceFolder =
      params.sourceFolder ||
      params.folderPath ||
      params.inputFolder ||
      params.directory;
    const outputFile =
      params.outputFile || params.targetWorkbook || params.outputPath;
    if (sourceFolder && outputFile) {
      return ` ${sourceFolder} → ${outputFile}`;
    }
    if (sourceFolder) return ` ${sourceFolder}`;
    if (outputFile) return ` ${outputFile}`;
  }

  const entries = Object.entries(params);
  if (entries.length > 0) {
    const [, value] = entries[0];
    if (typeof value === "string" && value.length < 30) {
      return ` ${value}`;
    }
  }

  return "";
};

export const extractKeyInfo = (
  toolName: string,
  content: string,
  _blockId: string,
) => {
  if (!content.trim()) {
    return "";
  }

  try {
    const params = JSON.parse(content);
    return formatKeyInfo(toolName, params);
  } catch {
    try {
      const partialJson = content.includes("{")
        ? content + "}"
        : `{${content}}`;
      const params = JSON.parse(partialJson);
      return formatKeyInfo(toolName, params);
    } catch {
      return "";
    }
  }
};

export const shouldDisplayRangeInfo = (toolName?: string) => {
  return toolName ? TOOL_NAME_WITH_RANGE.has(toolName) : false;
};
