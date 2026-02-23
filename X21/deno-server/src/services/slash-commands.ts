import baseCommands from "../data/slash-commands.json" with { type: "json" };
import { ExcelApiClient } from "../utils/excel-api-client.ts";
import { columnToNumber } from "../utils/excel-range.ts";
import { createLogger } from "../utils/logger.ts";
import {
  ReadValuesBatchRequest,
  ReadValuesBatchResponse,
  ToolNames,
  WorkbookMetadata,
} from "../types/index.ts";
import { fromFileUrl, join } from "@std/path";

export interface SlashCommandDefinition {
  id: string;
  name: string;
  title: string;
  description: string;
  prompt: string;
  icon?: string;
  category?: string;
  keywords?: string[];
  requiresInput?: boolean;
  inputPlaceholder?: string;
  defaultInput?: string;
  promptFile?: string;
}

type RawCellValue = {
  value?: unknown;
  formula?: unknown;
} | unknown;

type SheetSlashCommand = Partial<SlashCommandDefinition> & {
  requiresInput?: boolean | string | number;
  keywords?: string | string[];
  promptFile?: string;
};

type ParsedCell = {
  row: number;
  column: number;
  value: unknown;
};

const logger = createLogger("SlashCommandService");
const WORKSHEET_NAME = "X21_Commands";
const DEFAULT_RANGE = "A1:Z500";

const headerAliases: Record<string, keyof SheetSlashCommand> = {
  id: "id",
  name: "name",
  title: "title",
  description: "description",
  prompt: "prompt",
  icon: "icon",
  requiresinput: "requiresInput",
  inputplaceholder: "inputPlaceholder",
  defaultinput: "defaultInput",
  category: "category",
  keywords: "keywords",
};

const normalizeBool = (
  value: boolean | string | number | undefined,
): boolean => {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return ["true", "1", "yes", "y"].includes(normalized);
  }
  return false;
};

const normalizeKeywords = (value: unknown): string[] | undefined => {
  if (Array.isArray(value)) {
    const cleaned = value.map((item) => item?.toString().trim()).filter(
      Boolean,
    );
    return cleaned.length > 0 ? cleaned as string[] : undefined;
  }

  if (typeof value === "string") {
    const parts = value
      .split(/[,;]+/)
      .map((part) => part.trim())
      .filter(Boolean);
    return parts.length > 0 ? parts : undefined;
  }

  return undefined;
};

const promptCache = new Map<string, string>();
const dataDir = fromFileUrl(new URL("../data/", import.meta.url));
const isWatchMode = Deno.args.includes("--watch") ||
  Deno.env.get("DENO_WATCH") === "true";

const getPromptFromFile = (promptFile?: string): string => {
  if (!promptFile) return "";
  const fileName = promptFile.trim();
  if (!fileName) return "";
  if (promptCache.has(fileName)) {
    return promptCache.get(fileName) ?? "";
  }
  try {
    const filePath = join(dataDir, fileName);
    const content = Deno.readTextFileSync(filePath);
    promptCache.set(fileName, content);
    return content;
  } catch (error) {
    logger.warn("Failed to load prompt file", {
      promptFile: fileName,
      error: (error as Error)?.message,
    });
    promptCache.set(fileName, "");
    return "";
  }
};

const normalizeSlashCommand = (
  row: SheetSlashCommand,
): SlashCommandDefinition | null => {
  const id = (row.id ?? row.name)?.toString().trim();
  if (!id) return null;

  const name = row.name?.toString().trim() || id;
  const title = row.title?.toString().trim() || name;
  let prompt = row.prompt?.toString() || "";
  if (!prompt && row.promptFile) {
    prompt = getPromptFromFile(row.promptFile);
  }

  return {
    id,
    name,
    title,
    description: row.description?.toString().trim() || "",
    prompt,
    icon: row.icon?.toString().trim() || undefined,
    requiresInput: normalizeBool(row.requiresInput),
    inputPlaceholder: row.inputPlaceholder?.toString() || undefined,
    defaultInput: row.defaultInput?.toString() || undefined,
    category: row.category?.toString() || undefined,
    keywords: normalizeKeywords(row.keywords),
  };
};

const normalizeCommands = (
  rows: SheetSlashCommand[],
): SlashCommandDefinition[] => {
  if (!rows || rows.length === 0) return [];
  return rows.map(normalizeSlashCommand).filter((
    cmd,
  ): cmd is SlashCommandDefinition => Boolean(cmd));
};

const normalizeHeader = (header: unknown): string => {
  if (header === null || header === undefined) return "";
  return header.toString().replace(/[\s_]+/g, "").toLowerCase();
};

const parseCellAddress = (
  address: string,
): { row: number; column: number } | null => {
  const match = address.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;

  const [, columnLetters, rowStr] = match;
  const column = columnToNumber(columnLetters.toUpperCase());
  const row = Number(rowStr);

  if (!row || !column) return null;
  return { row, column };
};

class SlashCommandService {
  private static instance: SlashCommandService;
  private baseCommands: SlashCommandDefinition[] = [];
  private watcher?: Deno.FsWatcher;

  private constructor() {
    this.baseCommands = normalizeCommands(baseCommands as SheetSlashCommand[]);

    if (isWatchMode) {
      this.setupHotReload();
    }
  }

  private setupHotReload() {
    const commandsFilePath = join(dataDir, "slash-commands.json");

    try {
      this.watcher = Deno.watchFs(commandsFilePath);

      (async () => {
        for await (const event of this.watcher!) {
          if (event.kind === "modify") {
            logger.info(
              "Detected changes to slash-commands.json, reloading...",
            );
            await this.reloadBaseCommands();
          }
        }
      })();

      logger.info("Hot-reload enabled for slash-commands.json");
    } catch (error) {
      logger.warn("Failed to setup hot-reload for slash commands", {
        error: (error as Error)?.message,
      });
    }
  }

  private async reloadBaseCommands() {
    try {
      const commandsFilePath = join(dataDir, "slash-commands.json");
      const content = await Deno.readTextFile(commandsFilePath);
      const parsed = JSON.parse(content);
      this.baseCommands = normalizeCommands(parsed as SheetSlashCommand[]);

      // Clear prompt cache to reload any changed prompt files
      promptCache.clear();

      logger.info("Base commands reloaded successfully");
    } catch (error) {
      logger.error("Failed to reload base commands", {
        error: (error as Error)?.message,
      });
    }
  }

  static getInstance(): SlashCommandService {
    if (!this.instance) {
      this.instance = new SlashCommandService();
    }
    return this.instance;
  }

  async getCommands(
    workbookName?: string | null,
  ): Promise<SlashCommandDefinition[]> {
    const sheetCommands = await this.loadCommandsFromSheet(workbookName);
    return this.mergeCommands(sheetCommands);
  }

  async getCommandsSeparated(
    workbookName?: string | null,
  ): Promise<
    {
      baseCommands: SlashCommandDefinition[];
      excelCommands: SlashCommandDefinition[];
    }
  > {
    const sheetCommands = await this.loadCommandsFromSheet(workbookName);
    return {
      baseCommands: this.baseCommands,
      excelCommands: sheetCommands,
    };
  }

  private mergeCommands(
    sheetCommands: SlashCommandDefinition[],
  ): SlashCommandDefinition[] {
    const merged = new Map<string, SlashCommandDefinition>();

    const keyForCommand = (cmd: SlashCommandDefinition) =>
      cmd.id?.toString() || cmd.name?.toString() || cmd.title;

    for (const cmd of this.baseCommands) {
      const key = keyForCommand(cmd);
      if (key) merged.set(key, cmd);
    }

    for (const cmd of sheetCommands) {
      const key = keyForCommand(cmd);
      if (!key) continue;

      const existing = merged.get(key);
      const combined: SlashCommandDefinition = { ...existing, ...cmd };

      // Preserve existing prompt if the incoming command didn't provide one
      if (!combined.prompt && existing?.prompt) {
        combined.prompt = existing.prompt;
      }

      merged.set(key, combined);
    }

    return Array.from(merged.values());
  }

  private async loadCommandsFromSheet(
    workbookName?: string | null,
  ): Promise<SlashCommandDefinition[]> {
    if (!workbookName) return [];

    try {
      const rows = await this.readWorksheetCommands(workbookName);
      return normalizeCommands(rows);
    } catch (error) {
      logger.warn("Failed to load slash commands from sheet", {
        workbookName,
        error: (error as Error)?.message,
      });
      return [];
    }
  }

  private async readWorksheetCommands(
    workbookName: string,
  ): Promise<SheetSlashCommand[]> {
    const client = ExcelApiClient.getInstance();

    // Check if the commands sheet exists before trying to read it
    let metadata: WorkbookMetadata;
    try {
      metadata = await client.getMetadata({ workbookName });
    } catch (error) {
      logger.debug("Failed to get workbook metadata, skipping commands sheet", {
        workbookName,
        error: (error as Error)?.message,
      });
      return [];
    }

    // Check if the worksheet exists
    const worksheet = metadata.sheets?.find(
      (sheet: { name: string }) => sheet.name === WORKSHEET_NAME,
    );
    if (!worksheet) {
      // Sheet doesn't exist - this is ok, just return empty array
      logger.info("Commands sheet not found, skipping", {
        workbookName,
        worksheet: WORKSHEET_NAME,
        availableSheets: metadata.sheets?.map((s: { name: string }) =>
          s.name
        ) ||
          [],
      });
      return [];
    }

    // Get the used range for the worksheet instead of using DEFAULT_RANGE
    let range = DEFAULT_RANGE;
    if (worksheet.usedRangeAddress) {
      range = worksheet.usedRangeAddress;
    } else if (
      metadata.usedRange && metadata.activeSheet === WORKSHEET_NAME
    ) {
      // Fall back to active sheet's used range if it matches
      range = metadata.usedRange;
    }

    logger.info("Using range", { range });

    const request: ReadValuesBatchRequest = {
      operations: [{
        workbookName,
        worksheet: WORKSHEET_NAME,
        range,
      }],
    };

    const response = await client.executeExcelAction<
      ReadValuesBatchRequest,
      ReadValuesBatchResponse
    >(
      ToolNames.READ_VALUES_BATCH,
      request,
    );

    const firstResult = response?.results?.[0];
    const parsedCells = this.normalizeCellValues(firstResult?.cellValues);
    if (parsedCells.length === 0) return [];

    const valueMap = new Map<string, unknown>();
    let minRow = Number.POSITIVE_INFINITY;
    let maxRow = 0;

    for (const cell of parsedCells) {
      const key = `${cell.row}:${cell.column}`;
      valueMap.set(key, cell.value);
      minRow = Math.min(minRow, cell.row);
      maxRow = Math.max(maxRow, cell.row);
    }

    const headerRow = minRow === Number.POSITIVE_INFINITY ? 1 : minRow;
    const headerLookup = new Map<number, keyof SheetSlashCommand>();

    for (const cell of parsedCells) {
      if (cell.row !== headerRow) continue;
      const normalized = normalizeHeader(cell.value);
      if (!normalized) continue;

      const property = headerAliases[normalized];
      if (!property) continue;

      if (!headerLookup.has(cell.column)) {
        headerLookup.set(cell.column, property);
      }
    }

    if (headerLookup.size === 0) return [];

    const rows: SheetSlashCommand[] = [];
    for (let rowIndex = headerRow + 1; rowIndex <= maxRow; rowIndex++) {
      const rowData: SheetSlashCommand = {};
      let hasData = false;

      for (const [column, property] of headerLookup.entries()) {
        const key = `${rowIndex}:${column}`;
        if (!valueMap.has(key)) continue;

        const rawValue = valueMap.get(key);
        if (rawValue === null || rawValue === undefined) continue;

        const stringified = rawValue?.toString();
        if (!stringified || !stringified.trim()) continue;

        rowData[property] = rawValue as any;
        hasData = true;
      }

      if (hasData) {
        rows.push(rowData);
      }
    }

    return rows;
  }

  private normalizeCellValues(cellValues: unknown): ParsedCell[] {
    if (!cellValues || typeof cellValues !== "object") return [];

    let entries: Array<[string, RawCellValue]> = [];

    if (cellValues instanceof Map) {
      entries = Array.from(cellValues.entries());
    } else if (Array.isArray(cellValues)) {
      entries = (cellValues as RawCellValue[][]).filter((entry) =>
        Array.isArray(entry) && typeof entry[0] === "string"
      ) as Array<[string, RawCellValue]>;
    } else {
      entries = Object.entries(cellValues as Record<string, RawCellValue>);
    }

    const parsed: ParsedCell[] = [];
    for (const [address, cell] of entries) {
      const coords = parseCellAddress(address);
      if (!coords) continue;

      const value =
        (cell && typeof cell === "object" && "value" in (cell as any))
          ? (cell as any).value
          : cell;

      parsed.push({
        row: coords.row,
        column: coords.column,
        value,
      });
    }

    return parsed;
  }
}

export const slashCommandService = SlashCommandService.getInstance();
