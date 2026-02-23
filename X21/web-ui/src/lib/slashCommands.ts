import { webViewBridge } from "@/services/webViewBridge";
import { fetchSlashCommandsSeparated } from "@/services/slashCommandService";
import type { SlashCommandDefinition } from "@/types/slash-commands";

type ApiSlashCommand = Partial<SlashCommandDefinition> & {
  requiresInput?: boolean | string | number;
  keywords?: string | string[];
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
    const keywords = value
      .map((k) => k?.toString().trim())
      .filter(Boolean) as string[];
    return keywords.length > 0 ? keywords : undefined;
  }

  if (typeof value === "string") {
    const keywords = value
      .split(/[,;]+/)
      .map((k) => k.trim())
      .filter(Boolean);
    return keywords.length > 0 ? keywords : undefined;
  }

  return undefined;
};

const normalizeSlashCommand = (
  row: ApiSlashCommand,
): SlashCommandDefinition | null => {
  const id = (row.id ?? row.name)?.toString().trim();
  if (!id) return null;

  const name = row.name?.toString().trim() || id;
  const title = row.title?.toString().trim() || name;

  const prompt = row.prompt?.toString() || "";

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
  rows: ApiSlashCommand[],
): SlashCommandDefinition[] => {
  if (!rows || rows.length === 0) return [];

  return rows
    .map(normalizeSlashCommand)
    .filter((cmd): cmd is SlashCommandDefinition => Boolean(cmd));
};

let slashCommandList: SlashCommandDefinition[] = [];
let slashCommandLookup = new Map<string, SlashCommandDefinition>();
const listeners = new Set<() => void>();
let inFlightLoad: Promise<void> | null = null;
let baseCommands: SlashCommandDefinition[] = [];
let excelCommands: SlashCommandDefinition[] = [];

const notifyListeners = () => {
  listeners.forEach((listener) => {
    try {
      listener();
    } catch (error) {
      console.warn("slashCommands listener error", error);
    }
  });
};

const rebuildCaches = (newExcelCommands: SlashCommandDefinition[]) => {
  excelCommands = newExcelCommands;

  const keyForCommand = (cmd: SlashCommandDefinition) =>
    cmd.id?.toString() || cmd.name?.toString() || cmd.title;

  const merged = new Map<string, SlashCommandDefinition>();

  // First add base commands
  for (const cmd of baseCommands) {
    const key = keyForCommand(cmd);
    if (key) merged.set(key, cmd);
  }

  // Then add/override with Excel commands (preserving Excel order)
  for (const cmd of excelCommands) {
    const key = keyForCommand(cmd);
    if (key) merged.set(key, { ...merged.get(key), ...cmd });
  }

  // Build final list: Excel commands first (in sheet order), then base commands
  const excelCommandKeys = new Set(
    excelCommands.map((cmd) => keyForCommand(cmd)).filter(Boolean),
  );
  const orderedList: SlashCommandDefinition[] = [];

  // Add Excel commands first (preserving sheet order)
  for (const cmd of excelCommands) {
    const key = keyForCommand(cmd);
    if (key && merged.has(key)) {
      orderedList.push(merged.get(key)!);
    }
  }

  // Then add base commands (that aren't overridden by Excel)
  for (const cmd of baseCommands) {
    const key = keyForCommand(cmd);
    if (key && !excelCommandKeys.has(key) && merged.has(key)) {
      orderedList.push(merged.get(key)!);
    }
  }

  slashCommandList = orderedList;
  slashCommandLookup = new Map(
    slashCommandList.map((command) => [
      command.id ?? command.name ?? command.title,
      command,
    ]),
  );
  notifyListeners();
};

const loadSlashCommandsFromApi = async () => {
  try {
    const workbookName =
      (await webViewBridge.getWorkbookPath()) ||
      (await webViewBridge.getWorkbookName());
    const { baseCommands: apiBaseCommands, excelCommands: apiExcelCommands } =
      await fetchSlashCommandsSeparated(workbookName);

    // Normalize and store base commands
    baseCommands = normalizeCommands(apiBaseCommands as ApiSlashCommand[]);

    // Normalize and store Excel commands (preserving order from sheet)
    const normalizedExcelCommands = normalizeCommands(
      apiExcelCommands as ApiSlashCommand[],
    );

    rebuildCaches(normalizedExcelCommands);
  } catch (error) {
    // On error (e.g., network issue), clear excelCommands but preserve baseCommands
    // This ensures the app continues with base commands if they were previously loaded
    console.warn("Failed to load slash commands from API:", error);
    // Don't clear baseCommands on error - preserve what we have
    // Only clear excelCommands (set to empty array)
    rebuildCaches([]);
  }
};

export const refreshExcelSlashCommands = async () => {
  if (inFlightLoad) {
    return inFlightLoad;
  }

  inFlightLoad = loadSlashCommandsFromApi().finally(() => {
    inFlightLoad = null;
  });

  return inFlightLoad;
};

export const subscribeToSlashCommands = (
  listener: () => void,
): (() => void) => {
  listeners.add(listener);
  return () => listeners.delete(listener);
};

export const getExcelSlashCommands = (): SlashCommandDefinition[] =>
  excelCommands;
export const getBaseSlashCommands = (): SlashCommandDefinition[] =>
  baseCommands;

// Check if a command is from Excel sheet (custom) vs hardcoded
export const isExcelCommand = (command: SlashCommandDefinition): boolean => {
  const key =
    command.id?.toString() || command.name?.toString() || command.title;
  if (!key) return false;
  return excelCommands.some((cmd) => {
    const cmdKey = cmd.id?.toString() || cmd.name?.toString() || cmd.title;
    return cmdKey === key;
  });
};

export const findSlashCommandById = (
  id?: string | null,
): SlashCommandDefinition | undefined => {
  if (!id) return undefined;
  return slashCommandLookup.get(id);
};

export const searchSlashCommands = (
  query: string,
): SlashCommandDefinition[] => {
  if (!query) {
    return slashCommandList;
  }

  const normalized = query.trim().toLowerCase();
  if (!normalized) {
    return slashCommandList;
  }

  return slashCommandList.filter((command) => {
    const haystack = [
      command.id,
      command.title,
      command.description,
      command.category ?? "",
      ...(command.keywords ?? []),
    ]
      .join(" ")
      .toLowerCase();

    return haystack.includes(normalized);
  });
};
