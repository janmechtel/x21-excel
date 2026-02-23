import { getApiBase } from "./apiBase";
import type { SlashCommandDefinition } from "@/types/slash-commands";

export async function fetchSlashCommands(
  workbookName?: string | null,
): Promise<SlashCommandDefinition[]> {
  const base = await getApiBase();
  const params = new URLSearchParams();

  if (workbookName) {
    params.set("workbookName", workbookName);
  }

  const query = params.toString();
  const url = `${base}/api/slash-commands${query ? `?${query}` : ""}`;

  const res = await fetch(url);
  if (!res.ok) {
    // If sheet doesn't exist (400/404), return empty array instead of throwing
    // This allows the app to continue with base commands only
    if (res.status === 400 || res.status === 404) {
      return [];
    }
    throw new Error(`Failed to fetch slash commands: ${res.status}`);
  }

  const data = await res.json();
  return Array.isArray(data?.commands) ? data.commands : [];
}

export async function fetchSlashCommandsSeparated(
  workbookName?: string | null,
): Promise<{
  baseCommands: SlashCommandDefinition[];
  excelCommands: SlashCommandDefinition[];
}> {
  const base = await getApiBase();
  const params = new URLSearchParams();

  if (workbookName) {
    params.set("workbookName", workbookName);
  }
  params.set("separated", "true");

  const query = params.toString();
  const url = `${base}/api/slash-commands?${query}`;

  const res = await fetch(url);
  if (!res.ok) {
    // If sheet doesn't exist (400/404), the server should still return base commands
    // But if the request fails entirely, we can't get base commands
    // Return empty arrays - the app will work with whatever was previously loaded
    if (res.status === 400 || res.status === 404) {
      return { baseCommands: [], excelCommands: [] };
    }
    throw new Error(`Failed to fetch slash commands: ${res.status}`);
  }

  const data = await res.json();
  // Server should always return baseCommands, even if sheet doesn't exist
  return {
    baseCommands: Array.isArray(data?.baseCommands) ? data.baseCommands : [],
    excelCommands: Array.isArray(data?.excelCommands) ? data.excelCommands : [],
  };
}
