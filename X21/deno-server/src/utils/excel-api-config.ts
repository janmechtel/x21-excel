/// <reference lib="deno.ns" />

import { getEnvironment } from "./environment.ts";
import { createLogger } from "./logger.ts";

const logger = createLogger("ExcelApiConfig");

/**
 * Get the X21 app data directory path
 */
function getX21AppDataDirectory(): string {
  const localAppData = Deno.env.get("LOCALAPPDATA");
  if (!localAppData) {
    throw new Error("LOCALAPPDATA environment variable not found");
  }

  return `${localAppData}\\X21`;
}

/**
 * Read the Excel API port from the port file
 */
export function getExcelApiBaseUrl(): string {
  try {
    const appDataDir = getX21AppDataDirectory();
    const environment = getEnvironment();
    const fileName = `excel-api-port-${environment}`;
    const filePath = `${appDataDir}\\${fileName}`;

    try {
      const portText = Deno.readTextFileSync(filePath).trim();
      const port = parseInt(portText);

      logger.info(`📡 Excel API port: ${port}`);

      if (!isNaN(port)) {
        return `http://localhost:${port}`;
      }
    } catch {
      // File doesn't exist or can't be read
    }

    logger.info("📡 Excel API port not found, using default port 8080");
    // Fallback to default port
    return "http://localhost:8080";
  } catch (error) {
    logger.warn("Failed to read Excel API port file, using default:", error);
    return "http://localhost:8080";
  }
}

/**
 * Service class to manage Excel API configuration
 */
export class ExcelApiConfigService {
  private static instance: ExcelApiConfigService;
  private baseUrl: string;

  private constructor() {
    this.baseUrl = getExcelApiBaseUrl();
  }

  public static getInstance(): ExcelApiConfigService {
    if (!ExcelApiConfigService.instance) {
      ExcelApiConfigService.instance = new ExcelApiConfigService();
    }
    return ExcelApiConfigService.instance;
  }

  public getBaseUrl(): string {
    // Refresh the URL each time in case the port has changed
    this.baseUrl = getExcelApiBaseUrl();
    return this.baseUrl;
  }

  public getApiUrlActionExecution(): string {
    return `${this.getBaseUrl()}/api/actions/execute`;
  }

  public getApiUrlMetadata(workbookName?: string): string {
    const base = `${this.getBaseUrl()}/api/getMetadata`;
    if (!workbookName) return base;
    const encoded = encodeURIComponent(workbookName);
    return `${base}?workbook_name=${encoded}`;
  }

  public getApiUrlOpenWorkbooks(): string {
    return `${this.getBaseUrl()}/api/openWorkbooks`;
  }

  public getApiUrlOpenWorkbooksAll(): string {
    return `${this.getBaseUrl()}/api/openWorkbooksAll`;
  }
}
