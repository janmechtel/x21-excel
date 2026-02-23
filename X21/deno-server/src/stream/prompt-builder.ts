import Anthropic from "@anthropic-ai/sdk";
import { ExcelContext, OtherWorkbookContext } from "../types/index.ts";
import { createLogger } from "../utils/logger.ts";
import { ExcelApiConfigService } from "../utils/excel-api-config.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

const logger = createLogger("PromptBuilder");

export async function createPrompt(
  prompt: string,
  options?: {
    allowOtherWorkbooks?: boolean;
    activeWorkbookName?: string;
    attachments?: Array<{ name: string; type: string; size: number }>;
  },
): Promise<Anthropic.MessageParam> {
  let context: ExcelContext;
  try {
    context = await enrichRequestBody(options);
  } catch (error) {
    logger.error("Failed to fetch Excel context from VSTO:", {
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
    });
    context = {
      selectedRange: "",
      allSheets: [],
      usedRange: "",
      worksheet: "",
      workbookName: "",
      languageCode: "en-US",
      otherWorkbooks: [],
      dateLanguage: "english",
      listSeparator: ",",
      decimalSeparator: ".",
      thousandsSeparator: ",",
    };
  }

  logger.info("Preparing LLM context payload", {
    workbookName: context.workbookName,
    activeSheet: context.worksheet,
    selectedRange: context.selectedRange,
    usedRange: context.usedRange,
    sheetCount: context.allSheets?.length ?? 0,
    otherWorkbooks: context.otherWorkbooks ?? [],
  });

  return buildUserMessageWithContext(prompt, context, options?.attachments);
}

function buildUserMessageWithContext(
  prompt: string,
  context: ExcelContext,
  attachments?: Array<{ name: string; type: string; size: number }>,
): Anthropic.MessageParam {
  const currentDate = new Date().toLocaleDateString("en-US", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  const payload: Record<string, unknown> = {
    userMessage: prompt,
    excelContext: {
      activeSheet: context?.worksheet,
      selectedRange: context?.selectedRange,
      usedRange: context?.usedRange,
      isActiveSheetEmpty: !context?.usedRange ||
        context.usedRange.trim() === "",
      allSheets: context?.allSheets,
      displayLanguage: context?.languageCode || "en-US",
      workbookName: context?.workbookName,
      activeWorkbook: context?.workbookName,
      otherWorkbooks: context?.otherWorkbooks ?? [],
      currentDate,
      dateLanguage: context?.dateLanguage ?? "english",
      listSeparator: context?.listSeparator ?? ",",
      decimalSeparator: context?.decimalSeparator ?? ".",
      thousandsSeparator: context?.thousandsSeparator ?? ",",
    },
  };

  if (attachments && attachments.length > 0) {
    payload.attachments = attachments;
  }

  const contextMessage = JSON.stringify(payload, null, 2);
  logger.info(
    "Constructed user message with context:",
    JSON.parse(contextMessage),
  );
  // log the usedRange separately to avoid overly large logs
  logger.info("Used range in context:", { usedRange: context.usedRange });
  return { role: "user", content: contextMessage };
}

async function enrichRequestBody(
  options?: { allowOtherWorkbooks?: boolean; activeWorkbookName?: string },
): Promise<ExcelContext> {
  const excelApi = ExcelApiConfigService.getInstance();
  const apiUrl = excelApi.getApiUrlMetadata(options?.activeWorkbookName);
  const excelClient = ExcelApiClient.getInstance();

  logger.info("Fetching Excel context from VSTO:", {
    url: apiUrl,
    allowOtherWorkbooks: options?.allowOtherWorkbooks === true,
  });

  const response = await fetch(apiUrl, {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  });

  logger.info("VSTO API response received:", {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
  });

  if (!response.ok) {
    const errorText = await response.text();
    logger.error("VSTO API error response:", {
      status: response.status,
      statusText: response.statusText,
      body: errorText,
    });
    throw new Error(
      `HTTP error! status: ${response.status}, body: ${errorText}`,
    );
  }

  const metadata = await response.json();

  logger.info("Excel context successfully fetched:", {
    workbookName: metadata.workbookName,
    activeSheet: metadata.activeSheet,
    selectedRange: metadata.selectedRange,
    usedRange: metadata.usedRange,
    allSheets: Array.isArray(metadata.allSheets)
      ? metadata.allSheets.length
      : 0,
  });

  let otherWorkbooks: OtherWorkbookContext[] = [];

  if (options?.allowOtherWorkbooks) {
    try {
      const openWorkbooks = await excelClient.listOpenWorkbooks();
      const activeId = options?.activeWorkbookName ||
        metadata?.workbookName ||
        "";
      const activeLower = activeId.toLowerCase();
      const activeBase = activeLower.split(/[/\\]/).pop() || activeLower;

      const results: OtherWorkbookContext[] = [];
      for (const wb of openWorkbooks) {
        const workbookName = wb?.workbookName || "";
        const workbookFullName = wb?.workbookFullName || "";
        const compareId = workbookFullName || workbookName;
        if (!compareId) continue;
        const compareLower = compareId.toLowerCase();
        const compareBase = compareLower.split(/[/\\]/).pop() || compareLower;
        if (compareLower === activeLower || compareBase === activeBase) {
          continue;
        }
        try {
          const wbMetadata = await excelClient.getMetadata({
            workbookName: compareId,
          });
          const sheets = Array.isArray(wbMetadata?.sheets)
            ? wbMetadata.sheets.map((s) => s.name).filter(Boolean)
            : Array.isArray(wbMetadata?.allSheets)
            ? wbMetadata.allSheets
            : [];
          results.push({
            workbookName,
            workbookFullName,
            sheets,
          });
        } catch (innerErr) {
          logger.warn("Failed to fetch metadata for workbook", {
            workbookName: wb.workbookName,
            error: innerErr instanceof Error
              ? innerErr.message
              : String(innerErr),
          });
        }
      }
      otherWorkbooks = results;
    } catch (error) {
      logger.warn("Failed to list open workbooks for context", {
        error: error instanceof Error ? error.message : String(error),
      });
    }
  }

  const context: ExcelContext = {
    selectedRange: metadata.selectedRange,
    allSheets: metadata.allSheets,
    usedRange: metadata.usedRange,
    worksheet: metadata.activeSheet,
    workbookName: metadata.workbookName,
    languageCode: metadata.languageCode,
    otherWorkbooks,
    dateLanguage: metadata.dateLanguage,
    listSeparator: metadata.listSeparator,
    decimalSeparator: metadata.decimalSeparator,
    thousandsSeparator: metadata.thousandsSeparator,
  };

  return context;
}
