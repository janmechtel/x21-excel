import {
  basename,
  dirname,
  extname,
  isAbsolute,
  join,
  normalize,
} from "@std/path";
import { existsSync } from "node:fs";
import { MergeFilesRequest, Tool, ToolNames } from "../../types/index.ts";
import { getExcelApiBaseUrl } from "../../utils/excel-api-config.ts";
import { createLogger } from "../../utils/logger.ts";
import { WebSocketManager } from "../../services/websocket-manager.ts";

const logger = createLogger("MergeFilesTool");

const normalizeExtensions = (extensions?: string[]) => {
  if (!extensions || extensions.length === 0) {
    return [".xlsx", ".xlsm", ".xls", ".xlsb"];
  }
  const normalized = extensions
    .map((ext) => ext?.trim())
    .filter(Boolean)
    .map((ext) =>
      ext.startsWith(".") ? ext.toLowerCase() : `.${ext.toLowerCase()}`
    );
  return normalized.length > 0
    ? normalized
    : [".xlsx", ".xlsm", ".xls", ".xlsb"];
};

const resolveAndValidateFolderPath = async (folderPath: string) => {
  const rawSegments = folderPath.split(/[/\\]+/).filter(Boolean);
  if (rawSegments.some((segment) => segment === "..")) {
    throw new Error("folderPath must not contain traversal segments (..)");
  }

  const normalized = normalize(folderPath);

  if (!isAbsolute(normalized)) {
    throw new Error("folderPath must be an absolute path");
  }

  return await Deno.realPath(normalized);
};

export class MergeFilesTool implements Tool<MergeFilesRequest> {
  name = ToolNames.MERGE_FILES;
  description =
    "Merge multiple Excel files from a folder into one workbook (one sheet per file). WORKFLOW: Scan folder and merge in the same call. UI: only ask the user to choose a folder (folder_picker); default to opening the merged workbook and include all Excel formats without extra prompts.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["folderPath"],
    properties: {
      folderPath: {
        type: "string",
        description: "Folder containing Excel files to merge",
      },
      outputFileName: {
        type: "string",
        description:
          "Optional output filename (default: MergedWorkbook.xlsx). Only ask if the user wants a custom name.",
      },
      openAfter: {
        type: "boolean",
        description:
          "Open the merged workbook after creation (default: true). Skip prompting unless the user opts out.",
      },
      extensions: {
        type: "array",
        description:
          "Optional list of file extensions to include. Defaults to all Excel formats (xlsx, xlsm, xls, xlsb). Skip prompting unless the user asks for a filter.",
        items: { type: "string" },
      },
    },
  };

  async execute(params: MergeFilesRequest): Promise<any> {
    const rawFolderPath = params.folderPath?.trim();
    if (!rawFolderPath) {
      throw new Error("folderPath is required");
    }
    const folderPath = await resolveAndValidateFolderPath(rawFolderPath);
    const openAfter = params.openAfter !== false;

    logger.info("⚠️ EXECUTING MERGE - File operations starting");

    // Get workbookName from params for status updates
    const workbookName = params.workbookName;
    const wsManager = WebSocketManager.getInstance();

    // Helper to send status updates only if websocket is still connected
    const sendStatus = (
      status: string,
      message: string,
      progress?: { current: number; total: number; unit: string },
      operation?: string,
    ) => {
      if (workbookName && wsManager.isConnected(workbookName)) {
        wsManager.sendStatus(workbookName, status as any, message, progress, {
          toolName: ToolNames.MERGE_FILES,
          operation,
        });
      }
    };

    const includeExts = normalizeExtensions(params.extensions);
    const baseFileName = (
      params.outputFileName?.trim() || "MergedWorkbook.xlsx"
    ).replace(/[\\/:*?"<>|]/g, "");

    const stamp = (() => {
      const d = new Date();
      const pad = (n: number, len = 2) => n.toString().padStart(len, "0");
      return `${d.getFullYear()}${pad(d.getMonth() + 1)}${
        pad(
          d.getDate(),
        )
      }${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    })();

    const resolveOutputPath = () => {
      const ext = extname(baseFileName);
      const nameOnly = basename(baseFileName, ext);
      let candidate = `${nameOnly}_${stamp}${ext}`;
      let fullPath = join(folderPath, candidate);
      let counter = 1;
      while (existsSync(fullPath)) {
        candidate = `${nameOnly}_${stamp}_${counter++}${ext}`;
        fullPath = join(folderPath, candidate);
      }
      return fullPath;
    };

    let outputPath = resolveOutputPath();

    // Collect all Excel files
    sendStatus(
      "processing",
      "Scanning folder for Excel files...",
      undefined,
      "scan",
    );

    const entries: string[] = [];
    const outputFileName = basename(outputPath);
    // Pattern to match merged workbook files (e.g., MergedWorkbook_20260102150634.xlsx)
    const mergedPattern =
      /^MergedWorkbook_\d{14}(_\d+)?\.(xlsx|xlsm|xls|xlsb)$/i;

    for await (const entry of Deno.readDir(folderPath)) {
      if (
        entry.isFile &&
        includeExts.includes(extname(entry.name).toLowerCase())
      ) {
        // Skip temp/lock files, the output file itself, and any previously merged files
        if (
          !entry.name.startsWith("~$") &&
          entry.name !== outputFileName &&
          !mergedPattern.test(entry.name)
        ) {
          entries.push(entry.name);
        }
      }
    }

    if (entries.length === 0) {
      throw new Error(
        `No Excel files found in ${folderPath} matching ${
          includeExts.join(
            ", ",
          )
        }`,
      );
    }

    logger.info(`Found ${entries.length} Excel files to merge`);
    sendStatus(
      "processing",
      `Found ${entries.length} Excel files to merge`,
      {
        current: 0,
        total: entries.length,
        unit: "files",
      },
      "scan_complete",
    );

    // Execute merge
    logger.info("⚠️ Executing merge operation");

    const vstoBaseUrl = getExcelApiBaseUrl();

    // Create target workbook by copying the first file
    sendStatus(
      "processing",
      `Creating merged workbook from ${entries[0]}...`,
      {
        current: 0,
        total: entries.length,
        unit: "files",
      },
      "create_target",
    );

    const firstFilePath = join(folderPath, entries[0]);
    logger.info(`Creating target workbook at: ${outputPath}`);
    let targetOutputPath = outputPath;
    try {
      await Deno.copyFile(firstFilePath, targetOutputPath);
    } catch (error) {
      if (error instanceof Deno.errors.AlreadyExists) {
        logger.warn(
          `Output file already exists (${targetOutputPath}), retrying with a new name`,
        );
        targetOutputPath = resolveOutputPath();
        await Deno.copyFile(firstFilePath, targetOutputPath);
      } else {
        throw error;
      }
    }
    outputPath = targetOutputPath;

    // Open the target workbook (hidden)
    sendStatus(
      "processing",
      "Opening merged workbook...",
      {
        current: 0,
        total: entries.length,
        unit: "files",
      },
      "open_target",
    );
    logger.info(`Calling open_workbook API for: ${outputPath}`);
    const openTargetResponse = await fetch(
      `${vstoBaseUrl}/api/actions/execute`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          action: "open_workbook",
          filePath: outputPath,
          visible: true,
        }),
      },
    );

    logger.info(
      `Open target response received, status: ${openTargetResponse.status}, ok: ${openTargetResponse.ok}`,
    );
    const headers: Record<string, string> = {};
    openTargetResponse.headers.forEach((value, key) => (headers[key] = value));
    logger.info(`Response headers: ${JSON.stringify(headers)}`);

    if (!openTargetResponse.ok) {
      const errorText = await openTargetResponse.text();
      logger.info(`Error response text: ${errorText}`);
      throw new Error(`Failed to open target workbook: ${errorText}`);
    }

    logger.info("About to parse open target response JSON...");
    const responseText = await openTargetResponse.text();
    logger.info(
      `Response text received (length: ${responseText.length}): ${
        responseText.substring(0, 500)
      }`,
    );
    const openTargetResult = JSON.parse(responseText);
    logger.info(
      `Open target result parsed: ${JSON.stringify(openTargetResult)}`,
    );

    if (!openTargetResult.success) {
      throw new Error(
        `Failed to open target workbook: ${openTargetResult.message}`,
      );
    }

    const targetWorkbookName = openTargetResult.workbookName;
    logger.info(`Target workbook opened: ${targetWorkbookName}`);

    const processedFiles: string[] = [];
    const allCopiedSheets: string[] = [];

    // Get sheet names from the first file (which is now the target workbook)
    logger.info(
      `Fetching metadata for target workbook ${targetWorkbookName} (first file)...`,
    );
    const targetMetadataResponse = await fetch(
      `${vstoBaseUrl}/api/getMetadata?workbookName=${
        encodeURIComponent(
          targetWorkbookName,
        )
      }`,
    );

    if (targetMetadataResponse.ok) {
      const targetMetadata = await targetMetadataResponse.json();
      if (!targetMetadata.error) {
        const firstFileSheets = targetMetadata.sheets?.map((s: any) =>
          s.name
        ) || [];
        allCopiedSheets.push(...firstFileSheets);
        logger.info(`First file sheets tracked: ${firstFileSheets.join(", ")}`);
      }
    }

    processedFiles.push(entries[0]);

    // Process each file
    for (let i = 0; i < entries.length; i++) {
      const fileName = entries[i];
      const fullPath = join(folderPath, fileName);

      logger.info(`Processing ${i + 1}/${entries.length}: ${fileName}`);
      sendStatus(
        "processing",
        `Processing file ${i + 1}/${entries.length}: ${fileName}`,
        { current: i, total: entries.length, unit: "files" },
        "process_file",
      );

      try {
        // For the first file, we already have it as the target, skip opening and copying
        if (i === 0) {
          logger.info(
            `First file is already the base for merged workbook, sheets already tracked`,
          );
          continue;
        }

        // Open source workbook
        logger.info(`Opening source workbook: ${fullPath}`);
        const openSrcResponse = await fetch(
          `${vstoBaseUrl}/api/actions/execute`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "open_workbook",
              filePath: fullPath,
              visible: true,
            }),
          },
        );

        logger.info(`Open source response status: ${openSrcResponse.status}`);
        if (!openSrcResponse.ok) {
          logger.warn(`Failed to open ${fileName}, skipping`);
          continue;
        }

        logger.info(`Parsing open source response...`);
        const openSrcResult = await openSrcResponse.json();
        logger.info(`Open source result: ${JSON.stringify(openSrcResult)}`);
        if (!openSrcResult.success) {
          logger.warn(`Failed to open ${fileName}: ${openSrcResult.message}`);
          continue;
        }

        const sourceWorkbookName = openSrcResult.workbookName;
        logger.info(`Opened source workbook: ${sourceWorkbookName}`);

        // Get sheet names from metadata
        logger.info(`Fetching metadata for ${sourceWorkbookName}...`);
        const metadataResponse = await fetch(
          `${vstoBaseUrl}/api/getMetadata?workbookName=${
            encodeURIComponent(
              sourceWorkbookName,
            )
          }`,
        );

        logger.info(`Metadata response status: ${metadataResponse.status}`);
        if (!metadataResponse.ok) {
          const errorText = await metadataResponse.text();
          logger.warn(
            `Failed to get metadata for ${sourceWorkbookName}: ${errorText}`,
          );
          continue;
        }

        logger.info(`Parsing metadata response...`);
        const metadata = await metadataResponse.json();
        logger.info(`Metadata: ${JSON.stringify(metadata)}`);

        if (metadata.error) {
          logger.warn(
            `Metadata error for ${sourceWorkbookName}: ${metadata.error}`,
          );
          continue;
        }

        const sheetNames = metadata.sheets?.map((s: any) => s.name) || [];

        if (sheetNames.length === 0) {
          logger.warn(`No sheets found in ${fileName}, skipping`);
          continue;
        }

        logger.info(
          `Found ${sheetNames.length} sheets: ${sheetNames.join(", ")}`,
        );
        sendStatus(
          "processing",
          `Copying ${sheetNames.length} sheets from ${fileName}...`,
          { current: i, total: entries.length, unit: "files" },
          "copy_sheets",
        );

        // Copy all sheets from source to target with filename-based prefix
        const fileNameWithoutExt = basename(fileName, extname(fileName));
        const prefix = `${fileNameWithoutExt}_`;
        logger.info(
          `Calling copy_sheets API: source=${sourceWorkbookName}, target=${targetWorkbookName}, sheets=${
            sheetNames.join(
              ", ",
            )
          }, prefix=${prefix}`,
        );
        const copyResponse = await fetch(`${vstoBaseUrl}/api/actions/execute`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            action: "copy_sheets",
            sourceWorkbookName,
            targetWorkbookName,
            sheetNames,
            namePrefix: prefix,
          }),
        });

        logger.info(`Copy response status: ${copyResponse.status}`);
        if (!copyResponse.ok) {
          const errorText = await copyResponse.text();
          logger.warn(`Failed to copy sheets from ${fileName}: ${errorText}`);
          continue;
        }

        logger.info(`Parsing copy response...`);
        const copyResult = await copyResponse.json();
        logger.info(`Copy result: ${JSON.stringify(copyResult)}`);
        if (copyResult.success) {
          allCopiedSheets.push(...copyResult.copiedSheets);
          processedFiles.push(fileName);
          logger.info(
            `✅ Copied ${copyResult.copiedSheets.length} sheets: ${
              copyResult.copiedSheets.join(", ")
            }`,
          );
        } else {
          logger.warn(`Copy failed for ${fileName}: ${copyResult.message}`);
        }

        // Close the source workbook without saving
        try {
          logger.info(`Closing source workbook: ${sourceWorkbookName}`);
          const closeResponse = await fetch(
            `${vstoBaseUrl}/api/actions/execute`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                action: "close_workbook",
                workbookName: sourceWorkbookName,
                saveChanges: false,
              }),
            },
          );

          logger.info(`Close response status: ${closeResponse.status}`);
          if (closeResponse.ok) {
            const closeResult = await closeResponse.json();
            logger.info(`Close result: ${JSON.stringify(closeResult)}`);
            if (closeResult.success) {
              logger.info(`✅ Closed source workbook: ${sourceWorkbookName}`);
            } else {
              logger.warn(
                `Close reported unsuccessful: ${closeResult.message}`,
              );
            }
          } else {
            const errorText = await closeResponse.text();
            logger.warn(`Close response not OK: ${errorText}`);
          }
        } catch (closeError) {
          logger.warn(
            `Failed to close ${sourceWorkbookName}: ${
              (closeError as Error).message
            }`,
          );
        }
      } catch (error) {
        logger.warn(
          `Error processing ${fileName}: ${(error as Error).message}`,
        );
        continue;
      }
    }

    if (processedFiles.length === 0) {
      throw new Error(
        "No files were merged. Ensure the files can be opened and contain sheets.",
      );
    }

    // Save and close the target workbook
    sendStatus(
      "processing",
      "Saving merged workbook...",
      {
        current: entries.length,
        total: entries.length,
        unit: "files",
      },
      "save_target",
    );

    let targetClosed = false;
    try {
      logger.info(`Saving and closing merged workbook: ${targetWorkbookName}`);

      const closeTargetResponse = await fetch(
        `${vstoBaseUrl}/api/actions/execute`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            action: "close_workbook",
            workbookName: targetWorkbookName,
            saveChanges: true,
          }),
        },
      );

      logger.info(
        `Close target response status: ${closeTargetResponse.status}`,
      );
      if (closeTargetResponse.ok) {
        const closeResult = await closeTargetResponse.json();
        logger.info(`Close target result: ${JSON.stringify(closeResult)}`);
        if (closeResult.success) {
          logger.info(`✅ Saved and closed merged workbook`);
          targetClosed = true;
        } else {
          logger.warn(
            `Failed to close target workbook: ${closeResult.message}`,
          );
        }
      } else {
        const errorText = await closeTargetResponse.text();
        logger.warn(`Close target response not OK: ${errorText}`);
      }
    } catch (closeError) {
      logger.warn(
        `Error closing target workbook: ${(closeError as Error).message}`,
      );
    }

    if (openAfter && targetClosed) {
      try {
        sendStatus(
          "processing",
          "Opening merged workbook for review...",
          { current: entries.length, total: entries.length, unit: "files" },
          "open_final",
        );

        logger.info(`Opening merged workbook for user: ${outputPath}`);
        const reopenResponse = await fetch(
          `${vstoBaseUrl}/api/actions/execute`,
          {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              action: "open_workbook",
              filePath: outputPath,
              visible: true,
            }),
          },
        );

        logger.info(
          `Open merged workbook response status: ${reopenResponse.status}`,
        );
        if (!reopenResponse.ok) {
          const errorText = await reopenResponse.text();
          logger.warn(`Failed to open merged workbook for user: ${errorText}`);
        }
      } catch (openError) {
        logger.warn(
          `Error opening merged workbook for user: ${
            (openError as Error).message
          }`,
        );
      }
    }

    logger.info(
      `✅ Successfully merged ${processedFiles.length} files with ${allCopiedSheets.length} sheets`,
    );
    sendStatus(
      "idle",
      `✅ Successfully merged ${processedFiles.length} files with ${allCopiedSheets.length} sheets`,
      { current: entries.length, total: entries.length, unit: "files" },
      "complete",
    );

    const resultFileName = basename(outputPath);
    const resultDir = dirname(outputPath);

    return {
      outputPath,
      outputFileName: resultFileName,
      outputDir: resultDir,
      fileCount: processedFiles.length,
      sheetCount: allCopiedSheets.length,
      filesMerged: processedFiles,
      copiedSheets: allCopiedSheets,
    };
  }
}
