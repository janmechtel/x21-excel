import {
  CopyPasteRequest,
  CopyPasteResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function copyPaste(
  params: CopyPasteRequest,
): Promise<CopyPasteResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<CopyPasteRequest, CopyPasteResponse>(
    ToolNames.COPY_PASTE,
    params,
  );
}
