import { ToolNames, VBAToolRequest } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function createVBAMacro(params: VBAToolRequest): Promise<any> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<VBAToolRequest, any>(
    ToolNames.VBA_CREATE,
    params,
  );
}
