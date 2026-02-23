import { ToolNames, VBAReadRequest, VBAReadResponse } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function readVBAModules(
  params: VBAReadRequest,
): Promise<VBAReadResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<VBAReadRequest, VBAReadResponse>(
    ToolNames.VBA_READ,
    params,
  );
}
