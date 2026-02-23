import {
  ToolNames,
  VBAUpdateRequest,
  VBAUpdateResponse,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function updateVBAModule(
  params: VBAUpdateRequest,
): Promise<VBAUpdateResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<VBAUpdateRequest, VBAUpdateResponse>(
    ToolNames.VBA_UPDATE,
    params,
  );
}
