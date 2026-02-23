import {
  AddColumnsRequest,
  AddColumnsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function addColumns(
  params: AddColumnsRequest,
): Promise<AddColumnsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<AddColumnsRequest, AddColumnsResponse>(
    ToolNames.ADD_COLUMNS,
    params,
  );
}
