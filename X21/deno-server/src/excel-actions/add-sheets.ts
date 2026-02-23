import {
  AddSheetsRequest,
  AddSheetsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function addSheets(
  params: AddSheetsRequest,
): Promise<AddSheetsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<AddSheetsRequest, AddSheetsResponse>(
    ToolNames.ADD_SHEETS,
    params,
  );
}
