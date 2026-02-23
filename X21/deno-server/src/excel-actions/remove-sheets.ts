import {
  RemoveSheetsRequest,
  RemoveSheetsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function removeSheets(
  params: RemoveSheetsRequest,
): Promise<RemoveSheetsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<
    RemoveSheetsRequest,
    RemoveSheetsResponse
  >(
    ToolNames.REMOVE_SHEETS,
    params,
  );
}
