import { AddRowsRequest, AddRowsResponse, ToolNames } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function addRows(
  params: AddRowsRequest,
): Promise<AddRowsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<AddRowsRequest, AddRowsResponse>(
    ToolNames.ADD_ROWS,
    params,
  );
}
