import {
  RemoveRowsRequest,
  RemoveRowsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function removeRows(
  params: RemoveRowsRequest,
): Promise<RemoveRowsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<RemoveRowsRequest, RemoveRowsResponse>(
    ToolNames.REMOVE_ROWS,
    params,
  );
}
