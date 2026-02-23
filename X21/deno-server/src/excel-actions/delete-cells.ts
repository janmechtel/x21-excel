import {
  DeleteCellsRequest,
  DeleteCellsResponse,
  ToolNames,
} from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function deleteCells(
  params: DeleteCellsRequest,
): Promise<DeleteCellsResponse> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<
    DeleteCellsRequest,
    DeleteCellsResponse
  >(
    ToolNames.DELETE_CELLS,
    params,
  );
}
