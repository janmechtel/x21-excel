import { DragToolRequest, ToolNames } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function dragFormula(params: DragToolRequest): Promise<any> {
  const client = ExcelApiClient.getInstance();
  return await client.executeExcelAction<DragToolRequest, any>(
    ToolNames.DRAG_FORMULA,
    params,
  );
}
