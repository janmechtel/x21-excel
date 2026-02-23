import { OpenWorkbook } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function listOpenWorkbooks(): Promise<OpenWorkbook[]> {
  const client = ExcelApiClient.getInstance();
  return await client.listOpenWorkbooks();
}
