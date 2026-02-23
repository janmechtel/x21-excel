import { GetMetadataRequest, WorkbookMetadata } from "../types/index.ts";
import { ExcelApiClient } from "../utils/excel-api-client.ts";

export async function getMetadata(
  params: GetMetadataRequest,
): Promise<WorkbookMetadata> {
  const client = ExcelApiClient.getInstance();
  return await client.getMetadata(params);
}
