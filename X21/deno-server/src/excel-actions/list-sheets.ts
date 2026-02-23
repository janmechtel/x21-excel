import { ListSheetsRequest, SheetMetadata } from "../types/index.ts";
import { getMetadata } from "./get-metadata.ts";

export async function listSheets(
  params: ListSheetsRequest,
): Promise<SheetMetadata[]> {
  const metadata = await getMetadata({ workbookName: params.workbookName });
  return metadata.sheets ?? [];
}
