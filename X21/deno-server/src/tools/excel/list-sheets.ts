import {
  ListSheetsRequest,
  SheetMetadata,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { listSheets } from "../../excel-actions/list-sheets.ts";

export class ListSheetsTool implements Tool {
  name = ToolNames.LIST_SHEETS;
  description =
    "List sheets for a given workbook with basic range hints. Provide the workbook name to avoid ambiguity.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["workbookName"],
    properties: {
      workbookName: {
        type: "string",
        description: "Workbook name to inspect (required).",
      },
    },
  };

  async execute(params: ListSheetsRequest): Promise<SheetMetadata[]> {
    return await listSheets(params);
  }
}
