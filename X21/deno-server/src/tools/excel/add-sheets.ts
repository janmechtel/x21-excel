import { addSheets } from "../../excel-actions/add-sheets.ts";
import { AddSheetsRequest, Tool, ToolNames } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("AddSheetsTool");

export class AddSheetsTool implements Tool {
  name = ToolNames.ADD_SHEETS;
  description = "Add new worksheets to an Excel workbook.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      sheetNames: {
        type: "array",
        description: "Array of sheet names to add",
        items: { type: "string" },
      },
    },
  };

  async execute(params: AddSheetsRequest): Promise<any> {
    logger.info("🔧 add_sheets called with params:", params);
    return await addSheets(params);
  }
}
