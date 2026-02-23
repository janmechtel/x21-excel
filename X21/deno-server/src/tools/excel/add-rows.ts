import { addRows } from "../../excel-actions/add-rows.ts";
import { AddRowsRequest, Tool, ToolNames } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("AddRowsTool");

export class AddRowsTool implements Tool {
  name = ToolNames.ADD_ROWS;
  description = "Add rows to an Excel worksheet at a specific position.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["worksheet", "rowRange"],
    properties: {
      worksheet: { type: "string", description: "The worksheet name" },
      rowRange: {
        type: "string",
        description: "The row range to add e.g. `1:1` or `2:7` or `9:9`",
      },
    },
  };

  async execute(params: AddRowsRequest): Promise<any> {
    logger.info(
      "add_rows called with params:",
      params,
      "workbookName:",
      params.workbookName,
    );
    return await addRows(params);
  }
}
