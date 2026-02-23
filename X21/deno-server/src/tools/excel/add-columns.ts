import { addColumns } from "../../excel-actions/add-columns.ts";
import { AddColumnsRequest, Tool, ToolNames } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("AddColumnsTool");

export class AddColumnsTool implements Tool {
  name = ToolNames.ADD_COLUMNS;
  description =
    "Add columns to an Excel worksheet using column range specification.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["worksheet", "columnRange"],
    properties: {
      worksheet: { type: "string", description: "The worksheet name" },
      columnRange: {
        type: "string",
        description: "The column range to add e.g. `A:A` or `A:C` or `B:E`",
      },
    },
  };

  async execute(params: AddColumnsRequest): Promise<any> {
    logger.info(
      "add_columns called with params:",
      params,
      "workbookName:",
      params.workbookName,
    );
    return await addColumns(params);
  }
}
