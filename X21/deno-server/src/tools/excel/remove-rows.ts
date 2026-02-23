import { readValues } from "../../excel-actions/read-values.ts";
import { removeRows } from "../../excel-actions/remove-rows.ts";
import {
  ReadValuesRequest,
  RemoveRowsRequest,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("RemoveRowsTool");

export class RemoveRowsTool implements Tool {
  name = ToolNames.REMOVE_ROWS;
  description = "Delete rows from an Excel worksheet.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["worksheet", "rowRange"],
    properties: {
      worksheet: { type: "string", description: "The worksheet name" },
      rowRange: {
        type: "string",
        description: "The row range to remove e.g. `1:1` or `2:7` or `9:9`",
      },
    },
  };

  async execute(params: RemoveRowsRequest): Promise<any> {
    logger.info(
      "remove_rows called with params:",
      params,
      "workbookName:",
      params.workbookName,
    );

    const readValuesParams: ReadValuesRequest = {
      worksheet: params.worksheet,
      range: params.rowRange,
      workbookName: params.workbookName,
    };

    logger.info("Reading old values");
    const oldValues = await readValues(readValuesParams);

    logger.info("old values", { oldValues });

    const result = await removeRows(params);
    return {
      oldValues,
      result,
    };
  }
}
