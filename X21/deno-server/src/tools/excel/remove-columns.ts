import { removeColumns } from "../../excel-actions/remove-columns.ts";
import { readValues } from "../../excel-actions/read-values.ts";
import {
  ReadValuesRequest,
  RemoveColumnsRequest,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("RemoveColumnsTool");

export class RemoveColumnsTool implements Tool {
  name = ToolNames.REMOVE_COLUMNS;
  description = "Remove columns from an Excel worksheet.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["worksheet", "columnRange"],
    properties: {
      worksheet: { type: "string", description: "The worksheet name" },
      columnRange: {
        type: "string",
        description: "The column range to remove e.g. `A:A` or `A:C` or `B:E`",
      },
    },
  };

  async execute(params: RemoveColumnsRequest): Promise<any> {
    logger.info("remove_columns called with params:", params);

    const readValuesParams: ReadValuesRequest = {
      worksheet: params.worksheet,
      range: params.columnRange,
      workbookName: params.workbookName,
    };

    logger.info("Reading old values");
    const oldValues = await readValues(readValuesParams);
    logger.info("old values", { oldValues });

    const result = await removeColumns(params);

    logger.info("result", { result });

    return {
      oldValues,
      result,
    };
  }
}
