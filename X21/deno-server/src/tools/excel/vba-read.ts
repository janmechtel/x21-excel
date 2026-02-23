import { readVBAModules } from "../../excel-actions/vba-read.ts";
import { Tool, ToolNames, VBAReadRequest } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("VBAReadTool");

export class VBAReadTool implements Tool {
  name = ToolNames.VBA_READ;
  description = "Read VBA code from all modules in an Excel workbook.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["workbookName"],
    properties: {
      workbookName: {
        type: "string",
        description: "The name of the workbook to read VBA modules from",
      },
    },
  };

  async execute(params: VBAReadRequest): Promise<any> {
    logger.info("vba_read called with params:", params);

    const result = await readVBAModules(params);

    logger.info("result", { result });

    return result;
  }
}
