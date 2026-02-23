import { createVBAMacro } from "../../excel-actions/vba-create.ts";
import { Tool, ToolNames, VBAToolRequest } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("VBATool");

export class VBATool implements Tool {
  name = ToolNames.VBA_CREATE;
  description =
    "Create VBA macros in Excel workbooks with custom function names and VBA code.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["workbookName", "functionName", "vbaCode"],
    properties: {
      workbookName: {
        type: "string",
        description: "The name of the workbook to add the VBA macro to",
      },
      functionName: {
        type: "string",
        description: "The name of the VBA function/subroutine to create",
      },
      vbaCode: {
        type: "string",
        description: "The VBA code to add to the macro module",
      },
    },
  };

  async execute(params: VBAToolRequest): Promise<any> {
    logger.info("vba called with params:", params);

    const result = await createVBAMacro(params);

    logger.info("result", { result });

    return result;
  }
}
