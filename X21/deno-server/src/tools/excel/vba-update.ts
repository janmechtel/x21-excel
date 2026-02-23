import { updateVBAModule } from "../../excel-actions/vba-update.ts";
import { Tool, ToolNames, VBAUpdateRequest } from "../../types/index.ts";
import { createLogger } from "../../utils/logger.ts";

const logger = createLogger("VBAUpdateTool");

export class VBAUpdateTool implements Tool {
  name = ToolNames.VBA_UPDATE;
  description =
    "Update existing VBA code in a specific module of an Excel workbook.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["workbookName", "moduleName", "vbaCode"],
    properties: {
      workbookName: {
        type: "string",
        description:
          "The name of the workbook containing the VBA module to update",
      },
      moduleName: {
        type: "string",
        description: "The name of the VBA module to update",
      },
      vbaCode: {
        type: "string",
        description:
          "The new VBA code to replace the existing code in the module",
      },
    },
  };

  async execute(params: VBAUpdateRequest): Promise<any> {
    logger.info("vba_update called with params:", params);

    const result = await updateVBAModule(params);

    logger.info("result", { result });

    return result;
  }
}
