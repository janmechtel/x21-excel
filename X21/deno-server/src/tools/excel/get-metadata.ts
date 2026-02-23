import {
  GetMetadataRequest,
  Tool,
  ToolNames,
  WorkbookMetadata,
} from "../../types/index.ts";
import { getMetadata } from "../../excel-actions/get-metadata.ts";

export class GetMetadataTool implements Tool {
  name = ToolNames.GET_METADATA;
  description =
    "Fetch workbook metadata (sheets, used ranges) for planning reads. Defaults to the active workbook if none is provided.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      workbookName: {
        type: "string",
        description:
          "Optional workbook name to target. Defaults to the active workbook.",
      },
    },
  };

  async execute(params: GetMetadataRequest): Promise<WorkbookMetadata> {
    return await getMetadata(params);
  }
}
