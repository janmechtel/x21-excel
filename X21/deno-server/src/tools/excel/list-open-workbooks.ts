import {
  ListOpenWorkbooksRequest,
  OpenWorkbook,
  Tool,
  ToolNames,
} from "../../types/index.ts";
import { listOpenWorkbooks } from "../../excel-actions/list-open-workbooks.ts";

export class ListOpenWorkbooksTool implements Tool {
  name = ToolNames.LIST_OPEN_WORKBOOKS;
  description =
    "List all open workbooks in the current Excel session. Requires user approval before executing.";
  input_schema = {
    type: "object",
    additionalProperties: false,
    properties: {},
  };

  async execute(
    _params: ListOpenWorkbooksRequest,
  ): Promise<OpenWorkbook[]> {
    return await listOpenWorkbooks();
  }
}
