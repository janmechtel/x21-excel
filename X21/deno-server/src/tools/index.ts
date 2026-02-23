import { ReadValuesBatchTool } from "./excel/read-values-batch.ts";
import { WriteValuesBatchTool } from "./excel/write-values-batch.ts";
import { GetMetadataTool } from "./excel/get-metadata.ts";
import { ListSheetsTool } from "./excel/list-sheets.ts";
import { ListOpenWorkbooksTool } from "./excel/list-open-workbooks.ts";
import { Tool } from "../types/index.ts";
import { VBATool } from "./excel/vba-create.ts";
import { VBAReadTool } from "./excel/vba-read.ts";
import { VBAUpdateTool } from "./excel/vba-update.ts";
import { DragFormulaTool } from "./excel/drag-formula.ts";
import { AddSheetsTool } from "./excel/add-sheets.ts";
import { RemoveColumnsTool } from "./excel/remove-columns.ts";
import { AddColumnsTool } from "./excel/add-columns.ts";
import { RemoveRowsTool } from "./excel/remove-rows.ts";
import { AddRowsTool } from "./excel/add-rows.ts";
import { ReadFormatBatchTool } from "./excel/read-format-batch.ts";
import { WriteFormatBatchTool } from "./excel/write-format-batch.ts";
import { UiRequestTool } from "./ui-request.ts";
import { MergeFilesTool } from "./excel/merge-files.ts";
import { WorkbookChangelogTool } from "./workbook-changelog.ts";
import { CopyPasteTool } from "./excel/copy-paste.ts";

const toolsBase: Tool[] = [
  new UiRequestTool(),
  new ListOpenWorkbooksTool(),
  new GetMetadataTool(),
  new ListSheetsTool(),
  new WriteValuesBatchTool(),
  new WriteFormatBatchTool(),
  new ReadValuesBatchTool(),
  new ReadFormatBatchTool(),
  new DragFormulaTool(),
  new CopyPasteTool(),
  new AddSheetsTool(),
  new RemoveColumnsTool(),
  new AddColumnsTool(),
  new RemoveRowsTool(),
  new AddRowsTool(),
  new MergeFilesTool(),
  new VBATool(),
  new VBAReadTool(),
  new VBAUpdateTool(),
  new WorkbookChangelogTool(),
];

export const tools: Tool[] = toolsBase;

export function getToolByName(name: string): Tool | undefined {
  return tools.find((tool) => tool.name === name);
}

export {
  AddColumnsTool,
  AddRowsTool,
  AddSheetsTool,
  CopyPasteTool,
  DragFormulaTool,
  GetMetadataTool,
  ListOpenWorkbooksTool,
  ListSheetsTool,
  ReadFormatBatchTool,
  ReadValuesBatchTool,
  RemoveColumnsTool,
  RemoveRowsTool,
  UiRequestTool,
  VBAReadTool,
  VBATool,
  VBAUpdateTool,
  WorkbookChangelogTool,
  WriteFormatBatchTool,
  WriteValuesBatchTool,
};
