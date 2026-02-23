import { ToolNames } from "@/types/chat";

export const toolsNotRequiringApproval: string[] = [
  ToolNames.READ_VALUES_BATCH,
  ToolNames.READ_FORMAT_BATCH,
  ToolNames.GET_METADATA,
  ToolNames.LIST_SHEETS,
  ToolNames.VBA_READ,
  ToolNames.VBA_CREATE,
  ToolNames.COLLECT_INPUT,
  ToolNames.WORKBOOK_CHANGELOG,
  ToolNames.MERGE_FILES,
];
