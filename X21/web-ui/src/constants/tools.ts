import { ToolNames } from "@/types/chat";

export const AVAILABLE_TOOLS = [
  {
    id: ToolNames.COPY_PASTE,
    name: "Copy Paste",
    description: "Copy and paste data in Excel",
  },
  {
    id: ToolNames.COLLECT_INPUT,
    name: "Collect Input",
    description: "Ask the user for structured input via chat forms",
  },
  {
    id: ToolNames.WRITE_VALUES_BATCH,
    name: "Write Values",
    description: "Write data to multiple Excel ranges in one call",
  },
  {
    id: ToolNames.READ_VALUES_BATCH,
    name: "Read Values",
    description: "Read data from Excel ranges",
  },
  {
    id: ToolNames.WRITE_FORMAT_BATCH,
    name: "Write Format",
    description: "Format multiple Excel ranges in one call",
  },
  {
    id: ToolNames.READ_FORMAT_BATCH,
    name: "Read Format",
    description: "Read data from Excel ranges",
  },
  {
    id: ToolNames.GET_METADATA,
    name: "Get Metadata",
    description: "Fetch workbook metadata like sheets and used ranges",
  },
  {
    id: ToolNames.LIST_SHEETS,
    name: "List Sheets",
    description: "List sheets for a workbook",
  },
  {
    id: ToolNames.LIST_OPEN_WORKBOOKS,
    name: "List Open Workbooks",
    description: "List other open workbooks (requires approval)",
  },
  {
    id: ToolNames.ADD_ROWS,
    name: "Insert Rows",
    description: "Insert rows in Excel",
  },
  {
    id: ToolNames.REMOVE_ROWS,
    name: "Delete Rows",
    description: "Delete rows from Excel",
  },
  {
    id: ToolNames.ADD_COLUMNS,
    name: "Insert Columns",
    description: "Insert columns in Excel",
  },
  {
    id: ToolNames.REMOVE_COLUMNS,
    name: "Delete Columns",
    description: "Delete columns from Excel",
  },
  {
    id: ToolNames.ADD_SHEETS,
    name: "Add Sheets",
    description: "Add new worksheets to the workbook",
  },
  {
    id: ToolNames.VBA_READ,
    name: "VBA Read ",
    description: "Read VBA macros in Excel workbook ",
  },
  {
    id: ToolNames.VBA_CREATE,
    name: "VBA Create (experimental)",
    description:
      "Create a new custom VBA macro in Excel workbooks with specified function names and code",
  },
  {
    id: ToolNames.VBA_UPDATE,
    name: "VBA Update (experimental)",
    description:
      "Update existing VBA code in a specific module of an Excel workbook",
  },
  {
    id: ToolNames.DRAG_FORMULA,
    name: "Drag Formula",
    description:
      "Drag formulas and patterns using Excel AutoFill functionality",
  },
  {
    id: ToolNames.MERGE_FILES,
    name: "Merge Files (experimental)",
    description:
      "Merge multiple Excel files into one workbook (one sheet per file)",
  },
  {
    id: ToolNames.WORKBOOK_CHANGELOG,
    name: "Workbook Changelog (experimental)",
    description:
      "Retrieve the latest or date-filtered changelog summary for this workbook",
  },
];
