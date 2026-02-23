// Shared constants and types between deno-server and web-ui
// This is the single source of truth for common constants

// Status constants for type-safe usage (single source of truth)
export const OperationStatusValues = {
  IDLE: "idle",
  CONNECTING: "connecting",
  READING_EXCEL: "reading_excel",
  READING_EXCEL_FORMAT: "reading_excel_format",
  WRITING_EXCEL: "writing_excel",
  WRITING_EXCEL_FORMAT: "writing_excel_format",
  GENERATING_LLM: "generating_llm",
  EXECUTING_TOOL: "executing_tool",
  WAITING_APPROVAL: "waiting_approval",
  PROCESSING: "processing",
  ERROR: "error",
} as const;

// Derive the type from the constants - no duplication!
export type OperationStatus =
  typeof OperationStatusValues[keyof typeof OperationStatusValues];

// Tool name constants
export const ToolNames = {
  READ_VALUES_BATCH: "read_values_batch",
  WRITE_VALUES_BATCH: "write_values_batch",
  WRITE_FORMAT_BATCH: "write_format_batch",
  DRAG_FORMULA: "drag_formula",
  ADD_SHEETS: "add_sheets",
  REMOVE_SHEETS: "remove_sheets",
  ADD_COLUMNS: "add_columns",
  REMOVE_COLUMNS: "remove_columns",
  ADD_ROWS: "add_rows",
  REMOVE_ROWS: "remove_rows",
  VBA_CREATE: "vba_create",
  VBA_READ: "vba_read",
  VBA_UPDATE: "vba_update",
  COLLECT_INPUT: "collect_input",
  GET_METADATA: "get_metadata",
  LIST_SHEETS: "list_sheets",
  LIST_OPEN_WORKBOOKS: "list_open_workbooks",
  READ_FORMAT_BATCH: "read_format_batch",
  WORKBOOK_CHANGELOG: "workbook_changelog",
  MERGE_FILES: "merge_files",
  COPY_PASTE: "copy_paste",
  DELETE_CELLS: "delete_cells",
} as const;

export type ToolName = typeof ToolNames[keyof typeof ToolNames];

const NonExcelToolNames = [
  ToolNames.COLLECT_INPUT,
] as const;

// Tools that interact with Excel/COM surfaces and should be serialized
export type ExcelInteropToolName = Exclude<ToolName, typeof NonExcelToolNames[number]>;
export const ExcelInteropToolNames: readonly ExcelInteropToolName[] = Object
  .values(ToolNames)
  .filter((name): name is ExcelInteropToolName =>
    !(NonExcelToolNames as readonly string[]).includes(name)
  );

// WebSocket message type constants
export const WebSocketMessageTypes = {
  UI_REQUEST: "ui:request",
  TOOL_PERMISSION: "tool:permission",
  TOOL_AUTO_APPROVED: "tool:auto-approved",
  TOOL_ERROR: "tool:error",
  STREAM_CANCELLED: "stream:cancelled",
  MESSAGE_STOP: "message_stop",
  WORKBOOK_CHANGE_SUMMARY: "workbook:change_summary",
} as const;

// Claude API streaming event type constants
export const ClaudeEventTypes = {
  MESSAGE_START: "message_start",
  MESSAGE_DELTA: "message_delta",
  MESSAGE_STOP: "message_stop",
  CONTENT_BLOCK_START: "content_block_start",
  CONTENT_BLOCK_DELTA: "content_block_delta",
  CONTENT_BLOCK_STOP: "content_block_stop",
} as const;

// Claude API content type constants
export const ClaudeContentTypes = {
  TOOL_USE: "tool_use",
  TOOL_RESULT: "tool_result",
} as const;

// Claude API stop reason constants
export const ClaudeStopReasons = {
  TOOL_USE: "tool_use",
} as const;

// Revert operation key constants (used in revert payloads)
export const RevertOperationKeys = {
  WRITE_VALUES_BATCH: "write_values_batch",
  REMOVE_COLUMNS: "remove_columns",
  REMOVE_ROWS: "remove_rows",
  ADD_ROWS: "add_rows",
  REMOVE_SHEETS: "remove_sheets",
  ADD_COLUMNS: "add_columns",
  WRITE_FORMAT_PREFIX: "write_format_",
} as const;

// Content block type constants
export const ContentBlockTypes = {
  TEXT: "text",
  TOOL_USE: "tool_use",
  THINKING: "thinking",
  UI_REQUEST: "ui_request",
  TOOL_RESULT: "tool_result",
} as const;

// Derive the type from the constants - no duplication!
export type ContentBlockType =
  typeof ContentBlockTypes[keyof typeof ContentBlockTypes];

// Workbook resolution path constants
export const WorkbookResolutionPaths = {
  HOST_ACTIVE: "hostActive",
  SESSION: "session",
  SESSION_OPEN: "sessionOpen",
  PROVIDED_OPEN: "providedOpen",
  SINGLE_OPEN: "singleOpen",
  ERROR: "error",
} as const;

export type WorkbookResolutionPath =
  typeof WorkbookResolutionPaths[keyof typeof WorkbookResolutionPaths];
