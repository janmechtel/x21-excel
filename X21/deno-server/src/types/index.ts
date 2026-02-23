// Re-export shared constants and types from the shared location
export {
  ClaudeContentTypes,
  ClaudeEventTypes,
  ClaudeStopReasons,
  ContentBlockTypes,
  ExcelInteropToolNames,
  OperationStatusValues,
  RevertOperationKeys,
  ToolNames,
  WebSocketMessageTypes,
  WorkbookResolutionPaths,
} from "@shared/types";

export type { OperationStatus, WorkbookResolutionPath } from "@shared/types";

// Import for local use
import type { OperationStatus } from "@shared/types";

// WebSocket message types for status updates
export interface StatusUpdateMessage {
  type: "status:update";
  payload: {
    status: OperationStatus;
    message?: string;
    progress?: {
      current: number;
      total: number;
      unit?: string;
    };
    metadata?: {
      operation?: string;
      range?: string;
      toolName?: string;
      estimatedMs?: number;
    };
  };
}

export interface TokenUpdateMessage {
  type: "stream:token_update";
  payload: {
    inputTokens: number;
    outputTokens: number;
    totalTokens: number;
  };
}

import type { UiRequestPayload } from "./ui-request.ts";

// JSON Schema type for tool parameters (provider-agnostic).
// Compatible with both Anthropic and OpenAI function calling formats.
// Anthropic-supported schema features by 01.01.2026:
// - Basic types: object, array, string, integer, number, boolean, null
// - enum (strings, numbers, booleans, or null only; no complex types)
// - const
// - anyOf and allOf (allOf with $ref not supported)
// - $ref, $defs, and definitions (external $ref not supported)
// - default for all supported types
// - required and additionalProperties (objects must set additionalProperties: false)
// - String formats: date-time, time, date, duration, email, hostname, uri, ipv4,
//   ipv6, uuid
// - Array minItems (only values 0 and 1 supported)
export interface ToolParametersSchema {
  type?: string;
  properties?: Record<
    string,
    {
      type?: string;
      description?: string;
      additionalProperties?: boolean;
      enum?: readonly any[];
      items?: ToolParametersSchema;
      oneOf?: ToolParametersSchema[];
      anyOf?: ToolParametersSchema[];
      allOf?: ToolParametersSchema[];
      [key: string]: any; // Allow additional JSON Schema properties
    }
  >;
  required?: string[] | null;
  additionalProperties?: boolean | ToolParametersSchema;
  [key: string]: any; // Allow additional JSON Schema properties
}

export interface Tool<T extends ToolExecutionRequest = ToolExecutionRequest> {
  name: string;
  description: string;
  strict?: boolean;
  input_schema: ToolParametersSchema;
  execute(params: T): Promise<any>;
}

interface StandardExcelData {
  workbookName?: string;
  worksheet: string;
}

export interface ExcelContext extends StandardExcelData {
  selectedRange: string;
  allSheets: string[];
  usedRange?: string;
  isActiveSheetEmpty?: boolean;
  languageCode?: string;
  otherWorkbooks?: OtherWorkbookContext[];
  dateLanguage?: string;
  listSeparator?: string;
  decimalSeparator?: string;
  thousandsSeparator?: string;
}

export interface OtherWorkbookContext {
  workbookName: string;
  workbookFullName?: string;
  sheets: string[];
}

export type ToolExecutionRequest =
  | ReadValuesRequest
  | ReadValuesBatchRequest
  | WriteValuesRequest
  | WriteValuesBatchRequest
  | GetMetadataRequest
  | ListSheetsRequest
  | ListOpenWorkbooksRequest
  | DragToolRequest
  | ReadFormatRequest
  | ReadFormatBatchRequest
  | CopyPasteRequest
  | AddSheetsRequest
  | RemoveSheetsRequest
  | RemoveColumnsRequest
  | AddColumnsRequest
  | RemoveRowsRequest
  | AddRowsRequest
  | VBAToolRequest
  | VBAReadRequest
  | VBAUpdateRequest
  | WriteFormatBatchRequest
  | UiRequestPayload
  | MergeFilesRequest
  | WorkbookChangelogRequest;

export interface ReadValuesRequest {
  worksheet: string;
  range: string;
  workbookName?: string;
}

export interface ReadValuesResponse {
  cellValues: Map<string, CellValue>;
}

export interface ReadValuesBatchOperation {
  worksheet: string;
  range: string;
  workbookName?: string;
}

export interface ReadValuesBatchRequest {
  operations: ReadValuesBatchOperation[];
}

export interface ReadValuesBatchResult extends ReadValuesResponse {
  success: boolean;
  message: string;
  worksheet?: string;
  workbookName?: string;
  range?: string;
}

export interface ReadValuesBatchResponse {
  success: boolean;
  message: string;
  results: ReadValuesBatchResult[];
}

export interface GetMetadataRequest {
  workbookName?: string;
}

export interface ListSheetsRequest {
  workbookName: string;
}

export type ListOpenWorkbooksRequest = Record<string, never>;

export interface CellValue {
  value: string;
  formula: string;
}

export type ColumnWidthMode = "always" | "smart" | "never";

export interface WriteValuesRequest {
  workbookName?: string;
  worksheet: string;
  range: string;
  values: string[][];
  columnWidthMode?: ColumnWidthMode;
}

export interface WriteValuesBatchOperation {
  workbookName?: string;
  requestedWorkbook?: string;
  worksheet: string;
  range: string;
  values: string[][];
}

export interface WriteNumberFormatOperation {
  workbookName?: string;
  worksheet: string;
  range: string;
  numberFormat?: string;
}

export interface WriteValuesBatchRequest {
  operations: WriteValuesBatchOperation[];
  activeWorkbookName?: string;
  formats?: WriteNumberFormatOperation[];
  columnWidthMode?: ColumnWidthMode;
}

export interface WriteValuesResponse {
  success: boolean;
  message: string;
}

export type PolicyAction = "wrote" | "blocked" | "redirected" | "partial";

export interface WriteValuesBatchOperationResult {
  worksheet: string;
  range: string;
  requestedWorkbook?: string;
  resolvedWorkbook?: string | null;
  policyAction: PolicyAction;
  status: "success" | "error";
  oldValues?: any[][] | null;
  oldValuesOmittedReason?: string;
  newValues?: any[][] | null;
  errorCode?: string;
  errorMessage?: string;
}

export interface WriteValuesBatchResponse {
  success: boolean;
  message: string;
  tool?: string;
  activeWorkbook?: string;
  requestedWorkbook?: string;
  resolvedWorkbook?: string | null;
  policyAction: PolicyAction;
  operations: WriteValuesBatchOperationResult[];
  results?: WriteValuesResponse[];
  applied?: number;
  batches?: number;
  timingMs?: number;
  oldValues?: ReadValuesResponse[];
  oldFormats?: ReadFormatFinalResponseList[];
}

export interface DragToolRequest {
  workbookName?: string;
  worksheet: string;
  sourceRange: string;
  destinationRange: string;
  fillType: string;
}

export interface CopyPasteRequest {
  workbookName?: string;
  sourceWorkbookName?: string;
  sourceWorksheet: string;
  sourceRange: string;
  destinationWorksheet: string;
  destinationRange: string;
  pasteType?:
    | "all"
    | "values"
    | "formats"
    | "formulas"
    | "formulas_and_number_formats"
    | "values_and_number_formats"
    | "column_widths";
  operation?: "add" | "subtract" | "multiply" | "divide";
  skipBlanks?: boolean;
  transpose?: boolean;
  insertMode?: "none" | "shift_right" | "shift_down";
  includeColumnWidths?: boolean;
}

export interface CopyPasteResponse {
  success: boolean;
  message: string;
  workbookName?: string;
  sourceWorksheet?: string;
  sourceRange?: string;
  destinationWorksheet?: string;
  destinationRange?: string;
  pasteType?: string;
  insertMode?: string;
  rowsCopied?: number;
  columnsCopied?: number;
  oldValues?: ReadValuesResponse;
  oldFormats?: ReadFormatFinalResponseList;
  warnings?: string[];
}

export interface ReadFormatRequest {
  worksheet: string;
  range: string;
  workbookName?: string;
  propertiesToRead?: string[];
}

export interface ReadFormatFinalResponse {
  format: FormatSettings;
  ranges: string[];
}

export type ReadFormatFinalResponseList = ReadFormatFinalResponse[];

export interface ReadFormatRawResponse {
  cellFormats: Record<string, FormatSettings>;
}

export interface ReadFormatBatchOperation {
  worksheet: string;
  workbookName?: string;
  range: string;
  propertiesToRead?: string[];
}

export interface ReadFormatBatchRequest {
  operations: ReadFormatBatchOperation[];
}

export interface ReadFormatRawResponseWithMeta extends ReadFormatRawResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
  range: string;
}

export interface ReadFormatBatchResponse {
  success: boolean;
  message: string;
  results: ReadFormatRawResponseWithMeta[];
}

export interface FormatSettings {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  backgroundColor?: string;
  alignment?: string;
  numberFormat?: string;
  fontSize?: number;
  fontName?: string;
  clearBorders?: boolean;
}

export interface WriteFormatRequest {
  workbookName?: string;
  worksheet: string;
  range: string;
  format: FormatSettings;
}

export interface WriteFormatResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
}

export interface WriteFormatOperation {
  workbookName?: string;
  worksheet: string;
  range: string;
  format: FormatSettings;
}

export interface WriteFormatBatchRequest {
  operations: WriteFormatOperation[];
  readOldFormats?: boolean;
  collapseReadRanges?: boolean;
}

export interface WriteFormatBatchResponse {
  success: boolean;
  message: string;
  results: WriteFormatResponse[];
  applied?: number;
  batches?: number;
}

export interface RangeFormatPair {
  range: string;
  format: FormatSettings;
}

export interface AddSheetsRequest {
  sheetNames: string[];
  workbookName?: string;
}

export interface AddSheetsResponse {
  success: boolean;
  message: string;
  sheetsAdded: string[];
}

export interface RemoveSheetsRequest {
  sheetNames: string[];
  workbookName?: string;
}

export interface RemoveSheetsResponse {
  success: boolean;
  message: string;
  sheetsRemoved: string[];
}

export interface AddColumnsRequest {
  worksheet: string;
  workbookName?: string;
  columnRange: string;
}

export interface AddColumnsResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
  columnsInserted: number;
}

export interface RemoveColumnsRequest {
  worksheet: string;
  workbookName?: string;
  columnRange: string;
}

export interface RemoveColumnsResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
  deletedColumns: string;
}

export interface DeleteCellsRequest {
  worksheet: string;
  workbookName?: string;
  range: string;
  shiftDirection?: "left" | "up";
}

export interface DeleteCellsResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName?: string;
  range: string;
  shiftDirection?: string;
}

export interface AddRowsRequest {
  worksheet: string;
  workbookName?: string;
  rowRange: string;
}

export interface AddRowsResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
  rowsInserted: string;
}

export interface RemoveRowsRequest {
  worksheet: string;
  workbookName?: string;
  rowRange: string;
}

export interface RemoveRowsResponse {
  success: boolean;
  message: string;
  worksheet: string;
  workbookName: string;
  deletedRows: string;
}

export interface MergeFilesRequest {
  workbookName?: string;
  folderPath: string;
  outputFileName?: string;
  openAfter?: boolean;
  extensions?: string[];
}

export interface VBAToolRequest {
  workbookName?: string; // Optional, if not provided use current workbook
  functionName: string;
  vbaCode: string;
}

export interface VBAReadRequest {
  workbookName?: string;
}

export interface VBAReadResponse {
  success: boolean;
  message: string;
  modules: Array<{
    name: string;
    code: string;
  }>;
}

export interface VBAUpdateRequest {
  workbookName?: string;
  moduleName: string;
  vbaCode: string;
}

export interface VBAUpdateResponse {
  success: boolean;
  message: string;
  moduleName: string;
}

export interface SheetMetadata {
  name: string;
  usedRangeAddress?: string;
}

export interface WorkbookMetadata {
  workbookName: string;
  workbookFullName?: string;
  activeSheet?: string;
  usedRange?: string;
  selectedRange?: string;
  allSheets?: string[];
  sheets?: SheetMetadata[];
  languageCode?: string;
}

export interface OpenWorkbook {
  workbookName: string;
  workbookFullName?: string;
  sheets?: string[];
}

export interface WorkbookChangelogRequest {
  workbookName?: string;
  includeDiff?: boolean;
  /**
   * Start timestamp in ms since epoch (required for date filtering).
   */
  startTimestampMs?: number;
  /**
   * End timestamp in ms since epoch (optional, defaults to now).
   */
  endTimestampMs?: number;
}

export interface WorkbookChangelogSummary {
  id: string;
  createdAt: number;
  text: string;
  diffId: string | null;
  unifiedDiff?: string;
}

export interface WorkbookChangelogResponse {
  success: boolean;
  workbookName: string;
  message: string;
  summaries?: WorkbookChangelogSummary[]; // Multiple summaries when date range is specified
  summary?: WorkbookChangelogSummary; // Single summary for backward compatibility (latest only)
  dateRange?: {
    start: number;
    end: number;
  };
}

export * from "./ui-request.ts";
