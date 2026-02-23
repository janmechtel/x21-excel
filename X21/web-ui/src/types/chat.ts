// Status types for operation visibility
import type { UiRequestPayload, UiRequestResponse } from "./uiRequest";

// Re-export shared constants and types from the shared location
export {
  OperationStatusValues,
  ToolNames,
  WebSocketMessageTypes,
  ClaudeContentTypes,
  ClaudeEventTypes,
  ContentBlockTypes,
} from "@shared/types";

export type { OperationStatus, ContentBlockType } from "@shared/types";

// Import for local use
import type { OperationStatus, ContentBlockType } from "@shared/types";

export interface StatusUpdatePayload {
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
}

export interface TokenUpdatePayload {
  inputTokens: number;
  outputTokens: number;
  totalTokens: number;
}

export interface ChangeSummaryPayload {
  summary?: string;
  id?: string;
  timestamp?: number;
  sheetsAffected?: number;
  comparisonType?: "self" | "external";
  comparisonFilePath?: string | null;
}

export type ToolDecision = "approved" | "rejected";

export interface ToolDecisionData {
  decision: ToolDecision;
  message?: string;
}

export type ToolDecisionList = Map<string, ToolDecisionData>;
export type ToolGroupDecisions = Map<string, ToolDecisionList>;
export type ToolGroups = Map<string, string[]>;

export interface ActivityChange {
  timestamp: string;
  eventType: string;
  workbook: string;
  sheet?: string;
  cell?: string;
  range?: string;
  previousValue?: any;
  newValue?: any;
  previousFormula?: string;
  newFormula?: string;
  cellCount?: number;
  hasFormula?: boolean;
}

export interface ActivitySummary {
  type:
    | "cell_change"
    | "formula_change"
    | "sheet_action"
    | "workbook_action"
    | "other";
  count: number;
  description: string;
  details: ActivityChange[];
  firstOccurrence: string;
  lastOccurrence: string;
  children?: ActivitySummary[]; // For hierarchical structure
  level?: number; // Indentation level (0 = root, 1 = first child, etc.)
}

export interface ActivitySummaryData {
  timeRange: {
    start: string;
    end: string;
  };
  totalEvents: number;
  summaries: ActivitySummary[];
}

export interface ActivityLog {
  id: string;
  timestamp: number;
  data: ActivitySummaryData;
  feedback?: "up" | "down";
  // Original raw summary text as stored in the backend (checkbox list)
  rawSummaryText?: string;
  // Optional comparison context (when the summary was generated vs another file)
  comparisonType?: "self" | "external";
  comparisonFilePath?: string | null;
  comparisonFileName?: string | null;
}

export interface ContentBlock {
  id: string;
  type: ContentBlockType;
  content: string;
  toolName?: string;
  toolUseId?: string;
  toolNumber?: number;
  isComplete: boolean;
  startTime?: number;
  endTime?: number;
  uiRequest?: UiRequestPayload;
  uiRequestResponse?: UiRequestResponse;
  uiRequestSummary?: string;
  activitySummary?: ActivitySummaryData;
}

export interface AttachedFile {
  name: string;
  type: string;
  size: number;
  base64: string;
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  content?: string;
  timestamp: number;
  contentBlocks?: ContentBlock[];
  isStreaming?: boolean;
  score?: "up" | "down" | null;
  slashCommandId?: string;
  attachedFiles?: AttachedFile[];
}

export interface SendMessageOptions {
  overridePrompt?: string;
  attachmentsOverride?: AttachedFile[];
  slashCommandId?: string;
  nextPromptValue?: string;
}
