import {
  AlertCircle,
  CheckCircle2,
  Database,
  FileSpreadsheet,
  Loader2,
  Sparkles,
  Wrench,
} from "lucide-react";
import type { OperationStatus } from "@/types/chat";
import { OperationStatusValues } from "@/types/chat";

interface EnhancedStatusIndicatorProps {
  status: OperationStatus;
  message?: string | null;
  progress?: {
    current: number;
    total: number;
    unit?: string;
  } | null;
}

export function StatusIndicator({
  status,
  message,
  progress,
}: EnhancedStatusIndicatorProps) {
  // Icon, color, and label based on status
  const getStatusDisplay = () => {
    switch (status) {
      case OperationStatusValues.IDLE:
        return {
          icon: CheckCircle2,
          color: "text-green-500",
          bgColor: "",
          label: "Ready",
          showSpinner: false,
        };
      case OperationStatusValues.CONNECTING:
        return {
          icon: Loader2,
          color: "text-blue-500",
          bgColor: "",
          label: message || "Connecting...",
          showSpinner: true,
        };
      case OperationStatusValues.READING_EXCEL:
        return {
          icon: Database,
          color: "text-blue-500",
          bgColor: "",
          label: message || "Reading from Excel...",
          showSpinner: true,
        };
      case OperationStatusValues.READING_EXCEL_FORMAT:
        return {
          icon: Database,
          color: "text-blue-600",
          bgColor: "",
          label: message || "Reading cell formatting (may take a while)...",
          showSpinner: true,
        };
      case OperationStatusValues.WRITING_EXCEL:
        return {
          icon: FileSpreadsheet,
          color: "text-orange-500",
          bgColor: "",
          label: message || "Writing to Excel...",
          showSpinner: true,
        };
      case OperationStatusValues.WRITING_EXCEL_FORMAT:
        return {
          icon: FileSpreadsheet,
          color: "text-orange-600",
          bgColor: "",
          label: message || "Applying cell formatting (may take a while)...",
          showSpinner: true,
        };
      case OperationStatusValues.GENERATING_LLM:
        return {
          icon: Sparkles,
          color: "text-purple-500",
          bgColor: "",
          label: message || "Generating response...",
          showSpinner: true,
        };
      case OperationStatusValues.EXECUTING_TOOL:
        return {
          icon: Wrench,
          color: "text-indigo-500",
          bgColor: "",
          label: message || "Executing tool...",
          showSpinner: true,
        };
      case OperationStatusValues.WAITING_APPROVAL:
        return {
          icon: AlertCircle,
          color: "text-yellow-500",
          bgColor: "",
          label: message || "Waiting for approval",
          showSpinner: false,
        };
      case OperationStatusValues.ERROR:
        return {
          icon: AlertCircle,
          color: "text-red-500",
          bgColor: "",
          label: message || "Error occurred",
          showSpinner: false,
        };
      case OperationStatusValues.PROCESSING:
      default:
        return {
          icon: Loader2,
          color: "text-slate-500",
          bgColor: "",
          label: message || "Processing...",
          showSpinner: true,
        };
    }
  };

  const display = getStatusDisplay();
  const Icon = display.icon;
  const progressLabel =
    progress && progress.total > 0
      ? ` (${progress.current}/${progress.total} ${progress.unit || "ops"})`
      : "";

  return (
    <div className="">
      <div className="mb-2 max-w-4xl mx-auto space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-4 py-3 bg-slate-50/50 dark:bg-slate-800/50 border border-grey-300">
          {/* Left side: Status with icon and message */}
          <div className="flex items-center gap-3 flex-1 min-w-0">
            <div className="flex-shrink-0 p-1.5 rounded-md">
              {display.showSpinner ? (
                <Loader2
                  className={`h-3.5 w-3.5 animate-spin ${display.color}`}
                />
              ) : (
                <Icon className={`h-3.5 w-3.5 ${display.color}`} />
              )}
            </div>
            <span className="text-sm font-medium text-slate-700 dark:text-slate-300 whitespace-pre-wrap break-words">
              {display.label}
              {progressLabel}
            </span>

            {/* Progress bar (if available) */}
            {progress && progress.total > 0 && (
              <div className="flex items-center gap-2 flex-1 max-w-xs">
                <div className="flex-1 h-1.5 bg-slate-200 dark:bg-slate-700 rounded-full overflow-hidden">
                  <div
                    className={`h-full ${display.color.replace(
                      "text-",
                      "bg-",
                    )} transition-all duration-300`}
                    style={{
                      width: `${Math.min(
                        100,
                        (progress.current / progress.total) * 100,
                      )}%`,
                    }}
                  />
                </div>
                <span className="text-xs text-slate-500 dark:text-slate-400 whitespace-nowrap flex-shrink-0">
                  {progress.current}/{progress.total} {progress.unit || ""}
                </span>
              </div>
            )}
          </div>

          {/* Right side intentionally empty; secondary row handles meta */}
          <div className="flex items-center gap-3 flex-shrink-0" />
        </div>
      </div>
    </div>
  );
}
