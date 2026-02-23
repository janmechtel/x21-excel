import { clsx } from "clsx";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";

type StatusMetaProps = {
  hasMessages: boolean;
  totalTokens: number;
  inputTokens: number;
  outputTokens: number;
  wsConnected: boolean;
  wsUrl: string;
};

export function StatusMetaRow({
  hasMessages,
  totalTokens,
  inputTokens,
  outputTokens,
  wsConnected,
  wsUrl,
}: StatusMetaProps) {
  return (
    <div className="flex flex-nowrap items-center justify-between gap-2 text-[9px] text-slate-600 dark:text-slate-300 whitespace-nowrap">
      <Tooltip>
        <TooltipTrigger asChild>
          <div className="flex items-center cursor-default">
            <span
              className={clsx(
                "h-2.5 w-2.5 rounded-full",
                wsConnected ? "bg-green-500" : "bg-red-500",
              )}
              aria-label={wsConnected ? "Connected" : "Disconnected"}
            />
          </div>
        </TooltipTrigger>
        <TooltipContent>
          <div className="text-xs space-y-1">
            <div>
              <strong>Status:</strong>{" "}
              {wsConnected ? "Connected" : "Disconnected"}
            </div>
            <div>
              <strong>Endpoint:</strong> {wsUrl}
            </div>
            <div>
              <strong>↓ Input:</strong> {inputTokens.toLocaleString()} tokens
            </div>
            <div>
              <strong>↑ Output:</strong> {outputTokens.toLocaleString()} tokens
            </div>
            <div className="text-slate-500">
              WebSocket connection to Deno server
            </div>
          </div>
        </TooltipContent>
      </Tooltip>

      <div className="flex items-center flex-nowrap">
        {hasMessages && totalTokens > 0 && (
          <Tooltip>
            <TooltipTrigger asChild>
              <div className="flex items-center gap-1.5 cursor-default">
                <span className="text-slate-400" aria-label="Input tokens">
                  ↓
                </span>
                <span className="text-slate-400" aria-label="Output tokens">
                  ↑
                </span>
                <span className="font-mono font-semibold text-slate-700 dark:text-slate-100">
                  {totalTokens.toLocaleString()}
                </span>
              </div>
            </TooltipTrigger>
            <TooltipContent>
              <div className="text-xs space-y-1">
                <div>
                  <strong>↓ Input:</strong> {inputTokens.toLocaleString()}{" "}
                  tokens sent to model
                </div>
                <div>
                  <strong>↑ Output:</strong> {outputTokens.toLocaleString()}{" "}
                  tokens generated
                </div>
                <div>
                  <strong>Total:</strong> {totalTokens.toLocaleString()} tokens
                </div>
              </div>
            </TooltipContent>
          </Tooltip>
        )}
      </div>
    </div>
  );
}
