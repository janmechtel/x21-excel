import { useEffect, useMemo, useRef, useState } from "react";
import { Command, Sparkles } from "lucide-react";
import {
  getBaseSlashCommands,
  getExcelSlashCommands,
  isExcelCommand,
  refreshExcelSlashCommands,
  subscribeToSlashCommands,
} from "@/lib/slashCommands";
import { resolveIcon } from "@/lib/iconResolver";
import { Button } from "@/components/ui/button";
import type { SlashCommandDefinition } from "@/types/slash-commands";

interface EmptyStateProps {
  onCommandSelect: (command: SlashCommandDefinition) => void;
  onOpenCommands: () => void;
}

export function EmptyState({
  onCommandSelect,
  onOpenCommands,
}: EmptyStateProps) {
  const [commandRevision, setCommandRevision] = useState(0);
  const lastRefreshRef = useRef<number>(0);

  // Load Excel commands whenever EmptyState is shown (new chat opened)
  // Refresh if it's been more than 1 second since last refresh to avoid too many calls
  useEffect(() => {
    const now = Date.now();
    if (now - lastRefreshRef.current > 1000) {
      lastRefreshRef.current = now;
      refreshExcelSlashCommands();
    }
  }, []); // Only run on mount (when chat becomes empty)

  // Subscribe to command updates so we refresh when new commands are added
  useEffect(() => {
    const unsubscribe = subscribeToSlashCommands(() => {
      setCommandRevision((prev) => prev + 1);
    });
    return () => unsubscribe();
  }, []);

  const suggestedCommands = useMemo(() => {
    const excelCmds = getExcelSlashCommands();
    const baseCmds = getBaseSlashCommands();
    return [...excelCmds, ...baseCmds].slice(0, 5); // Show top 5 suggestions
  }, [commandRevision]);

  if (suggestedCommands.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full text-center px-6 py-12">
        <Sparkles className="w-12 h-12 text-slate-400 dark:text-slate-500 mb-4" />
        <h3 className="text-lg font-semibold text-slate-900 dark:text-slate-100 mb-2">
          Ready to get started?
        </h3>
        <p className="text-sm text-slate-600 dark:text-slate-400 mb-6 max-w-md">
          Type a question or use the Commands button to explore available
          actions.
        </p>
        <Button onClick={onOpenCommands} variant="outline" className="gap-2">
          <Command className="w-4 h-4" />
          View Commands
        </Button>
      </div>
    );
  }

  return (
    <div className="flex flex-col items-center justify-center h-full text-center px-1 py-1">
      <div className="flex items-center gap-2 mb-2">
        <Sparkles className="w-4 h-4 text-slate-400 dark:text-slate-500" />
        <span className="text-sm text-slate-600 dark:text-slate-400">
          Try these commands
        </span>
      </div>

      <div className="grid grid-cols-1 gap-1 w-full max-w-2xl mb-2">
        {suggestedCommands.map((command) => {
          const isCustom = isExcelCommand(command);
          return (
            <button
              key={command.id || command.name}
              onClick={() => onCommandSelect(command)}
              className={`text-left p-3 rounded-lg border transition-all group ${
                isCustom
                  ? "border-purple-300 dark:border-purple-700 bg-purple-50/50 dark:bg-purple-950/30 hover:border-purple-500 dark:hover:border-purple-400 hover:bg-purple-100/50 dark:hover:bg-purple-900/40"
                  : "border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 hover:border-blue-500 dark:hover:border-blue-400 hover:bg-blue-50/50 dark:hover:bg-blue-950/30"
              }`}
            >
              <div className="flex items-start gap-3">
                <div className="flex-shrink-0 mt-0.5">
                  {(() => {
                    const IconComponent = resolveIcon(command.icon);
                    return (
                      <div
                        className={`w-8 h-8 rounded-md flex items-center justify-center group-hover:transition-colors ${
                          isCustom
                            ? "bg-purple-100 dark:bg-purple-900/30 group-hover:bg-purple-200 dark:group-hover:bg-purple-900/50"
                            : "bg-blue-100 dark:bg-blue-900/30 group-hover:bg-blue-200 dark:group-hover:bg-blue-900/50"
                        }`}
                      >
                        <IconComponent
                          className={`w-4 h-4 ${
                            isCustom
                              ? "text-purple-600 dark:text-purple-400"
                              : "text-blue-600 dark:text-blue-400"
                          }`}
                        />
                      </div>
                    );
                  })()}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 mb-1">
                    <div className="font-medium text-sm text-slate-900 dark:text-slate-100 truncate">
                      {command.title || command.name}
                    </div>
                    {isCustom && (
                      <span className="text-xs px-1.5 py-0.5 rounded bg-purple-100 dark:bg-purple-900/50 text-purple-700 dark:text-purple-300 font-medium flex-shrink-0">
                        Custom
                      </span>
                    )}
                  </div>
                  <div className="text-xs text-slate-600 dark:text-slate-400 line-clamp-2 h-8">
                    {command.description}
                  </div>
                </div>
              </div>
            </button>
          );
        })}
      </div>

      <Button
        onClick={onOpenCommands}
        variant="ghost"
        size="sm"
        className="gap-2 text-slate-600 dark:text-slate-400"
      >
        <Command className="w-4 h-4" />
        View all commands
      </Button>
    </div>
  );
}
