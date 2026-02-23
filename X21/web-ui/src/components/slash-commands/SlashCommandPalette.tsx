import { memo, useEffect, useMemo, useState } from "react";
import { Sparkles } from "lucide-react";
import type { SlashCommandDefinition } from "@/types/slash-commands";
import { SearchableOverlay } from "@/components/shared/SearchableOverlay";
import { resolveIcon } from "@/lib/iconResolver";
import { isExcelCommand } from "@/lib/slashCommands";

interface SlashCommandPaletteProps {
  isOpen: boolean;
  commands: SlashCommandDefinition[];
  highlightedIndex: number;
  query: string;
  onSelect: (command: SlashCommandDefinition) => void;
  onHighlight: (index: number) => void;
  onClose: () => void;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
}

export const SlashCommandPalette = memo(
  ({
    isOpen,
    commands,
    highlightedIndex,
    query,
    onSelect,
    onHighlight,
    onClose,
    textareaRef,
  }: SlashCommandPaletteProps) => {
    const [searchValue, setSearchValue] = useState(query ?? "");

    useEffect(() => {
      if (isOpen) {
        setSearchValue(query ?? "");
        onHighlight(0);
      }
    }, [isOpen, query, onHighlight]);

    const filteredCommands = useMemo(() => {
      const term = searchValue?.trim().toLowerCase() ?? "";
      if (!term) return commands;
      return commands.filter((c) =>
        [
          c.id,
          (c as any).name ?? "",
          c.title,
          c.description,
          ...(c.keywords ?? []),
        ]
          .join(" ")
          .toLowerCase()
          .includes(term),
      );
    }, [commands, searchValue]);

    const renderEmptyState = (
      <div className="px-4 py-5 text-xs text-slate-500 dark:text-slate-400">
        {searchValue ? (
          <>No commands match "{searchValue}". Try a different keyword.</>
        ) : (
          <>Type to search quick commands for Excel automations.</>
        )}
      </div>
    );

    const handleSelect = () => {
      const selectedCommand = filteredCommands[highlightedIndex];
      if (selectedCommand) {
        onSelect(selectedCommand);
      }
    };

    return (
      <SearchableOverlay
        isOpen={isOpen}
        onClose={onClose}
        title="Commands"
        icon={Sparkles}
        searchValue={searchValue}
        onSearchChange={setSearchValue}
        searchPlaceholder="Search commands"
        highlightedIndex={highlightedIndex}
        itemCount={filteredCommands.length}
        onHighlight={onHighlight}
        onSelect={handleSelect}
        emptyStateMessage={
          filteredCommands.length === 0 ? renderEmptyState : undefined
        }
        textareaRef={textareaRef}
      >
        {filteredCommands.map((command, index) => {
          const isActive = index === highlightedIndex;
          const isCustom = isExcelCommand(command);
          const IconComponent = resolveIcon(command.icon);
          return (
            <button
              type="button"
              key={command.id}
              className={`w-full text-left px-4 py-2 flex items-start gap-3 text-xs transition-colors ${
                isActive
                  ? "bg-blue-50 dark:bg-blue-900/30 text-blue-900 dark:text-blue-100"
                  : "text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800/70"
              }`}
              onMouseEnter={() => onHighlight(index)}
              onFocus={() => onHighlight(index)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => onSelect(command)}
            >
              <div
                className={`w-8 h-8 flex items-center justify-center flex-shrink-0 mt-0.5 ${
                  isCustom
                    ? "bg-purple-100 dark:bg-purple-900/30"
                    : "bg-blue-100 dark:bg-blue-900/30"
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
              <div className="flex flex-col gap-1 flex-1">
                <div className="flex items-center gap-2">
                  <span className="text-sm font-semibold leading-tight">
                    {command.title}
                  </span>
                  {isCustom && (
                    <span className="text-xs px-1 py-0.5 bg-purple-100 dark:bg-purple-900/50 text-purple-700 dark:text-purple-300 font-medium">
                      Custom
                    </span>
                  )}
                </div>
                <p className="text-[11px] text-slate-500 dark:text-slate-400 leading-snug">
                  {command.description}
                </p>
              </div>
            </button>
          );
        })}
      </SearchableOverlay>
    );
  },
);

SlashCommandPalette.displayName = "SlashCommandPalette";
