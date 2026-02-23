import { memo, useEffect, useMemo, useState } from "react";
import { Wrench } from "lucide-react";

import { ToolNames } from "@/types/chat";
import { toolsNotRequiringApproval } from "@/utils/tools";
import { SearchableOverlay } from "@/components/shared/SearchableOverlay";
import { getApiBase } from "@/services/apiBase";
import {
  ColumnWidthModes,
  getColumnWidthMessage,
  type ColumnWidthMode,
} from "@/utils/columnWidth";

const COLUMN_WIDTH_MODE_KEY = "column_width_mode";

interface ToolTogglePanelProps {
  isOpen: boolean;
  activeTools: Set<string>;
  highlightedIndex: number;
  onSelect: (toolId: string) => void;
  onHighlight: (index: number) => void;
  onClose: () => void;
  onOpenSettings: () => void;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
  availableTools: {
    id: string;
    name: string;
    description: string;
  }[];
}

export const ToolTogglePanel = memo(
  ({
    isOpen,
    activeTools,
    highlightedIndex,
    onSelect,
    onHighlight,
    onClose,
    onOpenSettings,
    textareaRef,
    availableTools,
  }: ToolTogglePanelProps) => {
    const [searchValue, setSearchValue] = useState("");
    const [columnWidthMode, setColumnWidthMode] = useState<ColumnWidthMode>(
      ColumnWidthModes.SMART,
    );

    useEffect(() => {
      if (isOpen) {
        setSearchValue("");
        onHighlight(0);
      }
    }, [isOpen, onHighlight]);

    useEffect(() => {
      if (!isOpen) return;
      let isActive = true;

      const fetchPreference = async () => {
        try {
          const base = await getApiBase();
          const response = await fetch(
            `${base}/api/user-preference?key=${COLUMN_WIDTH_MODE_KEY}&type=string&default=${ColumnWidthModes.SMART}`,
          );
          if (!response.ok) return;
          const data = await response.json();
          const raw = String(data.preferenceValue || "").toLowerCase();
          const next =
            raw === ColumnWidthModes.ALWAYS ||
            raw === ColumnWidthModes.SMART ||
            raw === ColumnWidthModes.NEVER
              ? raw
              : ColumnWidthModes.SMART;
          if (isActive) {
            setColumnWidthMode(next);
          }
        } catch (error) {
          console.warn("Failed to fetch column width preference:", error);
        }
      };

      fetchPreference();

      return () => {
        isActive = false;
      };
    }, [isOpen]);

    const filteredTools = useMemo(() => {
      const term = searchValue?.trim().toLowerCase() ?? "";
      if (!term) return availableTools;
      return availableTools.filter((tool) =>
        [tool.id, tool.name, tool.description]
          .join(" ")
          .toLowerCase()
          .includes(term),
      );
    }, [availableTools, searchValue]);

    const renderEmptyState = (
      <div className="px-4 py-5 text-xs text-slate-500 dark:text-slate-400">
        {searchValue ? (
          <>No tools match "{searchValue}". Try a different keyword.</>
        ) : (
          <>Type to search available tools.</>
        )}
      </div>
    );

    const handleSelect = () => {
      const selectedTool = filteredTools[highlightedIndex];
      if (selectedTool) {
        onSelect(selectedTool.id);
      }
    };

    return (
      <SearchableOverlay
        isOpen={isOpen}
        onClose={onClose}
        title="Tools"
        icon={Wrench}
        searchValue={searchValue}
        onSearchChange={setSearchValue}
        searchPlaceholder="Search tools"
        highlightedIndex={highlightedIndex}
        itemCount={filteredTools.length}
        onHighlight={onHighlight}
        onSelect={handleSelect}
        emptyStateMessage={
          filteredTools.length === 0 ? renderEmptyState : undefined
        }
        textareaRef={textareaRef}
      >
        {filteredTools.map((tool, index) => {
          const isActive = index === highlightedIndex;
          const isToolActive = activeTools.has(tool.id);
          const isAutoApproved = toolsNotRequiringApproval.includes(tool.id);
          return (
            <button
              type="button"
              key={tool.id}
              className={`w-full text-left px-4 py-2 flex items-start gap-3 text-xs transition-colors ${
                isActive
                  ? "bg-blue-50 dark:bg-blue-900/30 text-blue-900 dark:text-blue-100"
                  : "text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800/70"
              }`}
              onMouseEnter={() => onHighlight(index)}
              onFocus={() => onHighlight(index)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => onSelect(tool.id)}
            >
              <div className="flex flex-col gap-1 flex-1">
                <div className="flex items-center gap-2">
                  <div
                    className={`w-2.5 h-2.5 flex-shrink-0 flex items-center justify-center ${
                      isToolActive
                        ? "text-blue-600 dark:text-blue-400"
                        : "text-slate-400 dark:text-slate-500"
                    }`}
                  >
                    <Wrench className="w-2.5 h-2.5" />
                  </div>
                  <span
                    className={`text-sm font-semibold leading-tight ${
                      isToolActive ? "text-blue-700 dark:text-blue-300" : ""
                    }`}
                  >
                    {tool.name}
                  </span>
                  {isToolActive && (
                    <span className="text-[9px] text-blue-600 dark:text-blue-400 font-medium px-1.5 py-0.5 bg-blue-100 dark:bg-blue-900/30 rounded">
                      ON
                    </span>
                  )}
                  {isAutoApproved && (
                    <span className="text-[9px] text-green-600 dark:text-green-400 font-medium px-1.5 py-0.5 bg-green-100 dark:bg-green-900/30 rounded">
                      AUTO
                    </span>
                  )}
                </div>
                <div className="ml-5 text-[11px] text-slate-500 dark:text-slate-400 leading-snug">
                  <p>{tool.description}</p>
                  {tool.id === ToolNames.WRITE_VALUES_BATCH && (
                    <div className="mt-0.5">
                      <span>{getColumnWidthMessage(columnWidthMode)} (</span>
                      <button
                        type="button"
                        onClick={(event) => {
                          event.preventDefault();
                          event.stopPropagation();
                          onOpenSettings();
                        }}
                        className="text-[11px] text-blue-600 dark:text-blue-400 hover:underline"
                      >
                        Open settings
                      </button>
                      <span>)</span>
                    </div>
                  )}
                </div>
              </div>
            </button>
          );
        })}
      </SearchableOverlay>
    );
  },
);

ToolTogglePanel.displayName = "ToolTogglePanel";
