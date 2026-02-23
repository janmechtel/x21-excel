import { memo, type RefObject, useEffect, useMemo, useState } from "react";
import { AtSign } from "lucide-react";

import { SearchableOverlay } from "@/components/shared/SearchableOverlay";

interface SheetMentionPaletteProps {
  isOpen: boolean;
  sheets: string[];
  otherWorkbookSheets?: Array<{
    workbookName: string;
    sheetName: string;
    workbookFullName?: string;
  }>;
  activeWorkbookName?: string;
  isRequestingOtherWorkbooks?: boolean;
  highlightedIndex: number;
  query: string;
  onSelect: (sheetName: string) => void;
  onHighlight: (index: number) => void;
  onClose: () => void;
  textareaRef?: RefObject<HTMLTextAreaElement>;
}

export const SheetMentionPalette = memo(
  ({
    isOpen,
    sheets,
    otherWorkbookSheets = [],
    activeWorkbookName,
    highlightedIndex,
    query,
    onSelect,
    onHighlight,
    onClose,
    textareaRef,
  }: SheetMentionPaletteProps) => {
    const [searchValue, setSearchValue] = useState(query ?? "");

    useEffect(() => {
      if (isOpen) {
        setSearchValue(query ?? "");
        onHighlight(0);
      }
    }, [isOpen, query, onHighlight]);

    const handleClose = () => {
      setSearchValue("");
      onHighlight(0);
      onClose();
    };

    type SheetOption = {
      value: string;
      display: string;
      group: "active" | "other";
      subtitle?: string;
      workbookName?: string;
      workbookFullName?: string;
    };

    const options: SheetOption[] = useMemo(() => {
      const active = sheets.map((sheet) => ({
        value: sheet,
        display: sheet,
        group: "active" as const,
        subtitle: undefined,
      }));
      const others = otherWorkbookSheets.map((entry) => ({
        value: `${entry.workbookName}/${entry.sheetName}`,
        display: entry.sheetName,
        group: "other" as const,
        subtitle: undefined,
        workbookName: entry.workbookName,
        workbookFullName: entry.workbookFullName,
      }));
      return [...active, ...others];
    }, [sheets, otherWorkbookSheets, activeWorkbookName]);

    useEffect(() => {
      console.log("Sheet mention options", options);
    }, [options]);

    const filteredOptions = useMemo(() => {
      const term = searchValue?.trim().toLowerCase() ?? "";
      if (!term) return options;
      return options.filter(
        (opt) =>
          opt.display.toLowerCase().includes(term) ||
          (opt.subtitle?.toLowerCase().includes(term) ?? false),
      );
    }, [options, searchValue]);

    const groupedOptions = useMemo(() => {
      const groupMap = new Map<string, SheetOption[]>();
      filteredOptions.forEach((opt) => {
        const title =
          opt.group === "active"
            ? activeWorkbookName || "Current workbook"
            : opt.workbookFullName || opt.workbookName || "Other workbooks";
        const items = groupMap.get(title) ?? [];
        items.push(opt);
        groupMap.set(title, items);
      });

      const currentTitle = activeWorkbookName || "Current workbook";
      const orderedKeys: string[] = [];
      if (groupMap.has(currentTitle)) orderedKeys.push(currentTitle);
      for (const key of groupMap.keys()) {
        if (key !== currentTitle) orderedKeys.push(key);
      }

      return orderedKeys.map((title) => ({
        title,
        options: groupMap.get(title) ?? [],
      }));
    }, [filteredOptions, activeWorkbookName]);

    const renderEmptyState = (
      <div className="px-4 py-5 text-xs text-slate-500 dark:text-slate-400">
        {searchValue ? (
          <>No sheets match "{searchValue}". Try a different keyword.</>
        ) : (
          <>Type to jump to another worksheet.</>
        )}
      </div>
    );

    const handleSelect = () => {
      if (!filteredOptions.length) return;
      const clampedIndex = Math.min(
        Math.max(highlightedIndex, 0),
        filteredOptions.length - 1,
      );
      const selected = filteredOptions[clampedIndex];
      if (selected) {
        onSelect(selected.value);
      }
    };

    return (
      <SearchableOverlay
        isOpen={isOpen}
        onClose={handleClose}
        title="Sheets"
        icon={AtSign}
        searchValue={searchValue}
        onSearchChange={setSearchValue}
        searchPlaceholder="Search sheets"
        highlightedIndex={highlightedIndex}
        itemCount={filteredOptions.length}
        onHighlight={onHighlight}
        onSelect={handleSelect}
        emptyStateMessage={
          filteredOptions.length === 0 ? renderEmptyState : undefined
        }
        textareaRef={textareaRef}
        maxWidth="max-w-lg"
      >
        {(() => {
          let optionRenderIndex = -1;
          return groupedOptions.map((group) => (
            <div key={group.title} className="pb-2">
              <div className="px-4 pt-2 pb-1 text-[10px] uppercase tracking-wide text-slate-500 dark:text-slate-400">
                {group.title}
              </div>
              <div className="space-y-1">
                {group.options.map((option) => {
                  optionRenderIndex += 1;
                  const isHighlighted = optionRenderIndex === highlightedIndex;

                  return (
                    <div
                      key={option.value}
                      className={`mx-2 rounded cursor-pointer transition-colors ${
                        isHighlighted
                          ? "bg-slate-100 dark:bg-slate-800/60 text-slate-900 dark:text-white"
                          : "text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-800/40"
                      }`}
                      onMouseDown={(event) => event.preventDefault()}
                      onMouseEnter={() => onHighlight(optionRenderIndex)}
                      onClick={() => {
                        onHighlight(optionRenderIndex);
                        onSelect(option.value);
                      }}
                    >
                      <div className="px-4 py-2 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <span className="text-base">@</span>
                          <div>
                            <div className="text-sm font-medium">
                              {option.display}
                            </div>
                            {option.subtitle && (
                              <div className="text-[11px] text-slate-500 dark:text-slate-400">
                                {option.subtitle}
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ));
        })()}
      </SearchableOverlay>
    );
  },
);
