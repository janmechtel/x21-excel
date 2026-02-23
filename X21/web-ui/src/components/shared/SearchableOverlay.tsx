import { ReactNode, useEffect, useRef } from "react";
import { LucideIcon, X } from "lucide-react";

export interface SearchableOverlayProps {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  icon: LucideIcon;
  searchValue: string;
  onSearchChange: (value: string) => void;
  searchPlaceholder: string;
  highlightedIndex: number;
  itemCount: number;
  onHighlight: (index: number) => void;
  onSelect: () => void;
  children: ReactNode;
  emptyStateMessage?: ReactNode;
  headerActions?: ReactNode;
  searchBarActions?: ReactNode; // Content to show above the search bar (e.g., scope selector)
  footer?: ReactNode;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
  maxWidth?: string;
  maxHeight?: string;
}

/**
 * Shared base component for searchable overlays (History, Commands, Tools)
 * Provides consistent layout, keyboard navigation, and focus management
 */
export function SearchableOverlay({
  isOpen,
  onClose,
  title,
  icon: Icon,
  searchValue,
  onSearchChange,
  searchPlaceholder,
  highlightedIndex,
  itemCount,
  onHighlight,
  onSelect,
  children,
  emptyStateMessage,
  headerActions,
  searchBarActions,
  footer,
  textareaRef,
  maxWidth = "max-w-2xl",
  maxHeight = "max-h-[80vh]",
}: SearchableOverlayProps) {
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  const panelRef = useRef<HTMLDivElement | null>(null);

  // Focus search input when opened
  useEffect(() => {
    if (isOpen) {
      // Focus immediately to catch fast typers on first open
      searchInputRef.current?.focus();

      // Also focus after DOM is ready, in case immediate focus failed
      // (e.g., WebView2 initialization delay on first load)
      requestAnimationFrame(() => {
        setTimeout(() => {
          searchInputRef.current?.focus();
        }, 100);
      });
    }
  }, [isOpen]);

  // Close when clicking outside the panel
  useEffect(() => {
    if (!isOpen) return;

    const handleClickOutside = (event: MouseEvent) => {
      if (
        panelRef.current &&
        !panelRef.current.contains(event.target as Node)
      ) {
        handleClose();
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    document.addEventListener("click", handleClickOutside, true);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      document.removeEventListener("click", handleClickOutside, true);
    };
  }, [isOpen]);

  const handleClose = () => {
    onClose();
    // Return focus to textarea after closing
    if (textareaRef) {
      setTimeout(() => {
        textareaRef.current?.focus();
      }, 0);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      onHighlight(Math.min(highlightedIndex + 1, itemCount - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      onHighlight(Math.max(highlightedIndex - 1, 0));
    } else if (e.key === "Enter") {
      e.preventDefault();
      onSelect();
    } else if (e.key === "Escape") {
      e.preventDefault();
      handleClose();
    }
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-[65] bg-black/30 dark:bg-black/40 backdrop-blur-sm flex items-start justify-center pt-9"
      onClick={handleClose}
    >
      <div
        ref={panelRef}
        className={`w-full ${maxWidth} ${maxHeight} bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 shadow-2xl flex flex-col overflow-hidden`}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="px-4 py-2 text-[11px] font-medium text-slate-500 dark:text-slate-400 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <Icon className="w-3 h-3" />
            <span className="tracking-wide uppercase text-[10px]">{title}</span>
          </div>
          <div className="flex items-center gap-2">
            {headerActions}
            <button
              type="button"
              onClick={handleClose}
              className="text-slate-400 hover:text-slate-600 dark:text-slate-500 dark:hover:text-slate-300 transition-colors"
              aria-label={`Close ${title}`}
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Search bar */}
        <div className="px-4 py-2 border-b border-slate-200 dark:border-slate-700">
          {searchBarActions && <div className="mb-0.5">{searchBarActions}</div>}
          <input
            type="text"
            value={searchValue}
            ref={searchInputRef}
            onChange={(e) => {
              onSearchChange(e.target.value);
              onHighlight(0);
            }}
            onKeyDown={handleKeyDown}
            placeholder={searchPlaceholder}
            className="w-full border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-900 px-2 py-1 text-[11px] text-slate-700 dark:text-slate-200 placeholder:text-slate-400 dark:placeholder:text-slate-500 focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500"
          />
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto">
          {itemCount === 0 && emptyStateMessage ? emptyStateMessage : children}
        </div>

        {/* Footer */}
        {footer && (
          <div className="border-t border-slate-200 dark:border-slate-700">
            {footer}
          </div>
        )}
      </div>
    </div>
  );
}
