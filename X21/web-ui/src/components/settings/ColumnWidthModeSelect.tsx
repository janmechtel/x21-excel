import { useEffect, useRef, useState } from "react";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { InfoTooltip } from "@/components/ui/info-tooltip";
import { getApiBase } from "@/services/apiBase";
import { ColumnWidthModes, type ColumnWidthMode } from "@/utils/columnWidth";

const COLUMN_WIDTH_MODE_KEY = "column_width_mode";

const MODE_OPTIONS = [
  { value: ColumnWidthModes.ALWAYS, label: "Always auto fit" },
  {
    value: ColumnWidthModes.SMART,
    label: "Smart auto-fit",
  },
  { value: ColumnWidthModes.NEVER, label: "No auto fit" },
] as const;

interface ColumnWidthModeSelectProps {
  disabled?: boolean;
}

export function ColumnWidthModeSelect({
  disabled = false,
}: ColumnWidthModeSelectProps) {
  const [mode, setMode] = useState<ColumnWidthMode>(ColumnWidthModes.SMART);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const isSavingRef = useRef(false);

  useEffect(() => {
    if (!toastMessage) return;
    const timeout = setTimeout(() => setToastMessage(null), 4000);
    return () => clearTimeout(timeout);
  }, [toastMessage]);

  useEffect(() => {
    const fetchPreference = async () => {
      try {
        const base = await getApiBase();
        const response = await fetch(
          `${base}/api/user-preference?key=${COLUMN_WIDTH_MODE_KEY}&type=string&default=${ColumnWidthModes.SMART}`,
        );
        if (response.ok) {
          const data = await response.json();
          const raw = String(data.preferenceValue || "").toLowerCase();
          const next = MODE_OPTIONS.find((option) => option.value === raw)
            ?.value as ColumnWidthMode | undefined;
          setMode(next ?? ColumnWidthModes.SMART);
        }
      } catch (error) {
        console.warn("Failed to fetch column width preference:", error);
        setMode(ColumnWidthModes.SMART);
      } finally {
        setIsLoading(false);
      }
    };

    fetchPreference();
  }, []);

  const handleChange = async (nextMode: ColumnWidthMode) => {
    if (isSavingRef.current) return;
    isSavingRef.current = true;
    const previous = mode;
    setMode(nextMode);
    setIsSaving(true);

    try {
      const base = await getApiBase();
      const response = await fetch(`${base}/api/user-preference`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          preferenceKey: COLUMN_WIDTH_MODE_KEY,
          preferenceValue: nextMode,
        }),
      });

      if (!response.ok) {
        throw new Error("Failed to save preference");
      }
    } catch (error) {
      console.error("Failed to save column width preference:", error);
      setMode(previous);
      setToastMessage("Failed to save column width preference.");
    } finally {
      setIsSaving(false);
      isSavingRef.current = false;
    }
  };

  if (isLoading) {
    return null;
  }

  const isDisabled = disabled || isSaving;

  return (
    <div>
      <label
        htmlFor="column-width-mode"
        className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
      >
        <span>Auto fit column width after writing</span>
        <InfoTooltip content="Controls whether columns auto-fit after AI writes to them. “Smart auto-fit” only resizes columns that have not been manually adjusted by the user." />
        {isSaving && (
          <span className="ml-auto flex items-center gap-1 text-xs text-slate-500 dark:text-slate-400">
            <span className="inline-block w-3 h-3 border-2 border-slate-400 border-t-transparent rounded-full animate-spin"></span>
            Saving...
          </span>
        )}
      </label>
      <select
        id="column-width-mode"
        value={mode}
        onChange={(e) => handleChange(e.target.value as ColumnWidthMode)}
        className={`w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 ${
          isDisabled ? "opacity-50 cursor-not-allowed" : ""
        }`}
        disabled={isDisabled}
      >
        {MODE_OPTIONS.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
      {toastMessage && (
        <Alert
          variant="destructive"
          className="fixed bottom-4 right-4 z-50 w-80 shadow-lg"
        >
          <AlertDescription className="text-sm">
            {toastMessage}
          </AlertDescription>
        </Alert>
      )}
    </div>
  );
}
