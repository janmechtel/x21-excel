import { useEffect, useState } from "react";
import { Checkbox } from "@/components/ui/checkbox";
import { getApiBase } from "@/services/apiBase";

interface SaveCopiesCheckboxProps {
  className?: string;
  disabled?: boolean;
  onChange?: (enabled: boolean) => void;
  size?: "default" | "small";
}

export function SaveCopiesCheckbox({
  className = "",
  disabled = false,
  onChange,
  size = "default",
}: SaveCopiesCheckboxProps) {
  const [saveSnapshotsEnabled, setSaveSnapshotsEnabled] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  // Load preference on mount
  useEffect(() => {
    const fetchPreference = async () => {
      try {
        const base = await getApiBase();
        const response = await fetch(
          `${base}/api/user-preference?key=save_snapshots`,
        );
        if (response.ok) {
          const data = await response.json();
          const enabled =
            data.preferenceValue === true || data.preferenceValue === "true";
          setSaveSnapshotsEnabled(enabled);
          // Inform parent about the initial loaded value so dependent UI
          // (like the Generate summary button) can reflect the true state.
          onChange?.(enabled);
        }
      } catch (error) {
        console.warn("Failed to fetch save copies preference:", error);
        setSaveSnapshotsEnabled(false);
        onChange?.(false);
      } finally {
        setIsLoading(false);
      }
    };

    fetchPreference();
  }, []);

  const handleToggle = async (enabled: boolean) => {
    setSaveSnapshotsEnabled(enabled);

    // Notify parent component
    onChange?.(enabled);

    // Send preference update to server via REST
    try {
      const base = await getApiBase();
      await fetch(`${base}/api/user-preference`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          preferenceKey: "save_snapshots",
          preferenceValue: enabled,
        }),
      });
    } catch (error) {
      console.error("Failed to save preference:", error);
      // Revert on error
      setSaveSnapshotsEnabled(!enabled);
      onChange?.(!enabled);
    }
  };

  if (isLoading) {
    return null; // Or a loading skeleton
  }

  const sizeClasses =
    size === "small" ? "text-xs text-slate-600 dark:text-slate-400" : "";

  return (
    <Checkbox
      id="save-copies-checkbox"
      checked={saveSnapshotsEnabled}
      onChange={(e) => handleToggle(e.target.checked)}
      disabled={disabled}
      label="Save milestone copies"
      tooltip="When enabled, X21 will automatically save copies of your workbook changes. These copies allow you to track modifications over time and generate summaries of your work. Copies are stored locally on your device."
      className={`${className} ${sizeClasses}`}
    />
  );
}
