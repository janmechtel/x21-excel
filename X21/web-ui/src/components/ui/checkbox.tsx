import * as React from "react";
import { InfoTooltip } from "./info-tooltip";

export interface CheckboxProps
  extends React.InputHTMLAttributes<HTMLInputElement> {
  label?: string;
  tooltip?: string;
}

const Checkbox = React.forwardRef<HTMLInputElement, CheckboxProps>(
  ({ className, label, tooltip, id, ...props }, ref) => {
    const checkboxId = id || React.useId();

    // Check if small size is requested via className
    const isSmall = className?.includes("text-xs");
    const inputSizeClass = isSmall ? "w-3.5 h-3.5" : "w-4 h-4";
    const labelSizeClass = isSmall ? "text-xs" : "text-sm";
    const labelColorClass = isSmall
      ? "text-slate-600 dark:text-slate-400"
      : "text-slate-700 dark:text-slate-300";

    return (
      <div className="flex items-center gap-2">
        <input
          id={checkboxId}
          type="checkbox"
          ref={ref}
          className={`${inputSizeClass} flex-shrink-0 text-blue-600 border-slate-300 rounded focus:ring-2 focus:ring-blue-500 ${
            className
              ?.replace(/text-\[10px\]|text-slate-\d+|dark:text-slate-\d+/g, "")
              .trim() || ""
          }`}
          {...props}
        />
        {label && (
          <label
            htmlFor={checkboxId}
            className={`ml-2 ${labelSizeClass} ${labelColorClass} cursor-pointer select-none`}
          >
            {label}
          </label>
        )}
        {tooltip && <InfoTooltip content={tooltip} />}
      </div>
    );
  },
);

Checkbox.displayName = "Checkbox";

export { Checkbox };
