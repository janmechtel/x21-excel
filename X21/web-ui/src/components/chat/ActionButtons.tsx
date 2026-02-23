import { Check, CheckCheck, Eye, EyeOff, X } from "lucide-react";

interface ActionButtonsProps {
  isViewingTool: boolean;
  onViewTool: () => void;
  onApproveTool: () => void;
  onRejectTool: () => void;
  onApproveAll: () => void;
}

export function ActionButtons({
  isViewingTool,
  onViewTool,
  onApproveTool,
  onRejectTool,
  onApproveAll,
}: ActionButtonsProps) {
  return (
    <div className="grid grid-cols-2 gap-2 w-full max-w-xs">
      <button
        onClick={onViewTool}
        className="flex items-center justify-center gap-1 px-2 py-1.5 text-xs border border-gray-300 hover:bg-blue-50 dark:hover:bg-blue-900/20 hover:border-blue-300 transition-colors"
      >
        {isViewingTool ? (
          <Eye className="w-3 h-3" />
        ) : (
          <EyeOff className="w-3 h-3" />
        )}
        {isViewingTool ? "Viewing" : "View"}
      </button>
      <button
        onClick={onApproveTool}
        className="flex items-center justify-center gap-1 px-2 py-1.5 text-xs bg-green-600 text-white hover:bg-green-700 transition-colors"
      >
        <Check className="w-3 h-3" />
        Accept
      </button>

      <button
        onClick={onRejectTool}
        className="flex items-center justify-center gap-1 px-2 py-1.5 text-xs border border-gray-300 hover:bg-red-50 dark:hover:bg-red-900/20 hover:border-red-300 text-red-600 transition-colors"
      >
        <X className="w-3 h-3" />
        Reject
      </button>
      <button
        onClick={onApproveAll}
        className="flex items-center justify-center gap-1 px-2 py-1.5 text-xs border border-gray-300 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 hover:border-emerald-300 text-emerald-600 transition-colors"
      >
        <CheckCheck className="w-3 h-3" />
        Accept All
      </button>
    </div>
  );
}
