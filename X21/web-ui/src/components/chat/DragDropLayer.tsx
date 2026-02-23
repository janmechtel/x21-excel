import { forwardRef, type ReactNode } from "react";
import { Upload } from "lucide-react";

interface DragDropLayerProps {
  isDragOver: boolean;
  className?: string;
  children: ReactNode;
  showOverlay?: boolean;
  onDragOver: React.DragEventHandler;
  onDragEnter: React.DragEventHandler;
  onDragLeave: React.DragEventHandler;
  onDrop: React.DragEventHandler;
  onScroll?: React.UIEventHandler<HTMLDivElement>;
}

export const DragDropLayer = forwardRef<HTMLDivElement, DragDropLayerProps>(
  (
    {
      isDragOver,
      className,
      children,
      showOverlay = true,
      onDragOver,
      onDragEnter,
      onDragLeave,
      onDrop,
      onScroll,
    },
    ref,
  ) => (
    <div
      ref={ref}
      className={className}
      onDragOver={onDragOver}
      onDragEnter={onDragEnter}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
      onScroll={onScroll}
    >
      {showOverlay && isDragOver && (
        <div className="fixed inset-0 bg-blue-500/30 backdrop-blur-md z-[60] flex items-center justify-center">
          <div className="bg-white dark:bg-slate-800 p-8 shadow-2xl border-2 border-dashed border-blue-400 dark:border-blue-600">
            <div className="text-center">
              <Upload className="w-12 h-12 text-blue-500 mx-auto mb-4" />
              <h3 className="text-lg font-semibold text-slate-900 dark:text-slate-100 mb-2">
                Drop PDFs anywhere
              </h3>
              <p className="text-sm text-slate-600 dark:text-slate-400">
                Release to upload your PDF documents
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-500 mt-1">
                Multiple files supported • Works anywhere on the interface!
              </p>
            </div>
          </div>
        </div>
      )}
      {children}
    </div>
  ),
);

DragDropLayer.displayName = "DragDropLayer";
