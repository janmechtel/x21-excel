import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";

interface NewChatDialogProps {
  open: boolean;
  isStreaming: boolean;
  onCancel: () => void;
  onConfirm: () => void;
}

export function NewChatDialog({
  open,
  isStreaming,
  onCancel,
  onConfirm,
}: NewChatDialogProps) {
  if (!open) return null;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <Card className="w-96 max-w-lg mx-4">
        <CardContent className="p-6">
          <h2 className="text-lg font-semibold mb-2">Start New Chat?</h2>
          <p className="text-slate-600 dark:text-slate-400 mb-6">
            {isStreaming
              ? "This will stop the current request and clear all chat history. This action cannot be undone."
              : "This will clear all chat history. This action cannot be undone."}
          </p>
          <div className="flex justify-end gap-3">
            <Button variant="outline" onClick={onCancel}>
              Cancel
            </Button>
            <Button
              onClick={onConfirm}
              className="bg-blue-600 hover:bg-blue-700 text-white"
            >
              {isStreaming ? "Stop and Start New Chat" : "Start New Chat"}
            </Button>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
