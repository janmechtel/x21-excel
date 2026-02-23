import { FileText, History, Menu, Plus } from "lucide-react";

import { Button } from "@/components/ui/button";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { useAuth } from "@/contexts/AuthContext";
import { isSupabaseConfigured } from "@/lib/supabaseClient";

interface AppHeaderProps {
  onToggleHistory: () => void;
  onToggleLogs: () => void;
  onNewChat: () => void;
  onOpenSettings: () => void;
}

export function AppHeader({
  onToggleHistory,
  onToggleLogs,
  onNewChat,
  onOpenSettings,
}: AppHeaderProps) {
  const { user, signOut } = useAuth();

  const handleSignOut = async () => {
    try {
      await signOut();
    } catch (error) {
      console.error("Error signing out:", error);
    }
  };

  return (
    <>
      <div className="flex-shrink-0">
        <div className="flex items-center justify-between py-1 max-w-4xl mx-auto">
          <div className="flex-1 flex items-center">
            {(user || !isSupabaseConfigured) && (
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-7 px-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"
                  >
                    <Menu className="w-4 h-4" />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start">
                  <DropdownMenuItem onClick={onOpenSettings}>
                    Settings
                  </DropdownMenuItem>
                  {user && (
                    <DropdownMenuItem onClick={handleSignOut}>
                      Sign out
                    </DropdownMenuItem>
                  )}
                </DropdownMenuContent>
              </DropdownMenu>
            )}
          </div>
          <div className="flex-1 flex items-center justify-center">
            {/* WebSocket indicator moved to status bar */}
          </div>
          <div className="flex-1 flex items-center justify-end gap-1">
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  onClick={onToggleLogs}
                  variant="ghost"
                  size="sm"
                  className="h-7 px-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"
                >
                  <FileText className="w-3 h-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>View activity logs</p>
              </TooltipContent>
            </Tooltip>

            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  onClick={onToggleHistory}
                  variant="ghost"
                  size="sm"
                  className="h-7 px-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"
                >
                  <History className="w-3 h-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>Browse recent conversations (Ctrl+H or Ctrl+Shift+H)</p>
              </TooltipContent>
            </Tooltip>

            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  onClick={onNewChat}
                  variant="ghost"
                  size="sm"
                  className="h-7 px-1.5 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-300"
                >
                  <Plus className="w-3 h-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent>
                <p>Start a new conversation and clear history (Ctrl+N)</p>
              </TooltipContent>
            </Tooltip>
          </div>
        </div>
      </div>
    </>
  );
}
