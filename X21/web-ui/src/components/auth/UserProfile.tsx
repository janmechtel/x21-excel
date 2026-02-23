import React from "react";
import { Menu } from "lucide-react";
import { useAuth } from "../../contexts/AuthContext";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Button } from "@/components/ui/button";

export const UserProfile: React.FC = () => {
  const { user, signOut } = useAuth();

  if (!user) return null;

  const handleSignOut = async () => {
    try {
      await signOut();
    } catch (error) {
      console.error("Error signing out:", error);
    }
  };

  return (
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
        <DropdownMenuItem onClick={handleSignOut}>Sign out</DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
};

export default UserProfile;
