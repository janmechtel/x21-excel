import React from "react";
import { useAuth } from "../../contexts/AuthContext";
import AuthPage from "./AuthPage";
import { isSupabaseConfigured } from "../../lib/supabaseClient";

interface ProtectedRouteProps {
  children: React.ReactNode;
}

export const ProtectedRoute: React.FC<ProtectedRouteProps> = ({ children }) => {
  const { user, loading } = useAuth();

  // Show loading spinner while checking authentication
  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-900 dark:to-slate-800">
        <div className="text-center">
          <div className="w-8 h-8 border-4 border-blue-600 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-slate-600 dark:text-slate-400">Loading...</p>
        </div>
      </div>
    );
  }

  // If no user is authenticated and Supabase is configured, show the auth page
  if (!user && isSupabaseConfigured) {
    return <AuthPage />;
  }

  // If user is authenticated, render the protected content
  return <>{children}</>;
};

export default ProtectedRoute;
