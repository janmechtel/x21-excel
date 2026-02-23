import React, { createContext, useContext, useEffect, useState } from "react";
import { Session, User } from "@supabase/supabase-js";
import { supabase, isSupabaseConfigured } from "../lib/supabaseClient";
import posthog from "../posthog";
import { webSocketChatService } from "../services/webSocketChatService";
import { webViewBridge } from "../services/webViewBridge";

interface AuthContextType {
  user: User | null;
  session: Session | null;
  loading: boolean;
  signOut: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({
  children,
}) => {
  const [user, setUser] = useState<User | null>(null);
  const [session, setSession] = useState<Session | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!isSupabaseConfigured) {
      setLoading(false);
      webSocketChatService.updateUserEmail(null);
      void webViewBridge.send("userEmailReady", { email: null });
      return;
    }

    // Get initial session
    const getInitialSession = async () => {
      const {
        data: { session },
      } = await supabase!.auth.getSession();
      setSession(session);
      setUser(session?.user ?? null);
      setLoading(false);

      // Identify user with PostHog if logged in (only in production)
      if (session?.user?.email && posthog.__loaded) {
        posthog.identify(session.user.email, {
          email: session.user.email,
          user_id: session.user.id,
          auth_event: "INITIAL_SESSION",
        });
        console.log("PostHog identified initial user:", session.user.email);
      }

      const email = session?.user?.email ?? null;
      webSocketChatService.updateUserEmail(email);
      void webViewBridge.send("userEmailReady", { email });
    };

    getInitialSession();

    // Listen for auth changes
    const {
      data: { subscription },
    } = supabase!.auth.onAuthStateChange(async (event, session) => {
      console.log("Auth state changed:", event, session);
      setSession(session);
      setUser(session?.user ?? null);
      setLoading(false);

      // Update PostHog identification (only in production)
      if (session?.user?.email && posthog.__loaded) {
        posthog.identify(session.user.email, {
          email: session.user.email,
          user_id: session.user.id,
          auth_event: event,
        });
        console.log("PostHog identified user:", session.user.email);
      } else if (event === "SIGNED_OUT" && posthog.__loaded) {
        posthog.reset();
        console.log("PostHog user reset on sign out");
      }

      const email = session?.user?.email ?? null;
      webSocketChatService.updateUserEmail(email);
      void webViewBridge.send("userEmailReady", { email });
    });

    return () => subscription.unsubscribe();
  }, []);

  const signOut = async () => {
    if (!isSupabaseConfigured) return;
    const { error } = await supabase!.auth.signOut();
    if (error) {
      console.error("Error signing out:", error);
    }
  };

  const value = {
    user,
    session,
    loading,
    signOut,
  };

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
};
