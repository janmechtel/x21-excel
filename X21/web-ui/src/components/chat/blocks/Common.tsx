import type { ReactNode } from "react";

export const IconContainer = ({ children }: { children: ReactNode }) => (
  <div className="w-4 h-4 flex items-center justify-center flex-shrink-0">
    {children}
  </div>
);

export const WaveAnimation = ({ className = "" }: { className?: string }) => (
  <div className={`flex items-center gap-0.5 ${className}`}>
    <div
      className="w-0.5 h-0.5 bg-current rounded-full animate-pulse"
      style={{ animationDelay: "0ms", animationDuration: "1.4s" }}
    />
    <div
      className="w-0.5 h-0.5 bg-current rounded-full animate-pulse"
      style={{ animationDelay: "200ms", animationDuration: "1.4s" }}
    />
    <div
      className="w-0.5 h-0.5 bg-current rounded-full animate-pulse"
      style={{ animationDelay: "400ms", animationDuration: "1.4s" }}
    />
  </div>
);
