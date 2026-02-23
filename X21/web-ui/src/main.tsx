import "./posthog";
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.tsx";
import "./index.css";
import { AuthProvider } from "./contexts/AuthContext";
import { initializeMockAuth } from "./utils/mockAuth";

// Initialize mock auth if VITE_SKIP_AUTH is enabled
initializeMockAuth();

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <AuthProvider>
      <App />
    </AuthProvider>
  </React.StrictMode>,
);
