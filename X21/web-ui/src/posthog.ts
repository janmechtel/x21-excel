// src/posthog.ts
import posthog from "posthog-js";

// Only initialize PostHog in production
const isProduction = import.meta.env.PROD;
const posthogApiKey = import.meta.env.VITE_POSTHOG_API_KEY;
const posthogApiHost =
  import.meta.env.VITE_POSTHOG_API_HOST ?? "https://eu.i.posthog.com";

if (isProduction && posthogApiKey) {
  posthog.init(posthogApiKey, {
    api_host: posthogApiHost,
    capture_pageview: true,
    defaults: "2025-05-24",
  });
  console.log("PostHog initialized for production");
} else if (isProduction) {
  console.log("PostHog disabled: missing VITE_POSTHOG_API_KEY");
} else {
  console.log("PostHog disabled in development mode");
}

export default posthog;
