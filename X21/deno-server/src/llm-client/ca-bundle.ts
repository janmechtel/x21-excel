import { createLogger } from "../utils/logger.ts";

const logger = createLogger("CA-Bundle");

const caBundleFetchCache = new Map<string, typeof fetch>();

/**
 * Returns a custom fetch bound to an HTTP client that trusts the PEM file at
 * `caBundlePath`, suitable for corporate proxy / TLS-intercept environments.
 *
 * Returns `undefined` when no path is provided so callers can conditionally
 * spread `{ fetch: fetchOverride }` without branching.
 *
 * The result is cached by path so repeated calls (e.g. per-request client
 * construction) do not re-read disk or re-allocate HttpClient handles.
 */
export function getFetchWithCaBundle(
  caBundlePath?: string,
): typeof fetch | undefined {
  if (!caBundlePath) {
    logger.info("CA bundle not configured; using default TLS trust");
    return undefined;
  }

  const cached = caBundleFetchCache.get(caBundlePath);
  if (cached) {
    logger.info("Using cached CA bundle HTTP client", { caBundlePath });
    return cached;
  }

  try {
    logger.info("Loading CA bundle from disk", { caBundlePath });
    const pemContents = Deno.readTextFileSync(caBundlePath);
    if (!pemContents.trim()) {
      logger.error("CA bundle file is empty", { caBundlePath });
      throw new Error("CA bundle file is empty");
    }
    logger.info("CA bundle loaded", {
      caBundlePath,
      pemLength: pemContents.length,
    });

    const httpClient = Deno.createHttpClient({ caCerts: [pemContents] });
    const fetchWithClient: typeof fetch = (input, init) => {
      const initWithClient = {
        ...(init ?? {}),
      } as RequestInit & { client: Deno.HttpClient };
      initWithClient.client = httpClient;
      return fetch(input, initWithClient);
    };

    caBundleFetchCache.set(caBundlePath, fetchWithClient);
    logger.info("CA bundle HTTP client ready", { caBundlePath });
    return fetchWithClient;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error("Failed to load CA bundle", { caBundlePath, error: message });
    throw error;
  }
}
