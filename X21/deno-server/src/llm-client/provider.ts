/**
 * LLM Provider Utilities
 *
 * This module provides centralized configuration management for LLM providers.
 * Configuration is retrieved from the database first, then falls back to environment variables.
 */

import { getLlmKeysConfigByProvider } from "../db/llm-keys-dal.ts";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("LLMProvider");

export type LLMProvider = "anthropic" | "azure";
type ReasoningEffort = "low" | "medium" | "high";

export interface AzureOpenAIConfig {
  apiKey: string;
  endpoint: string;
  deploymentName: string;
  model: string;
  reasoningEffort: ReasoningEffort;
  isActive: boolean;
  apiVersion?: string;
  caBundlePath?: string;
}

export interface AnthropicConfig {
  apiKey: string;
  model: string;
  baseUrl?: string;
  caBundlePath?: string;
}

/**
 * Get Azure OpenAI configuration from database or environment variables
 * Priority: Database (if active) > Environment Variables
 */
export function getAzureOpenAIConfig(): AzureOpenAIConfig | null {
  const normalizeReasoningEffort = (value?: string | null): ReasoningEffort => {
    const normalized = (value || "").toLowerCase();
    if (normalized === "low" || normalized === "high") {
      return normalized;
    }
    return "medium";
  };

  // Try to get from database first (will return null if DB not initialized yet)
  const dbConfig = getLlmKeysConfigByProvider("azure_openai");

  // If database config exists and is active and complete, use it
  if (
    dbConfig && dbConfig.isActive && dbConfig.azureOpenaiEndpoint &&
    dbConfig.azureOpenaiKey
  ) {
    const deployment = dbConfig.azureOpenaiDeploymentName || "gpt-5.2";
    // Use separate model field if provided, otherwise fall back to deployment name
    const model = dbConfig.azureOpenaiModel || deployment;
    const reasoningEffort = normalizeReasoningEffort(
      dbConfig.openaiReasoningEffort,
    );
    const cleanEndpoint = dbConfig.azureOpenaiEndpoint.replace(/\/+$/, "");

    logger.info("✓ Azure OpenAI config loaded from database", {
      configId: dbConfig.id,
      provider: dbConfig.provider,
      endpoint: cleanEndpoint,
      deploymentName: deployment,
      model: model,
      modelSource: dbConfig.azureOpenaiModel
        ? "explicit"
        : "derived from deployment",
      reasoningEffort: reasoningEffort,
      hasApiKey: !!dbConfig.azureOpenaiKey,
      apiKeyLength: dbConfig.azureOpenaiKey?.length || 0,
      apiKeyPrefix: dbConfig.azureOpenaiKey?.substring(0, 4) + "...",
      isActive: dbConfig.isActive,
      fullBaseUrl: `${cleanEndpoint}/openai/v1/`,
    });

    return {
      apiKey: dbConfig.azureOpenaiKey,
      endpoint: cleanEndpoint,
      deploymentName: deployment,
      model,
      reasoningEffort,
      isActive: true,
      apiVersion: "2024-08-01-preview", // Default API version for Responses API
      caBundlePath: dbConfig.anthropicCaBundlePath?.trim() || undefined,
    };
  }

  logger.debug("⚠️ No Azure OpenAI configuration found", {
    dbConfigExists: !!dbConfig,
    dbConfigActive: dbConfig?.isActive,
    hasEndpoint: !!dbConfig?.azureOpenaiEndpoint,
    hasKey: !!dbConfig?.azureOpenaiKey,
  });
  return null;
}

/**
 * Get Anthropic configuration from database or environment variables
 * Priority: Database (if active) > Environment Variables
 */
export function getAnthropicConfig(): AnthropicConfig | null {
  const defaultModel = "claude-sonnet-4-5-20250929";
  const defaultFoundryModel = "claude-sonnet-4-5";
  const normalizeBaseUrl = (value?: string | null): string | null => {
    if (!value) return null;
    const trimmed = value.trim();
    if (!trimmed) return null;

    if (!trimmed.includes("://")) {
      const isResource = !trimmed.includes(".");
      if (isResource) {
        return `https://${trimmed}.services.ai.azure.com/anthropic`;
      }
      return `https://${trimmed}`;
    }

    const normalized = trimmed.replace(/\/+$/, "");
    if (normalized.endsWith("/v1")) {
      return normalized.slice(0, -3);
    }
    return normalized;
  };
  const normalizeCaBundlePath = (value?: string | null): string | null => {
    if (!value) return null;
    const trimmed = value.trim();
    return trimmed ? trimmed : null;
  };

  const foundryBaseUrlFromEnv = normalizeBaseUrl(
    Deno.env.get("ANTHROPIC_FOUNDRY_BASE_URL"),
  );
  const foundryResourceFromEnv = normalizeBaseUrl(
    Deno.env.get("ANTHROPIC_FOUNDRY_RESOURCE"),
  );
  const baseUrlFromEnv = foundryBaseUrlFromEnv || foundryResourceFromEnv ||
    normalizeBaseUrl(Deno.env.get("ANTHROPIC_BASE_URL"));

  // Try to get from database first (will return null if DB not initialized yet)
  const dbConfig = getLlmKeysConfigByProvider("anthropic");

  if (dbConfig && dbConfig.isActive && dbConfig.anthropicApiKey) {
    const baseUrl = normalizeBaseUrl(dbConfig.anthropicBaseUrl) ||
      baseUrlFromEnv;
    const caBundlePath = normalizeCaBundlePath(
      dbConfig.anthropicCaBundlePath,
    );
    const useFoundryDefault = baseUrl?.includes(
      "services.ai.azure.com/anthropic",
    );
    const model = dbConfig.anthropicModel ||
      (useFoundryDefault ? defaultFoundryModel : defaultModel);
    logger.info("✓ Anthropic config loaded from database", {
      configId: dbConfig.id,
      provider: dbConfig.provider,
      model,
      baseUrl: baseUrl || "[default]",
      caBundlePath: caBundlePath || "[not set]",
      hasCaBundlePath: !!caBundlePath,
      hasApiKey: !!dbConfig.anthropicApiKey,
      apiKeyLength: dbConfig.anthropicApiKey?.length || 0,
      apiKeyPrefix: dbConfig.anthropicApiKey?.substring(0, 4) + "...",
      isActive: dbConfig.isActive,
    });

    return {
      apiKey: dbConfig.anthropicApiKey,
      model,
      baseUrl: baseUrl || undefined,
      caBundlePath: caBundlePath || undefined,
    };
  }

  // Fall back to environment variables
  const apiKey = Deno.env.get("ANTHROPIC_API_KEY");

  if (!apiKey || apiKey.trim().length === 0) {
    logger.warn("No Anthropic API key found");
    return null;
  }

  const useFoundryDefault = baseUrlFromEnv?.includes(
    "services.ai.azure.com/anthropic",
  );

  return {
    apiKey,
    model: useFoundryDefault ? defaultFoundryModel : defaultModel,
    baseUrl: baseUrlFromEnv || undefined,
  };
}

/**
 * Determines which LLM provider to use based on configuration availability.
 * Checks database first (is_active = 1 and provider = 'azure_openai'), then environment variables.
 *
 * @returns "azure" if Azure OpenAI config is available and active, otherwise "anthropic"
 */
export function getLLMProvider(): LLMProvider {
  //Always check from local db first
  // Auto-detect based on available configuration
  const azureConfig = getAzureOpenAIConfig();
  if (azureConfig && azureConfig.isActive) {
    logger.info("Using Azure OpenAI as provider");
    return "azure";
  }

  const anthropicConfig = getAnthropicConfig();
  if (anthropicConfig) {
    logger.info("Using Anthropic as provider");
    return "anthropic";
  }

  // Check explicit provider setting first
  const envProvider = (Deno.env.get("LLM_PROVIDER") || "").toLowerCase();

  if (envProvider === "azure" || envProvider === "anthropic") {
    return envProvider as LLMProvider;
  }

  // Default to anthropic
  logger.info("No valid config found, defaulting to Anthropic");
  return "anthropic";
}

/**
 * Mask API key for logging (show first 30 characters)
 */
function maskApiKey(key: string | undefined): string {
  if (!key || key.trim().length === 0) {
    return "[NOT SET]";
  }
  const trimmed = key.trim();
  if (trimmed.length <= 30) {
    return trimmed.substring(0, Math.min(10, trimmed.length)) + "...";
  }
  return trimmed.substring(0, 30) + "...";
}

/**
 * Get LLM configuration details for logging
 */
export function getLLMConfig(): {
  provider: LLMProvider;
  model: string;
  apiKey: string;
  endpoint?: string;
} {
  const provider = getLLMProvider();

  if (provider === "azure") {
    const config = getAzureOpenAIConfig();

    if (config) {
      logger.debug("🔍 Retrieved Azure OpenAI config for logging", {
        deploymentName: config.deploymentName,
        model: config.model,
        endpoint: config.endpoint,
        reasoningEffort: config.reasoningEffort,
      });

      return {
        provider,
        model: config.model,
        apiKey: maskApiKey(config.apiKey),
        endpoint: config.endpoint,
      };
    }

    logger.warn("⚠️ Azure provider selected but no config available");

    // Fallback if config not available
    return {
      provider,
      model: "[NOT SET]",
      apiKey: "[NOT SET]",
      endpoint: "[NOT SET]",
    };
  }

  // Anthropic
  const config = getAnthropicConfig();

  if (config) {
    return {
      provider,
      model: config.model,
      apiKey: maskApiKey(config.apiKey),
      endpoint: config.baseUrl,
    };
  }

  // Fallback if config not available
  return {
    provider,
    model: "[NOT SET]",
    apiKey: "[NOT SET]",
  };
}

/**
 * Reload LLM configuration and log the updated settings
 * This is called after configuration changes to pick up new values from the database
 */
export function reloadLLMConfig(): void {
  logger.info("🔄 Reloading LLM configuration...");

  const provider = getLLMProvider();
  const llmConfig = getLLMConfig();

  if (provider === "azure") {
    const azureConfig = getAzureOpenAIConfig();
    if (azureConfig) {
      logger.info("✓ LLM Configuration reloaded (Azure OpenAI):", {
        provider: llmConfig.provider,
        endpoint: llmConfig.endpoint,
        deploymentName: azureConfig.deploymentName,
        model: llmConfig.model,
        modelSource: azureConfig.azureOpenaiModel
          ? "explicit"
          : "from deployment",
        reasoningEffort: azureConfig.reasoningEffort,
        apiKey: llmConfig.apiKey,
        isActive: azureConfig.isActive,
      });
    } else {
      logger.warn("⚠️ LLM Configuration reload: Azure selected but no config");
    }
  } else {
    logger.info("✓ LLM Configuration reloaded:", {
      provider: llmConfig.provider,
      model: llmConfig.model,
      apiKey: llmConfig.apiKey,
      endpoint: llmConfig.endpoint,
    });
  }
}
