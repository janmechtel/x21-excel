import { getDb, isDbInitialized, nowMs } from "./sqlite.ts";
import { createLogger } from "../utils/logger.ts";

const logger = createLogger("LlmKeysDAL");

export interface LlmKeysConfig {
  id?: number;
  provider: string;
  azureOpenaiEndpoint: string | null;
  azureOpenaiKey: string | null;
  azureOpenaiDeploymentName: string | null;
  azureOpenaiModel: string | null;
  openaiReasoningEffort: string | null;
  anthropicApiKey: string | null;
  anthropicModel: string | null;
  anthropicBaseUrl: string | null;
  anthropicCaBundlePath: string | null;
  isActive: boolean;
  addedDate: number;
  modifiedDate: number | null;
}

export interface LlmKeysConfigRow {
  id: number;
  provider: string;
  azure_openai_endpoint: string | null;
  azure_openai_key: string | null;
  azure_openai_deployment_name: string | null;
  azure_openai_model: string | null;
  openai_reasoning_effort: string | null;
  anthropic_api_key: string | null;
  anthropic_model: string | null;
  anthropic_base_url: string | null;
  anthropic_ca_bundle_path: string | null;
  is_active: number;
  added_date: number;
  modified_date: number | null;
}

/**
 * Insert a new LLM keys configuration
 */
export function insertLlmKeysConfig(config: {
  provider?: string;
  azureOpenaiEndpoint?: string | null;
  azureOpenaiKey?: string | null;
  azureOpenaiDeploymentName?: string | null;
  azureOpenaiModel?: string | null;
  openaiReasoningEffort?: string | null;
  anthropicApiKey?: string | null;
  anthropicModel?: string | null;
  anthropicBaseUrl?: string | null;
  anthropicCaBundlePath?: string | null;
  isActive?: boolean;
}): number {
  const db = getDb();
  const now = nowMs();

  const result = db.query<[number]>(
    `INSERT INTO llm_keys_config
     (provider, azure_openai_endpoint, azure_openai_key, azure_openai_deployment_name, azure_openai_model, openai_reasoning_effort, anthropic_api_key, anthropic_model, anthropic_base_url, anthropic_ca_bundle_path, is_active, added_date)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
     RETURNING id`,
    [
      config.provider || "azure",
      config.azureOpenaiEndpoint || null,
      config.azureOpenaiKey || null,
      config.azureOpenaiDeploymentName || null,
      config.azureOpenaiModel || null,
      config.openaiReasoningEffort || null,
      config.anthropicApiKey || null,
      config.anthropicModel || null,
      config.anthropicBaseUrl || null,
      config.anthropicCaBundlePath || null,
      config.isActive === false ? 0 : 1,
      now,
    ],
  );

  const id = result[0]?.[0];
  logger.info(`Inserted LLM keys config with ID: ${id}`);
  return id;
}

/**
 * Update an existing LLM keys configuration
 */
export function updateLlmKeysConfig(
  id: number,
  config: {
    provider?: string;
    azureOpenaiEndpoint?: string | null;
    azureOpenaiKey?: string | null;
    azureOpenaiDeploymentName?: string | null;
    azureOpenaiModel?: string | null;
    openaiReasoningEffort?: string | null;
    anthropicApiKey?: string | null;
    anthropicModel?: string | null;
    anthropicBaseUrl?: string | null;
    anthropicCaBundlePath?: string | null;
    isActive?: boolean;
  },
): void {
  const db = getDb();
  const now = nowMs();

  db.query(
    `UPDATE llm_keys_config
     SET provider = ?,
         azure_openai_endpoint = ?,
         azure_openai_key = ?,
         azure_openai_deployment_name = ?,
         azure_openai_model = ?,
         openai_reasoning_effort = ?,
         anthropic_api_key = ?,
         anthropic_model = ?,
         anthropic_base_url = ?,
         anthropic_ca_bundle_path = ?,
         is_active = ?,
         modified_date = ?
     WHERE id = ?`,
    [
      config.provider || null,
      config.azureOpenaiEndpoint !== undefined
        ? config.azureOpenaiEndpoint
        : null,
      config.azureOpenaiKey !== undefined ? config.azureOpenaiKey : null,
      config.azureOpenaiDeploymentName !== undefined
        ? config.azureOpenaiDeploymentName
        : null,
      config.azureOpenaiModel !== undefined ? config.azureOpenaiModel : null,
      config.openaiReasoningEffort !== undefined
        ? config.openaiReasoningEffort
        : null,
      config.anthropicApiKey !== undefined ? config.anthropicApiKey : null,
      config.anthropicModel !== undefined ? config.anthropicModel : null,
      config.anthropicBaseUrl !== undefined ? config.anthropicBaseUrl : null,
      config.anthropicCaBundlePath !== undefined
        ? config.anthropicCaBundlePath
        : null,
      config.isActive !== undefined ? (config.isActive ? 1 : 0) : null,
      now,
      id,
    ],
  );

  logger.info(`Updated LLM keys config ID: ${id}`);
}

/**
 * Get LLM keys configuration by ID
 */
export function getLlmKeysConfig(id: number): LlmKeysConfig | null {
  const db = getDb();

  const rows = db.query<
    [
      number,
      string,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      number,
      number,
      number | null,
    ]
  >(
    `SELECT id, provider, azure_openai_endpoint, azure_openai_key, azure_openai_deployment_name, azure_openai_model, openai_reasoning_effort, anthropic_api_key, anthropic_model, anthropic_base_url, anthropic_ca_bundle_path, is_active, added_date, modified_date
     FROM llm_keys_config
     WHERE id = ?`,
    [id],
  );

  if (rows.length === 0) {
    return null;
  }

  const [
    id_,
    provider,
    endpoint,
    key,
    deploymentName,
    model,
    reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive,
    addedDate,
    modifiedDate,
  ] = rows[0];
  return {
    id: id_,
    provider,
    azureOpenaiEndpoint: endpoint,
    azureOpenaiKey: key,
    azureOpenaiDeploymentName: deploymentName,
    azureOpenaiModel: model,
    openaiReasoningEffort: reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive: isActive === 1,
    addedDate,
    modifiedDate,
  };
}

/**
 * Get the most recent LLM keys configuration
 */
export function getLatestLlmKeysConfig(): LlmKeysConfig | null {
  const db = getDb();

  const rows = db.query<
    [
      number,
      string,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      number,
      number,
      number | null,
    ]
  >(
    `SELECT id, provider, azure_openai_endpoint, azure_openai_key, azure_openai_deployment_name, azure_openai_model, openai_reasoning_effort, anthropic_api_key, anthropic_model, anthropic_base_url, anthropic_ca_bundle_path, is_active, added_date, modified_date
     FROM llm_keys_config
     ORDER BY added_date DESC
     LIMIT 1`,
  );

  if (rows.length === 0) {
    return null;
  }

  const [
    id,
    provider,
    endpoint,
    key,
    deploymentName,
    model,
    reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive,
    addedDate,
    modifiedDate,
  ] = rows[0];
  return {
    id,
    provider,
    azureOpenaiEndpoint: endpoint,
    azureOpenaiKey: key,
    azureOpenaiDeploymentName: deploymentName,
    azureOpenaiModel: model,
    openaiReasoningEffort: reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive: isActive === 1,
    addedDate,
    modifiedDate,
  };
}

/**
 * List all LLM keys configurations
 */
export function listAllLlmKeysConfigs(): LlmKeysConfig[] {
  const db = getDb();

  const rows = db.query<
    [
      number,
      string,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      number,
      number,
      number | null,
    ]
  >(
    `SELECT id, provider, azure_openai_endpoint, azure_openai_key, azure_openai_deployment_name, azure_openai_model, openai_reasoning_effort, anthropic_api_key, anthropic_model, anthropic_base_url, anthropic_ca_bundle_path, is_active, added_date, modified_date
     FROM llm_keys_config
     ORDER BY added_date DESC`,
  );

  return rows.map((
    [
      id,
      provider,
      endpoint,
      key,
      deploymentName,
      model,
      reasoningEffort,
      anthropicApiKey,
      anthropicModel,
      anthropicBaseUrl,
      anthropicCaBundlePath,
      isActive,
      addedDate,
      modifiedDate,
    ],
  ) => ({
    id,
    provider,
    azureOpenaiEndpoint: endpoint,
    azureOpenaiKey: key,
    azureOpenaiDeploymentName: deploymentName,
    azureOpenaiModel: model,
    openaiReasoningEffort: reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive: isActive === 1,
    addedDate,
    modifiedDate,
  }));
}

/**
 * Get LLM keys configuration by provider
 */
export function getLlmKeysConfigByProvider(
  provider: string,
): LlmKeysConfig | null {
  // Check if database is initialized before trying to access it
  if (!isDbInitialized()) {
    logger.debug("Database not initialized yet, skipping database lookup");
    return null;
  }

  const db = getDb();

  const rows = db.query<
    [
      number,
      string,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      string | null,
      number,
      number,
      number | null,
    ]
  >(
    `SELECT id, provider, azure_openai_endpoint, azure_openai_key, azure_openai_deployment_name, azure_openai_model, openai_reasoning_effort, anthropic_api_key, anthropic_model, anthropic_base_url, anthropic_ca_bundle_path, is_active, added_date, modified_date
     FROM llm_keys_config
     WHERE provider = ?
     ORDER BY modified_date DESC, added_date DESC
     LIMIT 1`,
    [provider],
  );

  if (rows.length === 0) {
    return null;
  }

  const [
    id,
    prov,
    endpoint,
    key,
    deploymentName,
    model,
    reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive,
    addedDate,
    modifiedDate,
  ] = rows[0];
  return {
    id,
    provider: prov,
    azureOpenaiEndpoint: endpoint,
    azureOpenaiKey: key,
    azureOpenaiDeploymentName: deploymentName,
    azureOpenaiModel: model,
    openaiReasoningEffort: reasoningEffort,
    anthropicApiKey,
    anthropicModel,
    anthropicBaseUrl,
    anthropicCaBundlePath,
    isActive: isActive === 1,
    addedDate,
    modifiedDate,
  };
}

/**
 * Upsert LLM keys configuration by provider
 * If a config with the provider exists, update it; otherwise insert a new one
 */
export function upsertLlmKeysConfigByProvider(config: {
  provider: string;
  azureOpenaiEndpoint: string | null;
  azureOpenaiKey: string | null;
  azureOpenaiDeploymentName: string | null;
  azureOpenaiModel: string | null;
  openaiReasoningEffort: string | null;
  anthropicApiKey: string | null;
  anthropicModel: string | null;
  anthropicBaseUrl: string | null;
  anthropicCaBundlePath: string | null;
  isActive: boolean;
}): number {
  const existing = getLlmKeysConfigByProvider(config.provider);

  if (existing && existing.id) {
    logger.info("🔄 DB: Updating existing LLM configuration", {
      configId: existing.id,
      provider: config.provider,
      endpoint: config.azureOpenaiEndpoint || "[not set]",
      deploymentName: config.azureOpenaiDeploymentName || "[not set]",
      model: config.azureOpenaiModel || "[not set]",
      modelChanged: existing.azureOpenaiModel !== config.azureOpenaiModel,
      deploymentChanged:
        existing.azureOpenaiDeploymentName !== config.azureOpenaiDeploymentName,
      hasApiKey: !!config.azureOpenaiKey,
      anthropicModel: config.anthropicModel || "[not set]",
      anthropicModelChanged: existing.anthropicModel !== config.anthropicModel,
      anthropicBaseUrl: config.anthropicBaseUrl || "[not set]",
      hasAnthropicCaBundlePath: !!config.anthropicCaBundlePath,
      hasAnthropicApiKey: !!config.anthropicApiKey,
      isActive: config.isActive,
    });

    updateLlmKeysConfig(existing.id, {
      provider: config.provider,
      azureOpenaiEndpoint: config.azureOpenaiEndpoint,
      azureOpenaiKey: config.azureOpenaiKey,
      azureOpenaiDeploymentName: config.azureOpenaiDeploymentName,
      azureOpenaiModel: config.azureOpenaiModel,
      openaiReasoningEffort: config.openaiReasoningEffort,
      anthropicApiKey: config.anthropicApiKey,
      anthropicModel: config.anthropicModel,
      anthropicBaseUrl: config.anthropicBaseUrl,
      anthropicCaBundlePath: config.anthropicCaBundlePath,
      isActive: config.isActive,
    });

    logger.info("✓ DB: Configuration updated successfully", {
      configId: existing.id,
      provider: config.provider,
    });

    return existing.id;
  } else {
    logger.info("➕ DB: Inserting new LLM configuration", {
      provider: config.provider,
      endpoint: config.azureOpenaiEndpoint || "[not set]",
      deploymentName: config.azureOpenaiDeploymentName || "[not set]",
      model: config.azureOpenaiModel || "[not set]",
      hasApiKey: !!config.azureOpenaiKey,
      anthropicModel: config.anthropicModel || "[not set]",
      anthropicBaseUrl: config.anthropicBaseUrl || "[not set]",
      hasAnthropicCaBundlePath: !!config.anthropicCaBundlePath,
      hasAnthropicApiKey: !!config.anthropicApiKey,
      isActive: config.isActive,
    });

    const id = insertLlmKeysConfig(config);

    logger.info("✓ DB: New configuration inserted successfully", {
      configId: id,
      provider: config.provider,
    });

    return id;
  }
}

/**
 * Delete an LLM keys configuration by ID
 */
export function deleteLlmKeysConfig(id: number): void {
  const db = getDb();

  db.query("DELETE FROM llm_keys_config WHERE id = ?", [id]);
  logger.info(`Deleted LLM keys config ID: ${id}`);
}
