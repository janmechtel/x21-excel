import { useEffect, useState } from "react";

import { Card, CardContent } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Checkbox } from "@/components/ui/checkbox";
import { SaveCopiesCheckbox } from "@/components/settings/SaveCopiesCheckbox";
import { ColumnWidthModeSelect } from "@/components/settings/ColumnWidthModeSelect";
import {
  AnthropicConfigFields,
  AzureConfigFields,
  InlineSpinner,
  ProviderSelect,
  type Provider,
} from "@/components/chat/settings/SettingsConfigFields";
import { webViewBridge } from "@/services/webViewBridge";

interface SettingsConfigDialogProps {
  open: boolean;
  onCancel: () => void;
  onSave: () => void;
  allowOtherWorkbookReads: boolean;
  onAllowOtherWorkbookReadsChange: (value: boolean) => void;
}

interface ConfigSnapshot {
  provider: Provider;
  isCustomEndpointEnabled: boolean;
  azure: {
    endpoint: string;
    apiKey: string;
    deploymentName: string;
    model: string;
    reasoningEffort: string;
  };
  anthropic: {
    baseUrl: string;
    apiKey: string;
    caBundlePath: string;
    model: string;
  };
}

const areSnapshotsEqual = (left: ConfigSnapshot, right: ConfigSnapshot) =>
  left.provider === right.provider &&
  left.isCustomEndpointEnabled === right.isCustomEndpointEnabled &&
  left.azure.endpoint === right.azure.endpoint &&
  left.azure.apiKey === right.azure.apiKey &&
  left.azure.deploymentName === right.azure.deploymentName &&
  left.azure.model === right.azure.model &&
  left.azure.reasoningEffort === right.azure.reasoningEffort &&
  left.anthropic.baseUrl === right.anthropic.baseUrl &&
  left.anthropic.apiKey === right.anthropic.apiKey &&
  left.anthropic.caBundlePath === right.anthropic.caBundlePath &&
  left.anthropic.model === right.anthropic.model;

// Helper to get API base URL from WebSocket URL
const getApiBaseUrl = async (): Promise<string> => {
  try {
    const wsUrl = await webViewBridge.getWebSocketUrl();
    // Convert ws://localhost:8085 to http://localhost:8085
    const httpUrl = wsUrl
      .replace("ws://", "http://")
      .replace("wss://", "https://");
    // Remove /ws path if present
    return httpUrl.replace(/\/ws$/, "");
  } catch (error) {
    console.warn("Failed to get WebSocket URL, using default:", error);
    return "http://localhost:8000";
  }
};

export function SettingsConfigDialog({
  open,
  onCancel,
  onSave,
  allowOtherWorkbookReads,
  onAllowOtherWorkbookReadsChange,
}: SettingsConfigDialogProps) {
  const [provider, setProvider] = useState<Provider>("azure_openai");
  const [isCustomEndpointEnabled, setIsCustomEndpointEnabled] = useState(false);
  const [endpoint, setEndpoint] = useState("");
  const [apiKey, setApiKey] = useState("");
  const [deploymentName, setDeploymentName] = useState("gpt-5.2");
  const [model, setModel] = useState("gpt-5.2");
  const [reasoningEffort, setReasoningEffort] = useState("medium");
  const [anthropicApiKey, setAnthropicApiKey] = useState("");
  const [anthropicBaseUrl, setAnthropicBaseUrl] = useState("");
  const [anthropicCaBundlePath, setAnthropicCaBundlePath] = useState("");
  const [anthropicModel, setAnthropicModel] = useState(
    "claude-sonnet-4-5-20250929",
  );
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [testing, setTesting] = useState(false);
  const [message, setMessage] = useState<{
    type: "error" | "success" | "info";
    text: string;
  } | null>(null);
  const [lastSavedSnapshot, setLastSavedSnapshot] =
    useState<ConfigSnapshot | null>(null);

  const currentSnapshot: ConfigSnapshot = {
    provider,
    isCustomEndpointEnabled,
    azure: {
      endpoint,
      apiKey,
      deploymentName,
      model,
      reasoningEffort,
    },
    anthropic: {
      baseUrl: anthropicBaseUrl,
      apiKey: anthropicApiKey,
      caBundlePath: anthropicCaBundlePath,
      model: anthropicModel,
    },
  };

  const isDirty =
    !lastSavedSnapshot ||
    !areSnapshotsEqual(currentSnapshot, lastSavedSnapshot);

  const hasAzureRequired =
    endpoint.trim().length > 0 && apiKey.trim().length > 0;
  const hasAnthropicRequired =
    anthropicBaseUrl.trim().length > 0 && anthropicApiKey.trim().length > 0;

  const testDisabledReason = !isCustomEndpointEnabled
    ? "Enable the custom endpoint to test."
    : isDirty
    ? "Save changes to enable testing."
    : provider === "azure_openai" && !hasAzureRequired
    ? "Enter the Azure endpoint and API key to test."
    : provider === "anthropic" && !hasAnthropicRequired
    ? "Enter the base URL and API key to test."
    : null;

  const validationError = !isCustomEndpointEnabled
    ? null
    : provider === "azure_openai" && !hasAzureRequired
    ? "Azure endpoint and API key are required to save."
    : provider === "anthropic" && !hasAnthropicRequired
    ? "Base URL and API key are required to save."
    : null;

  // Load existing config when dialog opens
  useEffect(() => {
    if (open) {
      setIsInitialLoad(true);
      loadConfigs();
    }
  }, [open]);

  const loadConfigs = async () => {
    setLoading(true);
    setMessage(null);

    try {
      const baseUrl = await getApiBaseUrl();
      const [azureResponse, anthropicResponse] = await Promise.all([
        fetch(`${baseUrl}/api/llm-config?provider=azure_openai`),
        fetch(`${baseUrl}/api/llm-config?provider=anthropic`),
      ]);
      const [azureData, anthropicData] = await Promise.all([
        azureResponse.json(),
        anthropicResponse.json(),
      ]);

      const azureConfig = azureData.config ?? {};
      const anthropicConfig = anthropicData.config ?? {};

      const nextAzure = {
        endpoint: azureConfig.azureOpenaiEndpoint || "",
        apiKey: azureConfig.azureOpenaiKey || "",
        deploymentName: azureConfig.azureOpenaiDeploymentName || "gpt-5.2",
        model:
          azureConfig.azureOpenaiModel ||
          azureConfig.azureOpenaiDeploymentName ||
          "gpt-5.2",
        reasoningEffort: (
          azureConfig.openaiReasoningEffort || "medium"
        ).toLowerCase(),
        isActive: !!azureConfig.isActive,
      };

      const nextAnthropic = {
        apiKey: anthropicConfig.anthropicApiKey || "",
        baseUrl: anthropicConfig.anthropicBaseUrl || "",
        caBundlePath: anthropicConfig.anthropicCaBundlePath || "",
        model: anthropicConfig.anthropicModel || "claude-sonnet-4-5-20250929",
        isActive: !!anthropicConfig.isActive,
      };

      const nextIsCustomEndpointEnabled =
        nextAzure.isActive || nextAnthropic.isActive;
      const nextProvider: Provider =
        nextAnthropic.isActive && !nextAzure.isActive
          ? "anthropic"
          : "azure_openai";

      setProvider(nextProvider);
      setIsCustomEndpointEnabled(nextIsCustomEndpointEnabled);
      setEndpoint(nextAzure.endpoint);
      setApiKey(nextAzure.apiKey);
      setDeploymentName(nextAzure.deploymentName);
      setModel(nextAzure.model);
      setReasoningEffort(nextAzure.reasoningEffort);
      setAnthropicApiKey(nextAnthropic.apiKey);
      setAnthropicBaseUrl(nextAnthropic.baseUrl);
      setAnthropicCaBundlePath(nextAnthropic.caBundlePath);
      setAnthropicModel(nextAnthropic.model);

      setLastSavedSnapshot({
        provider: nextProvider,
        isCustomEndpointEnabled: nextIsCustomEndpointEnabled,
        azure: {
          endpoint: nextAzure.endpoint,
          apiKey: nextAzure.apiKey,
          deploymentName: nextAzure.deploymentName,
          model: nextAzure.model,
          reasoningEffort: nextAzure.reasoningEffort,
        },
        anthropic: {
          baseUrl: nextAnthropic.baseUrl,
          apiKey: nextAnthropic.apiKey,
          caBundlePath: nextAnthropic.caBundlePath,
          model: nextAnthropic.model,
        },
      });
    } catch (error: any) {
      console.error("Failed to load LLM config:", error);
      setMessage({
        type: "error",
        text: "Failed to load configuration",
      });
    } finally {
      setLoading(false);
      setIsInitialLoad(false);
    }
  };

  const persistConfig = async (
    targetProvider: Provider,
    overrides: Partial<{
      isActive: boolean;
      endpoint: string;
      apiKey: string;
      deploymentName: string;
      model: string;
      reasoningEffort: string;
      anthropicApiKey: string;
      anthropicBaseUrl: string;
      anthropicCaBundlePath: string;
      anthropicModel: string;
    }> = {},
  ) => {
    const nextIsActive = overrides.isActive ?? false;
    const nextEndpoint = overrides.endpoint ?? endpoint;
    const nextApiKey = overrides.apiKey ?? apiKey;
    const nextDeploymentName = overrides.deploymentName ?? deploymentName;
    const nextModel = overrides.model ?? model;
    const nextReasoningEffort = overrides.reasoningEffort ?? reasoningEffort;
    const nextAnthropicApiKey = overrides.anthropicApiKey ?? anthropicApiKey;
    const nextAnthropicBaseUrl = overrides.anthropicBaseUrl ?? anthropicBaseUrl;
    const nextAnthropicCaBundlePath =
      overrides.anthropicCaBundlePath ?? anthropicCaBundlePath;
    const nextAnthropicModel = overrides.anthropicModel ?? anthropicModel;
    const sanitizedAnthropicApiKey = nextAnthropicApiKey.trim()
      ? nextAnthropicApiKey
      : null;
    const sanitizedAnthropicBaseUrl = nextAnthropicBaseUrl.trim()
      ? nextAnthropicBaseUrl
      : null;
    const sanitizedAnthropicCaBundlePath = nextAnthropicCaBundlePath.trim()
      ? nextAnthropicCaBundlePath
      : null;
    const sanitizedAnthropicModel = nextAnthropicModel.trim()
      ? nextAnthropicModel
      : null;

    const baseUrl = await getApiBaseUrl();
    const payload =
      targetProvider === "azure_openai"
        ? {
            provider: "azure_openai",
            azureOpenaiEndpoint: nextEndpoint,
            azureOpenaiKey: nextApiKey,
            azureOpenaiDeploymentName: nextDeploymentName,
            azureOpenaiModel: nextModel,
            openaiReasoningEffort: nextReasoningEffort,
            isActive: nextIsActive,
          }
        : {
            provider: "anthropic",
            anthropicApiKey: sanitizedAnthropicApiKey,
            anthropicBaseUrl: sanitizedAnthropicBaseUrl,
            anthropicCaBundlePath: sanitizedAnthropicCaBundlePath,
            anthropicModel: sanitizedAnthropicModel,
            isActive: nextIsActive,
          };
    const response = await fetch(`${baseUrl}/api/llm-config`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error("Failed to save configuration");
    }
  };

  const handleSave = async () => {
    if (validationError) {
      setMessage({
        type: "error",
        text: validationError,
      });
      return;
    }

    setSaving(true);
    setMessage(null);

    try {
      const activeProvider = isCustomEndpointEnabled ? provider : null;

      await persistConfig("azure_openai", {
        isActive: activeProvider === "azure_openai",
      });
      await persistConfig("anthropic", {
        isActive: activeProvider === "anthropic",
      });

      const baseUrl = await getApiBaseUrl();
      const reloadResponse = await fetch(`${baseUrl}/api/llm-config/reload`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
      });

      if (!reloadResponse.ok) {
        console.warn("Failed to reload LLM config, but save was successful");
      }

      setLastSavedSnapshot(currentSnapshot);
      onSave();
    } catch (error: any) {
      console.error("Failed to save LLM config:", error);
      setMessage({
        type: "error",
        text: error.message || "Failed to save configuration",
      });
    } finally {
      setSaving(false);
    }
  };

  const handlePickAnthropicCaBundle = async () => {
    try {
      const picked = await webViewBridge.pickFile({
        extensions: [".pem", ".crt", ".cer"],
        title: "Select CA bundle",
        filterLabel: "CA bundle files",
      });
      if (picked) {
        setAnthropicCaBundlePath(picked);
      }
    } catch (error: any) {
      console.warn("Failed to pick CA bundle:", error);
      setMessage({
        type: "error",
        text: "Failed to pick CA bundle file",
      });
    }
  };

  const handleTestConnection = async () => {
    if (testDisabledReason) {
      setMessage({
        type: "info",
        text: testDisabledReason,
      });
      return;
    }

    setTesting(true);
    setMessage(null);

    try {
      const baseUrl = await getApiBaseUrl();
      const response = await fetch(`${baseUrl}/api/llm-config/test`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ provider }),
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        setMessage({
          type: "error",
          text: data.error || "Connection test failed",
        });
      } else {
        setMessage({
          type: "success",
          text: `✓ Configuration loaded! Check console for details. Model: ${
            data.config?.model || "unknown"
          }`,
        });
        console.log("LLM Connection Test:", data.config);
      }
    } catch (error: any) {
      console.error("Failed to test connection:", error);
      setMessage({
        type: "error",
        text: error.message || "Failed to test connection",
      });
    } finally {
      setTesting(false);
    }
  };

  const handleCancel = () => {
    setMessage(null);
    onCancel();
  };

  if (!open) return null;

  const testDisabled = !!testDisabledReason || saving || testing;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[9999]">
      <Card className="w-full max-w-md mx-4 max-h-[85vh]">
        <CardContent className="p-6 max-h-[85vh] overflow-y-auto">
          <h2 className="text-lg font-semibold mb-4">Settings</h2>

          {message && (
            <Alert
              className={`mb-4 ${
                message.type === "error"
                  ? "border-red-200 bg-red-50"
                  : message.type === "info"
                  ? "border-blue-200 bg-blue-50"
                  : "border-green-200 bg-green-50"
              }`}
            >
              <AlertDescription
                className={
                  message.type === "error"
                    ? "text-red-800"
                    : message.type === "info"
                    ? "text-blue-800"
                    : "text-green-800"
                }
              >
                {message.text}
              </AlertDescription>
            </Alert>
          )}

          {isInitialLoad && loading ? (
            <div className="mb-4">
              <InlineSpinner label="Loading settings..." />
            </div>
          ) : null}

          <>
            {/* General Settings Section */}
            <div className="mb-6">
              <h3 className="text-sm font-semibold text-slate-700 dark:text-slate-300 mb-3">
                General
              </h3>
              <div className="space-y-3">
                <SaveCopiesCheckbox disabled={saving || loading} />
                <Checkbox
                  id="allow-other-workbooks"
                  checked={allowOtherWorkbookReads}
                  onChange={(e) =>
                    onAllowOtherWorkbookReadsChange(e.target.checked)
                  }
                  disabled={saving || loading}
                  label="Include file names of open workbooks"
                  tooltip="Helps when you refer to other workbooks by name, but may include file names of potentially unrelated workbooks in the prompt context."
                />
              </div>
            </div>

            <div className="mb-6">
              <div className="space-y-3">
                <ColumnWidthModeSelect disabled={saving || loading} />
              </div>
            </div>

            <div className="mb-6 space-y-4">
              <div className="flex items-center justify-between">
                <Checkbox
                  id="use-custom-endpoint"
                  name="useCustomEndpoint"
                  checked={isCustomEndpointEnabled}
                  onChange={(e) => setIsCustomEndpointEnabled(e.target.checked)}
                  disabled={saving || loading}
                  label="Use custom endpoint"
                  tooltip="Use your own provider configuration instead of the default settings."
                />
                {loading ? <InlineSpinner label="Refreshing..." /> : null}
              </div>

              {isCustomEndpointEnabled ? (
                <div className="space-y-4">
                  <ProviderSelect
                    provider={provider}
                    onChange={setProvider}
                    disabled={saving || loading}
                  />

                  <div>
                    <h4 className="text-sm font-semibold text-slate-700 dark:text-slate-300 mb-3">
                      {provider === "azure_openai"
                        ? "Azure OpenAI"
                        : "Anthropic"}
                    </h4>
                    {provider === "azure_openai" ? (
                      <AzureConfigFields
                        endpoint={endpoint}
                        apiKey={apiKey}
                        deploymentName={deploymentName}
                        model={model}
                        reasoningEffort={reasoningEffort}
                        onEndpointChange={setEndpoint}
                        onApiKeyChange={setApiKey}
                        onDeploymentNameChange={setDeploymentName}
                        onModelChange={setModel}
                        onReasoningEffortChange={setReasoningEffort}
                        disabled={saving || loading}
                      />
                    ) : (
                      <AnthropicConfigFields
                        baseUrl={anthropicBaseUrl}
                        apiKey={anthropicApiKey}
                        caBundlePath={anthropicCaBundlePath}
                        model={anthropicModel}
                        onBaseUrlChange={setAnthropicBaseUrl}
                        onApiKeyChange={setAnthropicApiKey}
                        onCaBundlePathChange={setAnthropicCaBundlePath}
                        onModelChange={setAnthropicModel}
                        onPickCaBundle={handlePickAnthropicCaBundle}
                        onClearCaBundle={() => setAnthropicCaBundlePath("")}
                        disabled={saving || loading}
                      />
                    )}
                  </div>
                </div>
              ) : null}
            </div>

            <div className="flex justify-between items-start gap-3">
              <div className="flex flex-col items-start gap-1">
                {isCustomEndpointEnabled ? (
                  <>
                    <Button
                      variant="outline"
                      onClick={handleTestConnection}
                      disabled={testDisabled}
                      className="border-blue-300 text-blue-700 hover:bg-blue-50"
                    >
                      {testing ? (
                        <div className="flex items-center">
                          <div className="w-4 h-4 border-2 border-blue-600 border-t-transparent rounded-full animate-spin mr-2"></div>
                          Testing...
                        </div>
                      ) : (
                        "Test Connection"
                      )}
                    </Button>
                    {testDisabledReason && !testing ? (
                      <span className="text-xs text-slate-500">
                        {testDisabledReason}
                      </span>
                    ) : null}
                  </>
                ) : null}
              </div>
              <div className="flex gap-3">
                <Button
                  variant="outline"
                  onClick={handleCancel}
                  disabled={saving || testing}
                >
                  Cancel
                </Button>
                <Button
                  onClick={handleSave}
                  className="bg-blue-600 hover:bg-blue-700 text-white"
                  disabled={saving || testing || !!validationError}
                >
                  {saving ? (
                    <div className="flex items-center">
                      <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin mr-2"></div>
                      Saving...
                    </div>
                  ) : (
                    "Save"
                  )}
                </Button>
              </div>
            </div>
          </>
        </CardContent>
      </Card>
    </div>
  );
}
