import { type ReactNode } from "react";

import { Button } from "@/components/ui/button";
import { InfoTooltip } from "@/components/ui/info-tooltip";

export type Provider = "azure_openai" | "anthropic";

const PROVIDER_OPTIONS: { value: Provider; label: string }[] = [
  { value: "azure_openai", label: "Azure OpenAI" },
  { value: "anthropic", label: "Anthropic" },
];

const baseInputClassName =
  "w-full px-3 py-2 border border-slate-300 dark:border-slate-600 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100";

const getInputClassName = (disabled: boolean, extra?: string) =>
  `${baseInputClassName} ${disabled ? "opacity-50 cursor-not-allowed" : ""} ${
    extra ?? ""
  }`.trim();

interface FormFieldProps {
  id: string;
  label: string;
  tooltip?: string;
  children: ReactNode;
}

const FormField = ({ id, label, tooltip, children }: FormFieldProps) => (
  <div>
    <label
      htmlFor={id}
      className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300 mb-1"
    >
      <span>{label}</span>
      {tooltip ? <InfoTooltip content={tooltip} /> : null}
    </label>
    {children}
  </div>
);

export const InlineSpinner = ({ label }: { label: string }) => (
  <div className="flex items-center gap-2 text-xs text-slate-500">
    <div className="w-3.5 h-3.5 border-2 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
    {label}
  </div>
);

interface ProviderSelectProps {
  provider: Provider;
  onChange: (value: Provider) => void;
  disabled: boolean;
}

export const ProviderSelect = ({
  provider,
  onChange,
  disabled,
}: ProviderSelectProps) => (
  <FormField
    id="provider"
    label="Provider"
    tooltip="Select the provider to use for the custom endpoint."
  >
    <select
      id="provider"
      value={provider}
      onChange={(e) => onChange(e.target.value as Provider)}
      className={getInputClassName(disabled)}
      disabled={disabled}
    >
      {PROVIDER_OPTIONS.map((option) => (
        <option key={option.value} value={option.value}>
          {option.label}
        </option>
      ))}
    </select>
  </FormField>
);

interface AzureConfigFieldsProps {
  endpoint: string;
  apiKey: string;
  deploymentName: string;
  model: string;
  reasoningEffort: string;
  onEndpointChange: (value: string) => void;
  onApiKeyChange: (value: string) => void;
  onDeploymentNameChange: (value: string) => void;
  onModelChange: (value: string) => void;
  onReasoningEffortChange: (value: string) => void;
  disabled: boolean;
}

export const AzureConfigFields = ({
  endpoint,
  apiKey,
  deploymentName,
  model,
  reasoningEffort,
  onEndpointChange,
  onApiKeyChange,
  onDeploymentNameChange,
  onModelChange,
  onReasoningEffortChange,
  disabled,
}: AzureConfigFieldsProps) => (
  <div className="space-y-4">
    <FormField
      id="endpoint"
      label="Azure OpenAI Endpoint"
      tooltip="Base URL of your Azure OpenAI resource (no /openai/... path). Example: https://your-resource.openai.azure.com. Note: This app uses the Responses API (/openai/v1/responses) which requires GPT-5.x models and may not be available on all Azure endpoints."
    >
      <input
        id="endpoint"
        type="text"
        value={endpoint}
        onChange={(e) => onEndpointChange(e.target.value)}
        placeholder="https://your-resource.openai.azure.com/"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="apiKey"
      label="Azure API Key"
      tooltip="Key from your Azure OpenAI resource (Azure Portal -> Keys and Endpoint). This is stored locally."
    >
      <input
        id="apiKey"
        type="password"
        value={apiKey}
        onChange={(e) => onApiKeyChange(e.target.value)}
        placeholder="Enter your Azure API key"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="deploymentName"
      label="Deployment Name"
      tooltip="The deployment name you created in Azure (e.g., 'aihub-gpt-5.2', 'my-gpt-deployment'). This is what you see in Azure Portal under your deployments."
    >
      <input
        id="deploymentName"
        type="text"
        value={deploymentName}
        onChange={(e) => onDeploymentNameChange(e.target.value)}
        placeholder="aihub-gpt-5.2"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="model"
      label="Model"
      tooltip="The underlying model name (e.g., 'gpt-5.2', 'gpt-4.1'). This is the actual model your deployment uses. If unsure, it often matches the deployment name."
    >
      <input
        id="model"
        type="text"
        value={model}
        onChange={(e) => onModelChange(e.target.value)}
        placeholder="gpt-5.2"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="reasoningEffort"
      label="Reasoning Effort"
      tooltip="Controls how much reasoning the model is allowed to use. Higher can improve results but may be slower and cost more."
    >
      <select
        id="reasoningEffort"
        value={reasoningEffort}
        onChange={(e) => onReasoningEffortChange(e.target.value)}
        className={getInputClassName(disabled)}
        disabled={disabled}
      >
        <option value="low">Low</option>
        <option value="medium">Medium</option>
        <option value="high">High</option>
      </select>
    </FormField>
  </div>
);

interface AnthropicConfigFieldsProps {
  baseUrl: string;
  apiKey: string;
  caBundlePath: string;
  model: string;
  onBaseUrlChange: (value: string) => void;
  onApiKeyChange: (value: string) => void;
  onCaBundlePathChange: (value: string) => void;
  onModelChange: (value: string) => void;
  onPickCaBundle: () => void;
  onClearCaBundle: () => void;
  disabled: boolean;
}

export const AnthropicConfigFields = ({
  baseUrl,
  apiKey,
  caBundlePath,
  model,
  onBaseUrlChange,
  onApiKeyChange,
  onCaBundlePathChange,
  onModelChange,
  onPickCaBundle,
  onClearCaBundle,
  disabled,
}: AnthropicConfigFieldsProps) => (
  <div className="space-y-4">
    <FormField
      id="anthropicBaseUrl"
      label="Base URL"
      tooltip="Base URL for your Anthropic-compatible endpoint. Example: https://api.anthropic.com. For Azure Foundry Claude, use https://{resource}.services.ai.azure.com/anthropic (no /v1)."
    >
      <input
        id="anthropicBaseUrl"
        type="text"
        value={baseUrl}
        onChange={(e) => onBaseUrlChange(e.target.value)}
        placeholder="https://api.anthropic.com"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="anthropicApiKey"
      label="API Key"
      tooltip="Key from your Anthropic Console or Azure Foundry Keys and Endpoint. This is stored locally."
    >
      <input
        id="anthropicApiKey"
        type="password"
        value={apiKey}
        onChange={(e) => onApiKeyChange(e.target.value)}
        placeholder="Enter your Anthropic API key"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="anthropicModel"
      label="Model"
      tooltip="Anthropic model or Foundry deployment name (e.g., 'claude-sonnet-4-5-20250929' or 'claude-sonnet-4-5')."
    >
      <input
        id="anthropicModel"
        type="text"
        value={model}
        onChange={(e) => onModelChange(e.target.value)}
        placeholder="claude-sonnet-4-5-20250929"
        className={getInputClassName(disabled)}
        disabled={disabled}
      />
    </FormField>

    <FormField
      id="anthropicCaBundlePath"
      label="CA Bundle (PEM) File (optional)"
      tooltip="Optional CA bundle for TLS verification (useful for Azure Foundry internal endpoints). Select a .pem file."
    >
      <div className="flex gap-2">
        <input
          id="anthropicCaBundlePath"
          type="text"
          value={caBundlePath}
          onChange={(e) => onCaBundlePathChange(e.target.value)}
          placeholder="C:\\path\\to\\bundle.pem"
          className={getInputClassName(disabled, "flex-1")}
          disabled={disabled}
        />
        <Button
          type="button"
          variant="outline"
          onClick={onPickCaBundle}
          disabled={disabled}
        >
          Browse
        </Button>
        {caBundlePath ? (
          <Button
            type="button"
            variant="ghost"
            onClick={onClearCaBundle}
            disabled={disabled}
          >
            Clear
          </Button>
        ) : null}
      </div>
    </FormField>
  </div>
);
