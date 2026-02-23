import { UiRequestControl, UiRequestResponse } from "../types/ui-request.ts";

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

export function validateUiRequestResult(
  controls: UiRequestControl[] | undefined,
  output: unknown,
): UiRequestResponse {
  if (!output || typeof output !== "object") {
    throw new Error(
      "collect_input result must be an object keyed by control id",
    );
  }

  const answers = output as Record<string, any>;
  const sanitized: UiRequestResponse = {};

  for (const control of controls || []) {
    const answer = answers[control.id];

    if (answer === undefined || answer === null) {
      if (control.required) {
        throw new Error(`Missing required response for ${control.id}`);
      }
      continue;
    }

    switch (control.kind) {
      case "boolean": {
        if (typeof answer.value !== "boolean") {
          throw new Error(`Control ${control.id} expected a boolean value`);
        }
        sanitized[control.id] = { value: answer.value };
        break;
      }
      case "segmented": {
        if (!isNonEmptyString(answer.choiceId)) {
          throw new Error(`Control ${control.id} requires a choiceId`);
        }
        const response: { choiceId: string; freeText?: string } = {
          choiceId: answer.choiceId,
        };
        if (answer.freeText && isNonEmptyString(answer.freeText)) {
          response.freeText = answer.freeText.trim();
        }
        sanitized[control.id] = response;
        break;
      }
      case "multi_choice": {
        const ids = Array.isArray(answer.choiceIds)
          ? answer.choiceIds.filter(isNonEmptyString)
          : [];
        if (control.required && ids.length === 0) {
          throw new Error(`Control ${control.id} needs at least one choice`);
        }
        const response: { choiceIds: string[]; freeText?: string } = {
          choiceIds: ids,
        };
        if (answer.freeText && isNonEmptyString(answer.freeText)) {
          response.freeText = answer.freeText.trim();
        }
        sanitized[control.id] = response;
        break;
      }
      case "range_picker": {
        if (!isNonEmptyString(answer.choiceId)) {
          throw new Error(`Control ${control.id} requires a choiceId`);
        }
        const response: { choiceId: string; rangeAddress?: string } = {
          choiceId: answer.choiceId,
        };
        if (answer.rangeAddress && isNonEmptyString(answer.rangeAddress)) {
          response.rangeAddress = answer.rangeAddress.trim();
        }
        sanitized[control.id] = response;
        break;
      }
      case "text": {
        if (!isNonEmptyString(answer.text)) {
          if (control.required) {
            throw new Error(`Control ${control.id} requires text`);
          }
          continue;
        }
        sanitized[control.id] = { text: answer.text.trim() };
        break;
      }
      case "folder_picker": {
        if (!isNonEmptyString(answer.path)) {
          if (control.required) {
            throw new Error(`Control ${control.id} requires a folder path`);
          }
          continue;
        }
        const files = Array.isArray(answer.files)
          ? answer.files.filter(isNonEmptyString)
          : undefined;
        sanitized[control.id] = { path: answer.path.trim(), files };
        break;
      }
      default:
        break;
    }
  }

  return sanitized;
}
