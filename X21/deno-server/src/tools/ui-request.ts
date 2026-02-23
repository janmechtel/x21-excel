import { Tool, ToolNames } from "../types/index.ts";
import { UiRequestPayload } from "../types/ui-request.ts";

export class UiRequestTool implements Tool<UiRequestPayload> {
  name = ToolNames.COLLECT_INPUT;
  description =
    "Request structured input from the user via an in-chat form. Use this whenever you need confirmations, scoping details, clarifications, or other decision points before proceeding. Prefer concise controls over free text and group related questions into a single request. Offer an 'Other' option with allowFreeText for important decisions. Set text.inputType='number' for numeric responses (amounts, counts, rates). NEVER guess numbers; instead, create a blocking collect_input that asks for missing loan amount, rate, term, and payment frequency,... before continuing.";

  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["title", "mode", "controls"],
    properties: {
      title: {
        type: "string",
        description: "Short headline for the form shown to the user.",
      },
      description: {
        type: "string",
        description:
          "Optional sentence or two explaining why the form is shown and what will happen next.",
      },
      mode: {
        type: "string",
        enum: ["blocking"],
        description:
          'Interaction mode. Use "blocking" to pause execution until the user completes the form.',
      },
      controls: {
        type: "array",
        description:
          "List of controls to render in the form. Combine related questions into one request.",
        minItems: 1,
        items: {
          type: "object",
          required: ["id", "kind", "label"],
          additionalProperties: false,
          properties: {
            id: {
              type: "string",
              description: "Stable identifier for the field",
            },
            kind: {
              type: "string",
              enum: [
                "boolean",
                "segmented",
                "multi_choice",
                "range_picker",
                "text",
              ],
              description: "Control type to render",
            },
            label: {
              type: "string",
              description: "User-facing label/question",
            },
            required: {
              type: "boolean",
              description:
                "Whether this control must be completed before continue",
            },
            inputType: {
              type: "string",
              enum: ["text", "number"],
              description:
                "Only for kind=text. Choose 'number' for numeric inputs (amounts, rates, counts); default is 'text'.",
            },
            placeholder: {
              type: "string",
              description: "Optional placeholder/help text for text inputs.",
            },
            yesLabel: {
              type: "string",
              description: "Optional label for the positive choice (boolean)",
            },
            noLabel: {
              type: "string",
              description: "Optional label for the negative choice (boolean)",
            },
            options: {
              type: "array",
              description:
                "Options for segmented or multi_choice controls. Include an 'Other' with allowFreeText for critical choices.",
              items: {
                type: "object",
                additionalProperties: false,
                required: ["id", "label"],
                properties: {
                  id: { type: "string" },
                  label: { type: "string" },
                  allowFreeText: {
                    type: "boolean",
                    description:
                      "If true, selecting this option opens a free-text field.",
                  },
                },
              },
            },
            presetOptions: {
              type: "array",
              description:
                "Preset locations for range_picker (e.g., used_range, selection).",
              items: {
                type: "object",
                additionalProperties: false,
                required: ["id", "label"],
                properties: {
                  id: { type: "string" },
                  label: { type: "string" },
                },
              },
            },
          },
        },
      },
    },
  };

  execute(_params: UiRequestPayload): Promise<any> {
    // This tool is orchestrated via the chat UI and should never execute directly.
    return Promise.reject(
      new Error(
        "collect_input is handled in the chat UI and should not be executed server-side.",
      ),
    );
  }
}
