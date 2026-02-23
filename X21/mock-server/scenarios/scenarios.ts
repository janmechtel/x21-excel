/**
 * Test Scenarios for Mock Backend
 *
 * Defines various response scenarios for testing different UI states
 */

import { ContentBlockTypes, ClaudeContentTypes, ToolNames, type ContentBlockType } from "../../shared/types/index.ts";

interface StreamScenario {
  blocks: Array<{
    type: ContentBlockType;
    content?: string;
    toolName?: string;
    toolInput?: Record<string, unknown>;
  }>;
  delayMs?: number;
  requestToolApproval?: boolean;
  toolPermissions?: Array<{
    toolId: string;
    toolName: string;
    toolInput: Record<string, unknown>;
  }>;
  uiRequest?: {
    toolUseId: string;
    payload: Record<string, unknown>;
  };
}

/**
 * Simple text response without tools
 */
export function getSimpleResponse(prompt: string): StreamScenario {
  return {
    blocks: [
      {
        type: ContentBlockTypes.TEXT,
        content: `I understand you asked: "${prompt}". This is a mock response from the testing server. In a real scenario, Claude would provide a helpful and detailed response to your query.`,
      },
    ],
    delayMs: 50,
  };
}

/**
 * Response with thinking block
 */
export function getResponseWithThinking(): StreamScenario {
  return {
    blocks: [
      {
        type: ContentBlockTypes.THINKING,
        content: "Let me analyze this request carefully. I need to consider the user's requirements and determine the best approach to help them.",
      },
      {
        type: ContentBlockTypes.TEXT,
        content: "Based on my analysis, I recommend the following approach. This will help you achieve your goals effectively.",
      },
    ],
    delayMs: 50,
  };
}

/**
 * Response that requests tool approval
 */
export function getToolApprovalScenario(prompt: string): StreamScenario {
  const toolId1 = `toolu_${Date.now()}_1`;
  const toolId2 = `toolu_${Date.now()}_2`;

  return {
    blocks: [
      {
        type: ContentBlockTypes.THINKING,
        content: "I need to make some changes to the Excel workbook. Let me prepare the necessary operations.",
      },
      {
        type: ContentBlockTypes.TEXT,
        content: "I'll help you with that. I need to perform a couple of operations in your workbook:",
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.WRITE_VALUES_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A1:B5",
              values: [
                ["Name", "Value"],
                ["Item 1", 100],
                ["Item 2", 200],
                ["Item 3", 300],
                ["Total", 600],
              ],
            },
          ],
        },
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.WRITE_FORMAT_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A1:B1",
              format: {
                bold: true,
                fontSize: 12,
                backgroundColor: "#4472C4",
                fontColor: "#FFFFFF",
              },
            },
          ],
        },
      },
    ],
    delayMs: 50,
    requestToolApproval: true,
    toolPermissions: [
      {
        toolId: toolId1,
        toolName: ToolNames.WRITE_VALUES_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A1:B5",
              values: [
                ["Name", "Value"],
                ["Item 1", 100],
                ["Item 2", 200],
                ["Item 3", 300],
                ["Total", 600],
              ],
            },
          ],
        },
      },
      {
        toolId: toolId2,
        toolName: ToolNames.WRITE_FORMAT_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A1:B1",
              format: {
                bold: true,
                fontSize: 12,
                backgroundColor: "#4472C4",
                fontColor: "#FFFFFF",
              },
            },
          ],
        },
      },
    ],
  };
}

/**
 * Response with single tool use
 */
export function getSingleToolScenario(): StreamScenario {
  const toolId = `toolu_${Date.now()}_1`;

  return {
    blocks: [
      {
        type: ContentBlockTypes.TEXT,
        content: "I'll read the current values from the selected range.",
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.READ_VALUES_BATCH,
        toolInput: {
          operations: [
            { worksheet: "Sheet1", range: "A1:C10" },
          ],
        },
      },
    ],
    delayMs: 50,
    requestToolApproval: true,
    toolPermissions: [
      {
        toolId,
        toolName: ToolNames.READ_VALUES_BATCH,
        toolInput: {
          operations: [
            { worksheet: "Sheet1", range: "A1:C10" },
          ],
        },
      },
    ],
  };
}

/**
 * Response with multiple tool uses (batch approval)
 */
export function getBatchToolScenario(): StreamScenario {
  const toolIds = Array.from({ length: 3 }, (_, i) => `toolu_${Date.now()}_${i + 1}`);

  return {
    blocks: [
      {
        type: ContentBlockTypes.THINKING,
        content: "I need to perform several operations to complete this task. I'll add some rows, write values, and apply formatting.",
      },
      {
        type: ContentBlockTypes.TEXT,
        content: "I'll perform the following operations:",
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.ADD_ROWS,
        toolInput: {
          row_index: 5,
          count: 3,
        },
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.WRITE_VALUES_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A5:A7",
              values: [["New Row 1"], ["New Row 2"], ["New Row 3"]],
            },
          ],
        },
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.WRITE_FORMAT_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A5:A7",
              format: {
                italic: true,
              },
            },
          ],
        },
      },
    ],
    delayMs: 50,
    requestToolApproval: true,
    toolPermissions: [
      {
        toolId: toolIds[0],
        toolName: ToolNames.ADD_ROWS,
        toolInput: {
          row_index: 5,
          count: 3,
        },
      },
      {
        toolId: toolIds[1],
        toolName: ToolNames.WRITE_VALUES_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A5:A7",
              values: [["New Row 1"], ["New Row 2"], ["New Row 3"]],
            },
          ],
        },
      },
      {
        toolId: toolIds[2],
        toolName: ToolNames.WRITE_FORMAT_BATCH,
        toolInput: {
          operations: [
            {
              worksheet: "Sheet1",
              range: "A5:A7",
              format: {
                italic: true,
              },
            },
          ],
        },
      },
    ],
  };
}

/**
 * UI form request scenario
 */
export function getUiRequestScenario(): StreamScenario {
  const toolUseId = `toolu_${Date.now()}_ui`;
  return {
    blocks: [
      {
        type: ContentBlockTypes.THINKING,
        content: "I need a few clarifications before drafting the amortization plan.",
      },
      {
        type: ContentBlockTypes.TEXT,
        content: "Please fill in a few quick details so I can tailor the schedule correctly.",
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.COLLECT_INPUT,
        toolInput: {
          title: "Amortization Schedule Parameters",
          description: "I need a few more details to create your $100,000 amortization schedule at 5% interest.",
          mode: "blocking",
          controls: [
            {
              id: "schedule_type",
              kind: "segmented",
              label: "What type of amortization schedule do you need?",
              required: true,
              options: [
                {
                  id: "loan",
                  label: "Loan Payment Schedule (paying down debt)",
                },
                {
                  id: "investment",
                  label: "Investment Growth Schedule (compound interest)",
                },
                {
                  id: "other",
                  label: "Other",
                  allowFreeText: true,
                },
              ],
            },
            {
              id: "term_years",
              kind: "text",
              label: "How many years? (e.g., 30, 15, 10)",
              required: true,
            },
            {
              id: "payment_frequency",
              kind: "segmented",
              label: "Payment/Compounding frequency",
              required: true,
              options: [
                { id: "monthly", label: "Monthly" },
                { id: "quarterly", label: "Quarterly" },
                { id: "annual", label: "Annual" },
              ],
            },
          ],
        },
      },
    ],
    delayMs: 40,
    uiRequest: {
      toolUseId,
      payload: {
        mode: "blocking",
        title: "Amortization Schedule Parameters",
        description: "I need a few more details to create your $100,000 amortization schedule at 5% interest.",
        controls: [
          {
            id: "schedule_type",
            kind: "segmented",
            label: "What type of amortization schedule do you need?",
            required: true,
            options: [
              { id: "loan", label: "Loan Payment Schedule (paying down debt)" },
              { id: "investment", label: "Investment Growth Schedule (compound interest)" },
              { id: "other", label: "Other", allowFreeText: true },
            ],
          },
          {
            id: "term_years",
            kind: "text",
            label: "How many years? (e.g., 30, 15, 10)",
            required: true,
            inputType: "number",
          },
          {
            id: "payment_frequency",
            kind: "segmented",
            label: "Payment/Compounding frequency",
            required: true,
            options: [
              { id: "monthly", label: "Monthly" },
              { id: "quarterly", label: "Quarterly" },
              { id: "annual", label: "Annual" },
            ],
          },
        ],
      },
    },
  };
}

/**
 * UI form request that showcases all control types and responses
 */
export function getUiControlsShowcaseScenario(): StreamScenario {
  const toolUseId = `toolu_${Date.now()}_ui_full`;
  return {
    blocks: [
      {
        type: ContentBlockTypes.TEXT,
        content: "Let me capture a few structured inputs to continue.",
      },
      {
        type: ClaudeContentTypes.TOOL_USE,
        toolName: ToolNames.COLLECT_INPUT,
        toolInput: {
          title: "Form Showcase",
          description: "Demo of all control types.",
          mode: "blocking",
          controls: [],
        },
      },
    ],
    delayMs: 30,
    uiRequest: {
      toolUseId,
      payload: {
        mode: "blocking",
        title: "Form Showcase",
        description: "Demo of all control types.",
        controls: [
          {
            id: "confirm_boolean",
            kind: "boolean",
            label: "Proceed with the default plan?",
            required: true,
            yesLabel: "Yes, continue",
            noLabel: "No, pause (but this is a really long label to test UI handling of long text in button labels)",
          },
          {
            id: "segmented_choice",
            kind: "segmented",
            label: "Pick one option",
            required: true,
            options: [
              { id: "alpha", label: "Alpha" },
              { id: "beta", label: "Beta" },
              { id: "custom", label: "Other", allowFreeText: true },
            ],
          },
          {
            id: "multi_choice",
            kind: "multi_choice",
            label: "Select all that apply",
            required: false,
            options: [
              { id: "numbers", label: "Numbers" },
              { id: "formats", label: "Formats" },
              { id: "notes", label: "Notes", allowFreeText: true },
            ],
          },
          {
            id: "range_picker",
            kind: "range_picker",
            label: "Where should I write results?",
            required: true,
            presetOptions: [
              { id: "used_range", label: "Used range on this sheet" },
              { id: "selection", label: "Current selection" },
            ],
          },
          {
            id: "text_notes",
            kind: "text",
            label: "Additional notes",
            required: false,
          },
          {
            id: "numeric_input",
            kind: "text",
            label: "Numeric input demo",
            required: false,
            inputType: "number",
          },
        ],
      },
    },
  };
}

/**
 * Long response for testing scrolling
 */
export function getLongResponse(): StreamScenario {
  return {
    blocks: [
      {
        type: ContentBlockTypes.TEXT,
        content: "Here's a comprehensive explanation:\n\n" +
          "1. First, let's understand the basics. Excel is a powerful spreadsheet application that allows you to organize, analyze, and visualize data.\n\n" +
          "2. When working with large datasets, it's important to use efficient formulas and techniques.\n\n" +
          "3. Some key concepts include: cell references, formulas, functions, pivot tables, and charts.\n\n" +
          "4. For data analysis, you can use functions like SUM, AVERAGE, COUNT, IF, VLOOKUP, and many more.\n\n" +
          "5. Formatting is crucial for making your data readable and professional-looking.\n\n" +
          "6. You can apply various formatting options like font styles, colors, borders, and number formats.\n\n" +
          "7. Advanced features include macros, VBA programming, and data validation.\n\n" +
          "8. Remember to always save your work regularly and use version control when collaborating.",
      },
    ],
    delayMs: 30,
  };
}

/**
 * Error scenario (not used in streaming, but for reference)
 */
export function getErrorResponse(): { error: string; message: string } {
  return {
    error: "MOCK_ERROR",
    message: "This is a simulated error for testing error handling in the UI.",
  };
}
