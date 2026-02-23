import { dragFormula } from "../../excel-actions/drag-formula.ts";
import {
  getModifiedRangeA1,
  isValidAutoFillOverlap,
} from "../../utils/drag-utils.ts";
import { readValues } from "../../excel-actions/read-values.ts";
import {
  DragToolRequest,
  ReadValuesRequest,
  Tool,
  ToolNames,
} from "../../types/index.ts";

export class DragFormulaTool implements Tool {
  name = ToolNames.DRAG_FORMULA;
  description =
    `Drag a formula or pattern from a source range to a destination range using Excel's AutoFill functionality.

Sequence guidance (dates/numbers):
- Use the minimal source range needed to establish the pattern.
- Use a single cell for simple increments.
- Use two cells only when you must define a step size.
- Avoid larger source ranges to prevent pattern repetition.

Range rules:
- destinationRange must be a valid range in the workbook.
- sourceRange must be contained inside destinationRange.
`;
  input_schema = {
    type: "object",
    additionalProperties: false,
    required: ["worksheet", "sourceRange", "destinationRange"],
    properties: {
      worksheet: { type: "string", description: "The worksheet name" },
      sourceRange: {
        type: "string",
        description:
          "The source range containing the formula/pattern to drag. For sequences: use single cell (A1) for simple +1 increments, or two cells (A1:A2) to establish step size. Avoid larger ranges to prevent pattern repetition.",
      },
      destinationRange: {
        type: "string",
        description:
          "The destination range to fill (e.g., A1:O3), the source should be contained in the destination. e.g. if sourceRange is A1:A1 then destinationRange should be A1:A10, or if sourceRange is A1:A10 then destinationRange should be A1:D10 ",
      },
      fillType: {
        type: "string",
        description:
          "AutoFill type: 'default', 'series', 'formats', 'values'. Default is 'default'",
        enum: ["default", "series", "formats", "values"],
      },
    },
  };

  async execute(params: DragToolRequest): Promise<any> {
    // Validate that the source and destination ranges have a valid AutoFill overlap
    if (!isValidAutoFillOverlap(params.sourceRange, params.destinationRange)) {
      throw new Error(
        `Invalid AutoFill operation: Source range '${params.sourceRange}' must be contained within destination range '${params.destinationRange}' in a valid AutoFill pattern. ` +
          `For example: source 'A1:A2' with destination 'A1:A10' (fill down), or source 'A1:B1' with destination 'A1:E1' (fill right).`,
      );
    }

    const modifiedRange = getModifiedRangeA1(
      params.sourceRange,
      params.destinationRange,
    );

    const readValuesParams: ReadValuesRequest = {
      worksheet: params.worksheet,
      range: params.destinationRange,
      workbookName: params.workbookName,
    };

    const oldValues = await readValues(readValuesParams);

    const result = await dragFormula(params);

    const newValues = await readValues(readValuesParams);

    return {
      oldValues,
      newValues,
      range: modifiedRange,
      result,
    };
  }
}
