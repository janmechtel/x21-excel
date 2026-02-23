import { assertEquals } from "std/assert";
import { getInputDataRevert } from "../src/stream/tool-logic.ts";
import type { ToolChangeInterface } from "../src/state/state-manager.ts";

Deno.test("getInputDataRevert handles sheet-prefixed ranges for write_values_batch", () => {
  const toolChange: ToolChangeInterface = {
    timestamp: new Date(),
    workbookName: "tmp8F8E.xlsx",
    worksheet: "Source Actuals",
    requestId: "req-1",
    toolId: "tool-1",
    toolName: "write_values_batch",
    applied: true,
    approved: true,
    pending: false,
    inputData: {
      operations: [
        {
          worksheet: "Source Actuals",
          range: "Source Actuals!J9:L13",
          values: [["1"]],
        },
      ],
    },
  };

  const result = {
    oldValues: [
      {
        cellValues: {
          "Source Actuals!J9": { value: "1" },
          "Source Actuals!K9": { value: "2" },
          "Source Actuals!L9": { value: "3" },
          "Source Actuals!J10": { value: "4" },
        },
      },
    ],
  };

  const revertPayload = getInputDataRevert(
    result,
    "tmp8F8E.xlsx",
    toolChange,
  );

  const writeValues = revertPayload["write_values_batch"][0];
  assertEquals(writeValues.range, "J9:L13");
  assertEquals(writeValues.values.length, 5);
  writeValues.values.forEach((row: string[]) => assertEquals(row.length, 3));
});
