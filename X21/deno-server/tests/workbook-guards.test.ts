import { assertEquals } from "std/assert";
import { applyWorkbookGuards } from "../src/stream/tool-logic.ts";
import { ToolNames } from "../src/types/index.ts";

Deno.test("applyWorkbookGuards ignores provided workbook for writes and uses active", () => {
  const input = {
    workbookName: "Other.xlsx",
    operations: [
      {
        worksheet: "Sheet1",
        range: "A1",
        values: [["x"]],
      },
    ],
  };

  const result = applyWorkbookGuards(
    ToolNames.WRITE_VALUES_BATCH,
    input,
    "Active.xlsx",
  );
  assertEquals(result.workbookName, "Active.xlsx");
  assertEquals(result.operations[0].workbookName, "Active.xlsx");
});

Deno.test("applyWorkbookGuards injects active workbook for writes", () => {
  const input = {
    operations: [
      {
        worksheet: "Sheet1",
        range: "A1",
        values: [["x"]],
      },
    ],
  };
  const result = applyWorkbookGuards(
    ToolNames.WRITE_VALUES_BATCH,
    input,
    "Active.xlsx",
  );
  assertEquals(result.workbookName, "Active.xlsx");
  assertEquals(result.operations[0].workbookName, "Active.xlsx");
});

Deno.test("applyWorkbookGuards keeps provided workbook on reads", () => {
  const input = {
    operations: [
      {
        worksheet: "Sheet1",
        range: "A1",
        workbookName: "Other.xlsx",
      },
    ],
  };
  const result = applyWorkbookGuards(
    ToolNames.READ_VALUES_BATCH,
    input,
    "Active.xlsx",
  );
  assertEquals(result.operations[0].workbookName, "Other.xlsx");
});

Deno.test("applyWorkbookGuards injects active workbook on reads when missing", () => {
  const input = {
    operations: [
      {
        worksheet: "Sheet1",
        range: "A1",
      },
    ],
  };
  const result = applyWorkbookGuards(
    ToolNames.READ_VALUES_BATCH,
    input,
    "Active.xlsx",
  );
  assertEquals(result.operations[0].workbookName, "Active.xlsx");
});
