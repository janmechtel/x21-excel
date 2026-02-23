import { assertEquals } from "std/assert";
import { WriteValuesBatchTool } from "../src/tools/excel/write-values-batch.ts";
import { sanitizeWorkbookName } from "../src/utils/workbook-name.ts";

Deno.test("sanitizeWorkbookName strips wrappers and trims whitespace", () => {
  assertEquals(
    sanitizeWorkbookName("`20260114_Deeploi_Buchungsstapel.xlsx`"),
    "20260114_Deeploi_Buchungsstapel.xlsx",
  );
  assertEquals(sanitizeWorkbookName("  'Book2'  "), "Book2");
  assertEquals(sanitizeWorkbookName('"Book3.xlsx"'), "Book3.xlsx");
  assertEquals(sanitizeWorkbookName("  "), undefined);
  assertEquals(sanitizeWorkbookName(123), undefined);
});

Deno.test("write_values_batch blocks cross-workbook requests and returns sanitized names", async () => {
  const tool = new WriteValuesBatchTool();
  const result = await tool.execute({
    activeWorkbookName: "Book1.xlsx",
    operations: [
      {
        workbookName: "`Book2.xlsx`",
        worksheet: "Sheet1",
        range: "A1",
        values: [["x"]],
      },
    ],
  });

  assertEquals(result.success, false);
  assertEquals(result.policyAction, "blocked");
  assertEquals(result.requestedWorkbook, "Book2.xlsx");
  assertEquals(result.resolvedWorkbook, null);
  assertEquals(result.activeWorkbook, "Book1.xlsx");

  const op = result.operations[0];
  assertEquals(op.policyAction, "blocked");
  assertEquals(op.requestedWorkbook, "Book2.xlsx");
  assertEquals(op.resolvedWorkbook, null);
  assertEquals(op.errorCode, "CROSS_WORKBOOK_WRITE_BLOCKED");
  assertEquals(op.status, "error");
  assertEquals(op.oldValues, null);
  assertEquals(op.newValues, null);
  assertEquals(
    op.errorMessage?.includes("`"),
    false,
    "errorMessage should not include backticks",
  );
});

Deno.test("sanitizeWorkbookName unwraps nested quotes and backticks", () => {
  assertEquals(sanitizeWorkbookName("''Book.xlsx''"), "Book.xlsx");
  assertEquals(sanitizeWorkbookName("``Book.xlsx``"), "Book.xlsx");
  assertEquals(sanitizeWorkbookName("'`Book.xlsx`'"), "Book.xlsx");
  assertEquals(sanitizeWorkbookName("  '  Book.xlsx  '  "), "Book.xlsx");
});
