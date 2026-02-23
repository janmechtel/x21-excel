import { assertEquals, assertThrows } from "std/assert";
import { chunkArray, EXCEL_BATCH_MAX_OPS } from "../src/utils/batching.ts";
import { ExcelApiClient } from "../src/utils/excel-api-client.ts";
import { readValuesBatch } from "../src/excel-actions/read-values-batch.ts";
import { writeValuesBatch } from "../src/excel-actions/write-values-batch.ts";
import { ToolNames } from "../src/types/index.ts";

Deno.test("chunkArray handles empty and boundary sizes", () => {
  assertEquals(chunkArray([]), []);
  assertEquals(chunkArray([1]).length, 1);
  assertEquals(chunkArray(Array.from({ length: 10 }, (_, i) => i)).length, 1);
  assertEquals(
    chunkArray(Array.from({ length: 11 }, (_, i) => i)).length,
    2,
  );
  assertEquals(
    chunkArray(Array.from({ length: 25 }, (_, i) => i)).length,
    3,
  );
});

Deno.test("chunkArray rejects oversize chunk size", () => {
  assertThrows(() => chunkArray([1, 2], EXCEL_BATCH_MAX_OPS + 1));
});

Deno.test("readValuesBatch chunks requests and preserves order", async () => {
  const operations = Array.from({ length: 25 }, (_, i) => ({
    workbookName: "Book1.xlsx",
    worksheet: "Sheet1",
    range: `A${i + 1}`,
  }));

  const calls: Array<{ action: string; count: number }> = [];
  const originalGetInstance = ExcelApiClient.getInstance;
  // deno-lint-ignore no-explicit-any
  (ExcelApiClient as any).getInstance = () =>
    ({
      executeExcelAction: (_action: string, params: any) => {
        calls.push({ action: _action, count: params.operations.length });
        return {
          success: true,
          message: "ok",
          results: params.operations.map((op: any) => ({
            success: true,
            message: "ok",
            worksheet: op.worksheet,
            workbookName: op.workbookName,
            range: op.range,
            cellValues: {
              [op.range]: { value: op.range, formula: "" },
            },
          })),
        };
      },
    }) as ExcelApiClient;

  try {
    const result = await readValuesBatch({ operations });
    assertEquals(calls.length, 3);
    calls.forEach((call) => {
      assertEquals(call.action, ToolNames.READ_VALUES_BATCH);
      assertEquals(call.count <= EXCEL_BATCH_MAX_OPS, true);
    });

    assertEquals(result.results.length, operations.length);
    const ranges = result.results.map((res) => res.range);
    assertEquals(ranges, operations.map((op) => op.range));
  } finally {
    // deno-lint-ignore no-explicit-any
    (ExcelApiClient as any).getInstance = originalGetInstance;
  }
});

Deno.test("writeValuesBatch chunks requests and aggregates results", async () => {
  const operations = Array.from({ length: 11 }, (_, i) => ({
    workbookName: "Book1.xlsx",
    worksheet: "Sheet1",
    range: `B${i + 1}`,
    values: [["x"]],
  }));

  const calls: Array<{ action: string; count: number }> = [];
  const originalGetInstance = ExcelApiClient.getInstance;
  // deno-lint-ignore no-explicit-any
  (ExcelApiClient as any).getInstance = () =>
    ({
      executeExcelAction: (_action: string, params: any) => {
        calls.push({ action: _action, count: params.operations.length });
        return {
          success: true,
          message: "ok",
          results: params.operations.map(() => ({
            success: true,
            message: "ok",
          })),
        };
      },
    }) as ExcelApiClient;

  try {
    const result = await writeValuesBatch({ operations });
    assertEquals(calls.length, 2);
    calls.forEach((call) => {
      assertEquals(call.action, ToolNames.WRITE_VALUES_BATCH);
      assertEquals(call.count <= EXCEL_BATCH_MAX_OPS, true);
    });

    assertEquals(result.operations.length, operations.length);
    assertEquals(result.batches, 2);
    assertEquals(result.applied, operations.length);
    assertEquals(result.policyAction, "wrote");
  } finally {
    // deno-lint-ignore no-explicit-any
    (ExcelApiClient as any).getInstance = originalGetInstance;
  }
});
