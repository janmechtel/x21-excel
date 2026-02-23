import { assertEquals, assertRejects } from "std/assert";
import {
  resolveWorkbookForWrite,
  WorkbookResolutionError,
} from "../src/utils/workbook-resolver.ts";
import { WorkbookResolutionPaths } from "../src/types/index.ts";

Deno.test("resolveWorkbookForWrite prefers session workbook over host active", async () => {
  const result = await resolveWorkbookForWrite(
    { sessionWorkbookName: "Session.xlsx" },
    {
      getHostMetadata: () => ({ workbookName: "Host.xlsx" } as any),
      listOpenWorkbooks: () => Promise.resolve([]),
    },
  );

  assertEquals(result.workbookName, "Session.xlsx");
  assertEquals(result.resolutionPath, WorkbookResolutionPaths.SESSION);
});

Deno.test("resolveWorkbookForWrite adopts provided workbook if open", async () => {
  const result = await resolveWorkbookForWrite(
    {
      sessionWorkbookName: "Book1.xlsx",
      providedWorkbookName: "Provided.xlsx",
    },
    {
      getHostMetadata: () => {
        throw new Error("fail");
      },
      listOpenWorkbooks: () =>
        Promise.resolve([{ workbookName: "Provided.xlsx" }] as any),
    },
  );

  assertEquals(result.workbookName, "Provided.xlsx");
  assertEquals(result.resolutionPath, WorkbookResolutionPaths.PROVIDED_OPEN);
});

Deno.test("resolveWorkbookForWrite uses single open workbook when session unknown", async () => {
  const result = await resolveWorkbookForWrite(
    { sessionWorkbookName: "" },
    {
      getHostMetadata: () => {
        throw new Error("fail");
      },
      listOpenWorkbooks: () =>
        Promise.resolve([{ workbookName: "Only.xlsx" }] as any),
    },
  );

  assertEquals(result.workbookName, "Only.xlsx");
  assertEquals(result.resolutionPath, WorkbookResolutionPaths.SINGLE_OPEN);
});

Deno.test("resolveWorkbookForWrite throws when nothing can be resolved", async () => {
  await assertRejects(
    () =>
      resolveWorkbookForWrite(
        { sessionWorkbookName: "" },
        {
          getHostMetadata: () => {
            throw new Error("fail");
          },
          listOpenWorkbooks: () =>
            Promise.resolve(
              [{ workbookName: "One.xlsx" }, {
                workbookName: "Two.xlsx",
              }] as any,
            ),
        },
      ),
    WorkbookResolutionError,
    "No active workbook is set",
  );
});
