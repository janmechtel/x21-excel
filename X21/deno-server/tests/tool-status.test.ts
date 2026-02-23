import { assertEquals } from "std/assert";
import { formatWriteValuesAction } from "../../web-ui/src/utils/toolStatus.ts";

Deno.test("formatWriteValuesAction returns blocked label for policy failures", () => {
  const label = formatWriteValuesAction({
    isComplete: true,
    isApproved: true,
    isRejected: false,
    isErrored: true,
    errorMessage: "CROSS_WORKBOOK_WRITE_BLOCKED",
  });
  assertEquals(label, "Write values (blocked)");
});

Deno.test("formatWriteValuesAction returns success label only on approved success", () => {
  const label = formatWriteValuesAction({
    isComplete: true,
    isApproved: true,
    isRejected: false,
    isErrored: false,
  });
  assertEquals(label, "Wrote values");
});
