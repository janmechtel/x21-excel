/** Build a user-facing action label for write_values_batch results. */
export function formatWriteValuesAction(options: {
  isComplete: boolean;
  isApproved: boolean;
  isRejected: boolean;
  isErrored: boolean;
  errorMessage?: string;
}): string {
  const { isComplete, isApproved, isRejected, isErrored, errorMessage } =
    options;

  if (!isComplete || (!isApproved && !isRejected)) {
    return "Writing values";
  }

  if (isErrored) {
    const blocked =
      errorMessage?.includes("CROSS_WORKBOOK_WRITE_BLOCKED") ||
      errorMessage?.toLowerCase().includes("blocked");
    return blocked ? "Write values (blocked)" : "Write values (failed)";
  }

  return isApproved ? "Wrote values" : "Write values (rejected)";
}
