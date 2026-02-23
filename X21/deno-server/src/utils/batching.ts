export const EXCEL_BATCH_MAX_OPS = 10;

export function chunkArray<T>(
  items: T[],
  chunkSize: number = EXCEL_BATCH_MAX_OPS,
): T[][] {
  if (!Array.isArray(items)) {
    throw new Error("items must be an array");
  }
  if (!Number.isFinite(chunkSize) || chunkSize <= 0) {
    throw new Error("chunkSize must be a positive number");
  }
  if (chunkSize > EXCEL_BATCH_MAX_OPS) {
    throw new Error(
      `chunkSize (${chunkSize}) exceeds max (${EXCEL_BATCH_MAX_OPS})`,
    );
  }

  if (items.length === 0) return [];

  const chunks: T[][] = [];
  for (let i = 0; i < items.length; i += chunkSize) {
    chunks.push(items.slice(i, i + chunkSize));
  }
  return chunks;
}
