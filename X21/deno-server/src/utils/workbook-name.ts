/**
 * Normalize workbook names coming from LLM/tool args.
 * Strips common formatting wrappers while preserving inner characters.
 */
export function sanitizeWorkbookName(input: unknown): string | undefined {
  if (typeof input !== "string") return undefined;

  const value = input.trim();
  if (!value) return undefined;

  const unwrap = (text: string): string => {
    const pairs: Array<[string, string]> = [
      ["`", "`"],
      ['"', '"'],
      ["'", "'"],
    ];
    for (const [start, end] of pairs) {
      if (text.startsWith(start) && text.endsWith(end) && text.length > 1) {
        return text.slice(1, -1).trim();
      }
    }
    return text;
  };

  let previous: string;
  let sanitized = value;
  // Keep unwrapping while we make progress (handles ```Book``` etc.)
  do {
    previous = sanitized;
    sanitized = unwrap(sanitized);
  } while (sanitized !== previous);

  sanitized = sanitized.trim();
  return sanitized.length > 0 ? sanitized : undefined;
}
