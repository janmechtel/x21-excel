export function normalizeCurrencyNumberFormat(
  format?: string,
): string | undefined {
  if (!format) return format;

  let inQuotes = false;
  let inBracket = false;
  let normalized = "";

  for (let i = 0; i < format.length; i++) {
    const char = format[i];
    const prevChar = i > 0 ? format[i - 1] : "";
    const isEscaped = prevChar === "\\";

    if (char === '"' && !isEscaped) {
      inQuotes = !inQuotes;
      normalized += char;
      continue;
    }

    if (!inQuotes) {
      if (char === "[" && !isEscaped) {
        inBracket = true;
        normalized += char;
        continue;
      }
      if (char === "]" && inBracket && !isEscaped) {
        inBracket = false;
        normalized += char;
        continue;
      }
    }

    if (char === "$" && !inQuotes && !inBracket && !isEscaped) {
      normalized += "\\$";
      continue;
    }

    normalized += char;
  }

  return normalized;
}
