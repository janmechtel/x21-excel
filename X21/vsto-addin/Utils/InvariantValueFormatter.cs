using System;
using System.Globalization;

namespace X21.Utils
{
    public static class InvariantValueFormatter
    {
        /// <summary>
        /// Convert common Excel Value2/object values to a stable, invariant string.
        /// This avoids culture-specific decimal separators (e.g. "1,23").
        /// </summary>
        public static string ToInvariantString(object value)
        {
            if (value == null) return string.Empty;

            if (value is string s) return s;

            // Prefer Excel-friendly boolean literals.
            if (value is bool b) return b ? "TRUE" : "FALSE";

            // Numeric types: force '.' decimal separator, no grouping.
            if (value is IConvertible c)
            {
                switch (c.GetTypeCode())
                {
                    case TypeCode.Byte:
                    case TypeCode.SByte:
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        return c.ToString(CultureInfo.InvariantCulture);
                }
            }

            // Everything else: preserve existing string form.
            return value.ToString();
        }
    }
}
