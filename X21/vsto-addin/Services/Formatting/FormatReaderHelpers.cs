using System;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace X21.Services.Formatting
{
    internal static class FormatReaderHelpers
    {
        public static bool? ToNullableBool(object value)
        {
            if (value == null) return null;
            try
            {
                if (value is bool b) return b;
                if (value is int i) return i != 0;
                if (bool.TryParse(value.ToString(), out var parsed)) return parsed;
            }
            catch { /* ignore */ }
            return null;
        }

        public static int? ToNullableInt(object value)
        {
            if (value == null) return null;
            try
            {
                if (value is int i) return i;
                if (value is double d) return Convert.ToInt32(d);
                if (int.TryParse(value.ToString(), out var parsed)) return parsed;
            }
            catch { /* ignore */ }
            return null;
        }

        public static double? ToNullableDouble(object value)
        {
            if (value == null) return null;
            try
            {
                if (value is double d) return d;
                if (value is float f) return f;
                if (double.TryParse(value.ToString(), out var parsed)) return parsed;
            }
            catch { /* ignore */ }
            return null;
        }

        public static string SafeString(object value)
        {
            return value?.ToString();
        }

        public static string ConvertOleColorToHex(object oleColor)
        {
            try
            {
                if (oleColor == null) return "none";
                var colorValue = Convert.ToInt32(oleColor);
                if (colorValue == -4142) return "none"; // xlColorIndexNone
                var color = ColorTranslator.FromOle(colorValue);
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }
            catch
            {
                return "none";
            }
        }

        public static string ConvertAlignment(object alignment)
        {
            if (alignment == null) return "left";
            try
            {
                // Excel may return XlHAlign, int, or double; normalize to int code first.
                int code;
                if (alignment is XlHAlign hAlign)
                {
                    code = (int)hAlign;
                }
                else
                {
                    code = Convert.ToInt32(alignment);
                }

                return ((XlHAlign)code) switch
                {
                    XlHAlign.xlHAlignCenter => "center",
                    XlHAlign.xlHAlignRight => "right",
                    XlHAlign.xlHAlignJustify => "justify",
                    _ => "left"
                };
            }
            catch
            {
                return "left";
            }
        }
    }
}
