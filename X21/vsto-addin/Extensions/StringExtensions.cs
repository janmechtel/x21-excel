using System;

namespace X21.Extensions
{
    public static class StringExtensions
    {
        public static string Safe(this string text)
        {
            return text ?? string.Empty;
        }

        public static bool EqualsSafe(this string text, string other)
        {
            return string.Equals(
                text.Safe(),
                other.Safe());
        }
        public static bool EqualsIgnoreCaseSafe(this string text, string other)
        {
            return string.Equals(
                text.Safe(),
                other.Safe(),
                StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
