using System;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Lightweight progress payload used while reading Excel formatting.
    /// </summary>
    public class FormatProgressUpdate
    {
        public int Current { get; set; }
        public int Total { get; set; }
        public string Message { get; set; } = string.Empty;

        public FormatProgressUpdate(int current, int total, string message)
        {
            Current = Math.Max(0, current);
            Total = Math.Max(1, total);
            Message = message ?? string.Empty;
        }
    }
}
