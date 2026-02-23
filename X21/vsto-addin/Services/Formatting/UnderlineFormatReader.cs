using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads underline formatting from a unified snapshot (CPU-only).
    /// </summary>
    public class UnderlineFormatReader : IFormatReader
    {
        public string Name => nameof(UnderlineFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "underline" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.Underline == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var raw = snapshot.Underline[r, c];
                    if (raw == null) continue;

                    bool isUnderlined;
                    if (raw is int intValue)
                    {
                        isUnderlined = intValue != (int)XlUnderlineStyle.xlUnderlineStyleNone;
                    }
                    else
                    {
                        isUnderlined = !raw.Equals(XlUnderlineStyle.xlUnderlineStyleNone);
                    }

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.Underline = isUnderlined;
                }
            }
        }
    }
}
