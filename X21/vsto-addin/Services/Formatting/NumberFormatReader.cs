using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads number formatting from a unified snapshot (CPU-only).
    /// </summary>
    public class NumberFormatReader : IFormatReader
    {
        public string Name => nameof(NumberFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "numberformat" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.NumberFormat == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var numberFormat = FormatReaderHelpers.SafeString(snapshot.NumberFormat[r, c]);
                    if (string.IsNullOrEmpty(numberFormat)) continue;

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.NumberFormat = numberFormat;
                }
            }
        }
    }
}
