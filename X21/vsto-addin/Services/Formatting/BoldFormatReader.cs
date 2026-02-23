using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads bold formatting from a unified snapshot (CPU-only).
    /// </summary>
    public class BoldFormatReader : IFormatReader
    {
        public string Name => nameof(BoldFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "bold" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.Bold == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var val = FormatReaderHelpers.ToNullableBool(snapshot.Bold[r, c]);
                    if (val == null) continue;

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.Bold = val;
                }
            }
        }
    }
}
