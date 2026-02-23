using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads italic formatting from a unified snapshot (CPU-only).
    /// </summary>
    public class ItalicFormatReader : IFormatReader
    {
        public string Name => nameof(ItalicFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "italic" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.Italic == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var val = FormatReaderHelpers.ToNullableBool(snapshot.Italic[r, c]);
                    if (val == null) continue;

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.Italic = val;
                }
            }
        }
    }
}
