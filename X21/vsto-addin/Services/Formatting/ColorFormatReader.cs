using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads font and background colors from a unified snapshot (CPU-only).
    /// </summary>
    public class ColorFormatReader : IFormatReader
    {
        public string Name => nameof(ColorFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "fontcolor", "backgroundcolor" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            ReadFontColors(snapshot, formattedCells);
            ReadBackgroundColors(snapshot, formattedCells);
        }

        private void ReadFontColors(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.FontColor == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var hex = FormatReaderHelpers.ConvertOleColorToHex(snapshot.FontColor[r, c]);
                    if (string.IsNullOrEmpty(hex) || hex == "none") continue;

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.FontColor = hex;
                }
            }
        }

        private void ReadBackgroundColors(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.BackgroundColor == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var colorIndex = snapshot.BackgroundColorIndex != null
                        ? FormatReaderHelpers.ToNullableInt(snapshot.BackgroundColorIndex[r, c])
                        : null;

                    string bgHex;
                    if (colorIndex.HasValue && colorIndex.Value == -4142)
                    {
                        bgHex = "none";
                    }
                    else
                    {
                        bgHex = FormatReaderHelpers.ConvertOleColorToHex(snapshot.BackgroundColor[r, c]);
                        if (string.IsNullOrEmpty(bgHex)) bgHex = "none";
                    }

                    var address = snapshot.Addresses[r, c];
                    if (!formattedCells.TryGetValue(address, out var settings))
                    {
                        settings = new FormatSettings();
                        formattedCells[address] = settings;
                    }
                    settings.BackgroundColor = bgHex;
                }
            }
        }
    }
}
