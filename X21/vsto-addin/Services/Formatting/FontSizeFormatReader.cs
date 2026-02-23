using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Reads font size and font name from a unified snapshot (CPU-only).
    /// </summary>
    public class FontSizeFormatReader : IFormatReader
    {
        public string Name => nameof(FontSizeFormatReader);
        public IReadOnlyCollection<string> SupportedProperties => _supportedProps;
        private static readonly string[] _supportedProps = { "fontsize", "fontname" };

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            if (snapshot.FontSize == null && snapshot.FontName == null) return;

            var rows = snapshot.RowCount;
            var cols = snapshot.ColumnCount;

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var address = snapshot.Addresses[r, c];
                    formattedCells.TryGetValue(address, out var settings);
                    var hasValue = false;

                    var size = snapshot.FontSize != null ? FormatReaderHelpers.ToNullableDouble(snapshot.FontSize[r, c]) : null;
                    if (size != null)
                    {
                        settings ??= new FormatSettings();
                        settings.FontSize = (int?)System.Convert.ToInt32(size.Value);
                        hasValue = true;
                    }

                    var name = snapshot.FontName != null ? FormatReaderHelpers.SafeString(snapshot.FontName[r, c]) : null;
                    if (!string.IsNullOrEmpty(name))
                    {
                        settings ??= new FormatSettings();
                        settings.FontName = name;
                        hasValue = true;
                    }

                    if (hasValue)
                    {
                        formattedCells[address] = settings;
                    }
                    else if (!hasValue && settings == null)
                    {
                        formattedCells.Remove(address);
                    }
                }
            }
        }
    }
}
