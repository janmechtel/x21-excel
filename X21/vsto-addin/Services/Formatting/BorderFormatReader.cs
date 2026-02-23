using System.Collections.Generic;
using X21.Logging;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Border reading is intentionally disabled in the unified snapshot pipeline.
    /// This stub satisfies the interface without performing COM work.
    /// </summary>
    public class BorderFormatReader : IFormatReader
    {
        private static readonly string[] _supportedProperties = { "border", "borders" };

        public string Name => nameof(BorderFormatReader);

        public IReadOnlyCollection<string> SupportedProperties => _supportedProperties;

        public void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells)
        {
            // Borders are currently excluded from formatting reads to avoid heavy COM scans.
            // Leave as a no-op; if border support is reintroduced, implement computation from the unified snapshot.
            Logger.Info("BorderFormatReader: skipping border computation in unified snapshot pipeline");
        }
    }
}
