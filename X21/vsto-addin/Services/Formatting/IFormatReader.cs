using System.Collections.Generic;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Snapshot-based format reader. Implementations MUST be CPU-only and never touch COM.
    /// </summary>
    public interface IFormatReader
    {
        /// <summary>
        /// Human-friendly name (typically the class name)
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Properties that this reader populates (lowercase, e.g., "bold", "fontcolor").
        /// Used to filter readers when a selective property list is supplied.
        /// </summary>
        IReadOnlyCollection<string> SupportedProperties { get; }

        /// <summary>
        /// Compute format settings from a unified snapshot and merge into the provided dictionary.
        /// Implementations MUST NOT perform COM calls; all data must come from <paramref name="snapshot" />.
        /// </summary>
        void ComputeFormats(FormatSnapshot snapshot, Dictionary<string, FormatSettings> formattedCells);
    }
}
