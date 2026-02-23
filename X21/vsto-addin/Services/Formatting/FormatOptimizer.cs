using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using X21.Logging;
using X21.Models;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Service for optimizing format operations through intelligent inference and caching
    /// </summary>
    public class FormatOptimizer
    {
        private readonly FormatManager _formatManager;

        public FormatOptimizer(FormatManager formatManager)
        {
            _formatManager = formatManager;
        }

        /// <summary>
        /// Infers new format values by merging old formats with applied changes
        /// This avoids the expensive second read operation
        /// </summary>
        public Dictionary<string, FormatSettings> InferNewFormatsFromAppliedChanges(
            Dictionary<string, FormatSettings> oldFormats,
            FormatSettings appliedFormat,
            Range targetRange)
        {
            if (oldFormats == null)
                return new Dictionary<string, FormatSettings>();

            var newFormats = new Dictionary<string, FormatSettings>();

            try
            {
                Logger.Info($"🧠 Inferring new formats by merging old formats with applied changes");

                // Get all cell addresses in the target range
                var allCellAddresses = GetAllCellAddressesInRange(targetRange);

                foreach (var cellAddress in allCellAddresses)
                {
                    // Start with old format (or default if cell had no formatting)
                    var newFormat = oldFormats.ContainsKey(cellAddress)
                        ? CloneFormatSettings(oldFormats[cellAddress])
                        : new FormatSettings(); // Default format for cells that had no formatting

                    // Override with applied changes (only non-null properties from appliedFormat)
                    MergeAppliedFormatIntoExisting(newFormat, appliedFormat);

                    // Only include cells that have some formatting (old or new)
                    if (HasAnyFormatting(newFormat))
                    {
                        newFormats[cellAddress] = newFormat;
                    }
                }

                Logger.Info($"✅ Inferred {newFormats.Count} new format entries from {oldFormats.Count} old entries");
                return newFormats;
            }
            catch (Exception ex)
            {
                Logger.Info($"❌ Error inferring new formats, falling back to reading: {ex.Message}");
                // Fallback to reading if inference fails
                return _formatManager.GetFormattedCells(targetRange);
            }
        }

        /// <summary>
        /// Gets all cell addresses in a range (not just formatted ones)
        /// </summary>
        public List<string> GetAllCellAddressesInRange(Range targetRange)
        {
            var addresses = new List<string>();

            try
            {
                foreach (Range cell in targetRange)
                {
                    addresses.Add(cell.Address[false, false]);
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting all cell addresses: {ex.Message}");
            }

            return addresses;
        }

        /// <summary>
        /// Creates a deep copy of format settings
        /// </summary>
        public FormatSettings CloneFormatSettings(FormatSettings original)
        {
            if (original == null) return new FormatSettings();

            return new FormatSettings
            {
                Bold = original.Bold,
                Italic = original.Italic,
                FontSize = original.FontSize,
                FontName = original.FontName,
                FontColor = original.FontColor,
                BackgroundColor = original.BackgroundColor,
                Alignment = original.Alignment,
                NumberFormat = original.NumberFormat,
                // Border = CloneBorderSettings(original.Border) // COMMENTED OUT TO STRIP BORDERS
            };
        }

        // /// <summary> // COMMENTED OUT TO STRIP BORDERS FROM FORMATTING
        // /// Creates a deep copy of border settings
        // /// </summary>
        // public BorderSettings CloneBorderSettings(BorderSettings original)
        // {
        //     if (original == null) return null;
        //
        //     return new BorderSettings
        //     {
        //         Top = original.Top,
        //         Bottom = original.Bottom,
        //         Left = original.Left,
        //         Right = original.Right
        //     };
        // }

        /// <summary>
        /// Merges applied format changes into existing format (only non-null properties)
        /// </summary>
        public void MergeAppliedFormatIntoExisting(FormatSettings existing, FormatSettings applied)
        {
            if (applied == null) return;

            // Only override with non-null/non-default values from applied format
            if (applied.Bold.HasValue)
                existing.Bold = applied.Bold.Value;

            if (applied.Italic.HasValue)
                existing.Italic = applied.Italic.Value;

            if (applied.FontSize.HasValue)
                existing.FontSize = applied.FontSize.Value;

            if (!string.IsNullOrEmpty(applied.FontName))
                existing.FontName = applied.FontName;

            if (!string.IsNullOrEmpty(applied.FontColor))
                existing.FontColor = applied.FontColor;

            if (!string.IsNullOrEmpty(applied.BackgroundColor))
                existing.BackgroundColor = applied.BackgroundColor;

            if (!string.IsNullOrEmpty(applied.Alignment))
                existing.Alignment = applied.Alignment;

            if (!string.IsNullOrEmpty(applied.NumberFormat))
                existing.NumberFormat = applied.NumberFormat;

            // Handle border settings - COMMENTED OUT TO STRIP BORDERS
            // if (applied.Border != null)
            // {
            //     if (existing.Border == null)
            //         existing.Border = new BorderSettings();
            //
            //     if (!string.IsNullOrEmpty(applied.Border.Top))
            //         existing.Border.Top = applied.Border.Top;
            //
            //     if (!string.IsNullOrEmpty(applied.Border.Bottom))
            //         existing.Border.Bottom = applied.Border.Bottom;
            //
            //     if (!string.IsNullOrEmpty(applied.Border.Left))
            //         existing.Border.Left = applied.Border.Left;
            //
            //     if (!string.IsNullOrEmpty(applied.Border.Right))
            //         existing.Border.Right = applied.Border.Right;
            // }
        }

        /// <summary>
        /// Checks if format settings contain any formatting (not all defaults)
        /// </summary>
        public bool HasAnyFormatting(FormatSettings format)
        {
            if (format == null) return false;

            return format.Bold.HasValue ||
                   format.Italic.HasValue ||
                   format.FontSize.HasValue ||
                   !string.IsNullOrEmpty(format.FontName) ||
                   !string.IsNullOrEmpty(format.FontColor) ||
                   !string.IsNullOrEmpty(format.BackgroundColor) ||
                   !string.IsNullOrEmpty(format.Alignment) ||
                   !string.IsNullOrEmpty(format.NumberFormat) ||
                   false; // (format.Border != null && HasAnyBorderFormatting(format.Border)); // COMMENTED OUT TO STRIP BORDERS
        }

        // /// <summary> // COMMENTED OUT TO STRIP BORDERS FROM FORMATTING
        // /// Checks if border settings contain any formatting
        // /// </summary>
        // public bool HasAnyBorderFormatting(BorderSettings border)
        // {
        //     if (border == null) return false;
        //
        //     return !string.IsNullOrEmpty(border.Top) ||
        //            !string.IsNullOrEmpty(border.Bottom) ||
        //            !string.IsNullOrEmpty(border.Left) ||
        //            !string.IsNullOrEmpty(border.Right);
        // }
    }
}
