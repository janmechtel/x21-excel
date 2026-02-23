using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using X21.Logging;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Builds a unified format snapshot for a given range using bulk COM reads on the STA thread.
    /// </summary>
    public class FormatSnapshotBuilder
    {
        public FormatSnapshot Build(Range targetRange, List<string> propertiesToRead)
        {
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            var propsLower = propertiesToRead == null
                ? null
                : new HashSet<string>(propertiesToRead.ConvertAll(p => p.ToLowerInvariant()));
            var isFullSnapshot = propsLower == null || propsLower.Count == 0;
            bool Needs(string prop) => isFullSnapshot || propsLower.Contains(prop.ToLowerInvariant());

            var sw = Stopwatch.StartNew();
            var rows = targetRange.Rows.Count;
            var cols = targetRange.Columns.Count;
            var startRow = targetRange.Row;
            var startCol = targetRange.Column;
            var addresses = FormatSnapshot.BuildAddresses(startRow, startCol, rows, cols);

            Logger.Info($"FormatSnapshotBuilder: building snapshot for {targetRange.Address[false, false]} ({rows}x{cols}, {rows * cols} cells)" +
                        (!isFullSnapshot ? $" selective: {string.Join(", ", propsLower)}" : " (full)"));

            object[,] bold = null, italic = null, underline = null, fontSize = null, fontName = null,
                fontColor = null, backgroundColor = null, backgroundColorIndex = null, numberFormat = null, alignment = null;

            if (Needs("bold")) bold = NormalizeTo2D(targetRange.Font.Bold, rows, cols);
            if (Needs("italic")) italic = NormalizeTo2D(targetRange.Font.Italic, rows, cols);
            if (Needs("underline")) underline = NormalizeTo2D(targetRange.Font.Underline, rows, cols);
            if (Needs("fontsize")) fontSize = NormalizeTo2D(targetRange.Font.Size, rows, cols);
            if (Needs("fontname")) fontName = NormalizeTo2D(targetRange.Font.Name, rows, cols);
            if (Needs("fontcolor")) fontColor = NormalizeTo2D(targetRange.Font.Color, rows, cols);
            if (Needs("backgroundcolor"))
            {
                backgroundColor = NormalizeTo2D(targetRange.Interior.Color, rows, cols);
                backgroundColorIndex = NormalizeTo2D(targetRange.Interior.ColorIndex, rows, cols);
            }
            if (Needs("numberformat")) numberFormat = NormalizeTo2D(targetRange.NumberFormat, rows, cols);
            if (Needs("alignment")) alignment = NormalizeTo2D(targetRange.HorizontalAlignment, rows, cols);

            Logger.Info($"FormatSnapshotBuilder: snapshot captured in {sw.ElapsedMilliseconds} ms for {rows * cols} cells");

            return new FormatSnapshot(
                startRow,
                startCol,
                rows,
                cols,
                addresses,
                bold,
                italic,
                underline,
                fontSize,
                fontName,
                fontColor,
                backgroundColor,
                backgroundColorIndex,
                numberFormat,
                alignment);
        }

        /// <summary>
        /// Normalize Excel COM return values into a 0-based rectangular object[,]
        /// </summary>
        private static object[,] NormalizeTo2D(object raw, int rows, int cols)
        {
            if (raw is Array array && array.Rank == 2)
            {
                var result = new object[rows, cols];
                var rowLb = array.GetLowerBound(0);
                var colLb = array.GetLowerBound(1);

                for (var r = 0; r < rows; r++)
                {
                    for (var c = 0; c < cols; c++)
                    {
                        result[r, c] = array.GetValue(r + rowLb, c + colLb);
                    }
                }

                return result;
            }

            // Scalar value – replicate across the 2D grid
            var replicated = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    replicated[r, c] = raw;
                }
            }
            return replicated;
        }
    }
}
