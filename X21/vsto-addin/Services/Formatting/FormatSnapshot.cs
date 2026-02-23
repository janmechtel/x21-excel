using System;
using X21.Logging;

namespace X21.Services.Formatting
{
    /// <summary>
    /// Immutable snapshot of range formatting captured via bulk COM reads on the STA thread.
    /// All arrays are zero-based and sized to RowCount x ColumnCount to allow CPU-only readers.
    /// </summary>
    public class FormatSnapshot
    {
        public int RowCount { get; }
        public int ColumnCount { get; }
        public int StartRow { get; }
        public int StartColumn { get; }
        public int CellCount => RowCount * ColumnCount;

        public string[,] Addresses { get; }

        public object[,] Bold { get; }
        public object[,] Italic { get; }
        public object[,] Underline { get; }
        public object[,] FontSize { get; }
        public object[,] FontName { get; }
        public object[,] FontColor { get; }
        public object[,] BackgroundColor { get; }
        public object[,] BackgroundColorIndex { get; }
        public object[,] NumberFormat { get; }
        public object[,] Alignment { get; }

        public FormatSnapshot(
            int startRow,
            int startColumn,
            int rowCount,
            int columnCount,
            string[,] addresses,
            object[,] bold,
            object[,] italic,
            object[,] underline,
            object[,] fontSize,
            object[,] fontName,
            object[,] fontColor,
            object[,] backgroundColor,
            object[,] backgroundColorIndex,
            object[,] numberFormat,
            object[,] alignment)
        {
            StartRow = startRow;
            StartColumn = startColumn;
            RowCount = rowCount;
            ColumnCount = columnCount;
            Addresses = addresses;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            FontSize = fontSize;
            FontName = fontName;
            FontColor = fontColor;
            BackgroundColor = backgroundColor;
            BackgroundColorIndex = backgroundColorIndex;
            NumberFormat = numberFormat;
            Alignment = alignment;
        }

        public static string ColumnNumberToName(int colNumber)
        {
            if (colNumber <= 0) throw new ArgumentOutOfRangeException(nameof(colNumber));
            var dividend = colNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static string[,] BuildAddresses(int startRow, int startColumn, int rows, int cols)
        {
            var addresses = new string[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                var rowNumber = startRow + r;
                for (var c = 0; c < cols; c++)
                {
                    var columnName = ColumnNumberToName(startColumn + c);
                    addresses[r, c] = $"{columnName}{rowNumber}";
                }
            }
            return addresses;
        }
    }
}
