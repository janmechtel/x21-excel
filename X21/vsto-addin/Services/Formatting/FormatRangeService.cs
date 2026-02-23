using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using X21.Excel;
using X21.Logging;
using X21.Models;

namespace X21.Services.Formatting
{

    public class FormatRangeService
    {
        private readonly FormatManager _formatManager;
        private readonly ExcelSelection _excelSelection;

        public FormatRangeService(FormatManager formatManager, ExcelSelection excelSelection)
        {
            _formatManager = formatManager;
            _excelSelection = excelSelection;
        }

        //public Dictionary<string, FormatSettings> ReadExcelRangeFormat(Worksheet worksheet, string range)
        //{
        //    return ReadExcelRangeFormat(worksheet, range, null);
        //}

        /// <summary>
        /// Reads format information for cells in a specified range with optional selective property reading
        /// </summary>
        /// <param name="worksheet">The worksheet containing the range</param>
        /// <param name="range">The range address (e.g., "A1:B10")</param>
        /// <param name="propertiesToRead">List of specific format properties to read (e.g., "bold", "italic"). If null, reads all properties.</param>
        //public Dictionary<string, FormatSettings> ReadExcelRangeFormat(Worksheet worksheet, string range, List<string> propertiesToRead)
        //{
        //    Logger.Info($"📖 Reading formats for all cells in range {range} from worksheet '{worksheet.Name}'" +
        //        (propertiesToRead != null ? $" (selective: {string.Join(", ", propertiesToRead)})" : ""));

        //    // Get the target range
        //    var targetRange = worksheet.Range[range];
        //    if (targetRange == null)
        //    {
        //        throw new ArgumentException($"Invalid range '{range}'");
        //    }

        //    // Get format information for all cells in the range using FormatManager
        //    var cellFormats = _formatManager.GetFormattedCells(targetRange, propertiesToRead);
        //    return cellFormats;
        //}


        public Task<Dictionary<string, FormatSettings>> ReadExcelRangeFormatAsync(Worksheet worksheet, string range, List<string> propertiesToRead, Action<FormatProgressUpdate> progress = null)
        {
            Logger.Info($"📖 Reading formats for all cells in range {range} from worksheet '{worksheet.Name}'" +
                (propertiesToRead != null ? $" (selective: {string.Join(", ", propertiesToRead)})" : ""));

            // Get the target range
            var targetRange = worksheet.Range[range];
            if (targetRange == null)
            {
                throw new ArgumentException($"Invalid range '{range}'");
            }

            // Get format information for all cells in the range using FormatManager
            var cellFormats = _formatManager.GetFormattedCellsAsync(targetRange, propertiesToRead, CancellationToken.None, null, progress);
            return cellFormats;
        }

        /// <summary>
        /// Reads format information for all cells in a specified range (with workbook context)
        /// </summary>
        //public ReadFormatResponse ReadExcelRangeFormat(string sheetName, string workbookName, string range)
        //{
        //    return ReadExcelRangeFormat(sheetName, workbookName, range, null);
        //}

        public Task<ReadFormatResponse> ReadExcelRangeFormatAsync(string sheetName, string workbookName, string range, Action<FormatProgressUpdate> progress = null)
        {
            return ReadExcelRangeFormatAsync(sheetName, workbookName, range, null, progress);
        }

        /// <summary>
        /// Reads format information for cells in a specified range with optional selective property reading (with workbook context)
        /// </summary>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="workbookName">The name of the workbook</param>
        /// <param name="range">The range address (e.g., "A1:B10")</param>
        /// <param name="propertiesToRead">List of specific format properties to read (e.g., "bold", "italic"). If null, reads all properties.</param>
        //public ReadFormatResponse ReadExcelRangeFormat(string sheetName, string workbookName, string range, List<string> propertiesToRead)
        //{
        //    try
        //    {
        //        var worksheet = _excelSelection.GetWorksheet(sheetName, workbookName);
        //        if (worksheet == null)
        //        {
        //            throw new ArgumentException($"Worksheet '{sheetName}' in workbook '{workbookName}' not found");
        //        }

        //        Logger.Info($"Reading formats with selective properties" +
        //            (propertiesToRead != null ? $": {string.Join(", ", propertiesToRead)}" : " (all properties)"));

        //        var cellFormats = ReadExcelRangeFormat(worksheet, range, propertiesToRead);
        //        return new ReadFormatResponse
        //        {
        //            Success = true,
        //            Message = $"Successfully read formats for range {range}" +
        //                (propertiesToRead != null ? $" (selective: {string.Join(", ", propertiesToRead)})" : ""),
        //            Worksheet = worksheet.Name,
        //            WorkbookName = workbookName,
        //            Range = range,
        //            CellFormats = cellFormats
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.LogException(ex);
        //        return new ReadFormatResponse
        //        {
        //            Success = false,
        //            Message = $"Failed to read format for range {range}: {ex.Message}",
        //            Worksheet = sheetName,
        //            WorkbookName = workbookName,
        //            Range = range,
        //            CellFormats = new Dictionary<string, FormatSettings>()
        //        };
        //    }
        //}

        public async Task<ReadFormatResponse> ReadExcelRangeFormatAsync(string sheetName, string workbookName, string range, List<string> propertiesToRead, Action<FormatProgressUpdate> progress = null)
        {
            try
            {
                var worksheet = _excelSelection.GetWorksheet(sheetName, workbookName);
                if (worksheet == null)
                {
                    throw new ArgumentException($"Worksheet '{sheetName}' in workbook '{workbookName}' not found");
                }

                Logger.Info($"Reading formats with selective properties" +
                    (propertiesToRead != null ? $": {string.Join(", ", propertiesToRead)}" : " (all properties)"));

                var cellFormats = await ReadExcelRangeFormatAsync(worksheet, range, propertiesToRead, progress);
                return new ReadFormatResponse
                {
                    Success = true,
                    Message = $"Successfully read formats for range {range}" +
                        (propertiesToRead != null ? $" (selective: {string.Join(", ", propertiesToRead)})" : ""),
                    Worksheet = worksheet.Name,
                    WorkbookName = workbookName,
                    Range = range,
                    CellFormats = cellFormats
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new ReadFormatResponse
                {
                    Success = false,
                    Message = $"Failed to read format for range {range}: {ex.Message}",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    Range = range,
                    CellFormats = new Dictionary<string, FormatSettings>()
                };
            }
        }
    }
}
