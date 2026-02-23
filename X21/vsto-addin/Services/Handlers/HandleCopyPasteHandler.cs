using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using X21.Common.Data;
using X21.Common.Model;
using X21.Models;
using X21.Excel;
using X21.Services.Formatting;
using Microsoft.Office.Interop.Excel;
using System.Text.Json.Serialization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using X21.Utils;

namespace X21.Services.Handlers
{
    public class HandleCopyPasteHandler : BaseExcelApiService
    {
        public HandleCopyPasteHandler(Container container) : base(container)
        {
        }

        public async Task HandleCopyPaste(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var copyPasteRequest = JsonSerializer.Deserialize<CopyPasteRequest>(requestBody, _jsonOptions);

                if (copyPasteRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrWhiteSpace(copyPasteRequest.DestinationWorksheet) &&
                    string.IsNullOrWhiteSpace(copyPasteRequest.SourceWorksheet))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "destinationWorksheet or sourceWorksheet is required");
                    return;
                }

                if (string.IsNullOrWhiteSpace(copyPasteRequest.SourceRange) ||
                    string.IsNullOrWhiteSpace(copyPasteRequest.DestinationRange))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "sourceRange and destinationRange are required");
                    return;
                }

                // Default missing worksheet names to the other value to keep intent clear
                if (string.IsNullOrWhiteSpace(copyPasteRequest.DestinationWorksheet))
                {
                    copyPasteRequest.DestinationWorksheet = copyPasteRequest.SourceWorksheet;
                }

                if (string.IsNullOrWhiteSpace(copyPasteRequest.SourceWorksheet))
                {
                    copyPasteRequest.SourceWorksheet = copyPasteRequest.DestinationWorksheet;
                }

                var result = InvokeExcel(() => CopyPasteRange(copyPasteRequest));

                await SendJsonResponse(response, result);

                var sanitizedSourceWorksheet = SanitizeForLog(copyPasteRequest.SourceWorksheet);
                var sanitizedSourceRange = SanitizeForLog(copyPasteRequest.SourceRange);
                var sanitizedDestinationWorksheet = SanitizeForLog(copyPasteRequest.DestinationWorksheet);
                var sanitizedDestinationRange = SanitizeForLog(copyPasteRequest.DestinationRange);
                var sanitizedPasteType = SanitizeForLog(copyPasteRequest.PasteType) ?? "all";
                var sanitizedInsertMode = SanitizeForLog(copyPasteRequest.InsertMode) ?? "none";
                Logger.Info($"Successfully copy/pasted {sanitizedSourceWorksheet}!{sanitizedSourceRange} -> {sanitizedDestinationWorksheet}!{sanitizedDestinationRange} | pasteType={sanitizedPasteType}, insertMode={sanitizedInsertMode}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (JsonException jsonEx)
            {
                Logger.LogException(jsonEx);
                await WriteJsonErrorAsync(response, $"Invalid JSON in request body: {jsonEx.Message}", 400);
            }
            catch (Exception ex) when (IsNonFatal(ex))
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error executing copy/paste: {ex.Message}", 500);
            }
        }


        public async Task HandleDeleteCells(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var deleteCellsRequest = JsonSerializer.Deserialize<DeleteCellsRequest>(requestBody, _jsonOptions);

                if (deleteCellsRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrWhiteSpace(deleteCellsRequest.Worksheet) ||
                    string.IsNullOrWhiteSpace(deleteCellsRequest.Range))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet and range are required");
                    return;
                }

                var result = InvokeExcel(() => DeleteCells(deleteCellsRequest));

                await SendJsonResponse(response, result);
                var sanitizedWorksheet = SanitizeForLog(deleteCellsRequest.Worksheet);
                var sanitizedRange = SanitizeForLog(deleteCellsRequest.Range);
                var sanitizedShift = SanitizeForLog(deleteCellsRequest.ShiftDirection) ?? "left/up";
                Logger.Info($"Deleted cells at {sanitizedWorksheet}!{sanitizedRange} with shift {sanitizedShift}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex) when (IsNonFatal(ex))
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error deleting cells: {ex.Message}", 500);
            }
        }


        private CopyPasteResponse CopyPasteRange(CopyPasteRequest request)
        {
            try
            {
                var sanitizedSourceWorksheet = SanitizeForLog(request.SourceWorksheet);
                var sanitizedSourceRange = SanitizeForLog(request.SourceRange);
                var sanitizedDestinationWorksheet = SanitizeForLog(request.DestinationWorksheet);
                var sanitizedDestinationRange = SanitizeForLog(request.DestinationRange);
                var sanitizedPasteType = SanitizeForLog(request.PasteType);
                var sanitizedInsertMode = SanitizeForLog(request.InsertMode);
                Logger.Info($"Copy/paste request: {sanitizedSourceWorksheet}!{sanitizedSourceRange} -> {sanitizedDestinationWorksheet}!{sanitizedDestinationRange} (pasteType={sanitizedPasteType}, insertMode={sanitizedInsertMode})");

                var sourceWorkbookName = string.IsNullOrWhiteSpace(request.SourceWorkbookName)
                    ? request.WorkbookName
                    : request.SourceWorkbookName;

                var warnings = new List<string>();

                var sourceWorksheet = ResolveWorksheetForRead(request.SourceWorksheet, sourceWorkbookName);
                var destinationWorksheet = GetWorksheetFromActiveWorkbook(request.DestinationWorksheet, request.WorkbookName);

                var sourceRange = sourceWorksheet.Range[request.SourceRange];
                var destinationRange = destinationWorksheet.Range[request.DestinationRange];

                if (sourceRange == null || destinationRange == null)
                {
                    throw new ArgumentException("Invalid source or destination range");
                }

                int rowCount = sourceRange.Rows.Count;
                int colCount = sourceRange.Columns.Count;
                var targetRange = destinationRange.Resize[rowCount, colCount];
                var targetAddress = targetRange.get_Address(RowAbsolute: false, ColumnAbsolute: false);

                // Capture existing values before making changes (used for revert and previews)
                var previousCells = GetRangeDataFromWorksheet(destinationWorksheet, targetAddress) ?? new List<Cell>();
                var oldValues = new ReadRangeResponse
                {
                    CellValues = previousCells.ToDictionary(
                        c => c.Address,
                        c => new CellValue
                        {
                            Value = InvariantValueFormatter.ToInvariantString(c.Value),
                            Formula = c.Formula ?? string.Empty
                        })
                };

                var insertDirection = MapInsertDirection(request.InsertMode);
                if (insertDirection.HasValue)
                {
                    targetRange.Insert(insertDirection.Value);
                }

                var pasteType = MapPasteType(request.PasteType);
                var operation = MapPasteOperation(request.Operation);
                bool skipBlanks = request.SkipBlanks ?? false;
                bool transpose = request.Transpose ?? false;

                sourceRange.Copy(Type.Missing);
                targetRange.PasteSpecial(pasteType, operation, skipBlanks, transpose);

                if (request.IncludeColumnWidths == true)
                {
                    try
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var sourceColumn = sourceRange.Columns[col];
                            var targetColumn = targetRange.Columns[col];
                            targetColumn.ColumnWidth = sourceColumn.ColumnWidth;
                        }
                    }
                    catch (COMException widthEx)
                    {
                        Logger.Info($"Warning: Failed to copy column widths: {SanitizeForLog(widthEx.Message)}");
                        warnings.Add("Failed to copy column widths.");
                    }
                }

                // Clear clipboard to avoid lingering copy state.
                destinationWorksheet.Application.CutCopyMode = (XlCutCopyMode)0;

                Logger.Info($"Copy/paste completed into {targetAddress} ({rowCount}x{colCount})");

                return new CopyPasteResponse
                {
                    Success = true,
                    Message = $"Copied {rowCount}x{colCount} from {request.SourceWorksheet}!{request.SourceRange} to {request.DestinationWorksheet}!{targetAddress}",
                    WorkbookName = (destinationWorksheet.Parent as Workbook)?.Name ?? request.WorkbookName,
                    SourceWorksheet = request.SourceWorksheet,
                    SourceRange = request.SourceRange,
                    DestinationWorksheet = request.DestinationWorksheet,
                    DestinationRange = targetAddress,
                    PasteType = NormalizePasteType(request.PasteType),
                    InsertMode = NormalizeInsertMode(request.InsertMode),
                    RowsCopied = rowCount,
                    ColumnsCopied = colCount,
                    OldValues = oldValues,
                    Warnings = warnings.Count > 0 ? warnings : null
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
                return CreateCopyPasteFailureResponse(request, ex.Message);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
                return CreateCopyPasteFailureResponse(request, ex.Message);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                throw;
            }
        }

        private DeleteCellsResponse DeleteCells(DeleteCellsRequest request)
        {
            try
            {
                Logger.Info($"Deleting cells at {request.Worksheet}!{request.Range} (shift {request.ShiftDirection ?? "left"})");

                var worksheet = GetWorksheetFromActiveWorkbook(request.Worksheet, request.WorkbookName);
                var targetRange = worksheet.Range[request.Range];
                if (targetRange == null)
                {
                    throw new ArgumentException($"Invalid range '{request.Range}'");
                }

                var shift = MapDeleteShiftDirection(request.ShiftDirection);
                targetRange.Delete(shift);

                var cleanedAddress = targetRange.get_Address(RowAbsolute: false, ColumnAbsolute: false);

                return new DeleteCellsResponse
                {
                    Success = true,
                    Message = $"Deleted cells at {cleanedAddress} with shift {request.ShiftDirection ?? "left"}",
                    Worksheet = worksheet.Name,
                    WorkbookName = (worksheet.Parent as Workbook)?.Name ?? request.WorkbookName,
                    Range = cleanedAddress,
                    ShiftDirection = shift == XlDeleteShiftDirection.xlShiftToLeft ? "left" : "up"
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex) when (IsNonFatal(ex))
            {
                Logger.LogException(ex);
                return new DeleteCellsResponse
                {
                    Success = false,
                    Message = $"Failed to delete cells: {ex.Message}",
                    Worksheet = request.Worksheet,
                    WorkbookName = request.WorkbookName,
                    Range = request.Range,
                    ShiftDirection = request.ShiftDirection
                };
            }
        }

        private CopyPasteResponse CreateCopyPasteFailureResponse(CopyPasteRequest request, string message)
        {
            return new CopyPasteResponse
            {
                Success = false,
                Message = $"Failed to copy/paste: {message}",
                WorkbookName = request.WorkbookName,
                SourceWorksheet = request.SourceWorksheet,
                SourceRange = request.SourceRange,
                DestinationWorksheet = request.DestinationWorksheet,
                DestinationRange = request.DestinationRange,
                PasteType = NormalizePasteType(request.PasteType),
                InsertMode = NormalizeInsertMode(request.InsertMode)
            };
        }

        private static bool IsNonFatal(Exception ex)
        {
            return !(ex is OutOfMemoryException ||
                     ex is ThreadAbortException ||
                     ex is StackOverflowException ||
                     ex is ThreadInterruptedException ||
                     ex is AccessViolationException);
        }

        private static string SanitizeForLog(string value)
        {
            return value?.Replace("\r", "").Replace("\n", "");
        }

        public class CopyPasteResponse
        {
            [JsonPropertyName("success")]
            public bool Success { get; set; }

            [JsonPropertyName("message")]
            public string Message { get; set; }

            [JsonPropertyName("workbookName")]
            public string WorkbookName { get; set; }

            [JsonPropertyName("sourceWorksheet")]
            public string SourceWorksheet { get; set; }

            [JsonPropertyName("sourceRange")]
            public string SourceRange { get; set; }

            [JsonPropertyName("destinationWorksheet")]
            public string DestinationWorksheet { get; set; }

            [JsonPropertyName("destinationRange")]
            public string DestinationRange { get; set; }

            [JsonPropertyName("pasteType")]
            public string PasteType { get; set; }

            [JsonPropertyName("insertMode")]
            public string InsertMode { get; set; }

            [JsonPropertyName("rowsCopied")]
            public int RowsCopied { get; set; }

            [JsonPropertyName("columnsCopied")]
            public int ColumnsCopied { get; set; }

            [JsonPropertyName("oldValues")]
            public ReadRangeResponse OldValues { get; set; }

            [JsonPropertyName("warnings")]
            public List<string> Warnings { get; set; }
        }

        private XlPasteType MapPasteType(string pasteType)
        {
            switch (pasteType?.ToLowerInvariant())
            {
                case "values":
                    return XlPasteType.xlPasteValues;
                case "formats":
                    return XlPasteType.xlPasteFormats;
                case "formulas":
                    return XlPasteType.xlPasteFormulas;
                case "formulas_and_number_formats":
                    return XlPasteType.xlPasteFormulasAndNumberFormats;
                case "values_and_number_formats":
                    return XlPasteType.xlPasteValuesAndNumberFormats;
                case "column_widths":
                    return XlPasteType.xlPasteColumnWidths;
                default:
                    return XlPasteType.xlPasteAll;
            }
        }

        private XlPasteSpecialOperation MapPasteOperation(string operation)
        {
            switch (operation?.ToLowerInvariant())
            {
                case "add":
                    return XlPasteSpecialOperation.xlPasteSpecialOperationAdd;
                case "subtract":
                    return XlPasteSpecialOperation.xlPasteSpecialOperationSubtract;
                case "multiply":
                    return XlPasteSpecialOperation.xlPasteSpecialOperationMultiply;
                case "divide":
                    return XlPasteSpecialOperation.xlPasteSpecialOperationDivide;
                default:
                    return XlPasteSpecialOperation.xlPasteSpecialOperationNone;
            }
        }

        private XlInsertShiftDirection? MapInsertDirection(string insertMode)
        {
            switch (insertMode?.ToLowerInvariant())
            {
                case "shiftright":
                case "shift_right":
                case "right":
                    return XlInsertShiftDirection.xlShiftToRight;
                case "shiftdown":
                case "shift_down":
                case "down":
                    return XlInsertShiftDirection.xlShiftDown;
                default:
                    return null;
            }
        }

        private XlDeleteShiftDirection MapDeleteShiftDirection(string shiftDirection)
        {
            switch (shiftDirection?.ToLowerInvariant())
            {
                case "up":
                case "shiftup":
                case "shift_up":
                    return XlDeleteShiftDirection.xlShiftUp;
                default:
                    return XlDeleteShiftDirection.xlShiftToLeft;
            }
        }

        private string NormalizePasteType(string pasteType)
        {
            switch (pasteType?.ToLowerInvariant())
            {
                case "values":
                case "formats":
                case "formulas":
                case "formulas_and_number_formats":
                case "values_and_number_formats":
                case "column_widths":
                    return pasteType.ToLowerInvariant();
                default:
                    return "all";
            }
        }

        private string NormalizeInsertMode(string insertMode)
        {
            switch (insertMode?.ToLowerInvariant())
            {
                case "shiftright":
                case "shift_right":
                case "right":
                    return "shift_right";
                case "shiftdown":
                case "shift_down":
                case "down":
                    return "shift_down";
                default:
                    return "none";
            }
        }


    }
}
