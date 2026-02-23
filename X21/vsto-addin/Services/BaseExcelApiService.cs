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

namespace X21.Services
{
    public abstract class BaseExcelApiService : Component
    {
        private readonly ExcelStaDispatcher _excelDispatcher;



        protected const string ErrorWorkbookNotFound = "WORKBOOK_NOT_FOUND";
        protected const string ErrorAmbiguousWorkbook = "AMBIGUOUS_WORKBOOK";
        protected const string ErrorCrossWorkbookWriteBlocked = "CROSS_WORKBOOK_WRITE_BLOCKED";
        private const double ColumnWidthTolerance = 0.01;

        private enum ColumnWidthMode
        {
            Always,
            Smart,
            Never
        }

        protected BaseExcelApiService(Container container) : base(container)
        {
            _excelDispatcher = Container.Resolve<ExcelStaDispatcher>();

        }

        protected readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        protected T InvokeExcel<T>(Func<T> func)
        {
            return _excelDispatcher.InvokeExcel(func);
        }

        protected void InvokeExcel(System.Action action)
        {
            _excelDispatcher.InvokeExcel(action);
        }

        protected Workbook ResolveWorkbook(Application application, string workbookName)
        {
            // Sanitize input for log line-forgery protection
            var sanitizedWorkbookName = workbookName?.Replace("\r", "").Replace("\n", "");
            Logger.Info($"[WorkbookResolver] Request for workbook='{sanitizedWorkbookName ?? "(null)"}'");


            if (application == null)
            {
                throw new WorkbookResolutionException("Excel application is unavailable", ErrorWorkbookNotFound);
            }

            if (string.IsNullOrWhiteSpace(workbookName))
            {
                var activeWorkbook = application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    throw new WorkbookResolutionException("No active workbook available", ErrorWorkbookNotFound);
                }

                Logger.Info($"[WorkbookResolver] Using active workbook '{activeWorkbook.Name}'");
                return activeWorkbook;
            }

            var matches = new List<Workbook>();
            var nameFromPath = string.IsNullOrWhiteSpace(workbookName)
                ? null
                : Path.GetFileName(workbookName);

            foreach (Workbook wb in application.Workbooks)
            {
                if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(nameFromPath) &&
                     wb.Name.Equals(nameFromPath, StringComparison.OrdinalIgnoreCase)))
                {
                    matches.Add(wb);
                }
                else
                {
                    try
                    {
                        if (wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            matches.Add(wb);
                        }
                    }
                    catch (COMException)
                    {
                        // Ignore FullName access issues and continue
                    }
                }
            }

            Logger.Info($"[WorkbookResolver] Matches found: {matches.Count}");

            if (matches.Count == 0)
            {
                throw new WorkbookResolutionException(
                    $"Workbook '{workbookName}' not found",
                    ErrorWorkbookNotFound
                );
            }

            if (matches.Count > 1)
            {
                var candidates = matches.Select(wb => new WorkbookCandidate
                {
                    WorkbookName = wb.Name,
                    WorkbookFullName = SafeGetWorkbookFullName(wb)
                });

                throw new WorkbookResolutionException(
                    $"Multiple workbooks found matching '{workbookName}'",
                    ErrorAmbiguousWorkbook,
                    candidates.ToList()
                );
            }

            Logger.Info($"[WorkbookResolver] Resolved to '{matches.First().Name}'");
            return matches.First();
        }
        protected Workbook ResolveActiveWorkbookForWrite(string providedWorkbookName)
        {
            var application = Container.Resolve<Application>();
            var activeWorkbook = application.ActiveWorkbook;
            var workbook = ResolveWorkbook(application, providedWorkbookName);

            Logger.Info($"[WorkbookResolver:Write] Resolved='{workbook?.Name}', Requested='{providedWorkbookName ?? "(null)"}'");

            if (activeWorkbook != null &&
                workbook != null &&
                !workbook.Name.Equals(activeWorkbook.Name, StringComparison.OrdinalIgnoreCase))
            {
                Logger.Info($"[WorkbookResolver:Write] Writing to non-active workbook '{workbook.Name}' (active='{activeWorkbook.Name}')");
            }

            return workbook;
        }

        /// <summary>
        /// Get all cell data from a specific worksheet range
        /// </summary>
        protected List<Cell> GetRangeDataFromWorksheet(Worksheet worksheet, string range)
        {
            var excelSelection = Container.Resolve<ExcelSelection>();
            return excelSelection.GetRangeDataFromWorksheet(worksheet, range);
        }

        protected Worksheet GetWorksheetFromActiveWorkbook(string sheetName, string providedWorkbookName)
        {
            var workbook = ResolveActiveWorkbookForWrite(providedWorkbookName);

            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }

            throw new WorkbookResolutionException(
                $"Worksheet '{sheetName}' not found in workbook '{workbook.Name}'",
                ErrorWorkbookNotFound,
                new List<WorkbookCandidate>
                {
                    new WorkbookCandidate
                    {
                        WorkbookName = workbook.Name,
                        WorkbookFullName = SafeGetWorkbookFullName(workbook)
                    }
                }
            );

        }

        protected string SafeGetWorkbookFullName(Workbook workbook)
        {
            try
            {
                return workbook.FullName;
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
                return string.Empty;
            }
        }
        protected Worksheet ResolveWorksheetForRead(string sheetName, string workbookName)
        {
            var application = Container.Resolve<Application>();
            var workbook = ResolveWorkbook(application, workbookName);

            Logger.Info($"[WorksheetResolver] Searching sheet '{sheetName}' in workbook '{workbook.Name}'");

            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Info($"[WorksheetResolver] Found worksheet '{ws.Name}' in workbook '{workbook.Name}'");
                    return ws;
                }
            }

            throw new Exception($"Worksheet '{sheetName}' not found in workbook '{workbook.Name}'");
        }
        protected async Task WriteJsonErrorAsync(HttpListenerResponse response, string message, int statusCode = 400, string errorCode = null, IEnumerable<WorkbookCandidate> candidates = null)
        {
            var errorResponse = new
            {
                error = message,
                errorCode,
                candidates
            };
            await SendJsonResponse(response, errorResponse, statusCode);
        }

        protected async Task SendJsonResponse(HttpListenerResponse response, object data, int statusCode = 200)
        {
            response.StatusCode = statusCode;
            response.ContentType = "application/json";

            var jsonResponse = JsonSerializer.Serialize(data, _jsonOptions);
            var buffer = Encoding.UTF8.GetBytes(jsonResponse);

            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.Close();
        }

        protected Task<WriteRangeResponse> WriteExcelRangeAsync(
            string sheetName,
            string workbookName,
            string range,
            object[][] values,
            string columnWidthMode = null)
        {
            try
            {
                Logger.Info($"📝 Writing to {sheetName}!{range} with {values.Length} rows in workbook {workbookName}");

                var result = InvokeExcel(() =>
                    WriteExcelRangeSynchronous(sheetName, workbookName, range, values, columnWidthMode));
                return Task.FromResult(result);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);

                var response = new WriteRangeResponse
                {
                    Success = false,
                    Message = $"Failed to write range {range} in sheet {sheetName} in workbook {workbookName}: {ex.Message}",
                };
                return Task.FromResult(response);
            }
        }

        protected WriteRangeResponse WriteExcelRangeSynchronous(
            string sheetName,
            string workbookName,
            string range,
            object[][] values,
            string columnWidthMode = null)
        {
            var totalStopwatch = Stopwatch.StartNew();
            try
            {
                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Validate input values first
                if (values == null || values.Length == 0)
                {
                    return new WriteRangeResponse
                    {
                        Success = false,
                        Message = "No values provided to write",
                    };
                }

                // Get target range and infer dimensions from it
                var rangeStopwatch = Stopwatch.StartNew();
                Range targetRange = worksheet.Range[range];
                int rowCount = targetRange.Rows.Count;
                int colCount = targetRange.Columns.Count;
                rangeStopwatch.Stop();

                if (values.Length != rowCount)
                {
                    return new WriteRangeResponse
                    {
                        Success = false,
                        Message = $"Values row count ({values.Length}) does not match range rows ({rowCount}) for {range}",
                    };
                }

                for (int row = 0; row < rowCount; row++)
                {
                    if (values[row] == null || values[row].Length != colCount)
                    {
                        var actualCols = values[row]?.Length ?? 0;
                        return new WriteRangeResponse
                        {
                            Success = false,
                            Message = $"Values column count ({actualCols}) does not match range columns ({colCount}) for {range} at row {row + 1}",
                        };
                    }
                }

                int totalCells = rowCount * colCount;
                Logger.Info($"📊 Writing {rowCount} rows × {colCount} columns to {range} = {totalCells} cells");

                // Convert object[][] to object[,] for bulk write
                var conversionStopwatch = Stopwatch.StartNew();
                var values2D = new object[rowCount, colCount];

                // Copy values from jagged array to 2D array
                for (int row = 0; row < rowCount; row++)
                {
                    var sourceRow = values[row];
                    for (int col = 0; col < colCount; col++)
                    {
                        values2D[row, col] = NormalizeJsonValue(sourceRow[col]);
                    }
                }
                conversionStopwatch.Stop();

                // Bulk write entire range at once (much faster than cell-by-cell)
                var writeStopwatch = Stopwatch.StartNew();
                try
                {
                    // Excel's Formula accepts object arrays where:
                    // - Strings starting with '=' are treated as formulas (invariant / English)
                    // - Other values are treated as constants
                    targetRange.FormulaLocal = values2D;
                    AdjustColumnWidthsAfterWrite(targetRange, columnWidthMode);
                }
                catch (Exception writeEx)
                {
                    Logger.Info($"❌ Error during write/autofit to range {range}: {writeEx.Message}");
                    throw;
                }
                finally
                {
                    writeStopwatch.Stop();
                }

                totalStopwatch.Stop();

                // Log detailed performance metrics
                Logger.Info($"⚡ PERFORMANCE: Total={totalStopwatch.ElapsedMilliseconds}ms | " +
                           $"Range={rangeStopwatch.ElapsedMilliseconds}ms | " +
                           $"Conversion={conversionStopwatch.ElapsedMilliseconds}ms | " +
                           $"ExcelWrite={writeStopwatch.ElapsedMilliseconds}ms | " +
                           $"Throughput={totalCells / (totalStopwatch.ElapsedMilliseconds / 1000.0):F0} cells/sec");

                Logger.Info($"✅ Successfully wrote {totalCells} cells to range {range} in bulk operation");

                return new WriteRangeResponse
                {
                    Success = true,
                    Message = $"Successfully wrote {totalCells} cells to range {range} in {totalStopwatch.ElapsedMilliseconds}ms",
                };
            }
            catch (WorkbookResolutionException)
            {
                totalStopwatch.Stop();
                throw;
            }
            catch (Exception ex)
            {
                totalStopwatch.Stop();
                if (ex is OperationCanceledException || ex is ThreadAbortException)
                {
                    throw;
                }
                Logger.Info($"❌ Error during write to range {range}: {ex.Message}");
                Logger.LogException(ex);
                return new WriteRangeResponse
                {
                    Success = false,
                    Message = $"Failed to write range {range} in sheet {sheetName}: {ex.Message}",
                };
            }
        }

        private ColumnWidthMode NormalizeColumnWidthMode(string columnWidthMode)
        {
            if (string.IsNullOrWhiteSpace(columnWidthMode)) return ColumnWidthMode.Smart;

            switch (columnWidthMode.Trim().ToLowerInvariant())
            {
                case "always":
                    return ColumnWidthMode.Always;
                case "never":
                    return ColumnWidthMode.Never;
                default:
                    return ColumnWidthMode.Smart;
            }
        }

        private void AdjustColumnWidthsAfterWrite(Range targetRange, string columnWidthMode)
        {
            if (targetRange == null) return;

            var mode = NormalizeColumnWidthMode(columnWidthMode);
            if (mode == ColumnWidthMode.Never) return;

            var stopwatch = Stopwatch.StartNew();
            bool restoreScreenUpdating = false;
            bool previousScreenUpdating = true;
            Worksheet worksheet = null;

            try
            {
                var application = Container.Resolve<Application>();
                var screenLimit = GetVisibleColumnWidthLimit(application);
                try
                {
                    worksheet = targetRange.Worksheet;
                }
                catch (COMException ex)
                {
                    Logger.LogException(ex);
                }
                if (application != null)
                {
                    previousScreenUpdating = application.ScreenUpdating;
                    application.ScreenUpdating = false;
                    restoreScreenUpdating = true;
                }

                const double NonHeaderMaxWidth = 60.0;
                const double HeaderMaxWidth = 90.0;

                double standardWidth = 0;
                bool hasStandardWidth = false;
                if (worksheet != null)
                {
                    try
                    {
                        standardWidth = worksheet.StandardWidth;
                        hasStandardWidth = true;
                    }
                    catch (COMException ex)
                    {
                        Logger.LogException(ex);
                    }
                }

                var columnCount = targetRange.Columns.Count;
                for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    Range columnRange = null;
                    Range entireColumn = null;
                    Range headerCell = null;
                    try
                    {
                        columnRange = targetRange.Columns[columnIndex] as Range;
                        if (columnRange == null) continue;

                        entireColumn = columnRange.EntireColumn as Range;
                        if (entireColumn == null) continue;

                        var previousWidth = entireColumn.ColumnWidth;
                        if (mode == ColumnWidthMode.Smart && hasStandardWidth &&
                            Math.Abs(previousWidth - standardWidth) > ColumnWidthTolerance)
                        {
                            continue;
                        }

                        try
                        {
                            entireColumn.AutoFit();
                        }
                        catch (COMException ex)
                        {
                            Logger.LogException(ex);
                            continue;
                        }

                        if (worksheet != null)
                        {
                            headerCell = worksheet.Cells[1, entireColumn.Column] as Range;
                        }

                        var fittedWidth = entireColumn.ColumnWidth;
                        var baseLimit = headerCell?.Value2 != null ? HeaderMaxWidth : NonHeaderMaxWidth;
                        var effectiveLimit = screenLimit > 0 ? Math.Min(baseLimit, screenLimit) : baseLimit;
                        var limitedWidth = Math.Min(fittedWidth, effectiveLimit);
                        var finalWidth = Math.Max(previousWidth, limitedWidth);

                        if (Math.Abs(finalWidth - fittedWidth) > ColumnWidthTolerance)
                        {
                            entireColumn.ColumnWidth = finalWidth;
                        }
                    }
                    finally
                    {
                        if (headerCell != null)
                        {
                            Marshal.ReleaseComObject(headerCell);
                        }
                        if (entireColumn != null)
                        {
                            Marshal.ReleaseComObject(entireColumn);
                        }
                        if (columnRange != null)
                        {
                            Marshal.ReleaseComObject(columnRange);
                        }
                    }
                }
            }
            catch (OutOfMemoryException ex)
            {
                Logger.Info($"Auto-fit skipped due to memory limits: {ex.Message}");
                Logger.LogException(ex);
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
            }
            catch (ExternalException ex)
            {
                Logger.LogException(ex);
            }
            finally
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                try
                {
                    if (restoreScreenUpdating)
                    {
                        var application = Container.Resolve<Application>();
                        if (application != null)
                        {
                            application.ScreenUpdating = previousScreenUpdating;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogException(ex);
                }

                stopwatch.Stop();
                Logger.Info($"📐 Column auto-fit adjusted in {stopwatch.ElapsedMilliseconds}ms");
            }
        }

        private double GetVisibleColumnWidthLimit(Application application)
        {
            if (application?.ActiveWindow == null) return 0;

            try
            {
                var visibleRange = application.ActiveWindow.VisibleRange;
                if (visibleRange == null) return 0;

                var widthPoints = visibleRange.Width;
                if (widthPoints <= 0) return 0;

                const double PointsPerCharacter = 7.0;
                return widthPoints / PointsPerCharacter;
            }
            catch (COMException)
            {
                return 0;
            }
            catch (InvalidOperationException ex)
            {
                Logger.LogException(ex);
                return 0;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        protected object NormalizeJsonValue(object value)
        {
            if (value is JsonElement element)
            {
                switch (element.ValueKind)
                {
                    case JsonValueKind.String:
                        return element.GetString();
                    case JsonValueKind.Number:
                        if (element.TryGetInt64(out var longVal)) return longVal;
                        if (element.TryGetDouble(out var doubleVal)) return doubleVal;
                        return element.GetRawText();
                    case JsonValueKind.True:
                    case JsonValueKind.False:
                        return element.GetBoolean();
                    case JsonValueKind.Null:
                    case JsonValueKind.Undefined:
                        return null;
                    default:
                        return element.GetRawText();
                }
            }
            return value;
        }

    }
}
