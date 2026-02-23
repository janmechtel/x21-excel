using X21.Common.Data;
using X21.Common.Model;
using X21.Constants;
using X21.Excel.Events;
using X21.Utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using X21.Models;
using System.Linq;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace X21.Excel
{
    public class ExcelSelection : Component
    {
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        private const int WorkbookValidationCacheMs = 250;
        private string _lastWorkbookValidationKey = string.Empty;
        private DateTime _lastWorkbookValidationAt = DateTime.MinValue;
        private bool _lastWorkbookValidationResult = false;

        public event EventHandler<SelectionChangedEventArgs> SelectionChanged;

        public ExcelSelection(Container container)
            : base(container)
        {
            Container = container;
        }

        public override void Init()
        {
            var application = Container.Resolve<Application>();
            application.SheetSelectionChange += OnSheetSelectionChange;
            application.SheetActivate += OnSheetActivate;
        }

        public override void Exit()
        {
            var application = Container.Resolve<Application>();
            application.SheetSelectionChange -= OnSheetSelectionChange;
            application.SheetActivate -= OnSheetActivate;
        }

        private void OnSheetSelectionChange(object sh, Range target)
        {
            if (target == null)
            {
                SelectionChanged?.Invoke(this, new SelectionChangedEventArgs(0, string.Empty));
                SendSelectionToTaskPane(string.Empty);
                return;
            }

            var cellCount = target.Cells.Count;
            var address = string.Empty;

            Execute.Call(
                () =>
                {
                    address = target.Address[false, false];

                    // Simple focus check: if WebView has focus when user clicks Excel cell, restore focus
                    if (WindowUtils.ShouldRestoreFocusToExcel(Container.Resolve<X21.TaskPane.TaskPaneManager>()))
                    {
                        WindowUtils.FocusExcelWindow(Container.Resolve<Application>());
                    }
                },
                Execute.CatchMode.DontLog
            );

            SelectionChanged?.Invoke(this, new SelectionChangedEventArgs(cellCount, address));
            SendSelectionToTaskPane(address);
        }

        private void OnSheetActivate(object sh)
        {
            try
            {
                var app = Container.Resolve<Application>();
                Range selection = null;
                Execute.Call(
                    () =>
                    {
                        selection = app.Selection as Range;
                    },
                    Execute.CatchMode.DontLog
                );

                var address = selection?.Address[false, false] ?? string.Empty;
                var cellCount = selection?.Cells.Count ?? 0;

                SelectionChanged?.Invoke(this, new SelectionChangedEventArgs(cellCount, address));
                SendSelectionToTaskPane(address);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error handling sheet activation: {ex.Message}");
            }
        }

        private void SendSelectionToTaskPane(string address)
        {
            try
            {
                var taskPaneManager = Container.Resolve<X21.TaskPane.TaskPaneManager>();
                var taskPane = taskPaneManager?.GetTaskPane();
                if (taskPane?.Control is X21.TaskPane.WebView2TaskPaneHost webViewHost)
                {
                    webViewHost.SendSelectionChangedEvent(address);
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending selection to task pane: {ex.Message}");
            }
        }


        /// <summary>
        /// Gets the address range of the currently selected cells
        /// </summary>
        /// <returns>Range address string like "A1" or "B2:D5", or empty string if no selection</returns>
        public string GetSelectedRange()
        {
            var app = Container.Resolve<Application>();
            return Execute.Call(
                () =>
                {
                    var selectedRange = app.Selection as Range;
                    if (selectedRange == null) return string.Empty;

                    return selectedRange.Address[false, false];
                },
                Execute.CatchMode.DontLog
            ) ?? string.Empty;
        }

        /// <summary>
        /// Gets the address range of the used range in the active worksheet
        /// </summary>
        /// <returns>Used range address string like "A1:Z100", or empty string if no used range</returns>
        public string GetUsedRange()
        {
            var app = Container.Resolve<Application>();
            return Execute.Call(
                () =>
                {
                    var worksheet = app.ActiveSheet as Worksheet;
                    if (worksheet == null) return string.Empty;

                    var usedRange = worksheet.UsedRange;
                    Logger.Info($"Chat: Used range from Excel Selection: {usedRange.Address[false, false]}");
                    if (usedRange == null) return string.Empty;

                    return usedRange.Address[false, false];
                },
                Execute.CatchMode.DontLog
            ) ?? string.Empty;
        }

        /// <summary>
        /// Gets the cell data for a specific range address
        /// </summary>
        /// <param name="rangeAddress">Range address like "A1", "B2:D5", or "A:A" (entire column)</param>
        /// <returns>List of cells in the specified range</returns>
        public List<Cell> GetRangeData(string rangeAddress)
        {
            if (string.IsNullOrWhiteSpace(rangeAddress))
                return new List<Cell>();

            var app = Container.Resolve<Application>();
            return Execute.Call(
                () =>
                {
                    var worksheet = app.ActiveSheet as Worksheet;
                    if (worksheet == null) return new List<Cell>();

                    try
                    {
                        var targetRange = worksheet.Range[rangeAddress];
                        if (targetRange == null) return new List<Cell>();

                        var cellsData = new List<Cell>();
                        foreach (Range cell in targetRange)
                        {
                            cellsData.Add(new Cell(address: cell.Address[false, false], value: cell.Value2, formula: cell.Formula.ToString()));
                        }

                        return cellsData;
                    }
                    catch
                    {
                        // Return empty list if range is invalid
                        return new List<Cell>();
                    }
                },
                Execute.CatchMode.DontLog
            );
        }

        public (List<Cell> currentCellInfo, List<Cell> sheetUsedCells) GetCurrentSelectionData()
        {
            var selection = Container.Resolve<ExcelSelection>();

            // Get the selected range first (IN → OUT approach)
            string selectedRange = selection.GetSelectedRange();


            // Build context range around selection instead of parsing massive used range
            string contextRange = BuildContextAroundSelection(selectedRange, 5000);

            // Now get the cell data from the focused ranges
            List<Cell> currentCellInfo = !string.IsNullOrEmpty(selectedRange)
                ? selection.GetRangeData(selectedRange)
                : new List<Cell>();


            List<Cell> sheetUsedCells = !string.IsNullOrEmpty(contextRange)
                ? selection.GetRangeData(contextRange)
                : new List<Cell>();


            return (currentCellInfo, sheetUsedCells);
        }

        /// <summary>
        /// Builds a context range around the selected range (IN → OUT approach)
        /// Much more efficient than parsing massive used ranges
        /// </summary>
        private string BuildContextAroundSelection(string selectedRange, int maxCells)
        {
            if (string.IsNullOrEmpty(selectedRange))
            {
                // No selection, return a default small area around A1
                return "A1:J20"; // 200 cells default
            }

            try
            {
                // Parse the selected range
                var (selStartRow, selStartCol, selEndRow, selEndCol) = ParseRangeBounds(selectedRange);
                if (selStartRow == 0 || selStartCol == 0) return "A1:J20";

                // Calculate center of selection
                int centerRow = (selStartRow + selEndRow) / 2;
                int centerCol = (selStartCol + selEndCol) / 2;

                // Calculate expansion size based on maxCells
                // Bias towards more rows (common in Excel data)
                int targetRows = (int)Math.Sqrt(maxCells * 2);
                int targetCols = Math.Max(1, maxCells / targetRows);

                // Ensure we don't exceed maxCells
                while (targetRows * targetCols > maxCells && targetRows > 1)
                {
                    targetRows--;
                }

                // Expand around the center
                int halfRows = targetRows / 2;
                int halfCols = targetCols / 2;

                int contextStartRow = Math.Max(1, centerRow - halfRows);
                int contextEndRow = Math.Min(1048576, centerRow + halfRows); // Excel max rows
                int contextStartCol = Math.Max(1, centerCol - halfCols);
                int contextEndCol = Math.Min(16384, centerCol + halfCols); // Excel max cols

                // Build the context range
                string startCell = ColumnNumberToLetter(contextStartCol) + contextStartRow;
                string endCell = ColumnNumberToLetter(contextEndCol) + contextEndRow;

                string contextRange = $"{startCell}:{endCell}";

                Logger.Info($"Context building: Selection center ({centerRow},{centerCol}) → Context {contextRange} ({targetRows}×{targetCols} = {targetRows * targetCols} cells)");

                return contextRange;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error building context around selection: {ex.Message}");
                return "A1:J20"; // Fallback
            }
        }

        /// <summary>
        /// Focuses and selects a specific cell or range in Excel
        /// </summary>
        /// <param name="cellAddress">Cell address like "A1" or "B2:D5"</param>
        public void FocusCell(string cellAddress)
        {
            var app = Container.Resolve<Application>();
            Execute.Call(
                () =>
                {
                    var worksheet = app.ActiveSheet as Worksheet;
                    if (worksheet != null && !string.IsNullOrWhiteSpace(cellAddress))
                    {
                        // Select the cell/range
                        var targetRange = worksheet.Range[cellAddress];
                        targetRange.Select();

                        // Center the range in the visible window for better UX
                        var window = app.ActiveWindow;
                        if (window != null)
                        {
                            try
                            {
                                // Get the visible rows and columns in the window
                                int visibleRows = window.VisibleRange.Rows.Count;
                                int visibleCols = window.VisibleRange.Columns.Count;

                                // Get the target range position
                                int targetRow = targetRange.Row;
                                int targetCol = targetRange.Column;

                                // Calculate centered scroll position
                                // Position the range in the middle of the visible area
                                int centerRow = Math.Max(1, targetRow - (visibleRows / 2));
                                int centerCol = Math.Max(1, targetCol - (visibleCols / 2));

                                // Apply the scroll to center the range
                                window.ScrollRow = centerRow;
                                window.ScrollColumn = centerCol;
                            }
                            catch (System.Runtime.InteropServices.COMException ex)
                            {
                                // If centering fails (e.g., after sheet switch), fall back to Goto
                                Logger.Info($"Could not center range, using fallback: {ex.Message}");
                                app.Goto(targetRange, false);
                            }
                        }

                        // Activate the Excel window to bring focus
                        app.ActiveWindow.Activate();
                    }
                },
                Execute.CatchMode.DontLog
            );
        }

        /// <summary>
        /// Parses a range address like "A1:C10" into start and end row/column numbers
        /// </summary>
        private (int startRow, int startCol, int endRow, int endCol) ParseRangeBounds(string rangeAddress)
        {
            if (string.IsNullOrEmpty(rangeAddress)) return (0, 0, 0, 0);

            try
            {
                // Handle single cell (like "A1")
                if (!rangeAddress.Contains(":"))
                {
                    var (row, col) = ParseCellAddress(rangeAddress);
                    return (row, col, row, col);
                }

                // Handle range (like "A1:C10")
                var parts = rangeAddress.Split(':');
                if (parts.Length != 2) return (0, 0, 0, 0);

                var (startRow, startCol) = ParseCellAddress(parts[0]);
                var (endRow, endCol) = ParseCellAddress(parts[1]);

                return (startRow, startCol, endRow, endCol);
            }
            catch
            {
                return (0, 0, 0, 0);
            }
        }

        /// <summary>
        /// Parses a cell address like "A1" or "AB123" into row and column numbers
        /// Also handles entire row/column references like "2" (entire row 2), "A" (entire column A)
        /// </summary>
        private (int row, int column) ParseCellAddress(string address)
        {
            if (string.IsNullOrEmpty(address)) return (0, 0);

            try
            {
                // Remove any sheet reference and absolute markers
                var cellRef = address.Contains("!") ? address.Split('!').Last() : address;
                cellRef = cellRef.Replace("$", "");

                // Handle entire row reference (just a number like "2")
                if (int.TryParse(cellRef, out int rowNumber))
                {
                    // For entire row, we'll use a special marker: column = -1 to indicate entire row
                    return (rowNumber, -1);
                }

                // Handle entire column reference (just letters like "A" or "AB")
                if (cellRef.All(char.IsLetter))
                {
                    // Convert column letters to number
                    int columnNumber = 0;
                    for (int i = 0; i < cellRef.Length; i++)
                    {
                        columnNumber = columnNumber * 26 + (cellRef[i] - 'A' + 1);
                    }
                    // For entire column, we'll use a special marker: row = -1 to indicate entire column
                    return (-1, columnNumber);
                }

                var columnPart = "";
                var rowPart = "";

                // Split into column letters and row numbers
                for (int i = 0; i < cellRef.Length; i++)
                {
                    if (char.IsLetter(cellRef[i]))
                    {
                        columnPart += cellRef[i];
                    }
                    else if (char.IsDigit(cellRef[i]))
                    {
                        rowPart = cellRef.Substring(i);
                        break;
                    }
                }

                if (string.IsNullOrEmpty(columnPart) || string.IsNullOrEmpty(rowPart))
                    return (0, 0);

                // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
                int column = 0;
                for (int i = 0; i < columnPart.Length; i++)
                {
                    column = column * 26 + (columnPart[i] - 'A' + 1);
                }

                // Convert row string to number
                if (!int.TryParse(rowPart, out int row))
                    return (0, 0);

                return (row, column);
            }
            catch
            {
                return (0, 0);
            }
        }

        /// <summary>
        /// Converts a column number to Excel column letters (1=A, 26=Z, 27=AA, etc.)
        /// </summary>
        private string ColumnNumberToLetter(int columnNumber)
        {
            string columnName = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        /// <summary>
        /// Gets a limited range around the specified selection, capped at maxCells (default 1000)
        /// </summary>
        /// <param name="worksheet">The worksheet to get data from</param>
        /// <param name="selectedRange">The specific range to center around (if null, uses current selection)</param>
        /// <param name="maxCells">Maximum number of cells to include (default 1000)</param>
        /// <param name="addEmptyCells">Whether to include empty cells in the result</param>
        /// <returns>List of cells in the limited range around selection</returns>
        public List<Cell> GetLimitedRangeAroundSelection(Worksheet worksheet, Range selectedRange = null, int maxCells = 1000, bool addEmptyCells = false)
        {
            if (worksheet == null)
                return new List<Cell>();

            try
            {
                // Get current selection or use provided range
                string selectedRangeAddress;
                if (selectedRange != null)
                {
                    selectedRangeAddress = selectedRange.Address[false, false];
                }
                else
                {
                    selectedRangeAddress = GetSelectedRange();
                    if (string.IsNullOrEmpty(selectedRangeAddress))
                        selectedRangeAddress = "A1"; // Default to A1 if no selection
                }

                var usedRange = worksheet.UsedRange;
                if (usedRange == null)
                    return new List<Cell>();

                // Calculate dimensions for approximately maxCells (prefer square-ish area)
                int maxRows = (int)Math.Ceiling(Math.Sqrt(maxCells));
                int maxCols = Math.Max(1, maxCells / maxRows);

                // Get the selection bounds
                var selectionRange = worksheet.Range[selectedRangeAddress];
                int selectionStartRow = selectionRange.Row;
                int selectionStartCol = selectionRange.Column;

                // Calculate bounds around selection
                int startRow = Math.Max(1, selectionStartRow - maxRows / 2);
                int endRow = Math.Min(usedRange.Row + usedRange.Rows.Count - 1, selectionStartRow + maxRows / 2);
                int startCol = Math.Max(1, selectionStartCol - maxCols / 2);
                int endCol = Math.Min(usedRange.Column + usedRange.Columns.Count - 1, selectionStartCol + maxCols / 2);

                // Ensure we don't exceed maxCells
                int actualRows = endRow - startRow + 1;
                int actualCols = endCol - startCol + 1;
                if (actualRows * actualCols > maxCells)
                {
                    // Reduce dimensions proportionally
                    double scaleFactor = Math.Sqrt((double)maxCells / (actualRows * actualCols));
                    int newRows = Math.Max(1, (int)(actualRows * scaleFactor));
                    int newCols = Math.Max(1, (int)(actualCols * scaleFactor));

                    // Recalculate bounds
                    startRow = Math.Max(1, selectionStartRow - newRows / 2);
                    endRow = Math.Min(usedRange.Row + usedRange.Rows.Count - 1, selectionStartRow + newRows / 2);
                    startCol = Math.Max(1, selectionStartCol - newCols / 2);
                    endCol = Math.Min(usedRange.Column + usedRange.Columns.Count - 1, selectionStartCol + newCols / 2);
                }

                // Create the limited range
                var limitedRange = worksheet.Range[worksheet.Cells[startRow, startCol], worksheet.Cells[endRow, endCol]];

                Logger.Info($"Limited range around selection: {limitedRange.Address[false, false]} ({(endRow - startRow + 1) * (endCol - startCol + 1)} cells)");

                return GetRangeDataFromWorksheet(worksheet, limitedRange, addEmptyCells);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting limited range around selection: {ex.Message}");
                return new List<Cell>();
            }
        }

        public List<Cell> GetRangeDataFromWorksheet(Worksheet worksheet, string rangeAddress)
        {
            var targetRange = worksheet.Range[rangeAddress];

            return GetRangeDataFromWorksheet(worksheet, targetRange);
        }

        public List<Cell> GetRangeDataFromWorksheet(Worksheet worksheet, Range targetRange, bool addEmptyCells = false)
        {
            if (targetRange == null || worksheet == null)
                return new List<Cell>();

            if (addEmptyCells)
            {
                return GetRangeDataFromWorksheetSlow(worksheet, targetRange, addEmptyCells);
            }

            try
            {

                if (targetRange == null)
                {
                    return new List<Cell>();
                }

                var cellsData = new List<Cell>();

                // Use Excel's native SpecialCells to filter non-empty cells efficiently
                try
                {
                    // Get cells with constants (non-formula values that are not empty)
                    var constantCells = targetRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeConstants);
                    foreach (Range cell in constantCells)
                    {
                        // Check if the cell is actually within our target range
                        if (worksheet.Application.Intersect(targetRange, cell) != null)
                        {
                            var cellValue = cell.Value2;
                            if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString().Trim()))
                            {
                                var address = cell.Address[false, false];
                                cellsData.Add(new Cell(
                                    address: address,
                                    value: cellValue,
                                    formula: ""
                                ));
                            }
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Logger.Info("No constant cells found in range");
                }

                try
                {
                    // Get cells with formulas
                    var formulaCells = targetRange.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeFormulas);
                    foreach (Range cell in formulaCells)
                    {
                        // Check if the cell is actually within our target range
                        if (worksheet.Application.Intersect(targetRange, cell) != null)
                        {
                            var cellValue = cell.Value2;
                            if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString().Trim()))
                            {
                                var address = cell.Address[false, false];
                                cellsData.Add(new Cell(
                                    address: address,
                                    value: cellValue,
                                    formula: cell.Formula?.ToString() ?? ""
                                ));
                            }
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Logger.Info("No formula cells found in range");
                }

                return cellsData;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);

                // Fallback to the original slow method
                return GetRangeDataFromWorksheetSlow(worksheet, targetRange);
            }
        }

        private List<Cell> GetRangeDataFromWorksheetSlow(Worksheet worksheet, Range targetRange, bool addEmptyCells = false)
        {

            if (targetRange == null || worksheet == null)
                return new List<Cell>();

            try
            {
                if (targetRange == null)
                {
                    return new List<Cell>();
                }

                var cellsData = new List<Cell>();
                foreach (Range cell in targetRange)
                {
                    // Only add cells with non-null, non-empty values
                    var cellValue = cell.Value2;
                    if (addEmptyCells || (cellValue != null &&
                        !string.IsNullOrEmpty(cellValue.ToString()) &&
                        cellValue.ToString().Trim() != ""))
                    {
                        var address = cell.Address[false, false];
                        cellsData.Add(new Cell(
                            address: address,
                            value: cellValue,
                            formula: cell.Formula?.ToString() ?? ""
                        ));
                    }
                }

                return cellsData;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new List<Cell>();
            }
        }

        private Workbook ResolveWorkbookForContext(out string resolutionPath)
        {
            resolutionPath = WorkbookResolutionPaths.ActiveWorkbook;
            Workbook workbook = null;

            try
            {
                workbook = Container.Resolve<Workbook>();
                if (workbook != null)
                {
                    var workbookName = workbook.Name; // Validate COM object
                    if (string.IsNullOrWhiteSpace(workbookName))
                    {
                        throw new COMException("Workbook name is invalid");
                    }
                    var excelApplication = Container.Resolve<Application>();
                    var workbookKey = !string.IsNullOrEmpty(workbook.FullName)
                        ? workbook.FullName
                        : workbookName;
                    var now = DateTime.UtcNow;
                    if (_lastWorkbookValidationKey == workbookKey &&
                        (now - _lastWorkbookValidationAt).TotalMilliseconds < WorkbookValidationCacheMs &&
                        _lastWorkbookValidationResult)
                    {
                        resolutionPath = WorkbookResolutionPaths.ContainerWorkbook;
                        return workbook;
                    }
                    var isOpen = false;
                    foreach (Workbook wb in excelApplication.Workbooks)
                    {
                        if (wb == workbook ||
                            wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            (!string.IsNullOrEmpty(workbook.FullName) &&
                             wb.FullName.Equals(workbook.FullName, StringComparison.OrdinalIgnoreCase)))
                        {
                            isOpen = true;
                            break;
                        }
                    }
                    _lastWorkbookValidationKey = workbookKey;
                    _lastWorkbookValidationAt = now;
                    _lastWorkbookValidationResult = isOpen;

                    if (isOpen)
                    {
                        resolutionPath = WorkbookResolutionPaths.ContainerWorkbook;
                        return workbook;
                    }

                    Logger.Info("Container workbook is not open; falling back to active workbook resolution");
                }
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException ||
                    ex is System.Threading.ThreadAbortException ||
                    ex is StackOverflowException ||
                    ex is System.Security.SecurityException)
                {
                    throw;
                }
                Logger.Info($"Container workbook resolution failed, falling back: {ex.Message}");
            }

            var application = Container.Resolve<Application>();
            workbook = application.ActiveWorkbook;
            if (workbook != null)
            {
                resolutionPath = WorkbookResolutionPaths.ActiveWorkbook;
                return workbook;
            }

            var activeSheet = application.ActiveWindow?.ActiveSheet as Worksheet;
            workbook = activeSheet?.Parent as Workbook;
            if (workbook != null)
            {
                resolutionPath = WorkbookResolutionPaths.ActiveWindow;
                return workbook;
            }

            var openCount = application.Workbooks?.Count ?? 0;
            if (openCount == 1)
            {
                resolutionPath = WorkbookResolutionPaths.SingleOpenWorkbook;
                return application.Workbooks[1];
            }

            resolutionPath = WorkbookResolutionPaths.None;
            return null;
        }

        public string GetActiveSheetName()
        {
            try
            {
                var workbook = ResolveWorkbookForContext(out _);
                if (workbook?.ActiveSheet is Worksheet sheet)
                {
                    return sheet.Name;
                }
            }
            catch (COMException ex)
            {
                Logger.Info($"Error getting active sheet name: {ex.Message}");
            }

            return string.Empty;
        }

        public List<string> GetAllSheetNames()
        {
            try
            {
                var workbook = ResolveWorkbookForContext(out _);
                if (workbook?.Sheets != null)
                {
                    return workbook.Sheets.Cast<Worksheet>()
                        .Select(sheet => sheet.Name)
                        .ToList();
                }
            }
            catch (Exception ex) when (ex is COMException || ex is InvalidCastException)
            {
                Logger.Info($"Error getting sheet names: {ex.Message}");
            }

            return new List<string>();
        }

        public string GetWorkbookName()
        {
            try
            {
                var workbook = ResolveWorkbookForContext(out var resolutionPath);
                if (workbook == null)
                {
                    Logger.Info("[WorkbookResolver] No active workbook found; returning empty.");
                    return string.Empty;
                }

                Logger.Info($"[WorkbookResolver] Using workbook via {resolutionPath}: '{workbook.Name}'");
                return workbook.Name;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting workbook name: {ex.Message}");
                return string.Empty;
            }
        }

        public string GetLanguageCode()
        {
            try
            {
                var application = Container.Resolve<Application>();
                // Get the UI language ID (what the user sees in Excel)
                int languageId = (int)application.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI];

                // Convert LCID to culture name (e.g., 1033 -> "en-US", 1031 -> "de-DE")
                var culture = new System.Globalization.CultureInfo(languageId);
                return culture.Name;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting language code: {ex.Message}");
                return "en-US"; // Default to English
            }
        }

        /// <summary>
        /// Gets the date language for Excel date formats
        /// Detects whether Excel uses English ("dd", "yyyy") or German ("tt", "jjjj") date format tokens
        /// </summary>
        /// <returns>Date language ("english" or "german")</returns>
        public string GetDateLanguage()
        {
            try
            {
                var application = Container.Resolve<Application>();

                try
                {
                    // If this evaluates without throwing, "tt" is supported
                    application.Evaluate("TEXT(TODAY(),\"tt\")");

                    Logger.Info("Supports German: true");
                    return "german";
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    Logger.Info($"Supports German: false (probe failed with COM error: {ex.Message})");
                    return "english";
                }
                catch (InvalidOperationException ex)
                {
                    Logger.Info($"Supports German: false (probe failed with invalid operation: {ex.Message})");
                    return "english";
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Logger.Info($"Error getting date language (COM): {ex.Message}");
                return "english"; // safe default
            }
            catch (InvalidOperationException ex)
            {
                Logger.Info($"Error getting date language (invalid operation): {ex.Message}");
                return "english"; // safe default
            }
        }

        /// <summary>
        /// Gets the list separator used in Excel formulas
        /// Controls formula syntax (; or ,)
        /// </summary>
        /// <returns>List separator character (e.g., "," or ";")</returns>
        public string GetListSeparator()
        {
            try
            {
                var application = Container.Resolve<Application>();
                string listSeparator = application.International[Microsoft.Office.Interop.Excel.XlApplicationInternational.xlListSeparator].ToString();
                return listSeparator;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Logger.Info($"Error getting list separator from Excel (COM): {ex.Message}");
                return ","; // Default to comma for Excel/COM issues
            }
            catch (Exception ex)
            {
                Logger.Info($"Unexpected error getting list separator: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Gets the decimal separator used in Excel
        /// Important for numeric literals (3,14 vs 3.14)
        /// </summary>
        /// <returns>Decimal separator character (e.g., "." or ",")</returns>
        public string GetDecimalSeparator()
        {
            try
            {
                var application = Container.Resolve<Application>();
                string decimalSeparator = application.International[Microsoft.Office.Interop.Excel.XlApplicationInternational.xlDecimalSeparator].ToString();
                return decimalSeparator;
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Logger.Info($"Error getting decimal separator from Excel (COM): {ex.Message}");
                return "."; // Default to period
            }
        }

        /// <summary>
        /// Gets the thousands (group) separator used in Excel.
        /// Useful for displaying localized numeric literals in chat (1 234,56 / 1.234,56 / 1,234.56).
        /// </summary>
        /// <returns>Thousands separator (e.g., ",", ".", " ")</returns>
        public string GetThousandsSeparator()
        {
            try
            {
                var application = Container.Resolve<Application>();
                string thousandsSeparator = application.International[
                    Microsoft.Office.Interop.Excel.XlApplicationInternational.xlThousandsSeparator
                ].ToString();
                return thousandsSeparator;
            }
            catch (Exception ex) when (
                ex is System.Runtime.InteropServices.COMException ||
                ex is System.Runtime.InteropServices.InvalidComObjectException ||
                ex is NullReferenceException ||
                ex is InvalidCastException
            )
            {
                // Avoid hard-defaulting to ',' because many locales use '.' or ' '.
                // Prefer a best-effort fallback based on the current culture, then infer from decimal separator.
                Logger.Info($"Warning: Error getting thousands separator from Excel (interop): {ex.GetType().Name}: {ex.Message}. Falling back to culture/inference.");

                var cultureSeparator = CultureInfo.CurrentCulture?.NumberFormat?.NumberGroupSeparator;
                if (!string.IsNullOrEmpty(cultureSeparator))
                {
                    return cultureSeparator;
                }

                // Last resort inference: choose a common group separator opposite the decimal separator.
                string decimalSeparator;
                try
                {
                    decimalSeparator = GetDecimalSeparator();
                }
                catch (Exception dex) when (
                    dex is System.Runtime.InteropServices.COMException ||
                    dex is System.Runtime.InteropServices.InvalidComObjectException ||
                    dex is NullReferenceException ||
                    dex is InvalidCastException
                )
                {
                    decimalSeparator = CultureInfo.CurrentCulture?.NumberFormat?.NumberDecimalSeparator ?? ".";
                }
                if (decimalSeparator == ",")
                {
                    return "."; // Common for many comma-decimal locales
                }

                if (decimalSeparator == ".")
                {
                    return ","; // Common for many dot-decimal locales
                }

                return ","; // Final fallback
            }
        }

        /// <summary>
        /// Gets the full path of the active workbook (directory + file name).
        /// </summary>
        /// <returns>Full workbook path like "D:\Folder\Book1.xlsx", or empty string if unavailable.</returns>
        public string GetWorkbookPath()
        {
            try
            {
                var workbook = ResolveWorkbookForContext(out _);
                if (workbook == null || string.IsNullOrEmpty(workbook.FullName))
                {
                    return string.Empty;
                }

                return workbook.FullName;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting workbook path: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Ensures a stable X21 file ID exists in the workbook's custom document properties
        /// and returns it. Creates one if missing.
        /// </summary>
        public string EnsureAndGetWorkbookFileId()
        {
            try
            {
                var workbook = ResolveWorkbookForContext(out _);
                if (workbook == null) return string.Empty;

                const string propName = "X21FileId";
                Office.DocumentProperties props = (Office.DocumentProperties)workbook.CustomDocumentProperties;

                // Try read existing
                try
                {
                    var existing = props[propName];
                    if (existing != null && existing.Value != null)
                    {
                        var valueStr = existing.Value.ToString();
                        if (!string.IsNullOrWhiteSpace(valueStr)) return valueStr;
                    }
                }
                catch
                {
                    // Property not found - fall through to create
                }

                // Create new GUID and save
                var newId = Guid.NewGuid().ToString();
                props.Add(propName, false, Office.MsoDocProperties.msoPropertyTypeString, newId, Type.Missing);
                return newId;
            }
            catch (Exception ex)
            {
                Logger.Info($"Error ensuring workbook file id: {ex.Message}");
                return string.Empty;
            }
        }

        public Worksheet GetWorksheet(string sheetName, string workbookName)
        {
            try
            {
                var application = Container.Resolve<Application>();

                Workbook targetWorkbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    // Use active workbook if no specific workbook name provided
                    targetWorkbook = application.ActiveWorkbook;
                }
                else
                {
                    // Find specific workbook by name
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWorkbook = wb;
                            Logger.Info($"📋 Found workbook: '{wb.Name}' (requested: '{workbookName}')");
                            break;
                        }
                    }

                    if (targetWorkbook == null)
                    {
                        Logger.Info($"❌ Workbook '{workbookName}' not found");
                        return null;
                    }
                }

                if (targetWorkbook?.Worksheets != null)
                {
                    foreach (Worksheet ws in targetWorkbook.Worksheets)
                    {
                        if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            Logger.Info($"📋 Found worksheet: '{ws.Name}' in workbook: '{targetWorkbook.Name}' (requested: '{sheetName}')");
                            return ws;
                        }
                    }
                }

                Logger.Info($"❌ Worksheet '{sheetName}' not found in workbook '{targetWorkbook?.Name ?? workbookName}'");
                return null;
            }
            catch (Exception ex)
            {
                Logger.Info($"❌ Error accessing worksheet '{sheetName}' in workbook '{workbookName}': {ex.Message}");
                return null;
            }
        }

    }
}
