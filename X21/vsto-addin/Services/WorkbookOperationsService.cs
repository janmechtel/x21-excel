using System;
using System.Collections.Generic;
using System.Net;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using X21.Common.Data;
using X21.Common.Model;

namespace X21.Services
{
    /// <summary>
    /// Service for workbook-level operations like opening, closing, and copying sheets
    /// </summary>
    public class WorkbookOperationsService : Component
    {
        private readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        };

        public WorkbookOperationsService(Container container) : base(container)
        {
        }

        private void ReleaseComObjectSafe(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (COMException)
            {
                // Ignore COM release failures to avoid masking primary errors
            }
            catch (InvalidComObjectException)
            {
                // Ignore invalid COM object failures during release
            }
        }

        private string SanitizeForLog(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace("\r", string.Empty)
                        .Replace("\n", string.Empty);
        }

        private static bool IsCriticalException(Exception ex)
        {
            return ex is OutOfMemoryException
                || ex is StackOverflowException
                || ex is ThreadAbortException
                || ex is AccessViolationException;
        }

        public async Task HandleOpenWorkbook(string requestBody, HttpListenerResponse response)
        {
            try
            {
                Logger.Info("HandleOpenWorkbook: Received request");
                Logger.Info($"HandleOpenWorkbook: Request body: {SanitizeForLog(requestBody)}");

                var openWorkbookRequest = JsonSerializer.Deserialize<OpenWorkbookRequest>(requestBody, _jsonOptions);

                if (openWorkbookRequest == null)
                {
                    Logger.Info("HandleOpenWorkbook: ⚠️ Request body deserialization resulted in null");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                Logger.Info($"HandleOpenWorkbook: FilePath: {openWorkbookRequest.FilePath}, Visible: {openWorkbookRequest.Visible}");

                if (string.IsNullOrEmpty(openWorkbookRequest.FilePath))
                {
                    Logger.Info("HandleOpenWorkbook: ⚠️ FilePath is empty");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "filePath is required");
                    return;
                }

                Logger.Info($"HandleOpenWorkbook: Calling InvokeExcel to open workbook");
                var result = InvokeExcel(() => OpenExcelWorkbook(openWorkbookRequest.FilePath, openWorkbookRequest.Visible));

                Logger.Info($"HandleOpenWorkbook: InvokeExcel completed, success: {result.Success}, workbook: {result.WorkbookName}");
                await SendJsonResponse(response, result);

                Logger.Info($"Successfully opened workbook: {openWorkbookRequest.FilePath}");
            }
            catch (JsonException jsonEx)
            {
                Logger.LogException(jsonEx);
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "Invalid request body", 400);
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, "An unexpected error occurred while opening the workbook.", 500);
            }
        }

        public async Task HandleCopySheets(string requestBody, HttpListenerResponse response)
        {
            try
            {
                Logger.Info("HandleCopySheets: Received request");
                Logger.Info($"HandleCopySheets: Request body: {SanitizeForLog(requestBody)}");

                var copySheetsRequest = JsonSerializer.Deserialize<CopySheetsRequest>(requestBody, _jsonOptions);

                if (copySheetsRequest == null)
                {
                    Logger.Info("HandleCopySheets: ⚠️ Request body deserialization resulted in null");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                Logger.Info($"HandleCopySheets: Deserialized request - Source: {copySheetsRequest.SourceWorkbookName}, Target: {copySheetsRequest.TargetWorkbookName}, Sheets: {string.Join(", ", copySheetsRequest.SheetNames ?? new string[0])}");

                if (string.IsNullOrEmpty(copySheetsRequest.SourceWorkbookName) ||
                    string.IsNullOrEmpty(copySheetsRequest.TargetWorkbookName) ||
                    copySheetsRequest.SheetNames == null || copySheetsRequest.SheetNames.Length == 0)
                {
                    Logger.Info("HandleCopySheets: ⚠️ Missing required parameters");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "sourceWorkbookName, targetWorkbookName, and sheetNames are required");
                    return;
                }

                Logger.Info($"HandleCopySheets: Calling InvokeExcel to copy {copySheetsRequest.SheetNames.Length} sheets");
                var result = InvokeExcel(() => CopyExcelSheets(
                    copySheetsRequest.SourceWorkbookName,
                    copySheetsRequest.TargetWorkbookName,
                    copySheetsRequest.SheetNames,
                    copySheetsRequest.NamePrefix));

                Logger.Info($"HandleCopySheets: InvokeExcel completed, success: {result.Success}");
                await SendJsonResponse(response, result);

                Logger.Info($"Successfully copied {copySheetsRequest.SheetNames.Length} sheets from {copySheetsRequest.SourceWorkbookName} to {copySheetsRequest.TargetWorkbookName}");
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "Invalid request body", 400);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, ex.Message, 400);
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, "An unexpected error occurred while copying sheets.", 500);
            }
        }

        public async Task HandleCloseWorkbook(string requestBody, HttpListenerResponse response)
        {
            try
            {
                Logger.Info("HandleCloseWorkbook: Received request");
                Logger.Info($"HandleCloseWorkbook: Request body: {SanitizeForLog(requestBody)}");

                var closeWorkbookRequest = JsonSerializer.Deserialize<CloseWorkbookRequest>(requestBody, _jsonOptions);

                if (closeWorkbookRequest == null)
                {
                    Logger.Info("HandleCloseWorkbook: ⚠️ Request body deserialization resulted in null");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                Logger.Info($"HandleCloseWorkbook: WorkbookName: {closeWorkbookRequest.WorkbookName}, SaveChanges: {closeWorkbookRequest.SaveChanges}");

                if (string.IsNullOrEmpty(closeWorkbookRequest.WorkbookName))
                {
                    Logger.Info("HandleCloseWorkbook: ⚠️ WorkbookName is empty");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "workbookName is required");
                    return;
                }

                Logger.Info($"HandleCloseWorkbook: Calling InvokeExcel to close workbook");
                var result = InvokeExcel(() => CloseExcelWorkbook(closeWorkbookRequest.WorkbookName, closeWorkbookRequest.SaveChanges));

                Logger.Info($"HandleCloseWorkbook: InvokeExcel completed, success: {result.Success}");
                await SendJsonResponse(response, result);

                Logger.Info($"Successfully closed workbook: {closeWorkbookRequest.WorkbookName}");
            }
            catch (JsonException ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "Invalid request body", 400);
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, $"Invalid request data: {ex.Message}", 400);
            }
            catch (COMException ex) when (!IsCriticalException(ex))
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, "An unexpected error occurred while closing the workbook.", 500);
            }
        }

        private OpenWorkbookResponse OpenExcelWorkbook(string filePath, bool visible)
        {
            try
            {
                Logger.Info($"📂 OpenExcelWorkbook: Entry - Opening workbook: {filePath}, visible: {visible}");

                Logger.Info("OpenExcelWorkbook: Resolving Excel Application from container");
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Logger.Info($"OpenExcelWorkbook: Excel Application resolved, open workbooks count: {application.Workbooks.Count}");

                // Check if already open
                Logger.Info($"OpenExcelWorkbook: Checking if workbook is already open");
                for (int i = 1; i <= application.Workbooks.Count; i++)
                {
                    Workbook wb = null;
                    try
                    {
                        wb = application.Workbooks[i];
                        Logger.Info($"OpenExcelWorkbook: Checking workbook - Name: '{wb.Name}', FullName: '{wb.FullName}'");
                        if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                        {
                            var existingName = wb.Name;
                            var existingFullName = wb.FullName;
                            Logger.Info($"OpenExcelWorkbook: ✅ Workbook already open: {existingName}");
                            return new OpenWorkbookResponse
                            {
                                Success = true,
                                Message = $"Workbook already open: {existingName}",
                                WorkbookName = existingName,
                                WorkbookFullName = existingFullName,
                                AlreadyOpen = true
                            };
                        }
                    }
                    finally
                    {
                        ReleaseComObjectSafe(wb);
                    }
                }

                Logger.Info("OpenExcelWorkbook: Workbook not currently open, proceeding to open");

                // Disable screen updating when opening hidden workbooks
                bool originalScreenUpdating = application.ScreenUpdating;
                if (!visible)
                {
                    Logger.Info($"OpenExcelWorkbook: Disabling screen updating (was: {originalScreenUpdating})");
                    application.ScreenUpdating = false;
                }

                try
                {
                    // Open workbook
                    Logger.Info($"OpenExcelWorkbook: Calling Excel COM Workbooks.Open for: {filePath}");
                    Workbook workbook = null;
                    Windows windows = null;
                    try
                    {
                        workbook = application.Workbooks.Open(
                        Filename: filePath,
                        UpdateLinks: 0,
                        ReadOnly: false,
                        Format: Type.Missing,
                        Password: Type.Missing,
                        WriteResPassword: Type.Missing,
                        IgnoreReadOnlyRecommended: true,
                        Origin: Type.Missing,
                        Delimiter: Type.Missing,
                        Editable: Type.Missing,
                        Notify: Type.Missing,
                        Converter: Type.Missing,
                        AddToMru: Type.Missing,
                        Local: Type.Missing,
                        CorruptLoad: Type.Missing
                    );

                        Logger.Info($"OpenExcelWorkbook: Workbooks.Open completed, workbook name: {workbook.Name}");

                        // Hide the window if not visible
                        if (!visible)
                        {
                            Logger.Info($"OpenExcelWorkbook: Hiding workbook windows (count: {workbook.Windows.Count})");
                            windows = workbook.Windows;
                            var windowCount = windows.Count;
                            for (int i = 1; i <= windowCount; i++)
                            {
                                Window window = null;
                                try
                                {
                                    window = windows[i];
                                    window.Visible = false;
                                }
                                finally
                                {
                                    ReleaseComObjectSafe(window);
                                }
                            }
                            Logger.Info("OpenExcelWorkbook: All windows hidden");
                        }

                        Logger.Info($"✅ OpenExcelWorkbook: Successfully opened workbook: {workbook.Name}");

                        return new OpenWorkbookResponse
                        {
                            Success = true,
                            Message = $"Successfully opened workbook: {workbook.Name}",
                            WorkbookName = workbook.Name,
                            WorkbookFullName = workbook.FullName,
                            AlreadyOpen = false
                        };
                    }
                    finally
                    {
                        ReleaseComObjectSafe(windows);
                        ReleaseComObjectSafe(workbook);
                    }
                }
                finally
                {
                    // Restore screen updating
                    if (!visible)
                    {
                        application.ScreenUpdating = originalScreenUpdating;
                    }
                }
            }
            catch (Exception ex) when (!IsCriticalException(ex))
            {
                Logger.LogException(ex);
                return new OpenWorkbookResponse
                {
                    Success = false,
                    Message = $"Failed to open workbook: {ex.Message}",
                    WorkbookName = null,
                    WorkbookFullName = null,
                    AlreadyOpen = false
                };
            }
        }

        private CopySheetsResponse CopyExcelSheets(string sourceWorkbookName, string targetWorkbookName, string[] sheetNames, string namePrefix)
        {
            Workbook sourceWorkbook = null;
            Workbook targetWorkbook = null;
            Worksheet lastSheet = null;
            var sourceSheetsList = new List<Worksheet>();
            var conflictingSheets = new Dictionary<string, (Worksheet sheet, string originalName, string tempName)>();
            Sheets targetSheets = null;
            try
            {
                Logger.Info($"📋 CopyExcelSheets: Entry - Copying {sheetNames.Length} sheets from '{sourceWorkbookName}' to '{targetWorkbookName}'");

                Logger.Info("CopyExcelSheets: Resolving Excel Application from container");
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Logger.Info($"CopyExcelSheets: Excel Application resolved, open workbooks count: {application.Workbooks.Count}");

                // Disable screen updating for performance and to prevent flashing
                bool originalScreenUpdating = application.ScreenUpdating;
                bool originalDisplayAlerts = application.DisplayAlerts;
                Logger.Info($"CopyExcelSheets: Disabling screen updating (was: {originalScreenUpdating}) and alerts (was: {originalDisplayAlerts})");
                application.ScreenUpdating = false;
                application.DisplayAlerts = false;

                try
                {
                    // Find source workbook
                    Logger.Info($"CopyExcelSheets: Searching for source workbook '{sourceWorkbookName}'");
                    for (int i = 1; i <= application.Workbooks.Count; i++)
                    {
                        Workbook wb = null;
                        try
                        {
                            wb = application.Workbooks[i];
                            Logger.Info($"CopyExcelSheets: Checking workbook - Name: '{wb.Name}', FullName: '{wb.FullName}'");
                            if (wb.Name.Equals(sourceWorkbookName, StringComparison.OrdinalIgnoreCase) ||
                                wb.FullName.Equals(sourceWorkbookName, StringComparison.OrdinalIgnoreCase))
                            {
                                sourceWorkbook = wb;
                                Logger.Info($"CopyExcelSheets: ✅ Found source workbook: '{wb.Name}'");
                                wb = null; // keep reference alive
                                break;
                            }
                        }
                        finally
                        {
                            ReleaseComObjectSafe(wb);
                        }
                    }

                    if (sourceWorkbook == null)
                    {
                        Logger.Info($"CopyExcelSheets: ❌ Source workbook '{sourceWorkbookName}' not found");
                        throw new Exception($"Source workbook '{sourceWorkbookName}' not found");
                    }

                    // Find target workbook
                    Logger.Info($"CopyExcelSheets: Searching for target workbook '{targetWorkbookName}'");
                    for (int i = 1; i <= application.Workbooks.Count; i++)
                    {
                        Workbook wb = null;
                        try
                        {
                            wb = application.Workbooks[i];
                            Logger.Info($"CopyExcelSheets: Checking workbook - Name: '{wb.Name}', FullName: '{wb.FullName}'");
                            if (wb.Name.Equals(targetWorkbookName, StringComparison.OrdinalIgnoreCase) ||
                                wb.FullName.Equals(targetWorkbookName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetWorkbook = wb;
                                Logger.Info($"CopyExcelSheets: ✅ Found target workbook: '{wb.Name}'");
                                wb = null; // keep reference alive
                                break;
                            }
                        }
                        finally
                        {
                            ReleaseComObjectSafe(wb);
                        }
                    }

                    if (targetWorkbook == null)
                    {
                        Logger.Info($"CopyExcelSheets: ❌ Target workbook '{targetWorkbookName}' not found");
                        throw new Exception($"Target workbook '{targetWorkbookName}' not found");
                    }

                    Logger.Info($"CopyExcelSheets: Target workbook has {targetWorkbook.Worksheets.Count} worksheets");
                    targetSheets = targetWorkbook.Worksheets;
                    lastSheet = targetSheets[targetSheets.Count];
                    Logger.Info($"CopyExcelSheets: Last sheet in target: '{lastSheet.Name}'");

                    // PHASE 1: Collect existing sheet names from target
                    Logger.Info("CopyExcelSheets: PHASE 1 - Collecting existing sheet names from target");
                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    var targetSheetCount = targetSheets.Count;
                    for (int i = 1; i <= targetSheetCount; i++)
                    {
                        Worksheet ws = null;
                        try
                        {
                            ws = targetSheets[i];
                            existingNames.Add(ws.Name);
                        }
                        finally
                        {
                            ReleaseComObjectSafe(ws);
                        }
                    }
                    Logger.Info($"CopyExcelSheets: Found {existingNames.Count} existing sheets in target: {string.Join(", ", existingNames)}");

                    // PHASE 2: Build mapping of original names to new names
                    Logger.Info("CopyExcelSheets: PHASE 2 - Building name mapping for all sheets to be copied");
                    var nameMapping = new Dictionary<string, string>();

                    foreach (string sheetName in sheetNames)
                    {
                        string finalName = sheetName;

                        // If name exists, prefer filename-based prefix; fall back to underscore behavior.
                        if (existingNames.Contains(sheetName))
                        {
                            if (!string.IsNullOrWhiteSpace(namePrefix))
                            {
                                finalName = $"{namePrefix}{sheetName}";
                                int suffix = 1;
                                while (existingNames.Contains(finalName))
                                {
                                    finalName = $"{namePrefix}{sheetName}_{suffix}";
                                    suffix++;
                                }
                            }
                            else
                            {
                                finalName = $"_{sheetName}";

                                // Keep adding underscores until we find a unique name
                                while (existingNames.Contains(finalName))
                                {
                                    finalName = $"_{finalName}";
                                }
                            }
                        }

                        nameMapping[sheetName] = finalName;
                        existingNames.Add(finalName); // Add to set so next sheets see this as taken
                    }

                    // Log the complete mapping
                    Logger.Info("CopyExcelSheets: Sheet name mapping:");
                    foreach (var mapping in nameMapping)
                    {
                        if (mapping.Key == mapping.Value)
                        {
                            Logger.Info($"  '{mapping.Key}' → '{mapping.Value}' (no conflict)");
                        }
                        else
                        {
                            Logger.Info($"  '{mapping.Key}' → '{mapping.Value}' (conflict resolved)");
                        }
                    }

                    // PHASE 3: Copy ALL sheets at once to preserve inter-sheet references/formulas
                    Logger.Info($"CopyExcelSheets: PHASE 3 - Copying all {sheetNames.Length} sheets in one operation to preserve formulas");
                    var copiedSheets = new List<string>();

                    // Find all source sheets and store conflicting sheets info
                    foreach (string sheetName in sheetNames)
                    {
                        Worksheet sourceSheet = null;
                        Sheets sourceSheets = null;
                        try
                        {
                            sourceSheets = sourceWorkbook.Worksheets;
                            var sourceCount = sourceSheets.Count;
                            for (int i = 1; i <= sourceCount; i++)
                            {
                                Worksheet ws = null;
                                try
                                {
                                    ws = sourceSheets[i];
                                    if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        sourceSheet = ws;
                                        ws = null; // keep alive beyond loop
                                        break;
                                    }
                                }
                                finally
                                {
                                    ReleaseComObjectSafe(ws);
                                }
                            }
                        }
                        finally
                        {
                            ReleaseComObjectSafe(sourceSheets);
                        }

                        if (sourceSheet == null)
                        {
                            Logger.Info($"CopyExcelSheets: ⚠️ Sheet '{sheetName}' not found in source workbook, skipping");
                            continue;
                        }

                        sourceSheetsList.Add(sourceSheet);

                        // Check if this sheet has a conflict
                        string mappedName = nameMapping[sheetName];
                        bool hasConflict = !sheetName.Equals(mappedName, StringComparison.OrdinalIgnoreCase);

                        if (hasConflict)
                        {
                            for (int i = 1; i <= targetSheetCount; i++)
                            {
                                Worksheet ws = null;
                                try
                                {
                                    ws = targetSheets[i];
                                    if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        string tempName = $"TEMP_{Guid.NewGuid().ToString().Substring(0, 8)}";
                                        Logger.Info($"CopyExcelSheets: Temporarily renaming conflicting sheet '{ws.Name}' → '{tempName}'");
                                        ws.Name = tempName;
                                        conflictingSheets[sheetName] = (ws, sheetName, tempName);
                                        ws = null; // keep reference until restoration
                                        break;
                                    }
                                }
                                finally
                                {
                                    ReleaseComObjectSafe(ws);
                                }
                            }
                        }
                    }

                    if (sourceSheetsList.Count == 0)
                    {
                        Logger.Info($"CopyExcelSheets: No sheets to copy");
                        return new CopySheetsResponse
                        {
                            Success = true,
                            Message = "No sheets to copy",
                            SourceWorkbookName = sourceWorkbookName,
                            TargetWorkbookName = targetWorkbookName,
                            CopiedSheets = new string[0]
                        };
                    }

                    try
                    {
                        // Create array of sheets to copy
                        Logger.Info($"CopyExcelSheets: Copying {sourceSheetsList.Count} sheets as a group to preserve formulas");

                        // Copy all sheets at once - this preserves inter-sheet references
                        if (sourceSheetsList.Count == 1)
                        {
                            // Single sheet copy
                            sourceSheetsList[0].Copy(After: lastSheet);
                        }
                        else
                        {
                            // Multiple sheets - select them all and copy together
                            // This preserves formulas that reference between sheets
                            Logger.Info($"CopyExcelSheets: Selecting all {sourceSheetsList.Count} sheets for group copy");

                            // Create array of sheet objects for Select method
                            var sheetsToSelect = new object[sourceSheetsList.Count];
                            for (int i = 0; i < sourceSheetsList.Count; i++)
                            {
                                sheetsToSelect[i] = sourceSheetsList[i];
                            }

                            // Select all sheets (this groups them)
                            sourceSheetsList[0].Select(Type.Missing);
                            for (int i = 1; i < sourceSheetsList.Count; i++)
                            {
                                sourceSheetsList[i].Select(false); // false = add to selection
                            }

                            // Copy the selected group to target workbook
                            Logger.Info($"CopyExcelSheets: Copying selected group after '{lastSheet.Name}'");
                            application.ActiveWindow.SelectedSheets.Copy(After: lastSheet);
                        }

                        Logger.Info($"CopyExcelSheets: Bulk copy operation completed");
                    }
                    catch (Exception ex)
                    {
                        Logger.Info($"CopyExcelSheets: Bulk copy failed: {ex.Message}, falling back to individual copy");

                        // Restore conflicting sheet names on error
                        foreach (var conflict in conflictingSheets.Values)
                        {
                            Logger.Info($"CopyExcelSheets: Restoring conflicting sheet name to '{conflict.originalName}'");
                            conflict.sheet.Name = conflict.originalName;
                        }
                        throw;
                    }

                    // Now rename all the copied sheets according to the mapping
                    // The copied sheets will be at the end, after lastSheet
                    Logger.Info($"CopyExcelSheets: Renaming copied sheets according to mapping");
                    var currentTargetSheetCount = targetSheets.Count;

                    // Process sheets that need renaming (have conflicts)
                    foreach (string originalName in sheetNames.Where(name => nameMapping.TryGetValue(name, out _)))
                    {
                        var mappedName = nameMapping[originalName];
                        bool needsRename = !originalName.Equals(mappedName, StringComparison.OrdinalIgnoreCase);

                        if (needsRename)
                        {
                            // Find the copied sheet by its original name
                            Worksheet copiedSheet = null;
                            // Find the copied sheet by its original name
                            for (int i = 1; i <= currentTargetSheetCount; i++)
                            {
                                Worksheet ws = null;
                                try
                                {
                                    ws = targetSheets[i];
                                    // The copied sheet still has its original name and is not in the conflictingSheets dict
                                    if (ws.Name.Equals(originalName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        bool isOriginalConflicting = conflictingSheets.Values.Any(conflict => conflict.sheet == ws);

                                        if (!isOriginalConflicting)
                                        {
                                            copiedSheet = ws;
                                            ws = null; // keep reference
                                            break;
                                        }
                                    }
                                }
                                finally
                                {
                                    ReleaseComObjectSafe(ws);
                                }
                            }

                            if (copiedSheet != null)
                            {
                                Logger.Info($"CopyExcelSheets: Renaming copied sheet '{copiedSheet.Name}' → '{mappedName}'");
                                copiedSheet.Name = mappedName;
                                copiedSheets.Add(mappedName);
                            }
                            else
                            {
                                Logger.Info($"CopyExcelSheets: ⚠️ Could not find copied sheet '{originalName}' to rename");
                            }
                        }
                        else
                        {
                            // No rename needed, just add to list
                            copiedSheets.Add(mappedName);
                            Logger.Info($"CopyExcelSheets: Sheet '{originalName}' kept its original name");
                        }
                    }

                    // Restore all conflicting sheet names
                    foreach (var conflict in conflictingSheets.Values)
                    {
                        Logger.Info($"CopyExcelSheets: Restoring original sheet name to '{conflict.originalName}'");
                        conflict.sheet.Name = conflict.originalName;
                    }

                    Logger.Info($"✅ CopyExcelSheets: Successfully copied and renamed {copiedSheets.Count} sheets with formulas preserved");

                    Logger.Info($"✅ Successfully copied {copiedSheets.Count} sheets");

                    return new CopySheetsResponse
                    {
                        Success = true,
                        Message = $"Successfully copied {copiedSheets.Count} sheets",
                        SourceWorkbookName = sourceWorkbookName,
                        TargetWorkbookName = targetWorkbookName,
                        CopiedSheets = copiedSheets.ToArray()
                    };
                }
                finally
                {
                    // Restore screen updating and alerts
                    application.ScreenUpdating = originalScreenUpdating;
                    application.DisplayAlerts = originalDisplayAlerts;
                    foreach (var sheet in sourceSheetsList)
                    {
                        ReleaseComObjectSafe(sheet);
                    }
                    foreach (var conflict in conflictingSheets.Values)
                    {
                        ReleaseComObjectSafe(conflict.sheet);
                    }
                    ReleaseComObjectSafe(lastSheet);
                    ReleaseComObjectSafe(targetSheets);
                    ReleaseComObjectSafe(targetWorkbook);
                    ReleaseComObjectSafe(sourceWorkbook);
                }
            }
            catch (Exception ex) when (!IsCriticalException(ex))
            {
                Logger.LogException(ex);
                return new CopySheetsResponse
                {
                    Success = false,
                    Message = $"Failed to copy sheets: {ex.Message}",
                    SourceWorkbookName = sourceWorkbookName,
                    TargetWorkbookName = targetWorkbookName,
                    CopiedSheets = new string[0]
                };
            }
        }

        private CloseWorkbookResponse CloseExcelWorkbook(string workbookName, bool saveChanges)
        {
            Workbook workbook = null;
            bool? originalDisplayAlerts = null;
            Microsoft.Office.Interop.Excel.Application application = null;
            try
            {
                Logger.Info($"🚪 CloseExcelWorkbook: Entry - Closing workbook: {workbookName}, saveChanges: {saveChanges}");

                Logger.Info("CloseExcelWorkbook: Resolving Excel Application from container");
                application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Logger.Info($"CloseExcelWorkbook: Excel Application resolved, open workbooks count: {application.Workbooks.Count}");

                originalDisplayAlerts = application.DisplayAlerts;
                application.DisplayAlerts = false;

                // Find the workbook
                Logger.Info($"CloseExcelWorkbook: Searching for workbook '{workbookName}'");
                Workbooks workbooks = null;
                try
                {
                    workbooks = application.Workbooks;
                    var workbookCount = workbooks.Count;
                    for (int i = 1; i <= workbookCount; i++)
                    {
                        Workbook wb = null;
                        try
                        {
                            wb = workbooks[i];
                            Logger.Info($"CloseExcelWorkbook: Checking workbook - Name: '{wb.Name}', FullName: '{wb.FullName}'");
                            if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                                wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = wb;
                                wb = null; // keep reference alive for close
                                Logger.Info($"CloseExcelWorkbook: ✅ Found workbook: '{workbookName}'");
                                break;
                            }
                        }
                        finally
                        {
                            ReleaseComObjectSafe(wb);
                        }
                    }
                }
                finally
                {
                    ReleaseComObjectSafe(workbooks);
                }

                if (workbook == null)
                {
                    Logger.Info($"CloseExcelWorkbook: ⚠️ Workbook '{workbookName}' not found");
                    return new CloseWorkbookResponse
                    {
                        Success = false,
                        Message = $"Workbook '{workbookName}' not found",
                        WorkbookName = workbookName
                    };
                }

                // Close the workbook
                Logger.Info($"CloseExcelWorkbook: Calling Excel COM Close operation, SaveChanges: {saveChanges}");
                workbook.Close(SaveChanges: saveChanges);
                Logger.Info($"CloseExcelWorkbook: Close operation completed");

                Logger.Info($"✅ CloseExcelWorkbook: Successfully closed workbook: {workbookName}");

                return new CloseWorkbookResponse
                {
                    Success = true,
                    Message = $"Successfully closed workbook: {workbookName}",
                    WorkbookName = workbookName
                };
            }
            catch (Exception ex) when (!IsCriticalException(ex))
            {
                Logger.LogException(ex);
                return new CloseWorkbookResponse
                {
                    Success = false,
                    Message = $"Failed to close workbook: {ex.Message}",
                    WorkbookName = workbookName
                };
            }
            finally
            {
                if (application != null && originalDisplayAlerts.HasValue)
                {
                    try
                    {
                        application.DisplayAlerts = originalDisplayAlerts.Value;
                    }
                    catch (Exception ex) when (!IsCriticalException(ex))
                    {
                        Logger.LogException(ex);
                    }
                }

                ReleaseComObjectSafe(workbook);
            }
        }

        private async Task WriteJsonErrorAsync(HttpListenerResponse response, string message, int statusCode = 400)
        {
            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            var error = new { error = message };
            var json = JsonSerializer.Serialize(error, _jsonOptions);
            var buffer = System.Text.Encoding.UTF8.GetBytes(json);
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Close();
            Logger.Info($"WriteJsonErrorAsync: Error response sent and stream closed (status: {statusCode}, message: {message})");
        }

        private async Task SendJsonResponse(HttpListenerResponse response, object data, int statusCode = 200)
        {
            response.StatusCode = statusCode;
            response.ContentType = "application/json";
            var json = JsonSerializer.Serialize(data, _jsonOptions);
            var buffer = System.Text.Encoding.UTF8.GetBytes(json);
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.OutputStream.Close();
            Logger.Info($"SendJsonResponse: Response sent and stream closed (status: {statusCode})");
        }

        private T InvokeExcel<T>(Func<T> action)
        {
            var dispatcher = Container.Resolve<ExcelStaDispatcher>();
            return dispatcher.InvokeExcel(action);
        }
    }

    #region Request/Response Models

    public class OpenWorkbookRequest
    {
        [JsonPropertyName("filePath")]
        public string FilePath { get; set; }

        [JsonPropertyName("visible")]
        public bool Visible { get; set; } = false;
    }

    public class OpenWorkbookResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("workbookFullName")]
        public string WorkbookFullName { get; set; }

        [JsonPropertyName("alreadyOpen")]
        public bool AlreadyOpen { get; set; }
    }

    public class CopySheetsRequest
    {
        [JsonPropertyName("sourceWorkbookName")]
        public string SourceWorkbookName { get; set; }

        [JsonPropertyName("targetWorkbookName")]
        public string TargetWorkbookName { get; set; }

        [JsonPropertyName("sheetNames")]
        public string[] SheetNames { get; set; }

        [JsonPropertyName("namePrefix")]
        public string NamePrefix { get; set; }
    }

    public class CopySheetsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("sourceWorkbookName")]
        public string SourceWorkbookName { get; set; }

        [JsonPropertyName("targetWorkbookName")]
        public string TargetWorkbookName { get; set; }

        [JsonPropertyName("copiedSheets")]
        public string[] CopiedSheets { get; set; }
    }

    public class CloseWorkbookRequest
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("saveChanges")]
        public bool SaveChanges { get; set; } = false;
    }

    public class CloseWorkbookResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }
    }

    #endregion
}
