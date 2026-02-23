using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using X21.Common.Data;
using X21.Services;

namespace X21.Services.Handlers
{
    public class HandleOpenWorkbooksAllHandler : BaseExcelApiService
    {
        private static bool _openXmlNotAvailableLogged;
        private static Action<string> _logInfo;
        private static Action<Exception> _logException;

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

        public HandleOpenWorkbooksAllHandler(Container container) : base(container)
        {
            if (_logInfo == null)
            {
                _logInfo = message => Logger.Info(message);
            }

            if (_logException == null)
            {
                _logException = ex => Logger.LogException(ex);
            }
        }

        public async Task HandleListOpenWorkbooksAll(HttpListenerResponse response)
        {
            try
            {
                var workbooks = InvokeExcel(() => ListOpenWorkbooksAllWithSheets());
                Logger.Info($"[OpenWorkbooksAll] Returned {workbooks?.Count ?? 0} workbooks");
                await SendJsonResponse(response, new { workbooks });
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException || ex is System.Threading.ThreadAbortException)
                {
                    throw;
                }

                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Failed to list open workbooks (all): {ex.Message}", 500);
            }
        }

        private List<OpenWorkbookWithSheets> ListOpenWorkbooksAllWithSheets()
        {
            var results = new List<OpenWorkbookWithSheets>();
            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var stats = EnumerateRotEntries(results, seenKeys);
            AddWorkbooksFromCurrentApplication(results, seenKeys);

            Logger.Info(
                $"[OpenWorkbooksAll] ROT entries={stats.RotEntries}, apps={stats.ApplicationCount}, workbooks={stats.WorkbookCount}, results={results.Count}"
            );
            return results;
        }

        private void AddWorkbooksFromCurrentApplication(
            List<OpenWorkbookWithSheets> results,
            HashSet<string> seenKeys
        )
        {
            try
            {
                var application = Container.Resolve<Application>();
                if (application == null)
                {
                    return;
                }

                AddWorkbooksFromApplication(application, results, seenKeys);
            }
            catch (COMException ex)
            {
                LogComException("[OpenWorkbooksAll][WARN] Current application enumeration failed", ex);
            }
            catch (ObjectDisposedException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] Current application enumeration failed: {SanitizeForLog(ex.Message)}");
            }
            catch (InvalidOperationException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] Current application enumeration failed: {SanitizeForLog(ex.Message)}");
            }
            catch (ArgumentException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] Current application enumeration failed: {SanitizeForLog(ex.Message)}");
            }
        }

        private RotStats EnumerateRotEntries(List<OpenWorkbookWithSheets> results, HashSet<string> seenKeys)
        {
            var stats = new RotStats();
            IRunningObjectTable rot = null;
            IEnumMoniker enumMoniker = null;

            try
            {
                try
                {
                    GetRunningObjectTable(0, out rot);
                    if (rot == null)
                    {
                        return stats;
                    }
                }
                catch (COMException ex)
                {
                    LogComException("[OpenWorkbooksAll][WARN] Failed to access ROT", ex);
                    return stats;
                }

                rot.EnumRunning(out enumMoniker);
                enumMoniker?.Reset();

                var monikers = new IMoniker[1];
                while (enumMoniker != null && enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    stats.RotEntries++;
                    var moniker = monikers[0];
                    if (moniker == null)
                    {
                        continue;
                    }

                    IBindCtx bindCtx = null;
                    object comObject = null;
                    try
                    {
                        CreateBindCtx(0, out bindCtx);
                        if (bindCtx == null)
                        {
                            continue;
                        }

                        moniker.GetDisplayName(bindCtx, null, out var displayName);
                        var isApplication = IsExcelApplicationMoniker(displayName);
                        var isWorkbook = IsExcelWorkbookMoniker(displayName);
                        if (!isApplication && !isWorkbook)
                        {
                            continue;
                        }

                        rot.GetObject(moniker, out comObject);
                        if (comObject is Application app && isApplication)
                        {
                            stats.ApplicationCount++;
                            AddWorkbooksFromApplication(app, results, seenKeys);
                        }
                        else if (comObject is Workbook workbook && isWorkbook)
                        {
                            stats.WorkbookCount++;
                            AddWorkbookFromInstance(workbook, results, seenKeys);
                        }
                    }
                    catch (COMException ex)
                    {
                        LogComException("[OpenWorkbooksAll][WARN] ROT enumeration failed", ex);
                    }
                    finally
                    {
                        ReleaseComObjectSafe(comObject);
                        ReleaseComObjectSafe(bindCtx);
                        ReleaseComObjectSafe(moniker);
                    }
                }
            }
            finally
            {
                ReleaseComObjectSafe(enumMoniker);
                ReleaseComObjectSafe(rot);
            }

            return stats;
        }

        private static bool IsExcelApplicationMoniker(string displayName)
        {
            return !string.IsNullOrWhiteSpace(displayName) &&
                   displayName.IndexOf("Excel.Application", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsExcelWorkbookMoniker(string displayName)
        {
            if (string.IsNullOrWhiteSpace(displayName))
            {
                return false;
            }

            if (displayName.IndexOf("Excel.Workbook", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return displayName.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void AddWorkbooksFromApplication(
            Application application,
            List<OpenWorkbookWithSheets> results,
            HashSet<string> seenKeys
        )
        {
            Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                var count = workbooks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Workbook workbook = null;
                    try
                    {
                        workbook = workbooks.Item[i];
                        AddWorkbookFromInstance(workbook, results, seenKeys, application);
                    }
                    finally
                    {
                        ReleaseComObjectSafe(workbook);
                    }
                }
            }
            catch (COMException ex)
            {
                LogComException("[OpenWorkbooksAll][ERROR] Failed to enumerate workbooks", ex);
            }
            finally
            {
                ReleaseComObjectSafe(workbooks);
            }
        }

        private void AddWorkbookFromInstance(
            Workbook workbook,
            List<OpenWorkbookWithSheets> results,
            HashSet<string> seenKeys,
            Application applicationOverride = null
        )
        {
            Application application = null;
            try
            {
                var name = workbook.Name;
                var fullName = SafeGetWorkbookFullName(workbook);
                if (string.IsNullOrWhiteSpace(fullName))
                {
                    Logger.Info($"[OpenWorkbooksAll][WARN] Workbook full path unavailable for '{SanitizeForLog(name)}'");
                }
                var sheets = GetWorksheetNames(workbook);
                application = applicationOverride ?? workbook.Application;
                var key = BuildWorkbookKey(application, name, fullName);

                if (!seenKeys.Add(key))
                {
                    return;
                }

                results.Add(new OpenWorkbookWithSheets
                {
                    WorkbookName = name,
                    WorkbookFullName = fullName,
                    Sheets = sheets
                });
            }
            catch (COMException ex)
            {
                LogComException("[OpenWorkbooksAll][ERROR] Failed to read workbook", ex);
            }
            finally
            {
                if (applicationOverride == null)
                {
                    ReleaseComObjectSafe(application);
                }
            }
        }

        private static string BuildWorkbookKey(Application application, string workbookName, string workbookFullName)
        {
            var instanceKey = GetApplicationInstanceKey(application);
            var identity = string.IsNullOrWhiteSpace(workbookFullName) ? workbookName : workbookFullName;
            return $"{instanceKey}|{identity}";
        }

        private static string GetApplicationInstanceKey(Application application)
        {
            try
            {
                return application.Hwnd.ToString();
            }
            catch (COMException ex)
            {
                LogComException("[OpenWorkbooksAll][WARN] Failed to read application HWND", ex);
                return application.GetHashCode().ToString();
            }
            catch (InvalidOperationException)
            {
                return application.GetHashCode().ToString();
            }
        }

        private List<string> GetWorksheetNames(Workbook workbook)
        {
            if (workbook == null)
            {
                return new List<string>();
            }

            var path = SafeGetWorkbookFullName(workbook);
            if (!string.IsNullOrWhiteSpace(path) &&
                File.Exists(path) &&
                IsOpenXmlWorkbookExtension(path))
            {
                var openXmlNames = GetWorksheetNamesFromFile(path);
                if (openXmlNames != null)
                {
                    return openXmlNames;
                }
            }

            return GetWorksheetNamesComSafe(workbook);
        }

        private static List<string> GetWorksheetNamesFromFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            try
            {
                var spreadsheetDocumentType = Type.GetType(
                    "DocumentFormat.OpenXml.Packaging.SpreadsheetDocument, DocumentFormat.OpenXml");
                if (spreadsheetDocumentType == null)
                {
                    if (!_openXmlNotAvailableLogged)
                    {
                        _openXmlNotAvailableLogged = true;
                        LogInfo("[OpenWorkbooksAll][INFO] OpenXML SDK not available; using COM enumeration.");
                    }
                    return null;
                }

                var openMethod = spreadsheetDocumentType.GetMethod(
                    "Open",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(Stream), typeof(bool) },
                    null);
                if (openMethod == null)
                {
                    LogInfo("[OpenWorkbooksAll][WARN] OpenXML SDK Open method not found; using COM enumeration.");
                    return null;
                }

                object document = null;
                try
                {
                    using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        document = openMethod.Invoke(null, new object[] { fileStream, false });
                        if (document == null)
                        {
                            LogInfo("[OpenWorkbooksAll][WARN] OpenXML failed to open document; using COM enumeration.");
                            return null;
                        }

                        var names = new List<string>();
                        var workbookPart = document.GetType().GetProperty("WorkbookPart")?.GetValue(document);
                        var workbook = workbookPart?.GetType().GetProperty("Workbook")?.GetValue(workbookPart);
                        var sheets = workbook?.GetType().GetProperty("Sheets")?.GetValue(workbook);

                        if (sheets is IEnumerable enumerable)
                        {
                            foreach (var sheet in enumerable)
                            {
                                var name = GetOpenXmlStringProperty(sheet, "Name");
                                if (!string.IsNullOrWhiteSpace(name))
                                {
                                    names.Add(name);
                                }
                            }
                        }

                        return names;
                    }
                }
                finally
                {
                    if (document != null)
                    {
                        try
                        {
                            document.GetType()
                                .GetMethod("Dispose", BindingFlags.Public | BindingFlags.Instance)
                                ?.Invoke(document, null);
                        }
                        catch (TargetInvocationException)
                        {
                            // Ignore dispose failures to keep fallback path safe
                        }
                        catch (MethodAccessException)
                        {
                            // Ignore dispose failures to keep fallback path safe
                        }
                        catch (TargetException)
                        {
                            // Ignore dispose failures to keep fallback path safe
                        }
                        catch (InvalidOperationException)
                        {
                            // Ignore dispose failures to keep fallback path safe
                        }
                        catch (ArgumentException)
                        {
                            // Ignore dispose failures to keep fallback path safe
                        }
                    }
                }
            }
            catch (TargetInvocationException ex)
            {
                var inner = ex.InnerException ?? ex;
                LogException(inner);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(inner.Message)}");
                return null;
            }
            catch (IOException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (UnauthorizedAccessException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (NotSupportedException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (System.Security.SecurityException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (ArgumentException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (MethodAccessException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (TargetException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
            catch (InvalidOperationException ex)
            {
                LogException(ex);
                LogInfo($"[OpenWorkbooksAll][WARN] OpenXML sheet enumeration failed, falling back to COM: {SanitizeForLog(ex.Message)}");
                return null;
            }
        }

        private static string GetOpenXmlStringProperty(object instance, string propertyName)
        {
            if (instance == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            try
            {
                var property = instance.GetType().GetProperty(propertyName);
                var value = property?.GetValue(instance);
                if (value == null)
                {
                    return null;
                }

                var valueProperty = value.GetType().GetProperty("Value");
                var raw = valueProperty?.GetValue(value);
                return raw?.ToString() ?? value.ToString();
            }
            catch (ArgumentException)
            {
                return null;
            }
            catch (AmbiguousMatchException)
            {
                return null;
            }
            catch (TargetInvocationException)
            {
                return null;
            }
            catch (TargetException)
            {
                return null;
            }
            catch (MethodAccessException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        private static bool IsOpenXmlWorkbookExtension(string path)
        {
            var extension = Path.GetExtension(path);
            return extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> GetWorksheetNamesComSafe(Workbook workbook)
        {
            var names = new List<string>();
            Sheets worksheets = null;
            try
            {
                worksheets = workbook.Worksheets;
                var count = worksheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    Worksheet worksheet = null;
                    try
                    {
                        worksheet = (Worksheet)worksheets.Item[i];
                        if (worksheet != null)
                        {
                            names.Add(worksheet.Name);
                        }
                    }
                    finally
                    {
                        ReleaseComObjectSafe(worksheet);
                    }
                }
            }
            catch (COMException ex)
            {
                LogComException("[OpenWorkbooksAll][ERROR] Failed to enumerate sheets", ex);
            }
            finally
            {
                ReleaseComObjectSafe(worksheets);
            }

            return names;
        }

        private static void ReleaseComObjectSafe(object comObject)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.FinalReleaseComObject(comObject);
                }
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

        private static void LogComException(string context, COMException ex)
        {
            LogException(ex);
            LogInfo($"{context} (HRESULT=0x{ex.ErrorCode:X8}): {SanitizeForLog(ex.Message)}");
        }

        private static void LogInfo(string message)
        {
            _logInfo?.Invoke(message);
        }

        private static void LogException(Exception ex)
        {
            _logException?.Invoke(ex);
        }

        private static string SanitizeForLog(string value)
        {
            return value?.Replace("\r", "").Replace("\n", "");
        }

        private class RotStats
        {
            public int RotEntries { get; set; }
            public int ApplicationCount { get; set; }
            public int WorkbookCount { get; set; }
        }

        public class OpenWorkbookWithSheets
        {
            [JsonPropertyName("workbookName")]
            public string WorkbookName { get; set; }

            [JsonPropertyName("workbookFullName")]
            public string WorkbookFullName { get; set; }

            [JsonPropertyName("sheets")]
            public List<string> Sheets { get; set; }
        }
    }
}
