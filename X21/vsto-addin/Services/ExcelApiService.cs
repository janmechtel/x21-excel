using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using X21.Common.Data;
using X21.Models;
using X21.Excel;
using X21.Services.Formatting;
using X21.Constants;
using Microsoft.Office.Interop.Excel;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Threading;
using X21.Services.Handlers;
using X21.Utils;


namespace X21.Services
{
    public partial class ExcelApiService : BaseExcelApiService
    {
        private HttpListener _httpListener;
        private bool _isRunning;
        private string _baseUrl;
        private readonly int _preferredPort = 8080;
        private readonly FormatManager _formatManager;
        private readonly FormatWriter _formatWriter;
        private readonly FormatRangeService _formatRangeService;
        private readonly ExcelSelection _excelSelection;
        private readonly HttpClient _progressHttpClient = new HttpClient();
        private readonly HandleCopyPasteHandler _handleCopyPasteHandler;
        private readonly HandleWriteValuesBatchHandler _handleWriteValuesBatchHandler;
        private readonly WorkbookOperationsService _workbookOperations;
        private readonly HandleOpenWorkbooksAllHandler _handleOpenWorkbooksAllHandler;
        private const int MaxBatchOperations = 10;

        public ExcelApiService(Container container)
            : base(container)
        {
            _excelSelection = Container.Resolve<ExcelSelection>();
            _formatManager = new FormatManager();
            _formatWriter = new FormatWriter();
            _formatRangeService = new FormatRangeService(_formatManager, _excelSelection);
            _workbookOperations = new WorkbookOperationsService(container);
            _handleCopyPasteHandler = new HandleCopyPasteHandler(Container);
            _handleWriteValuesBatchHandler = new HandleWriteValuesBatchHandler(Container);
            _handleOpenWorkbooksAllHandler = new HandleOpenWorkbooksAllHandler(Container);
        }

        public override void Init()
        {
            base.Init();
            Task.Run(StartAsync);
        }

        public override void Exit()
        {
            Stop();
            base.Exit();
        }

        public async Task StartAsync()
        {
            if (_isRunning) return;

            try
            {
                // Step 1: Find an available port starting from the preferred port
                var actualPort = ExcelApiConfigService.Instance.FindAvailablePort(_preferredPort);
                ExcelApiConfigService.Instance.SetPort(actualPort);
                _baseUrl = ExcelApiConfigService.Instance.BaseUrl;

                // Step 2: Start the HTTP listener
                _httpListener = new HttpListener();
                _httpListener.Prefixes.Add(_baseUrl);
                _httpListener.Start();
                _isRunning = true;

                Logger.Info($"Excel API Service started on {_baseUrl}");
                if (actualPort != _preferredPort)
                {
                    Logger.Info($"Note: Preferred port {_preferredPort} was not available, using port {actualPort}");
                }

                // Start listening for requests directly
                await ListenAsync();
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Failed to start Excel API Service: {ex.Message}");
            }
        }

        public void Stop()
        {
            if (!_isRunning) return;

            try
            {
                _httpListener?.Stop();
                _httpListener?.Close();
                _isRunning = false;
                Logger.Info("Excel API Service stopped");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
            }
        }

        private async Task ListenAsync()
        {
            while (_isRunning)
            {
                try
                {
                    var context = await _httpListener.GetContextAsync();
                    await ProcessRequestAsync(context);
                }
                catch (Exception ex)
                {
                    if (_isRunning)
                    {
                        Logger.LogException(ex);
                    }
                }
            }
        }

        private async Task ProcessRequestAsync(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                // Enable CORS - use dynamic backend URL from BackendConfigService
                var backendUrl = BackendConfigService.Instance.BaseUrl;
                response.Headers.Add("Access-Control-Allow-Origin", backendUrl);
                response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
                response.Headers.Add("Access-Control-Allow-Headers", "Content-Type, Authorization");
                response.Headers.Add("Access-Control-Allow-Credentials", "true");

                // Handle preflight requests
                if (request.HttpMethod == "OPTIONS")
                {
                    response.StatusCode = 200;
                    response.Close();
                    return;
                }

                if (request.HttpMethod == "GET" && request.Url.AbsolutePath == "/api/health")
                {
                    await HandleHealthAsync(response);
                    return;
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath == ApiEndpoints.GetMetadata)
                {
                    await HandleGetMetadataAsync(request, response);
                    return;
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath == ApiEndpoints.OpenWorkbooks)
                {
                    await HandleListOpenWorkbooksAsync(response);
                    return;
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath == ApiEndpoints.OpenWorkbooksAll)
                {
                    await _handleOpenWorkbooksAllHandler.HandleListOpenWorkbooksAll(response);
                    return;
                }
                else if (request.HttpMethod == "GET" && request.Url.AbsolutePath == ApiEndpoints.User)
                {
                    await HandleGetUserAsync(response);
                    return;
                }

                // Handle API requests
                if (request.HttpMethod == "POST" && request.Url.AbsolutePath == ApiEndpoints.ActionsExecute)
                {
                    await HandleActionExecutionAsync(request, response);
                    return;
                }
                else
                {
                    Logger.Info($"Endpoint not found: {request.Url.AbsolutePath.Replace("\r", "").Replace("\n", "")}");
                    response.StatusCode = 404;
                    await WriteJsonErrorAsync(response, "Endpoint not found");
                    return;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 500;
                await WriteJsonErrorAsync(response, "Internal server error", 500);
            }
        }

        private async Task HandleGetMetadataAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            if (!IsExcelReady(out var readinessReason))
            {
                Logger.Info($"[ExcelReady] Not ready for metadata request: {readinessReason}");
                response.StatusCode = (int)HttpStatusCode.Conflict;
                await WriteJsonErrorAsync(
                    response,
                    $"Excel is busy (edit mode or modal dialog). Exit edit mode to continue. {readinessReason}".Trim(),
                    (int)HttpStatusCode.Conflict,
                    "EXCEL_NOT_READY");
                return;
            }

            // NOTE: HttpListenerRequest.QueryString decoding uses system code-page and can corrupt
            // UTF-8 percent-encoded values (e.g. "Ü" -> "Ãœ"). Decode from the raw URL using UTF-8.
            var requestedWorkbookName =
                GetQueryParameterUtf8(request, "workbook_name") ??
                GetQueryParameterUtf8(request, "workbookName");

            try
            {
                var safeWorkbookName = (requestedWorkbookName ?? "(active)").Replace("\r", "").Replace("\n", "");
                Logger.Info($"[Metadata] Requested workbook='{safeWorkbookName}'");
                var metadata = InvokeExcel(() => GetWorkbookMetadata(requestedWorkbookName));

                Logger.Info($"metadata: {JsonSerializer.Serialize(metadata)}");

                await SendJsonResponse(response, metadata);
            }
            catch (WorkbookResolutionException wre)
            {
                Logger.Info($"Workbook resolution error while getting metadata: {wre.Message}");
                await WriteJsonErrorAsync(
                    response,
                    wre.Message,
                    400,
                    wre.ErrorCode,
                    wre.Candidates
                );
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await SendJsonResponse(response, new { error = ex.Message }, 500);
            }
        }

        private static string GetQueryParameterUtf8(HttpListenerRequest request, string key)
        {
            if (request?.Url == null || string.IsNullOrEmpty(key))
            {
                return null;
            }

            var query = request.Url.Query;
            if (string.IsNullOrEmpty(query))
            {
                return null;
            }

            // Trim leading '?'
            if (query.Length > 0 && query[0] == '?')
            {
                query = query.Substring(1);
            }

            // Parse manually so we can decode using UTF-8 percent-decoding.
            // Also normalize '+' to space for form-style query strings.
            var pairs = query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var pair in pairs)
            {
                var idx = pair.IndexOf('=');
                var rawKey = idx >= 0 ? pair.Substring(0, idx) : pair;
                var rawValue = idx >= 0 ? pair.Substring(idx + 1) : string.Empty;

                // Keys in our API are ASCII; still unescape for completeness.
                string decodedKey;
                try
                {
                    decodedKey = Uri.UnescapeDataString(rawKey);
                }
                catch (UriFormatException)
                {
                    decodedKey = rawKey;
                }
                if (!string.Equals(decodedKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.IsNullOrEmpty(rawValue))
                {
                    return string.Empty;
                }

                var normalized = rawValue.Replace("+", "%20");
                try
                {
                    return Uri.UnescapeDataString(normalized);
                }
                catch (UriFormatException)
                {
                    // Fall back to the raw value if decoding fails
                    return rawValue;
                }
            }

            return null;
        }

        private async Task HandleHealthAsync(HttpListenerResponse response)
        {
            try
            {
                var payload = new
                {
                    success = true,
                    timestamp = DateTime.UtcNow.ToString("O")
                };
                await SendJsonResponse(response, payload);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await SendJsonResponse(response, new { success = false, error = ex.Message }, 500);
            }
        }

        private async Task HandleListOpenWorkbooksAsync(HttpListenerResponse response)
        {
            try
            {
                var workbooks = InvokeExcel(() => ListOpenWorkbooks());
                Logger.Info($"[OpenWorkbooks] Returned {workbooks?.Count ?? 0} workbooks");
                await SendJsonResponse(response, new { workbooks });
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException || ex is System.Threading.ThreadAbortException)
                {
                    throw;
                }

                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Failed to list open workbooks: {ex.Message}", 500);
            }
        }


        private async Task HandleGetUserAsync(HttpListenerResponse response)
        {
            try
            {
                var userId = Utils.UserUtils.GetUserId();

                var userData = new
                {
                    userId,
                    timestamp = DateTime.UtcNow.ToString("O")
                };

                Logger.Info($"User data requested: {userId}");

                await SendJsonResponse(response, userData);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await SendJsonResponse(response, new { error = ex.Message }, 500);
            }
        }

        private WorkbookMetadataResponse GetWorkbookMetadata(string requestedWorkbookName)
        {


            var application = Container.Resolve<Application>();
            var workbook = ResolveWorkbook(application, requestedWorkbookName);

            var sheetMetadata = new List<SheetMetadata>();
            string activeSheetName = string.Empty;
            string usedRangeAddress = string.Empty;
            string selectedRange = string.Empty;
            string languageCode = string.Empty;
            string dateLanguage = string.Empty;
            string listSeparator = string.Empty;
            string decimalSeparator = string.Empty;
            string thousandsSeparator = string.Empty;


            foreach (Worksheet ws in workbook.Worksheets)
            {
                var usedRange = GetUsedRangeAddress(ws);
                sheetMetadata.Add(new SheetMetadata
                {
                    Name = ws.Name,
                    UsedRangeAddress = usedRange
                });
            }

            if (workbook.ActiveSheet is Worksheet activeSheet)
            {
                activeSheetName = activeSheet.Name;
                usedRangeAddress = GetUsedRangeAddress(activeSheet);
            }

            // Only return current selection for the active workbook to avoid stale data
            if (application.ActiveWorkbook != null &&
                string.Equals(application.ActiveWorkbook.Name, workbook.Name, StringComparison.OrdinalIgnoreCase))
            {
                selectedRange = _excelSelection.GetSelectedRange();
            }

            if (application.ActiveWorkbook != null)
            {
                languageCode = _excelSelection.GetLanguageCode();
                dateLanguage = _excelSelection.GetDateLanguage();
                listSeparator = _excelSelection.GetListSeparator();
                decimalSeparator = _excelSelection.GetDecimalSeparator();
                thousandsSeparator = _excelSelection.GetThousandsSeparator();
            }

            var allSheetNames = sheetMetadata.Select(s => s.Name).ToList();

            return new WorkbookMetadataResponse
            {
                WorkbookName = workbook.Name,
                WorkbookFullName = SafeGetWorkbookFullName(workbook),
                ActiveSheet = activeSheetName,
                UsedRange = usedRangeAddress,
                SelectedRange = selectedRange,
                AllSheets = allSheetNames,
                Sheets = sheetMetadata,
                LanguageCode = languageCode,
                DateLanguage = dateLanguage,
                ListSeparator = listSeparator,
                DecimalSeparator = decimalSeparator,
                ThousandsSeparator = thousandsSeparator,
            };
        }

        private new Workbook ResolveWorkbook(Application application, string workbookName)
        {
            // Sanitize input for log line-forgery protection
            var sanitizedWorkbookName = workbookName?.Replace("\r", "").Replace("\n", "");
            Logger.Info($"[WorkbookResolver] Request for workbook='{sanitizedWorkbookName ?? "(null)"}'");


            if (application == null)
            {
                throw new WorkbookResolutionException("Excel application is unavailable", ErrorCodes.WorkbookNotFound);
            }

            if (string.IsNullOrWhiteSpace(workbookName))
            {
                var activeWorkbook = application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    throw new WorkbookResolutionException("No active workbook available", ErrorCodes.WorkbookNotFound);
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
                    ErrorCodes.WorkbookNotFound
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
                    ErrorCodes.AmbiguousWorkbook,
                    candidates.ToList()
                );
            }

            Logger.Info($"[WorkbookResolver] Resolved to '{matches.First().Name}'");
            return matches.First();
        }

        private new Worksheet ResolveWorksheetForRead(string sheetName, string workbookName)
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
        private bool IsNonCritical(Exception ex)
        {
            return !(ex is OutOfMemoryException)
                && !(ex is ThreadAbortException)
                && !(ex is ThreadInterruptedException)
                && !(ex is StackOverflowException)
                && !(ex is AccessViolationException);
        }

        private string GetUsedRangeAddress(Worksheet worksheet)
        {
            try
            {
                var usedRange = worksheet.UsedRange;
                if (usedRange == null)
                {
                    return string.Empty;
                }

                // Check if the sheet is truly empty
                // UsedRange returns A1 even for empty sheets, so we need to verify
                if (usedRange.Cells.Count == 1)
                {
                    var cell = usedRange.Cells[1, 1];
                    // Check if A1 is empty (no value, no formula)
                    if (cell.Value2 == null && string.IsNullOrEmpty(cell.Formula as string))
                    {
                        return string.Empty;
                    }
                }

                return usedRange.Address[false, false];
            }
            catch (COMException ex)
            {
                Logger.LogException(ex);
                return string.Empty;
            }
            catch (ArgumentException ex)
            {
                Logger.LogException(ex);
                return string.Empty;
            }
            catch (InvalidOperationException ex)
            {
                Logger.LogException(ex);
                return string.Empty;
            }
            catch (Exception ex) when (IsNonCritical(ex))
            {
                Logger.LogException(ex);
                return string.Empty;
            }
        }

        private List<WorkbookCandidate> ListOpenWorkbooks()
        {
            var application = Container.Resolve<Application>();
            var workbooks = new List<WorkbookCandidate>();

            foreach (Workbook wb in application.Workbooks)
            {
                workbooks.Add(new WorkbookCandidate
                {
                    WorkbookName = wb.Name,
                    WorkbookFullName = SafeGetWorkbookFullName(wb)
                });
            }

            return workbooks;
        }

        private new Workbook ResolveActiveWorkbookForWrite(string providedWorkbookName)
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

        private new Worksheet GetWorksheetFromActiveWorkbook(string sheetName, string providedWorkbookName)
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
                ErrorCodes.WorkbookNotFound,
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

        private async Task HandleActionExecutionAsync(HttpListenerRequest request, HttpListenerResponse response)
        {
            try
            {
                // Read request body
                string requestBody;
                using (var reader = new StreamReader(request.InputStream, Encoding.UTF8))
                {
                    requestBody = await reader.ReadToEndAsync();
                }

                // Parse as generic JSON to get the tool/action type
                var jsonDoc = JsonSerializer.Deserialize<JsonElement>(requestBody, _jsonOptions);

                if (!jsonDoc.TryGetProperty("action", out var actionElement))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Missing 'action' property in request");
                    return;
                }
                Logger.Info($"Action type: {actionElement.GetString()}");
                Logger.Info($"jsonDoc: {jsonDoc}");


                var actionType = actionElement.GetString();

                // Gateway switch statement
                if (!IsExcelReady(out var readinessReason))
                {
                    Logger.Info($"[ExcelReady] Not ready for action '{actionType}': {readinessReason}");
                    response.StatusCode = (int)HttpStatusCode.Conflict;
                    await WriteJsonErrorAsync(
                        response,
                        $"Excel is busy (edit mode or modal dialog). Exit edit mode to continue. {readinessReason}".Trim(),
                        (int)HttpStatusCode.Conflict,
                        "EXCEL_NOT_READY");
                    return;
                }

                switch (actionType?.ToLower())
                {
                    case ExcelToolNames.ReadValuesBatch:
                        await HandleReadValuesBatch(requestBody, response);
                        break;

                    case ExcelToolNames.WriteValuesBatch:
                        await _handleWriteValuesBatchHandler.HandleWriteValuesBatch(requestBody, response);
                        break;

                    case ExcelToolNames.WriteFormatBatch:
                        await HandleWriteFormatBatch(requestBody, response);
                        break;

                    case ExcelToolNames.ReadFormatBatch:
                        await HandleReadFormatBatch(requestBody, response);
                        break;

                    case ExcelToolNames.RemoveColumns:
                        await HandleDeleteColumns(requestBody, response);
                        break;

                    case ExcelToolNames.RemoveRows:
                        await HandleDeleteRows(requestBody, response);
                        break;

                    case ExcelToolNames.AddRows:
                        await HandleInsertRows(requestBody, response);
                        break;

                    case ExcelToolNames.AddColumns:
                        await HandleInsertColumns(requestBody, response);
                        break;

                    case ExcelToolNames.AddSheets:
                        await HandleAddSheets(requestBody, response);
                        break;

                    case ExcelToolNames.RemoveSheets:
                        await HandleRemoveSheets(requestBody, response);
                        break;

                    case ExcelToolNames.CreateChart:
                        await HandleCreateChart(requestBody, response);
                        break;

                    case ExcelToolNames.VbaCreate:
                        await HandleVBA(requestBody, response);
                        break;

                    case ExcelToolNames.VbaRead:
                        await HandleVBAReadTool(requestBody, response);
                        break;

                    case ExcelToolNames.VbaUpdate:
                        await HandleVBAUpdateTool(requestBody, response);
                        break;

                    case ExcelToolNames.DragFormula:
                        await HandleDragFormula(requestBody, response);
                        break;

                    case "open_workbook":
                        await _workbookOperations.HandleOpenWorkbook(requestBody, response);
                        break;

                    case "copy_sheets":
                        await _workbookOperations.HandleCopySheets(requestBody, response);
                        break;

                    case "close_workbook":
                        await _workbookOperations.HandleCloseWorkbook(requestBody, response);
                        break;

                    case ExcelToolNames.CopyPaste:
                        await _handleCopyPasteHandler.HandleCopyPaste(requestBody, response);
                        break;

                    case ExcelToolNames.DeleteCells:
                        await _handleCopyPasteHandler.HandleDeleteCells(requestBody, response);
                        break;

                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                response.StatusCode = 500;
                await WriteJsonErrorAsync(response, "Internal server error", 500);
            }
        }

        private bool IsExcelReady(out string reason)
        {
            reason = string.Empty;
            try
            {
                var application = Container.Resolve<Application>();
                if (application == null)
                {
                    reason = "Excel application unavailable.";
                    Logger.Info($"[ExcelReady] {reason}");
                    return false;
                }

                try
                {
                    // Touch a lightweight property to ensure the COM proxy is still valid.
                    InvokeExcel(() => application.Hwnd);
                }
                catch (InvalidComObjectException ex)
                {
                    reason = $"Excel application unavailable (COM object released). ({ex.Message})";
                    Logger.Info($"[ExcelReady] {reason}");
                    return false;
                }
                catch (COMException ex)
                {
                    reason = $"Excel application unavailable. ({ex.Message})";
                    Logger.Info($"[ExcelReady] {reason}");
                    return false;
                }

                bool isReady;
                try
                {
                    isReady = InvokeExcel(() => application.Ready);
                }
                catch (COMException ex)
                {
                    reason = ex.Message;
                    Logger.Info($"[ExcelReady] COMException while checking Application.Ready: {reason}");
                    return false;
                }

                if (!isReady)
                {
                    reason = "Excel is not ready.";
                    Logger.Info($"[ExcelReady] {reason}");
                    return false;
                }

                try
                {
                    InvokeExcel(() =>
                    {
                        // ReferenceStyle self-assignment throws in edit mode.
                        var style = application.ReferenceStyle;
                        application.ReferenceStyle = style;
                    });
                }
                catch (COMException ex)
                {
                    reason = $"Excel is in edit mode. ({ex.Message})";
                    Logger.Info($"[ExcelReady] {reason}");
                    return false;
                }

                try
                {
                    var calcState = InvokeExcel(() => application.CalculationState);
                    if (calcState != XlCalculationState.xlDone)
                    {
                        reason = "Excel calculation in progress.";
                        Logger.Info($"[ExcelReady] {reason}");
                        return false;
                    }
                }
                catch (COMException ex)
                {
                    reason = ex.Message;
                    Logger.Info($"[ExcelReady] COMException while checking CalculationState: {reason}");
                    return false;
                }

                Logger.Info("[ExcelReady] Excel is ready.");
                return true;
            }
            catch (COMException ex)
            {
                reason = ex.Message;
                Logger.Info($"[ExcelReady] COMException while checking readiness: {reason}");
                return false;
            }
        }

        private async Task HandleRemoveSheets(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var removeSheetsRequest = JsonSerializer.Deserialize<RemoveSheetsRequest>(requestBody, _jsonOptions);
                Logger.Info($"removeSheetsRequest: {removeSheetsRequest}");

                if (removeSheetsRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (removeSheetsRequest.SheetNames == null || removeSheetsRequest.SheetNames.Length == 0 || string.IsNullOrEmpty(removeSheetsRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "sheetNames array and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => RemoveExcelSheets(removeSheetsRequest.WorkbookName, removeSheetsRequest.SheetNames));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully removed {result.SheetsRemoved?.Length ?? 0} sheets from workbook {removeSheetsRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error removing sheets: {ex.Message}", 500);
            }
        }

        private async Task HandleReadValues(string requestBody, HttpListenerResponse response)
        {
            var readRangeRequest = JsonSerializer.Deserialize<ReadRangeRequest>(requestBody, _jsonOptions);
            Logger.Info($"readRangeRequest: {readRangeRequest}");
            if (readRangeRequest == null)
            {
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "Invalid request body");
                return;
            }

            if (string.IsNullOrEmpty(readRangeRequest.Worksheet) || string.IsNullOrEmpty(readRangeRequest.Range))
            {
                Logger.Info($"one is empty: worksheet: {readRangeRequest.Worksheet}, range: {readRangeRequest.Range}, workbookName: {readRangeRequest.WorkbookName}");
                response.StatusCode = 400;
                await WriteJsonErrorAsync(response, "worksheet and range are required");
                return;
            }

            try
            {
                var result = InvokeExcel(() => ReadExcelRange(readRangeRequest.Worksheet, readRangeRequest.WorkbookName, readRangeRequest.Range));
                Logger.Info($"result read values: {JsonSerializer.Serialize(result)}");
                await SendJsonResponse(response, result);
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error reading values: {ex.Message}", 500);
            }
        }

        private async Task HandleReadValuesBatch(string requestBody, HttpListenerResponse response)
        {
            try
            {
                Logger.Info($"?? Read values batch request body: {requestBody.Replace("\r", "").Replace("\n", "")}");

                var batchRequest = JsonSerializer.Deserialize<ReadRangeBatchRequest>(requestBody, _jsonOptions);
                if (batchRequest?.Operations == null || batchRequest.Operations.Length == 0)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "operations array is required and cannot be empty");
                    return;
                }

                if (batchRequest.Operations.Length > MaxBatchOperations)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(
                        response,
                        $"operations array exceeds maximum of {MaxBatchOperations}",
                        400);
                    return;
                }

                var batchResponse = InvokeExcel(() =>
                {
                    var results = new List<ReadRangeBatchResult>();
                    Logger.Info($"?? Read values batch start: operations={batchRequest.Operations.Length}");

                    foreach (var op in batchRequest.Operations)
                    {
                        Logger.Info($"?? Read values op -> workbook={op.WorkbookName}, sheet={op.Worksheet}, range={op.Range}");

                        if (string.IsNullOrEmpty(op.Worksheet) || string.IsNullOrEmpty(op.Range))
                        {
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = "worksheet and range are required"
                            });
                            continue;
                        }

                        try
                        {
                            var res = ReadExcelRange(op.Worksheet, op.WorkbookName, op.Range);
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = true,
                                Message = "Read values",
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                CellValues = res.CellValues
                            });
                        }
                        catch (WorkbookResolutionException wre)
                        {
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = wre.Message
                            });
                        }
                        catch (COMException comEx)
                        {
                            Logger.LogException(comEx);
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = $"Error reading values from Excel (COM error): {comEx.Message}"
                            });
                        }
                        catch (ArgumentException argEx)
                        {
                            Logger.LogException(argEx);
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = $"Invalid range or arguments: {argEx.Message}"
                            });
                        }
                        catch (InvalidOperationException invOpEx)
                        {
                            Logger.LogException(invOpEx);
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = $"Error reading values due to workbook or worksheet state: {invOpEx.Message}"
                            });
                        }
                        catch (Exception ex)
                        {
                            // Rethrow critical exceptions so they are not silently converted into normal failures.
                            if (ex is OutOfMemoryException ||
                                ex is StackOverflowException ||
                                ex is ThreadAbortException ||
                                ex is ThreadInterruptedException ||
                                ex is AccessViolationException)
                            {
                                throw;
                            }

                            Logger.LogException(ex);
                            results.Add(new ReadRangeBatchResult
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Range = op.Range,
                                Message = $"Unexpected error ({ex.GetType().Name}) reading values: {ex.Message}"
                            });
                        }
                    }

                    var successCount = results.Count(r => r.Success);
                    return new ReadRangeBatchResponse
                    {
                        Success = successCount == results.Count,
                        Message = $"Read values for {successCount}/{results.Count} range(s)",
                        Results = results.ToArray()
                    };
                });

                await SendJsonResponse(response, batchResponse);
                Logger.Info($"?? Read values batch completed: {batchResponse.Message}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (JsonException jex)
            {
                Logger.LogException(jex);
                await WriteJsonErrorAsync(response, $"Invalid JSON in read values batch request: {jex.Message}", 400);
            }
            catch (IOException ioex)
            {
                Logger.LogException(ioex);
                await WriteJsonErrorAsync(response, $"I/O error while processing read values batch: {ioex.Message}", 500);
            }
        }


        private ReadRangeResponse ReadExcelRange(string sheetName, string workbookName, string range)
        {
            try
            {
                Logger.Info($"Reading range {range} from worksheet '{sheetName}' with format information");

                var worksheet = ResolveWorksheetForRead(sheetName, workbookName);
                var resolvedWbName = (worksheet?.Parent as Workbook)?.Name ?? workbookName;
                Logger.Info($"[ReadValues] Resolved workbook='{resolvedWbName}', sheet='{worksheet?.Name}'");

                var cellData = GetRangeDataFromWorksheet(worksheet, range);
                var cellValues = cellData.ToDictionary(
                    c => c.Address,
                    c => new CellValue
                    {
                        Value = InvariantValueFormatter.ToInvariantString(c.Value),
                        Formula = c.Formula ?? ""
                    });

                return new ReadRangeResponse
                {
                    CellValues = cellValues,
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                throw new Exception($"Failed to read range {range} from sheet {sheetName}: {ex.Message}");
            }
        }

        private new async Task WriteJsonErrorAsync(HttpListenerResponse response, string message, int statusCode = 400, string errorCode = null, IEnumerable<WorkbookCandidate> candidates = null)
        {
            var errorResponse = new
            {
                error = message,
                errorCode,
                candidates
            };
            await SendJsonResponse(response, errorResponse, statusCode);
        }

        private new async Task SendJsonResponse(HttpListenerResponse response, object data, int statusCode = 200)
        {
            response.StatusCode = statusCode;
            response.ContentType = "application/json";

            var jsonResponse = JsonSerializer.Serialize(data, _jsonOptions);
            var buffer = Encoding.UTF8.GetBytes(jsonResponse);

            response.ContentLength64 = buffer.Length;
            await response.OutputStream.WriteAsync(buffer, 0, buffer.Length);
            response.Close();
        }

        private async Task SendProgressUpdate(
            string workbookName,
            string status,
            string message,
            int? current = null,
            int? total = null,
            string unit = "cells",
            string worksheet = null,
            string range = null)
        {
            try
            {
                var payload = new
                {
                    status,
                    message,
                    progress = current.HasValue && total.HasValue
                        ? new
                        {
                            current = Math.Max(0, current.Value),
                            total = Math.Max(1, total.Value),
                            unit = unit ?? "cells"
                        }
                        : null,
                    metadata = new
                    {
                        range,
                        worksheet
                    }
                };

                var progressUrl = $"{BackendConfigService.Instance.BaseUrl.TrimEnd('/')}/api/progress";
                var json = JsonSerializer.Serialize(payload, _jsonOptions);

                using var requestMessage = new HttpRequestMessage(HttpMethod.Post, progressUrl);
                requestMessage.Content = new StringContent(json, Encoding.UTF8, "application/json");

                if (!string.IsNullOrEmpty(workbookName))
                {
                    requestMessage.Headers.Add("X-Workbook-Name", workbookName);
                }

                await _progressHttpClient.SendAsync(requestMessage);
            }
            catch (Exception ex)
            {
                // Progress updates are non-critical, don't throw
                Logger.Info($"Progress update failed: {ex.Message}");
            }
        }

        private string GetFormatSummary(FormatSettings format)
        {
            if (format == null) return "none";

            var parts = new List<string>();
            if (format.Bold.HasValue) parts.Add($"bold={format.Bold.Value}");
            if (format.Italic.HasValue) parts.Add($"italic={format.Italic.Value}");
            if (format.Underline.HasValue) parts.Add($"underline={format.Underline.Value}");
            if (format.FontSize.HasValue) parts.Add($"fontSize={format.FontSize.Value}");
            if (!string.IsNullOrEmpty(format.FontName)) parts.Add($"fontName={format.FontName}");
            if (!string.IsNullOrEmpty(format.FontColor)) parts.Add($"fontColor={format.FontColor}");
            if (!string.IsNullOrEmpty(format.BackgroundColor)) parts.Add($"background={format.BackgroundColor}");
            if (!string.IsNullOrEmpty(format.Alignment)) parts.Add($"alignment={format.Alignment}");
            if (!string.IsNullOrEmpty(format.NumberFormat)) parts.Add($"numberFormat={format.NumberFormat}");

            return parts.Count == 0 ? "empty" : string.Join(", ", parts);
        }

        private async Task<(FormatRangeResponse Response, string ErrorMessage)> ApplyRangeFormatting(FormatRangeRequest request, bool enableProgressUpdates = true)
        {
            Logger.Info($"🎨 Applying formatting to {request.Range} in worksheet '{request.Worksheet}' in workbook '{request.WorkbookName}'");

            var worksheet = GetWorksheetFromActiveWorkbook(request.Worksheet, request.WorkbookName);

            var ranges = (request.Range ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(r => r.Trim())
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .ToList();

            if (ranges.Count == 0)
            {
                return (null, "range is required and cannot be empty");
            }

            int successCount = 0;
            int processedCells = 0;
            int totalCells = 0;

            // Pre-compute total cells for progress
            foreach (var r in ranges)
            {
                try
                {
                    var area = worksheet.Range[r];
                    totalCells += Math.Max(1, Convert.ToInt32(area.Cells.Count));
                }
                catch
                {
                    // Ignore invalid area during pre-scan; it will be handled in the main loop
                }
            }
            totalCells = Math.Max(1, totalCells);

            foreach (var areaRange in ranges)
            {
                try
                {
                    Logger.Info($"🎨 Applying format to range {areaRange}");

                    var targetRange = worksheet.Range[areaRange];
                    var areaCells = Math.Max(1, Convert.ToInt32(targetRange.Cells.Count));

                    if (enableProgressUpdates)
                    {
                        await SendProgressUpdate(
                            request.WorkbookName,
                            OperationStatusValues.WritingExcelFormat,
                            $"Applying formatting to {areaRange}...",
                            processedCells,
                            totalCells,
                            "cells",
                            request.Worksheet,
                            areaRange);
                    }

                    _formatWriter.ApplyFormattingToRange(targetRange, request.Format);

                    successCount++;
                    processedCells += areaCells;

                    if (enableProgressUpdates)
                    {
                        await SendProgressUpdate(
                            request.WorkbookName,
                            OperationStatusValues.WritingExcelFormat,
                            $"Finished applying formatting to {areaRange}",
                            processedCells,
                            totalCells,
                            "cells",
                            request.Worksheet,
                            areaRange);
                    }
                    Logger.Info($"✅ Successfully applied formatting to range {areaRange}");
                }
                catch (Exception ex)
                {
                    Logger.Info($"🎨 Error applying formatting to range {areaRange}: {ex.Message}");
                    if (enableProgressUpdates)
                    {
                        await SendProgressUpdate(
                            request.WorkbookName,
                            OperationStatusValues.Error,
                            $"Error applying formatting to {areaRange}: {ex.Message}",
                            processedCells,
                            totalCells,
                            "cells",
                            request.Worksheet,
                            areaRange);
                    }
                }
            }

            // Don't send terminal "idle" status - server handles completion status
            // when HTTP response returns. Sending "idle" here causes race conditions.
            if (enableProgressUpdates && successCount != ranges.Count)
            {
                // Only send error status if operation had failures
                await SendProgressUpdate(
                    request.WorkbookName,
                    OperationStatusValues.Error,
                    $"Formatting completed with errors ({successCount}/{ranges.Count} range(s) succeeded)",
                    null,
                    null,
                    "cells",
                    request.Worksheet,
                    null);
            }

            var response = new FormatRangeResponse
            {
                Worksheet = request.Worksheet,
                WorkbookName = request.WorkbookName,
                Success = successCount == ranges.Count,
                Message = $"Successfully formatted {successCount}/{ranges.Count} range(s)",
            };

            return (response, null);
        }

        private async Task<WriteFormatBatchResponse> ApplyRangeFormattingBatch(WriteFormatBatchRequest batchRequest)
        {
            var results = new List<FormatRangeResponse>();
            int successCount = 0;

            var application = Container.Resolve<Application>();
            var originalScreenUpdating = application.ScreenUpdating;
            var originalEnableEvents = application.EnableEvents;
            var originalCalculation = application.Calculation;

            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;
                application.Calculation = XlCalculation.xlCalculationManual;

                Logger.Info($"?? Applying batch formatting to {batchRequest.Operations.Count} range(s)");

                foreach (var op in batchRequest.Operations)
                {
                    if (string.IsNullOrEmpty(op.Worksheet) ||
                        string.IsNullOrEmpty(op.Range) || op.Format == null)
                    {
                        results.Add(new FormatRangeResponse
                        {
                            Worksheet = op.Worksheet,
                            WorkbookName = op.WorkbookName,
                            Success = false,
                            Message = "worksheet, range, and format are required"
                        });
                        continue;
                    }

                    try
                    {
                        Logger.Info($"[BatchFormat][NoUI] workbook={op.WorkbookName}, sheet={op.Worksheet}, range={op.Range}, format={GetFormatSummary(op.Format)}");

                        var (formatResponse, errorMessage) = await ApplyRangeFormatting(op, false);

                        if (!string.IsNullOrEmpty(errorMessage))
                        {
                            results.Add(new FormatRangeResponse
                            {
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Success = false,
                                Message = errorMessage
                            });
                            continue;
                        }

                        results.Add(formatResponse);
                        if (formatResponse.Success) successCount++;
                    }
                    catch (WorkbookResolutionException wre)
                    {
                        results.Add(new FormatRangeResponse
                        {
                            Worksheet = op.Worksheet,
                            WorkbookName = op.WorkbookName,
                            Success = false,
                            Message = $"{wre.ErrorCode}: {wre.Message}"
                        });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new FormatRangeResponse
                        {
                            Worksheet = op.Worksheet,
                            WorkbookName = op.WorkbookName,
                            Success = false,
                            Message = $"Error applying format: {ex.Message}"
                        });
                    }
                }
            }
            finally
            {
                try { application.ScreenUpdating = originalScreenUpdating; } catch { }
                try { application.EnableEvents = originalEnableEvents; } catch { }
                try { application.Calculation = originalCalculation; } catch { }
            }

            return new WriteFormatBatchResponse
            {
                Success = successCount == results.Count,
                Message = $"Successfully formatted {successCount}/{results.Count} range(s)",
                Results = results
            };
        }

        private async Task HandleWriteFormatBatch(string requestBody, HttpListenerResponse response)
        {
            try
            {
                Logger.Info($"?? Write format batch request body: {requestBody.Replace("\r", "").Replace("\n", "")}");

                var batchRequest = JsonSerializer.Deserialize<WriteFormatBatchRequest>(requestBody, _jsonOptions);
                if (batchRequest?.Operations == null || batchRequest.Operations.Count == 0)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "operations array is required and cannot be empty");
                    return;
                }

                if (batchRequest.Operations.Count > MaxBatchOperations)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(
                        response,
                        $"operations array exceeds maximum of {MaxBatchOperations}",
                        400);
                    return;
                }

                var batchResponse = await InvokeExcel(async () => await ApplyRangeFormattingBatch(batchRequest));

                await SendJsonResponse(response, batchResponse);
                Logger.Info($"?? Write format batch completed: {batchResponse.Message}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error writing format batch: {ex.Message}", 500);
            }
        }

        private async Task HandleReadFormatBatch(string requestBody, HttpListenerResponse response)
        {
            ReadFormatBatchRequest batchRequest = null;
            try
            {
                Logger.Info($"?? Read format batch request body: {requestBody.Replace("\r", "").Replace("\n", "")}");

                batchRequest = JsonSerializer.Deserialize<ReadFormatBatchRequest>(requestBody, _jsonOptions);
                if (batchRequest?.Operations == null || batchRequest.Operations.Count == 0)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "operations array is required and cannot be empty");
                    return;
                }

                if (batchRequest.Operations.Count > MaxBatchOperations)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(
                        response,
                        $"operations array exceeds maximum of {MaxBatchOperations}",
                        400);
                    return;
                }

                var batchResponse = await InvokeExcel(async () =>
                {
                    var results = new List<ReadFormatResponse>();
                    Logger.Info($"?? Read format batch start: operations={batchRequest.Operations.Count}");
                    foreach (var op in batchRequest.Operations)
                    {
                        Logger.Info($"?? Read format op -> workbook={op.WorkbookName}, sheet={op.Worksheet}, range={op.Range}");
                        if (string.IsNullOrEmpty(op.Worksheet) || string.IsNullOrEmpty(op.WorkbookName) ||
                            string.IsNullOrEmpty(op.Range))
                        {
                            results.Add(new ReadFormatResponse
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Message = "worksheet, workbookName, and range are required"
                            });
                            continue;
                        }

                        try
                        {
                            var progressCallback = new Action<FormatProgressUpdate>(update =>
                            {
                                // Fire-and-forget for intermediate progress updates
                                _ = SendProgressUpdate(
                                    op.WorkbookName,
                                    OperationStatusValues.ReadingExcelFormat,
                                    update.Message,
                                    update.Current,
                                    update.Total,
                                    "cells",
                                    op.Worksheet,
                                    op.Range);
                            });

                            var res = await _formatRangeService.ReadExcelRangeFormatAsync(
                                op.Worksheet,
                                op.WorkbookName,
                                op.Range,
                                op.PropertiesToRead,
                                progressCallback);
                            Logger.Info($"?? Read format op success -> workbook={op.WorkbookName}, sheet={op.Worksheet}, range={op.Range}, cells={res.CellFormats?.Count ?? 0}");
                            results.Add(res);
                        }
                        catch (Exception ex)
                        {
                            Logger.Info($"?? Read format op failed -> workbook={op.WorkbookName}, sheet={op.Worksheet}, range={op.Range}, error={ex.Message}");
                            await SendProgressUpdate(
                                op.WorkbookName,
                                OperationStatusValues.Error,
                                $"Error reading format for {op.Range}: {ex.Message}",
                                null,
                                null,
                                "cells",
                                op.Worksheet,
                                op.Range);
                            results.Add(new ReadFormatResponse
                            {
                                Success = false,
                                Worksheet = op.Worksheet,
                                WorkbookName = op.WorkbookName,
                                Message = $"Error reading format: {ex.Message}"
                            });
                        }
                    }

                    var successCount = results.Count(r => r.Success);
                    return new ReadFormatBatchResponse
                    {
                        Success = successCount == results.Count,
                        Message = $"Successfully read formats for {successCount}/{results.Count} range(s)",
                        Results = results
                    };
                });

                await SendJsonResponse(response, batchResponse);
                Logger.Info($"?? Read format batch completed: {batchResponse.Message}");

                if (!batchResponse.Success)
                {
                    // Send error status to all workbooks in the batch (best-effort)
                    var workbookNames = batchRequest?.Operations
                        ?.Select(op => op?.WorkbookName)
                        .Where(w => !string.IsNullOrEmpty(w))
                        .Distinct()
                        .ToList();

                    if (workbookNames != null)
                    {
                        foreach (var wb in workbookNames)
                        {
                            try
                            {
                                await SendProgressUpdate(wb, OperationStatusValues.Error, batchResponse.Message, null, null, "cells", null, null);
                            }
                            catch (Exception sendEx)
                            {
                                Logger.Info($"Non-fatal error sending batch error status: {sendEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);

                await WriteJsonErrorAsync(response, $"Error reading format batch: {ex.Message}", 500);
            }
        }

        private async Task HandleReadFormat(string requestBody, HttpListenerResponse response)
        {
            X21.Models.ReadFormatRequest readFormatRequest = null;
            try
            {
                readFormatRequest = JsonSerializer.Deserialize<X21.Models.ReadFormatRequest>(requestBody, _jsonOptions);

                if (readFormatRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(readFormatRequest.Worksheet) || string.IsNullOrEmpty(readFormatRequest.Range) || string.IsNullOrEmpty(readFormatRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, range and workbookName are required");
                    return;
                }

                var progressCallback = new Action<FormatProgressUpdate>(update =>
                {
                    // Fire-and-forget for intermediate progress updates
                    _ = SendProgressUpdate(
                        readFormatRequest.WorkbookName,
                        OperationStatusValues.ReadingExcelFormat,
                        update.Message,
                        update.Current,
                        update.Total,
                        "cells",
                        readFormatRequest.Worksheet,
                        readFormatRequest.Range);
                });

                var result = await InvokeExcel(async () => await _formatRangeService.ReadExcelRangeFormatAsync(
                    readFormatRequest.Worksheet,
                    readFormatRequest.WorkbookName,
                    readFormatRequest.Range,
                    readFormatRequest.PropertiesToRead,
                    progressCallback));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully read format for range {readFormatRequest.Range} in sheet {readFormatRequest.Worksheet} in workbook {readFormatRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);

                await SendProgressUpdate(
                    readFormatRequest?.WorkbookName,
                    OperationStatusValues.Error,
                    $"Failed to read format for {readFormatRequest?.Range}: {ex.Message}",
                    null,
                    null,
                    "cells",
                    readFormatRequest?.Worksheet,
                    readFormatRequest?.Range);

                await WriteJsonErrorAsync(response, $"Error reading range format: {ex.Message}", 500);
            }
        }

        private async Task HandleDeleteColumns(string requestBody, HttpListenerResponse response)
        {
            try
            {
                // Logger.Info($"🔧 Delete columns request: {requestBody}");
                var deleteColumnsRequest = JsonSerializer.Deserialize<DeleteColumnsRequest>(requestBody, _jsonOptions);

                if (deleteColumnsRequest == null)
                {
                    // Logger.Info($"❌ Invalid request body: {requestBody}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(deleteColumnsRequest.Worksheet) || string.IsNullOrEmpty(deleteColumnsRequest.ColumnRange) || string.IsNullOrEmpty(deleteColumnsRequest.WorkbookName))
                {
                    Logger.Info($"❌ worksheet: {deleteColumnsRequest.Worksheet}, columns: {deleteColumnsRequest.ColumnRange}, workbookName: {deleteColumnsRequest.WorkbookName}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, columns and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => DeleteExcelColumns(deleteColumnsRequest.Worksheet, deleteColumnsRequest.WorkbookName, deleteColumnsRequest.ColumnRange));
                Logger.Info($"🔧 Delete columns result: {JsonSerializer.Serialize(result, _jsonOptions)}");
                await SendJsonResponse(response, result);

                Logger.Info($"Successfully deleted columns {deleteColumnsRequest.ColumnRange} in sheet {deleteColumnsRequest.Worksheet} in workbook {deleteColumnsRequest.WorkbookName}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error deleting columns: {ex.Message}", 500);
            }
        }

        private async Task HandleDeleteRows(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var deleteRowsRequest = JsonSerializer.Deserialize<DeleteRowsRequest>(requestBody, _jsonOptions);
                Logger.Info($"🔧 Delete rows request: {JsonSerializer.Serialize(deleteRowsRequest, _jsonOptions)}");
                if (deleteRowsRequest == null)
                {
                    // Logger.Info($"❌ Invalid request body: {requestBody}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(deleteRowsRequest.Worksheet) || string.IsNullOrEmpty(deleteRowsRequest.RowRange) || string.IsNullOrEmpty(deleteRowsRequest.WorkbookName))
                {
                    Logger.Info($"❌ worksheet: {deleteRowsRequest.Worksheet}, rowRange: {deleteRowsRequest.RowRange}, workbookName: {deleteRowsRequest.WorkbookName}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, rows and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => DeleteExcelRows(deleteRowsRequest.Worksheet, deleteRowsRequest.WorkbookName, deleteRowsRequest.RowRange));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully deleted rows {deleteRowsRequest.RowRange} in sheet {deleteRowsRequest.Worksheet} in workbook {deleteRowsRequest.WorkbookName}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error deleting rows: {ex.Message}", 500);
            }
        }

        private async Task HandleInsertRows(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var insertRowsRequest = JsonSerializer.Deserialize<InsertRowsRequest>(requestBody, _jsonOptions);

                if (insertRowsRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(insertRowsRequest.Worksheet) || string.IsNullOrEmpty(insertRowsRequest.RowRange) || string.IsNullOrEmpty(insertRowsRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, rowRange and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => InsertExcelRows(insertRowsRequest.Worksheet, insertRowsRequest.WorkbookName, insertRowsRequest.RowRange));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully inserted rows at range {insertRowsRequest.RowRange} in sheet {insertRowsRequest.Worksheet} in workbook {insertRowsRequest.WorkbookName}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error inserting rows: {ex.Message}", 500);
            }
        }

        private async Task HandleInsertColumns(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var insertColumnsRequest = JsonSerializer.Deserialize<InsertColumnsRequest>(requestBody, _jsonOptions);

                if (insertColumnsRequest == null)
                {
                    // var sanitizedRequestBody = requestBody?
                    //     .Replace("\r\n", " ")
                    //     .Replace("\n", " ")
                    //     .Replace("\r", " ");
                    // Logger.Info($"❌ Invalid request body: {sanitizedRequestBody}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(insertColumnsRequest.Worksheet) || string.IsNullOrEmpty(insertColumnsRequest.ColumnRange) || string.IsNullOrEmpty(insertColumnsRequest.WorkbookName))
                {
                    Logger.Info($"❌ worksheet: {insertColumnsRequest.Worksheet}, columnRange: {insertColumnsRequest.ColumnRange}, workbookName: {insertColumnsRequest.WorkbookName}");
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, columnRange and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => InsertExcelColumns(insertColumnsRequest.Worksheet, insertColumnsRequest.WorkbookName, insertColumnsRequest.ColumnRange));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully inserted columns at range {insertColumnsRequest.ColumnRange} in sheet {insertColumnsRequest.Worksheet} in workbook {insertColumnsRequest.WorkbookName}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error inserting columns: {ex.Message}", 500);
            }
        }

        private async Task HandleAddSheets(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var addSheetsRequest = JsonSerializer.Deserialize<AddSheetsRequest>(requestBody, _jsonOptions);

                if (addSheetsRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (addSheetsRequest.SheetNames == null || addSheetsRequest.SheetNames.Length == 0 || string.IsNullOrEmpty(addSheetsRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "sheetNames array and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => AddExcelSheets(addSheetsRequest.WorkbookName, addSheetsRequest.SheetNames));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully added {result.SheetsAdded?.Length ?? 0} sheets in workbook {addSheetsRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error adding sheets: {ex.Message}", 500);
            }
        }

        private async Task HandleCreateChart(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var createChartRequest = JsonSerializer.Deserialize<CreateChartRequest>(requestBody, _jsonOptions);

                if (createChartRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(createChartRequest.ChartType) || string.IsNullOrEmpty(createChartRequest.DataRange) || string.IsNullOrEmpty(createChartRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "chartType, dataRange and workbookName are required");
                    return;
                }

                if (createChartRequest.ChartType != ChartTypes.Line && createChartRequest.ChartType != ChartTypes.Histogram && createChartRequest.ChartType != ChartTypes.Pie)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "chartType must be 'line', 'histogram', or 'pie'");
                    return;
                }

                var result = InvokeExcel(() => CreateExcelChart(createChartRequest.WorkbookName, createChartRequest.ChartType,
                    createChartRequest.DataRange, createChartRequest.Title, createChartRequest.XAxisTitle,
                    createChartRequest.YAxisTitle, createChartRequest.ChartLocation));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully created {createChartRequest.ChartType} chart in workbook {createChartRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error creating chart: {ex.Message}", 500);
            }
        }

        private async Task HandleVBA(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var vbaRequest = JsonSerializer.Deserialize<VBARequest>(requestBody, _jsonOptions);

                if (vbaRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(vbaRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "workbookName is required");
                    return;
                }

                if (string.IsNullOrEmpty(vbaRequest.FunctionName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "functionName is required");
                    return;
                }

                if (string.IsNullOrEmpty(vbaRequest.VbaCode))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "vbaCode is required");
                    return;
                }

                var result = InvokeExcel(() => CreateVBAMacro(vbaRequest.WorkbookName, vbaRequest.FunctionName, vbaRequest.VbaCode));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully created VBA macro in workbook {vbaRequest.WorkbookName}");

            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error creating VBA macro: {ex.Message}", 500);
            }
        }

        private async Task HandleVBAReadTool(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var vbaReadRequest = JsonSerializer.Deserialize<VBAReadRequest>(requestBody, _jsonOptions);

                if (vbaReadRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(vbaReadRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "workbookName is required");
                    return;
                }

                var result = InvokeExcel(() => ReadVBAModules(vbaReadRequest.WorkbookName));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully read VBA modules from workbook {vbaReadRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error reading VBA modules: {ex.Message}", 500);
            }
        }

        private async Task HandleVBAUpdateTool(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var vbaUpdateRequest = JsonSerializer.Deserialize<VBAUpdateRequest>(requestBody, _jsonOptions);

                if (vbaUpdateRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(vbaUpdateRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "workbookName is required");
                    return;
                }

                if (string.IsNullOrEmpty(vbaUpdateRequest.ModuleName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "moduleName is required");
                    return;
                }

                if (string.IsNullOrEmpty(vbaUpdateRequest.VbaCode))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "vbaCode is required");
                    return;
                }

                var result = InvokeExcel(() => UpdateVBAMacro(vbaUpdateRequest.WorkbookName, vbaUpdateRequest.ModuleName, vbaUpdateRequest.VbaCode));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully updated VBA module {vbaUpdateRequest.ModuleName} in workbook {vbaUpdateRequest.WorkbookName}");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error updating VBA module: {ex.Message}", 500);
            }
        }


        private async Task HandleDragFormula(string requestBody, HttpListenerResponse response)
        {
            try
            {
                var dragFormulaRequest = JsonSerializer.Deserialize<DragFormulaRequest>(requestBody, _jsonOptions);

                if (dragFormulaRequest == null)
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "Invalid request body");
                    return;
                }

                if (string.IsNullOrEmpty(dragFormulaRequest.Worksheet) || string.IsNullOrEmpty(dragFormulaRequest.SourceRange) ||
                    string.IsNullOrEmpty(dragFormulaRequest.DestinationRange) || string.IsNullOrEmpty(dragFormulaRequest.WorkbookName))
                {
                    response.StatusCode = 400;
                    await WriteJsonErrorAsync(response, "worksheet, sourceRange, destinationRange and workbookName are required");
                    return;
                }

                var result = InvokeExcel(() => DragExcelFormula(dragFormulaRequest.Worksheet, dragFormulaRequest.WorkbookName,
                    dragFormulaRequest.SourceRange, dragFormulaRequest.DestinationRange, dragFormulaRequest.FillType));

                await SendJsonResponse(response, result);

                Logger.Info($"Successfully dragged formula from {dragFormulaRequest.SourceRange} to {dragFormulaRequest.DestinationRange} in sheet {dragFormulaRequest.Worksheet}");
            }
            catch (WorkbookResolutionException wre)
            {
                await WriteJsonErrorAsync(response, wre.Message, 400, wre.ErrorCode, wre.Candidates);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await WriteJsonErrorAsync(response, $"Error dragging formula: {ex.Message}", 500);
            }
        }



        private VBAResponse CreateVBAMacro(string workbookName, string functionName, string vbaCode)
        {
            try
            {
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook workbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    workbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = wb;
                            break;
                        }
                    }
                }

                if (workbook == null)
                {
                    Logger.Info("No active workbook found");
                    application.StatusBar = "❌ No active workbook found";
                    return new VBAResponse
                    {
                        Success = false,
                        Message = "No active workbook found",
                        WorkbookName = workbookName
                    };
                }

                // Access VBA project
                var vbProject = workbook.VBProject;

                // Generate unique module name to avoid conflicts (use function name as base)
                string baseModuleName = $"{functionName}Module";
                string moduleName = baseModuleName;
                int moduleCounter = 1;

                // Check if module already exists and generate unique name
                bool moduleExists = true;
                while (moduleExists)
                {
                    moduleExists = false;
                    foreach (Microsoft.Vbe.Interop.VBComponent existingComponent in vbProject.VBComponents)
                    {
                        if (existingComponent.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                        {
                            moduleExists = true;
                            moduleName = $"{baseModuleName}{moduleCounter}";
                            moduleCounter++;
                            break;
                        }
                    }
                }

                // Create new VBA module
                var vbComponent = vbProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                vbComponent.Name = moduleName;

                // Add the provided VBA code to the module
                var codeModule = vbComponent.CodeModule;
                codeModule.AddFromString(vbaCode);

                Logger.Info($"VBA macro created successfully in module: {moduleName}");
                application.StatusBar = $"✅ VBA macro '{functionName}' created in module '{moduleName}'";

                return new VBAResponse
                {
                    Success = true,
                    Message = $"VBA macro '{functionName}' created in module '{moduleName}'",
                    WorkbookName = workbookName,
                    ModuleName = moduleName,
                    MacroName = functionName
                };
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("not trusted"))
            {
                Logger.LogException(ex);
                Logger.Info("VBA project access is not trusted - user needs to enable macro security setting");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = "❌ Enable 'Trust access to VBA project' in Excel Trust Center settings";

                return new VBAResponse
                {
                    Success = false,
                    Message = "Enable 'Trust access to VBA project' in Excel Trust Center settings",
                    WorkbookName = workbookName
                };
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A802D))
            {
                Logger.LogException(ex);
                Logger.Info("Module name conflict detected - this should not happen with unique naming");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = "❌ VBA module name conflict (0x800A802D)";

                return new VBAResponse
                {
                    Success = false,
                    Message = "VBA module name conflict detected. The module name generation failed to create a unique name.",
                    WorkbookName = workbookName
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error creating VBA macro: {ex.Message}");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = $"❌ Error creating VBA macro: {ex.Message}";

                return new VBAResponse
                {
                    Success = false,
                    Message = $"Error creating VBA macro: {ex.Message}",
                    WorkbookName = workbookName
                };
            }
        }

        private VBAUpdateResponse UpdateVBAMacro(string workbookName, string moduleName, string vbaCode)
        {
            try
            {
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook workbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    workbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = wb;
                            break;
                        }
                    }
                }

                if (workbook == null)
                {
                    Logger.Info("No workbook found");
                    return new VBAUpdateResponse
                    {
                        Success = false,
                        Message = "No workbook found",
                        WorkbookName = workbookName,
                        ModuleName = moduleName
                    };
                }

                // Access VBA project
                var vbProject = workbook.VBProject;
                Microsoft.Vbe.Interop.VBComponent targetComponent = null;

                // Find the module by name
                foreach (Microsoft.Vbe.Interop.VBComponent component in vbProject.VBComponents)
                {
                    if (component.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetComponent = component;
                        break;
                    }
                }

                if (targetComponent == null)
                {
                    Logger.Info($"Module '{moduleName}' not found in workbook '{workbookName}'");
                    return new VBAUpdateResponse
                    {
                        Success = false,
                        Message = $"Module '{moduleName}' not found in workbook '{workbookName}'",
                        WorkbookName = workbookName,
                        ModuleName = moduleName
                    };
                }

                // Store the old code
                var codeModule = targetComponent.CodeModule;
                string oldCode = codeModule.CountOfLines > 0
                    ? codeModule.Lines[1, codeModule.CountOfLines]
                    : "";

                // Clear the existing code and add the new code
                if (codeModule.CountOfLines > 0)
                {
                    codeModule.DeleteLines(1, codeModule.CountOfLines);
                }
                codeModule.AddFromString(vbaCode);

                Logger.Info($"VBA module '{moduleName}' updated successfully");
                application.StatusBar = $"✅ VBA module '{moduleName}' updated successfully";

                return new VBAUpdateResponse
                {
                    Success = true,
                    Message = $"VBA module '{moduleName}' updated successfully",
                    WorkbookName = workbookName,
                    ModuleName = moduleName,
                    OldCode = oldCode,
                    NewCode = vbaCode
                };
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("not trusted"))
            {
                Logger.LogException(ex);
                Logger.Info("VBA project access is not trusted - user needs to enable macro security setting");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = "❌ Enable 'Trust access to VBA project' in Excel Trust Center settings";

                return new VBAUpdateResponse
                {
                    Success = false,
                    Message = "Enable 'Trust access to VBA project' in Excel Trust Center settings",
                    WorkbookName = workbookName,
                    ModuleName = moduleName
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error updating VBA module: {ex.Message}");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = $"❌ Error updating VBA module: {ex.Message}";

                return new VBAUpdateResponse
                {
                    Success = false,
                    Message = $"Error updating VBA module: {ex.Message}",
                    WorkbookName = workbookName,
                    ModuleName = moduleName
                };
            }
        }

        private VBAReadResponse ReadVBAModules(string workbookName)
        {
            try
            {
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook workbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    workbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = wb;
                            break;
                        }
                    }
                }

                if (workbook == null)
                {
                    Logger.Info("No workbook found");
                    return new VBAReadResponse
                    {
                        Success = false,
                        Message = "No workbook found",
                        WorkbookName = workbookName,
                        Modules = new List<VBAModuleInfo>()
                    };
                }

                var modules = new List<VBAModuleInfo>();

                try
                {
                    // Access VBA project
                    var vbProject = workbook.VBProject;

                    // Iterate through all VBA components (modules)
                    foreach (Microsoft.Vbe.Interop.VBComponent component in vbProject.VBComponents)
                    {
                        var moduleInfo = new VBAModuleInfo
                        {
                            ModuleName = component.Name,
                            ModuleType = GetModuleTypeName(component.Type),
                            Code = component.CodeModule.CountOfLines > 0
                                ? component.CodeModule.Lines[1, component.CodeModule.CountOfLines]
                                : ""
                        };

                        modules.Add(moduleInfo);
                        Logger.Info($"Read VBA module: {moduleInfo.ModuleName} ({moduleInfo.ModuleType}) - {component.CodeModule.CountOfLines} lines");
                    }

                    Logger.Info($"Successfully read {modules.Count} VBA modules from workbook: {workbookName}");

                    var application2 = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                    application2.StatusBar = $"✅ Read {modules.Count} VBA modules";

                    return new VBAReadResponse
                    {
                        Success = true,
                        Message = $"Successfully read {modules.Count} VBA modules",
                        WorkbookName = workbookName,
                        Modules = modules
                    };
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("not trusted"))
                {
                    Logger.LogException(ex);
                    Logger.Info("VBA project access is not trusted - user needs to enable macro security setting");

                    var application2 = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                    application2.StatusBar = "❌ Enable 'Trust access to VBA project' in Excel Trust Center settings";

                    return new VBAReadResponse
                    {
                        Success = false,
                        Message = "Enable 'Trust access to VBA project' in Excel Trust Center settings",
                        WorkbookName = workbookName,
                        Modules = new List<VBAModuleInfo>()
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error reading VBA modules: {ex.Message}");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                application.StatusBar = $"❌ Error reading VBA modules: {ex.Message}";

                return new VBAReadResponse
                {
                    Success = false,
                    Message = $"Error reading VBA modules: {ex.Message}",
                    WorkbookName = workbookName,
                    Modules = new List<VBAModuleInfo>()
                };
            }
        }

        private string GetModuleTypeName(Microsoft.Vbe.Interop.vbext_ComponentType moduleType)
        {
            switch (moduleType)
            {
                case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule:
                    return "Standard Module";
                case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule:
                    return "Class Module";
                case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm:
                    return "UserForm";
                case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document:
                    return "Document Module";
                default:
                    return "Unknown";
            }
        }

        // private void FormatExcelRange(string sheetName, string workbookName, string range, FormatSettings format, bool isRevert)
        // {
        //     try
        //     {
        //         Logger.Info($"🎨 Formatting range {range} in worksheet '{sheetName}' in workbook '{workbookName}'");

        //         // Use centralized worksheet access
        //         var worksheet = ExcelSelection.GetWorksheet(sheetName, workbookName);
        //         if (worksheet == null)
        //         {
        //             throw new ArgumentException($"Worksheet '{sheetName}' in workbook '{workbookName}' not found");
        //         }

        //         // Get the target range
        //         var targetRange = worksheet.Range[range];
        //         if (targetRange == null)
        //         {
        //             throw new ArgumentException($"Invalid range '{range}'");
        //         }

        //         // Apply formatting using Microsoft Excel Interop APIs
        //         _formatWriter.ApplyFormattingToRange(targetRange, format);

        //         Logger.Info($"✅ Successfully formatted range {range} with optimized change tracking");

        //         return;
        //     }
        //     catch (Exception ex)
        //     {
        //         Logger.LogException(ex);
        //         return;
        //     }
        // }

        private DeleteColumnsResponse DeleteExcelColumns(string sheetName, string workbookName, string columns)
        {
            try
            {
                Logger.Info($"🗑️ Deleting columns {columns} in worksheet '{sheetName}' in workbook '{workbookName}'");

                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Get the columns range (e.g., "A:C" or "A,C,E")
                var columnsRange = worksheet.Range[columns];
                if (columnsRange == null)
                {
                    throw new ArgumentException($"Invalid columns specification '{columns}'");
                }

                // Count columns before deletion for response
                int columnsDeleted = columnsRange.Columns.Count;

                // Delete the columns
                columnsRange.EntireColumn.Delete();

                Logger.Info($"✅ Successfully deleted {columnsDeleted} columns");

                return new DeleteColumnsResponse
                {
                    Success = true,
                    Message = $"Successfully deleted {columnsDeleted} columns, all columns on the right of the deleted ones and their cell references have been updated",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    DeletedColumns = columns,
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new DeleteColumnsResponse
                {
                    Success = false,
                    Message = $"Failed to delete columns {columns}, cannot delete a column if the workbook is not focused",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    DeletedColumns = ""
                };
            }
        }

        private DeleteRowsResponse DeleteExcelRows(string sheetName, string workbookName, string rows)
        {
            try
            {
                Logger.Info($"🗑️ Deleting rows {rows} in worksheet '{sheetName}' in workbook '{workbookName}'");

                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Get the rows range (e.g., "1:5" or "1,3,5")
                var rowsRange = worksheet.Range[rows];
                if (rowsRange == null)
                {
                    throw new ArgumentException($"Invalid rows specification '{rows}'");
                }

                // Count rows before deletion for response
                int rowsDeleted = rowsRange.Rows.Count;

                // Delete the rows
                rowsRange.EntireRow.Delete();

                Logger.Info($"✅ Successfully deleted {rowsDeleted} rows");

                return new DeleteRowsResponse
                {
                    Success = true,
                    Message = $"Successfully deleted {rowsDeleted} rows, all rows below and their cell references have been updated",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    DeletedRows = rows
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new DeleteRowsResponse
                {
                    Success = false,
                    Message = $"Failed to delete rows {rows}, cannot delete a row if the workbook is not focused",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    DeletedRows = ""
                };
            }
        }

        private InsertRowsResponse InsertExcelRows(string sheetName, string workbookName, string rowRange)
        {
            try
            {
                Logger.Info($"🔄 Inserting rows at range {rowRange} in worksheet '{sheetName}'");

                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Validate row range format
                if (string.IsNullOrEmpty(rowRange) || !rowRange.Contains(":"))
                {
                    throw new ArgumentException($"Invalid row range format '{rowRange}'. Expected format like '1:1' or '2:5'");
                }

                var targetRange = worksheet.Range[rowRange];
                targetRange.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);

                Logger.Info($"✅ Successfully inserted {rowRange} rows");

                return new InsertRowsResponse
                {
                    Success = true,
                    Message = $"Successfully inserted {rowRange} rows, all rows below and their cell references have been updated",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    RowsInserted = rowRange
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new InsertRowsResponse
                {
                    Success = false,
                    Message = $"Failed to insert rows at range {rowRange}: {ex.Message}",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    RowsInserted = ""
                };
            }
        }

        private InsertColumnsResponse InsertExcelColumns(string sheetName, string workbookName, string columnRange)
        {
            try
            {
                Logger.Info($"🔄 Inserting columns at range {columnRange} in worksheet '{sheetName}'");

                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Validate and parse the column range
                if (string.IsNullOrEmpty(columnRange) || !columnRange.Contains(":"))
                {
                    Logger.Info($"Validate and parse ❌ worksheet: {sheetName}, columnRange: {columnRange}, workbookName: {workbookName}");
                    throw new ArgumentException($"Invalid column range format '{columnRange}'. Expected format like 'A:A' or 'B:D'");
                }

                // Get the range and insert columns
                var targetRange = worksheet.Range[columnRange];
                if (targetRange == null)
                {
                    throw new ArgumentException($"Invalid column range specification '{columnRange}'");
                }

                // Insert the columns using Excel's range insertion
                targetRange.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight);
                Logger.Info($"✅ Successfully inserted {columnRange} columns");

                return new InsertColumnsResponse
                {
                    Success = true,
                    Message = $"Successfully inserted {columnRange} columns, all columns on the right of the inserted ones and their cell references have been updated",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    ColumnsInserted = columnRange
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new InsertColumnsResponse
                {
                    Success = false,
                    Message = $"Failed to insert columns at range {columnRange}: {ex.Message}",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    ColumnsInserted = ""
                };
            }
        }

        private AddSheetsResponse AddExcelSheets(string workbookName, string[] sheetNames)
        {
            try
            {
                Logger.Info($"📋 Adding {sheetNames.Length} sheets to workbook '{workbookName}'");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook targetWorkbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    targetWorkbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWorkbook = wb;
                            break;
                        }
                    }
                }

                if (targetWorkbook == null)
                {
                    throw new ArgumentException($"Workbook '{workbookName}' not found");
                }

                var addedSheets = new List<string>();

                foreach (var sheetName in sheetNames)
                {
                    try
                    {
                        // Check if sheet already exists
                        bool sheetExists = false;
                        foreach (Worksheet existingSheet in targetWorkbook.Worksheets)
                        {
                            if (existingSheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                sheetExists = true;
                                break;
                            }
                        }

                        if (!sheetExists)
                        {
                            // Add new worksheet
                            var newSheet = (Worksheet)targetWorkbook.Worksheets.Add(After: targetWorkbook.Worksheets[targetWorkbook.Worksheets.Count]);
                            newSheet.Name = sheetName;
                            addedSheets.Add(sheetName);
                            Logger.Info($"✅ Added sheet: '{sheetName}'");
                        }
                        else
                        {
                            Logger.Info($"⚠️ Sheet '{sheetName}' already exists, skipping");
                        }
                    }
                    catch (Exception sheetEx)
                    {
                        Logger.Info($"❌ Error adding sheet '{sheetName}': {sheetEx.Message}");
                    }
                }

                Logger.Info($"✅ Successfully added {addedSheets.Count}/{sheetNames.Length} sheets");

                return new AddSheetsResponse
                {
                    Success = true,
                    Message = $"Successfully added {addedSheets.Count} sheets",
                    WorkbookName = workbookName,
                    Worksheet = "", // Not applicable for this operation
                    SheetsAdded = addedSheets.ToArray()
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new AddSheetsResponse
                {
                    Success = false,
                    Message = $"Failed to add sheets: {ex.Message}",
                    WorkbookName = workbookName,
                    Worksheet = "",
                    SheetsAdded = new string[0]
                };
            }
        }

        private RemoveSheetsResponse RemoveExcelSheets(string workbookName, string[] sheetNames)
        {
            Logger.Info($"🗑️ Attempting to remove {sheetNames.Length} sheets from workbook: {workbookName}");

            try
            {
                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook targetWorkbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    targetWorkbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWorkbook = wb;
                            break;
                        }
                    }
                }

                if (targetWorkbook == null)
                {
                    Logger.Info($"❌ Workbook '{workbookName}' not found");
                    return new RemoveSheetsResponse
                    {
                        Success = false,
                        Message = $"Workbook '{workbookName}' not found",
                        WorkbookName = workbookName,
                        Worksheet = "",
                        SheetsRemoved = new string[0]
                    };
                }

                var removedSheets = new List<string>();

                foreach (string sheetName in sheetNames)
                {
                    try
                    {
                        // Check if sheet exists
                        Worksheet sheetToRemove = null;
                        bool sheetExists = false;

                        foreach (Worksheet sheet in targetWorkbook.Worksheets)
                        {
                            if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                sheetToRemove = sheet;
                                sheetExists = true;
                                break;
                            }
                        }

                        if (sheetExists && sheetToRemove != null)
                        {
                            // Check if this is the last sheet (Excel requires at least one sheet)
                            if (targetWorkbook.Worksheets.Count <= 1)
                            {
                                Logger.Info($"⚠️ Cannot remove sheet '{sheetName}' - workbook must have at least one sheet");
                                continue;
                            }

                            // Remove the sheet
                            sheetToRemove.Delete();
                            removedSheets.Add(sheetName);
                            Logger.Info($"✅ Removed sheet: '{sheetName}'");
                        }
                        else
                        {
                            Logger.Info($"⚠️ Sheet '{sheetName}' does not exist, skipping");
                        }
                    }
                    catch (Exception sheetEx)
                    {
                        Logger.Info($"❌ Error removing sheet '{sheetName}': {sheetEx.Message}");
                    }
                }

                Logger.Info($"✅ Successfully removed {removedSheets.Count}/{sheetNames.Length} sheets");

                return new RemoveSheetsResponse
                {
                    Success = true,
                    Message = $"Successfully removed {removedSheets.Count} sheets",
                    WorkbookName = workbookName,
                    Worksheet = "", // Not applicable for this operation
                    SheetsRemoved = removedSheets.ToArray()
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new RemoveSheetsResponse
                {
                    Success = false,
                    Message = $"Failed to remove sheets: {ex.Message}",
                    WorkbookName = workbookName,
                    Worksheet = "",
                    SheetsRemoved = new string[0]
                };
            }
        }




        private DragFormulaResponse DragExcelFormula(string sheetName, string workbookName, string sourceRange, string destinationRange, string fillType)
        {
            try
            {
                Logger.Info($"🔄 Dragging formula from {sourceRange} to {destinationRange} in worksheet '{sheetName}' in workbook '{workbookName}'");

                var worksheet = GetWorksheetFromActiveWorkbook(sheetName, workbookName);

                // Get source and destination ranges
                var source = worksheet.Range[sourceRange];
                var destination = worksheet.Range[destinationRange];

                if (source == null || destination == null)
                {
                    throw new ArgumentException($"Invalid source range '{sourceRange}' or destination range '{destinationRange}'");
                }

                // Map fill type to Excel enum
                XlAutoFillType autoFillType = XlAutoFillType.xlFillDefault;
                switch (fillType?.ToLower())
                {
                    case "series":
                        autoFillType = XlAutoFillType.xlFillSeries;
                        break;
                    case "formats":
                        autoFillType = XlAutoFillType.xlFillFormats;
                        break;
                    case "values":
                        autoFillType = XlAutoFillType.xlFillValues;
                        break;
                    default:
                        autoFillType = XlAutoFillType.xlFillDefault;
                        break;
                }

                // Perform AutoFill
                source.AutoFill(destination, autoFillType);

                // Count filled cells
                int cellsFilled = destination.Cells.Count;

                Logger.Info($"✅ Successfully filled {cellsFilled} cells using AutoFill");

                return new DragFormulaResponse
                {
                    Success = true,
                    Message = $"Successfully dragged formula from {sourceRange} to {destinationRange}",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    SourceRange = sourceRange,
                    DestinationRange = destinationRange,
                    CellsFilled = cellsFilled
                };
            }
            catch (WorkbookResolutionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new DragFormulaResponse
                {
                    Success = false,
                    Message = $"Failed to drag formula: {ex.Message}",
                    Worksheet = sheetName,
                    WorkbookName = workbookName,
                    SourceRange = sourceRange,
                    DestinationRange = destinationRange,
                    CellsFilled = 0
                };
            }
        }

        private CreateChartResponse CreateExcelChart(string workbookName, string chartType, string dataRange,
            string title, string xAxisTitle, string yAxisTitle, string chartLocation)
        {
            try
            {
                Logger.Info($"📊 Creating {chartType} chart in workbook '{workbookName}' from range '{dataRange}'");

                var application = Container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                Workbook targetWorkbook = null;

                if (string.IsNullOrEmpty(workbookName))
                {
                    targetWorkbook = application.ActiveWorkbook;
                }
                else
                {
                    foreach (Workbook wb in application.Workbooks)
                    {
                        if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                            wb.FullName.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetWorkbook = wb;
                            break;
                        }
                    }
                }

                if (targetWorkbook == null)
                {
                    throw new ArgumentException($"Workbook '{workbookName}' not found");
                }

                // Get the active worksheet (or first worksheet if no active sheet)
                Worksheet targetWorksheet = (Worksheet)targetWorkbook.ActiveSheet;
                if (targetWorksheet == null && targetWorkbook.Worksheets.Count > 0)
                {
                    targetWorksheet = (Worksheet)targetWorkbook.Worksheets[1];
                }

                if (targetWorksheet == null)
                {
                    throw new ArgumentException("No worksheets found in workbook");
                }

                // Get the data range
                Range sourceRange = targetWorksheet.Range[dataRange];
                if (sourceRange == null)
                {
                    throw new ArgumentException($"Invalid data range: {dataRange}");
                }

                // Determine chart type for Excel
                XlChartType excelChartType;
                switch (chartType.ToLower())
                {
                    case ChartTypes.Line:
                        excelChartType = XlChartType.xlLine;
                        break;
                    case ChartTypes.Histogram:
                        excelChartType = XlChartType.xlColumnClustered; // Using column chart for histogram
                        break;
                    case ChartTypes.Pie:
                        excelChartType = XlChartType.xlPie;
                        break;
                    default:
                        throw new ArgumentException($"Unsupported chart type: {chartType}");
                }

                // Create the chart
                ChartObject chartObject;
                Chart chart;
                string finalChartLocation = chartLocation ?? "A1";

                if (string.IsNullOrEmpty(chartLocation) || chartLocation.ToLower() != ChartLocations.NewSheet)
                {
                    // Create chart on existing sheet
                    var chartObjects = (ChartObjects)targetWorksheet.ChartObjects();

                    // Determine position for chart (default to A1 area)
                    double left = 50;
                    double top = 50;
                    double width = 400;
                    double height = 300;

                    if (!string.IsNullOrEmpty(chartLocation) && chartLocation.ToLower() != ChartLocations.NewSheet)
                    {
                        try
                        {
                            var locationRange = targetWorksheet.Range[chartLocation];
                            left = locationRange.Left;
                            top = locationRange.Top;
                        }
                        catch
                        {
                            Logger.Info($"⚠️ Could not parse chart location '{chartLocation}', using default position");
                        }
                    }

                    chartObject = chartObjects.Add(left, top, width, height);
                    chart = chartObject.Chart;
                    finalChartLocation = targetWorksheet.Name;
                }
                else
                {
                    // Create chart on new sheet
                    chart = (Chart)targetWorkbook.Charts.Add();
                    finalChartLocation = ChartLocations.NewSheet;
                }

                // Set the data source and chart type
                if (chartType.ToLower() == ChartTypes.Pie)
                {
                    // For pie charts, data is typically organized in rows (categories in first row, values in second row)
                    chart.SetSourceData(sourceRange, XlRowCol.xlRows);
                }
                else
                {
                    // For line plots and histograms, use columns to support multiple series
                    // This allows for multiple data series (multiple lines in line charts, grouped bars in histograms)
                    // Data structure expected:
                    // Row 1: Category labels (X-axis values)
                    // Row 2+: Series data (each row becomes a separate line/bar group)
                    chart.SetSourceData(sourceRange, XlRowCol.xlColumns);

                    // For histograms, ensure we're using the right chart type for multiple series
                    if (chartType.ToLower() == ChartTypes.Histogram && sourceRange.Rows.Count > 2)
                    {
                        // Use clustered column chart for side-by-side comparison when multiple series
                        chart.ChartType = XlChartType.xlColumnClustered;
                        Logger.Info($"📊 Using clustered column chart for multi-series histogram with {sourceRange.Rows.Count - 1} data series");
                    }

                    // For line charts with multiple series, Excel will automatically create multiple lines
                    if (chartType.ToLower() == ChartTypes.Line && sourceRange.Rows.Count > 2)
                    {
                        Logger.Info($"📊 Creating multi-series line chart with {sourceRange.Rows.Count - 1} data series");
                    }
                }
                chart.ChartType = excelChartType;

                // Set chart title
                if (!string.IsNullOrEmpty(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }

                // Set axis titles (only for non-pie charts)
                if (chartType.ToLower() != ChartTypes.Pie)
                {
                    if (!string.IsNullOrEmpty(xAxisTitle))
                    {
                        chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).HasTitle = true;
                        chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary).AxisTitle.Text = xAxisTitle;
                    }

                    if (!string.IsNullOrEmpty(yAxisTitle))
                    {
                        chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).HasTitle = true;
                        chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary).AxisTitle.Text = yAxisTitle;
                    }
                }
                else
                {
                    // For pie charts, show data labels if axis titles are provided (as a way to enhance the chart)
                    if (!string.IsNullOrEmpty(xAxisTitle) || !string.IsNullOrEmpty(yAxisTitle))
                    {
                        try
                        {
                            chart.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowPercent);
                        }
                        catch
                        {
                            // Fallback if data labels fail
                            Logger.Info("⚠️ Could not apply data labels to pie chart");
                        }
                    }
                }

                string chartName = chart.Name ?? $"Chart_{DateTime.Now:yyyyMMdd_HHmmss}";

                Logger.Info($"✅ Successfully created {chartType} chart: '{chartName}'");

                return new CreateChartResponse
                {
                    Success = true,
                    Message = $"Successfully created {chartType} chart",
                    WorkbookName = workbookName,
                    Worksheet = targetWorksheet.Name,
                    ChartName = chartName,
                    ChartLocation = finalChartLocation
                };
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                return new CreateChartResponse
                {
                    Success = false,
                    Message = $"Failed to create chart: {ex.Message}",
                    WorkbookName = workbookName,
                    Worksheet = "",
                    ChartName = "",
                    ChartLocation = ""
                };
            }
        }


    }

    public class WorkbookMetadataResponse
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("workbookFullName")]
        public string WorkbookFullName { get; set; }

        [JsonPropertyName("activeSheet")]
        public string ActiveSheet { get; set; }

        [JsonPropertyName("usedRange")]
        public string UsedRange { get; set; }

        [JsonPropertyName("selectedRange")]
        public string SelectedRange { get; set; }

        [JsonPropertyName("allSheets")]
        public List<string> AllSheets { get; set; }

        [JsonPropertyName("sheets")]
        public List<SheetMetadata> Sheets { get; set; }

        [JsonPropertyName("languageCode")]
        public string LanguageCode { get; set; }

        [JsonPropertyName("dateLanguage")]
        public string DateLanguage { get; set; }

        [JsonPropertyName("listSeparator")]
        public string ListSeparator { get; set; }

        [JsonPropertyName("decimalSeparator")]
        public string DecimalSeparator { get; set; }

        [JsonPropertyName("thousandsSeparator")]
        public string ThousandsSeparator { get; set; }
    }

    public class SheetMetadata
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("usedRangeAddress")]
        public string UsedRangeAddress { get; set; }
    }

    public class WorkbookCandidate
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("workbookFullName")]
        public string WorkbookFullName { get; set; }
    }

    public class WorkbookResolutionException : Exception
    {
        public WorkbookResolutionException(string message, string errorCode, List<WorkbookCandidate> candidates = null)
            : base(message)
        {
            ErrorCode = errorCode;
            Candidates = candidates ?? new List<WorkbookCandidate>();
        }

        public string ErrorCode { get; }
        public List<WorkbookCandidate> Candidates { get; }
    }

    public class FormatRangeRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("format")]
        public FormatSettings Format { get; set; }

    }

    public class ReadFormatBatchOperation
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("propertiesToRead")]
        public List<string> PropertiesToRead { get; set; }
    }

    public class ReadFormatBatchRequest
    {
        [JsonPropertyName("operations")]
        public List<ReadFormatBatchOperation> Operations { get; set; }
    }

    public class ReadFormatBatchResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("results")]
        public List<ReadFormatResponse> Results { get; set; }
    }

    public class WriteFormatBatchRequest
    {
        [JsonPropertyName("operations")]
        public List<FormatRangeRequest> Operations { get; set; }

        [JsonPropertyName("readOldFormats")]
        public bool? ReadOldFormats { get; set; }

        [JsonPropertyName("collapseReadRanges")]
        public bool? CollapseReadRanges { get; set; }
    }

    public class WriteFormatBatchResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("results")]
        public List<FormatRangeResponse> Results { get; set; }
    }

    // Keep RangeFormatPair for backwards compatibility with revert operations
    public class RangeFormatPair
    {
        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("format")]
        public FormatSettings Format { get; set; }

        public void Deconstruct(out string range, out FormatSettings format)
        {
            range = Range;
            format = Format;
        }
    }
    public class DeleteColumnsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("columnRange")]
        public string ColumnRange { get; set; }
    }

    public class DeleteColumnsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("deletedColumns")]
        public string DeletedColumns { get; set; }
    }

    public class DeleteRowsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("rowRange")]
        public string RowRange { get; set; }
    }

    public class DeleteRowsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("deletedRows")]
        public string DeletedRows { get; set; }
    }

    public class InsertRowsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("rowRange")]
        public string RowRange { get; set; }
    }

    public class InsertRowsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("rowsInserted")]
        public string RowsInserted { get; set; }
    }

    public class InsertColumnsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("columnRange")]
        public string ColumnRange { get; set; }
    }

    public class InsertColumnsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("columnsInserted")]
        public string ColumnsInserted { get; set; }
    }

    public class CopyPasteRequest
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sourceWorkbookName")]
        public string SourceWorkbookName { get; set; }

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

        [JsonPropertyName("operation")]
        public string Operation { get; set; }

        [JsonPropertyName("skipBlanks")]
        public bool? SkipBlanks { get; set; }

        [JsonPropertyName("transpose")]
        public bool? Transpose { get; set; }

        [JsonPropertyName("insertMode")]
        public string InsertMode { get; set; }

        [JsonPropertyName("includeColumnWidths")]
        public bool? IncludeColumnWidths { get; set; }
    }


    public class DeleteCellsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("shiftDirection")]
        public string ShiftDirection { get; set; }
    }

    public class DeleteCellsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("range")]
        public string Range { get; set; }

        [JsonPropertyName("shiftDirection")]
        public string ShiftDirection { get; set; }
    }


    public class AddSheetsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sheetNames")]
        public string[] SheetNames { get; set; }
    }

    public class AddSheetsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sheetsAdded")]
        public string[] SheetsAdded { get; set; }
    }

    public class CreateChartRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("chartType")]
        public string ChartType { get; set; }

        [JsonPropertyName("dataRange")]
        public string DataRange { get; set; }

        [JsonPropertyName("title")]
        public string Title { get; set; }

        [JsonPropertyName("xAxisTitle")]
        public string XAxisTitle { get; set; }

        [JsonPropertyName("yAxisTitle")]
        public string YAxisTitle { get; set; }

        [JsonPropertyName("chartLocation")]
        public string ChartLocation { get; set; }
    }

    public class CreateChartResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("chartName")]
        public string ChartName { get; set; }

        [JsonPropertyName("chartLocation")]
        public string ChartLocation { get; set; }
    }

    public class VBARequest
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("functionName")]
        public string FunctionName { get; set; }

        [JsonPropertyName("vbaCode")]
        public string VbaCode { get; set; }
    }

    public class VBAResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; }

        [JsonPropertyName("macroName")]
        public string MacroName { get; set; }
    }

    public class VBAReadRequest
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }
    }

    public class VBAModuleInfo
    {
        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; }

        [JsonPropertyName("moduleType")]
        public string ModuleType { get; set; }

        [JsonPropertyName("code")]
        public string Code { get; set; }
    }

    public class VBAReadResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("modules")]
        public List<VBAModuleInfo> Modules { get; set; }
    }

    public class VBAUpdateRequest
    {
        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; }

        [JsonPropertyName("vbaCode")]
        public string VbaCode { get; set; }
    }

    public class VBAUpdateResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("moduleName")]
        public string ModuleName { get; set; }

        [JsonPropertyName("oldCode")]
        public string OldCode { get; set; }

        [JsonPropertyName("newCode")]
        public string NewCode { get; set; }
    }
    public class DragFormulaRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sourceRange")]
        public string SourceRange { get; set; }

        [JsonPropertyName("destinationRange")]
        public string DestinationRange { get; set; }

        [JsonPropertyName("fillType")]
        public string FillType { get; set; }
    }

    public class DragFormulaResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sourceRange")]
        public string SourceRange { get; set; }

        [JsonPropertyName("destinationRange")]
        public string DestinationRange { get; set; }

        [JsonPropertyName("cellsFilled")]
        public int CellsFilled { get; set; }
    }

    public class RemoveSheetsRequest
    {
        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sheetNames")]
        public string[] SheetNames { get; set; }
    }

    public class RemoveSheetsResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("message")]
        public string Message { get; set; }

        [JsonPropertyName("worksheet")]
        public string Worksheet { get; set; }

        [JsonPropertyName("workbookName")]
        public string WorkbookName { get; set; }

        [JsonPropertyName("sheetsRemoved")]
        public string[] SheetsRemoved { get; set; }
    }
}
