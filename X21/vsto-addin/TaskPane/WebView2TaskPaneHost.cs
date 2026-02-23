using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Office.Tools;
using X21.Common.Data;
using X21.Excel;
using X21.Logging;
using X21.Services;
using System.Text.Json.Serialization;
using System.IO;
using X21.Utils;
using System.Threading;
using System.Net.Http;
using System.Text;

namespace X21.TaskPane
{
    public partial class WebView2TaskPaneHost : UserControl
    {
        private readonly Container _container;
        private WebView2 webView;
        private bool _isInitialized = false;
        private bool _frontEndIsReady = false;
        private bool _pendingFocusRequest = false;
        private bool _webViewHasFocus = false; // Track WebView2 focus state
        private DateTime _lastFocusLossTime = DateTime.MinValue; // Track when WebView lost focus
        private bool _intentionalFocusRequest = false; // Track when focus is intentionally set via keyboard shortcut
        private bool _fileDialogOpen = false; // Track when file picker dialog is open

        public WebView2TaskPaneHost(Container container)
        {
            _container = container;
            InitializeComponent();
            InitializeWebView();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // WebView2 control
            this.webView = new WebView2()
            {
                Dock = DockStyle.Fill
            };

            this.Controls.Add(this.webView);

            this.ResumeLayout(false);
        }

        private async void InitializeWebView()
        {
            try
            {
                var options = new CoreWebView2EnvironmentOptions()
                {
                    AdditionalBrowserArguments = "--disable-features=msWebOOUI,msPdfOOUI --disable-web-security"
                };

                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "X21", "WebView2");

                // Ensure the directory exists and is writable
                if (!Directory.Exists(userDataFolder))
                {
                    Directory.CreateDirectory(userDataFolder);
                }

                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                await webView.EnsureCoreWebView2Async(env);

                // Configure WebView2 settings
                webView.CoreWebView2.Settings.IsScriptEnabled = true;
                webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = true;
                webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
                webView.CoreWebView2.Settings.AreHostObjectsAllowed = false;
                webView.CoreWebView2.Settings.AreDevToolsEnabled = true;

                // Set up message handling
                webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // Simple focus handling and keyboard shortcut to recreate the pane (Ctrl+Alt+R)
                webView.KeyDown += (s, e) =>
                {
                    if (e.Control && e.Alt && e.KeyCode == Keys.R)
                    {
                        RecreateTaskPane();
                        e.Handled = true;
                    }
                    else
                    {
                        e.Handled = false; // Don't block other keyboard events
                    }
                };
                webView.KeyUp += (s, e) => e.Handled = false;   // Don't block keyboard events

                // Navigate to our React app
                if (EnvironmentHelper.IsDebugMode())
                {
                    webView.CoreWebView2.Navigate("http://localhost:5173");
                }
                else
                {
                    // Published builds: serve from local WebAssets
                    // Use PathResolver to handle deployment scenarios
                    string localPath = PathResolver.GetFilePath("TaskPane\\WebAssets\\index.html");

                    var url = $"file:///{localPath.Replace('\\', '/')}";
                    Logger.Info($"Navigating to local path: {url}");
                    if (System.IO.File.Exists(localPath))
                    {
                        webView.CoreWebView2.Navigate(url);
                    }
                    else
                    {
                        // Fallback: show error message
                        webView.NavigateToString(@"
                            <html>
                                <body style='font-family: Arial; padding: 20px; text-align: center;'>
                                    <h2>React App Not Found</h2>
                                    <p>WebView2 is working, but React build files are missing.</p>
                                    <p>Expected path: " + localPath + @"</p>
                                    <p>Run: <code>npm run build</code> in the web-ui folder</p>
                                </body>
                            </html>");
                    }
                }

                _isInitialized = true;
                Logger.Info("WebView2 initialized successfully");
            }
            catch (Exception ex)
            {
                Logger.Info($"Failed to initialize WebView2: {ex.Message}");
            }
        }

        private async Task HandleNavigateToRange(string messageId, string range, string workbookName)
        {
            await Task.Run(() =>
            {
                try
                {
                    Logger.Info($"Navigating to range: {range} (workbook: {workbookName ?? "(active)"} )");

                    if (_container == null)
                    {
                        Logger.Info("Container not initialized");
                        return;
                    }

                    var excelSelection = _container.Resolve<ExcelSelection>();
                    if (excelSelection == null)
                    {
                        Logger.Info("Excel service not available");
                        return;
                    }

                    // Parse the range to handle sheet references
                    string targetRange = range;
                    string sheetName = null;
                    string workbookFromRange = null;

                    var parts = range.Split('!');
                    if (parts.Length >= 2)
                    {
                        // Format: [workbook!]sheet!range
                        if (parts.Length == 3)
                        {
                            workbookFromRange = parts[0].Trim('\'');
                            sheetName = parts[1].Trim('\'');
                            targetRange = parts[2];
                        }
                        else
                        {
                            sheetName = parts[0].Trim('\'');
                            targetRange = parts[1];
                        }
                    }

                    var resolvedWorkbook = string.IsNullOrWhiteSpace(workbookName)
                        ? workbookFromRange
                        : workbookName;

                    // Switch to target sheet if specified
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        // Reuse existing GetWorksheet method, allowing workbook resolution
                        var targetWorksheet = excelSelection.GetWorksheet(sheetName, resolvedWorkbook);

                        if (targetWorksheet != null)
                        {
                            targetWorksheet.Activate();
                            Logger.Info($"Activated worksheet: '{sheetName}' (workbook: '{resolvedWorkbook ?? "active"}')");
                        }
                        else
                        {
                            Logger.Info($"Worksheet '{sheetName}' not found in workbook '{resolvedWorkbook ?? "active"}'");
                            return;
                        }
                    }

                    Logger.Info($"Focusing on range: '{targetRange}' in sheet: '{sheetName ?? "current"}'");

                    // Focus the range using existing functionality
                    excelSelection.FocusCell(targetRange);

                    Logger.Info($"Successfully navigated to range: {range} (workbook: {resolvedWorkbook ?? "active"})");
                }
                catch (Exception ex)
                {
                    Logger.Info($"Error navigating to range {range}: {ex.Message}");
                }
            });
        }

        private async Task HandleGenerateChangelog(string messageId, JsonElement payload)
        {
            try
            {
                string comparisonFilePath = null;
                if (payload.ValueKind != JsonValueKind.Undefined &&
                    payload.TryGetProperty("comparisonFilePath", out var filePathProp))
                {
                    comparisonFilePath = filePathProp.GetString();
                    Logger.Info($"Manual changelog generation requested with comparison file: {comparisonFilePath}");
                }
                else
                {
                    Logger.Info("Manual changelog generation requested");
                }

                var application = _container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                var workbook = application?.ActiveWorkbook;

                if (workbook == null)
                {
                    Logger.Info("No active workbook");
                    await SendResponse(messageId, new { success = false, error = "No active workbook" });
                    return;
                }

                var changeTracker = _container.Resolve<X21.Services.WorkbookChangeTrackerService>();
                if (changeTracker == null)
                {
                    Logger.Info("WorkbookChangeTrackerService not available");
                    await SendResponse(messageId, new { success = false, error = "Change tracker service not available" });
                    return;
                }

                // Generate changelog in background
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await changeTracker.GenerateChangelog(workbook, comparisonFilePath);
                    }
                    catch (Exception asyncEx)
                    {
                        Logger.LogException(asyncEx);
                        Logger.Info($"Error generating changelog: {asyncEx.Message}");
                    }
                });

                // Return immediately
                await SendResponse(messageId, new { success = true, message = "Changelog generation started" });
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                await SendResponse(messageId, new { success = false, error = ex.Message });
            }
        }

        private async void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var messageJson = e.WebMessageAsJson;
                // Only log message type, not full JSON content

                var actualJsonString = JsonSerializer.Deserialize<string>(messageJson);

                using var doc = JsonDocument.Parse(actualJsonString);
                var root = doc.RootElement;

                // Now try class deserialization with options
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    AllowTrailingCommas = true
                };

                var message = JsonSerializer.Deserialize<WebViewMessage>(actualJsonString, options);
                var requestId = message.Payload?.RequestId?.ToString();
                var conversationId = message.Payload?.ConversationId;

                var excelSelection = _container.Resolve<ExcelSelection>();
                if (excelSelection == null)
                {
                    Logger.Info("Excel service not available");
                    return;
                }

                Logger.Info($"🎯 {message.Type} (ID: {requestId})");

                // Handle the message
                switch (message.Type)
                {
                    case "FrontEndIsReady":
                        _frontEndIsReady = true;
                        Logger.Info("FrontEndIsReady received from React");

                        // Send current selection when frontend is ready
                        SendCurrentSelection();
                        SendWorkbookContext();

                        // Send user ID for PostHog analytics
                        SendEventToReact("userIdReady", new { userId = Utils.UserUtils.GetUserId() });

                        if (_pendingFocusRequest)
                        {
                            _pendingFocusRequest = false;
                            FocusTextInput();
                        }
                        break;
                    case "userEmailReady":
                        if (root.TryGetProperty("payload", out var emailPayload) &&
                            emailPayload.TryGetProperty("email", out var emailProp))
                        {
                            var email = emailProp.GetString();
                            Utils.UserUtils.SetUserEmail(email);
                            Logger.Info($"User email received from React: {email ?? "null"}");
                        }
                        break;
                    case "getUserIdentity":
                        await HandleGetUserIdentity(requestId);
                        break;
                    case "getWorkbookName":
                        await HandleGetWorkbookName(requestId);
                        break;
                    case "getWorkbookPath":
                        await HandleGetWorkbookPath(requestId);
                        break;
                    case "getWorksheetNames":
                        await HandleGetWorksheetNames(requestId);
                        break;
                    case "getWebSocketUrl":
                        await HandleGetWebSocketUrl(requestId);
                        break;
                    case "pickFolder":
                        Logger.Info($"[pickFolder] Received request (ID: {requestId})");
                        await HandlePickFolder(requestId, root.TryGetProperty("payload", out var pickPayload) ? pickPayload : default);
                        break;
                    case "pickFile":
                        Logger.Info($"[pickFile] Received request (ID: {requestId})");
                        await HandlePickFile(requestId, root.TryGetProperty("payload", out var filePayload) ? filePayload : default);
                        break;
                    case "openWorkbook":
                        Logger.Info($"[openWorkbook] Received request (ID: {requestId})");
                        await HandleOpenWorkbookFromWebView(requestId, root.TryGetProperty("payload", out var openPayload) ? openPayload : default);
                        break;
                    case "getSlashCommandsFromSheet":
                        await HandleGetSlashCommandsFromSheet(requestId);
                        break;
                    case "navigateToRange":
                        var rangePayload = root.GetProperty("payload");
                        string range = rangePayload.GetProperty("range").GetString();
                        string workbookName = rangePayload.TryGetProperty("workbookName", out var wbProp) ? wbProp.GetString() : null;
                        await HandleNavigateToRange(requestId, range, workbookName);
                        break;
                    case "focusHost":
                        HandleFocusHost();
                        break;
                    case "fileDialogOpening":
                        _fileDialogOpen = true;
                        Logger.Info("File dialog opening - preventing auto-focus to Excel");
                        break;
                    case "webViewFocusChanged":
                        try
                        {
                            // Access the hasFocus property through the JSON document
                            if (root.TryGetProperty("payload", out var payload) &&
                                payload.TryGetProperty("hasFocus", out var hasFocusElement))
                            {
                                bool newFocusState = hasFocusElement.GetBoolean();

                                // If WebView gained focus, clear the intentional focus flag
                                if (newFocusState && !_webViewHasFocus)
                                {
                                    Logger.Info("WebView gained focus, clearing intentional focus flag");
                                    _intentionalFocusRequest = false;

                                    // If file dialog was open, it has now closed - restore focus to textarea
                                    if (_fileDialogOpen)
                                    {
                                        Logger.Info("File dialog closed (focus returned to WebView) - restoring focus to textarea");
                                        _fileDialogOpen = false;

                                        // WebView already has focus (we're in the gained focus branch)
                                        // Just need to focus the textarea in React
                                        SendEventToReact("focusTextInput", new { });
                                    }
                                }

                                // Track focus loss timing for smart restoration
                                if (_webViewHasFocus && !newFocusState)
                                {
                                    _lastFocusLossTime = DateTime.Now;
                                    Logger.Info($"WebView lost focus at {_lastFocusLossTime:HH:mm:ss.fff}");

                                    // Only auto-focus Excel if:
                                    // 1. This was NOT an intentional focus request AND
                                    // 2. The file dialog is NOT open
                                    if (_fileDialogOpen)
                                    {
                                        Logger.Info("Skipping auto-focus to Excel (file dialog is open)");
                                    }
                                    else if (_intentionalFocusRequest)
                                    {
                                        Logger.Info("Skipping auto-focus to Excel (intentional focus request in progress)");
                                    }
                                    else
                                    {
                                        HandleFocusHost();
                                        Logger.Info("Automatically focusing Excel after WebView focus loss");
                                    }
                                }

                                _webViewHasFocus = newFocusState;
                                Logger.Info($"WebView focus changed: {_webViewHasFocus}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Info($"Error handling webViewFocusChanged: {ex.Message}");
                        }
                        break;
                    case "generateChangelog":
                        var changelogPayload = root.TryGetProperty("payload", out var changelogPayloadProp) ? changelogPayloadProp : default;
                        await HandleGenerateChangelog(requestId, changelogPayload);
                        break;
                    default:
                        Logger.Info($"❓ Unknown message: {message.Type}");
                        await SendResponse(requestId, new { error = "Unknown message type" });
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error: {ex.Message}");
                await SendResponse("Something went wrong", new { error = ex.Message });
            }
        }

        #region Message Handlers

        private void HandleFocusHost()
        {
            var app = _container.Resolve<Microsoft.Office.Interop.Excel.Application>();
            WindowUtils.FocusExcelWindow(app);
        }

        private async Task HandleGetUserIdentity(string messageId)
        {
            try
            {
                var userId = UserUtils.GetUserId();
                var userName = UserUtils.GetUserId();

                await SendResponse(messageId, new { success = true, userId, userName });
            }
            catch (Exception ex)
            {
                await SendResponse(messageId, new { success = false, error = ex.Message });
            }
        }

        private async Task HandleGetWorkbookName(string messageId)
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var workbookName = excelSelection.GetWorkbookName();

                await SendResponse(messageId, workbookName);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting workbook name: {ex.Message}");
                await SendResponse(messageId, string.Empty); // Explicitly indicate missing workbook
            }
        }

        private async Task HandleGetWorkbookPath(string messageId)
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var workbookPath = excelSelection.GetWorkbookPath();

                await SendResponse(messageId, workbookPath);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting workbook path: {ex.Message}");
                await SendResponse(messageId, string.Empty);
            }
        }

        private async Task HandleGetWorksheetNames(string messageId)
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var sheetNames = excelSelection?.GetAllSheetNames() ?? new List<string>();

                await SendResponse(messageId, sheetNames);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting worksheet names: {ex.Message}");
                await SendResponse(messageId, new List<string>());
            }
        }

        private async Task HandleGetWebSocketUrl(string messageId)
        {
            try
            {
                // If the backend isn't running (e.g., the Deno server crashed), attempt to restart it
                Globals.ThisAddIn?.EnsureBackendServerRunning("websocket url request");

                var websocketUrl = BackendConfigService.Instance.WebSocketUrl;
                await SendResponse(messageId, websocketUrl);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error getting WebSocket URL: {ex.Message}");
                await SendResponse(messageId, "ws://localhost:8085"); // Fallback to default
            }
        }

        private async Task HandlePickFolder(string messageId, JsonElement payload)
        {
            try
            {
                Logger.Info($"[pickFolder] Start handling (requestId: {messageId})");
                _fileDialogOpen = true;
                var allowListing = true;
                if (payload.ValueKind != JsonValueKind.Undefined &&
                    payload.TryGetProperty("allowFileListing", out var allowProp))
                {
                    if (allowProp.ValueKind == JsonValueKind.False) allowListing = false;
                    if (allowProp.ValueKind == JsonValueKind.True) allowListing = true;
                }
                var description = payload.ValueKind != JsonValueKind.Undefined &&
                                  payload.TryGetProperty("description", out var descProp)
                                    ? descProp.GetString()
                                    : "Select a folder that contains Excel files to merge";

                var extensions = new List<string> { ".xlsx", ".xlsm" };
                if (payload.ValueKind != JsonValueKind.Undefined && payload.TryGetProperty("extensions", out var extProp) && extProp.ValueKind == JsonValueKind.Array)
                {
                    try
                    {
                        extensions = extProp.EnumerateArray()
                            .Select(e => e.GetString())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .Select(s => s.StartsWith(".") ? s : "." + s)
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();
                        if (extensions.Count == 0)
                        {
                            extensions = new List<string> { ".xlsx", ".xlsm" };
                        }
                    }
                    catch (JsonException)
                    {
                        // If the JSON structure for extensions is invalid, fall back to default extensions.
                        extensions = new List<string> { ".xlsx", ".xlsm" };
                    }
                    catch (InvalidOperationException ex)
                    {
                        // Log unexpected errors while still falling back to safe default extensions.
                        Logger.Info($"Error parsing extensions from payload: {ex.Message}");
                        extensions = new List<string> { ".xlsx", ".xlsm" };
                    }
                    catch (ArgumentException ex)
                    {
                        // Log unexpected errors while still falling back to safe default extensions.
                        Logger.Info($"Error parsing extensions from payload: {ex.Message}");
                        extensions = new List<string> { ".xlsx", ".xlsm" };
                    }
                    catch (FormatException ex)
                    {
                        // Log unexpected errors while still falling back to safe default extensions.
                        Logger.Info($"Error parsing extensions from payload: {ex.Message}");
                        extensions = new List<string> { ".xlsx", ".xlsm" };
                    }
                }

                Logger.Info($"[pickFolder] allowListing={allowListing}, extensions={string.Join(",", extensions)}");

                string selectedPath = null;
                DialogResult result = DialogResult.None;
                Exception dialogException = null;

                void RunDialog()
                {
                    try
                    {
                        using (var dialog = new FolderBrowserDialog
                        {
                            Description = description,
                            ShowNewFolderButton = true,
                        })
                        {
                            Logger.Info("[pickFolder] Showing folder dialog on STA thread");
                            result = dialog.ShowDialog();
                            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                            {
                                selectedPath = dialog.SelectedPath;
                            }
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        dialogException = ex;
                    }
                    catch (IOException ex)
                    {
                        dialogException = ex;
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        dialogException = ex;
                    }
                }

                var dialogThread = new Thread(RunDialog);
                dialogThread.SetApartmentState(ApartmentState.STA);
                dialogThread.Start();
                dialogThread.Join();

                _fileDialogOpen = false;

                if (dialogException != null)
                {
                    Logger.Info($"[pickFolder] Dialog threw exception: {dialogException}");
                    await SendResponse(messageId, new { path = (string)null, files = Array.Empty<string>(), error = dialogException.Message });
                    SendEventToReact("focusTextInput", new { });
                    return;
                }

                Logger.Info($"[pickFolder] Dialog result={result}, selectedPath={(string.IsNullOrWhiteSpace(selectedPath) ? "<none>" : selectedPath)}");

                if (!string.IsNullOrWhiteSpace(selectedPath))
                {
                    string[] files = Array.Empty<string>();
                    if (allowListing && Directory.Exists(selectedPath))
                    {
                        try
                        {
                            files = Directory.GetFiles(selectedPath)
                                .Where(f => extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
                                .Select(Path.GetFileName)
                                .ToArray();
                            Logger.Info($"[pickFolder] Listed {files.Length} files");
                        }
                        catch (Exception listEx) when (
                            listEx is UnauthorizedAccessException ||
                            listEx is DirectoryNotFoundException ||
                            listEx is IOException ||
                            listEx is PathTooLongException)
                        {
                            Logger.Info($"Error listing files for folder {selectedPath}: {listEx.Message}");
                        }
                    }

                    await SendResponse(messageId, new { path = selectedPath, files });
                    SendEventToReact("focusTextInput", new { });
                }
                else
                {
                    await SendResponse(messageId, new { path = (string)null, files = Array.Empty<string>() });
                    SendEventToReact("focusTextInput", new { });
                }
            }
            catch (OperationCanceledException)
            {
                // Allow cancellations to propagate instead of being treated as generic errors.
                throw;
            }
            catch (Exception ex) when (
                ex is IOException ||
                ex is UnauthorizedAccessException ||
                ex is PathTooLongException ||
                ex is DirectoryNotFoundException ||
                ex is InvalidOperationException ||
                ex is ArgumentException)
            {
                _fileDialogOpen = false;
                Logger.Info($"Error handling pickFolder: {ex}");
                await SendResponse(messageId, new { path = (string)null, files = Array.Empty<string>(), error = ex.Message });
            }
        }

        private async Task HandlePickFile(string messageId, JsonElement payload)
        {
            try
            {
                Logger.Info($"[pickFile] Start handling (requestId: {messageId})");
                _fileDialogOpen = true;

                var extensions = new List<string> { ".xlsx", ".xlsm", ".xls" };
                var dialogTitle = "Select Excel file to compare";
                var filterLabel = "Excel Files";
                if (payload.ValueKind != JsonValueKind.Undefined && payload.TryGetProperty("extensions", out var extProp) && extProp.ValueKind == JsonValueKind.Array)
                {
                    try
                    {
                        extensions = extProp.EnumerateArray()
                            .Select(e => e.GetString())
                            .Where(s => !string.IsNullOrWhiteSpace(s))
                            .Select(s => s.StartsWith(".") ? s : "." + s)
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();
                        if (extensions.Count == 0)
                        {
                            extensions = new List<string> { ".xlsx", ".xlsm", ".xls" };
                        }
                    }
                    catch (JsonException)
                    {
                        extensions = new List<string> { ".xlsx", ".xlsm", ".xls" };
                    }
                }

                if (payload.ValueKind != JsonValueKind.Undefined &&
                    payload.TryGetProperty("title", out var titleProp) &&
                    titleProp.ValueKind == JsonValueKind.String)
                {
                    var requestedTitle = titleProp.GetString();
                    if (!string.IsNullOrWhiteSpace(requestedTitle))
                    {
                        dialogTitle = requestedTitle;
                    }
                }

                if (payload.ValueKind != JsonValueKind.Undefined &&
                    payload.TryGetProperty("filterLabel", out var labelProp) &&
                    labelProp.ValueKind == JsonValueKind.String)
                {
                    var requestedLabel = labelProp.GetString();
                    if (!string.IsNullOrWhiteSpace(requestedLabel))
                    {
                        filterLabel = requestedLabel;
                    }
                }

                Logger.Info($"[pickFile] extensions={string.Join(",", extensions)}");

                string selectedFilePath = null;
                DialogResult result = DialogResult.None;
                Exception dialogException = null;

                void RunDialog()
                {
                    try
                    {
                        var filter = string.Join(";", extensions.Select(ext => $"*{ext}"));
                        using (var dialog = new OpenFileDialog
                        {
                            Filter = $"{filterLabel} ({filter})|{filter}|All Files (*.*)|*.*",
                            FilterIndex = 1,
                            Title = dialogTitle
                        })
                        {
                            Logger.Info("[pickFile] Showing file dialog on STA thread");
                            result = dialog.ShowDialog();
                            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.FileName))
                            {
                                selectedFilePath = dialog.FileName;
                            }
                        }
                    }
                    catch (InvalidOperationException ex)
                    {
                        dialogException = ex;
                    }
                    catch (IOException ex)
                    {
                        dialogException = ex;
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        dialogException = ex;
                    }
                }

                var dialogThread = new Thread(RunDialog);
                dialogThread.SetApartmentState(ApartmentState.STA);
                dialogThread.Start();
                dialogThread.Join();

                _fileDialogOpen = false;

                if (dialogException != null)
                {
                    Logger.Info($"[pickFile] Dialog threw exception: {dialogException}");
                    await SendResponse(messageId, (string)null);
                    SendEventToReact("focusTextInput", new { });
                    return;
                }

                Logger.Info($"[pickFile] Dialog result={result}, selectedFilePath={(string.IsNullOrWhiteSpace(selectedFilePath) ? "<none>" : selectedFilePath)}");

                await SendResponse(messageId, selectedFilePath ?? (string)null);
                SendEventToReact("focusTextInput", new { });
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex) when (
                ex is IOException ||
                ex is UnauthorizedAccessException ||
                ex is PathTooLongException ||
                ex is DirectoryNotFoundException ||
                ex is InvalidOperationException ||
                ex is ArgumentException)
            {
                _fileDialogOpen = false;
                Logger.Info($"Error handling pickFile: {ex}");
                await SendResponse(messageId, (string)null);
            }
        }

        private async Task HandleOpenWorkbookFromWebView(string messageId, JsonElement payload)
        {
            try
            {
                Logger.Info($"[openWorkbook] Start handling (requestId: {messageId})");

                if (payload.ValueKind == JsonValueKind.Undefined || !payload.TryGetProperty("filePath", out var filePathProp))
                {
                    Logger.Info("[openWorkbook] No filePath provided");
                    await SendResponse(messageId, new { success = false, error = "filePath is required" });
                    return;
                }

                var filePath = filePathProp.GetString();
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    Logger.Info("[openWorkbook] Empty filePath");
                    await SendResponse(messageId, new { success = false, error = "filePath cannot be empty" });
                    return;
                }

                Logger.Info($"[openWorkbook] Opening workbook: {filePath}");

                // Use the Excel API to open the workbook
                var application = _container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                if (application == null)
                {
                    Logger.Info("[openWorkbook] Excel application not available");
                    await SendResponse(messageId, new { success = false, error = "Excel application not available" });
                    return;
                }

                await Task.Run(() =>
                {
                    try
                    {
                        if (!System.IO.File.Exists(filePath))
                        {
                            Logger.Info($"[openWorkbook] File not found: {filePath}");
                            SendResponse(messageId, new { success = false, error = "File not found" }).Wait();
                            return;
                        }

                        // Open the workbook
                        var workbook = application.Workbooks.Open(
                            Filename: filePath,
                            UpdateLinks: 0,
                            ReadOnly: false,
                            Format: 5,
                            Password: Type.Missing,
                            WriteResPassword: Type.Missing,
                            IgnoreReadOnlyRecommended: true,
                            Origin: Type.Missing,
                            Delimiter: Type.Missing,
                            Editable: false,
                            Notify: false,
                            Converter: Type.Missing,
                            AddToMru: true,
                            Local: Type.Missing,
                            CorruptLoad: Type.Missing
                        );

                        Logger.Info($"[openWorkbook] Successfully opened workbook: {workbook.Name}");
                        SendResponse(messageId, new { success = true, workbookName = workbook.Name }).Wait();
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Logger.Info($"[openWorkbook] Access denied when opening workbook '{filePath}': {ex.Message}");
                        SendResponse(messageId, new { success = false, error = "Access denied when opening the file." }).Wait();
                    }
                    catch (IOException ex)
                    {
                        Logger.Info($"[openWorkbook] I/O error when opening workbook '{filePath}': {ex.Message}");
                        SendResponse(messageId, new { success = false, error = "I/O error when opening the file." }).Wait();
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Logger.Info($"[openWorkbook] COM error when opening workbook '{filePath}': {ex.Message}");
                        SendResponse(messageId, new { success = false, error = "Excel reported an error while opening the workbook." }).Wait();
                    }
                    catch (OperationCanceledException ex)
                    {
                        Logger.Info($"[openWorkbook] Operation was canceled when opening workbook '{filePath}': {ex.Message}");
                        SendResponse(messageId, new { success = false, error = "The operation was canceled while opening the workbook." }).Wait();
                    }
                });
            }
            catch (OperationCanceledException ex)
            {
                Logger.Info($"[openWorkbook] Operation was canceled: {ex.Message}");
                await SendResponse(messageId, new { success = false, error = "The operation was canceled while opening the workbook." });
            }
            catch (Exception ex)
            {
                Logger.Info($"[openWorkbook] Unexpected error: {ex}");
                throw;
            }
        }

        private async Task HandleGetSlashCommandsFromSheet(string messageId)
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var application = _container.Resolve<Microsoft.Office.Interop.Excel.Application>();
                var workbook = application?.ActiveWorkbook;

                if (workbook == null)
                {
                    await SendResponse(messageId, new { success = false, commands = Array.Empty<object>(), error = "No active workbook" });
                    return;
                }

                var worksheet = excelSelection.GetWorksheet("X21_Commands", workbook.Name);
                if (worksheet == null)
                {
                    await SendResponse(messageId, new { success = true, commands = Array.Empty<object>() });
                    return;
                }

                var usedRange = worksheet.UsedRange;
                if (usedRange == null)
                {
                    await SendResponse(messageId, new { success = true, commands = Array.Empty<object>() });
                    return;
                }

                var rawValues = usedRange.Value2 as object[,];
                if (rawValues == null)
                {
                    await SendResponse(messageId, new { success = true, commands = Array.Empty<object>() });
                    return;
                }

                int rowStart = rawValues.GetLowerBound(0);
                int rowEnd = rawValues.GetUpperBound(0);
                int colStart = rawValues.GetLowerBound(1);
                int colEnd = rawValues.GetUpperBound(1);

                // Build header lookup
                var headerLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = colStart; col <= colEnd; col++)
                {
                    var header = rawValues[rowStart, col]?.ToString()?.Trim();
                    var normalizedHeader = NormalizeHeader(header);
                    if (!string.IsNullOrEmpty(normalizedHeader) && !headerLookup.ContainsKey(normalizedHeader))
                    {
                        headerLookup[normalizedHeader] = col;
                    }
                }

                var commands = new List<object>();

                var nameColumn = ResolveColumn(headerLookup, "name");

                for (int row = rowStart + 1; row <= rowEnd; row++)
                {
                    string name = GetCellString(rawValues, row, nameColumn);


                    var title = GetCellString(rawValues, row, ResolveColumn(headerLookup, "title"));
                    var description = GetCellString(rawValues, row, ResolveColumn(headerLookup, "description"));
                    var prompt = GetCellString(rawValues, row, ResolveColumn(headerLookup, "prompt"));
                    var requiresInput = ParseBool(GetCell(rawValues, row, ResolveColumn(headerLookup, "requiresInput")));

                    commands.Add(new
                    {
                        name,
                        title = string.IsNullOrWhiteSpace(title) ? name : title,
                        description = description ?? string.Empty,
                        prompt = prompt ?? string.Empty,
                        requiresInput
                    });
                }

                await SendResponse(messageId, new { success = true, commands });
            }
            catch (Exception ex)
            {
                Logger.Info($"Error reading slash commands sheet: {ex.Message}");
                await SendResponse(messageId, new { success = false, commands = Array.Empty<object>(), error = ex.Message });
            }
        }

        private static object GetCell(object[,] values, int row, int col)
        {
            if (col <= 0) return null;
            try
            {
                return values[row, col];
            }
            catch
            {
                return null;
            }
        }

        private static string GetCellString(object[,] values, int row, int col)
        {
            var value = GetCell(values, row, col);
            return value?.ToString();
        }

        private static int ResolveColumn(Dictionary<string, int> headerLookup, string headerName)
        {
            var normalized = NormalizeHeader(headerName);
            if (string.IsNullOrEmpty(normalized))
            {
                return -1;
            }

            if (headerLookup.TryGetValue(normalized, out var col))
            {
                return col;
            }

            return -1;
        }

        private static bool ParseBool(object value)
        {
            if (value == null) return false;

            switch (value)
            {
                case bool b:
                    return b;
                case double d:
                    return Math.Abs(d) > double.Epsilon;
                case int i:
                    return i != 0;
                case string s:
                    var normalized = s.Trim().ToLowerInvariant();
                    return normalized == "true" || normalized == "yes" || normalized == "y" || normalized == "1";
                default:
                    return false;
            }
        }

        private static string NormalizeHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return string.Empty;
            return header.Replace(" ", string.Empty).Replace("_", string.Empty).ToLowerInvariant();
        }

        private async Task HandleRevertFromToolId(string messageId, string toolUseId, string workbookName, string userId)
        {
            try
            {
                Logger.Info($"Reverting from tool ID: {toolUseId} for workbook: {workbookName}");

                var httpClient = new HttpClient();
                var requestData = new
                {
                    toolUseId = toolUseId,
                    workbookName = workbookName,
                    userId = userId
                };

                var json = JsonSerializer.Serialize(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{BackendConfigService.Instance.BaseUrl}/revertFromToolId", content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    Logger.Info($"Successfully reverted from tool ID: {toolUseId}");
                    await SendResponse(messageId, new { success = true, message = "Successfully reverted changes" });
                }
                else
                {
                    Logger.Info($"Failed to revert from tool ID: {toolUseId}. Response: {responseContent}");
                    await SendResponse(messageId, new { success = false, error = $"Failed to revert: {responseContent}" });
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error reverting from tool ID: {ex.Message}");
                await SendResponse(messageId, new { success = false, error = ex.Message });
            }
        }

        private async Task HandleApplyFromToolId(string messageId, string toolUseId, string workbookName, string userId)
        {
            try
            {
                Logger.Info($"Applying from tool ID: {toolUseId} for workbook: {workbookName}");

                var httpClient = new HttpClient();
                var requestData = new
                {
                    toolUseId = toolUseId,
                    workbookName = workbookName,
                    userId = userId
                };

                var json = JsonSerializer.Serialize(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync($"{BackendConfigService.Instance.BaseUrl}/applyFromToolId", content);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    Logger.Info($"Successfully applied from tool ID: {toolUseId}");
                    await SendResponse(messageId, new { success = true, message = "Successfully applied changes" });
                }
                else
                {
                    Logger.Info($"Failed to apply from tool ID: {toolUseId}. Response: {responseContent}");
                    await SendResponse(messageId, new { success = false, error = $"Failed to apply: {responseContent}" });
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error applying from tool ID: {ex.Message}");
                await SendResponse(messageId, new { success = false, error = ex.Message });
            }
        }

        #endregion

        #region Chat Service Event Handlers

        private void OnChatStreamDelta(object sender, string deltaText)
        {
            try
            {
                // Send streaming text to React UI
                SendEventToReact("streamDelta", deltaText);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending stream delta event: {ex.Message}");
            }
        }

        private void OnChatStreamComplete(object sender, Models.Message message)
        {
            try
            {
                // Generate a unique message ID that will be sent to frontend
                var messageId = $"msg_{DateTime.Now.Ticks}";

                // Send stream completion to React UI
                SendEventToReact("streamComplete", new
                {
                    role = message.Role,
                    content = message.Content,
                    timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    traceId = message.TraceId,
                    messageId = messageId // Include message ID for frontend mapping
                });
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending stream complete event: {ex.Message}");
            }
        }

        private void OnChatStreamError(object sender, Exception ex)
        {
            try
            {
                Logger.Info($"🔴 StreamError received: {ex.Message}");
                Logger.Info($"🔴 Sending streamError event to React UI");

                // Send error to React UI
                SendEventToReact("streamError", new { error = ex.Message });

                Logger.Info($"🔴 streamError event sent successfully");
            }
            catch (Exception logEx)
            {
                Logger.Info($"Error sending stream error event: {logEx.Message}");
            }
        }

        #endregion

        #region Public Methods



        public void FocusTextInput()
        {
            if (!_isInitialized || !_frontEndIsReady)
            {
                Logger.Info("WebView2 or frontend not ready, queueing focus request");
                _pendingFocusRequest = true;
                return;
            }
            try
            {
                // Mark this as an intentional focus request to prevent auto-focus back to Excel
                _intentionalFocusRequest = true;
                Logger.Info("Intentional focus request detected (keyboard shortcut)");

                if (InvokeRequired)
                {
                    Invoke(new Action(() =>
                    {
                        webView.Focus();
                        SendEventToReact("focusTextInput", new { });
                        Logger.Info("WebView2 focused and focus text input event sent to React");
                    }));
                }
                else
                {
                    webView.Focus();
                    SendEventToReact("focusTextInput", new { });
                    Logger.Info("WebView2 focused and focus text input event sent to React");
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error focusing text input: {ex.Message}");
            }
        }

        // Scoring is now handled via WebSocket - old HandleScoreMessage method removed

        // Feedback is now handled via WebSocket - old HandleFeedbackMessage method removed

        /// <summary>
        /// Gets whether the WebView2 currently has focus
        /// </summary>
        public bool WebViewHasFocus => _webViewHasFocus;

        /// <summary>
        /// Sends selection change event to React UI
        /// </summary>
        /// <param name="selectedRange">The currently selected range address</param>
        public void SendSelectionChangedEvent(string selectedRange)
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var activeSheetName = string.Empty;

                try
                {
                    activeSheetName = excelSelection?.GetActiveSheetName() ?? string.Empty;
                }
                catch (Exception innerEx)
                {
                    Logger.Info($"Error resolving active sheet name: {innerEx.Message}");
                }

                SendEventToReact("selectionChanged", new { selectedRange, activeSheet = activeSheetName });
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending selection changed event: {ex.Message}");
            }
        }

        /// <summary>
        /// Sends the current Excel selection to React UI
        /// </summary>
        private void SendCurrentSelection()
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var currentRange = excelSelection?.GetSelectedRange() ?? string.Empty;
                SendSelectionChangedEvent(currentRange);
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending current selection: {ex.Message}");
            }
        }

        private void SendWorkbookContext()
        {
            try
            {
                var excelSelection = _container.Resolve<ExcelSelection>();
                var workbookName = excelSelection?.GetWorkbookName() ?? string.Empty;
                var workbookPath = excelSelection?.GetWorkbookPath() ?? string.Empty;

                SendEventToReact("workbookContext", new { workbookName, workbookPath });
            }
            catch (Exception ex)
            {
                if (ex is OutOfMemoryException ||
                    ex is ThreadAbortException ||
                    ex is AccessViolationException ||
                    ex is StackOverflowException)
                {
                    throw;
                }

                Logger.Info($"Error sending workbook context: {ex.Message}");
            }
        }

        #endregion

        #region Helper Methods

        private void RecreateTaskPane()
        {
            try
            {
                Logger.Info("Recreating CustomTaskPane via TaskPaneManager...");
                var manager = Globals.ThisAddIn.Container.Resolve<X21.TaskPane.TaskPaneManager>();
                manager?.RecreateActiveTaskPane();
            }
            catch (Exception ex)
            {
                Logger.Info($"Pane recreate failed → {ex}");
            }
        }

        private Task SendResponse(string id, object data)
        {
            if (!_isInitialized) return Task.CompletedTask;

            try
            {
                // Match the WebViewBridgeMessage format expected by React
                var response = new
                {
                    type = "response",
                    payload = new
                    {
                        requestId = id,
                        data = data
                    }
                };
                var responseJson = JsonSerializer.Serialize(response);

                // Marshal to UI thread if needed
                if (InvokeRequired)
                {
                    Invoke(new Action(() => webView.CoreWebView2.PostWebMessageAsString(responseJson)));
                }
                else
                {
                    webView.CoreWebView2.PostWebMessageAsString(responseJson);
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending response: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        private void SendEventToReact(string eventType, object data)
        {
            if (!_isInitialized) return;

            try
            {
                var eventMessage = new { type = "event", eventType, data };
                var eventJson = JsonSerializer.Serialize(eventMessage);

                // Marshal to UI thread if needed
                if (InvokeRequired)
                {
                    Invoke(new Action(() => webView.CoreWebView2.PostWebMessageAsString(eventJson)));
                }
                else
                {
                    webView.CoreWebView2.PostWebMessageAsString(eventJson);
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Error sending event to React: {ex.Message}");
            }
        }

        #endregion

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (webView != null)
                {
                    webView.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private async Task HandleApproveToolUse(string messageId, string toolId)
        {
            try
            {
                Logger.Info($"Approving tool use: {toolId}");

                var httpClient = new HttpClient();
                var requestData = new
                {
                    toolId = toolId,
                    userId = UserUtils.GetUserId()
                };

                var jsonContent = JsonSerializer.Serialize(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var backendUrl = BackendConfigService.Instance.BaseUrl;
                var response = await httpClient.PostAsync($"{backendUrl}/api/tool/approve", content);

                if (response.IsSuccessStatusCode)
                {
                    var responseBody = await response.Content.ReadAsStringAsync();
                    Logger.Info($"Tool approval successful: {responseBody}");
                    await SendResponse(messageId, new { success = true, message = $"Tool {toolId} approved successfully" });
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    Logger.Info($"Tool approval failed: {response.StatusCode} - {errorResponse}");
                    await SendResponse(messageId, new { success = false, error = $"Failed to approve tool: {response.StatusCode}" });
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Tool approval error: {ex.Message}");
                await SendResponse(messageId, new { success = false, error = $"Tool approval failed: {ex.Message}" });
            }
        }

        private async Task HandleRejectToolUse(string messageId, string toolId)
        {
            try
            {
                Logger.Info($"Rejecting tool use: {toolId}");

                var httpClient = new HttpClient();
                var requestData = new
                {
                    toolId = toolId,
                    userId = UserUtils.GetUserId()
                };

                var jsonContent = JsonSerializer.Serialize(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var backendUrl = BackendConfigService.Instance.BaseUrl;
                var response = await httpClient.PostAsync($"{backendUrl}/api/tool/reject", content);

                if (response.IsSuccessStatusCode)
                {
                    var responseBody = await response.Content.ReadAsStringAsync();
                    Logger.Info($"Tool rejection successful: {responseBody}");
                    await SendResponse(messageId, new { success = true, message = $"Tool {toolId} rejected successfully" });
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    Logger.Info($"Tool rejection failed: {response.StatusCode} - {errorResponse}");
                    await SendResponse(messageId, new { success = false, error = $"Failed to reject tool: {response.StatusCode}" });
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Tool rejection error: {ex.Message}");
                await SendResponse(messageId, new { success = false, error = $"Tool rejection failed: {ex.Message}" });
            }
        }
    }

    // Simple message classes for JSON deserialization
    public class WebViewMessage
    {
        public WebViewMessage() { }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("payload")]
        public MessagePayload Payload { get; set; }
    }

    public class MessagePayload
    {
        public MessagePayload() { }

        [JsonPropertyName("requestId")]
        public int? RequestId { get; set; }

        [JsonPropertyName("conversationId")]
        public string ConversationId { get; set; }

        [JsonPropertyName("score")]
        public double? Score { get; set; }

        [JsonPropertyName("traceId")]
        public string TraceId { get; set; }

        [JsonPropertyName("comment")]
        public string Comment { get; set; }
    }
}
