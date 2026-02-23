using X21.Common.Data;
using X21.Interfaces;
using X21.Logging;
using X21.Utils;
using X21.Services;
using System;
using System.Linq;
using System.Net;
using System.IO;
using System.Diagnostics;
using System.Threading.Tasks;

namespace X21
{
    public partial class ThisAddIn
    {
        public Container Container { get; private set; } = new Container();
        private X21.Utils.KeyboardHook _keyboardHook;
        private Process _backendProcess;
        private readonly object _backendProcessLock = new object();
        private bool _isShuttingDown;
        private ExcelStaDispatcher _excelStaDispatcher;
        // Track pending Save As operations: workbook FullName -> old workbook Name
        private readonly System.Collections.Generic.Dictionary<string, string> _pendingSaveAs = new System.Collections.Generic.Dictionary<string, string>();


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Configure TLS
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                NLogLogger.Init();

                EnvFileLoader.Load();

                // Test log message to verify logging is working
                Logger.Info("=== Excel Add-in Starting Up ===");
                Logger.Info($"Environment: {EnvironmentHelper.GetEnvironmentName()}");
                Logger.Info($"Assembly Name: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}");
                Logger.Info($"Excel Log Path: {EnvironmentHelper.GetExcelLogPath()}");
                Logger.Info($"Deno Log Path: {EnvironmentHelper.GetDenoLogPath()}");


                var application = this.Application;
                Container.RegisterSingleton(application);

                // Marshal every Excel COM call through a dedicated STA dispatcher thread.
                _excelStaDispatcher = new ExcelStaDispatcher();
                Container.RegisterSingleton(_excelStaDispatcher);

                OleMessageFilter.Register();

                ClassLoader.Instance.Init(typeof(ThisAddIn).Assembly);

                AddComponents();
                InitComponents();

                // Initialize keyboard shortcut handling using Microsoft's recommended approach
                InitializeKeyboardShortcuts();

                // Check if this is the first run and open samples.xlsx if it is
                CheckFirstRunAndOpenSamples();

                // Launch the backend server
                LaunchBackendServer();

                // Initialize PostHog analytics
                InitializePostHog();

                // Subscribe to workbook events
                SubscribeToWorkbookEvents();

                // Capture initial snapshots for any already-open workbooks
                CaptureSnapshotsForOpenWorkbooks();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info("Unhandled exception during startup.");
                if (IsFatalException(ex))
                {
                    throw;
                }
            }
        }

        private void CheckFirstRunAndOpenSamples()
        {
            try
            {
                // Check if this is the first run using RegistryHelper
                // GetBool will handle backward compatibility with old nested key format
                bool isFirstRun = RegistryHelper.GetBool("FirstRun", true);

                if (isFirstRun)
                {
                    // Mark as not first run (only in published builds, not Debug mode)
                    if (!EnvironmentHelper.IsDebugMode())
                    {
                        RegistryHelper.SetBool("FirstRun", false);
                    }

                    // Open the samples workbook
                    string tempPath = Path.Combine(Path.GetTempPath(), "x21-samples.xlsx");
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    using (var stream = assembly.GetManifestResourceStream("X21.x21-samples.xlsx")) // Namespace + filename
                    {
                        if (stream != null)
                        {
                            using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                            {
                                stream.CopyTo(fileStream);
                            }
                            Application.Workbooks.Open(tempPath);
                        }
                        else
                        {
                            Logger.Info("x21-samples.xlsx embedded resource not found.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info("Error checking first run status: " + ex.Message);
            }
        }

        private void LaunchBackendServer()
        {
            EnsureBackendServerRunning("startup");
        }

        /// <summary>
        /// Ensure the Deno backend process is running; restart if it exited.
        /// </summary>
        /// <param name="reason">Short label for logging why the restart was attempted.</param>
        /// <returns>True if the backend is running or was started successfully.</returns>
        public bool EnsureBackendServerRunning(string reason = "runtime check")
        {
            try
            {
                // Only launch backend in published builds (not Debug mode)
                if (EnvironmentHelper.IsDebugMode())
                {
                    Logger.Info("Skipping backend launch in Debug mode (run 'deno task dev' manually)");
                    return false;
                }

                lock (_backendProcessLock)
                {
                    if (_isShuttingDown)
                    {
                        Logger.Info("Add-in is shutting down, skipping backend restart");
                        return false;
                    }

                    if (_backendProcess != null && !_backendProcess.HasExited)
                    {
                        return true;
                    }

                    // Clean up stale handles from a previous process
                    if (_backendProcess != null)
                    {
                        _backendProcess.Exited -= OnBackendProcessExited;
                        _backendProcess.Dispose();
                        _backendProcess = null;
                    }

                    Logger.Info($"Attempting to launch backend server ({reason})...");

                    // Use PathResolver to find the backend executable (environment-specific)
                    var backendPath = PathResolver.GetBackendExecutablePath();

                    Logger.Info($"Backend path: {backendPath}");

                    // Check if the backend executable exists
                    if (!File.Exists(backendPath))
                    {
                        Logger.Info($"Backend executable not found at: {backendPath}");
                        return false;
                    }

                    // Create process start info
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = backendPath,
                        WorkingDirectory = Path.GetDirectoryName(backendPath),
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    // Set environment variable so the backend knows which environment it's running in
                    var environmentName = EnvironmentHelper.GetEnvironmentName();
                    startInfo.EnvironmentVariables["X21_ENVIRONMENT"] = environmentName;
                    Logger.Info($"Setting X21_ENVIRONMENT={environmentName} for backend process");

                    // Start the backend process
                    var backendProcess = new Process
                    {
                        StartInfo = startInfo,
                        EnableRaisingEvents = true
                    };

                    backendProcess.Exited += OnBackendProcessExited;

                    // Start the process
                    backendProcess.Start();
                    backendProcess.BeginOutputReadLine();
                    backendProcess.BeginErrorReadLine();

                    _backendProcess = backendProcess;

                    Logger.Info($"Backend server launched successfully ({reason}). Process ID: {_backendProcess.Id}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error launching backend server ({reason}): {ex.Message}");
            }

            return false;
        }

        private void OnBackendProcessExited(object sender, EventArgs e)
        {
            Logger.Info("Backend process has exited.");

            lock (_backendProcessLock)
            {
                if (_backendProcess != null)
                {
                    _backendProcess.Exited -= OnBackendProcessExited;
                    _backendProcess.Dispose();
                    _backendProcess = null;
                }
            }

            if (_isShuttingDown)
            {
                Logger.Info("Add-in is shutting down; skipping backend restart.");
                return;
            }

            Task.Run(() =>
            {
                Logger.Info("Attempting to restart backend server after unexpected exit...");
                EnsureBackendServerRunning("auto-restart after exit");
            });
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs args)
        {
            try
            {
                _isShuttingDown = true;
                // Unsubscribe from workbook events
                UnsubscribeFromWorkbookEvents();

                // Clean up keyboard hook
                _keyboardHook?.Dispose();

                // Clean up backend process
                CleanupBackendProcess();

                ExitComponents();

                PostHogService.Instance.Dispose();

                OleMessageFilter.Revoke();

                _excelStaDispatcher?.Dispose();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info("Unhandled exception during shutdown.");
                if (IsFatalException(ex))
                {
                    throw;
                }
            }
        }

        private static bool IsFatalException(Exception ex)
        {
            if (ex is AggregateException aggregateException)
            {
                return aggregateException.Flatten().InnerExceptions.Any(IsFatalException);
            }

            return ex is OutOfMemoryException
                || ex is AccessViolationException
                || ex is AppDomainUnloadedException;
        }

        private void UnsubscribeFromWorkbookEvents()
        {
            try
            {
                Logger.Info("Unsubscribing from workbook events");
                this.Application.WorkbookOpen -= Application_WorkbookOpen;
                ((Microsoft.Office.Interop.Excel.AppEvents_Event)this.Application).NewWorkbook -= Application_NewWorkbook;
                this.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                this.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
                this.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
                Logger.Info("Workbook events unsubscribed successfully");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error unsubscribing from workbook events: {ex.Message}");
            }
        }

        private void CleanupBackendProcess()
        {
            try
            {
                lock (_backendProcessLock)
                {
                    if (_backendProcess != null)
                    {
                        _backendProcess.Exited -= OnBackendProcessExited;

                        if (!_backendProcess.HasExited)
                        {
                            Logger.Info("Shutting down backend process...");

                            // Kill the process immediately without waiting
                            // This allows Excel to close quickly without blocking
                            _backendProcess.Kill();

                            Logger.Info("Backend process kill signal sent.");
                        }

                        // Dispose without waiting for exit
                        _backendProcess.Dispose();
                        _backendProcess = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error cleaning up backend process: {ex.Message}");
            }
        }

        private void InitializePostHog()
        {
            try
            {
                var userId = UserUtils.GetUserId();
                Logger.Info($"Initializing PostHog analytics for user: {userId}");

                PostHogService.Instance.Identify(userId);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error initializing PostHog: {ex.Message}");
            }
        }

        private void SubscribeToWorkbookEvents()
        {
            try
            {
                Logger.Info("Subscribing to workbook events");
                this.Application.WorkbookOpen += Application_WorkbookOpen;
                ((Microsoft.Office.Interop.Excel.AppEvents_Event)this.Application).NewWorkbook += Application_NewWorkbook;
                this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
                this.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
                Logger.Info("Workbook events subscribed successfully");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error subscribing to workbook events: {ex.Message}");
            }
        }

        private void CaptureSnapshotsForOpenWorkbooks()
        {
            try
            {
                Logger.Info("Capturing initial snapshots for already-open workbooks");
                var changeTracker = Container.Resolve<X21.Services.WorkbookChangeTrackerService>();
                if (changeTracker == null)
                {
                    Logger.Info("WorkbookChangeTrackerService not available");
                    return;
                }

                foreach (Microsoft.Office.Interop.Excel.Workbook workbook in Application.Workbooks)
                {
                    try
                    {
                        Logger.Info($"Capturing snapshot for open workbook: {workbook.Name}");
                        _ = changeTracker.SendInitialSnapshot(workbook);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogException(ex);
                        Logger.Info($"Error capturing snapshot for {workbook.Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error capturing snapshots for open workbooks: {ex.Message}");
            }
        }

        private void Application_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            try
            {
                Logger.Info($"Workbook opened: {Wb.Name}");

                // Capture initial snapshot for change tracking
                var changeTracker = Container.Resolve<X21.Services.WorkbookChangeTrackerService>();
                if (changeTracker != null)
                {
                    _ = changeTracker.SendInitialSnapshot(Wb);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error handling WorkbookOpen: {ex.Message}");
            }
        }

        private void Application_NewWorkbook(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            try
            {
                Logger.Info($"New workbook created: {Wb.Name}");

                // Capture initial snapshot for change tracking
                var changeTracker = Container.Resolve<X21.Services.WorkbookChangeTrackerService>();
                if (changeTracker != null)
                {
                    _ = changeTracker.SendInitialSnapshot(Wb);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error handling NewWorkbook: {ex.Message}");
            }
        }

        private void Application_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
        {
            try
            {
                Logger.Info($"Workbook closing: {Wb.Name}");

                // Get the task pane manager and close the task pane for this workbook
                var taskPaneManager = Container.Resolve<X21.TaskPane.TaskPaneManager>();
                if (taskPaneManager != null)
                {
                    taskPaneManager.CleanupTaskPaneForWorkbook(Wb);
                }

                // Note: Snapshot cleanup is now handled server-side in Deno state manager
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error handling WorkbookBeforeClose: {ex.Message}");
            }
        }

        private void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                Logger.Info($"Workbook about to be saved: {Wb.Name}, SaveAsUI: {SaveAsUI}");

                // Track Save As operations to copy snapshot after save
                if (SaveAsUI)
                {
                    // Store the old name keyed by FullName (more stable than Name)
                    var fullName = Wb.FullName;
                    var oldName = Wb.Name;
                    _pendingSaveAs[fullName] = oldName;
                    Logger.Info($"Save As detected, tracking old name '{oldName}' for workbook at '{fullName}'");
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error handling WorkbookBeforeSave: {ex.Message}");
            }
        }

        private void Application_WorkbookAfterSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool Success)
        {
            try
            {
                if (Success)
                {
                    Logger.Info($"Workbook saved successfully: {Wb.Name}");

                    // Check if this was a Save As operation
                    var currentFullName = Wb.FullName;
                    var currentName = Wb.Name;
                    string oldName = null;

                    // Check if we have a pending Save As for this workbook
                    // Try to match by FullName first (might have changed on Save As)
                    if (_pendingSaveAs.TryGetValue(currentFullName, out oldName))
                    {
                        // Found by FullName - remove from pending
                        _pendingSaveAs.Remove(currentFullName);
                    }
                    else
                    {
                        // FullName might have changed on Save As, check all pending entries
                        // Find entry where old name no longer exists as an open workbook
                        foreach (var kvp in _pendingSaveAs.ToList())
                        {
                            var candidateOldName = kvp.Value;

                            // Check if any open workbook still has the old name
                            bool oldNameStillExists = false;
                            foreach (Microsoft.Office.Interop.Excel.Workbook openWb in this.Application.Workbooks)
                            {
                                if (openWb.Name == candidateOldName && openWb.FullName != currentFullName)
                                {
                                    oldNameStillExists = true;
                                    break;
                                }
                            }

                            // If old name doesn't exist anymore, this is likely our Save As
                            if (!oldNameStillExists && candidateOldName != currentName)
                            {
                                oldName = candidateOldName;
                                _pendingSaveAs.Remove(kvp.Key);
                                break;
                            }
                        }
                    }

                    // If we found an old name and it's different from current, copy snapshot
                    if (oldName != null && oldName != currentName)
                    {
                        Logger.Info($"Save As detected: copying snapshot from '{oldName}' to '{currentName}'");
                        var changeTracker = Container.Resolve<X21.Services.WorkbookChangeTrackerService>();
                        if (changeTracker != null)
                        {
                            // Run async operation without blocking
                            Task.Run(async () =>
                            {
                                try
                                {
                                    await changeTracker.CopySnapshot(oldName, currentName);
                                }
                                catch (IOException ioEx)
                                {
                                    Logger.LogException(ioEx);
                                    Logger.Info($"I/O error copying snapshot from '{oldName}' to '{currentName}': {ioEx.Message}");
                                }
                                catch (UnauthorizedAccessException authEx)
                                {
                                    Logger.LogException(authEx);
                                    Logger.Info($"Access error copying snapshot from '{oldName}' to '{currentName}': {authEx.Message}");
                                }
                                catch (WebException webEx)
                                {
                                    Logger.LogException(webEx);
                                    Logger.Info($"Network error copying snapshot from '{oldName}' to '{currentName}': {webEx.Message}");
                                }
                                catch (Exception asyncEx)
                                {
                                    // Unexpected error: log and rethrow to avoid silently swallowing programming or severe faults
                                    Logger.LogException(asyncEx);
                                    Logger.Info($"Unexpected error copying snapshot from '{oldName}' to '{currentName}': {asyncEx.Message}");
                                    throw;
                                }
                            });
                        }
                    }
                }
                else
                {
                    Logger.Info($"Workbook save failed: {Wb.Name}");
                    // Clean up pending Save As if save failed
                    var currentFullName = Wb.FullName;
                    _pendingSaveAs.Remove(currentFullName);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info($"Error handling WorkbookAfterSave: {ex.Message}");
            }
        }

        private void InitializeKeyboardShortcuts()
        {
            try
            {
                Logger.Info("Initializing keyboard shortcuts: Ctrl+Shift+A (chat)");
                _keyboardHook = new X21.Utils.KeyboardHook();
                _keyboardHook.CtrlShiftAPressed += OnCtrlShiftAPressed;
                Logger.Info("Keyboard hooks initialized successfully");
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info("Failed to initialize keyboard hook: " + ex.Message);
            }
        }

        private void OnCtrlShiftAPressed()
        {
            try
            {
                var taskPaneManager = Container.Resolve<X21.TaskPane.TaskPaneManager>();
                if (taskPaneManager == null)
                {
                    Logger.Info("Task pane manager not found");
                    return;
                }

                // Check if WebView currently has focus
                if (taskPaneManager.WebViewHasFocus())
                {
                    Logger.Info("Ctrl+Shift+A detected in WebView - returning focus to Excel");
                    WindowUtils.FocusExcelWindow(Application);
                }
                else
                {
                    Logger.Info("Ctrl+Shift+A detected - focusing chat input");

                    // Show the task pane if it's not visible
                    var toggleCommand = Container.CommandById("CommandToggleTaskPane");
                    if (toggleCommand?.CanExecute(null) == true)
                    {
                        var taskPane = taskPaneManager.GetTaskPane();
                        if (taskPane == null || !taskPane.Visible)
                        {
                            Logger.Info("Task pane not visible, showing it");
                            toggleCommand.Execute(null);
                        }

                        // Focus the text input in the React app
                        taskPaneManager.FocusTextInput();
                    }
                    else
                    {
                        Logger.Info("Toggle task pane command not found or cannot execute");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex);
                Logger.Info("Error in OnCtrlShiftAPressed: " + ex.Message);
            }
        }

        private void AddComponents()
        {
            // Register core components
            Container.RegisterSingleton<X21.Excel.ExcelSelection>();

            Container.RegisterSingleton<X21.TaskPane.TaskPaneManager>();

            // Register utility commands (folder opening commands)
            Container.RegisterSingleton<X21.Common.Commands.UtilityCommands>();

            // Register Excel API service
            Container.RegisterSingleton<X21.Services.ExcelApiService>();

            // Register workbook change tracker service
            Container.RegisterSingleton<X21.Services.WorkbookChangeTrackerService>();
        }

        private void InitComponents()
        {
            foreach (var component in Container.ResolveAll<IComponent>())
            {
                component.Init();
            }
        }

        private void ExitComponents()
        {
            foreach (var component in Container.ResolveAll<IComponent>())
            {
                Execute.Call(() => component.Exit(), Execute.CatchMode.LogFileOnly);
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new X21.Ribbon.Ribbon(Container);

            Container.RegisterSingleton(ribbon);

            return ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
