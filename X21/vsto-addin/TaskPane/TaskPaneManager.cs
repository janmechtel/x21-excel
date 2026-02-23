using X21.Common.Data;
using X21.Common.Model;
using X21.TaskPane.Commands;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using X21.Utils;

namespace X21.TaskPane
{
    /// <summary>
    /// Manages the chat taskpane lifecycle and visibility per workbook
    /// </summary>
    public class TaskPaneManager : Component
    {
        // Dictionary to track task panes per window handle (HWND)
        private readonly Dictionary<int, CustomTaskPane> _taskPanes = new Dictionary<int, CustomTaskPane>();
        private readonly Dictionary<int, WebView2TaskPaneHost> _taskPaneControls = new Dictionary<int, WebView2TaskPaneHost>();
        private readonly Dictionary<int, Container> _workbookContainers = new Dictionary<int, Container>();
        private readonly Application _excelApplication;

        public TaskPaneManager(Container container) : base(container)
        {
            _excelApplication = container.Resolve<Application>();
            Logging.Logger.Info($"TaskPaneManager created - Container Hash: {container.GetHashCode()}");
        }

        protected override void InitCommands()
        {
            var toggleTaskPaneAnnotation = new Annotation(Container)
            {
                Title = "AI Chat"
            };
            Commands.Add(new CommandToggleTaskPane(Container, toggleTaskPaneAnnotation));

            // Register the same command for the Home ribbon button
            Commands.Add(new CommandToggleTaskPane(Container, toggleTaskPaneAnnotation, "CommandToggleTaskPaneHome"));

            var tutorialAnnotation = new Annotation(Container)
            {
                Title = "Tutorial"
            };
            Commands.Add(new X21.Common.Commands.CommandTutorial(Container, tutorialAnnotation));

            var feedbackAnnotation = new Annotation(Container)
            {
                Title = "Feedback"
            };
            Commands.Add(new X21.Common.Commands.CommandFeedback(Container, feedbackAnnotation));
        }

        public override void Init()
        {
            base.Init();
            // Don't create task panes here - create them on-demand per workbook
        }

        public override void Exit()
        {
            base.Exit();
            CleanupAllTaskPanes();
        }

        /// <summary>
        /// Gets whether the taskpane is currently visible for the active workbook
        /// </summary>
        public bool IsVisible
        {
            get
            {
                var taskPane = GetOrCreateTaskPaneForActiveWorkbook();
                return taskPane?.Visible ?? false;
            }
        }

        /// <summary>
        /// Shows the chat taskpane for the active workbook
        /// </summary>
        public void ShowTaskPane()
        {
            var taskPane = GetOrCreateTaskPaneForActiveWorkbook();
            if (taskPane != null)
            {
                taskPane.Visible = true;
            }
        }

        /// <summary>
        /// Hides the chat taskpane for the active workbook
        /// </summary>
        public void HideTaskPane()
        {
            var taskPane = GetOrCreateTaskPaneForActiveWorkbook();
            if (taskPane != null)
            {
                taskPane.Visible = false;
            }
        }

        /// <summary>
        /// Toggles the taskpane visibility for the active workbook
        /// </summary>
        public void ToggleTaskPane()
        {
            var activeWindow = _excelApplication.ActiveWindow;
            if (activeWindow == null)
            {
                Logging.Logger.Info("ToggleTaskPane called but no active window.");
                return;
            }
            var windowKey = activeWindow.Hwnd;

            Logging.Logger.Info($"ToggleTaskPane called for window: {activeWindow.Caption} (Key: {windowKey})");

            if (IsVisible)
            {
                HideTaskPane();
            }
            else
            {
                ShowTaskPane();
            }
        }

        /// <summary>
        /// Gets or creates a task pane for the currently active workbook
        /// </summary>
        private CustomTaskPane GetOrCreateTaskPaneForActiveWorkbook()
        {
            try
            {
                var activeWorkbook = _excelApplication.ActiveWorkbook;
                var activeWindow = _excelApplication.ActiveWindow;

                if (activeWorkbook == null || activeWindow == null)
                {
                    Logging.Logger.Info("No active workbook or window");
                    return null;
                }

                var windowKey = activeWindow.Hwnd;

                // Check if we already have a task pane for this window
                if (_taskPanes.TryGetValue(windowKey, out var existingTaskPane))
                {
                    Logging.Logger.Info($"Using existing task pane for window: {activeWindow.Caption}");
                    return existingTaskPane;
                }

                // Create a new task pane for this window
                Logging.Logger.Info($"Creating new WebView2 task pane for window: {activeWindow.Caption}");

                // Create an isolated container for this window
                var windowContainer = CreateIsolatedContainer(windowKey, activeWorkbook);
                _workbookContainers[windowKey] = windowContainer;

                // Create the WebView2 chat control with the isolated container
                var chatControl = new WebView2TaskPaneHost(windowContainer);
                var thisAddIn = Globals.ThisAddIn;

                // Associate the task pane with the specific workbook window
                var taskPane = thisAddIn.CustomTaskPanes.Add(chatControl, "x²¹", activeWindow);

                // Configure taskpane properties
                taskPane.Width = 350;
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Visible = false; // Start hidden

                // Handle taskpane events
                taskPane.VisibleChanged += (sender, e) => OnTaskPaneVisibleChanged(sender, e, windowKey);

                // Store the task pane and control
                _taskPanes[windowKey] = taskPane;
                _taskPaneControls[windowKey] = chatControl;

                Logging.Logger.Info($"WebView2 task pane created successfully for window: {activeWindow.Caption} with isolated container");

                return taskPane;
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
                return null;
            }
        }

        /// <summary>
        /// Creates an isolated container for a workbook with its own ChatService instance
        /// </summary>
        private Container CreateIsolatedContainer(int windowKey, Workbook workbook)
        {
            var isolatedContainer = new Container();

            // Register the same Excel Application and ThisAddIn instances
            isolatedContainer.RegisterSingleton(_excelApplication);
            isolatedContainer.RegisterSingleton(Globals.ThisAddIn);

            if (workbook != null)
            {
                isolatedContainer.RegisterSingleton(workbook);
            }

            // Register components same as main container but with new instances
            isolatedContainer.RegisterSingleton<X21.Excel.ExcelSelection>();

            // Register the workbook change tracker service (shared from main container)
            var changeTracker = Container.Resolve<X21.Services.WorkbookChangeTrackerService>();
            if (changeTracker != null)
            {
                isolatedContainer.RegisterSingleton(changeTracker);
            }

            Logging.Logger.Info($"Created isolated container for window key: {windowKey}");

            return isolatedContainer;
        }

        private void OnTaskPaneVisibleChanged(object sender, EventArgs e, int windowKey)
        {
            var taskPane = sender as CustomTaskPane;
            Logging.Logger.Info($"Task pane visibility changed for window {windowKey}: {taskPane?.Visible}");
        }

        /// <summary>
        /// Cleans up the task pane associated with a specific workbook when it closes
        /// </summary>
        public void CleanupTaskPaneForWorkbook(Workbook workbook)
        {
            try
            {
                if (workbook == null)
                {
                    Logging.Logger.Info("CleanupTaskPaneForWorkbook called with null workbook");
                    return;
                }

                // Find all windows for this workbook and clean them up
                var windowsToCleanup = new List<int>();

                foreach (Window window in workbook.Windows)
                {
                    try
                    {
                        var windowKey = window.Hwnd;
                        if (_taskPanes.ContainsKey(windowKey))
                        {
                            windowsToCleanup.Add(windowKey);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.Logger.Info($"Error getting window handle: {ex.Message}");
                    }
                }

                // Clean up each window's task pane
                foreach (var windowKey in windowsToCleanup)
                {
                    CleanupTaskPaneByWindowKey(windowKey);
                }

                Logging.Logger.Info($"Cleaned up task panes for workbook: {workbook.Name}");
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
                Logging.Logger.Info($"Error cleaning up task pane for workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Cleans up a specific task pane by window key
        /// </summary>
        private void CleanupTaskPaneByWindowKey(int windowKey)
        {
            try
            {
                var thisAddIn = Globals.ThisAddIn;
                if (thisAddIn?.CustomTaskPanes == null)
                {
                    return;
                }

                // Get the task pane for this window
                if (_taskPanes.TryGetValue(windowKey, out var taskPane))
                {
                    // Remove event handlers
                    taskPane.VisibleChanged -= (sender, e) => OnTaskPaneVisibleChanged(sender, e, windowKey);

                    TryRemoveTaskPane(taskPane);

                    _taskPanes.Remove(windowKey);
                    Logging.Logger.Info($"Removed task pane for window key: {windowKey}");
                }

                // Dispose of the control
                if (_taskPaneControls.TryGetValue(windowKey, out var control))
                {
                    control?.Dispose();
                    _taskPaneControls.Remove(windowKey);
                    Logging.Logger.Info($"Disposed task pane control for window key: {windowKey}");
                }

                // Clean up isolated container
                if (_workbookContainers.TryGetValue(windowKey, out var container))
                {
                    try
                    {
                        // Exit components in the isolated container
                        foreach (var component in container.ResolveAll<X21.Interfaces.IComponent>())
                        {
                            component?.Exit();
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.Logger.LogException(ex);
                    }

                    _workbookContainers.Remove(windowKey);
                    Logging.Logger.Info($"Cleaned up isolated container for window key: {windowKey}");
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
                Logging.Logger.Info($"Error cleaning up task pane for window key {windowKey}: {ex.Message}");
            }
        }

        private void CleanupAllTaskPanes()
        {
            try
            {
                var thisAddIn = Globals.ThisAddIn;
                if (thisAddIn?.CustomTaskPanes == null)
                {
                    return;
                }

                foreach (var kvp in _taskPanes.ToList())
                {
                    var taskPane = kvp.Value;
                    var windowKey = kvp.Key;

                    if (taskPane != null)
                    {
                        // Remove event handlers
                        taskPane.VisibleChanged -= (sender, e) => OnTaskPaneVisibleChanged(sender, e, windowKey);

                        TryRemoveTaskPane(taskPane);
                    }
                }

                // Dispose of all controls
                foreach (var control in _taskPaneControls.Values)
                {
                    control?.Dispose();
                }

                // Clean up isolated containers
                foreach (var container in _workbookContainers.Values)
                {
                    try
                    {
                        // Exit components in the isolated container
                        foreach (var component in container.ResolveAll<X21.Interfaces.IComponent>())
                        {
                            component?.Exit();
                        }
                    }
                    catch (Exception ex)
                    {
                        Logging.Logger.LogException(ex);
                    }
                }

                _taskPanes.Clear();
                _taskPaneControls.Clear();
                _workbookContainers.Clear();
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
            }
        }

        private bool TryRemoveTaskPane(CustomTaskPane taskPane)
        {
            if (taskPane == null)
            {
                return false;
            }

            var thisAddIn = Globals.ThisAddIn;
            if (thisAddIn == null || thisAddIn.CustomTaskPanes == null)
            {
                return false;
            }

            try
            {
                if (thisAddIn.CustomTaskPanes.Cast<CustomTaskPane>().Any(pane => pane == taskPane))
                {
                    thisAddIn.CustomTaskPanes.Remove(taskPane);
                    return true;
                }
            }
            catch (ObjectDisposedException)
            {
                Logging.Logger.Info("CustomTaskPanes already disposed; skipping task pane removal.");
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
            }

            return false;
        }

        /// <summary>
        /// Gets the task pane for the active workbook (if it exists)
        /// </summary>
        public CustomTaskPane GetTaskPane()
        {
            var activeWindow = _excelApplication.ActiveWindow;
            if (activeWindow == null) return null;

            var windowKey = activeWindow.Hwnd;
            _taskPanes.TryGetValue(windowKey, out var taskPane);
            return taskPane;
        }

        /// <summary>
        /// Recreates the task pane for the active workbook/window, preserving width and dock position
        /// </summary>
        public void RecreateActiveTaskPane()
        {
            try
            {
                var activeWindow = _excelApplication.ActiveWindow;
                if (activeWindow == null)
                {
                    Logging.Logger.Info("RecreateActiveTaskPane called but no active window.");
                    return;
                }

                var windowKey = activeWindow.Hwnd;

                // Preserve layout from existing pane if present
                int preservedWidth = 350;
                var preservedDock = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                if (_taskPanes.TryGetValue(windowKey, out var existingPane) && existingPane != null)
                {
                    preservedWidth = existingPane.Width;
                    preservedDock = existingPane.DockPosition;
                }

                // Clean up existing pane/control/container for this window
                CleanupTaskPaneByWindowKey(windowKey);

                // Recreate pane using the centralized creation path
                var newPane = GetOrCreateTaskPaneForActiveWorkbook();
                if (newPane != null)
                {
                    newPane.Width = preservedWidth;
                    newPane.DockPosition = preservedDock;
                    newPane.Visible = true;
                    Logging.Logger.Info($"Task pane recreated for window {windowKey} (width={preservedWidth}, dock={preservedDock}).");
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Focuses the text input in the React app for the active workbook
        /// </summary>
        public void FocusTextInput()
        {
            try
            {
                var activeWindow = _excelApplication.ActiveWindow;
                if (activeWindow == null)
                {
                    Logging.Logger.Info("No active window for focusing text input");
                    return;
                }

                var windowKey = activeWindow.Hwnd;

                // Ensure TaskPane exists and is visible
                var taskPane = GetOrCreateTaskPaneForActiveWorkbook();
                if (taskPane == null)
                {
                    Logging.Logger.Info("Failed to get or create task pane");
                    return;
                }

                taskPane.Visible = true;

                // Send focus event to React - React handles all focus timing
                if (_taskPaneControls.TryGetValue(windowKey, out var control))
                {
                    control.FocusTextInput();
                    Logging.Logger.Info($"Focus event sent to React for window key: {windowKey}");
                }
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
            }
        }

        /// <summary>
        /// Checks if the WebView currently has focus for the active workbook
        /// </summary>
        /// <returns>True if WebView has focus, false otherwise</returns>
        public bool WebViewHasFocus()
        {
            try
            {
                var activeWindow = _excelApplication.ActiveWindow;
                if (activeWindow == null) return false;

                var windowKey = activeWindow.Hwnd;

                if (_taskPaneControls.TryGetValue(windowKey, out var control))
                {
                    return control.WebViewHasFocus;
                }

                return false;
            }
            catch (Exception ex)
            {
                Logging.Logger.LogException(ex);
                return false;
            }
        }

    }
}
