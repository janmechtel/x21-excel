using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using X21.Common.Data;
using X21.TaskPane;
using X21.Logging;

namespace X21.Utils
{
    /// <summary>
    /// Contains utility methods for managing windows, particularly for focus management.
    /// </summary>
    public static class WindowUtils
    {
        /// <summary>
        /// Uses Windows API to return keyboard focus back to the main Excel application window.
        /// This is often needed when focus is stolen by other windows, like WebView2.
        /// </summary>
        public static void FocusExcelWindow(Application app)
        {
            if (app == null) return;

            try
            {
                Logger.Info("[Focus] Attempting to restore focus to Excel...");

                // Make sure Excel is interactive and ready for input
                app.Interactive = true;

                // Use Windows API to force focus back to Excel.
                // Focusing the desktop first can help in some edge cases.
                NativeMethods.SetForegroundWindow(NativeMethods.GetDesktopWindow());

                // Then focus the main Excel window.
                IntPtr excelHwnd = new IntPtr(app.Hwnd);
                bool success = NativeMethods.SetForegroundWindow(excelHwnd);

                Logger.Info($"[Focus] Focus restoration result: {success}");
            }
            catch (Exception ex)
            {
                Logger.Info($"[Focus] Error during focus restoration: {ex.Message}");
                // Ignore errors in focus management, as it's not critical if it fails.
            }
        }

        /// <summary>
        /// Determines if we should restore focus to Excel based on current WebView focus state
        /// Simple and direct: if WebView has focus when user clicks Excel, restore focus
        /// </summary>
        public static bool ShouldRestoreFocusToExcel(TaskPaneManager taskPaneManager)
        {
            try
            {
                if (taskPaneManager != null)
                {
                    var taskPane = taskPaneManager.GetTaskPane();
                    if (taskPane?.Control is X21.TaskPane.WebView2TaskPaneHost webViewHost)
                    {
                        // Simple check: if WebView currently has focus, restore it to Excel
                        bool webViewHasFocus = webViewHost.WebViewHasFocus;
                        if (webViewHasFocus)
                        {
                            Logger.Info("[Focus] WebView has focus when cell clicked - will restore to Excel");
                        }
                        return webViewHasFocus;
                    }
                }

                return false; // No restoration needed
            }
            catch (Exception ex)
            {
                Logger.Info($"[Focus] Error checking focus state: {ex.Message}");
                return false; // If we can't determine, don't call FocusExcel
            }
        }

        /// <summary>
        /// Windows API methods for focus management.
        /// See: https://github.com/MicrosoftEdge/WebView2Feedback/issues/951
        /// </summary>
        private static class NativeMethods
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetDesktopWindow();

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
    }
}
