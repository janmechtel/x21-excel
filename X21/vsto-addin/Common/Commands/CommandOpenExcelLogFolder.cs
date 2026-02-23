using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;
using X21.Utils;
using System;
using System.Diagnostics;
using System.IO;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the Excel add-in log folder in Windows Explorer
    /// </summary>
    public class CommandOpenExcelLogFolder : CommandBase
    {
        public CommandOpenExcelLogFolder(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandOpenExcelLogFolder))
        {
            // No specific execution conditions - should always be available
        }

        protected override void ExecuteCore(object value)
        {
            // Get the environment-specific Excel log folder path
            var logPath = EnvironmentHelper.GetExcelLogPath();

            // Create the directory if it doesn't exist
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            // Open Windows Explorer to the log folder
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{logPath}\"",
                UseShellExecute = true
            });
        }
    }
}
