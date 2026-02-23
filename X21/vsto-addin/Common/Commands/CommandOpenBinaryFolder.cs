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
    /// Command to open the binary folder where x21-backend.exe is located
    /// </summary>
    public class CommandOpenBinaryFolder : CommandBase
    {
        public CommandOpenBinaryFolder(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandOpenBinaryFolder))
        {
            // No specific execution conditions - should always be available
        }

        protected override void ExecuteCore(object value)
        {
            // Get the binary folder path using PathResolver
            var backendPath = PathResolver.GetBackendExecutablePath();
            var binaryFolder = Path.GetDirectoryName(backendPath);

            // If the backend path doesn't exist, try to find it in the assembly directory
            if (string.IsNullOrEmpty(binaryFolder) || !Directory.Exists(binaryFolder))
            {
                var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                binaryFolder = Path.GetDirectoryName(assemblyLocation);
            }

            // Open Windows Explorer to the binary folder
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{binaryFolder}\"",
                UseShellExecute = true
            });
        }
    }
}
