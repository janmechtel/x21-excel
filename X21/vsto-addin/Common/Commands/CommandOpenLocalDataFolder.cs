using X21.Common.Commands;
using X21.Common.Data;
using X21.Common.Model;
using System;
using System.Diagnostics;
using System.IO;

namespace X21.Common.Commands
{
    /// <summary>
    /// Command to open the root X21 data folder in LocalAppData (e.g. contains sqlite DB).
    /// </summary>
    public class CommandOpenLocalDataFolder : CommandBase
    {
        public CommandOpenLocalDataFolder(Container container, Annotation annotation)
            : base(container, annotation, nameof(CommandOpenLocalDataFolder))
        {
        }

        protected override void ExecuteCore(object value)
        {
            var dataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "X21");

            if (!Directory.Exists(dataPath))
            {
                Directory.CreateDirectory(dataPath);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"\"{dataPath}\"",
                UseShellExecute = true
            });
        }
    }
}
